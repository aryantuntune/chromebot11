"""Replay the recorded HP Gas login flow for every row of the input Excel.

This runner reproduces, per spreadsheet row, the exact sequence of actions
captured with the Playwright recorder (see ``recorded_flow.py``):

  1. Open the HP Gas portal login page.
  2. Fill the Mobile/E-Mail field with the row's email (column B / ``id``).
  3. PAUSE: the human solves the CAPTCHA in the visible Chrome window and
     presses <Enter> in the CAPTCHA box (this reveals the password field),
     then returns to THIS terminal and presses <Enter> to continue.
  4. Fill the password (column C) once the password field appears.
  5. Click Login, dismiss the "OK" confirmation dialog (best-effort), and
     click "View Cylinder Booking history" (best-effort).
  6. Record a result string into the ``code`` column (column D) and save a
     copy of the workbook into ``output/`` after every row.

Run it from the project root, in YOUR OWN terminal (so the Enter-pause works):

    .venv/bin/python run_bot.py

Each row opens a fresh, logged-out browser context. Press Ctrl-C any time to
stop; progress is saved incrementally after each completed row.
"""

from __future__ import annotations

import logging
import os
import pathlib
import sys
import time

from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

# --------------------------------------------------------------------------- #
# Settings (kept inline; the YAML config model does not fit this staged flow).
# --------------------------------------------------------------------------- #
LOGIN_URL = "https://myhpgas.in/myHPGas/PortalLogin.aspx"
HEADLESS = False             # MUST be headed so the human can solve the CAPTCHA

# Which browser to drive. Override with BOT_BROWSER=chrome|edge|brave|opera.
# The booking page on this portal is flaky; if one browser loads it poorly,
# try another. All are Chromium-based, so the automation is identical.
BROWSER = os.environ.get("BOT_BROWSER", "chrome").strip().lower()

NAV_TIMEOUT_MS = 90_000
ELEMENT_TIMEOUT_MS = 45_000
# The booking-history page is server-rendered and can be slow; give the grid
# plenty of time to appear after we click into it before we read the OTP.
BOOKING_NAV_TIMEOUT_MS = 120_000

# How long to wait for the human to solve the CAPTCHA (the password field
# appearing is the signal that the CAPTCHA was accepted).
CAPTCHA_SOLVE_TIMEOUT_MS = 300_000  # 5 minutes per row

# Selectors captured by the recorder.
SEL_EMAIL = "#ContentPlaceHolder1_txtUserNameEmail"
SEL_CAPTCHA = "#ContentPlaceHolder1_loginCaptcha_tbCaptchaInput"
SEL_PASSWORD = "#ContentPlaceHolder1_txtPassword"
SEL_LOGIN_BTN = "#ContentPlaceHolder1_btnLogin"

# Booking-history grid shown after login. Newest booking is the FIRST data row.
# 0-based <td> indices in a row:
#   0=Order Date 1=Order Ref No 2=Order No 3=No.of Cyl 4=Status 5=Date
#   6=Cash Memo No 7=Cash Memo Amount 8=Cash Memo Date 9=delivery/OTP cell ...
# The delivery cell (index 9) shows "OutForDelivery (OTP: 7149)" while the
# cylinder is out for delivery, or "Delivered" once delivered. The Status cell
# (index 4) shows "In process" / "Delivered". The OTP we want is in cell 9.
SEL_BOOKING_TABLE = "#ContentPlaceHolder1_gvBookingHistory"
BOOKING_ROW = "tr.GridviewScrollItem"
STATUS_COL_IDX = 4       # 0-based <td> index of the Status cell
ORDERREF_COL_IDX = 1     # 0-based <td> index of the Order Ref No cell
DELIVERY_COL_IDX = 9     # 0-based <td> index of the "OutForDelivery (OTP: …)" cell

# --------------------------------------------------------------------------- #
# File layout. Two input workbooks live in input/:
#   * MASTER  -- the big SHREELALJI credentials DB (>=5 columns, no header):
#         A=conno(consumer id)  B=name  C=email(login id)  D=Hpgas-pwd  E=password
#   * TARGETS -- a single column of consumer IDs to actually process this run.
# We match each target consumer ID against column A of the master to pull that
# account's login email (C) + password (E).
# --------------------------------------------------------------------------- #
INPUT_DIR = "input"
OUTPUT_DIR = "output"

MASTER_SHEET = "Sheet1"
MASTER_CONNO_COL = "A"     # consumer id (matched against the targets)
MASTER_EMAIL_COL = "C"     # login id (email / mobile)
MASTER_PWD_COL = "E"       # login password (e.g. 'shree1234')

TARGETS_CONNO_COL = "A"    # the targets file's consumer-id column

OTP_RESULTS_NAME = "OTP_results.xlsx"   # separate output: Consumer ID | OTP

# Optional run controls via environment variables (run.bat stays simple):
#   BOT_MAX_ROWS=5  -> only process the first 5 target IDs (great for a test run)
#   BOT_DEBUG=1     -> dump screenshot + page text/html per account to output/debug/
MAX_ROWS = int(os.environ["BOT_MAX_ROWS"]) if os.environ.get("BOT_MAX_ROWS") else None
DEBUG_DUMP = os.environ.get("BOT_DEBUG", "").strip() not in ("", "0", "false", "no")


def _setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def _wait_for_captcha_solved(page, email: str) -> None:
    """Block until the human solves the CAPTCHA in the browser.

    The signal is the password field becoming visible: on this site the
    password box only appears AFTER the CAPTCHA is typed and accepted (Enter in
    the CAPTCHA box). We poll for it so there is no dependency on terminal input
    — the operator interacts only with the visible Chrome window.
    """
    banner = (
        "\n"
        "============================================================\n"
        f"  ACTION REQUIRED  —  signing in as: {email}\n"
        "------------------------------------------------------------\n"
        "  In the visible Chrome window:\n"
        "    1. Type the CAPTCHA code shown on the page.\n"
        "    2. Press <Enter> in the CAPTCHA box.\n"
        "  The PASSWORD field then appears and the bot continues\n"
        "  automatically — nothing to do in this terminal.\n"
        "============================================================\n"
    )
    sys.stdout.write(banner)
    sys.stdout.flush()

    deadline = time.monotonic() + (CAPTCHA_SOLVE_TIMEOUT_MS / 1000.0)
    pwd = page.locator(SEL_PASSWORD).first
    while time.monotonic() < deadline:
        try:
            if pwd.is_visible():
                return
        except Exception:  # noqa: BLE001 - transient DOM states during reload
            pass
        time.sleep(0.5)
    raise TimeoutError(
        "CAPTCHA not solved in time: the password field "
        f"({SEL_PASSWORD!r}) never appeared within "
        f"{CAPTCHA_SOLVE_TIMEOUT_MS // 1000}s."
    )


def _debug_dump(page, tag: str) -> None:
    """Save a screenshot + visible page text so we can locate the OTP element.

    Only runs when BOT_DEBUG is set. Files land in output/debug/ and are named
    after the account so the screenshot and text line up.
    """
    import pathlib
    import re

    safe = re.sub(r"[^A-Za-z0-9._-]", "_", tag)[:60] or "row"
    out = pathlib.Path("output") / "debug"
    out.mkdir(parents=True, exist_ok=True)
    try:
        page.screenshot(path=str(out / f"{safe}.png"), full_page=True)
    except Exception:  # noqa: BLE001 - best-effort
        pass
    try:
        text = page.locator("body").inner_text(timeout=5_000)
        (out / f"{safe}.txt").write_text(text, encoding="utf-8")
    except Exception:  # noqa: BLE001 - best-effort
        pass
    try:
        (out / f"{safe}.html").write_text(page.content(), encoding="utf-8")
    except Exception:  # noqa: BLE001 - best-effort
        pass


def _extract_otp(text: str) -> str:
    """Pull the delivery OTP out of a cell like 'OutForDelivery (OTP: 7149)'.

    Returns the digit code, or '' when there is no OTP (e.g. 'Delivered' or a
    still-'In process' booking that has not gone out for delivery yet).
    """
    import re

    if not text:
        return ""
    # Preferred: an explicit 'OTP: 7149' token.
    m = re.search(r"OTP[:\s)]*?(\d{3,8})", text, re.IGNORECASE)
    if m:
        return m.group(1)
    low = text.lower()
    # Plain 'Delivered' (not 'OutForDelivery') has no OTP.
    if "deliver" in low and "outfor" not in low and "out for" not in low:
        return ""
    # Fallback: a standalone digit group near a delivery status.
    m2 = re.search(r"\b(\d{3,8})\b", text)
    return m2.group(1) if m2 else ""


def _capture_result(page) -> tuple[str, str]:
    """Read the OTP/Status and Order Ref No of the LATEST booking.

    On the booking-history grid the newest booking is the first data row. Its
    Status cell holds the delivery OTP while the booking is pending, or
    "Delivered" once the cylinder has been delivered. Returns
    ``(status_or_otp, order_ref_no)``.
    """
    try:
        table = page.locator(SEL_BOOKING_TABLE).first
        try:
            table.wait_for(state="visible", timeout=BOOKING_NAV_TIMEOUT_MS)
        except PlaywrightTimeoutError:
            return ("NO_BOOKING_TABLE", "")

        row = table.locator(BOOKING_ROW).first
        if row.count() == 0:
            return ("NO_BOOKINGS", "")
        cells = row.locator("td")
        n = cells.count()

        def _cell(i: int) -> str:
            if i >= n:
                return ""
            text = (cells.nth(i).inner_text(timeout=5_000) or "").strip()
            return " ".join(text.split())  # collapse whitespace/newlines

        status = _cell(STATUS_COL_IDX)
        delivery = _cell(DELIVERY_COL_IDX)
        order_ref = _cell(ORDERREF_COL_IDX)
        # The OTP, when present, sits in the delivery cell as
        # "OutForDelivery (OTP: 7149)". Prefer that; fall back to the Status
        # cell. If there's genuinely no OTP yet, return the human-readable
        # status so the column explains why (e.g. "Delivered" / "In process").
        otp = _extract_otp(delivery) or _extract_otp(status)
        if otp:
            return (otp, order_ref)
        return (delivery or status or "EMPTY_STATUS", order_ref)
    except Exception as e:  # noqa: BLE001 - best-effort
        return (f"CAPTURE_ERROR: {e}", "")


def _txt(value) -> str:
    """Coerce a cell value to a stripped string ('' for None)."""
    return "" if value is None else str(value).strip()


# Result values that represent a technical failure worth retrying on re-run
# (as opposed to a real status like an OTP, 'Delivered', or 'In process').
_RETRY_MARKERS = (
    "NO_BOOKING_TABLE", "NO_BOOKINGS", "EMPTY_STATUS", "CAPTURE_ERROR", "ERROR",
)


def _needs_retry(value) -> bool:
    """True if a previously recorded result should be attempted again."""
    s = _txt(value).upper()
    return (not s) or any(s.startswith(m) for m in _RETRY_MARKERS)


def _classify_inputs(input_dir: str):
    """Locate the master credentials file and the targets (consumer-ID list).

    The master is the workbook with many columns (>=5, i.e. it holds the
    credentials); the targets file is a single-column list of consumer IDs. If
    several of either exist, the most recently modified one wins.
    """
    xlsx = [
        p
        for p in pathlib.Path(input_dir).glob("*.xlsx")
        if p.is_file() and not p.name.startswith("~$")
    ]
    if not xlsx:
        raise FileNotFoundError(f"No .xlsx files found in {input_dir!r}.")

    masters, lists = [], []
    for p in xlsx:
        wb = load_workbook(p, read_only=True)
        max_col = wb.active.max_column or 1
        wb.close()
        (masters if max_col >= 5 else lists).append(p)

    if not masters:
        raise FileNotFoundError(
            "No master credentials file (a workbook with >=5 columns) found in input/."
        )
    if not lists:
        raise FileNotFoundError(
            "No targets file (a single-column list of consumer IDs) found in input/."
        )
    master = max(masters, key=lambda p: p.stat().st_mtime)
    targets = max(lists, key=lambda p: p.stat().st_mtime)
    return master, targets


def _load_targets(path) -> list[str]:
    """Return the ordered consumer IDs (as strings) from the targets workbook."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    col = column_index_from_string(TARGETS_CONNO_COL)
    ids: list[str] = []
    for row in ws.iter_rows(min_col=col, max_col=col, values_only=True):
        v = row[0]
        if v not in (None, ""):
            ids.append(_txt(v))
    wb.close()
    return ids


def _load_master(path):
    """Open the master workbook (writable) and index it by consumer ID.

    Returns ``(workbook, worksheet, info)`` where ``info[consumer_id]`` is a
    dict ``{"row": int, "email": str, "password": str}``. The first occurrence
    of a consumer ID wins.
    """
    wb = load_workbook(path, data_only=False)
    ws = wb[MASTER_SHEET] if MASTER_SHEET in wb.sheetnames else wb.active
    a = column_index_from_string(MASTER_CONNO_COL)
    c = column_index_from_string(MASTER_EMAIL_COL)
    e = column_index_from_string(MASTER_PWD_COL)
    info: dict[str, dict] = {}
    for r in range(1, ws.max_row + 1):
        cid = ws.cell(row=r, column=a).value
        if cid in (None, ""):
            continue
        key = _txt(cid)
        if key in info:
            continue
        info[key] = {
            "row": r,
            "email": _txt(ws.cell(row=r, column=c).value),
            "password": _txt(ws.cell(row=r, column=e).value),
        }
    return wb, ws, info


def _first_existing(paths: list[str]) -> str | None:
    for p in paths:
        if p and pathlib.Path(p).exists():
            return p
    return None


# Detected locations for the non-channel Chromium browsers.
_LOCALAPPDATA = os.environ.get("LOCALAPPDATA", "")
_PF = os.environ.get("ProgramFiles", r"C:\Program Files")
_PF86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
_BRAVE_PATHS = [
    rf"{_PF}\BraveSoftware\Brave-Browser\Application\brave.exe",
    rf"{_PF86}\BraveSoftware\Brave-Browser\Application\brave.exe",
    rf"{_LOCALAPPDATA}\BraveSoftware\Brave-Browser\Application\brave.exe",
]
_OPERA_PATHS = [
    rf"{_LOCALAPPDATA}\Programs\Opera GX\opera.exe",
    rf"{_LOCALAPPDATA}\Programs\Opera\opera.exe",
]


def _launch_browser(pw):
    """Launch the browser selected by BOT_BROWSER (default Chrome).

    Chrome/Edge use Playwright channels; Brave/Opera launch by executable path.
    Each account already runs in its own fresh context, so we don't need an
    explicit incognito flag for isolation.
    """
    log = logging.getLogger("bot")
    if BROWSER in ("edge", "msedge"):
        log.info("Browser: Microsoft Edge")
        return pw.chromium.launch(headless=HEADLESS, channel="msedge")
    if BROWSER == "brave":
        path = _first_existing(_BRAVE_PATHS)
        if path:
            log.info("Browser: Brave (%s)", path)
            return pw.chromium.launch(headless=HEADLESS, executable_path=path)
        log.warning("Brave not found; falling back to Chrome.")
    if BROWSER == "opera":
        path = _first_existing(_OPERA_PATHS)
        if path:
            log.info("Browser: Opera (%s)", path)
            return pw.chromium.launch(headless=HEADLESS, executable_path=path)
        log.warning("Opera not found; falling back to Chrome.")
    log.info("Browser: Google Chrome")
    return pw.chromium.launch(headless=HEADLESS, channel="chrome")


def _open_booking_history(page) -> bool:
    """Navigate to the booking-history grid, retrying because the page is flaky.

    On each attempt we (re)click the menu link if it's visible, then wait for
    the grid. If it doesn't appear we reload and try again. Re-clicking/reloading
    fixes most blank loads WITHOUT needing a fresh login or CAPTCHA. Returns True
    if the grid became visible.
    """
    for attempt in range(1, 4):
        try:
            link = page.get_by_role(
                "link", name="View Cylinder Booking history"
            ).first
            if link.is_visible():
                link.click()
        except Exception:  # noqa: BLE001 - best-effort
            pass
        try:
            page.locator(SEL_BOOKING_TABLE).first.wait_for(
                state="visible", timeout=45_000
            )
            return True
        except PlaywrightTimeoutError:
            logging.getLogger("bot").info(
                "  booking grid not ready (attempt %d/3) -> reloading...", attempt
            )
            try:
                page.reload(wait_until="domcontentloaded")
            except Exception:  # noqa: BLE001 - best-effort
                pass
    return False


def _process_row(browser, email: str, password: str) -> tuple[str, str]:
    """Run the recorded flow once for a single row; return the result string."""
    context = browser.new_context()
    context.set_default_timeout(ELEMENT_TIMEOUT_MS)
    context.set_default_navigation_timeout(NAV_TIMEOUT_MS)
    # Auto-accept any native JS dialogs so they never block the run.
    page = context.new_page()
    page.on("dialog", lambda d: d.accept())
    try:
        page.goto(LOGIN_URL)

        # 1. Email.
        email_field = page.locator(SEL_EMAIL).first
        email_field.wait_for(state="visible")
        email_field.fill(email)

        # 2. Human solves the CAPTCHA in the browser; we wait for the password
        #    field to appear (the signal the CAPTCHA was accepted).
        _wait_for_captcha_solved(page, email)

        # 3. Password (appears after the CAPTCHA is accepted).
        pwd_field = page.locator(SEL_PASSWORD).first
        pwd_field.wait_for(state="visible")
        pwd_field.fill(password)

        # 4. Submit.
        page.locator(SEL_LOGIN_BTN).first.click()

        # 5. Best-effort: dismiss an "OK" confirmation modal if one appears.
        try:
            ok = page.get_by_role("button", name="OK")
            ok.wait_for(state="visible", timeout=5_000)
            ok.click()
        except PlaywrightTimeoutError:
            pass

        # 6. Open the booking-history grid, retrying/reloading if the flaky
        #    page comes up blank (no new CAPTCHA needed -- we stay logged in).
        _open_booking_history(page)

        if DEBUG_DUMP:
            _debug_dump(page, email)

        return _capture_result(page)
    finally:
        try:
            context.close()
        except Exception:  # noqa: BLE001 - best-effort cleanup
            pass


def main() -> int:
    _setup_logging()
    log = logging.getLogger("bot")

    master_path, targets_path = _classify_inputs(INPUT_DIR)
    log.info("Master credentials file : %s", master_path.name)
    log.info("Targets (consumer IDs)  : %s", targets_path.name)

    target_ids = _load_targets(targets_path)
    log.info("Loaded %d target consumer ID(s).", len(target_ids))
    if MAX_ROWS is not None:
        target_ids = target_ids[:MAX_ROWS]
        log.info("BOT_MAX_ROWS set -> processing only the first %d.", len(target_ids))

    log.info("Loading master workbook (this can take ~20s on the big file)...")
    master_wb, master_ws, info = _load_master(master_path)
    log.info("Master indexed: %d unique consumer IDs.", len(info))

    # The OTP goes into a brand-new column at the very end of the master.
    otp_col_idx = (master_ws.max_column or 0) + 1
    otp_col_letter = get_column_letter(otp_col_idx)
    log.info("OTP -> master column %s (new last column).", otp_col_letter)

    out_dir = pathlib.Path(OUTPUT_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)
    master_out = out_dir / f"{master_path.stem}.result.xlsx"
    otp_out = out_dir / OTP_RESULTS_NAME

    # Separate OTP results workbook (Consumer ID | OTP). Resume from a prior run
    # if one exists, so we only re-solve CAPTCHAs for what's left / what failed.
    if otp_out.exists():
        res_wb = load_workbook(str(otp_out))
        res_ws = res_wb.active
        log.info("Resuming from existing %s.", otp_out.name)
    else:
        res_wb = Workbook()
        res_ws = res_wb.active
        res_ws.title = "OTP"
        res_ws.append(["Consumer ID", "OTP"])
    # Index existing result rows by consumer ID for in-place upsert.
    res_rows: dict[str, int] = {}
    for r in range(2, res_ws.max_row + 1):
        key = _txt(res_ws.cell(row=r, column=1).value)
        if key:
            res_rows[key] = r

    def _record(cid: str, otp: str) -> None:
        if cid in res_rows:
            res_ws.cell(row=res_rows[cid], column=2, value=otp)
        else:
            res_ws.append([cid, otp])
            res_rows[cid] = res_ws.max_row

    processed = 0
    skipped = 0
    total = len(target_ids)
    with sync_playwright() as pw:
        browser = _launch_browser(pw)
        try:
            for i, cid in enumerate(target_ids, 1):
                rec = info.get(cid)
                if rec is None or not rec["email"]:
                    log.warning("[%d/%d] Consumer %s: not found in master / no email -> skipped.",
                                i, total, cid)
                    _record(cid, "NOT_IN_MASTER")
                    res_wb.save(str(otp_out))
                    continue

                # Already captured in a previous run? Mirror it into the master
                # (so master_out stays complete) and skip the browser/CAPTCHA.
                prev = res_ws.cell(row=res_rows[cid], column=2).value if cid in res_rows else None
                if prev is not None and not _needs_retry(prev):
                    master_ws.cell(row=rec["row"], column=otp_col_idx, value=prev)
                    skipped += 1
                    log.info("[%d/%d] Consumer %s -> already have %r, skipping.",
                             i, total, cid, _txt(prev))
                    continue

                log.info("[%d/%d] Consumer %s -> signing in as %s",
                         i, total, cid, rec["email"])
                try:
                    otp, order_ref = _process_row(browser, rec["email"], rec["password"])
                except KeyboardInterrupt:
                    log.warning("Interrupted by user. Saving progress and exiting.")
                    break
                except Exception as ex:  # noqa: BLE001 - keep going on the next account
                    log.exception("Consumer %s failed: %s", cid, ex)
                    otp, order_ref = (f"ERROR: {ex}", "")

                master_ws.cell(row=rec["row"], column=otp_col_idx, value=otp)
                _record(cid, otp)
                processed += 1
                log.info("[%d/%d] Consumer %s -> OTP=%r (order ref %s)",
                         i, total, cid, otp, order_ref)

                # The small OTP file is saved every account (never lose an OTP);
                # the big master is saved every 10 accounts and again at the end.
                res_wb.save(str(otp_out))
                if processed % 10 == 0:
                    master_wb.save(str(master_out))
                    log.info("Saved master progress (%d processed).", processed)
        finally:
            try:
                browser.close()
            except Exception:  # noqa: BLE001
                pass

    res_wb.save(str(otp_out))
    master_wb.save(str(master_out))
    log.info("Done. %d processed, %d already-had, %d total target(s).",
             processed, skipped, total)
    log.info("OTP list  -> %s", otp_out)
    log.info("Master+OTP (column %s) -> %s", otp_col_letter, master_out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
