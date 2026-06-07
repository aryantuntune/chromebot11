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
import sys
import time

from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from src.config import ExcelConfig
from src.excel_io import (
    find_input_file,
    load_workbook_and_rows,
    save_workbook,
    write_result,
)

# --------------------------------------------------------------------------- #
# Settings (kept inline; the YAML config model does not fit this staged flow).
# --------------------------------------------------------------------------- #
LOGIN_URL = "https://myhpgas.in/myHPGas/PortalLogin.aspx"
BROWSER_CHANNEL = "chrome"   # drive real Google Chrome
HEADLESS = False             # MUST be headed so the human can solve the CAPTCHA

NAV_TIMEOUT_MS = 45_000
ELEMENT_TIMEOUT_MS = 30_000

# How long to wait for the human to solve the CAPTCHA (the password field
# appearing is the signal that the CAPTCHA was accepted).
CAPTCHA_SOLVE_TIMEOUT_MS = 300_000  # 5 minutes per row

# Selectors captured by the recorder.
SEL_EMAIL = "#ContentPlaceHolder1_txtUserNameEmail"
SEL_CAPTCHA = "#ContentPlaceHolder1_loginCaptcha_tbCaptchaInput"
SEL_PASSWORD = "#ContentPlaceHolder1_txtPassword"
SEL_LOGIN_BTN = "#ContentPlaceHolder1_btnLogin"

# The input file columns: A=conno, B=id(email), C=password, D=code(result).
EXCEL_CFG = ExcelConfig(
    input_dir="input",
    output_dir="output",
    sheet_name=None,
    header_row=1,
    first_data_row=2,
    email_col="B",
    password_col="C",
    result_col="D",
)


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


def _capture_result(page) -> str:
    """Best-effort: read something meaningful into the 'code' column.

    We don't know the exact element you want recorded yet, so this grabs the
    first table row's text on the post-login page if present, otherwise a short
    status string. Tell me the precise field and I'll point this at it.
    """
    for sel in ("table tbody tr:first-child", "table tr:nth-child(2)", "table tr"):
        try:
            loc = page.locator(sel).first
            if loc.count() > 0 and loc.is_visible():
                text = (loc.inner_text(timeout=3_000) or "").strip()
                if text:
                    return " | ".join(t.strip() for t in text.splitlines() if t.strip())[:300]
        except Exception:  # noqa: BLE001 - best-effort
            continue
    return "LOGGED_IN_OK"


def _process_row(browser, email: str, password: str) -> str:
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

        # 6. Best-effort: open "View Cylinder Booking history" if present.
        try:
            link = page.get_by_role("link", name="View Cylinder Booking history")
            link.wait_for(state="visible", timeout=5_000)
            link.click()
            page.wait_for_load_state("domcontentloaded")
        except PlaywrightTimeoutError:
            pass

        return _capture_result(page)
    finally:
        try:
            context.close()
        except Exception:  # noqa: BLE001 - best-effort cleanup
            pass


def main() -> int:
    _setup_logging()
    log = logging.getLogger("bot")

    input_path = find_input_file(EXCEL_CFG.input_dir)
    log.info("Using input file: %s", input_path)

    wb, ws, rows = load_workbook_and_rows(input_path, EXCEL_CFG)
    log.info("Found %d data row(s) to process.", len(rows))

    out_path = None
    with sync_playwright() as pw:
        launch_kwargs = {"headless": HEADLESS, "args": ["--incognito"]}
        if BROWSER_CHANNEL:
            launch_kwargs["channel"] = BROWSER_CHANNEL
        browser = pw.chromium.launch(**launch_kwargs)
        try:
            for row in rows:
                log.info("Row %d: signing in as %s", row.row_number, row.email)
                try:
                    result = _process_row(browser, row.email, row.password)
                    write_result(ws, EXCEL_CFG, row.row_number, result)
                    log.info("Row %d: done -> %r", row.row_number, result)
                except KeyboardInterrupt:
                    log.warning("Interrupted by user. Saving progress and exiting.")
                    break
                except Exception as e:  # noqa: BLE001 - keep going on the next row
                    log.exception("Row %d: failed: %s", row.row_number, e)
                    write_result(ws, EXCEL_CFG, row.row_number, f"ERROR: {e}")
                out_path = save_workbook(wb, input_path, EXCEL_CFG.output_dir)
        finally:
            try:
                browser.close()
            except Exception:  # noqa: BLE001
                pass

    if out_path is not None:
        log.info("Done. Output written to: %s", out_path)
    else:
        log.warning("Done. No output produced.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
