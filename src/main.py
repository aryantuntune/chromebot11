"""Orchestrator for the sign-in / scrape bot.

Run from the project root with:

    python -m src.main

(Running ``python src/main.py`` directly will break the relative imports;
always invoke it as a module so that ``from .config import ...`` resolves.)

Flow per row of the input spreadsheet:
  1. Start a fresh, logged-out browser session.
  2. Sign in using the email/password read from the Excel row
     (pausing for a human to solve the CAPTCHA).
  3. Open the sub-page and read the first/top data row's message cell.
  4. Write that message back into the result column and save incrementally.

NOTE: This bot ONLY uses credentials supplied in the input Excel file. It never
authenticates to any external/cloud account.
"""

from __future__ import annotations

import logging
import sys

from .browser import (
    close,
    close_session,
    launch,
    new_session,
    sign_in,
    sign_out,
)
from .config import load_config
from .excel_io import (
    find_input_file,
    load_workbook_and_rows,
    save_workbook,
    write_result,
)
from .scraper import read_message


def _setup_logging() -> None:
    """Configure INFO-level, timestamped logging for the run."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def run() -> int:
    """Execute the full bot run. Returns a process exit code (0 == success)."""
    _setup_logging()
    log = logging.getLogger("bot")

    # --- Load configuration -------------------------------------------------
    cfg = load_config()

    # --- Locate and load the input workbook --------------------------------
    input_path = find_input_file(cfg.excel.input_dir)
    log.info("Using input file: %s", input_path)

    wb, ws, rows = load_workbook_and_rows(input_path, cfg.excel)
    log.info("Found %d data row(s) to process.", len(rows))

    # --- Launch the browser once for the whole run -------------------------
    pw, browser = launch(cfg)
    out_path = None
    try:
        for row in rows:
            ctx = None
            try:
                # Each row gets a fresh, logged-out session.
                ctx, page = new_session(browser, cfg)
                log.info("Row %d: signing in as %s", row.row_number, row.email)

                sign_in(page, cfg, row.email, row.password)
                msg = read_message(page, cfg)

                write_result(ws, cfg.excel, row.row_number, msg)
                out_path = save_workbook(wb, input_path, cfg.excel.output_dir)
                log.info(
                    "Row %d: success -> wrote message (%d chars), saved to %s",
                    row.row_number,
                    len(msg),
                    out_path,
                )

                # Best-effort sign-out, then tear down the session.
                sign_out(page, cfg)
            except Exception as e:  # noqa: BLE001 - keep processing remaining rows
                log.exception("Row %d: failed: %s", row.row_number, e)
                # Record the error in the result column so it is visible to the user.
                try:
                    write_result(ws, cfg.excel, row.row_number, f"ERROR: {e}")
                    out_path = save_workbook(
                        wb, input_path, cfg.excel.output_dir
                    )
                except Exception:  # noqa: BLE001 - never let bookkeeping abort the run
                    log.exception(
                        "Row %d: additionally failed to write error to workbook",
                        row.row_number,
                    )
            finally:
                close_session(ctx)
    finally:
        close(pw, browser)

    if out_path is not None:
        log.info("Done. Final output written to: %s", out_path)
    else:
        log.warning("Done. No output file was produced (no rows processed?).")

    return 0


if __name__ == "__main__":
    sys.exit(run())
