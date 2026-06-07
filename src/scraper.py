"""Scraper module: reads the first/top data row's message cell from the sub-page.

Uses the synchronous Playwright API (playwright.sync_api). The public surface is a
single function, ``read_message``, matching the shared interface contract.
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

if TYPE_CHECKING:
    # Imported only for type checking to avoid hard runtime coupling / circular imports.
    from playwright.sync_api import Page

    from .config import AppConfig

logger = logging.getLogger(__name__)


class ScrapeError(Exception):
    """Raised when the sub-page table or message cell cannot be located/read."""


def read_message(page: "Page", cfg: "AppConfig") -> str:
    """Navigate to the sub-page and return the first data row's message text.

    Steps:
      1. Reach the sub-page: click ``selectors.subpage_link`` if it is set,
         otherwise ``page.goto(cfg.subpage_url)``.
      2. Wait for ``selectors.table`` to appear.
      3. Locate ``selectors.message_row`` (the first/top data row) and, within it,
         ``selectors.message_cell``.
      4. Return the cell's ``inner_text()`` stripped of surrounding whitespace.

    Raises:
        ScrapeError: if the sub-page link, table, message row, or message cell
            cannot be found (the message names the specific selector that failed).
    """
    selectors = cfg.selectors

    # --- 1. Reach the sub-page --------------------------------------------------
    if selectors.subpage_link:
        logger.info("Navigating to sub-page via link selector: %s", selectors.subpage_link)
        try:
            page.click(selectors.subpage_link)
        except PlaywrightTimeoutError as exc:
            raise ScrapeError(
                f"Could not find/click the sub-page link selector "
                f"{selectors.subpage_link!r}: {exc}"
            ) from exc
        except PlaywrightError as exc:
            raise ScrapeError(
                f"Error clicking the sub-page link selector "
                f"{selectors.subpage_link!r}: {exc}"
            ) from exc
    else:
        logger.info("Navigating to sub-page URL: %s", cfg.subpage_url)
        try:
            page.goto(cfg.subpage_url)
        except PlaywrightError as exc:
            raise ScrapeError(
                f"Failed to navigate to sub-page URL {cfg.subpage_url!r}: {exc}"
            ) from exc

    # --- 2. Wait for the results table -----------------------------------------
    try:
        page.wait_for_selector(selectors.table)
    except PlaywrightTimeoutError as exc:
        raise ScrapeError(
            f"Results table not found using selector {selectors.table!r} "
            f"on sub-page: {exc}"
        ) from exc

    table = page.locator(selectors.table).first

    # --- 3. Locate the first/top data row, then the message cell within it ------
    row = table.locator(selectors.message_row).first
    try:
        # Ensure the top row actually exists before drilling into its cell.
        row.wait_for(state="attached")
    except PlaywrightTimeoutError as exc:
        raise ScrapeError(
            f"Message row not found using selector {selectors.message_row!r} "
            f"within table {selectors.table!r}: {exc}"
        ) from exc

    cell = row.locator(selectors.message_cell).first
    try:
        cell.wait_for(state="attached")
    except PlaywrightTimeoutError as exc:
        raise ScrapeError(
            f"Message cell not found using selector {selectors.message_cell!r} "
            f"within the first data row {selectors.message_row!r}: {exc}"
        ) from exc

    # --- 4. Read and return the stripped text ----------------------------------
    try:
        text = cell.inner_text()
    except PlaywrightError as exc:
        raise ScrapeError(
            f"Failed to read text from message cell selector "
            f"{selectors.message_cell!r}: {exc}"
        ) from exc

    message = text.strip()
    logger.info("Read message from top data row (%d chars).", len(message))
    return message
