"""Browser automation layer built on the synchronous Playwright API.

This module owns the lifecycle of Playwright, the browser (real Google Chrome by
default, via Playwright's ``channel="chrome"``), and the
per-row browser sessions, plus the sign-in / sign-out flows. The most important
piece of UX here is the CAPTCHA pause inside :func:`sign_in`: the browser runs
HEADED so a human can solve the CAPTCHA, and we block until they tell us to
continue (or until the password field becomes editable, in "auto" mode).
"""

from __future__ import annotations

import sys
import time

from playwright.sync_api import (
    Browser,
    BrowserContext,
    Page,
    Playwright,
    sync_playwright,
)
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from .config import AppConfig


def launch(cfg: AppConfig) -> tuple[Playwright, Browser]:
    """Start Playwright and launch the browser.

    By default this drives **real Google Chrome** via Playwright's ``channel``
    mechanism (``cfg.browser_channel`` defaults to ``"chrome"``). Set
    ``browser_channel`` to null/empty in the config to fall back to Playwright's
    bundled Chromium instead.

    The browser is launched HEADED by default (``cfg.headless`` defaults to
    ``False``) so a human can interact with the visible window to solve the
    CAPTCHA.

    Returns the ``(playwright, browser)`` pair; the caller is responsible for
    eventually calling :func:`close`.
    """
    playwright = sync_playwright().start()
    launch_kwargs: dict = {"headless": cfg.headless}
    if cfg.browser_channel:
        # channel="chrome" launches the locally-installed Google Chrome rather
        # than the bundled Chromium build.
        launch_kwargs["channel"] = cfg.browser_channel
    browser = playwright.chromium.launch(**launch_kwargs)
    return playwright, browser


def new_session(browser: Browser, cfg: AppConfig) -> tuple[BrowserContext, Page]:
    """Create a fresh, isolated browser context and page.

    Each row gets its own context so cookies/storage do not leak between rows
    and every row starts logged-out. Default timeouts are applied from the
    config so all subsequent waits/actions inherit sensible limits.
    """
    context = browser.new_context()

    # Apply timeouts to the whole context so every page/locator inherits them.
    context.set_default_timeout(cfg.timeouts.element_ms)
    context.set_default_navigation_timeout(cfg.timeouts.navigation_ms)

    page = context.new_page()
    # Set on the page too, for clarity / belt-and-suspenders.
    page.set_default_timeout(cfg.timeouts.element_ms)
    page.set_default_navigation_timeout(cfg.timeouts.navigation_ms)

    return context, page


def _print_captcha_banner() -> None:
    """Print a clear, multi-line instruction banner for the human operator."""
    banner = (
        "\n"
        "============================================================\n"
        "  ACTION REQUIRED: SOLVE THE CAPTCHA\n"
        "------------------------------------------------------------\n"
        "  1. Switch to the visible Chrome browser window.\n"
        "  2. Solve the CAPTCHA / complete any human verification.\n"
        "  3. Come back HERE and press <Enter> to continue.\n"
        "============================================================\n"
    )
    # Write directly to stdout and flush so the banner appears immediately,
    # even if logging is buffering elsewhere.
    sys.stdout.write(banner)
    sys.stdout.flush()


def _wait_for_password_editable(page: Page, cfg: AppConfig) -> None:
    """Poll until the password field is visible AND editable ("auto" mode).

    We give the human a generous window to solve the CAPTCHA before the
    password field becomes interactable. We poll instead of relying on a single
    wait so a CAPTCHA that takes a while still succeeds.
    """
    # Be generous: allow at least two minutes (or the configured element
    # timeout, whichever is larger) for a human to solve the CAPTCHA.
    generous_ms = max(cfg.timeouts.element_ms, 120_000)
    deadline = time.monotonic() + (generous_ms / 1000.0)
    poll_interval_s = 0.5

    locator = page.locator(cfg.selectors.password_input).first
    last_error: Exception | None = None

    while time.monotonic() < deadline:
        try:
            if locator.is_visible() and locator.is_editable():
                return
        except Exception as exc:  # noqa: BLE001 - transient DOM states are expected
            last_error = exc
        time.sleep(poll_interval_s)

    raise TimeoutError(
        "Timed out waiting for the password field "
        f"({cfg.selectors.password_input!r}) to become editable in 'auto' "
        f"captcha mode after {generous_ms} ms."
        + (f" Last error: {last_error}" if last_error else "")
    )


def sign_in(page: Page, cfg: AppConfig, email: str, password: str) -> None:
    """Perform the full sign-in flow, pausing for a human-solved CAPTCHA.

    Steps:
      1. Navigate to the sign-in page.
      2. Fill the email field.
      3. CAPTCHA pause (mode-dependent):
         - "enter": print a banner and block on ``input()``.
         - "auto":  wait until the password field becomes editable.
      4. Fill the password field.
      5. Click the sign-in button.
      6. Wait for the post-login marker (raises on timeout).
    """
    selectors = cfg.selectors

    # 1. Go to the sign-in page.
    page.goto(cfg.signin_url)

    # 2. Fill the email. Wait for it to be visible/editable first for robustness.
    email_field = page.locator(selectors.email_input).first
    email_field.wait_for(state="visible")
    email_field.fill(email)

    # 3. CAPTCHA pause.
    mode = cfg.timeouts.captcha_wait_mode
    if mode == "enter":
        _print_captcha_banner()
        # Block on terminal input. This is the critical human-in-the-loop step.
        input()
    else:  # "auto"
        _wait_for_password_editable(page, cfg)

    # 4. Fill the password. Re-wait in case the CAPTCHA flow re-rendered the form.
    password_field = page.locator(selectors.password_input).first
    password_field.wait_for(state="visible")
    password_field.fill(password)

    # 5. Click the sign-in button.
    signin_button = page.locator(selectors.signin_button).first
    signin_button.wait_for(state="visible")
    signin_button.click()

    # 6. Wait for the post-login marker that appears ONLY after success.
    try:
        page.locator(selectors.post_login_marker).first.wait_for(state="visible")
    except PlaywrightTimeoutError as exc:
        raise RuntimeError(
            "Sign-in failed: post-login marker "
            f"({selectors.post_login_marker!r}) did not appear. The "
            "credentials may be wrong or the CAPTCHA was not solved correctly."
        ) from exc


def sign_out(page: Page, cfg: AppConfig) -> None:
    """Best-effort sign out between rows.

    If a sign-out selector is configured and present, click it. Any error is
    swallowed so a flaky sign-out never aborts the run (each new row gets a
    fresh context anyway).
    """
    selector = cfg.selectors.signout_button
    if not selector:
        return

    try:
        locator = page.locator(selector).first
        if locator.count() > 0 and locator.is_visible():
            locator.click()
    except Exception:  # noqa: BLE001 - best-effort, never raise
        pass


def close_session(context: BrowserContext) -> None:
    """Close a per-row browser context, swallowing any error."""
    try:
        context.close()
    except Exception:  # noqa: BLE001 - best-effort cleanup
        pass


def close(playwright: Playwright, browser: Browser) -> None:
    """Tear down the browser and Playwright, swallowing any error."""
    try:
        browser.close()
    except Exception:  # noqa: BLE001 - best-effort cleanup
        pass
    try:
        playwright.stop()
    except Exception:  # noqa: BLE001 - best-effort cleanup
        pass
