"""Configuration loading and validation for the sign-in / scrape bot.

This module defines strongly-typed dataclasses describing every piece of
configuration the bot needs, plus a ``load_config`` function that reads a YAML
file, validates that all required keys are present and well-typed, and returns a
fully-populated :class:`AppConfig`.

All validation failures raise :class:`ConfigError` with a clear, specific
message that names the offending key, so a misconfigured ``config.yaml`` is easy
to fix.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import yaml


class ConfigError(Exception):
    """Raised when the configuration file is missing keys or has invalid values."""


# --------------------------------------------------------------------------- #
# Dataclasses describing the configuration schema.
# --------------------------------------------------------------------------- #
@dataclass
class Selectors:
    """CSS/Playwright selectors used to drive the target site."""

    email_input: str
    password_input: str
    signin_button: str
    post_login_marker: str   # selector that appears ONLY after successful login
    subpage_link: str | None  # optional selector clicked to reach subpage; if None use subpage_url
    table: str               # selector for the results table
    message_row: str         # selector for the first/top data row within the table
    message_cell: str        # selector for the message cell within that row (relative to row)
    signout_button: str | None  # optional; best-effort sign out between rows


@dataclass
class ExcelConfig:
    """Where input/output workbooks live and how to read/write them."""

    input_dir: str
    output_dir: str
    sheet_name: str | None   # None = active sheet
    header_row: int          # 1-based worksheet row holding headers; 0 = no header
    first_data_row: int      # 1-based worksheet row where data starts
    email_col: str           # column letter, e.g. "A"
    password_col: str        # column letter, e.g. "B"
    result_col: str          # column letter, e.g. "C"


@dataclass
class Timeouts:
    """Timeout values and CAPTCHA-handling strategy."""

    navigation_ms: int
    element_ms: int
    captcha_wait_mode: str   # "enter" (wait for Enter key in terminal) OR "auto" (wait for marker/password field)


@dataclass
class AppConfig:
    """Top-level configuration object consumed by the rest of the bot."""

    site_url: str
    signin_url: str          # absolute URL of the sign-in page
    subpage_url: str         # absolute URL of the sub-page with the table
    headless: bool           # default False so a human can solve the CAPTCHA
    browser_channel: str | None  # "chrome" (default) drives real Google Chrome; None = bundled Chromium
    selectors: Selectors
    excel: ExcelConfig
    timeouts: Timeouts


# --------------------------------------------------------------------------- #
# Validation helpers.
# --------------------------------------------------------------------------- #
def _require(mapping: dict[str, Any], key: str, context: str) -> Any:
    """Return ``mapping[key]`` or raise ConfigError naming the missing key."""
    if not isinstance(mapping, dict):
        raise ConfigError(f"'{context}' must be a mapping/block, got {type(mapping).__name__}.")
    if key not in mapping:
        raise ConfigError(f"Missing required config key: '{context}.{key}'." if context
                          else f"Missing required config key: '{key}'.")
    return mapping[key]


def _as_str(value: Any, key: str) -> str:
    """Coerce/validate a non-empty string value."""
    if value is None:
        raise ConfigError(f"Config key '{key}' must be a string, got null/None.")
    if not isinstance(value, str):
        # Coerce simple scalars (e.g. numbers) to str but reject containers.
        if isinstance(value, (dict, list)):
            raise ConfigError(f"Config key '{key}' must be a string, got {type(value).__name__}.")
        value = str(value)
    if value.strip() == "":
        raise ConfigError(f"Config key '{key}' must be a non-empty string.")
    return value


def _as_opt_str(value: Any, key: str) -> str | None:
    """Coerce/validate an optional string (None allowed)."""
    if value is None:
        return None
    return _as_str(value, key)


def _as_int(value: Any, key: str) -> int:
    """Coerce/validate an integer value."""
    if isinstance(value, bool):
        # bool is a subclass of int; reject it explicitly to avoid surprises.
        raise ConfigError(f"Config key '{key}' must be an integer, got a boolean.")
    if isinstance(value, int):
        return value
    if isinstance(value, str) and value.strip().lstrip("-").isdigit():
        return int(value.strip())
    raise ConfigError(f"Config key '{key}' must be an integer, got {value!r}.")


def _as_bool(value: Any, key: str) -> bool:
    """Coerce/validate a boolean value."""
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in ("true", "yes", "1"):
            return True
        if lowered in ("false", "no", "0"):
            return False
    raise ConfigError(f"Config key '{key}' must be a boolean (true/false), got {value!r}.")


def _as_col_letter(value: Any, key: str) -> str:
    """Validate that a value is an Excel column letter like 'A' or 'AB'."""
    text = _as_str(value, key).strip().upper()
    if not text.isalpha():
        raise ConfigError(
            f"Config key '{key}' must be a column letter (e.g. 'A', 'B', 'AB'), got {value!r}."
        )
    return text


# --------------------------------------------------------------------------- #
# Sub-block builders.
# --------------------------------------------------------------------------- #
def _build_selectors(block: Any) -> Selectors:
    if not isinstance(block, dict):
        raise ConfigError("'selectors' must be a mapping/block.")
    return Selectors(
        email_input=_as_str(_require(block, "email_input", "selectors"), "selectors.email_input"),
        password_input=_as_str(_require(block, "password_input", "selectors"), "selectors.password_input"),
        signin_button=_as_str(_require(block, "signin_button", "selectors"), "selectors.signin_button"),
        post_login_marker=_as_str(
            _require(block, "post_login_marker", "selectors"), "selectors.post_login_marker"
        ),
        subpage_link=_as_opt_str(block.get("subpage_link"), "selectors.subpage_link"),
        table=_as_str(_require(block, "table", "selectors"), "selectors.table"),
        message_row=_as_str(_require(block, "message_row", "selectors"), "selectors.message_row"),
        message_cell=_as_str(_require(block, "message_cell", "selectors"), "selectors.message_cell"),
        signout_button=_as_opt_str(block.get("signout_button"), "selectors.signout_button"),
    )


def _build_excel(block: Any) -> ExcelConfig:
    if not isinstance(block, dict):
        raise ConfigError("'excel' must be a mapping/block.")

    header_row = _as_int(_require(block, "header_row", "excel"), "excel.header_row")
    if header_row < 0:
        raise ConfigError("Config key 'excel.header_row' must be >= 0 (0 means no header).")

    first_data_row = _as_int(_require(block, "first_data_row", "excel"), "excel.first_data_row")
    if first_data_row < 1:
        raise ConfigError("Config key 'excel.first_data_row' must be >= 1.")

    return ExcelConfig(
        input_dir=_as_str(_require(block, "input_dir", "excel"), "excel.input_dir"),
        output_dir=_as_str(_require(block, "output_dir", "excel"), "excel.output_dir"),
        sheet_name=_as_opt_str(block.get("sheet_name"), "excel.sheet_name"),
        header_row=header_row,
        first_data_row=first_data_row,
        email_col=_as_col_letter(_require(block, "email_col", "excel"), "excel.email_col"),
        password_col=_as_col_letter(_require(block, "password_col", "excel"), "excel.password_col"),
        result_col=_as_col_letter(_require(block, "result_col", "excel"), "excel.result_col"),
    )


def _build_timeouts(block: Any) -> Timeouts:
    if not isinstance(block, dict):
        raise ConfigError("'timeouts' must be a mapping/block.")

    navigation_ms = _as_int(_require(block, "navigation_ms", "timeouts"), "timeouts.navigation_ms")
    element_ms = _as_int(_require(block, "element_ms", "timeouts"), "timeouts.element_ms")
    if navigation_ms <= 0:
        raise ConfigError("Config key 'timeouts.navigation_ms' must be a positive integer.")
    if element_ms <= 0:
        raise ConfigError("Config key 'timeouts.element_ms' must be a positive integer.")

    mode = _as_str(
        _require(block, "captcha_wait_mode", "timeouts"), "timeouts.captcha_wait_mode"
    ).strip().lower()
    if mode not in ("enter", "auto"):
        raise ConfigError(
            "Config key 'timeouts.captcha_wait_mode' must be 'enter' or 'auto', "
            f"got {mode!r}."
        )

    return Timeouts(
        navigation_ms=navigation_ms,
        element_ms=element_ms,
        captcha_wait_mode=mode,
    )


# --------------------------------------------------------------------------- #
# Public loader.
# --------------------------------------------------------------------------- #
def load_config(path: str = "config.yaml") -> AppConfig:
    """Load and validate the YAML config at ``path`` into an :class:`AppConfig`.

    Raises :class:`ConfigError` with a specific message naming any missing or
    invalid key.
    """
    try:
        with open(path, "r", encoding="utf-8") as fh:
            raw = yaml.safe_load(fh)
    except FileNotFoundError as exc:
        raise ConfigError(f"Config file not found: {path!r}.") from exc
    except yaml.YAMLError as exc:
        raise ConfigError(f"Could not parse YAML config {path!r}: {exc}") from exc

    if raw is None:
        raise ConfigError(f"Config file {path!r} is empty.")
    if not isinstance(raw, dict):
        raise ConfigError(f"Config file {path!r} must contain a top-level mapping/block.")

    selectors = _build_selectors(_require(raw, "selectors", ""))
    excel = _build_excel(_require(raw, "excel", ""))
    timeouts = _build_timeouts(_require(raw, "timeouts", ""))

    # headless defaults to False (headed) when omitted, so a human can solve the CAPTCHA.
    headless_raw = raw.get("headless", False)
    headless = _as_bool(headless_raw, "headless")

    # browser_channel defaults to "chrome" so the bot drives real Google Chrome.
    # Set it to null/empty in YAML to fall back to Playwright's bundled Chromium.
    browser_channel = _as_opt_str(raw.get("browser_channel", "chrome"), "browser_channel")

    return AppConfig(
        site_url=_as_str(_require(raw, "site_url", ""), "site_url"),
        signin_url=_as_str(_require(raw, "signin_url", ""), "signin_url"),
        subpage_url=_as_str(_require(raw, "subpage_url", ""), "subpage_url"),
        headless=headless,
        browser_channel=browser_channel,
        selectors=selectors,
        excel=excel,
        timeouts=timeouts,
    )
