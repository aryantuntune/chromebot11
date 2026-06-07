"""Tests for :mod:`src.config` loading and validation."""

from __future__ import annotations

import pytest

from src.config import AppConfig, ConfigError, load_config

# A complete, valid configuration covering every required key in the contract.
VALID_CONFIG_YAML = """\
site_url: "https://example.com"
signin_url: "https://example.com/signin"
subpage_url: "https://example.com/messages"
headless: false

selectors:
  email_input: "#email"
  password_input: "#password"
  signin_button: "#signin"
  post_login_marker: "#dashboard"
  subpage_link: null
  table: "table.results"
  message_row: "tbody tr"
  message_cell: "td.message"
  signout_button: "#signout"

excel:
  input_dir: "input"
  output_dir: "output"
  sheet_name: null
  header_row: 1
  first_data_row: 2
  email_col: "A"
  password_col: "B"
  result_col: "C"

timeouts:
  navigation_ms: 30000
  element_ms: 15000
  captcha_wait_mode: "enter"
"""


def _write(tmp_path, text):
    path = tmp_path / "config.yaml"
    path.write_text(text)
    return path


def test_load_config_returns_populated_appconfig(tmp_path):
    path = _write(tmp_path, VALID_CONFIG_YAML)
    cfg = load_config(str(path))

    assert isinstance(cfg, AppConfig)
    assert cfg.site_url == "https://example.com"
    assert cfg.signin_url == "https://example.com/signin"
    assert cfg.subpage_url == "https://example.com/messages"
    assert cfg.headless is False

    # Selectors
    assert cfg.selectors.email_input == "#email"
    assert cfg.selectors.password_input == "#password"
    assert cfg.selectors.signin_button == "#signin"
    assert cfg.selectors.post_login_marker == "#dashboard"
    assert cfg.selectors.subpage_link is None
    assert cfg.selectors.table == "table.results"
    assert cfg.selectors.message_row == "tbody tr"
    assert cfg.selectors.message_cell == "td.message"
    assert cfg.selectors.signout_button == "#signout"

    # Excel
    assert cfg.excel.input_dir == "input"
    assert cfg.excel.output_dir == "output"
    assert cfg.excel.sheet_name is None
    assert cfg.excel.header_row == 1
    assert cfg.excel.first_data_row == 2
    assert cfg.excel.email_col == "A"
    assert cfg.excel.password_col == "B"
    assert cfg.excel.result_col == "C"

    # Timeouts
    assert cfg.timeouts.navigation_ms == 30000
    assert cfg.timeouts.element_ms == 15000
    assert cfg.timeouts.captcha_wait_mode == "enter"


def test_load_config_missing_required_key_raises_configerror(tmp_path):
    # Drop the top-level "signin_url" key to trigger a validation error.
    lines = [
        line
        for line in VALID_CONFIG_YAML.splitlines()
        if not line.startswith("signin_url:")
    ]
    broken = "\n".join(lines)
    path = _write(tmp_path, broken)

    with pytest.raises(ConfigError):
        load_config(str(path))


def test_load_config_missing_nested_key_raises_configerror(tmp_path):
    # Remove a required nested selector key.
    lines = [
        line
        for line in VALID_CONFIG_YAML.splitlines()
        if "email_input:" not in line
    ]
    broken = "\n".join(lines)
    path = _write(tmp_path, broken)

    with pytest.raises(ConfigError):
        load_config(str(path))
