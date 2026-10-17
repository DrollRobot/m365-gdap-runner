"""Tests for this package's Settings fields resolved through the config system."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from m365_gdap_runner.config import ConfigError, load_settings
from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.config.schema import ENV_PREFIX

pytestmark = pytest.mark.integration


@pytest.fixture(autouse=True)
def _isolated_config(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Strip this package's env vars and point the config dir at an empty temp dir."""
    for key in list(os.environ):
        if key.startswith(ENV_PREFIX):
            monkeypatch.delenv(key)
    monkeypatch.setenv(CONFIG_DIR_ENV, str(tmp_path))


def _required(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", "partner")
    monkeypatch.setenv(ENV_PREFIX + "USERNAME", "admin@partner.example")


def test_defaults(monkeypatch: pytest.MonkeyPatch) -> None:
    _required(monkeypatch)
    settings = load_settings()
    assert settings.partner_tenant_id == "partner"
    assert settings.username == "admin@partner.example"
    assert settings.client_id == ""
    assert settings.cache == "memory"
    assert settings.cloud == ""
    assert settings.tenants_file == ""
    assert settings.debug_log_path == ""


@pytest.mark.parametrize("missing", ["PARTNER_TENANT_ID", "USERNAME"])
def test_partner_tenant_and_username_are_required(
    monkeypatch: pytest.MonkeyPatch, missing: str
) -> None:
    _required(monkeypatch)
    monkeypatch.delenv(ENV_PREFIX + missing)
    with pytest.raises(ConfigError, match=missing.lower()):
        load_settings()


def test_env_vars_override_defaults(monkeypatch: pytest.MonkeyPatch) -> None:
    _required(monkeypatch)
    monkeypatch.setenv(ENV_PREFIX + "CLIENT_ID", "custom-app")
    monkeypatch.setenv(ENV_PREFIX + "CACHE", "disk")
    monkeypatch.setenv(ENV_PREFIX + "CLOUD", "USGov")
    settings = load_settings()
    assert (settings.client_id, settings.cache, settings.cloud) == ("custom-app", "disk", "USGov")


def test_overrides_beat_env_vars(monkeypatch: pytest.MonkeyPatch) -> None:
    _required(monkeypatch)
    monkeypatch.setenv(ENV_PREFIX + "CLIENT_ID", "from-env")
    assert load_settings(overrides={"client_id": "from-flag"}).client_id == "from-flag"
