"""Tests for this package's Settings fields resolved through the config system."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from m365_gdap_runner.config import ConfigError, load_settings
from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.config.schema import ENV_PREFIX

pytestmark = pytest.mark.integration

GRAPH_COMMAND_LINE_TOOLS = "14d82eec-204b-4c2f-b7e8-296a70dab67e"


@pytest.fixture(autouse=True)
def _isolated_config(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Strip this package's env vars and point the config dir at an empty temp dir."""
    for key in list(os.environ):
        if key.startswith(ENV_PREFIX):
            monkeypatch.delenv(key)
    monkeypatch.setenv(CONFIG_DIR_ENV, str(tmp_path))


def test_defaults(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", "partner")
    settings = load_settings()
    assert settings.partner_tenant_id == "partner"
    assert settings.client_id == GRAPH_COMMAND_LINE_TOOLS
    assert settings.persist_token_cache is False


def test_partner_tenant_id_is_required() -> None:
    with pytest.raises(ConfigError, match="partner_tenant_id"):
        load_settings()


def test_env_vars_override_defaults(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", "partner")
    monkeypatch.setenv(ENV_PREFIX + "CLIENT_ID", "custom-app")
    monkeypatch.setenv(ENV_PREFIX + "PERSIST_TOKEN_CACHE", "true")
    settings = load_settings()
    assert settings.client_id == "custom-app"
    assert settings.persist_token_cache is True


def test_overrides_beat_env_vars(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", "partner")
    monkeypatch.setenv(ENV_PREFIX + "CLIENT_ID", "from-env")
    assert load_settings(overrides={"client_id": "from-flag"}).client_id == "from-flag"
