"""Integration tests for standalone(): one tenant, real session and clients, fake msal."""

from __future__ import annotations

import os
from pathlib import Path

import httpx
import pytest

from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.config.schema import ENV_PREFIX
from m365_gdap_runner.loader import load_payload
from m365_gdap_runner.standalone import standalone
from m365_gdap_runner.tenants import Tenant
from tests._gdap_fakes import (
    CONSENT_REQUIRED,
    PARTNER_TENANT,
    TENANT_A,
    TENANT_B,
    USER,
    FakeMsal,
    fake_services,
)

pytestmark = pytest.mark.integration

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"
TRANSPORT = httpx.MockTransport(fake_services)


@pytest.fixture(autouse=True)
def _config(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Isolate config and supply the partner tenant and user through the environment."""
    for key in list(os.environ):
        if key.startswith(ENV_PREFIX):
            monkeypatch.delenv(key)
    monkeypatch.setenv(CONFIG_DIR_ENV, str(tmp_path))
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", PARTNER_TENANT)
    monkeypatch.setenv(ENV_PREFIX + "USERNAME", USER)


def test_returns_context_for_named_tenant_and_leaves_other_args(fake_msal: FakeMsal) -> None:
    with standalone(
        ["User.Read.All"],
        argv=["--tenant", "BETA CORP", "--payload-flag", "value"],
        transport=TRANSPORT,
    ) as ctx:
        users = ctx.graph.get_all("/users")
    assert ctx.tenant == Tenant(TENANT_B, "beta corp")
    assert ctx.auth.tenant.id == TENANT_B
    assert users[0]["userPrincipalName"] == f"admin@{TENANT_B}.example"
    assert fake_msal.interactive_calls == 1


def test_prompts_for_tenant_when_flag_is_absent(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch
) -> None:
    prompts: list[str] = []

    def fake_input(prompt: str) -> str:
        prompts.append(prompt)
        return "Alpha"

    monkeypatch.setattr("builtins.input", fake_input)
    with standalone(["User.Read.All"], argv=[], transport=TRANSPORT) as ctx:
        assert ctx.tenant.tenant_id == TENANT_A
    assert prompts == ["Tenant name or ID: "]


@pytest.mark.parametrize("example", ["licensed_users.py", "audit_disabled.py"])
def test_example_payloads_run_standalone(fake_msal: FakeMsal, example: str) -> None:
    payload = load_payload(str(EXAMPLES / example))
    with standalone(
        payload.scopes, payload.clients, ["--tenant", "Alpha"], transport=TRANSPORT
    ) as ctx:
        result = payload.run(ctx)
    expected: object
    if example == "licensed_users.py":
        expected = [
            {
                "userPrincipalName": f"admin@{TENANT_A}.example",
                "displayName": "Admin",
                "skuIds": "sku-1;sku-2",
            }
        ]
    else:
        expected = {"name": f"{TENANT_A}.onmicrosoft.com", "auditDisabled": False}
    assert result == expected


def test_exits_1_when_scope_is_not_consented(
    fake_msal: FakeMsal, capsys: pytest.CaptureFixture[str]
) -> None:
    fake_msal.tenant_errors[TENANT_A] = CONSENT_REQUIRED
    with pytest.raises(SystemExit) as info:
        standalone(["User.Read.All"], argv=["--tenant", "Alpha"], transport=TRANSPORT)
    assert info.value.code == 1
    assert "AADSTS65001" in capsys.readouterr().err


def test_exits_2_on_unknown_tenant(fake_msal: FakeMsal, capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as info:
        standalone(["User.Read.All"], argv=["--tenant", "Contoso"], transport=TRANSPORT)
    assert info.value.code == 2
    assert "No GDAP tenant matches 'Contoso'" in capsys.readouterr().err


@pytest.mark.parametrize(
    ("scopes", "clients", "message"),
    [
        ([".default"], ["graph"], "must not use"),
        (["User.Read.All"], ["exchange"], "'graph' is not in CLIENTS"),
        ([], ["teams"], "unknown client 'teams'"),
    ],
)
def test_exits_2_on_invalid_declarations_before_sign_in(
    fake_msal: FakeMsal,
    capsys: pytest.CaptureFixture[str],
    scopes: list[str],
    clients: list[str],
    message: str,
) -> None:
    with pytest.raises(SystemExit) as info:
        standalone(scopes, clients, ["--tenant", "Alpha"], transport=TRANSPORT)
    assert info.value.code == 2
    assert fake_msal.interactive_calls == 0
    assert message in capsys.readouterr().err
