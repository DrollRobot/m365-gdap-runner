"""Integration tests for standalone(): one tenant, real session and Graph client, fake MSAL."""

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
    FakeMsal,
    fake_access,
    fake_graph,
)

pytestmark = pytest.mark.integration

EXAMPLE = Path(__file__).resolve().parents[1] / "examples" / "licensed_users.py"
TRANSPORT = httpx.MockTransport(fake_graph)


@pytest.fixture(autouse=True)
def _config(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Isolate config and supply the partner tenant through the environment."""
    for key in list(os.environ):
        if key.startswith(ENV_PREFIX):
            monkeypatch.delenv(key)
    monkeypatch.setenv(CONFIG_DIR_ENV, str(tmp_path))
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", PARTNER_TENANT)


def test_returns_context_for_named_tenant_and_leaves_other_args() -> None:
    ctx = standalone(
        ["User.Read.All"],
        ["--tenant", "BETA CORP", "--payload-flag", "value"],
        app_factory=FakeMsal(),
        transport=TRANSPORT,
    )
    with ctx.graph:
        users = ctx.graph.get_all("/users")
    assert ctx.tenant == Tenant(TENANT_B, "beta corp")
    assert users[0]["userPrincipalName"] == f"admin@{TENANT_B}.example"
    assert ctx.credential.get_token("User.Read.All").token == fake_access(TENANT_B)


def test_prompts_for_tenant_when_flag_is_absent(monkeypatch: pytest.MonkeyPatch) -> None:
    prompts: list[str] = []

    def fake_input(prompt: str) -> str:
        prompts.append(prompt)
        return "Alpha"

    monkeypatch.setattr("builtins.input", fake_input)
    ctx = standalone(["User.Read.All"], [], app_factory=FakeMsal(), transport=TRANSPORT)
    ctx.graph.close()
    assert ctx.tenant.tenant_id == TENANT_A
    assert prompts == ["Tenant name or ID: "]


def test_example_payload_runs_standalone() -> None:
    payload = load_payload(str(EXAMPLE))
    ctx = standalone(
        payload.scopes, ["--tenant", "Alpha"], app_factory=FakeMsal(), transport=TRANSPORT
    )
    with ctx.graph:
        rows = payload.run(ctx)
    assert rows == [
        {
            "userPrincipalName": f"admin@{TENANT_A}.example",
            "displayName": "Admin",
            "skuIds": "sku-1;sku-2",
        }
    ]


def test_exits_1_when_scope_is_not_consented(capsys: pytest.CaptureFixture[str]) -> None:
    fake = FakeMsal(tenant_errors={TENANT_A: CONSENT_REQUIRED})
    with pytest.raises(SystemExit) as info:
        standalone(["User.Read.All"], ["--tenant", "Alpha"], app_factory=fake, transport=TRANSPORT)
    assert info.value.code == 1
    assert "consent_required" in capsys.readouterr().err


def test_exits_2_on_unknown_tenant(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as info:
        standalone(
            ["User.Read.All"], ["--tenant", "Contoso"], app_factory=FakeMsal(), transport=TRANSPORT
        )
    assert info.value.code == 2
    assert "No GDAP tenant matches 'Contoso'" in capsys.readouterr().err


def test_exits_2_on_invalid_scopes_before_sign_in(capsys: pytest.CaptureFixture[str]) -> None:
    fake = FakeMsal()
    with pytest.raises(SystemExit) as info:
        standalone([".default"], ["--tenant", "Alpha"], app_factory=fake, transport=TRANSPORT)
    assert info.value.code == 2
    assert fake.interactive_calls == 0
    assert "must not use" in capsys.readouterr().err
