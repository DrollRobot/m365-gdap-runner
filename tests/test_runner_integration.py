"""Integration tests: runner with the real session, azure-auth clients, loader and examples."""

from __future__ import annotations

import os
from pathlib import Path

import httpx
import pytest
from azure_auth.constants import EXCHANGE_POWERSHELL_CLIENT_ID, GRAPH_CLI_CLIENT_ID

from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.config.schema import ENV_PREFIX
from m365_gdap_runner.loader import load_payload
from m365_gdap_runner.runner import run_payload
from m365_gdap_runner.session import PartnerSession
from tests._gdap_fakes import (
    CONSENT_REQUIRED,
    TENANT_A,
    TENANT_B,
    FakeMsal,
    bearer_tenant,
    fake_services,
    make_settings,
)

pytestmark = pytest.mark.integration

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"
LICENSED_USERS = EXAMPLES / "licensed_users.py"
AUDIT_DISABLED = EXAMPLES / "audit_disabled.py"


@pytest.fixture(autouse=True)
def _isolated_config(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Strip this package's env vars and point the config dir at a temp dir."""
    for key in list(os.environ):
        if key.startswith(ENV_PREFIX):
            monkeypatch.delenv(key)
    monkeypatch.setenv(CONFIG_DIR_ENV, str(tmp_path))


def test_graph_example_runs_through_session_and_azure_auth(fake_msal: FakeMsal) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return fake_services(request)

    payload = load_payload(str(LICENSED_USERS))
    session = PartnerSession(make_settings(), transport=httpx.MockTransport(handler))
    session.login(payload.clients)
    tenants = session.list_tenants()
    results = run_payload(session, payload, tenants, output_format="csv")

    assert [t.tenant_id for t in tenants] == [TENANT_A, TENANT_B]
    assert [r.status for r in results] == ["ok", "auth_failed"]
    assert results[0].result == [
        {
            "userPrincipalName": f"admin@{TENANT_A}.example",
            "displayName": "Admin",
            "skuIds": "sku-1;sku-2",
        }
    ]
    assert results[1].error is not None
    assert "AADSTS65001" in results[1].error

    # Consent was checked with the payload's own scopes, not .default.
    assert fake_msal.silent_calls(TENANT_B) == [
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/User.Read.All",))
    ]
    # The payload never ran for the tenant that refused the token.
    users_requests = [r for r in seen if r.url.path == "/v1.0/users"]
    assert [bearer_tenant(r) for r in users_requests] == [TENANT_A]
    assert users_requests[0].headers["ConsistencyLevel"] == "eventual"
    assert fake_msal.interactive_calls == 1


def test_exchange_example_signs_in_to_exchange_and_runs_the_cmdlet(fake_msal: FakeMsal) -> None:
    payload = load_payload(str(AUDIT_DISABLED))
    assert (payload.clients, payload.scopes) == (("exchange",), ())

    session = PartnerSession(make_settings(), transport=httpx.MockTransport(fake_services))
    session.login(payload.clients)
    results = run_payload(session, payload, session.list_tenants())

    assert [r.status for r in results] == ["ok", "ok"]
    assert [r.result for r in results] == [
        {"name": f"{TENANT_A}.onmicrosoft.com", "auditDisabled": False},
        {"name": f"{TENANT_B}.onmicrosoft.com", "auditDisabled": True},
    ]
    # Graph for the tenant list, Exchange for the cmdlet: two home sign-ins.
    assert [c.client_id for c in fake_msal.calls if c.method == "interactive"] == [
        GRAPH_CLI_CLIENT_ID,
        EXCHANGE_POWERSHELL_CLIENT_ID,
    ]
    # Each customer was checked with the Exchange client id before the cmdlet ran.
    assert fake_msal.silent_calls(TENANT_A)[0] == (
        EXCHANGE_POWERSHELL_CLIENT_ID,
        ("https://outlook.office365.com/.default",),
    )
