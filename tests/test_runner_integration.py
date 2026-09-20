"""Integration test: runner with the real session, credential, Graph client, and loader."""

from __future__ import annotations

import os
from pathlib import Path

import httpx
import pytest

from m365_gdap_runner.auth import PartnerSession
from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.config.schema import ENV_PREFIX
from m365_gdap_runner.loader import load_payload
from m365_gdap_runner.runner import run_payload
from m365_gdap_runner.tenants import list_gdap_tenants
from tests._gdap_fakes import (
    CONSENT_REQUIRED,
    TENANT_A,
    TENANT_B,
    FakeMsal,
    bearer_tenant,
    fake_graph,
    make_settings,
)

pytestmark = pytest.mark.integration

EXAMPLE = Path(__file__).resolve().parents[1] / "examples" / "licensed_users.py"


@pytest.fixture(autouse=True)
def _isolated_config(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Strip this package's env vars and point the config dir at a temp dir."""
    for key in list(os.environ):
        if key.startswith(ENV_PREFIX):
            monkeypatch.delenv(key)
    monkeypatch.setenv(CONFIG_DIR_ENV, str(tmp_path))


def test_example_payload_runs_through_real_session_and_graph_client() -> None:
    fake = FakeMsal(tenant_errors={TENANT_B: CONSENT_REQUIRED})
    session = PartnerSession(make_settings(), app_factory=fake, token_cache=object())
    session.login()
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return fake_graph(request)

    transport = httpx.MockTransport(handler)
    tenants = list_gdap_tenants(session, transport=transport)
    results = run_payload(
        session, load_payload(str(EXAMPLE)), tenants, output_format="csv", transport=transport
    )

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
    assert "consent_required" in results[1].error

    # Consent was checked with the payload's own scopes, not .default.
    assert (TENANT_B, ("User.Read.All",), None) in fake.silent_calls
    # The payload never ran for the tenant that refused the token.
    users_requests = [r for r in seen if r.url.path == "/v1.0/users"]
    assert [bearer_tenant(r) for r in users_requests] == [TENANT_A]
    assert users_requests[0].headers["ConsistencyLevel"] == "eventual"
    assert fake.interactive_calls == 1
