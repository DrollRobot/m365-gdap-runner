"""Integration tests for PartnerSession: real azure-auth over a fake msal and transport."""

from __future__ import annotations

import httpx
import pytest
from azure_auth import ConsentRequired
from azure_auth.constants import (
    AZURE_POWERSHELL_CLIENT_ID,
    EXCHANGE_POWERSHELL_CLIENT_ID,
    GRAPH_CLI_CLIENT_ID,
)

from m365_gdap_runner.config import ConfigError
from m365_gdap_runner.session import PartnerSession
from m365_gdap_runner.tenants import Tenant
from tests._gdap_fakes import (
    CONSENT_REQUIRED,
    PARTNER_TENANT,
    TENANT_A,
    TENANT_B,
    TENANT_C,
    FakeMsal,
    bearer_tenant,
    customers_body,
    fake_services,
    make_settings,
)

pytestmark = pytest.mark.integration

PARTNER_SCOPE = "https://graph.microsoft.com/DelegatedAdminRelationship.Read.All"
TRANSPORT = httpx.MockTransport(fake_services)


def test_login_signs_in_once_per_client_id(fake_msal: FakeMsal) -> None:
    session = PartnerSession(make_settings(), transport=TRANSPORT)
    session.login(["graph", "exchange", "ipps", "azure"])
    interactive = [c for c in fake_msal.calls if c.method == "interactive"]
    assert [(c.tenant_id, c.client_id, c.scopes) for c in interactive] == [
        (PARTNER_TENANT, GRAPH_CLI_CLIENT_ID, (PARTNER_SCOPE,)),
        (
            PARTNER_TENANT,
            EXCHANGE_POWERSHELL_CLIENT_ID,
            ("https://outlook.office365.com/.default",),
        ),
        (PARTNER_TENANT, AZURE_POWERSHELL_CLIENT_ID, ("https://management.azure.com/.default",)),
    ]
    # Security & Compliance shares Exchange's client id, so it signed in silently.
    assert (
        EXCHANGE_POWERSHELL_CLIENT_ID,
        ("https://ps.compliance.protection.outlook.com/.default",),
    ) in fake_msal.silent_calls(PARTNER_TENANT)


def test_login_always_covers_graph_for_listing_tenants(fake_msal: FakeMsal) -> None:
    PartnerSession(make_settings(), transport=TRANSPORT).login(["exchange"])
    assert [c.client_id for c in fake_msal.calls if c.method == "interactive"] == [
        GRAPH_CLI_CLIENT_ID,
        EXCHANGE_POWERSHELL_CLIENT_ID,
    ]


def test_login_with_own_client_id_is_one_sign_in(fake_msal: FakeMsal) -> None:
    session = PartnerSession(make_settings(client_id="my-app"), transport=TRANSPORT)
    session.login(["graph", "exchange"])
    assert fake_msal.interactive_calls == 1
    assert {app.client_id for app in fake_msal.apps} == {"my-app"}


def test_login_rejects_unknown_client(fake_msal: FakeMsal) -> None:
    with pytest.raises(ValueError, match="Unknown client kind 'teams'"):
        PartnerSession(make_settings(), transport=TRANSPORT).login(["teams"])


def test_list_tenants_dedupes_sorts_and_names_nameless_by_id(fake_msal: FakeMsal) -> None:
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        body = customers_body(
            (TENANT_B, "beta corp"), (TENANT_A, "Alpha"), (TENANT_B.upper(), "beta corp")
        )
        body["value"].append({"id": "no-tenant"})
        body["value"].append({"tenantId": TENANT_C})
        return httpx.Response(200, json=body)

    session = PartnerSession(make_settings(), transport=httpx.MockTransport(handler))
    session.login()
    tenants = session.list_tenants()

    assert tenants == [
        Tenant(TENANT_A, "Alpha"),
        Tenant(TENANT_B, "beta corp"),
        Tenant(TENANT_C, TENANT_C),
    ]
    (request,) = seen
    assert request.url.path == "/v1.0/tenantRelationships/delegatedAdminCustomers"
    assert bearer_tenant(request) == PARTNER_TENANT
    assert fake_msal.interactive_calls == 1


def test_context_for_is_a_sibling_whose_probe_checks_the_customer(fake_msal: FakeMsal) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    session = PartnerSession(make_settings(), transport=TRANSPORT)
    session.login(["graph"])

    with session.context_for(Tenant(TENANT_A, "Alpha"), scopes=["User.Read.All"]) as alpha:
        assert alpha.auth.is_sibling
        assert alpha.auth.tenant_id == TENANT_A
        alpha.probe()
        assert bearer_tenant(alpha.graph.request("GET", "/users").request) == TENANT_A

    with (
        session.context_for(Tenant(TENANT_B, "beta corp"), scopes=["User.Read.All"]) as beta,
        pytest.raises(ConsentRequired) as info,
    ):
        beta.probe()
    assert info.value.tenant_id == TENANT_B
    assert fake_msal.silent_calls(TENANT_B) == [
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/User.Read.All",))
    ]
    assert fake_msal.interactive_calls == 1


def test_partner_tenant_id_is_the_guid(fake_msal: FakeMsal) -> None:
    session = PartnerSession(make_settings(partner_tenant_id="partner.example"))
    assert session.partner_tenant_id != "partner.example"
    assert session.auth.tenant_name == "partner.example"


@pytest.mark.parametrize(
    ("field", "value", "message"),
    [("cache", "plaintext", "cache must be one of"), ("cloud", "Mars", "Unknown cloud 'Mars'")],
)
def test_bad_cache_or_cloud_is_a_config_error(
    fake_msal: FakeMsal, field: str, value: str, message: str
) -> None:
    with pytest.raises(ConfigError, match=message):
        PartnerSession(make_settings(**{field: value}))


def test_cloud_setting_skips_the_lookup(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch
) -> None:
    def no_lookup(tenant: str) -> object:
        raise AssertionError(f"looked up {tenant}")

    monkeypatch.setattr("azure_auth.auth.context.discover_tenant", no_lookup)
    session = PartnerSession(make_settings(cloud="USGov"))
    assert session.auth.cloud.name == "USGov"
