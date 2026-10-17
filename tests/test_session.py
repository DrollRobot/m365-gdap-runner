"""Integration tests for PartnerSession: real azure-auth over a fake msal and transport."""

from __future__ import annotations

import httpx
import pytest
from azure_auth import AuthError, ConsentRequired, TenantNotFound
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
    GAMMA_ADMIN,
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
        assert alpha.auth.tenant.id == TENANT_A
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
    assert session.auth.tenant.domain == "partner.example"


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
    def no_lookup(name: str, *args: object, **kwargs: object) -> object:
        raise AssertionError(f"looked up {name}")

    monkeypatch.setattr("azure_auth.Tenant.lookup", no_lookup)
    session = PartnerSession(make_settings(cloud="USGov"))
    assert session.auth.tenant.cloud.name == "USGov"
    assert session.partner_tenant_id == PARTNER_TENANT
    assert session.cloud_name == "USGov"


def test_cloud_setting_with_a_domain_looks_it_up_in_that_cloud(fake_msal: FakeMsal) -> None:
    session = PartnerSession(make_settings(partner_tenant_id="partner.example", cloud="USGov"))
    assert session.cloud_name == "USGov"
    assert session.auth.tenant.domain == "partner.example"
    assert session.partner_tenant_id != "partner.example"


def test_cloud_of_finds_a_tenant_by_discovery(fake_msal: FakeMsal) -> None:
    session = PartnerSession(make_settings(), transport=TRANSPORT)
    assert session.cloud_name == "Commercial"
    assert session.cloud_of(TENANT_A) == "Commercial"


@pytest.mark.parametrize(
    "error",
    [
        TenantNotFound("No tenant", name=TENANT_A),
        httpx.ConnectError("no route to login.microsoftonline.com"),
    ],
)
def test_cloud_of_is_blank_when_discovery_fails(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, error: Exception
) -> None:
    def failing(name: str, *args: object, **kwargs: object) -> object:
        raise error

    session = PartnerSession(make_settings())
    monkeypatch.setattr("azure_auth.Tenant.lookup", failing)
    assert session.cloud_of(TENANT_A) == ""


# --- check_access ----------------------------------------------------------------------

ALPHA = Tenant(TENANT_A, "Alpha", "Commercial")
BETA = Tenant(TENANT_B, "beta corp", "Commercial")
GAMMA_GOV = Tenant(TENANT_C, "Gamma", "USGov", GAMMA_ADMIN)
BETA_ADMIN = "admin@beta.example"


def _signed_in(clients: list[str]) -> PartnerSession:
    session = PartnerSession(make_settings(), transport=TRANSPORT)
    session.login(clients)
    return session


def test_check_access_reaches_gdap_tenants_through_the_partner_sign_in(
    fake_msal: FakeMsal,
) -> None:
    session = _signed_in(["graph"])
    session.check_access([ALPHA, BETA], scopes=["User.Read.All"], jobs=2)
    with session.context_for(ALPHA, scopes=["User.Read.All"]) as alpha:
        assert alpha.auth.is_sibling
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]
    assert fake_msal.silent_calls(TENANT_B) == [
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/User.Read.All",))
    ]


def test_check_access_signs_in_directly_where_the_partner_sign_in_fails(
    fake_msal: FakeMsal,
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    beta_with_admin = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)
    session = _signed_in(["graph"])
    session.check_access([ALPHA, beta_with_admin], scopes=["User.Read.All"])
    with session.context_for(beta_with_admin, scopes=["User.Read.All"]) as beta:
        assert not beta.auth.is_sibling
        assert (beta.auth.tenant.id, beta.auth.username) == (TENANT_B, BETA_ADMIN)
        assert bearer_tenant(beta.graph.request("GET", "/users").request) == TENANT_B
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]


@pytest.mark.regression
@pytest.mark.parametrize(
    ("tenant", "reason"),
    [(BETA, "AADSTS65001"), (Tenant(TENANT_C, "Gamma", "USGov"), "Gamma is in the USGov cloud")],
)
def test_check_access_never_signs_in_directly_as_the_partner_account(
    fake_msal: FakeMsal, tenant: Tenant, reason: str
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    session = _signed_in(["graph"])
    session.check_access([tenant], scopes=["User.Read.All"])
    with pytest.raises(AuthError, match="m365_username is blank") as info:
        session.context_for(tenant, scopes=["User.Read.All"])
    assert reason in str(info.value)
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]


def test_check_access_signs_in_to_another_cloud_once_per_client_as_its_account(
    fake_msal: FakeMsal,
) -> None:
    session = _signed_in(["graph", "exchange"])
    session.check_access([GAMMA_GOV], scopes=["User.Read.All"], clients=["graph", "exchange"])
    # Straight to Gamma's own sign-in, never through the commercial partner sign-in: Graph
    # first, then Exchange, which has no refresh token of its own for this account yet.
    assert [(c.method, c.client_id) for c in fake_msal.calls if c.tenant_id == TENANT_C] == [
        ("interactive", GRAPH_CLI_CLIENT_ID),
        ("silent", EXCHANGE_POWERSHELL_CLIENT_ID),
        ("interactive", EXCHANGE_POWERSHELL_CLIENT_ID),
    ]
    assert fake_msal.authority(TENANT_C) == f"https://login.microsoftonline.us/{TENANT_C}"
    with session.context_for(
        GAMMA_GOV, scopes=["User.Read.All"], clients=["graph", "exchange"]
    ) as gamma:
        assert (gamma.auth.tenant.cloud.name, gamma.auth.username) == ("USGov", GAMMA_ADMIN)
        assert gamma.graph.request("GET", "/users").request.url.host == "graph.microsoft.us"


def test_check_access_without_direct_sign_ins_leaves_the_reason_for_context_for(
    fake_msal: FakeMsal,
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    session = _signed_in(["graph"])
    session.check_access([ALPHA, BETA, GAMMA_GOV], scopes=["User.Read.All"], direct=False)
    with pytest.raises(ConsentRequired):
        session.context_for(BETA, scopes=["User.Read.All"])
    with pytest.raises(AuthError, match="Gamma is in the USGov cloud"):
        session.context_for(GAMMA_GOV, scopes=["User.Read.All"])
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]


def test_a_failed_direct_sign_in_is_what_context_for_raises(fake_msal: FakeMsal) -> None:
    fake_msal.interactive_errors[TENANT_C] = {"error": "access_denied", "error_description": "no"}
    session = _signed_in(["graph"])
    session.check_access([GAMMA_GOV], scopes=["User.Read.All"])
    with pytest.raises(AuthError, match="access_denied"):
        session.context_for(GAMMA_GOV, scopes=["User.Read.All"])
