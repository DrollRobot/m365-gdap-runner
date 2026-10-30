"""Integration tests for PartnerSession: real azure-auth over a fake msal and transport."""

from __future__ import annotations

import logging
import os
import threading
import webbrowser
from itertools import count
from pathlib import Path
from typing import Any

import azure_auth
import httpx2
import pytest
from azure_auth import AuthContext, AuthError, ConsentRequired, NetworkError, TenantNotFound
from azure_auth.auth import cache as azure_auth_cache
from azure_auth.constants import (
    AZURE_POWERSHELL_CLIENT_ID,
    EXCHANGE_POWERSHELL_CLIENT_ID,
    GRAPH_CLI_CLIENT_ID,
)

from m365_gdap_runner.cache_reminder import reminder_path
from m365_gdap_runner.config import ConfigError
from m365_gdap_runner.runner import TenantLogFilter
from m365_gdap_runner.session import PartnerSession
from m365_gdap_runner.tenants import Tenant
from tests._gdap_fakes import (
    CONSENT_REQUIRED,
    DEVICE_NOT_TRUSTED,
    GAMMA_ADMIN,
    MFA_TOO_OLD,
    PARTNER_TENANT,
    REFRESH_REFUSED,
    TENANT_A,
    TENANT_B,
    TENANT_C,
    USER,
    FakeApp,
    FakeMsal,
    bearer_tenant,
    customers_body,
    fake_services,
    make_settings,
)

pytestmark = pytest.mark.integration

PARTNER_SCOPE = "https://graph.microsoft.com/DelegatedAdminRelationship.Read.All"
# What the log line adds for a direct sign-in that shows the consent screen.
TICK = '; tick "Consent on behalf of your organization" in the consent screen'
TRANSPORT = httpx2.MockTransport(fake_services)


def test_login_signs_in_once_per_client_id(fake_msal: FakeMsal) -> None:
    session = PartnerSession(make_settings(), transport=TRANSPORT)
    session.login(["graph", "exchange", "ipps", "arm"])
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


def test_login_can_ask_for_the_password_again_for_every_client(fake_msal: FakeMsal) -> None:
    session = _signed_in(["graph", "exchange"])
    fake_msal.calls.clear()
    session.login(["exchange"], reauthenticate=True)
    # A window per client, whatever the cache holds, and no request after it.
    assert [(c.method, c.client_id, c.prompt) for c in fake_msal.calls] == [
        ("interactive", GRAPH_CLI_CLIENT_ID, "login"),
        ("interactive", EXCHANGE_POWERSHELL_CLIENT_ID, "login"),
    ]


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
    seen: list[httpx2.Request] = []

    def handler(request: httpx2.Request) -> httpx2.Response:
        seen.append(request)
        body = customers_body(
            (TENANT_B, "beta corp"), (TENANT_A, "Alpha"), (TENANT_B.upper(), "beta corp")
        )
        body["value"].append({"tenantId": TENANT_C})
        return httpx2.Response(200, json=body)

    session = PartnerSession(make_settings(), transport=httpx2.MockTransport(handler))
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


def test_context_for_uses_the_partner_sign_in_and_probes_the_customer(fake_msal: FakeMsal) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    session = PartnerSession(make_settings(), transport=TRANSPORT)
    session.login(["graph"])

    with session.context_for(Tenant(TENANT_A, "Alpha"), scopes=["User.Read.All"]) as alpha:
        assert alpha.auth.home is session.auth
        assert alpha.auth.tenant.id == TENANT_A
        alpha.probe()
        assert bearer_tenant(alpha.graph.request("GET", "/users").request) == TENANT_A

    with (
        session.context_for(Tenant(TENANT_B, "beta corp"), scopes=["User.Read.All"]) as beta,
        pytest.raises(ConsentRequired) as info,
    ):
        beta.probe()
    assert info.value.tenant_id == TENANT_B
    assert info.value.missing == ("User.Read.All",)
    # .default, cached then fresh, shows nothing is granted, so User.Read.All is never asked for.
    assert fake_msal.silent_calls(TENANT_B) == [
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/.default",)),
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/.default",)),
    ]
    assert fake_msal.interactive_calls == 1


def test_partner_tenant_id_is_the_guid(fake_msal: FakeMsal) -> None:
    session = PartnerSession(make_settings(partner_tenant_id="partner.example"))
    assert session.partner_tenant_id != "partner.example"
    assert session.auth.tenant.domain == "partner.example"


@pytest.mark.parametrize(
    ("field", "value", "message"),
    [
        ("cache", "plaintext", "cache must be one of"),
        ("cloud", "Mars", "Unknown cloud 'Mars'"),
        ("browser_timeout", 0.0, "browser_timeout must be more than 0 seconds, not 0.0"),
    ],
)
def test_bad_cache_cloud_or_browser_timeout_is_a_config_error(
    fake_msal: FakeMsal, field: str, value: object, message: str
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
        NetworkError("GET https://login.microsoftonline.com failed: ConnectError: no route"),
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


# azure-auth's own lookup, before fake_msal replaces it.
REAL_LOOKUP = azure_auth.Tenant.lookup


@pytest.mark.regression
def test_cloud_of_is_blank_when_no_sign_in_host_answers(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch
) -> None:
    # azure-auth 0.10.0 raises NetworkError here, where 0.9.1 raised AuthError.
    def unreachable(request: httpx2.Request) -> httpx2.Response:
        raise httpx2.ConnectError("no route", request=request)

    session = PartnerSession(make_settings(), transport=httpx2.MockTransport(unreachable))
    monkeypatch.setattr("azure_auth.Tenant.lookup", REAL_LOOKUP)
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


def _tagged(caplog: pytest.LogCaptureFixture) -> list[tuple[str, str]]:
    """Each session record caplog kept, as its tenant tag and message.

    Needs :class:`TenantLogFilter` on ``caplog.handler``, which sets the tag.
    """
    return [
        (r.__dict__["tenant"], r.getMessage())
        for r in caplog.records
        if r.name == "m365_gdap_runner.session"
    ]


def test_check_access_reaches_gdap_tenants_through_the_partner_sign_in(
    fake_msal: FakeMsal,
) -> None:
    session = _signed_in(["graph"])
    list(session.check_access([ALPHA, BETA], scopes=["User.Read.All"], threads=2))
    with session.context_for(ALPHA, scopes=["User.Read.All"]) as alpha:
        assert alpha.auth.home is session.auth
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]
    # The consent check, the token, then the token again from Entra rather than the cache.
    assert fake_msal.silent_calls(TENANT_B) == [
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/.default",)),
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/User.Read.All",)),
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/User.Read.All",)),
    ]
    assert fake_msal.refreshed(TENANT_B) == [
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/User.Read.All",)),
    ]


@pytest.mark.parametrize(
    ("seconds_per_check", "lines"),
    [
        (60, ["Checked 1 of 2 tenant(s), about 1 min left"]),
        (40, ["Checked 1 of 2 tenant(s), under a minute left"]),
        (10, []),
    ],
    ids=["minutes-left", "under-a-minute", "too-soon"],
)
def test_check_access_logs_progress_every_30_seconds(
    fake_msal: FakeMsal,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
    seconds_per_check: int,
    lines: list[str],
) -> None:
    ticks = count(0, seconds_per_check)
    monkeypatch.setattr("m365_gdap_runner.session.monotonic", lambda: float(next(ticks)))
    session = _signed_in(["graph"])
    with caplog.at_level(logging.INFO, logger="m365_gdap_runner.session"):
        # Gamma is in another cloud, so it gets no silent check to count, and with no
        # m365_username, no sign-in of its own either.
        gamma = Tenant(TENANT_C, "Gamma", "USGov")
        list(session.check_access([ALPHA, BETA, gamma], scopes=["User.Read.All"]))
    # Nothing after the last check: the lines that follow say how it went.
    assert [r.getMessage() for r in caplog.records if r.msg.startswith("Checked")] == lines


def test_check_access_signs_in_directly_where_the_partner_sign_in_fails(
    fake_msal: FakeMsal,
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    beta_with_admin = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)
    session = _signed_in(["graph"])
    list(session.check_access([ALPHA, beta_with_admin], scopes=["User.Read.All"]))
    with session.context_for(beta_with_admin, scopes=["User.Read.All"]) as beta:
        assert beta.auth.home is beta.auth
        assert (beta.auth.tenant.id, beta.auth.username) == (TENANT_B, BETA_ADMIN)
        assert bearer_tenant(beta.graph.request("GET", "/users").request) == TENANT_B
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]


@pytest.mark.acceptance
def test_check_access_tries_the_partner_sign_in_then_a_cached_direct_one(
    fake_msal: FakeMsal,
    cache_files: list[tuple[str, Path | None]],
    caplog: pytest.LogCaptureFixture,
) -> None:
    # beta corp's own account signed in on an earlier run, and the disk cache kept it.
    fake_msal.sign_in(GRAPH_CLI_CLIENT_ID, BETA_ADMIN)
    fake_msal.account_errors[(TENANT_B, USER)] = DEVICE_NOT_TRUSTED
    beta_with_admin = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)
    session = PartnerSession(make_settings(cache="disk"), transport=TRANSPORT)
    session.login(["graph"])
    with caplog.at_level(logging.INFO, logger="m365_gdap_runner.session"):
        list(session.check_access([beta_with_admin], scopes=["User.Read.All"]))
    assert (
        "AADSTS530004: AcceptCompliantDevice setting isn't configured for this organization; "
        f"direct: signing in as {BETA_ADMIN}"
    ) in caplog.text
    # Only the partner sign-in opened a window; beta corp's own came from the cache.
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]
    with session.context_for(beta_with_admin, scopes=["User.Read.All"]) as beta:
        assert (beta.auth.home is beta.auth, beta.auth.username) == (True, BETA_ADMIN)


@pytest.mark.acceptance
def test_a_tenant_lacking_consent_gets_the_consent_screen_in_its_own_sign_in(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture
) -> None:
    # beta corp's administrator consented for themselves alone on an earlier run, so their
    # cached sign-in holds every scope while the partner sign-in has none of them. Entra
    # shows the consent screen again only when asked to.
    fake_msal.sign_in(GRAPH_CLI_CLIENT_ID, BETA_ADMIN)
    fake_msal.account_errors[(TENANT_B, USER)] = CONSENT_REQUIRED
    beta_with_admin = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)
    session = _signed_in(["graph", "exchange"])
    with caplog.at_level(logging.INFO, logger="m365_gdap_runner.session"):
        list(
            session.check_access(
                [beta_with_admin], scopes=["User.Read.All"], clients=["graph", "exchange"]
            )
        )
    assert f"partner: no Graph consent; direct: signing in as {BETA_ADMIN}{TICK}" in caplog.text
    # Graph's window opens for the consent screen, cached sign-in or not; Exchange has no
    # consent screen, and its window opens only because this account has no sign-in for it.
    assert fake_msal.interactive_prompts() == [
        (PARTNER_TENANT, None),
        (PARTNER_TENANT, None),
        (TENANT_B, "consent"),
        (TENANT_B, None),
    ]
    with session.context_for(beta_with_admin, scopes=["User.Read.All"]) as beta:
        assert (beta.auth.home is beta.auth, beta.auth.username) == (True, BETA_ADMIN)


def test_check_access_never_tries_the_partner_sign_in_where_gdap_is_false(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture
) -> None:
    # The partner sign-in would reach beta corp; the tenants file says not to try.
    beta = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN, gdap=False)
    session = _signed_in(["graph"])
    with caplog.at_level(logging.INFO, logger="m365_gdap_runner.session"):
        list(session.check_access([ALPHA, beta], scopes=["User.Read.All"]))
    assert fake_msal.silent_calls(TENANT_B, USER) == []
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]
    assert f"partner: m365_gdap is FALSE; direct: signing in as {BETA_ADMIN}" in caplog.text
    with session.context_for(beta, scopes=["User.Read.All"]) as context:
        assert (context.auth.home is context.auth, context.auth.username) == (True, BETA_ADMIN)


@pytest.mark.acceptance
def test_the_sign_in_checks_name_their_tenant_in_the_log(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture
) -> None:
    fake_msal.account_errors[(TENANT_B, USER)] = CONSENT_REQUIRED
    beta_with_admin = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)
    caplog.handler.addFilter(TenantLogFilter())
    session = _signed_in(["graph"])
    with caplog.at_level(logging.DEBUG, logger="m365_gdap_runner.session"):
        list(session.check_access([beta_with_admin], scopes=["User.Read.All"]))
    tagged = _tagged(caplog)
    # The partner check runs on a worker thread; the direct sign-in on this one.
    assert ("beta corp", "The partner sign-in cannot reach the tenant") in tagged
    assert (
        "beta corp",
        f"partner: no Graph consent; direct: signing in as {BETA_ADMIN}{TICK}",
    ) in tagged
    # A line about no one tenant names none.
    assert ("-", "Checking access to 1 tenant(s)") in tagged


@pytest.mark.regression
@pytest.mark.parametrize(
    ("tenant", "reason"),
    [
        (BETA, "no Graph consent"),
        (Tenant(TENANT_C, "Gamma", "USGov"), "tenant is in the USGov cloud"),
        (Tenant(TENANT_B, "beta corp", "Commercial", gdap=False), "m365_gdap is FALSE"),
    ],
)
def test_check_access_never_signs_in_directly_as_the_partner_account(
    fake_msal: FakeMsal, tenant: Tenant, reason: str
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    session = _signed_in(["graph"])
    list(session.check_access([tenant], scopes=["User.Read.All"]))
    with pytest.raises(AuthError) as info:
        session.context_for(tenant, scopes=["User.Read.All"])
    assert str(info.value) == f"partner: {reason}; direct: no m365_username"
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]


def test_a_tenant_with_no_account_of_its_own_adds_no_warning_to_the_check(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture
) -> None:
    # The run's error line for beta corp says why; a warning would say it twice.
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    session = _signed_in(["graph"])
    with caplog.at_level(logging.INFO, logger="m365_gdap_runner.session"):
        list(session.check_access([BETA], scopes=["User.Read.All"]))
    assert [r.getMessage() for r in caplog.records if r.levelno >= logging.WARNING] == []


@pytest.mark.acceptance
def test_an_excessively_long_partner_failure_is_cut_but_the_direct_one_still_shows(
    fake_msal: FakeMsal,
) -> None:
    fake_msal.tenant_errors[TENANT_B] = {
        "error": "invalid_grant",
        "error_description": (
            "AADSTS50105: Your administrator has configured the application Microsoft Graph "
            "Command Line Tools ('14d82eec-204b-4c2f-b7e8-296a70dab67e') to block users unless "
            "they are specifically granted ('assigned') access to the application.  The signed "
            "in user '{EUII Hidden}' is blocked because they are not a direct member of a group "
            "with access, nor had access directly assigned by an administrator. Please contact "
            "your administrator to assign access to this application.\r\nTrace ID: 0\r\n"
            "Correlation ID: 0\r\nTimestamp: 2026-10-22 01:51:05Z"
        ),
        "error_codes": [50105],
    }
    session = _signed_in(["graph"])
    list(session.check_access([BETA], scopes=["User.Read.All"]))
    with pytest.raises(AuthError) as info:
        session.context_for(BETA, scopes=["User.Read.All"])
    assert str(info.value) == (
        "partner: invalid_grant: AADSTS50105: Your administrator has configured the application "
        "Microsoft Graph Command Line Tools ('14d82eec-204b-4c2f-b7e8-296a70dab67e') to block "
        "users unless they are specifically granted ('assigned') access to the application; "
        "direct: no m365_username"
    )


def test_check_access_signs_in_to_another_cloud_once_per_client_as_its_account(
    fake_msal: FakeMsal,
) -> None:
    session = _signed_in(["graph", "exchange"])
    list(session.check_access([GAMMA_GOV], scopes=["User.Read.All"], clients=["graph", "exchange"]))
    # Straight to Gamma's own sign-in, never through the commercial partner sign-in: Graph
    # first, then Exchange, which has no refresh token of its own for this account yet. After
    # each window, the new sign-in is asked for a new token, like a cached one would be.
    assert [(c.method, c.client_id) for c in fake_msal.calls if c.tenant_id == TENANT_C] == [
        ("interactive", GRAPH_CLI_CLIENT_ID),
        ("silent", GRAPH_CLI_CLIENT_ID),
        ("silent", EXCHANGE_POWERSHELL_CLIENT_ID),
        ("interactive", EXCHANGE_POWERSHELL_CLIENT_ID),
        ("silent", EXCHANGE_POWERSHELL_CLIENT_ID),
    ]
    assert [client for client, _ in fake_msal.refreshed(TENANT_C)] == [
        GRAPH_CLI_CLIENT_ID,
        EXCHANGE_POWERSHELL_CLIENT_ID,
    ]
    assert fake_msal.authority(TENANT_C) == f"https://login.microsoftonline.us/{TENANT_C}"
    with session.context_for(
        GAMMA_GOV, scopes=["User.Read.All"], clients=["graph", "exchange"]
    ) as gamma:
        assert (gamma.auth.tenant.cloud.name, gamma.auth.username) == ("USGov", GAMMA_ADMIN)
        assert gamma.graph.request("GET", "/users").request.url.host == "graph.microsoft.us"


@pytest.mark.acceptance
@pytest.mark.regression
def test_a_cached_direct_sign_in_entra_will_not_renew_opens_the_window_in_the_check(
    fake_msal: FakeMsal,
) -> None:
    # beta corp's own account signed in on an earlier run, and its token is still in the
    # cache, but the tenant's sign-in frequency policy has since started refusing new ones.
    # Before, the check took the cached token and the script failed an hour later; the window
    # opens here instead, where it may.
    fake_msal.sign_in(GRAPH_CLI_CLIENT_ID, BETA_ADMIN)
    fake_msal.refresh_errors[(TENANT_B, BETA_ADMIN)] = REFRESH_REFUSED
    beta = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN, gdap=False)
    session = _signed_in(["graph"])
    assert list(session.check_access([beta], scopes=["User.Read.All"])) == [beta]
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]
    with session.context_for(beta, scopes=["User.Read.All"]) as context:
        assert context.auth.username == BETA_ADMIN


BETA_WITH_ADMIN = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)


@pytest.mark.acceptance
@pytest.mark.parametrize(
    ("refusal", "code"),
    [(REFRESH_REFUSED, "70043"), (MFA_TOO_OLD, "50078")],
    ids=["sign-in", "mfa"],
)
def test_a_customer_that_finds_the_partner_sign_in_too_old_gets_a_new_one(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture, refusal: dict[str, Any], code: str
) -> None:
    fake_msal.refresh_errors[(TENANT_B, USER)] = refusal
    caplog.handler.addFilter(TenantLogFilter())
    session = _signed_in(["graph"])
    with caplog.at_level(logging.INFO, logger="m365_gdap_runner.session"):
        list(session.check_access([BETA_WITH_ADMIN], scopes=["User.Read.All"]))
    # The partner tenant again, asking for the password; beta corp's own account never.
    assert fake_msal.interactive_prompts() == [(PARTNER_TENANT, None), (PARTNER_TENANT, "login")]
    with session.context_for(BETA_WITH_ADMIN, scopes=["User.Read.All"]) as beta:
        assert beta.auth.home is session.auth
    (tag, message) = next((tag, message) for tag, message in _tagged(caplog) if "again" in message)
    assert tag == "beta corp"
    assert message.startswith(f"partner: invalid_grant: AADSTS{code}: ")
    assert message.endswith("; signing in to the partner tenant again")


def test_the_partner_tenant_is_signed_in_to_again_once_a_run(fake_msal: FakeMsal) -> None:
    fake_msal.refresh_errors[(TENANT_A, USER)] = MFA_TOO_OLD
    fake_msal.refresh_errors[(TENANT_B, USER)] = REFRESH_REFUSED
    session = _signed_in(["graph"])
    assert list(session.check_access([ALPHA, BETA], scopes=["User.Read.All"])) == [ALPHA, BETA]
    assert fake_msal.interactive_prompts() == [(PARTNER_TENANT, None), (PARTNER_TENANT, "login")]
    for tenant in (ALPHA, BETA):
        with session.context_for(tenant, scopes=["User.Read.All"]) as context:
            assert context.auth.home is session.auth


def test_a_refusal_no_new_sign_in_answers_goes_straight_to_the_direct_sign_in(
    fake_msal: FakeMsal,
) -> None:
    fake_msal.refresh_errors[(TENANT_B, USER)] = DEVICE_NOT_TRUSTED
    session = _signed_in(["graph"])
    list(session.check_access([BETA_WITH_ADMIN], scopes=["User.Read.All"]))
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]
    with session.context_for(BETA_WITH_ADMIN, scopes=["User.Read.All"]) as beta:
        assert beta.auth.username == BETA_ADMIN


def test_a_tenant_still_refused_after_the_new_sign_in_goes_to_its_own(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture
) -> None:
    # The new sign-in succeeds, but the tenant refuses again, for a reason it does not answer.
    fake_msal.refresh_errors[(TENANT_B, USER)] = MFA_TOO_OLD
    session = _signed_in(["graph"])
    original = session.login

    def login_then_refuse(*args: Any, **kwargs: Any) -> None:
        original(*args, **kwargs)
        fake_msal.refresh_errors[(TENANT_B, USER)] = DEVICE_NOT_TRUSTED

    session.login = login_then_refuse  # type: ignore[method-assign]
    with caplog.at_level(logging.INFO, logger="m365_gdap_runner.session"):
        list(session.check_access([BETA_WITH_ADMIN], scopes=["User.Read.All"]))
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, PARTNER_TENANT, TENANT_B]
    # The direct sign-in's line gives the second refusal.
    assert "partner: invalid_grant: AADSTS530004" in caplog.text
    with session.context_for(BETA_WITH_ADMIN, scopes=["User.Read.All"]) as beta:
        assert beta.auth.username == BETA_ADMIN


def test_a_failed_new_partner_sign_in_sends_each_tenant_to_its_own(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture
) -> None:
    fake_msal.refresh_errors[(TENANT_A, USER)] = MFA_TOO_OLD
    fake_msal.refresh_errors[(TENANT_B, USER)] = MFA_TOO_OLD
    alpha = Tenant(TENANT_A, "Alpha", "Commercial", "admin@alpha.example")
    session = _signed_in(["graph"])
    fake_msal.interactive_errors[PARTNER_TENANT] = {
        "error": "access_denied",
        "error_description": "no",
    }
    with caplog.at_level(logging.WARNING, logger="m365_gdap_runner.session"):
        list(session.check_access([alpha, BETA_WITH_ADMIN], scopes=["User.Read.All"]))
    # One try at the partner tenant, then each tenant's own sign-in.
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, PARTNER_TENANT, TENANT_A, TENANT_B]
    assert "Could not sign in to the partner tenant again: access_denied: no" in caplog.text


def test_ctrl_c_then_enter_during_the_new_partner_sign_in_skips_it(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    fake_msal.refresh_errors[(TENANT_B, USER)] = MFA_TOO_OLD
    monkeypatch.setattr("builtins.input", lambda: "")
    session = _signed_in(["graph"])
    fake_msal.interactive_interrupts.add(PARTNER_TENANT)
    assert list(session.check_access([BETA_WITH_ADMIN], scopes=["User.Read.All"])) == [
        BETA_WITH_ADMIN
    ]
    err = capsys.readouterr().err
    assert "Stopped the sign-in to the partner tenant. Press Enter to skip it" in err
    # Skipping it sends beta corp to its own sign-in.
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, PARTNER_TENANT, TENANT_B]


def test_a_failed_direct_sign_in_is_what_context_for_raises(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture
) -> None:
    fake_msal.interactive_errors[TENANT_C] = {"error": "access_denied", "error_description": "no"}
    caplog.handler.addFilter(TenantLogFilter())
    session = _signed_in(["graph"])
    list(session.check_access([GAMMA_GOV], scopes=["User.Read.All"]))
    with pytest.raises(AuthError) as info:
        session.context_for(GAMMA_GOV, scopes=["User.Read.All"])
    # azure-auth's advice after "access_denied: no." is past the length limit, so it is cut.
    assert str(info.value) == "partner: tenant is in the USGov cloud; direct: access_denied: no"
    assert ("Gamma", f"Could not sign in as {GAMMA_ADMIN}: access_denied: no") in _tagged(caplog)


@pytest.mark.acceptance
def test_ctrl_c_then_enter_during_a_direct_sign_in_skips_that_tenant(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    fake_msal.interactive_interrupts.add(TENANT_B)
    monkeypatch.setattr("builtins.input", lambda: "")
    beta = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN, gdap=False)
    session = _signed_in(["graph"])
    assert list(session.check_access([beta, GAMMA_GOV], scopes=["User.Read.All"])) == [
        beta,
        GAMMA_GOV,
    ]
    assert "Stopped the sign-in to beta corp. Press Enter to skip it" in capsys.readouterr().err
    with pytest.raises(
        AuthError, match=r"^partner: m365_gdap is FALSE; direct: skipped with Ctrl\+C$"
    ):
        session.context_for(beta, scopes=["User.Read.All"])
    # Gamma's sign-in still opened, after beta corp's.
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B, TENANT_C]


@pytest.mark.parametrize("answer", [KeyboardInterrupt, EOFError], ids=["ctrl-c", "stdin-closed"])
def test_ctrl_c_during_a_direct_sign_in_stops_the_run_unless_enter_skips_it(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, answer: type[BaseException]
) -> None:
    def ask() -> str:
        raise answer

    fake_msal.interactive_interrupts.add(TENANT_B)
    monkeypatch.setattr("builtins.input", ask)
    beta = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN, gdap=False)
    session = _signed_in(["graph"])
    with pytest.raises(KeyboardInterrupt):
        list(session.check_access([beta, GAMMA_GOV], scopes=["User.Read.All"]))
    # No sign-in window opened after beta corp's.
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]


def test_check_access_hands_out_every_tenant_once(fake_msal: FakeMsal) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    blank_gamma = Tenant(TENANT_C, "Gamma", "USGov")
    session = _signed_in(["graph"])
    handed_out = list(session.check_access([ALPHA, BETA, blank_gamma], threads=2))
    # Reached or not: context_for() then has a context or the reason for each.
    assert sorted(t.tenant_id for t in handed_out) == [TENANT_A, TENANT_B, TENANT_C]


@pytest.mark.acceptance
@pytest.mark.parametrize(
    "beta",
    [
        Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN, gdap=False),
        Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN),
    ],
    ids=["never-checked", "check-failed"],
)
def test_a_direct_sign_in_does_not_wait_for_the_other_checks(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, beta: Tenant
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    beta_signed_in = threading.Event()
    through_partner = PartnerSession._through_partner
    sign_in_directly = PartnerSession._sign_in_directly

    def alpha_waits_for_beta(
        self: PartnerSession, tenant: Tenant, *args: Any
    ) -> AuthContext | Exception:
        if tenant.tenant_id == TENANT_A:
            assert beta_signed_in.wait(5), "beta corp's sign-in waited for Alpha's check"
        return through_partner(self, tenant, *args)

    def sign_in(
        self: PartnerSession, tenant: Tenant, *args: Any, **kwargs: Any
    ) -> AuthContext | str:
        try:
            return sign_in_directly(self, tenant, *args, **kwargs)
        finally:
            beta_signed_in.set()

    monkeypatch.setattr(PartnerSession, "_through_partner", alpha_waits_for_beta)
    monkeypatch.setattr(PartnerSession, "_sign_in_directly", sign_in)
    session = _signed_in(["graph"])
    handed_out = list(session.check_access([ALPHA, beta], scopes=["User.Read.All"], threads=2))
    assert handed_out == [beta, ALPHA]
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]


def test_direct_sign_ins_open_one_at_a_time_on_the_thread_that_iterates(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch
) -> None:
    fake_msal.tenant_errors[TENANT_A] = CONSENT_REQUIRED
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    threads: list[threading.Thread] = []
    interactive = FakeApp.acquire_token_interactive

    def recording(self: FakeApp, scopes: list[str], **kwargs: Any) -> dict[str, Any]:
        threads.append(threading.current_thread())
        return interactive(self, scopes, **kwargs)

    monkeypatch.setattr(FakeApp, "acquire_token_interactive", recording)
    alpha = Tenant(TENANT_A, "Alpha", "Commercial", "admin@alpha.example")
    beta = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)
    session = _signed_in(["graph"])
    list(session.check_access([alpha, beta], scopes=["User.Read.All"], threads=2))
    # The partner sign-in, then each tenant's own, in the order their checks failed.
    assert sorted(fake_msal.interactive_tenants()[1:]) == [TENANT_A, TENANT_B]
    assert threads == [threading.current_thread()] * 3


def test_closing_check_access_opens_no_more_sign_in_windows(fake_msal: FakeMsal) -> None:
    beta = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN, gdap=False)
    session = _signed_in(["graph"])
    ready = session.check_access([beta, GAMMA_GOV], scopes=["User.Read.All"])
    assert next(ready) == beta
    ready.close()
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]


# --- the browser -----------------------------------------------------------------------


@pytest.fixture
def browsers(monkeypatch: pytest.MonkeyPatch) -> list[tuple[str, ...]]:
    """The command of every sign-in azure-auth opens in a chosen browser, in order.

    Each then signs in through the fake msal, as a sign-in in the default browser does.
    """
    opened: list[tuple[str, ...]] = []

    def browser_sign_in(
        self: AuthContext,
        app: Any,
        scopes: list[str],
        claims: str | None,
        prompt: str | None,
        browser: tuple[str, ...],
    ) -> dict[str, Any]:
        opened.append(browser)
        return dict(app.acquire_token_interactive(scopes, login_hint=self.username))

    monkeypatch.setattr("azure_auth.auth.context.AuthContext._browser_sign_in", browser_sign_in)
    return opened


@pytest.mark.parametrize(
    ("overrides", "opened"),
    [
        ({}, []),
        ({"browser_path": "~/edge"}, [(str(Path.home() / "edge"),)] * 2),
        (
            {"browser_path": "msedge", "browser_args": ["--profile-directory=Profile 2"]},
            [("msedge", "--profile-directory=Profile 2")] * 2,
        ),
    ],
    ids=["system-default", "browser-path", "browser-path-and-args"],
)
def test_the_partner_and_direct_sign_ins_use_the_browser_settings(
    fake_msal: FakeMsal,
    browsers: list[tuple[str, ...]],
    overrides: dict[str, Any],
    opened: list[tuple[str, ...]],
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    beta_with_admin = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)
    session = PartnerSession(make_settings(**overrides), transport=TRANSPORT)
    session.login(["graph"])
    list(session.check_access([beta_with_admin], scopes=["User.Read.All"]))
    # The partner sign-in, then beta corp's own.
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]
    assert browsers == opened


@pytest.mark.parametrize(
    ("overrides", "timeout"),
    [({}, 120.0), ({"browser_timeout": 45.0}, 45.0)],
    ids=["default", "browser-timeout"],
)
def test_the_partner_and_direct_sign_ins_wait_browser_timeout_seconds(
    fake_msal: FakeMsal,
    monkeypatch: pytest.MonkeyPatch,
    overrides: dict[str, Any],
    timeout: float,
) -> None:
    timeouts: list[object] = []
    interactive = FakeApp.acquire_token_interactive

    def recording(self: FakeApp, scopes: list[str], **kwargs: Any) -> dict[str, Any]:
        timeouts.append(kwargs.get("timeout"))
        return interactive(self, scopes, **kwargs)

    monkeypatch.setattr(FakeApp, "acquire_token_interactive", recording)
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    beta_with_admin = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)
    session = PartnerSession(make_settings(**overrides), transport=TRANSPORT)
    session.login(["graph"])
    list(session.check_access([beta_with_admin], scopes=["User.Read.All"]))
    # The partner sign-in, then beta corp's own.
    assert timeouts == [timeout, timeout]


def test_browser_args_without_a_browser_path_is_a_config_error(fake_msal: FakeMsal) -> None:
    with pytest.raises(ConfigError, match="browser_args needs browser_path"):
        PartnerSession(make_settings(browser_args=["--inprivate"]))


# --- the password page -----------------------------------------------------------------

PW_URL = "https://passwords.example/beta"
BETA_WITH_PW_URL = Tenant(
    TENANT_B, "beta corp", "Commercial", BETA_ADMIN, gdap=False, pw_url=PW_URL
)

PwPage = tuple[tuple[str, ...] | None, str, list[str]]


@pytest.fixture
def pw_pages(monkeypatch: pytest.MonkeyPatch, fake_msal: FakeMsal) -> list[PwPage]:
    """Every password page opened, in order.

    Each is the browser command (None for the system default), the address, and the tenant of
    every sign-in window opened before it.
    """
    opened: list[PwPage] = []

    def default(url: str, *args: Any, **kwargs: Any) -> bool:
        opened.append((None, url, fake_msal.interactive_tenants()))
        return True

    def chosen(self: webbrowser.BackgroundBrowser, url: str, *args: Any, **kwargs: Any) -> bool:
        opened.append(((self.name, *self.args), url, fake_msal.interactive_tenants()))
        return True

    monkeypatch.setattr("webbrowser.open", default)
    monkeypatch.setattr("webbrowser.BackgroundBrowser.open", chosen)
    return opened


@pytest.mark.acceptance
def test_a_direct_sign_in_opens_the_password_page_right_before_its_first_window(
    fake_msal: FakeMsal, pw_pages: list[PwPage]
) -> None:
    session = _signed_in(["graph", "exchange"])
    list(
        session.check_access(
            [BETA_WITH_PW_URL], scopes=["User.Read.All"], clients=["graph", "exchange"]
        )
    )
    # A window per client at the partner tenant, then at beta corp; the page opens once.
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, PARTNER_TENANT, TENANT_B, TENANT_B]
    assert pw_pages == [(None, PW_URL, [PARTNER_TENANT, PARTNER_TENANT])]


@pytest.mark.acceptance
def test_a_direct_sign_in_from_the_cache_opens_no_password_page(
    fake_msal: FakeMsal, pw_pages: list[PwPage]
) -> None:
    # beta corp's own account signed in on an earlier run, so its sign-in opens no window.
    fake_msal.sign_in(GRAPH_CLI_CLIENT_ID, BETA_ADMIN)
    session = _signed_in(["graph"])
    list(session.check_access([BETA_WITH_PW_URL], scopes=["User.Read.All"]))
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]
    assert pw_pages == []


def test_a_tenant_the_partner_sign_in_reaches_opens_no_password_page(
    fake_msal: FakeMsal, pw_pages: list[PwPage]
) -> None:
    alpha = Tenant(TENANT_A, "Alpha", "Commercial", "admin@alpha.example", pw_url=PW_URL)
    session = _signed_in(["graph"])
    list(session.check_access([alpha], scopes=["User.Read.All"]))
    assert pw_pages == []


@pytest.mark.parametrize(
    ("overrides", "browser"),
    [
        ({}, None),
        ({"pw_browser_path": "~/chrome"}, (str(Path.home() / "chrome"), "%s")),
        (
            {"pw_browser_path": "chrome", "pw_browser_args": ["--profile-directory=Default"]},
            ("chrome", "--profile-directory=Default", "%s"),
        ),
    ],
    ids=["system-default", "pw-browser-path", "pw-browser-path-and-args"],
)
def test_the_password_page_opens_in_the_pw_browser_settings(
    fake_msal: FakeMsal,
    pw_pages: list[PwPage],
    overrides: dict[str, Any],
    browser: tuple[str, ...] | None,
) -> None:
    session = PartnerSession(make_settings(**overrides), transport=TRANSPORT)
    session.login(["graph"])
    list(session.check_access([BETA_WITH_PW_URL], scopes=["User.Read.All"]))
    # "%s" is where webbrowser puts the address.
    assert [(command, url) for command, url, _ in pw_pages] == [(browser, PW_URL)]


def test_a_password_page_that_cannot_be_opened_is_a_warning(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    monkeypatch.setattr("webbrowser.open", lambda url, *args, **kwargs: False)
    caplog.handler.addFilter(TenantLogFilter())
    session = _signed_in(["graph"])
    with caplog.at_level(logging.WARNING, logger="m365_gdap_runner.session"):
        list(session.check_access([BETA_WITH_PW_URL], scopes=["User.Read.All"]))
    assert ("beta corp", f"Could not open the password page {PW_URL}") in _tagged(caplog)
    # The sign-in goes on without it.
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]
    with session.context_for(BETA_WITH_PW_URL, scopes=["User.Read.All"]) as beta:
        assert beta.auth.username == BETA_ADMIN


def test_pw_browser_args_without_a_pw_browser_path_is_a_config_error(fake_msal: FakeMsal) -> None:
    with pytest.raises(ConfigError, match="pw_browser_args needs pw_browser_path"):
        PartnerSession(make_settings(pw_browser_args=["--profile-directory=Default"]))


# --- the disk cache --------------------------------------------------------------------


@pytest.fixture
def cache_files(monkeypatch: pytest.MonkeyPatch) -> list[tuple[str, Path | None]]:
    """The cache kind and file every ``AuthContext`` built during the test asks for, in order.

    Each gets an empty cache in memory instead, so a disk cache needs no encryption here. The
    fake msal keeps every sign-in in one place, as a disk cache shared by every context does.
    """
    asked: list[tuple[str, Path | None]] = []

    def build_cache(kind: str, cache_path: str | Path | None = None) -> object:
        asked.append((kind, None if cache_path is None else Path(cache_path)))
        return azure_auth_cache.build_cache("memory")

    monkeypatch.setattr("azure_auth.auth.context.build_cache", build_cache)
    return asked


def _points_at(link: Path, folder: Path) -> bool:
    """Say whether a reminder opens ``folder``: a ``.url`` file on Windows, else a link."""
    if os.name == "nt":
        return f"URL={folder.resolve().as_uri()}/" in link.read_text(encoding="utf-8")
    return link.resolve() == folder.resolve()


@pytest.mark.parametrize(
    ("overrides", "asked"),
    [
        ({}, ("memory", None)),
        ({"cache": "disk"}, ("disk", None)),
        ({"cache": "disk", "cache_path": "~/tokens.bin"}, ("disk", Path.home() / "tokens.bin")),
    ],
    ids=["memory-by-default", "disk-in-azure-auths-file", "disk-in-cache-path"],
)
def test_the_partner_and_direct_sign_ins_use_the_cache_settings(
    fake_msal: FakeMsal,
    cache_files: list[tuple[str, Path | None]],
    overrides: dict[str, str],
    asked: tuple[str, Path | None],
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    beta_with_admin = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)
    session = PartnerSession(make_settings(**overrides), transport=TRANSPORT)
    session.login(["graph"])
    list(session.check_access([beta_with_admin], scopes=["User.Read.All"]))
    # The partner sign-in, then beta corp's own.
    assert cache_files == [asked, asked]


def test_a_disk_cache_leaves_the_reminder_pointing_at_its_folder(
    fake_msal: FakeMsal, cache_files: list[tuple[str, Path | None]], tmp_path: Path
) -> None:
    cache_file = tmp_path / "cache" / "tokens.bin"
    desktop = tmp_path / "desktop"
    PartnerSession(
        make_settings(cache="disk", cache_path=str(cache_file), cache_reminder_path=str(desktop))
    )
    assert _points_at(reminder_path(desktop), cache_file.parent)


def test_the_reminder_finds_azure_auths_file_without_a_cache_path(
    fake_msal: FakeMsal, cache_files: list[tuple[str, Path | None]], tmp_path: Path
) -> None:
    desktop = tmp_path / "desktop"
    PartnerSession(make_settings(cache="disk", cache_reminder_path=str(desktop)))
    assert _points_at(reminder_path(desktop), azure_auth_cache.default_cache_path().parent)


def test_the_memory_cache_leaves_no_reminder(
    fake_msal: FakeMsal, cache_files: list[tuple[str, Path | None]], tmp_path: Path
) -> None:
    desktop = tmp_path / "desktop"
    PartnerSession(make_settings(cache_reminder_path=str(desktop)))
    assert not desktop.exists()


def test_a_reminder_that_cannot_be_placed_is_a_warning(
    fake_msal: FakeMsal,
    cache_files: list[tuple[str, Path | None]],
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
) -> None:
    desktop = tmp_path / "desktop"
    # A folder in the reminder's place is someone's, and is left alone.
    reminder_path(desktop).mkdir(parents=True)
    with caplog.at_level(logging.WARNING, logger="m365_gdap_runner.session"):
        session = PartnerSession(
            make_settings(
                cache="disk",
                cache_path=str(tmp_path / "tokens.bin"),
                cache_reminder_path=str(desktop),
            )
        )
    assert session.partner_tenant_id == PARTNER_TENANT
    assert f"Could not put the token cache reminder in {desktop}" in caplog.text


# --- writing the disk cache ------------------------------------------------------------


@pytest.fixture
def flushes(monkeypatch: pytest.MonkeyPatch) -> list[str | None]:
    """The username of every context whose cache the session wrote, in order.

    azure-auth's ``flush_cache()`` does the writing; what is recorded here is each call.
    """
    flushed: list[str | None] = []
    monkeypatch.setattr(AuthContext, "flush_cache", lambda self: flushed.append(self.username))
    return flushed


def test_each_partner_sign_in_is_written_at_once(
    fake_msal: FakeMsal, flushes: list[str | None]
) -> None:
    # A sign-in that opened a window is worth keeping even if the run is killed later, so it
    # is written as soon as it is done, once per client.
    session = PartnerSession(make_settings(), transport=TRANSPORT)
    session.login(["graph", "exchange"])

    assert flushes == [USER, USER]


def test_a_direct_sign_in_is_written_at_once(
    fake_msal: FakeMsal, flushes: list[str | None]
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    beta = Tenant(TENANT_B, "beta corp", "Commercial", "admin@beta.example")
    session = _signed_in(["graph"])
    del flushes[:]

    list(session.check_access([beta], scopes=["User.Read.All"]))

    assert "admin@beta.example" in flushes


# --- A tuple in SCOPES: any of its scopes will do ----------------------------------------

EITHER = ("Application.Read.All", "Application.ReadWrite.All")
READ = "https://graph.microsoft.com/Application.Read.All"
READ_WRITE = "https://graph.microsoft.com/Application.ReadWrite.All"
DEFAULT = "https://graph.microsoft.com/.default"


def _interactive_scopes(fake_msal: FakeMsal, tenant_id: str) -> list[tuple[str, ...]]:
    """The scopes of every sign-in window a tenant opened, in order."""
    return [
        call.scopes
        for call in fake_msal.calls
        if call.method == "interactive" and call.tenant_id.lower() == tenant_id.lower()
    ]


@pytest.mark.acceptance
def test_a_tenant_holding_the_second_scope_in_a_tuple_is_asked_for_that_one(
    fake_msal: FakeMsal,
) -> None:
    fake_msal.granted[TENANT_B] = "Application.ReadWrite.All"
    session = _signed_in(["graph"])
    assert list(session.check_access([BETA], scopes=[EITHER])) == [BETA]
    # The grant is read once, for the choice and azure-auth's own check; then the token for
    # the scope the tenant holds, and that token again from Entra.
    assert fake_msal.silent_calls(TENANT_B) == [
        (GRAPH_CLI_CLIENT_ID, (DEFAULT,)),
        (GRAPH_CLI_CLIENT_ID, (READ_WRITE,)),
        (GRAPH_CLI_CLIENT_ID, (READ_WRITE,)),
    ]
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]
    with session.context_for(BETA, scopes=[EITHER]) as beta:
        assert beta.auth.home is session.auth
        assert beta.graph.scopes == (READ_WRITE,)


def test_a_tenant_holding_both_scopes_in_a_tuple_is_asked_for_the_first(
    fake_msal: FakeMsal,
) -> None:
    fake_msal.granted[TENANT_B] = "User.Read.All Application.ReadWrite.All Application.Read.All"
    session = _signed_in(["graph"])
    assert list(session.check_access([BETA], scopes=["User.Read.All", EITHER])) == [BETA]
    with session.context_for(BETA, scopes=["User.Read.All", EITHER]) as beta:
        assert beta.graph.scopes == ("https://graph.microsoft.com/User.Read.All", READ)


@pytest.mark.acceptance
def test_a_tenant_holding_neither_scope_in_a_tuple_consents_to_the_first_in_its_own_sign_in(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture
) -> None:
    fake_msal.granted[TENANT_B] = "User.Read.All"
    session = _signed_in(["graph"])
    with caplog.at_level(logging.INFO, logger="m365_gdap_runner.session"):
        list(session.check_access([BETA_WITH_ADMIN], scopes=[EITHER]))
    assert (
        "partner: missing 1 of 1 Graph scopes, can't self-grant; "
        f"direct: signing in as {BETA_ADMIN}{TICK}" in caplog.text
    )
    assert fake_msal.interactive_prompts() == [(PARTNER_TENANT, None), (TENANT_B, "consent")]
    assert _interactive_scopes(fake_msal, TENANT_B) == [(READ,)]
    with session.context_for(BETA_WITH_ADMIN, scopes=[EITHER]) as beta:
        assert beta.auth.home is beta.auth
        assert beta.graph.scopes == (READ,)


def test_a_tenant_whose_grant_cannot_be_read_is_asked_for_the_first_scope_in_a_tuple(
    fake_msal: FakeMsal,
) -> None:
    # Alpha's tokens carry no scopes, so nobody can tell what it has granted, and the check
    # asks for the first scope as it would for a plain entry.
    session = _signed_in(["graph"])
    assert list(session.check_access([ALPHA], scopes=[EITHER])) == [ALPHA]
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]
    with session.context_for(ALPHA, scopes=[EITHER]) as alpha:
        assert alpha.graph.scopes == (READ,)


def test_a_tenant_never_checked_is_asked_for_the_first_scope_in_a_tuple(
    fake_msal: FakeMsal,
) -> None:
    fake_msal.granted[TENANT_B] = "Application.ReadWrite.All"
    session = _signed_in(["graph"])
    with session.context_for(BETA, scopes=[EITHER]) as beta:
        assert beta.graph.scopes == (READ,)


def test_a_script_without_a_tuple_never_reads_the_grant_first(fake_msal: FakeMsal) -> None:
    fake_msal.granted[TENANT_B] = "User.Read.All"
    session = _signed_in(["graph"])
    assert list(session.check_access([BETA], scopes=["User.Read.All"])) == [BETA]
    assert fake_msal.silent_calls(TENANT_B)[0] == (GRAPH_CLI_CLIENT_ID, (DEFAULT,))
    assert fake_msal.silent_calls(TENANT_B)[1] != (GRAPH_CLI_CLIENT_ID, (DEFAULT,))
