"""Integration tests for PartnerSession: real azure-auth over a fake msal and transport."""

from __future__ import annotations

import logging
import os
from itertools import count
from pathlib import Path
from typing import Any

import httpx
import pytest
from azure_auth import AuthContext, AuthError, ConsentRequired, TenantNotFound
from azure_auth.auth import cache as azure_auth_cache
from azure_auth.constants import (
    AZURE_POWERSHELL_CLIENT_ID,
    EXCHANGE_POWERSHELL_CLIENT_ID,
    GRAPH_CLI_CLIENT_ID,
)

from m365_gdap_runner.cache_reminder import reminder_path
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
TRANSPORT = httpx.MockTransport(fake_services)


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
        assert alpha.auth.home is session.auth
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]
    assert fake_msal.silent_calls(TENANT_B) == [
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/.default",)),
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
        session.check_access([ALPHA, BETA, gamma], scopes=["User.Read.All"])
    # Nothing after the last check: the lines that follow say how it went.
    assert [r.getMessage() for r in caplog.records if r.msg.startswith("Checked")] == lines


def test_check_access_signs_in_directly_where_the_partner_sign_in_fails(
    fake_msal: FakeMsal,
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    beta_with_admin = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)
    session = _signed_in(["graph"])
    session.check_access([ALPHA, beta_with_admin], scopes=["User.Read.All"])
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
    fake_msal.account_errors[(TENANT_B, USER)] = CONSENT_REQUIRED
    beta_with_admin = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN)
    session = PartnerSession(make_settings(cache="disk"), transport=TRANSPORT)
    session.login(["graph"])
    with caplog.at_level(logging.INFO, logger="m365_gdap_runner.session"):
        session.check_access([beta_with_admin], scopes=["User.Read.All"])
    assert (
        f"Signing in to beta corp directly as {BETA_ADMIN}: the partner sign-in cannot reach it"
        in caplog.text
    )
    # Only the partner sign-in opened a window; beta corp's own came from the cache.
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]
    with session.context_for(beta_with_admin, scopes=["User.Read.All"]) as beta:
        assert (beta.auth.home is beta.auth, beta.auth.username) == (True, BETA_ADMIN)


def test_check_access_never_tries_the_partner_sign_in_where_gdap_is_false(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture
) -> None:
    # The partner sign-in would reach beta corp; the tenants file says not to try.
    beta = Tenant(TENANT_B, "beta corp", "Commercial", BETA_ADMIN, gdap=False)
    session = _signed_in(["graph"])
    with caplog.at_level(logging.INFO, logger="m365_gdap_runner.session"):
        session.check_access([ALPHA, beta], scopes=["User.Read.All"])
    assert fake_msal.silent_calls(TENANT_B) == []
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]
    assert (
        f"Signing in to beta corp directly as {BETA_ADMIN}: its m365_gdap is not TRUE"
        in caplog.text
    )
    with session.context_for(beta, scopes=["User.Read.All"]) as context:
        assert (context.auth.home is context.auth, context.auth.username) == (True, BETA_ADMIN)


@pytest.mark.regression
@pytest.mark.parametrize(
    ("tenant", "reason"),
    [
        (BETA, "has no consent in tenant"),
        (Tenant(TENANT_C, "Gamma", "USGov"), "Gamma is in the USGov cloud"),
        (
            Tenant(TENANT_B, "beta corp", "Commercial", gdap=False),
            "beta corp's m365_gdap is not TRUE",
        ),
    ],
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


def test_a_failed_direct_sign_in_is_what_context_for_raises(fake_msal: FakeMsal) -> None:
    fake_msal.interactive_errors[TENANT_C] = {"error": "access_denied", "error_description": "no"}
    session = _signed_in(["graph"])
    session.check_access([GAMMA_GOV], scopes=["User.Read.All"])
    with pytest.raises(AuthError, match="access_denied"):
        session.context_for(GAMMA_GOV, scopes=["User.Read.All"])


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
    session.check_access([beta_with_admin], scopes=["User.Read.All"])
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
    session.check_access([beta_with_admin], scopes=["User.Read.All"])
    # The partner sign-in, then beta corp's own.
    assert timeouts == [timeout, timeout]


def test_browser_args_without_a_browser_path_is_a_config_error(fake_msal: FakeMsal) -> None:
    with pytest.raises(ConfigError, match="browser_args needs browser_path"):
        PartnerSession(make_settings(browser_args=["--inprivate"]))


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
    session.check_access([beta_with_admin], scopes=["User.Read.All"])
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
