"""Unit tests for TenantContext in context.py."""

from __future__ import annotations

import azure_auth
import pytest
from azure_auth import AuthContext, InteractionRequired
from azure_auth.clouds import COMMERCIAL
from azure_auth.constants import EXCHANGE_POWERSHELL_CLIENT_ID, GRAPH_CLI_CLIENT_ID
from azure_auth.sync import ArmClient, ExchangeClient, GraphClient, IppsClient

from m365_gdap_runner.context import CLIENT_KINDS, ClientNotDeclared, TenantContext, build_client
from m365_gdap_runner.tenants import Tenant
from tests._gdap_fakes import REFRESH_REFUSED, TENANT_A, USER, FakeMsal

pytestmark = pytest.mark.unit

ALPHA = Tenant(TENANT_A, "Alpha")


def _auth() -> AuthContext:
    # A GUID with its cloud needs no lookup, so no fake discovery is required here.
    return AuthContext(azure_auth.Tenant(TENANT_A, COMMERCIAL), username=USER)


def test_declared_clients_are_built_with_the_right_types() -> None:
    with TenantContext(ALPHA, _auth(), scopes=["User.Read.All"], clients=CLIENT_KINDS) as ctx:
        assert ctx.clients == ("graph", "exchange", "ipps", "arm")
        assert isinstance(ctx.graph, GraphClient)
        assert isinstance(ctx.exchange, ExchangeClient)
        assert isinstance(ctx.ipps, IppsClient)
        assert isinstance(ctx.arm, ArmClient)
        assert ctx.graph.scopes == ("https://graph.microsoft.com/User.Read.All",)
        assert ctx.graph.client_id == GRAPH_CLI_CLIENT_ID
        assert ctx.exchange.client_id == EXCHANGE_POWERSHELL_CLIENT_ID
        assert ctx.graph.auth is ctx.auth


def test_undeclared_client_raises_naming_the_kind() -> None:
    ctx = TenantContext(ALPHA, _auth(), scopes=["User.Read.All"])
    assert ctx.clients == ("graph",)
    with pytest.raises(ClientNotDeclared, match=r"ctx\.exchange is not available") as info:
        _ = ctx.exchange
    assert info.value.kind == "exchange"
    assert "CLIENTS" in str(info.value)
    ctx.close()


def test_close_closes_every_client() -> None:
    ctx = TenantContext(ALPHA, _auth(), clients=["exchange", "arm"])
    exchange, arm = ctx.exchange, ctx.arm
    ctx.close()
    assert exchange._http.is_closed
    assert arm._http.is_closed


def test_close_writes_the_tenants_tokens_to_the_disk_cache(monkeypatch: pytest.MonkeyPatch) -> None:
    # The disk cache is written once per tenant, when its context closes, rather than once per
    # token. azure-auth's flush_cache() does the writing; what is checked here is the call.
    flushed: list[AuthContext] = []
    monkeypatch.setattr(AuthContext, "flush_cache", lambda self: flushed.append(self))
    auth = _auth()

    TenantContext(ALPHA, auth, clients=["graph"]).close()

    assert flushed == [auth]


def test_build_client_rejects_unknown_kind() -> None:
    with pytest.raises(ValueError, match="Unknown client kind 'teams'"):
        build_client("teams", _auth())


def test_probe_requests_a_token_for_every_declared_client(fake_msal: FakeMsal) -> None:
    fake_msal.sign_in(GRAPH_CLI_CLIENT_ID)
    fake_msal.sign_in(EXCHANGE_POWERSHELL_CLIENT_ID)
    auth = AuthContext("partner.example", username=USER).get_tenant_context(TENANT_A)
    with TenantContext(ALPHA, auth, scopes=["User.Read.All"], clients=["graph", "exchange"]) as ctx:
        ctx.probe()
    # The Graph client first asks .default what the tenant has granted. Each client then asks
    # Entra for a new token, after the one the cache answered with.
    assert fake_msal.silent_calls(TENANT_A) == [
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/.default",)),
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/User.Read.All",)),
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/User.Read.All",)),
        (EXCHANGE_POWERSHELL_CLIENT_ID, ("https://outlook.office365.com/.default",)),
        (EXCHANGE_POWERSHELL_CLIENT_ID, ("https://outlook.office365.com/.default",)),
    ]
    assert fake_msal.refreshed(TENANT_A) == [
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/User.Read.All",)),
        (EXCHANGE_POWERSHELL_CLIENT_ID, ("https://outlook.office365.com/.default",)),
    ]
    assert fake_msal.interactive_calls == 0


@pytest.mark.acceptance
def test_probe_fails_when_entra_will_not_renew_a_cached_sign_in(fake_msal: FakeMsal) -> None:
    # The cache still answers for Alpha; Entra, asked for a new token, refuses.
    fake_msal.sign_in(GRAPH_CLI_CLIENT_ID)
    fake_msal.refresh_errors[(TENANT_A, USER)] = REFRESH_REFUSED
    auth = AuthContext("partner.example", username=USER).get_tenant_context(TENANT_A)
    with (
        TenantContext(ALPHA, auth, scopes=["User.Read.All"]) as ctx,
        pytest.raises(InteractionRequired, match="AADSTS70043"),
    ):
        ctx.probe()
    assert fake_msal.interactive_calls == 0


def test_probe_through_the_partner_sign_in_never_prompts(fake_msal: FakeMsal) -> None:
    auth = AuthContext("partner.example", username=USER).get_tenant_context(TENANT_A)
    with (
        TenantContext(ALPHA, auth, scopes=["User.Read.All"]) as ctx,
        pytest.raises(InteractionRequired) as info,
    ):
        ctx.probe()
    assert info.value.tenant_id == TENANT_A
    assert fake_msal.interactive_calls == 0
