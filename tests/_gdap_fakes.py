"""Shared fakes for the runner tests.

Two boundaries are faked: the ``msal`` module azure-auth signs in with, and the HTTP
transport its clients send requests through. Everything between them (``AuthContext``, the
contexts ``get_tenant_context()`` returns, the Graph and Exchange clients) is real.

Fake access tokens are unsigned JWTs whose ``tid`` claim is the tenant they were issued
for, so a fake service can tell which tenant a request is for from its bearer token, and
azure-auth's Exchange client can read the tenant GUID out of it the way it does live.
"""

from __future__ import annotations

import base64
import importlib.util
import json
import sys
import threading
import uuid
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from types import ModuleType
from typing import Any

import azure_auth
import httpx
from azure_auth.clouds import COMMERCIAL, Cloud

from m365_gdap_runner.config import Settings
from m365_gdap_runner.context import TenantContext
from m365_gdap_runner.tenants import Tenant

PARTNER_TENANT = "11111111-1111-1111-1111-111111111111"
TENANT_A = "aaaaaaaa-0000-0000-0000-00000000000a"
TENANT_B = "bbbbbbbb-0000-0000-0000-00000000000b"
TENANT_C = "cccccccc-0000-0000-0000-00000000000c"
USER = "admin@partner.example"
# An account of a tenant the partner has no GDAP relationship with.
GAMMA_ADMIN = "admin@gamma.example"

# Customers the fake partner tenant has GDAP relationships with.
CUSTOMERS = ((TENANT_A, "Alpha"), (TENANT_B, "beta corp"))

CONSENT_REQUIRED: dict[str, Any] = {
    "error": "invalid_grant",
    "error_description": (
        "AADSTS65001: The user or administrator has not consented to use the application."
        "\r\nTrace ID: 0\r\nCorrelation ID: 0"
    ),
    "error_codes": [65001],
    "suberror": "consent_required",
}

# What Entra says when the partner account asks a tenant it has no access to for a token.
NOT_IN_TENANT: dict[str, Any] = {
    "error": "invalid_grant",
    "error_description": (
        f"AADSTS50020: User account '{USER}' from identity provider does not exist in tenant."
    ),
    "error_codes": [50020],
}


def make_settings(**overrides: Any) -> Settings:
    """Build Settings for the fake partner tenant, with field overrides."""
    values: dict[str, Any] = {"partner_tenant_id": PARTNER_TENANT, "username": USER}
    values.update(overrides)
    return Settings(**values)


def load_script(path: Path) -> ModuleType:
    """Import a script file as its own module, the way running it would, minus ``__main__``."""
    name = f"_test_script_{uuid.uuid4().hex}"
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    # Registered first, as for any import: dataclasses look their module up while being built.
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


# --- tokens ------------------------------------------------------------------------------


def _b64(value: dict[str, Any]) -> str:
    return base64.urlsafe_b64encode(json.dumps(value).encode()).rstrip(b"=").decode()


def fake_access(tenant_id: str, scp: str | None = None) -> str:
    """Return the fake access token for a tenant: an unsigned JWT with a ``tid`` claim.

    With *scp*, the token also carries it, the scopes granted there, space-separated. Without
    it, azure-auth cannot tell what the tenant has granted and asks for the scopes anyway.
    """
    claims = {"tid": tenant_id} if scp is None else {"tid": tenant_id, "scp": scp}
    return f"{_b64({'alg': 'none'})}.{_b64(claims)}.signature"


def bearer_tenant(request: httpx.Request) -> str:
    """Return the tenant a request's fake bearer token was issued for."""
    token = request.headers["Authorization"].removeprefix("Bearer ")
    payload = token.split(".")[1]
    claims = json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
    return str(claims["tid"])


# --- msal --------------------------------------------------------------------------------


@dataclass
class Call:
    """One recorded token request."""

    method: str
    tenant_id: str
    client_id: str
    scopes: tuple[str, ...]


class FakeApp:
    """Fake ``PublicClientApplication`` bound to one client id and one tenant authority."""

    def __init__(
        self, module: FakeMsal, client_id: str, authority: str, token_cache: object
    ) -> None:
        self.module = module
        self.client_id = client_id
        self.authority = authority
        self.token_cache = token_cache
        self.tenant_id = authority.rsplit("/", 1)[1]

    def get_accounts(self, username: str | None = None) -> list[dict[str, Any]]:
        return [
            account
            for account in self.module.accounts
            if username is None or account["username"].lower() == username.lower()
        ]

    def acquire_token_silent_with_error(
        self, scopes: list[str], account: Mapping[str, Any], **kwargs: Any
    ) -> dict[str, Any] | None:
        self.module.record("silent", self, scopes)
        # A refresh token belongs to one account and one client id: no interactive sign-in
        # for both yet means MSAL finds nothing to redeem, whatever the tenant.
        if (self.client_id, account["username"].lower()) not in self.module.signed_in:
            return None
        error = self.module.account_errors.get(
            (self.tenant_id.lower(), account["username"].lower())
        ) or self.module.tenant_errors.get(self.tenant_id.lower())
        if error is not None:
            return dict(error)
        return {"access_token": self.module.access(self.tenant_id), "expires_in": 3600}

    def acquire_token_interactive(self, scopes: list[str], **kwargs: Any) -> dict[str, Any]:
        self.module.record("interactive", self, scopes)
        if self.module.interactive_result is not None:
            return dict(self.module.interactive_result)
        refused = self.module.interactive_errors.get(self.tenant_id.lower())
        if refused is not None:
            return dict(refused)
        # The person signs in as the account the page was pre-filled with.
        username = kwargs.get("login_hint") or USER
        self.module.sign_in(self.client_id, username)
        return {
            "access_token": self.module.access(self.tenant_id),
            "expires_in": 3600,
            "id_token_claims": {"preferred_username": username, "tid": self.tenant_id},
        }


@dataclass
class FakeMsal:
    """Replacement for the ``msal`` module: one shared cache, scriptable per tenant.

    Attributes:
        tenant_errors: Tenant id to the MSAL error result its silent requests return.
        account_errors: ``(tenant id, username)``, lower-cased, to the MSAL error result that
            account's silent requests in that tenant return, ahead of ``tenant_errors``.
        interactive_result: When set, what every interactive sign-in returns (a failure).
        interactive_errors: Tenant id to the MSAL error result its interactive sign-ins return.
        granted: Tenant id to the scopes it has granted, space-separated, which its tokens
            carry as their ``scp`` claim. A tenant not in it gets tokens with none.
        accounts: Cached accounts, shared by every application.
        signed_in: ``(client_id, username)`` pairs that have signed in interactively, and so
            hold a refresh token.
        calls: Every token request, in order.
    """

    tenant_errors: dict[str, dict[str, Any]] = field(default_factory=dict)
    account_errors: dict[tuple[str, str], dict[str, Any]] = field(default_factory=dict)
    interactive_result: dict[str, Any] | None = None
    interactive_errors: dict[str, dict[str, Any]] = field(default_factory=dict)
    granted: dict[str, str] = field(default_factory=dict)
    accounts: list[dict[str, Any]] = field(default_factory=list)
    signed_in: set[tuple[str, str]] = field(default_factory=set)
    calls: list[Call] = field(default_factory=list)
    apps: list[FakeApp] = field(default_factory=list)
    lock: threading.Lock = field(default_factory=threading.Lock)

    def TokenCache(self) -> object:
        return object()

    def PublicClientApplication(
        self, client_id: str, *, authority: str, token_cache: object, **kwargs: Any
    ) -> FakeApp:
        app = FakeApp(self, client_id, authority, token_cache)
        with self.lock:
            self.apps.append(app)
        return app

    def access(self, tenant_id: str) -> str:
        """Return the fake access token for a tenant, with what it has granted."""
        return fake_access(tenant_id, self.granted.get(tenant_id.lower()))

    def record(self, method: str, app: FakeApp, scopes: Sequence[str]) -> None:
        with self.lock:
            self.calls.append(Call(method, app.tenant_id, app.client_id, tuple(scopes)))

    def sign_in(self, client_id: str, username: str = USER) -> None:
        """Put an account and a refresh token for *client_id* in the cache."""
        with self.lock:
            self.signed_in.add((client_id, username.lower()))
            if not any(a["username"] == username for a in self.accounts):
                self.accounts.append({"username": username, "home_account_id": f"oid.{username}"})

    @property
    def interactive_calls(self) -> int:
        return sum(1 for call in self.calls if call.method == "interactive")

    def silent_calls(self, tenant_id: str) -> list[tuple[str, tuple[str, ...]]]:
        """Return ``(client_id, scopes)`` of every silent request for a tenant."""
        return [
            (call.client_id, call.scopes)
            for call in self.calls
            if call.method == "silent" and call.tenant_id.lower() == tenant_id.lower()
        ]

    def interactive_tenants(self) -> list[str]:
        """Return the tenant of every interactive sign-in, in order."""
        return [call.tenant_id for call in self.calls if call.method == "interactive"]

    def authority(self, tenant_id: str) -> str:
        """Return the sign-in authority the applications for a tenant were built with."""
        (authority,) = {app.authority for app in self.apps if app.tenant_id == tenant_id}
        return authority


@dataclass
class FakeDiscovery:
    """Replacement for ``azure_auth.Tenant.lookup``: every tenant is in the cloud asked for.

    A lookup that names no cloud finds the tenant in ``cloud``.
    """

    cloud: Cloud = COMMERCIAL
    tenants: list[str] = field(default_factory=list)

    def __call__(self, name: str, cloud: Cloud | None = None, **kwargs: Any) -> azure_auth.Tenant:
        self.tenants.append(name)
        try:
            guid, domain = str(uuid.UUID(name)), None
        except ValueError:
            guid, domain = str(uuid.uuid5(uuid.NAMESPACE_DNS, name.lower())), name.lower()
        return azure_auth.Tenant(guid, cloud or self.cloud, domain=domain)


# --- services ----------------------------------------------------------------------------


def customers_body(*customers: tuple[str, str]) -> dict[str, Any]:
    """Build a delegatedAdminCustomers page for (tenant_id, display_name) pairs."""
    return {
        "value": [
            {"id": tenant_id, "tenantId": tenant_id, "displayName": name}
            for tenant_id, name in customers
        ]
    }


def fake_services(request: httpx.Request) -> httpx.Response:
    """Fake Graph and Exchange: GDAP customers, one licensed user, and the org config."""
    path = request.url.path
    tenant = bearer_tenant(request)
    if path == "/v1.0/tenantRelationships/delegatedAdminCustomers":
        return httpx.Response(200, json=customers_body(*CUSTOMERS))
    if path == "/v1.0/users":
        user = {
            "userPrincipalName": f"admin@{tenant}.example",
            "displayName": "Admin",
            "assignedLicenses": [{"skuId": "sku-2"}, {"skuId": "sku-1"}],
        }
        return httpx.Response(200, json={"value": [user]})
    if path == f"/adminapi/beta/{tenant}/InvokeCommand":
        cmdlet = json.loads(request.content)["CmdletInput"]["CmdletName"]
        if cmdlet == "Get-OrganizationConfig":
            config = {"Name": f"{tenant}.onmicrosoft.com", "AuditDisabled": tenant == TENANT_B}
            return httpx.Response(200, json={"value": [config]})
        return httpx.Response(403, json={"error": {"code": "Forbidden", "message": cmdlet}})
    return httpx.Response(404, json={"error": {"code": "NotFound", "message": path}})


# --- runner fakes ------------------------------------------------------------------------


class FakeContext(TenantContext):
    """A tenant context with no clients."""

    def __init__(self, tenant: Tenant) -> None:
        self.tenant = tenant
        self.closed = False
        self._clients = {}

    def close(self) -> None:
        self.closed = True


class FakeSession:
    """Session that could not reach the tenants given in *auth_errors*, as check_access() left it.

    ``context_for()`` raises a tenant's error, as the real session does for a tenant it could
    not reach.
    """

    def __init__(self, auth_errors: Mapping[str, Exception] | None = None) -> None:
        self.auth_errors = dict(auth_errors or {})
        self.contexts: dict[str, FakeContext] = {}
        self.requests: list[tuple[str, tuple[str, ...], tuple[str, ...]]] = []
        self.lock = threading.Lock()

    def context_for(
        self, tenant: Tenant, *, scopes: Sequence[str] = (), clients: Sequence[str] = ("graph",)
    ) -> TenantContext:
        with self.lock:
            self.requests.append((tenant.tenant_id, tuple(scopes), tuple(clients)))
        error = self.auth_errors.get(tenant.tenant_id)
        if error is not None:
            raise error
        context = FakeContext(tenant)
        with self.lock:
            self.contexts[tenant.tenant_id] = context
        return context
