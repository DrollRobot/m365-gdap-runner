"""Shared fakes for the runner tests.

Two boundaries are faked: the ``msal`` module azure-auth signs in with, and the HTTP
transport its clients send requests through. Everything between them (``AuthContext``,
sibling contexts, the Graph and Exchange clients) is real.

Fake access tokens are unsigned JWTs whose ``tid`` claim is the tenant they were issued
for, so a fake service can tell which tenant a request is for from its bearer token, and
azure-auth's Exchange client can read the tenant GUID out of it the way it does live.
"""

from __future__ import annotations

import base64
import json
import threading
import uuid
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
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


def make_settings(**overrides: Any) -> Settings:
    """Build Settings for the fake partner tenant, with field overrides."""
    values: dict[str, Any] = {"partner_tenant_id": PARTNER_TENANT, "username": USER}
    values.update(overrides)
    return Settings(**values)


# --- tokens ------------------------------------------------------------------------------


def _b64(value: dict[str, Any]) -> str:
    return base64.urlsafe_b64encode(json.dumps(value).encode()).rstrip(b"=").decode()


def fake_access(tenant_id: str) -> str:
    """Return the fake access token for a tenant: an unsigned JWT with a ``tid`` claim."""
    return f"{_b64({'alg': 'none'})}.{_b64({'tid': tenant_id})}.signature"


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
        # A refresh token belongs to one client id: no interactive sign-in for this client
        # id yet means MSAL finds nothing to redeem, whatever the tenant.
        if self.client_id not in self.module.signed_in:
            return None
        error = self.module.tenant_errors.get(self.tenant_id.lower())
        if error is not None:
            return dict(error)
        return {"access_token": fake_access(self.tenant_id), "expires_in": 3600}

    def acquire_token_interactive(self, scopes: list[str], **kwargs: Any) -> dict[str, Any]:
        self.module.record("interactive", self, scopes)
        if self.module.interactive_result is not None:
            return dict(self.module.interactive_result)
        self.module.sign_in(self.client_id)
        return {
            "access_token": fake_access(self.tenant_id),
            "expires_in": 3600,
            "id_token_claims": {"preferred_username": USER, "tid": self.tenant_id},
        }


@dataclass
class FakeMsal:
    """Replacement for the ``msal`` module: one shared cache, scriptable per tenant.

    Attributes:
        tenant_errors: Tenant id to the MSAL error result its silent requests return.
        interactive_result: When set, what every interactive sign-in returns (a failure).
        accounts: Cached accounts, shared by every application.
        signed_in: Client ids that have signed in interactively (and so hold a refresh token).
        calls: Every token request, in order.
    """

    tenant_errors: dict[str, dict[str, Any]] = field(default_factory=dict)
    interactive_result: dict[str, Any] | None = None
    accounts: list[dict[str, Any]] = field(default_factory=list)
    signed_in: set[str] = field(default_factory=set)
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

    def record(self, method: str, app: FakeApp, scopes: Sequence[str]) -> None:
        with self.lock:
            self.calls.append(Call(method, app.tenant_id, app.client_id, tuple(scopes)))

    def sign_in(self, client_id: str) -> None:
        """Put the partner user's account and a refresh token for *client_id* in the cache."""
        with self.lock:
            self.signed_in.add(client_id)
            if not any(a["username"] == USER for a in self.accounts):
                self.accounts.append({"username": USER, "home_account_id": f"oid.{PARTNER_TENANT}"})

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
    """A tenant context with no clients: ``probe()`` raises what it is told to."""

    def __init__(self, tenant: Tenant, error: BaseException | None = None) -> None:
        self.tenant = tenant
        self.error = error
        self.closed = False
        self._clients = {}

    def probe(self) -> None:
        if self.error is not None:
            raise self.error

    def close(self) -> None:
        self.closed = True


class FakeSession:
    """Session whose contexts fail their probe for the tenants given in *auth_errors*."""

    def __init__(self, auth_errors: Mapping[str, BaseException] | None = None) -> None:
        self.auth_errors = dict(auth_errors or {})
        self.contexts: dict[str, FakeContext] = {}
        self.requests: list[tuple[str, tuple[str, ...], tuple[str, ...]]] = []
        self.lock = threading.Lock()

    def context_for(
        self, tenant: Tenant, *, scopes: Sequence[str] = (), clients: Sequence[str] = ("graph",)
    ) -> TenantContext:
        context = FakeContext(tenant, self.auth_errors.get(tenant.tenant_id))
        with self.lock:
            self.contexts[tenant.tenant_id] = context
            self.requests.append((tenant.tenant_id, tuple(scopes), tuple(clients)))
        return context
