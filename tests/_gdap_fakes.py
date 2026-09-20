"""Shared fakes for the runner tests: MSAL, sessions, credentials, and Graph responses.

Fake access tokens are ``at:<tenant_id>`` so Graph fakes can tell which tenant a
request is for from its Authorization header.
"""

from __future__ import annotations

import threading
from collections.abc import Mapping
from typing import Any

import httpx
from azure.core.credentials import AccessToken, TokenCredential

from m365_gdap_runner.config import Settings

PARTNER_TENANT = "11111111-1111-1111-1111-111111111111"
TENANT_A = "aaaaaaaa-0000-0000-0000-00000000000a"
TENANT_B = "bbbbbbbb-0000-0000-0000-00000000000b"
TENANT_C = "cccccccc-0000-0000-0000-00000000000c"
USER_OID = "22222222-2222-2222-2222-222222222222"
HOME_ACCOUNT_ID = f"{USER_OID}.{PARTNER_TENANT}"

# A far-future expiry for fake tokens (2100-01-01).
FAR_FUTURE = 4102444800

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
    values: dict[str, Any] = {"partner_tenant_id": PARTNER_TENANT}
    values.update(overrides)
    return Settings(**values)


def fake_access(tenant_id: str) -> str:
    """Return the fake access token value for a tenant."""
    return f"at:{tenant_id}"


def bearer_tenant(request: httpx.Request) -> str:
    """Return the tenant a fake bearer token was issued for."""
    return request.headers["Authorization"].removeprefix("Bearer at:")


class FakeMsalApp:
    """Fake ``PublicClientApplication`` bound to one authority."""

    def __init__(self, broker: FakeMsal, authority: str) -> None:
        self.broker = broker
        self.authority = authority
        self.tenant_id = authority.rsplit("/", 1)[1]

    def get_accounts(self) -> list[dict[str, Any]]:
        return list(self.broker.accounts)

    def acquire_token_silent_with_error(
        self,
        scopes: list[str],
        account: Mapping[str, Any] | None,
        *,
        claims_challenge: str | None = None,
    ) -> dict[str, Any] | None:
        with self.broker.lock:
            self.broker.silent_calls.append((self.tenant_id, tuple(scopes), claims_challenge))
        if account is None or not self.broker.has_refresh_token:
            return None
        error = self.broker.tenant_errors.get(self.tenant_id.lower())
        if error is not None:
            return dict(error)
        return {"access_token": fake_access(self.tenant_id), "expires_in": 3600}

    def acquire_token_interactive(self, scopes: list[str]) -> dict[str, Any]:
        self.broker.interactive_calls += 1
        if self.broker.interactive_result is not None:
            return dict(self.broker.interactive_result)
        self.broker.seed_cached_account()
        return {
            "access_token": fake_access(self.tenant_id),
            "expires_in": 3600,
            "id_token_claims": {"oid": USER_OID, "tid": PARTNER_TENANT},
        }


class FakeMsal:
    """Fake MSAL app factory; its apps share state the way a shared token cache would."""

    def __init__(
        self,
        *,
        tenant_errors: Mapping[str, Mapping[str, Any]] | None = None,
        interactive_result: Mapping[str, Any] | None = None,
    ) -> None:
        self.tenant_errors = {k.lower(): dict(v) for k, v in (tenant_errors or {}).items()}
        self.interactive_result = dict(interactive_result) if interactive_result else None
        self.accounts: list[dict[str, Any]] = []
        self.has_refresh_token = False
        self.interactive_calls = 0
        self.silent_calls: list[tuple[str, tuple[str, ...], str | None]] = []
        self.created: list[tuple[str, str, object]] = []
        self.lock = threading.Lock()

    def __call__(self, client_id: str, *, authority: str, token_cache: object) -> FakeMsalApp:
        with self.lock:
            self.created.append((client_id, authority, token_cache))
        return FakeMsalApp(self, authority)

    def seed_cached_account(self) -> None:
        """Put the partner user's account and refresh token in the fake cache."""
        if not any(a["home_account_id"] == HOME_ACCOUNT_ID for a in self.accounts):
            self.accounts.append({"home_account_id": HOME_ACCOUNT_ID})
        self.has_refresh_token = True


class FakeCredential(TokenCredential):
    """Credential that returns ``at:<tenant>`` tokens or raises a set error."""

    def __init__(self, tenant_id: str, error: Exception | None = None) -> None:
        self.tenant_id = tenant_id
        self.error = error
        self.calls: list[tuple[str, ...]] = []

    def get_token(
        self,
        *scopes: str,
        claims: str | None = None,
        tenant_id: str | None = None,
        enable_cae: bool = False,
        **kwargs: Any,
    ) -> AccessToken:
        self.calls.append(scopes)
        if self.error is not None:
            raise self.error
        return AccessToken(fake_access(self.tenant_id), FAR_FUTURE)


class FakeSession:
    """Session whose credentials fail for the tenants given in *auth_errors*."""

    def __init__(self, auth_errors: Mapping[str, Exception] | None = None) -> None:
        self.partner_tenant_id = PARTNER_TENANT
        self.auth_errors = dict(auth_errors or {})
        self.credentials: dict[str, FakeCredential] = {}
        self.lock = threading.Lock()

    def credential_for(self, tenant_id: str) -> FakeCredential:
        credential = FakeCredential(tenant_id, self.auth_errors.get(tenant_id))
        with self.lock:
            self.credentials[tenant_id] = credential
        return credential


def fake_graph(request: httpx.Request) -> httpx.Response:
    """Fake Graph: GDAP relationships for the partner, one licensed user per customer."""
    if request.url.path == "/v1.0/tenantRelationships/delegatedAdminRelationships":
        return httpx.Response(200, json=relationships_body(*CUSTOMERS))
    if request.url.path == "/v1.0/users":
        user = {
            "userPrincipalName": f"admin@{bearer_tenant(request)}.example",
            "displayName": "Admin",
            "assignedLicenses": [{"skuId": "sku-2"}, {"skuId": "sku-1"}],
        }
        return httpx.Response(200, json={"value": [user]})
    return httpx.Response(404, json={"error": {"code": "NotFound", "message": request.url.path}})


# Customers the fake partner tenant has active GDAP relationships with.
CUSTOMERS = ((TENANT_A, "Alpha"), (TENANT_B, "beta corp"))


def relationships_body(*customers: tuple[str, str]) -> dict[str, Any]:
    """Build a delegatedAdminRelationships page for (tenant_id, display_name) pairs."""
    return {
        "value": [
            {
                "id": f"relationship-{index}",
                "status": "active",
                "customer": {"tenantId": tenant_id, "displayName": name},
            }
            for index, (tenant_id, name) in enumerate(customers)
        ]
    }
