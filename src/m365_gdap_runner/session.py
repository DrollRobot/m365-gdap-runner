"""The partner sign-in and what hangs off it: tenant listing and per-tenant contexts.

azure-auth does the authentication. This module only decides *what* to sign in to: one
sign-in per client id the payload uses, all at the partner (home) tenant, before the loop
starts. Customer tenants are reached through azure-auth sibling contexts, which never prompt;
they redeem the home sign-in's refresh token in the customer tenant.

Why sign in per client: Microsoft's first-party apps differ per resource (Graph PowerShell,
Exchange Online PowerShell, Azure PowerShell) and a refresh token belongs to one app, so a
payload that uses Graph and Exchange needs two home sign-ins. With an encrypted disk cache
later runs are silent.
"""

from __future__ import annotations

import logging
from collections.abc import Iterable, Sequence
from typing import Literal, cast

import httpx
from azure_auth import AuthContext, AuthError, discover_tenant
from azure_auth.clouds import get_cloud
from azure_auth.sync import GraphClient

from m365_gdap_runner.config import ConfigError, Settings
from m365_gdap_runner.context import CLIENT_KINDS, TenantContext, build_client
from m365_gdap_runner.tenants import Tenant

logger = logging.getLogger(__name__)

# Delegated scope the partner sign-in needs to list GDAP customers. Payload scopes are never
# added here: consent in the partner tenant says nothing about the customer tenants.
PARTNER_SCOPES = ("DelegatedAdminRelationship.Read.All",)

# Partner-tenant Graph collection listing GDAP customers. azure-auth will grow a method that
# returns these with their display names; until then the runner reads it directly.
CUSTOMERS_PATH = "/tenantRelationships/delegatedAdminCustomers"

CACHE_KINDS = ("memory", "disk")


class PartnerSession:
    """One partner sign-in, and the customer tenants reachable from it.

    Attributes:
        auth: The root azure-auth context, for the partner tenant.
    """

    def __init__(self, settings: Settings, *, transport: httpx.BaseTransport | None = None) -> None:
        """Build the root context from settings. No sign-in happens yet.

        Args:
            settings: Resolved settings.
            transport: Optional httpx transport for every client this session builds (tests).

        Raises:
            ConfigError: If ``cache`` or ``cloud`` holds a value azure-auth does not know.
            azure_auth.AuthError: If the partner tenant cannot be looked up, or the disk
                cache cannot be encrypted on this platform.
        """
        if settings.cache not in CACHE_KINDS:
            raise ConfigError(
                f"cache must be one of {', '.join(CACHE_KINDS)}, not {settings.cache!r}."
            )
        cloud = None
        if settings.cloud:
            try:
                cloud = get_cloud(settings.cloud)
            except ValueError as exc:
                raise ConfigError(str(exc)) from exc
        self._transport = transport
        self.auth = AuthContext(
            settings.partner_tenant_id,
            username=settings.username,
            client_id=settings.client_id or None,
            cache=cast(Literal["memory", "disk"], settings.cache),
            cloud=cloud,
        )

    @property
    def partner_tenant_id(self) -> str:
        """The partner tenant's GUID."""
        return self.auth.tenant_id

    @property
    def cloud_name(self) -> str:
        """Name of the cloud this session signs in to, as azure-auth spells it."""
        return self.auth.cloud.name

    def cloud_of(self, tenant_id: str) -> str:
        """Find the cloud a tenant lives in, by OpenID Connect discovery.

        Discovery is unauthenticated and asks each public sign-in host once, so it costs
        three requests and works for any tenant. It fills the tenants file's ``m365_cloud``
        column.

        Args:
            tenant_id: The tenant's GUID.

        Returns:
            The cloud's name (``Commercial``, ``USGov``, ``USGovDoD`` or ``China``), or ``""``
            when no sign-in host knows the tenant or none can be reached.
        """
        try:
            return discover_tenant(tenant_id, transport=self._transport).cloud.name
        except (AuthError, httpx.HTTPError, ValueError) as exc:
            logger.debug("Could not look up the cloud of %s: %s", tenant_id, exc)
            return ""

    def login(self, clients: Iterable[str] = ()) -> None:
        """Sign in at the partner tenant, once per client id the payload needs.

        Graph is always signed in, for listing customers. The browser opens only for a client
        id with no usable cached sign-in.

        Args:
            clients: The payload's ``CLIENTS``.

        Raises:
            azure_auth.AuthError: If a sign-in fails or is cancelled.
        """
        kinds = ["graph", *(kind for kind in clients if kind != "graph")]
        for kind in kinds:
            if kind not in CLIENT_KINDS:
                raise ValueError(f"Unknown client kind {kind!r}.")
            logger.debug("Signing in to the partner tenant for %s", kind)
            scopes = PARTNER_SCOPES if kind == "graph" else ()
            with build_client(kind, self.auth, scopes=scopes, transport=self._transport) as client:
                client.login()

    def list_tenants(self) -> list[Tenant]:
        """List the customer tenants with a GDAP relationship to the partner.

        Returns:
            Tenants deduplicated by ID, sorted by display name. A customer without a display
            name is named by its ID.

        Raises:
            azure_auth.AuthError: If no partner token can be obtained.
            azure_auth.GraphError: If Graph rejects the request.
        """
        with GraphClient(
            self.auth, scopes=list(PARTNER_SCOPES), transport=self._transport
        ) as graph:
            customers = graph.get_all(CUSTOMERS_PATH)
        tenants: dict[str, Tenant] = {}
        for customer in customers:
            if not isinstance(customer, dict) or not customer.get("tenantId"):
                continue
            tenant_id = str(customer["tenantId"])
            display_name = str(customer.get("displayName") or tenant_id)
            tenants.setdefault(tenant_id.lower(), Tenant(tenant_id, display_name))
        return sorted(
            tenants.values(), key=lambda t: (t.display_name.casefold(), t.tenant_id.lower())
        )

    def context_for(
        self, tenant: Tenant, *, scopes: Sequence[str] = (), clients: Sequence[str] = ("graph",)
    ) -> TenantContext:
        """Build the context a payload receives for one customer tenant.

        Args:
            tenant: The customer tenant.
            scopes: The payload's Graph scopes.
            clients: The payload's ``CLIENTS``.

        Returns:
            The context, with one client per declared kind. No network call is made.
        """
        return TenantContext(
            tenant,
            self.auth.for_tenant(tenant.tenant_id),
            scopes=scopes,
            clients=clients,
            transport=self._transport,
        )
