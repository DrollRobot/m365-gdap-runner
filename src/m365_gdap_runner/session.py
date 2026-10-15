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

import azure_auth
import httpx
from azure_auth import AuthContext, AuthError
from azure_auth.clouds import Cloud, get_cloud
from azure_auth.sync import GraphClient

from m365_gdap_runner.config import ConfigError, Settings
from m365_gdap_runner.context import CLIENT_KINDS, TenantContext, build_client
from m365_gdap_runner.tenants import Tenant

logger = logging.getLogger(__name__)

# Delegated scope the partner sign-in needs to list GDAP customers. Payload scopes are never
# added here: consent in the partner tenant says nothing about the customer tenants.
PARTNER_SCOPES = ("DelegatedAdminRelationship.Read.All",)

CACHE_KINDS = ("memory", "disk")


def _partner_tenant(name: str, cloud: Cloud) -> azure_auth.Tenant:
    """Build the partner tenant in a configured cloud, looking it up only when needed.

    A GUID with its cloud is all azure-auth needs, so no request is made. A domain name is
    still looked up, for its GUID, in that cloud only.

    Args:
        name: The ``partner_tenant_id`` setting: a GUID or a domain name.
        cloud: The ``cloud`` setting.

    Returns:
        The tenant, for :class:`~azure_auth.AuthContext`.

    Raises:
        azure_auth.AuthError: If a domain name cannot be looked up in that cloud.
    """
    try:
        return azure_auth.Tenant(name, cloud)
    except ValueError:
        return azure_auth.Tenant.lookup(name, cloud)


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
        tenant: str | azure_auth.Tenant = settings.partner_tenant_id
        if settings.cloud:
            try:
                cloud = get_cloud(settings.cloud)
            except ValueError as exc:
                raise ConfigError(str(exc)) from exc
            tenant = _partner_tenant(settings.partner_tenant_id, cloud)
        self._transport = transport
        self.auth = AuthContext(
            tenant,
            username=settings.username,
            client_id=settings.client_id or None,
            cache=cast(Literal["memory", "disk"], settings.cache),
        )

    @property
    def partner_tenant_id(self) -> str:
        """The partner tenant's GUID."""
        return self.auth.tenant.id

    @property
    def cloud_name(self) -> str:
        """Name of the cloud this session signs in to, as azure-auth spells it."""
        return self.auth.tenant.cloud.name

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
            return azure_auth.Tenant.lookup(tenant_id, transport=self._transport).cloud.name
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
            customers = graph.list_customers()
        tenants: dict[str, Tenant] = {}
        for customer in customers:
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
