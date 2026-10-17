"""The partner sign-in and what hangs off it: tenant listing and per-tenant contexts.

azure-auth does the authentication. This module only decides *what* to sign in to: one
sign-in per client id the script uses, all at the partner (home) tenant, before the loop
starts. Customer tenants are reached through azure-auth sibling contexts, which never prompt;
they redeem the home sign-in's refresh token in the customer tenant.

A tenant the partner sign-in cannot reach (no GDAP relationship, missing consent, a policy
that wants a fresh sign-in, or another Microsoft cloud) gets a direct sign-in instead: a
context of its own, in its own cloud, as the tenant's own account. A tenant with no account
of its own gets none. All of that happens in :meth:`PartnerSession.check_access`, before the
loop, so the loop never opens a browser.

Why sign in per client: Microsoft's first-party apps differ per resource (Graph PowerShell,
Exchange Online PowerShell, Azure PowerShell) and a refresh token belongs to one app, so a
script that uses Graph and Exchange needs two home sign-ins. With an encrypted disk cache
later runs are silent.
"""

from __future__ import annotations

import logging
from collections.abc import Iterable, Sequence
from concurrent.futures import ThreadPoolExecutor
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
        self._settings = settings
        # How check_access() found each tenant can be reached, by lower-cased tenant ID.
        self._routes: dict[str, AuthContext | Exception] = {}
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

    def check_access(
        self,
        tenants: Sequence[Tenant],
        *,
        scopes: Sequence[str] = (),
        clients: Sequence[str] = ("graph",),
        direct: bool = True,
        jobs: int = 1,
    ) -> None:
        """Find a way into every tenant before the run, so the run itself never prompts.

        A tenant in the partner's cloud is checked first through the partner sign-in, silently
        and the way the run will use it: every client the script uses must get a token
        (:meth:`TenantContext.probe`). A tenant that fails that check, for whatever reason,
        and a tenant in another cloud, then get a direct sign-in: a context of their own, in
        their own cloud, as their ``username``, signed in once per client. One with no
        ``username`` gets none, and no prompt. The browser opens only where no cached sign-in
        will do. Direct sign-ins run one at a time, on this thread.

        :meth:`context_for` then hands out what was found, and raises the reason for a tenant
        that could not be reached.

        Args:
            tenants: The selected tenants, with their cloud filled in.
            scopes: The script's Graph scopes.
            clients: The script's ``CLIENTS``.
            direct: Sign in directly where the partner sign-in fails. False leaves those
                tenants unreachable, for unattended runs.
            jobs: How many silent checks run at once.
        """
        partner_cloud = self.cloud_name.lower()
        nearby = [t for t in tenants if (t.cloud or self.cloud_name).lower() == partner_cloud]
        logger.info("Checking access to %d tenant(s)", len(tenants))
        with ThreadPoolExecutor(
            max_workers=max(1, min(jobs, len(nearby))), thread_name_prefix="check"
        ) as pool:
            checked = dict(
                zip(
                    (t.tenant_id.lower() for t in nearby),
                    pool.map(lambda t: self._through_partner(t, scopes, clients), nearby),
                    strict=True,
                )
            )
        for tenant in tenants:
            key = tenant.tenant_id.lower()
            route = checked.get(key)
            if isinstance(route, AuthContext):
                self._routes[key] = route
                continue
            if route is None:
                why = f"it is in the {tenant.cloud} cloud"
                route = AuthError(
                    f"{tenant.display_name} is in the {tenant.cloud} cloud, which the partner "
                    "sign-in cannot reach; it needs a direct sign-in."
                )
            else:
                why = f"the partner sign-in cannot reach it ({route})"
            if not direct:
                logger.warning("Skipping %s: %s", tenant.display_name, why)
                self._routes[key] = route
                continue
            if not tenant.username:
                logger.warning(
                    "Skipping %s: %s, and its m365_username is blank", tenant.display_name, why
                )
                self._routes[key] = AuthError(
                    f"{route} No direct sign-in to {tenant.display_name}: its m365_username is "
                    "blank."
                )
                continue
            logger.info(
                "Signing in to %s directly as %s: %s", tenant.display_name, tenant.username, why
            )
            self._routes[key] = self._sign_in_directly(tenant, scopes, clients)

    def _through_partner(
        self, tenant: Tenant, scopes: Sequence[str], clients: Sequence[str]
    ) -> AuthContext | Exception:
        """Check a tenant through the partner sign-in, silently.

        Args:
            tenant: The tenant, in the partner's cloud.
            scopes: The script's Graph scopes.
            clients: The script's ``CLIENTS``.

        Returns:
            The sibling context when every client got a token, else what went wrong.
        """
        try:
            auth = self.auth.for_tenant(tenant.tenant_id)
            with TenantContext(
                tenant, auth, scopes=scopes, clients=clients, transport=self._transport
            ) as context:
                context.probe()
        except Exception as exc:
            logger.debug("Partner sign-in cannot reach %s", tenant.display_name, exc_info=True)
            return exc
        return auth

    def _sign_in_directly(
        self, tenant: Tenant, scopes: Sequence[str], clients: Sequence[str]
    ) -> AuthContext | Exception:
        """Sign in to a tenant on its own, as its own account, once per client the script uses.

        Args:
            tenant: The tenant, with its cloud and username filled in.
            scopes: The script's Graph scopes.
            clients: The script's ``CLIENTS``.

        Returns:
            The tenant's own context, signed in to every client, else what went wrong.
        """
        try:
            cloud = get_cloud(tenant.cloud or self.cloud_name)
            auth = AuthContext(
                azure_auth.Tenant(tenant.tenant_id, cloud),
                username=tenant.username,
                client_id=self._settings.client_id or None,
                cache=cast(Literal["memory", "disk"], self._settings.cache),
            )
            for kind in clients:
                with build_client(kind, auth, scopes=scopes, transport=self._transport) as client:
                    client.login()
        except Exception as exc:
            logger.warning(
                "Could not sign in to %s as %s: %s", tenant.display_name, tenant.username, exc
            )
            return exc
        return auth

    def context_for(
        self, tenant: Tenant, *, scopes: Sequence[str] = (), clients: Sequence[str] = ("graph",)
    ) -> TenantContext:
        """Build the context a script receives for one customer tenant.

        Args:
            tenant: The customer tenant.
            scopes: The script's Graph scopes.
            clients: The script's ``CLIENTS``.

        Returns:
            The context, with one client per declared kind: on the route
            :meth:`check_access` found, or through the partner sign-in for a tenant it did
            not check. No network call is made.

        Raises:
            Exception: Why :meth:`check_access` could not reach the tenant.
        """
        route = self._routes.get(tenant.tenant_id.lower())
        if isinstance(route, Exception):
            raise route
        return TenantContext(
            tenant,
            route or self.auth.for_tenant(tenant.tenant_id),
            scopes=scopes,
            clients=clients,
            transport=self._transport,
        )
