"""GDAP customer tenants: enumeration, selection, and the per-tenant payload context."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

import httpx
from azure.core.credentials import TokenCredential

from m365_gdap_runner.auth import PARTNER_SCOPES, TokenSession
from m365_gdap_runner.graph import GraphClient

# Partner-tenant Graph collection listing GDAP relationships.
RELATIONSHIPS_PATH = "/tenantRelationships/delegatedAdminRelationships"

# Only active relationships grant access.
ACTIVE_FILTER = "status eq 'active'"


class TenantSelectionError(ValueError):
    """Raised when a ``--tenant``/``--exclude`` value matches no tenant or several."""


@dataclass(frozen=True)
class Tenant:
    """A GDAP customer tenant.

    Attributes:
        tenant_id: Entra tenant ID.
        display_name: Tenant display name as set in Entra ID.
    """

    tenant_id: str
    display_name: str


@dataclass(frozen=True)
class TenantContext:
    """Everything a payload's ``run()`` receives for one tenant.

    Attributes:
        tenant: The tenant being processed.
        graph: Graph client bound to this tenant and the payload's Graph scopes.
        credential: Token credential bound to this tenant, for other SDKs or resources.
    """

    tenant: Tenant
    graph: GraphClient
    credential: TokenCredential


def list_gdap_tenants(
    session: TokenSession, *, transport: httpx.BaseTransport | None = None
) -> list[Tenant]:
    """Enumerate customer tenants with an active GDAP relationship.

    Args:
        session: A signed-in session.
        transport: Optional httpx transport for the Graph client (tests).

    Returns:
        Tenants deduplicated by tenant ID (a customer can have several active
        relationships), sorted by display name.

    Raises:
        TenantAuthError: If no partner-tenant token can be acquired.
        GraphError: If Graph rejects the request.
    """
    credential = session.credential_for(session.partner_tenant_id)
    with GraphClient(credential, PARTNER_SCOPES, transport=transport) as graph:
        relationships = graph.get_all(RELATIONSHIPS_PATH, params={"$filter": ACTIVE_FILTER})
    tenants: dict[str, Tenant] = {}
    for relationship in relationships:
        customer = relationship.get("customer")
        if not isinstance(customer, dict) or not customer.get("tenantId"):
            continue
        tenant_id = str(customer["tenantId"])
        display_name = str(customer.get("displayName") or tenant_id)
        tenants.setdefault(tenant_id.lower(), Tenant(tenant_id, display_name))
    return sorted(tenants.values(), key=lambda t: (t.display_name.casefold(), t.tenant_id.lower()))


def _match(tenants: Sequence[Tenant], spec: str) -> Tenant:
    """Resolve one tenant by ID or display name (case-insensitive).

    Args:
        tenants: Candidate tenants.
        spec: Tenant ID or display name.

    Returns:
        The single matching tenant. An ID match wins over a name match.

    Raises:
        TenantSelectionError: If nothing matches, or the name matches several tenants.
    """
    needle = spec.strip().casefold()
    for tenant in tenants:
        if tenant.tenant_id.casefold() == needle:
            return tenant
    by_name = [t for t in tenants if t.display_name.casefold() == needle]
    if not by_name:
        raise TenantSelectionError(f"No GDAP tenant matches {spec!r}.")
    if len(by_name) > 1:
        ids = ", ".join(t.tenant_id for t in by_name)
        raise TenantSelectionError(
            f"{spec!r} matches {len(by_name)} tenants by display name; use a tenant ID ({ids})."
        )
    return by_name[0]


def select(
    tenants: Sequence[Tenant],
    include: Sequence[str] | None = None,
    exclude: Sequence[str] | None = None,
) -> list[Tenant]:
    """Filter tenants by ID or display name, preserving order.

    Every ``include`` and ``exclude`` value must match exactly one tenant: a
    typo is an error, never a silent no-op.

    Args:
        tenants: All enumerated tenants.
        include: Keep only these tenants. None or empty keeps all.
        exclude: Drop these tenants.

    Returns:
        The selected tenants, in input order.

    Raises:
        TenantSelectionError: If a value matches no tenant or several.
    """
    chosen = list(tenants)
    if include:
        wanted = {_match(tenants, spec).tenant_id for spec in include}
        chosen = [t for t in chosen if t.tenant_id in wanted]
    if exclude:
        unwanted = {_match(tenants, spec).tenant_id for spec in exclude}
        chosen = [t for t in chosen if t.tenant_id not in unwanted]
    return chosen
