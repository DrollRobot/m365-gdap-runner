"""GDAP customer tenants and how ``--tenant``/``--exclude`` pick them."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass


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
