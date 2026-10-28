"""Dump what Graph says about this partner's GDAP customers and relationships.

Signs in to the partner tenant with your saved runner config, then writes Graph's answers,
unchanged, to two files:

- ``.local/gdap_customers.json``: ``GET /tenantRelationships/delegatedAdminCustomers``, the
  list the runner picks tenants from.
- ``.local/gdap_relationships.json``: ``GET /tenantRelationships/delegatedAdminRelationships``,
  which carries each relationship's status and end date.

It also prints one line per relationship, then the customers with no active relationship.
Run it from the repository root::

    uv run python scripts/dump_gdap_relationships.py
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from azure_auth.sync import GraphClient

from m365_gdap_runner.config import load_settings
from m365_gdap_runner.session import PARTNER_SCOPES, PartnerSession

OUT = Path(".local")


def _customer_of(relationship: dict[str, Any]) -> dict[str, Any]:
    """Return the customer a relationship is with, or an empty dict when Graph names none.

    Args:
        relationship: One relationship, as Graph returned it.

    Returns:
        Its ``customer`` object.
    """
    customer = relationship.get("customer")
    return customer if isinstance(customer, dict) else {}


def main() -> None:
    """Sign in, fetch both lists, write them out, and print the summary."""
    session = PartnerSession(load_settings())
    session.login()
    with GraphClient(session.auth, scopes=list(PARTNER_SCOPES)) as graph:
        customers = graph.get_all("/tenantRelationships/delegatedAdminCustomers")
        relationships = graph.get_all("/tenantRelationships/delegatedAdminRelationships")

    for name, items in (
        ("gdap_customers.json", customers),
        ("gdap_relationships.json", relationships),
    ):
        path = OUT / name
        path.write_text(json.dumps(items, indent=2), encoding="utf-8")
        print(f"Wrote {len(items)} item(s) to {path}")

    print("\nRelationships (customer, status, end date, relationship name):")
    for relationship in sorted(
        relationships, key=lambda r: str(_customer_of(r).get("displayName", "")).casefold()
    ):
        print(
            f"  {_customer_of(relationship).get('displayName', '?')!s:45} "
            f"{relationship.get('status', '?')!s:12} "
            f"{relationship.get('endDateTime', '?')!s:22} "
            f"{relationship.get('displayName', '')}"
        )

    active = {
        str(_customer_of(r).get("tenantId", "")).lower()
        for r in relationships
        if r.get("status") == "active"
    }
    idle = sorted(
        (c for c in customers if str(c.get("tenantId", "")).lower() not in active),
        key=lambda c: str(c.get("displayName", "")).casefold(),
    )
    print(f"\n{len(idle)} of {len(customers)} listed customer(s) have no active relationship:")
    for customer in idle:
        print(f"  {customer.get('displayName')}  {customer.get('tenantId')}")


if __name__ == "__main__":
    main()
