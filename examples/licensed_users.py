"""Example payload: licensed users in each tenant.

Copy this file as the starting point for new payloads. Run it across a group of tenants::

    uv run m365-gdap-runner examples/licensed_users.py --group pilot --format csv --output users.csv

or directly against one tenant::

    uv run python examples/licensed_users.py --tenant Contoso
"""

from __future__ import annotations

from m365_gdap_runner import TenantContext

# Which of ctx.graph, ctx.exchange, ctx.ipps and ctx.azure this payload uses. The runner
# signs in once per client at the partner tenant and checks each in every customer tenant.
CLIENTS = ["graph"]

# Delegated Graph scopes this payload needs in every customer tenant.
SCOPES = ["User.Read.All"]


def run(ctx: TenantContext) -> list[dict[str, str]]:
    """Return every licensed user in the tenant.

    Args:
        ctx: The tenant context supplied by the runner or ``standalone()``.

    Returns:
        One row per licensed user with ``userPrincipalName``, ``displayName``,
        and ``skuIds`` (semicolon-separated, so rows stay flat for CSV output).
    """
    users = ctx.graph.get_all(
        "/users",
        params={
            "$filter": "assignedLicenses/$count ne 0",
            "$count": "true",
            "$select": "userPrincipalName,displayName,assignedLicenses",
        },
        headers={"ConsistencyLevel": "eventual"},
    )
    return [
        {
            "userPrincipalName": str(user.get("userPrincipalName") or ""),
            "displayName": str(user.get("displayName") or ""),
            "skuIds": ";".join(
                sorted(
                    str(license["skuId"])
                    for license in user.get("assignedLicenses") or []
                    if license.get("skuId")
                )
            ),
        }
        for user in users
    ]


if __name__ == "__main__":
    import json

    from m365_gdap_runner import standalone

    with standalone(SCOPES, CLIENTS) as ctx:
        print(json.dumps(run(ctx), indent=2))
