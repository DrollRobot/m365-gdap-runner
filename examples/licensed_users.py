# /// script
# requires-python = ">=3.14"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
# ]
# ///
"""Example script: licensed users in each tenant.

Copy this file as the starting point for new scripts. The header above tells uv what the
script needs, so ``uv run`` builds its environment; run it across a group of tenants::

    uv run licensed_users.py --group pilot --format csv --output users.csv

or against one tenant::

    uv run licensed_users.py --tenant Contoso

In this repository, ``uv run python examples/licensed_users.py ...`` runs it against the
working copy instead of the release the header names.
"""

from __future__ import annotations

from m365_gdap_runner import TenantContext

# Which of ctx.graph, ctx.exchange, ctx.ipps and ctx.azure this script uses. The runner
# signs in once per client at the partner tenant and checks each in every customer tenant.
CLIENTS = ["graph"]

# Delegated Graph scopes this script needs in every customer tenant.
SCOPES = ["User.Read.All"]


def run(ctx: TenantContext) -> list[dict[str, str]]:
    """Return every licensed user in the tenant.

    Args:
        ctx: The tenant context supplied by the runner.

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
    from m365_gdap_runner import main

    raise SystemExit(main(__name__))
