# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
# ]
# ///
"""Example script: stop users signing in to the ConnectWise Cloud Backup enterprise app.

Sets "Enabled for users to sign-in?" to No on the app in each tenant, which is
``accountEnabled`` false on its service principal. Try it first with ``--dry-run``, which
reads every tenant and changes nothing::

    uv run disable_connectwise_backup_sign_in.py --include pilot --dry-run

then without it::

    uv run disable_connectwise_backup_sign_in.py --include pilot

A dry run keeps going after a tenant's error, so it shows every tenant; a real run stops at
the first, rather than repeat a broken change.

The app is found by its exact display name; a tenant without it is left alone. ``report()``
prints a line per tenant: what was done, or would be in a dry run, or why the tenant failed.

The dry run asks for the same Graph scope as the real run, so a tenant that has not consented
to ``Application.ReadWrite.All`` fails the dry run too. Changing the app also needs a GDAP role
such as Cloud Application Administrator; a dry run cannot check that.
"""

from __future__ import annotations

import logging

from m365_gdap_runner import TenantContext, TenantResult

logger = logging.getLogger(__name__)

CLIENTS = ["graph"]

# Write access to service principals. A dry run asks for it too, so it checks consent.
SCOPES = ["Application.ReadWrite.All"]

# The enterprise app's display name, matched exactly.
APP_NAME = "ConnectWise Cloud Backup"

# Report what would change instead of changing it. ``--dry-run`` sets it before the run.
DRY_RUN = False

# A dry run only reads, so it keeps going after a tenant's error. ``--dry-run`` sets it too.
CONTINUE_ON_ERROR = False


def run(ctx: TenantContext) -> list[dict[str, str]]:
    """Disable user sign-in on every service principal named :data:`APP_NAME`.

    Args:
        ctx: The tenant context supplied by the runner.

    Returns:
        One row per app found, with ``appId``, ``id`` and ``action``: ``disabled``,
        ``would disable`` (dry run) or ``already disabled``. Empty when the tenant has no
        such app.
    """
    apps = ctx.graph.get_all(
        "/servicePrincipals",
        params={
            "$filter": f"displayName eq '{APP_NAME}'",
            "$select": "id,appId,displayName,accountEnabled",
        },
    )
    if not apps:
        logger.info("%s not found", APP_NAME)
    rows = []
    for app in apps:
        # A missing value is treated as enabled, which is the default.
        if app.get("accountEnabled") is False:
            action = "already disabled"
        elif DRY_RUN:
            action = "would disable"
        else:
            ctx.graph.patch(f"/servicePrincipals/{app['id']}", json={"accountEnabled": False})
            action = "disabled"
        logger.info("%s (appId %s): %s", APP_NAME, app["appId"], action)
        rows.append({"appId": str(app["appId"]), "id": str(app["id"]), "action": action})
    return rows


def report(results: list[TenantResult]) -> None:
    """Print one line per app found, or per tenant without one, or that failed.

    Args:
        results: Every tenant's result, in run order. A tenant that failed has its status
            and error, and no rows.
    """
    if DRY_RUN:
        print("Dry run: nothing was changed.")
    lines: list[tuple[str, str]] = []
    for result in results:
        name = result.tenant.display_name
        rows = result.result if isinstance(result.result, list) else []
        if result.status != "ok":
            lines.append((name, f"{result.status}: {result.error}"))
        elif not rows:
            lines.append((name, "not found"))
        else:
            lines += [(name, f"{row['action']} (appId {row['appId']})") for row in rows]
    width = max((len(name) for name, _ in lines), default=0)
    for name, outcome in lines:
        print(f"{name:<{width}}  {outcome}")


if __name__ == "__main__":
    import argparse

    from m365_gdap_runner import add_arguments, main

    parser = argparse.ArgumentParser(
        description=f"Disable user sign-in on the {APP_NAME} enterprise app in each tenant."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would change in each tenant, and change nothing.",
    )
    add_arguments(parser)
    args = parser.parse_args()
    DRY_RUN = CONTINUE_ON_ERROR = args.dry_run
    raise SystemExit(main(__name__, args))
