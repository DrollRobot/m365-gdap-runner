# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
#   "xlsxwriter>=3.3.1",
# ]
# ///
"""Example script: can users sign in to an enterprise app, in each tenant?

Finds every service principal whose display name or app ID matches one of the regular
expressions given, and reports whether users can sign in to it: the "Enabled for users to
sign-in?" switch, which is ``accountEnabled`` on the service principal. By default it only
reads::

    uv run service_principal_enabled.py "Contoso Cloud Backup" --include pilot

``--disable`` also turns sign-in off on every match where it is on::

    uv run service_principal_enabled.py "Contoso Cloud Backup" --include pilot --disable

``--delete`` instead deletes every match, enabled or not::

    uv run service_principal_enabled.py "^Contoso Cloud Backup$" --include pilot --delete

A deleted service principal stays in the tenant's deleted items for 30 days, where it can be
restored; after that it is gone for good.

A pattern is case-insensitive and matches anywhere in the value, so ``Backup`` matches every
app with Backup in its name; anchor it (``^Contoso Cloud Backup$``) to match one name
exactly. Run without ``--disable`` or ``--delete`` first to see what the patterns match.

Nothing is written while tenants run. After the last, one workbook goes in the current
folder, or the one ``--out`` names: a row per tenant, failed ones included, and a column per
matched app name.
Each cell is ``enabled`` or ``disabled`` (after the change, with ``--disable``), ``deleted``
with ``--delete``, or blank where the tenant has no such app. Apps that share a name in one
tenant share a cell, which lists each state they are in. ``--export-json`` also writes the
matched service principals, whole, as Graph returns them, in one JSON object keyed by tenant
ID; with ``--delete``, as they were before it, a record of what was deleted. Neither file
replaces another: where a name is taken, the file gets the first free ``_2``, ``_3``,
and so on.

A tenant's error does not stop the others, whether reading or changing apps: a tenant that
refuses, say for want of a role, is no reason to leave an app enabled in the rest.
Reading asks each tenant for ``Application.Read.All``, or ``Application.ReadWrite.All``
where that is the one granted. ``--disable`` and ``--delete`` ask for
``Application.ReadWrite.All`` alone, and changing or deleting an app also needs a GDAP role
such as Cloud Application Administrator.
"""

from __future__ import annotations

import argparse
import datetime
import io
import json
import re
from collections.abc import Mapping
from pathlib import Path
from typing import Any, BinaryIO

import xlsxwriter

from m365_gdap_runner import TenantContext, TenantResult

CLIENTS = ["graph"]

# Read access to service principals: either scope will do, the first preferred.
# ``--disable`` and ``--delete`` swap in Application.ReadWrite.All alone.
SCOPES: list[str | tuple[str, ...]] = [("Application.Read.All", "Application.ReadWrite.All")]

# Matched against each service principal's display name and app ID. The command line sets
# them before the run.
PATTERNS: list[re.Pattern[str]] = []

# Turn sign-in off where it is on. ``--disable`` sets it before the run.
DISABLE = False

# Delete every match. ``--delete`` sets it before the run.
DELETE = False

# Also write the matched service principals as JSON. ``--export-json`` sets it.
EXPORT_JSON = False

# One tenant's error is no reason to skip the others, whether reading, disabling or deleting.
CONTINUE_ON_ERROR = True

# Where the workbook and JSON go: the current folder. ``--out`` replaces it before the run.
OUT_DIR = Path()

# What a match needs from each service principal; ``--export-json`` reads matches whole.
PROPERTIES = "id,appId,displayName,accountEnabled"

# The largest page Graph returns of service principals.
PAGE_SIZE = 999

# Output files are named this, then today's date.
FILE_STEM = "ServicePrincipalEnabled"

# The workbook's font, and its size in points.
FONT = "Consolas"
FONT_SIZE = 11


def run(ctx: TenantContext) -> list[dict[str, Any]]:
    """Find the service principals that match, then disable or delete them as the flags say.

    ``--disable`` turns sign-in off on every match where it is on; ``--delete`` deletes every
    match.

    Args:
        ctx: The tenant context supplied by the runner.

    Returns:
        One service principal per match: the properties in :data:`PROPERTIES`, or the whole
        object with ``--export-json``. With ``--disable``, ``accountEnabled`` is as the run
        left it; with ``--delete``, each is as it was before. Empty when nothing matches.
    """
    apps = ctx.graph.get_all(
        "/servicePrincipals", params={"$select": PROPERTIES, "$top": PAGE_SIZE}
    )
    found = [app for app in apps if is_match(app)]
    if EXPORT_JSON:
        # Read before any change, so a replica that has not seen it yet cannot undo it here.
        found = [ctx.graph.get(f"/servicePrincipals/{app['id']}") for app in found]
    if DISABLE:
        for app in found:
            if _state(app) == "enabled":
                ctx.graph.patch(f"/servicePrincipals/{app['id']}", json={"accountEnabled": False})
                app["accountEnabled"] = False
    if DELETE:
        for app in found:
            ctx.graph.delete(f"/servicePrincipals/{app['id']}")
    return found


def report(results: list[TenantResult]) -> None:
    """Write the workbook, and with ``--export-json`` the JSON file.

    Args:
        results: Every tenant's result, in run order. A tenant that failed has its status
            and blank app cells.
    """
    if not results:
        print("No tenants ran; nothing written.")
        return
    cells = [_cells(result) for result in results]
    names = sorted({name for row in cells for name in row}, key=str.casefold)
    columns = ["Tenant", "Status", "TenantId", *names]
    rows = [
        [
            result.tenant.display_name,
            result.status,
            result.tenant.tenant_id,
            *(row.get(name) for name in names),
        ]
        for result, row in zip(results, cells, strict=True)
    ]
    options = {
        # Every cell's font, unless its format says otherwise. With no font scheme, Excel shows
        # FONT itself, not the theme's font, Calibri.
        "default_format_properties": {
            "font_name": FONT,
            "font_size": FONT_SIZE,
            "font_scheme": "none",
        },
        # Text stays text, never a formula or a link.
        "strings_to_formulas": False,
        "strings_to_urls": False,
    }
    data = io.BytesIO()
    with xlsxwriter.Workbook(data, options) as workbook:
        sheet = workbook.add_worksheet("Tenants")
        sheet.add_table(
            0,
            0,
            len(rows),
            len(columns) - 1,
            {
                "name": "Tenants",
                "style": "Table Style Medium 8",
                "columns": [{"header": column} for column in columns],
                "data": rows,
            },
        )
        sheet.freeze_panes(1, 0)
        for number, column in enumerate(zip(columns, *rows, strict=True)):
            longest = max(len(str(value)) for value in column if value is not None)
            sheet.set_column(number, number, min(longest + 2, 80))

    today = datetime.date.today().isoformat()
    path, file = _create(f"{FILE_STEM}_{today}", ".xlsx")
    with file:
        file.write(data.getvalue())
    print(f"Wrote {path}")
    if EXPORT_JSON:
        exported = {
            result.tenant.tenant_id: result.result
            for result in results
            if isinstance(result.result, list)
        }
        path, file = _create(f"{FILE_STEM}_{today}", ".json")
        with file:
            file.write(json.dumps(exported, indent=2).encode("utf-8"))
        print(f"Wrote {path}")


def is_match(app: Mapping[str, Any]) -> bool:
    """Tell whether a service principal's display name or app ID matches any of PATTERNS.

    Args:
        app: The service principal.

    Returns:
        True when any pattern is found anywhere in either value.
    """
    values = (str(app.get("displayName") or ""), str(app.get("appId") or ""))
    return any(regex.search(value) for regex in PATTERNS for value in values)


def pattern(text: str) -> re.Pattern[str]:
    """Compile a pattern from the command line, case-insensitive.

    Args:
        text: The regular expression.

    Returns:
        The compiled pattern.

    Raises:
        argparse.ArgumentTypeError: If it is not a valid regular expression.
    """
    try:
        return re.compile(text, re.IGNORECASE)
    except re.error as exc:
        raise argparse.ArgumentTypeError(
            f"not a valid regular expression: {text!r} ({exc})"
        ) from exc


def _cells(result: TenantResult) -> dict[str, str]:
    """Return one tenant's workbook cell per matched app name.

    Args:
        result: The tenant's result.

    Returns:
        App name to ``enabled``, ``disabled``, or both, comma-separated, when apps sharing
        the name differ; ``deleted`` with ``--delete``. Empty for a tenant that failed or
        matched nothing.
    """
    apps = result.result if isinstance(result.result, list) else []
    states: dict[str, set[str]] = {}
    for app in apps:
        # An app without a display name is shown by its app ID.
        name = str(app.get("displayName") or app.get("appId"))
        # A tenant whose run finished deleted every match it returned.
        states.setdefault(name, set()).add("deleted" if DELETE else _state(app))
    return {name: ", ".join(sorted(found)) for name, found in states.items()}


def _state(app: Mapping[str, Any]) -> str:
    """Return whether users can sign in to a service principal.

    Args:
        app: The service principal.

    Returns:
        ``disabled`` when ``accountEnabled`` is false, else ``enabled``: a missing value is
        treated as enabled, which is the default.
    """
    return "disabled" if app.get("accountEnabled") is False else "enabled"


def _create(stem: str, suffix: str) -> tuple[Path, BinaryIO]:
    """Create a new file in :data:`OUT_DIR`, never one that is already there.

    Each name is created only if it does not exist, so an earlier run's file is never
    replaced.

    Args:
        stem: The file name before its suffix, such as ``ServicePrincipalEnabled_2026-10-29``.
        suffix: The file name's suffix, such as ``.xlsx``.

    Returns:
        Where the file is, and the file, open for writing bytes. The name is *stem* and
        *suffix*, or where that is taken, the first free one with ``_2``, ``_3``, and so on
        between them.
    """
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"{stem}{suffix}"
    number = 1
    while True:
        try:
            return path, path.open("xb")
        except FileExistsError:
            number += 1
            path = OUT_DIR / f"{stem}_{number}{suffix}"


if __name__ == "__main__":
    from m365_gdap_runner import add_arguments, main

    parser = argparse.ArgumentParser(
        description="Report user sign-in on matching enterprise apps, and optionally turn it "
        "off or delete the apps."
    )
    parser.add_argument(
        "patterns",
        nargs="+",
        type=pattern,
        metavar="PATTERN",
        help="Regular expression matched, case-insensitive, against each service principal's "
        "display name and app ID.",
    )
    change = parser.add_mutually_exclusive_group()
    change.add_argument(
        "--disable",
        action="store_true",
        help="Turn user sign-in off on every match where it is on.",
    )
    change.add_argument(
        "--delete",
        action="store_true",
        help="Delete every match. Deleted items keep it for 30 days, where it can be restored.",
    )
    parser.add_argument(
        "--export-json",
        action="store_true",
        help="Also write the matched service principals, whole, keyed by tenant ID.",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=OUT_DIR,
        metavar="DIR",
        help="Folder for the workbook and JSON (default: the current folder).",
    )
    add_arguments(parser)
    args = parser.parse_args()
    PATTERNS = args.patterns
    DISABLE = args.disable
    DELETE = args.delete
    EXPORT_JSON = args.export_json
    OUT_DIR = args.out
    if DISABLE or DELETE:
        SCOPES = ["Application.ReadWrite.All"]
    raise SystemExit(main(__name__, args))
