# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
#   "openpyxl>=3.1.5",
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

A pattern is case-insensitive and matches anywhere in the value, so ``Backup`` matches every
app with Backup in its name; anchor it (``^Contoso Cloud Backup$``) to match one name
exactly. Run without ``--disable`` first to see what the patterns match.

Nothing is written while tenants run. After the last, one workbook goes in the current
folder, or the one ``--out`` names: a row per tenant, failed ones included, and a column per
matched app name.
Each cell is ``enabled`` or ``disabled`` (after the change, with ``--disable``), or blank
where the tenant has no such app. Apps that share a name in one tenant share a cell, which
lists each state they are in. ``--export-json`` also writes the matched service principals,
whole, as Graph returns them, in one JSON object keyed by tenant ID.

A tenant's error does not stop the others, with or without ``--disable``: a tenant that
refuses, say for want of a role, is no reason to leave an app enabled in the rest.
``--disable`` asks for ``Application.ReadWrite.All`` in place of ``Application.Read.All``, and
changing an app also needs a GDAP role such as Cloud Application Administrator.
"""

from __future__ import annotations

import argparse
import datetime
import json
import re
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from m365_gdap_runner import TenantContext, TenantResult

CLIENTS = ["graph"]

# Read access to service principals. ``--disable`` swaps in Application.ReadWrite.All.
SCOPES = ["Application.Read.All"]

# Matched against each service principal's display name and app ID. The command line sets
# them before the run.
PATTERNS: list[re.Pattern[str]] = []

# Turn sign-in off where it is on. ``--disable`` sets it before the run.
DISABLE = False

# Also write the matched service principals as JSON. ``--export-json`` sets it.
EXPORT_JSON = False

# One tenant's error is no reason to skip the others, whether reading or disabling.
CONTINUE_ON_ERROR = True

# Where the workbook and JSON go: the current folder. ``--out`` replaces it before the run.
OUT_DIR = Path()

# What a match needs from each service principal; ``--export-json`` reads matches whole.
PROPERTIES = "id,appId,displayName,accountEnabled"

# The largest page Graph returns of service principals.
PAGE_SIZE = 999

# Output files are named this, then today's date.
FILE_STEM = "ServicePrincipalEnabled"


def run(ctx: TenantContext) -> list[dict[str, Any]]:
    """Find the service principals that match, and turn sign-in off on them with ``--disable``.

    Args:
        ctx: The tenant context supplied by the runner.

    Returns:
        One service principal per match, with ``accountEnabled`` as it is after the run: the
        properties in :data:`PROPERTIES`, or the whole object with ``--export-json``. Empty
        when nothing matches.
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
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None  # noqa: S101 (a new workbook always has a sheet)
    sheet.title = "Tenants"
    sheet.append(["Tenant", "Status", "TenantId", *names])
    for result, row in zip(results, cells, strict=True):
        sheet.append(
            [
                result.tenant.display_name,
                result.status,
                result.tenant.tenant_id,
                *(row.get(name) for name in names),
            ]
        )
    table = Table(displayName="Tenants", ref=sheet.dimensions)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium8", showRowStripes=True)
    sheet.add_table(table)
    sheet.freeze_panes = "A2"
    for number, column in enumerate(sheet.iter_cols(values_only=True), start=1):
        longest = max(len(str(value)) for value in column if value is not None)
        sheet.column_dimensions[get_column_letter(number)].width = min(longest + 2, 80)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    today = datetime.date.today().isoformat()
    path = OUT_DIR / f"{FILE_STEM}_{today}.xlsx"
    workbook.save(path)
    print(f"Wrote {path}")
    if EXPORT_JSON:
        exported = {
            result.tenant.tenant_id: result.result
            for result in results
            if isinstance(result.result, list)
        }
        path = OUT_DIR / f"{FILE_STEM}_{today}.json"
        path.write_text(json.dumps(exported, indent=2), encoding="utf-8")
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
        the name differ. Empty for a tenant that failed or matched nothing.
    """
    apps = result.result if isinstance(result.result, list) else []
    states: dict[str, set[str]] = {}
    for app in apps:
        # An app without a display name is shown by its app ID.
        name = str(app.get("displayName") or app.get("appId"))
        states.setdefault(name, set()).add(_state(app))
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


if __name__ == "__main__":
    from m365_gdap_runner import add_arguments, main

    parser = argparse.ArgumentParser(
        description="Report, and optionally disable, user sign-in on matching enterprise apps."
    )
    parser.add_argument(
        "patterns",
        nargs="+",
        type=pattern,
        metavar="PATTERN",
        help="Regular expression matched, case-insensitive, against each service principal's "
        "display name and app ID.",
    )
    parser.add_argument(
        "--disable",
        action="store_true",
        help="Turn user sign-in off on every match where it is on.",
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
    EXPORT_JSON = args.export_json
    OUT_DIR = args.out
    if DISABLE:
        SCOPES = ["Application.ReadWrite.All"]
    raise SystemExit(main(__name__, args))
