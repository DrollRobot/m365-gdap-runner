# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
#   "xlsxwriter>=3.3.1",
# ]
# ///
"""Which Microsoft Graph scopes does an app hold, in each tenant?

Signs in as :data:`APP_ID`, Microsoft Graph Command Line Tools, unless ``--client-id`` names
another app, and checks the scopes it holds in each tenant for the ones in
:data:`REQUESTED_SCOPES`; ``--scopes`` replaces the list. By default it only reads, asking
for no scope in particular, so every tenant where the app holds any scope reports::

    uv run app_scopes.py --include pilot

``--write`` asks for every requested scope at sign-in instead::

    uv run app_scopes.py --include pilot --write

Where a tenant's grant lacks some of them and the app holds
``DelegatedPermissionGrant.ReadWrite.All`` there, they are added to the app's tenant-wide
grant with no prompt (``GRANT_WHEN_POSSIBLE``; ``--no-grant-when-possible`` turns it off).
Every user of the app in that tenant then gets them, each still limited by their own roles.
Any other tenant that lacks some gets a direct sign-in, as its ``m365_username``, whose
consent screen grants them. Tick "Consent on behalf of your organization" there, or only
that account gets them and the partner sign-in still lacks them. A tenant with no
``m365_username`` is ``auth_failed``, naming the scopes it lacks.
Entra can take about a minute to put newly granted scopes in tokens, so a tenant consented
during the run may still show them missing; run the report again to check.

Nothing is written while tenants run. After the last, one workbook goes in the current
folder, or the one ``--out`` names: a row per tenant, failed ones included, with whether the
app holds every requested scope, the ones it lacks, every scope it holds, and why a tenant
failed. ``--export-json`` also writes each tenant's scopes as JSON. ``--import-json`` builds
the workbook from such a file instead, with no sign-in, and checks it against the scopes
requested now, so one export can be checked against other lists. Neither file replaces
another: where a name is taken, the file gets the first free ``_2``, ``_3``, and so on.
"""

from __future__ import annotations

import argparse
import datetime
import io
import json
from collections.abc import Sequence
from pathlib import Path
from typing import Any, BinaryIO

import xlsxwriter

from m365_gdap_runner import Tenant, TenantContext, TenantResult

CLIENTS = ["graph"]

# Reading the scopes needs none in particular. ``--write`` replaces it with REQUESTED_SCOPES.
SCOPES: list[str] = []

# The Microsoft Graph scopes to check for. ``--scopes`` replaces them before the run.
REQUESTED_SCOPES = [
    "Application.Read.All",
    "Application.ReadWrite.All",
    "AuditLog.Read.All",
    "AuditLogsQuery.Read.All",
    "CrossTenantInformation.ReadBasic.All",
    "DelegatedPermissionGrant.Read.All",
    "Device.Read.All",
    "DeviceManagementManagedDevices.Read.All",
    "DeviceManagementManagedDevices.ReadWrite.All",
    "Directory.AccessAsUser.All",
    "Directory.Read.All",
    "Domain.Read.All",
    "Organization.Read.All",
    "Policy.Read.All",
    "User-LifeCycleInfo.Read.All",
    "User-PasswordProfile.Read.All",
    "User-PasswordProfile.ReadWrite.All",
    "User.Read.All",
    "User.ReadWrite.All",
    "UserAuthenticationMethod.Read.All",
    "UserAuthenticationMethod.ReadWrite.All",
    "Mail.ReadBasic.Shared",
    "Policy.Read.ConditionalAccess",
    "SecurityAlert.Read.All",
    "SecurityIncident.Read.All",
]

# The app to sign in as: Microsoft Graph Command Line Tools. It is the default of
# ``--client-id``, so it wins over the client_id setting; ``--client-id`` replaces it.
APP_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

# Add the scopes a tenant lacks to the app's grant where it holds
# DelegatedPermissionGrant.ReadWrite.All there. ``--write`` sets it.
GRANT_WHEN_POSSIBLE = False

# Also write each tenant's result as JSON. ``--export-json`` sets it.
EXPORT_JSON = False

# One tenant's error is no reason to skip the others, whether reading or consenting.
CONTINUE_ON_ERROR = True

# Where the workbook and JSON go: the current folder. ``--out`` replaces it before the run.
OUT_DIR = Path()

# Output files are named this, then today's date.
FILE_STEM = "AppScopes"

# The workbook's font, and its size in points.
FONT = "Consolas"
FONT_SIZE = 11


def run(ctx: TenantContext) -> dict[str, Any]:
    """Return the Graph scopes the app holds in the tenant.

    Args:
        ctx: The tenant context supplied by the runner.

    Returns:
        ``appId``, the app signed in as, and ``scopes``, every Graph scope it holds there,
        sorted.
    """
    scopes = sorted(ctx.graph.granted_scopes(), key=str.casefold)
    return {"appId": ctx.graph.client_id, "scopes": scopes}


def report(results: list[TenantResult]) -> None:
    """Write the workbook, and with ``--export-json`` the JSON file.

    Args:
        results: Every tenant's result, in run order. A tenant that failed has its status
            and error, and blank scope cells.
    """
    if not results:
        print("No tenants ran; nothing written.")
        return
    columns = ["Tenant", "Status", "TenantId", "HasScopes", "Missing", "Scopes", "Error"]
    rows = [
        [
            result.tenant.display_name,
            result.status,
            result.tenant.tenant_id,
            *_cells(result),
            result.error,
        ]
        for result in results
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
        exported = [
            {
                "tenant": result.tenant.display_name,
                "tenantId": result.tenant.tenant_id,
                "status": result.status,
                "error": result.error,
                "result": result.result,
            }
            for result in results
        ]
        path, file = _create(f"{FILE_STEM}_{today}", ".json")
        with file:
            file.write(json.dumps(exported, indent=2).encode("utf-8"))
        print(f"Wrote {path}")


def load_results(path: Path) -> list[TenantResult]:
    """Read the tenants' results back from a file ``--export-json`` wrote.

    Args:
        path: The JSON file.

    Returns:
        One result per tenant, in the order the file lists them.

    Raises:
        OSError: If the file cannot be read.
        ValueError: If it is not JSON.
        KeyError: If an entry lacks a field the export writes.
        TypeError: If the file is not a list of entries.
    """
    entries = json.loads(path.read_text(encoding="utf-8"))
    return [
        TenantResult(
            Tenant(entry["tenantId"], entry["tenant"]),
            entry["status"],
            entry["result"],
            entry["error"],
        )
        for entry in entries
    ]


def missing(requested: Sequence[str], held: Sequence[str]) -> list[str]:
    """Return the requested scopes the app lacks. Scope names are compared ignoring case.

    Args:
        requested: The scopes to check for.
        held: The scopes the app holds.

    Returns:
        The requested scopes not held, in the order requested.
    """
    have = {scope.lower() for scope in held}
    return [scope for scope in requested if scope.lower() not in have]


def _cells(result: TenantResult) -> list[object]:
    """Return one tenant's HasScopes, Missing and Scopes cells.

    Args:
        result: The tenant's result.

    Returns:
        Whether the app holds every requested scope, then the scopes it lacks and every
        scope it holds, each space-separated, or blank when there are none. All three are
        blank for a tenant that failed.
    """
    if not isinstance(result.result, dict):
        return [None, None, None]
    scopes = result.result["scopes"]
    lacking = missing(REQUESTED_SCOPES, scopes)
    return [not lacking, " ".join(lacking) or None, " ".join(scopes) or None]


def _create(stem: str, suffix: str) -> tuple[Path, BinaryIO]:
    """Create a new file in :data:`OUT_DIR`, never one that is already there.

    Each name is created only if it does not exist, so an earlier run's file is never
    replaced.

    Args:
        stem: The file name before its suffix, such as ``AppScopes_2026-10-29``.
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
        description="Report, and optionally consent to, the Microsoft Graph scopes an app holds."
    )
    parser.add_argument(
        "--scopes",
        nargs="+",
        default=REQUESTED_SCOPES,
        metavar="SCOPE",
        help="Microsoft Graph scopes to check for (default: REQUESTED_SCOPES in this script).",
    )
    parser.add_argument(
        "--write",
        action="store_true",
        help="Ask for every requested scope at sign-in: add them to the app's grant where it "
        "can, else ask the tenant's own account to consent.",
    )
    parser.add_argument(
        "--export-json",
        action="store_true",
        help="Also write each tenant's result as JSON, for --import-json.",
    )
    parser.add_argument(
        "--import-json",
        type=Path,
        metavar="FILE",
        help="Build the workbook from a file --export-json wrote, without signing in.",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=OUT_DIR,
        metavar="DIR",
        help="Folder for the workbook and JSON (default: the current folder).",
    )
    add_arguments(parser)
    parser.set_defaults(client_id=APP_ID)
    args = parser.parse_args()
    REQUESTED_SCOPES = args.scopes
    EXPORT_JSON = args.export_json
    OUT_DIR = args.out
    if args.import_json:
        clashes = ("--write", "--export-json", "--include", "--exclude")
        given = [flag for flag in clashes if getattr(args, flag[2:].replace("-", "_"))]
        if given:
            parser.error(f"--import-json runs no tenants, so it takes no {', '.join(given)}.")
        try:
            imported = load_results(args.import_json)
        except (OSError, ValueError, KeyError, TypeError) as exc:
            parser.error(f"cannot read {args.import_json}: {exc!r}")
        report(imported)
        raise SystemExit(0)
    if args.write:
        SCOPES = REQUESTED_SCOPES
        GRANT_WHEN_POSSIBLE = True
    raise SystemExit(main(__name__, args))
