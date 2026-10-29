# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
#   "xlsxwriter>=3.3.1",
# ]
# ///
"""Example script: the Entra audit logs that match a filter, in each tenant.

Reads ``/auditLogs/directoryAudits`` from Graph's beta endpoint with the OData ``$filter``
given on the command line, so one script answers any question the audit log can::

    uv run entra_audit_logs.py "activityDisplayName eq 'Add member to role'" --include pilot

``--v1`` reads v1.0 instead. Its audit events leave out one property beta has, ``userAgent``,
which Microsoft's reference doesn't list as one to filter on::

    uv run entra_audit_logs.py "activityDateTime ge 2026-10-01T00:00:00Z" --v1

Each tenant with matches gets one JSON file of its audit events, whole, as Graph returns
them. After the last tenant, one workbook counts the matches in every tenant, failed ones
included, under a title row that holds the filter. Both go in the current folder, or the one
``--out`` names. A file never replaces another: where its name is taken, it gets the first
free ``_2``, ``_3``, and so on.

Reading the audit log asks each tenant for ``AuditLog.Read.All``, and needs a GDAP role
such as Reports Reader, Global Reader or Security Reader. Unlike the sign-in log, it needs
no Entra ID P1 or P2 licence.

Common filters
--------------

Join filters with ``and`` and ``or``, and group them with parentheses. Text goes in single
quotes; dates and times go without. The audit log takes ``eq``, ``ge``, ``le`` and
``startswith``, each on only the properties Microsoft lists for it; ``ne``, ``contains`` and
``endswith`` are not among them.

**Date range.** ``activityDateTime`` takes ``ge``, ``le`` and ``eq``, not ``gt`` or ``lt``.
Entra keeps audit events for seven days in a tenant without P1 or P2, and 30 days in one
with it::

    activityDateTime ge 2026-10-01T00:00:00Z
    activityDateTime ge 2026-10-01T00:00:00Z and activityDateTime le 2026-10-07T23:59:59Z

**Activity.** ``activityDisplayName`` is the name in the Activity column of the audit log
in the Entra admin center. ``eq`` for the whole name, ``startswith`` for its start::

    activityDisplayName eq 'Add member to role'
    activityDisplayName eq 'Consent to application'
    activityDisplayName eq 'Add service principal credentials'
    activityDisplayName eq 'Update Conditional Access policy'
    startswith(activityDisplayName, 'Reset password')

**Who did it.** ``initiatedBy`` holds the user or the app that did it: a user by user name,
object ID or display name, an app by app ID or display name. ``startswith`` takes the user
name alone::

    initiatedBy/user/userPrincipalName eq 'admin@contoso.com'
    startswith(initiatedBy/user/userPrincipalName, 'admin')
    initiatedBy/user/id eq '00000000-0000-0000-0000-000000000000'
    initiatedBy/app/appId eq '00000000-0000-0000-0000-000000000000'
    initiatedBy/app/displayName eq 'Contoso Backup'

**What it changed.** ``targetResources`` lists the objects the activity changed, and
``any`` looks through them. ``eq`` for an object ID or display name, ``startswith`` for the
start of a name::

    targetResources/any(t: t/id eq '00000000-0000-0000-0000-000000000000')
    targetResources/any(t: t/displayName eq 'Contoso Admins')
    targetResources/any(t: startswith(t/displayName, 'Contoso'))

A user it changed often has a user name and no display name. Microsoft's reference doesn't
list ``userPrincipalName`` as one to filter on, so Graph may refuse it::

    targetResources/any(t: t/userPrincipalName eq 'user@contoso.com')

**Service.** ``loggedByService`` is the part of Entra that logged the event, as the Service
column of the audit log shows it::

    loggedByService eq 'Core Directory'
    loggedByService eq 'Conditional Access'
    loggedByService eq 'Privileged Identity Management'
    loggedByService eq 'Self-service Password Management'

**Category and result.** Microsoft's reference doesn't list ``category`` or ``result`` as
ones to filter on, so Graph may refuse them. Each event's JSON has both either way.
``result`` is ``success``, ``failure`` or ``timeout``::

    category eq 'RoleManagement'
    category eq 'ApplicationManagement'
    result eq 'failure'

**One operation.** ``correlationId`` is shared by every event one operation logged, across
services::

    correlationId eq '00000000-0000-0000-0000-000000000000'

A filter that joins two, for this month's role assignments::

    activityDateTime ge 2026-10-01T00:00:00Z and activityDisplayName eq 'Add member to role'

Microsoft's documentation
-------------------------

Listing audit events, with its permissions, roles and examples:
https://learn.microsoft.com/en-us/graph/api/directoryaudit-list?view=graph-rest-beta

Every audit event property and the operators it takes, in beta and in v1.0:
https://learn.microsoft.com/en-us/graph/api/resources/directoryaudit?view=graph-rest-beta
https://learn.microsoft.com/en-us/graph/api/resources/directoryaudit?view=graph-rest-1.0

The ``$filter`` syntax:
https://learn.microsoft.com/en-us/graph/filter-query-parameter

Every activity the audit log records, by service and category:
https://learn.microsoft.com/en-us/entra/identity/monitoring-health/reference-audit-activities

How long Entra keeps audit events:
https://learn.microsoft.com/en-us/entra/identity/monitoring-health/reference-reports-data-retention
"""

from __future__ import annotations

import datetime
import io
import json
import logging
import re
from pathlib import Path
from typing import BinaryIO

import xlsxwriter

from m365_gdap_runner import TenantContext, TenantResult

logger = logging.getLogger(__name__)

CLIENTS = ["graph"]

# The one scope Graph asks of a delegated sign-in that lists audit events.
SCOPES = ["AuditLog.Read.All"]

# A report only reads, so one tenant's error is no reason to stop the others.
CONTINUE_ON_ERROR = True

# The OData filter sent as ``$filter``. The command line sets it before the run.
FILTER = ""

# Read Graph's v1.0 endpoint instead of beta. ``--v1`` sets it before the run.
V1 = False

# Where the JSON files and the workbook go: the current folder. ``--out`` replaces it.
OUT_DIR = Path()

# Output files are named this, then the tenant for a JSON file, then today's date.
FILE_STEM = "EntraAuditLogs"

# The workbook's columns, in order, below the title row.
COLUMNS = ["Tenant", "TenantId", "Status", "AuditEvents", "Error"]

# Widest a column is made to fit its longest value, in characters.
MAX_COLUMN_WIDTH = 80

# The workbook's font, its size in points, and the title row's size.
FONT = "Consolas"
FONT_SIZE = 11
TITLE_FONT_SIZE = 20


def run(ctx: TenantContext) -> int:
    """Read the tenant's matching audit events, and write them as JSON when there are any.

    Args:
        ctx: The tenant context supplied by the runner.

    Returns:
        How many audit events match.
    """
    events = ctx.graph.get_all(_collection(ctx), params={"$filter": FILTER})
    if not events:
        logger.info("No audit events match")
        return 0
    path, file = _create(f"{FILE_STEM}_{_file_safe(ctx.tenant.display_name)}_{_today()}", ".json")
    with file:
        file.write(json.dumps(events, indent=2).encode("utf-8"))
    logger.info("%d audit events match; wrote %s", len(events), path)
    return len(events)


def report(results: list[TenantResult]) -> None:
    """Write one workbook with each tenant's count of matches, under the filter as its title.

    Args:
        results: Every tenant's result, in run order. A tenant that failed has its status
            and error, and a blank count.
    """
    if not results:
        print("No tenants ran; no workbook written.")
        return
    rows = [
        [
            result.tenant.display_name,
            result.tenant.tenant_id,
            result.status,
            result.result if isinstance(result.result, int) else None,
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
        title = workbook.add_format({"bold": True, "font_size": TITLE_FONT_SIZE})
        sheet.merge_range(0, 0, 0, len(COLUMNS) - 1, FILTER, title)
        sheet.add_table(
            1,
            0,
            len(rows) + 1,
            len(COLUMNS) - 1,
            {
                "name": "Tenants",
                "style": "Table Style Medium 8",
                "columns": [{"header": column} for column in COLUMNS],
                "data": rows,
            },
        )
        sheet.freeze_panes(2, 0)
        # The title row spans the table, so it is left out of the widths.
        for number, column in enumerate(zip(COLUMNS, *rows, strict=True)):
            longest = max(len(str(value)) for value in column if value is not None)
            sheet.set_column(number, number, min(longest + 2, MAX_COLUMN_WIDTH))

    path, file = _create(f"{FILE_STEM}_{_today()}", ".xlsx")
    with file:
        file.write(data.getvalue())
    print(f"Wrote {path}")


def _collection(ctx: TenantContext) -> str:
    """Return the audit events collection to read: beta by its whole URL, v1.0 by its path.

    ``ctx.graph`` sends a path to v1.0, and an absolute URL on its own host as it is.

    Args:
        ctx: The tenant context, for the Graph address of the tenant's cloud.

    Returns:
        The beta collection's URL, or with ``--v1`` ``/auditLogs/directoryAudits``.
    """
    if V1:
        return "/auditLogs/directoryAudits"
    return f"{ctx.auth.tenant.cloud.graph}/beta/auditLogs/directoryAudits"


def _create(stem: str, suffix: str) -> tuple[Path, BinaryIO]:
    """Create a new file in :data:`OUT_DIR`, never one that is already there.

    Each name is created only if it does not exist, so neither an earlier run's file nor one
    another tenant of the same name just wrote is replaced.

    Args:
        stem: The file name before its suffix, such as ``EntraAuditLogs_2026-10-29``.
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


def _file_safe(name: str) -> str:
    """Turn a tenant's display name into the part of a file name that names it.

    Args:
        name: The display name, such as ``Contoso Ltd.``.

    Returns:
        The name with anything but letters, digits, ``_`` and ``-`` replaced by ``_``.
    """
    return re.sub(r"[^\w-]", "_", name)


def _today() -> str:
    """Return today's date as it goes in a file name.

    Returns:
        The date, such as ``2026-10-29``.
    """
    return datetime.date.today().isoformat()


if __name__ == "__main__":
    import argparse

    from m365_gdap_runner import add_arguments, main

    parser = argparse.ArgumentParser(
        description="Write the Entra audit logs that match a filter, for each tenant."
    )
    parser.add_argument(
        "filter",
        metavar="FILTER",
        help="OData $filter for /auditLogs/directoryAudits, such as "
        "\"activityDisplayName eq 'Add member to role'\".",
    )
    parser.add_argument(
        "--v1",
        action="store_true",
        help="Read Graph's v1.0 endpoint instead of beta.",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=OUT_DIR,
        metavar="DIR",
        help="Folder for the JSON files and the workbook (default: the current folder).",
    )
    add_arguments(parser)
    args = parser.parse_args()
    FILTER = args.filter
    V1 = args.v1
    OUT_DIR = args.out
    raise SystemExit(main(__name__, args))
