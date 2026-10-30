# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.2.0",
#   "xlsxwriter>=3.3.1",
# ]
# ///
"""Example script: the message trace records that match a filter, in each tenant.

Runs ``Get-MessageTraceV2`` through Exchange Online's admin API,
``/adminapi/beta/<tenant>/InvokeCommand``. That API takes the cmdlet's parameters as a JSON
object, and the filter is that object, sent as given::

    uv run exchange_message_trace.py '{"SenderAddress": "admin@contoso.com"}' --include pilot
    uv run exchange_message_trace.py '{"Status": ["Failed", "Quarantined"]}'
    uv run exchange_message_trace.py '{}' --include pilot --size 100

PowerShell before 7.3 strips the inner double quotes: put a backslash before each one there.

Where Exchange refuses ``Get-MessageTraceV2`` with a 403, as it does in GCC High, the script runs
``Get-MessageTrace`` instead, with the same filter. Exchange refuses a parameter the cmdlet does
not take, and the workbook says so: only ``Get-MessageTraceV2`` takes ``Subject`` and
``SubjectFilterType``, and only ``Get-MessageTrace`` takes ``Page``.

Writes one JSON file per tenant with matches, and one workbook of match counts per tenant
with the filter in its title row, in the current folder or ``--out``. A file never replaces
another: where its name is taken, it gets the first free ``_2``, ``_3``, and so on.

Needs a GDAP role such as Exchange Administrator. A record is one message to one recipient.

Each tenant gets one query, for up to ``--size`` records: 5000 by default, the most a query
returns. It goes as ``ResultSize`` to ``Get-MessageTraceV2`` and as ``PageSize`` to
``Get-MessageTrace``, unless the filter sets its own. A count of exactly that many in the
workbook means there may be more.

Common filters. Several values for one parameter go in a list, and one value goes as text::

    # Who sent it, and who it went to.
    {"SenderAddress": "admin@contoso.com"}
    {"RecipientAddress": ["user@contoso.com", "other@contoso.com"]}

    # Date range: both dates or neither. One alone is refused, with no reason given. Neither is
    # the last 48 hours. At most 10 days at once, from the last 90 days, or the last 10 with
    # Get-MessageTrace. A date with no time is the start of that day, so give the end a time to
    # include its day.
    {"StartDate": "2026-10-20T00:00:00Z", "EndDate": "2026-10-29T23:59:59Z"}

    # Delivery status: Delivered, Expanded, Failed, FilteredAsSpam, GettingStatus, Pending
    # or Quarantined.
    {"Status": ["Failed", "Quarantined"]}

    # Subject, and how to match it, which it needs: Contains, StartsWith or EndsWith.
    {"Subject": "Weekly digest", "SubjectFilterType": "StartsWith"}

    # One message, by its Message-ID header or its trace ID.
    {"MessageId": "d9683b4c-127b-413a-ae2e-fa7dfb32c69d@contoso.com"}
    {"MessageTraceId": "00000000-0000-0000-0000-000000000000"}

    # The server that sent it, and the one it went to.
    {"FromIP": "203.0.113.10"}
    {"ToIP": "203.0.113.20"}

    # All messages from a specific domain over the past 2 days:
    {"SenderAddress": "*@gmx.de", "StartDate": "2026-10-28T18:00:00Z",
     "EndDate": "2026-10-30T23:59:59Z"}

Microsoft allows each tenant 100 queries in 5 minutes. A tenant refused for that, or for
anything else, has the reason in the workbook and no JSON file; run it again.

Microsoft's docs: the two cmdlets, and the message trace feature::

    https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-messagetracev2
    https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-messagetrace
    https://learn.microsoft.com/en-us/exchange/monitoring/trace-an-email-message/message-trace-modern-eac
"""

from __future__ import annotations

import datetime
import io
import json
import logging
import re
from http import HTTPStatus
from pathlib import Path
from typing import Any, BinaryIO

import xlsxwriter
from azure_auth import InvokeCommandError

from m365_gdap_runner import TenantContext, TenantResult

logger = logging.getLogger(__name__)

# Exchange only: no Graph, so no SCOPES.
CLIENTS = ["exchange"]

# A report only reads, so one tenant's error is no reason to stop the others.
CONTINUE_ON_ERROR = True

# The cmdlet's parameters, as a JSON object. The command line sets it before the run.
FILTER = "{}"

# The most records to ask each tenant for: the most a query returns. ``--size`` replaces it.
SIZE = 5000

# Where the JSON files and the workbook go: the current folder. ``--out`` replaces it.
OUT_DIR = Path()

# Output files are named this, then the tenant for a JSON file, then today's date.
FILE_STEM = "ExchangeMessageTrace"

# The workbook's columns, in order, below the title row.
COLUMNS = ["Tenant", "TenantId", "Status", "Records", "Error"]

# Widest a column is made to fit its longest value, in characters.
MAX_COLUMN_WIDTH = 80

# The workbook's font, its size in points, and the title row's size.
FONT = "Consolas"
FONT_SIZE = 11
TITLE_FONT_SIZE = 20


def run(ctx: TenantContext) -> int:
    """Read the tenant's matching records, and write them as JSON when there are any.

    Args:
        ctx: The tenant context supplied by the runner.

    Returns:
        How many records match.
    """
    records = _read_records(ctx)
    if not records:
        logger.info("No records match")
        return 0
    path, file = _create(f"{FILE_STEM}_{_file_safe(ctx.tenant.display_name)}_{_today()}", ".json")
    with file:
        file.write(json.dumps(records, indent=2).encode("utf-8"))
    logger.info("%d records match; wrote %s", len(records), path)
    return len(records)


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
        sheet.write(0, 0, FILTER, title)
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
        # The title runs on across the empty cells to its right, so it is left out of the widths.
        for number, column in enumerate(zip(COLUMNS, *rows, strict=True)):
            longest = max(len(str(value)) for value in column if value is not None)
            sheet.set_column(number, number, min(longest + 2, MAX_COLUMN_WIDTH))

    path, file = _create(f"{FILE_STEM}_{_today()}", ".xlsx")
    with file:
        file.write(data.getvalue())
    print(f"Wrote {path}")


def _read_records(ctx: TenantContext) -> list[dict[str, Any]]:
    """Run ``Get-MessageTraceV2``, or ``Get-MessageTrace`` where Exchange refuses it with a 403.

    A 403 with no error body is how Exchange answers a cmdlet it does not know.

    Args:
        ctx: The tenant context supplied by the runner.

    Returns:
        The records, whole, as the cmdlet returns them.

    Raises:
        InvokeCommandError: If Exchange refuses ``Get-MessageTraceV2`` for any other reason, or
            refuses both cmdlets. The error for both opens with the first cmdlet's refusal,
            because the runner cuts a long message to its first sentence.
    """
    parameters = json.loads(FILTER)
    try:
        return ctx.exchange.run("Get-MessageTraceV2", **{"ResultSize": SIZE, **parameters})
    except InvokeCommandError as refused:
        if refused.status != HTTPStatus.FORBIDDEN:
            raise
    logger.info("Exchange refused Get-MessageTraceV2; trying Get-MessageTrace")
    try:
        return ctx.exchange.run("Get-MessageTrace", **{"PageSize": SIZE, **parameters})
    except InvokeCommandError as error:
        message = f"Exchange refused Get-MessageTraceV2 with a 403, then {error}"
        raise InvokeCommandError(message, status=error.status) from error


def _create(stem: str, suffix: str) -> tuple[Path, BinaryIO]:
    """Create a new file in :data:`OUT_DIR`, never one that is already there.

    Each name is created only if it does not exist, so neither an earlier run's file nor one
    another tenant of the same name just wrote is replaced.

    Args:
        stem: The file name before its suffix, such as ``ExchangeMessageTrace_2026-10-29``.
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
        description="Write the message trace records that match a filter, for each tenant."
    )
    parser.add_argument(
        "filter",
        metavar="FILTER",
        help="Get-MessageTraceV2's parameters as a JSON object, such as "
        '\'{"SenderAddress": "admin@contoso.com"}\'.',
    )
    parser.add_argument(
        "--size",
        type=int,
        default=SIZE,
        metavar="N",
        help=f"The most records to read from each tenant, up to 5000 (default: {SIZE}).",
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
    SIZE = args.size
    OUT_DIR = args.out
    raise SystemExit(main(__name__, args))
