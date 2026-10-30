# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.2.0",
#   "xlsxwriter>=3.3.1",
# ]
# ///
"""Example script: is the unified audit log on in each tenant?

``run()`` reads one tenant's ``UnifiedAuditLogIngestionEnabled`` and warns when it is off.
After the last tenant, ``report()`` writes one workbook with a row per tenant, failed ones
included, in the current folder or ``--output``::

    uv run unified_audit_log.py --include all
    uv run unified_audit_log.py --include all --output reports

A workbook never replaces another: where its name is taken, it gets the first free
``_2``, ``_3``, and so on.

``Get-AdminAuditLogConfig`` reaches Exchange here, because this script declares no Security &
Compliance client. That matters: the Security & Compliance version of the cmdlet always
reports ``UnifiedAuditLogIngestionEnabled`` as False.
"""

from __future__ import annotations

import datetime
import io
import logging
from pathlib import Path
from typing import BinaryIO

import xlsxwriter

from m365_gdap_runner import TenantContext, TenantResult

logger = logging.getLogger(__name__)

# Exchange only: no Graph, so no SCOPES.
CLIENTS = ["exchange"]

# A report only reads, so one tenant's error is no reason to stop the others.
CONTINUE_ON_ERROR = True

# Where the workbook goes: the current folder. ``--output`` replaces it.
OUT_DIR = Path()

# The workbook's columns, in order.
COLUMNS = ["Tenant", "TenantId", "Status", "UnifiedAuditLogIngestionEnabled", "Error"]

# The workbook's font, and its size in points.
FONT = "Consolas"
FONT_SIZE = 11


def run(ctx: TenantContext) -> bool:
    """Read whether audit records reach the tenant's unified audit log.

    Args:
        ctx: The tenant context supplied by the runner.

    Returns:
        ``UnifiedAuditLogIngestionEnabled``.
    """
    (config,) = ctx.exchange.run("Get-AdminAuditLogConfig")
    enabled = bool(config.get("UnifiedAuditLogIngestionEnabled"))
    if enabled:
        logger.info("Unified audit log is on")
    else:
        logger.warning("Unified audit log is off (UnifiedAuditLogIngestionEnabled is False)")
    return enabled


def report(results: list[TenantResult]) -> None:
    """Write one workbook with a row per tenant, as a filterable table.

    Args:
        results: Every tenant's result, in run order. A tenant that failed has its status
            and error, and a blank ``UnifiedAuditLogIngestionEnabled``.
    """
    if not results:
        print("No tenants ran; no workbook written.")
        return
    rows = [
        [
            result.tenant.display_name,
            result.tenant.tenant_id,
            result.status,
            result.result if isinstance(result.result, bool) else None,
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
            len(COLUMNS) - 1,
            {
                "name": "Tenants",
                "style": "Table Style Medium 8",
                "columns": [{"header": column} for column in COLUMNS],
                "data": rows,
            },
        )
        sheet.freeze_panes(1, 0)
        for number, column in enumerate(zip(COLUMNS, *rows, strict=True)):
            longest = max(len(str(value)) for value in column if value is not None)
            sheet.set_column(number, number, min(longest + 2, 80))

    path, file = _create(f"UnifiedAuditLog_{datetime.date.today().isoformat()}", ".xlsx")
    with file:
        file.write(data.getvalue())
    print(f"Wrote {path}")


def _create(stem: str, suffix: str) -> tuple[Path, BinaryIO]:
    """Create a new file in :data:`OUT_DIR`, never one that is already there.

    Each name is created only if it does not exist, so an earlier run's file is never
    replaced.

    Args:
        stem: The file name before its suffix, such as ``UnifiedAuditLog_2026-10-29``.
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
    import argparse

    from m365_gdap_runner import add_arguments, main

    parser = argparse.ArgumentParser(description="Is the unified audit log on in each tenant?")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=OUT_DIR,
        metavar="DIR",
        help="Folder for the workbook (default: the current folder).",
    )
    add_arguments(parser)
    args = parser.parse_args()
    OUT_DIR = args.output
    raise SystemExit(main(__name__, args))
