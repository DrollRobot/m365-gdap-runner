# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
#   "openpyxl>=3.1.5",
# ]
# ///
"""Example script: is the unified audit log on in each tenant?

``run()`` reads one tenant's ``UnifiedAuditLogIngestionEnabled`` and warns when it is off.
After the last tenant, ``report()`` writes one workbook with a row per tenant, failed ones
included, in the folder ``--out`` names::

    uv run unified_audit_log.py --include all --out reports

A workbook never replaces another: where its name is taken, it gets the first free
``_2``, ``_3``, and so on.

``Get-AdminAuditLogConfig`` reaches Exchange here, because this script declares no Security &
Compliance client. That matters: the Security & Compliance version of the cmdlet always
reports ``UnifiedAuditLogIngestionEnabled`` as False.
"""

from __future__ import annotations

import datetime
import logging
from pathlib import Path
from typing import BinaryIO

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from m365_gdap_runner import TenantContext, TenantResult

logger = logging.getLogger(__name__)

# Exchange only: no Graph, so no SCOPES.
CLIENTS = ["exchange"]

# A report only reads, so one tenant's error is no reason to stop the others.
CONTINUE_ON_ERROR = True

# Where the workbook goes. ``--out`` replaces it before the run starts.
OUT_DIR = Path("reports")

# The workbook's columns, in order.
COLUMNS = ["Tenant", "TenantId", "Status", "UnifiedAuditLogIngestionEnabled", "Error"]


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
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None  # noqa: S101 (a new workbook always has a sheet)
    sheet.title = "Tenants"
    sheet.append(COLUMNS)
    for result in results:
        sheet.append(
            [
                result.tenant.display_name,
                result.tenant.tenant_id,
                result.status,
                result.result if isinstance(result.result, bool) else None,
                result.error,
            ]
        )
    table = Table(displayName="Tenants", ref=sheet.dimensions)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium8", showRowStripes=True)
    sheet.add_table(table)
    sheet.freeze_panes = "A2"
    for number, column in enumerate(sheet.iter_cols(values_only=True), start=1):
        longest = max(len(str(value)) for value in column if value is not None)
        sheet.column_dimensions[get_column_letter(number)].width = min(longest + 2, 80)

    path, file = _create(f"UnifiedAuditLog_{datetime.date.today().isoformat()}", ".xlsx")
    with file:
        workbook.save(file)
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
        "--out",
        type=Path,
        default=OUT_DIR,
        metavar="DIR",
        help=f"Folder for the workbook (default: {OUT_DIR}).",
    )
    add_arguments(parser)
    args = parser.parse_args()
    OUT_DIR = args.out
    raise SystemExit(main(__name__, args))
