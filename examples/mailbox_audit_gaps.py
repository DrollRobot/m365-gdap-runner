# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
#   "openpyxl>=3.1.5",
# ]
# ///
"""Example script: mailboxes whose audit action sets lag Microsoft's defaults.

A port of a PowerShell report, and an example of both kinds of report. ``run()`` writes each
tenant's workbook as soon as that tenant is done: its audit switches, the mailboxes missing
actions, and the audit bypass associations. After the last tenant, ``report()`` writes one
summary workbook: a row per tenant, failed ones included, and every tenant's gaps on one
sheet. ``--out`` is the script's own flag, beside the library's::

    uv run mailbox_audit_gaps.py --include pilot --out reports

The reference action set per logon type (Admin, Delegate, Owner) is derived from the
mailboxes still enrolled in Microsoft's managed defaults, then every mailbox that left a
default set is compared with it. The tenant-wide audit switches and audit bypass associations
are reported too, since either can invalidate the per-mailbox findings.

Mailboxes are read the way ``Get-EXOMailbox -PropertySets Audit`` reads them, through
Exchange's REST endpoint (``ctx.exchange.mailboxes()``): only the audit properties and the
three the report names, which is much faster than ``Get-Mailbox`` on a large tenant.

One thing differs from the PowerShell version: ``Get-AdminAuditLogConfig`` always reaches
Exchange here, because this script declares no Security & Compliance client, so
``UnifiedAuditLogIngestionEnabled`` can be trusted.

Note:
    ``MailItemsAccessed`` and ``Send`` appear in the default sets only on Audit Premium
    licensed mailboxes. In a mixed-license tenant, E5 mailboxes contribute those actions to
    the derived reference, so a customised non-E5 mailbox is reported as missing them. That is
    a licensing finding, not a configuration one; no ``Set-Mailbox`` change adds them.
"""

from __future__ import annotations

import datetime
import logging
import re
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from m365_gdap_runner import TenantContext, TenantResult

logger = logging.getLogger(__name__)

# Exchange only: no Graph, so no SCOPES.
CLIENTS = ["exchange"]

# A report only reads, so one tenant's error is no reason to stop the others.
CONTINUE_ON_ERROR = True

# The logon types a mailbox audits, each with its own action set and default-set flag.
LOGON_TYPES = ("Admin", "Delegate", "Owner")

# Mailbox kinds the report covers.
RECIPIENT_TYPES = ["UserMailbox", "SharedMailbox", "RoomMailbox", "EquipmentMailbox"]

# Properties read beside the Audit property set, which carries each logon type's actions.
AUDIT_PROPERTIES = ["DefaultAuditSet", "UserPrincipalName", "RecipientTypeDetails"]

# Where the workbooks go. ``--out`` replaces it before the run starts.
OUT_DIR = Path("reports")

# The look ImportExcel gave the PowerShell report: a striped table with filters on each sheet.
TABLE_STYLE = "TableStyleMedium8"

# Widest a column is made to fit its longest value, in characters.
MAX_COLUMN_WIDTH = 80


@dataclass(frozen=True)
class TenantAudit:
    """What ``run()`` found in one tenant, for ``report()``.

    Attributes:
        tenant: The tenant's initial ``*.onmicrosoft.com`` domain.
        audit_disabled: Whether mailbox auditing is off tenant-wide.
        unified_audit_log: Whether records reach the unified audit log.
        mailboxes: How many mailboxes were read.
        gaps: One row per mailbox missing actions, as the workbook's Audit Gaps sheet has it.
        bypassed: One row per principal whose mailbox access is not audited.
        workbook: The tenant's workbook.
    """

    tenant: str
    audit_disabled: bool
    unified_audit_log: bool
    mailboxes: int
    gaps: list[dict[str, str]]
    bypassed: list[dict[str, str]]
    workbook: Path


def run(ctx: TenantContext) -> TenantAudit:
    """Audit one tenant and write its workbook.

    Args:
        ctx: The tenant context supplied by the runner.

    Returns:
        What was found, and where the workbook is.
    """
    exchange = ctx.exchange
    tenant = _initial_domain(exchange.run("Get-AcceptedDomain"))

    (org_config,) = exchange.run("Get-OrganizationConfig")
    tenant = tenant or str(org_config.get("Name") or ctx.tenant.display_name)
    audit_disabled = bool(org_config.get("AuditDisabled"))
    if audit_disabled:
        logger.warning("AuditDisabled is True: mailbox auditing is off tenant-wide.")

    (audit_config,) = exchange.run("Get-AdminAuditLogConfig")
    unified_audit_log = bool(audit_config.get("UnifiedAuditLogIngestionEnabled"))
    if not unified_audit_log:
        logger.warning("UnifiedAuditLogIngestionEnabled is False: nothing reaches the audit log.")

    mailboxes = exchange.mailboxes(
        PropertySet="Audit", RecipientTypeDetails=RECIPIENT_TYPES, select=AUDIT_PROPERTIES
    )
    logger.info("%d mailboxes retrieved", len(mailboxes))
    gaps = audit_gaps(mailboxes)

    logger.info("Checking audit bypass associations (enumerates every principal; can be slow)")
    bypassed = sorted(
        (
            {"Name": str(item.get("Name") or ""), "Identity": str(item.get("Identity") or "")}
            for item in exchange.run("Get-MailboxAuditBypassAssociation")
            if item.get("AuditBypassEnabled")
        ),
        key=lambda row: row["Name"].casefold(),
    )

    workbook = _workbook()
    settings = {
        "Tenant": tenant,
        "Collected": datetime.datetime.now().replace(microsecond=0),
        "AuditDisabled": audit_disabled,
        "UnifiedAuditLogIngestionEnabled": unified_audit_log,
    }
    _add_sheet(workbook, "Tenant", "TenantSettings", [settings])
    if gaps:
        _add_sheet(workbook, "Audit Gaps", "AuditGaps", gaps)
    if bypassed:
        _add_sheet(workbook, "Audit Bypass", "AuditBypass", bypassed)
    path = _save(workbook, f"MailboxAuditGaps_{_file_safe(tenant)}")
    logger.info(
        "Wrote %s: %d mailboxes, %d with gaps, %d bypassed",
        path,
        len(mailboxes),
        len(gaps),
        len(bypassed),
    )
    return TenantAudit(
        tenant, audit_disabled, unified_audit_log, len(mailboxes), gaps, bypassed, path
    )


def report(results: list[TenantResult]) -> None:
    """Write the run's summary workbook: one row per tenant, and every tenant's gaps.

    Args:
        results: Every tenant's result, in run order. A tenant that failed has its status
            and error and no counts.
    """
    tenants: list[dict[str, object]] = []
    gaps: list[dict[str, object]] = []
    for result in results:
        audit = result.result if isinstance(result.result, TenantAudit) else None
        tenants.append(
            {
                "Tenant": result.tenant.display_name,
                "TenantId": result.tenant.tenant_id,
                "Status": result.status,
                "AuditDisabled": audit.audit_disabled if audit else None,
                "UnifiedAuditLogIngestionEnabled": audit.unified_audit_log if audit else None,
                "Mailboxes": audit.mailboxes if audit else None,
                "WithGaps": len(audit.gaps) if audit else None,
                "Bypassed": len(audit.bypassed) if audit else None,
                "Workbook": audit.workbook.name if audit else "",
                "Error": result.error or "",
            }
        )
        if audit:
            gaps += [{"Tenant": result.tenant.display_name, **row} for row in audit.gaps]
    workbook = _workbook()
    _add_sheet(workbook, "Tenants", "Tenants", tenants)
    if gaps:
        _add_sheet(workbook, "Audit Gaps", "AuditGaps", gaps)
    path = _save(workbook, "MailboxAuditGaps_summary")
    print(f"Wrote {path}")


def audit_gaps(mailboxes: Iterable[Mapping[str, Any]]) -> list[dict[str, str]]:
    """Find mailboxes missing audit actions, measured against the tenant's own defaults.

    Args:
        mailboxes: The audit read of every mailbox.

    Returns:
        One row per mailbox with missing actions, sorted by user principal name:
        ``UserPrincipalName``, ``RecipientTypeDetails`` and ``MissingActions``
        (comma-separated).
    """
    mailboxes = list(mailboxes)
    reference = reference_sets(mailboxes)
    rows: list[dict[str, str]] = []
    for mailbox in mailboxes:
        defaults = _names(mailbox.get("DefaultAuditSet"))
        missing: set[str] = set()
        for logon_type in LOGON_TYPES:
            if logon_type not in defaults:
                missing |= reference[logon_type] - _names(mailbox.get(f"Audit{logon_type}"))
        if missing:
            rows.append(
                {
                    "UserPrincipalName": str(mailbox.get("UserPrincipalName") or ""),
                    "RecipientTypeDetails": str(mailbox.get("RecipientTypeDetails") or ""),
                    "MissingActions": ",".join(sorted(missing)),
                }
            )
    return sorted(rows, key=lambda row: row["UserPrincipalName"].casefold())


def reference_sets(mailboxes: Iterable[Mapping[str, Any]]) -> dict[str, frozenset[str]]:
    """Derive the default action set per logon type from the mailboxes still enrolled in it.

    Args:
        mailboxes: The audit read of every mailbox.

    Returns:
        Logon type to the union of actions audited by the mailboxes whose ``DefaultAuditSet``
        includes that type. Empty, with a warning, when no mailbox is enrolled.
    """
    reference = {logon_type: set[str]() for logon_type in LOGON_TYPES}
    for mailbox in mailboxes:
        for logon_type in _names(mailbox.get("DefaultAuditSet")) & set(LOGON_TYPES):
            reference[logon_type] |= _names(mailbox.get(f"Audit{logon_type}"))
    for logon_type, actions in reference.items():
        if not actions:
            logger.warning(
                "No mailbox is enrolled in the default %s set; its gaps cannot be detected.",
                logon_type,
            )
    return {logon_type: frozenset(actions) for logon_type, actions in reference.items()}


def _initial_domain(accepted_domains: Iterable[Mapping[str, Any]]) -> str:
    """Return the tenant's initial ``*.onmicrosoft.com`` domain, or ``""``.

    Args:
        accepted_domains: ``Get-AcceptedDomain`` output.

    Returns:
        The ``DomainName`` of the domain flagged ``InitialDomain``.
    """
    for domain in accepted_domains:
        if domain.get("InitialDomain"):
            return str(domain.get("DomainName") or "")
    return ""


def _names(value: object) -> frozenset[str]:
    """Read a multi-valued cmdlet property as a set of names.

    Exchange returns these as JSON arrays; a comma-separated string is accepted too, since
    that is how PowerShell renders them.

    Args:
        value: The property value: a list, a string, or None.

    Returns:
        The non-empty names.
    """
    if value is None:
        return frozenset()
    items = value if isinstance(value, list | tuple) else str(value).split(",")
    return frozenset(name for name in (str(item).strip() for item in items) if name)


def _file_safe(tenant: str) -> str:
    """Turn a tenant's domain into the part of a file name that names it.

    Args:
        tenant: The tenant's initial domain, such as ``contoso.onmicrosoft.com``.

    Returns:
        Its first label with anything but letters, digits, ``_`` and ``-`` replaced.
    """
    return re.sub(r"[^\w-]", "_", tenant.split(".")[0])


def _workbook() -> Workbook:
    """Return a workbook with no sheets, ready for :func:`_add_sheet`.

    Returns:
        The workbook.
    """
    workbook = Workbook()
    for sheet in list(workbook.worksheets):
        workbook.remove(sheet)
    return workbook


def _add_sheet(
    workbook: Workbook, title: str, table_name: str, rows: Sequence[Mapping[str, object]]
) -> None:
    """Add a sheet holding *rows* as a styled table, filters on, top row frozen.

    Args:
        workbook: The workbook.
        title: The sheet's name.
        table_name: The table's name, unique in the workbook and without spaces.
        rows: At least one row; the first row's keys are the columns.
    """
    sheet = workbook.create_sheet(title)
    columns = list(rows[0])
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column) for column in columns])
    last_cell = f"{get_column_letter(len(columns))}{len(rows) + 1}"
    table = Table(displayName=table_name, ref=f"A1:{last_cell}")
    table.tableStyleInfo = TableStyleInfo(name=TABLE_STYLE, showRowStripes=True)
    sheet.add_table(table)
    sheet.freeze_panes = "A2"
    for number, column in enumerate(columns, start=1):
        longest = max(len(_text(value)) for value in [column, *(row.get(column) for row in rows)])
        width = min(longest + 2, MAX_COLUMN_WIDTH)
        sheet.column_dimensions[get_column_letter(number)].width = width


def _text(value: object) -> str:
    """Show a cell value as the text a column has to fit.

    Args:
        value: The cell value.

    Returns:
        The text; empty for None.
    """
    if value is None:
        return ""
    if isinstance(value, datetime.datetime):
        return value.isoformat(sep=" ")
    return str(value)


def _save(workbook: Workbook, stem: str) -> Path:
    """Save a workbook in :data:`OUT_DIR`, named with today's date.

    Args:
        workbook: The workbook.
        stem: The file name before the date.

    Returns:
        Where it was saved.
    """
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"{stem}_{datetime.date.today().isoformat()}.xlsx"
    workbook.save(path)
    return path


if __name__ == "__main__":
    import argparse

    from m365_gdap_runner import add_arguments, main

    parser = argparse.ArgumentParser(
        description="Mailboxes whose audit action sets lag Microsoft's defaults."
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=OUT_DIR,
        metavar="DIR",
        help=f"Folder for the workbooks (default: {OUT_DIR}).",
    )
    add_arguments(parser)
    args = parser.parse_args()
    OUT_DIR = args.out
    raise SystemExit(main(__name__, args))
