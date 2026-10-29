# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
#   "xlsxwriter>=3.3.1",
# ]
# ///
"""Example script: mailboxes whose audit action sets lag Microsoft's defaults.

A port of a PowerShell report, and an example of both kinds of report. ``run()`` writes each
tenant's workbook as soon as that tenant is done: its audit switches, the mailboxes missing
actions, and the audit bypass associations. After the last tenant, ``report()`` writes one
summary workbook: a row per tenant, failed ones included, and every tenant's gaps on one
sheet. ``--out`` is the script's own flag, beside the library's::

    uv run mailbox_audit_gaps.py --include pilot --out reports

A workbook never replaces another: where its name is taken, it gets the first free
``_2``, ``_3``, and so on.

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
import io
import logging
import re
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, BinaryIO

import xlsxwriter

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
TABLE_STYLE = "Table Style Medium 8"

# Widest a column is made to fit its longest value, in characters.
MAX_COLUMN_WIDTH = 80

# The workbook's font, and its size in points.
FONT = "Consolas"
FONT_SIZE = 11


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

    settings: dict[str, object] = {
        "Tenant": tenant,
        "Collected": datetime.datetime.now().replace(microsecond=0),
        "AuditDisabled": audit_disabled,
        "UnifiedAuditLogIngestionEnabled": unified_audit_log,
    }
    sheets: list[tuple[str, str, Sequence[Mapping[str, object]]]] = [
        ("Tenant", "TenantSettings", [settings])
    ]
    if gaps:
        sheets.append(("Audit Gaps", "AuditGaps", gaps))
    if bypassed:
        sheets.append(("Audit Bypass", "AuditBypass", bypassed))
    path = _save(f"MailboxAuditGaps_{_file_safe(tenant)}", sheets)
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
    sheets: list[tuple[str, str, Sequence[Mapping[str, object]]]] = [
        ("Tenants", "Tenants", tenants)
    ]
    if gaps:
        sheets.append(("Audit Gaps", "AuditGaps", gaps))
    path = _save("MailboxAuditGaps_summary", sheets)
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


def _add_sheet(
    workbook: xlsxwriter.Workbook,
    title: str,
    table_name: str,
    rows: Sequence[Mapping[str, object]],
) -> None:
    """Add a sheet holding *rows* as a styled table, filters on, top row frozen.

    Args:
        workbook: The workbook.
        title: The sheet's name.
        table_name: The table's name, unique in the workbook and without spaces.
        rows: At least one row; the first row's keys are the columns.
    """
    sheet = workbook.add_worksheet(title)
    columns = list(rows[0])
    data = [[row.get(column) for column in columns] for row in rows]
    sheet.add_table(
        0,
        0,
        len(data),
        len(columns) - 1,
        {
            "name": table_name,
            "style": TABLE_STYLE,
            "columns": [{"header": column} for column in columns],
            "data": data,
        },
    )
    sheet.freeze_panes(1, 0)
    for number, column in enumerate(columns):
        longest = max(len(_text(value)) for value in [column, *(row.get(column) for row in rows)])
        sheet.set_column(number, number, min(longest + 2, MAX_COLUMN_WIDTH))


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


def _save(stem: str, sheets: Sequence[tuple[str, str, Sequence[Mapping[str, object]]]]) -> Path:
    """Write a workbook in :data:`OUT_DIR`, named with today's date, never over another file.

    Args:
        stem: The file name before the date.
        sheets: Each sheet's name, its table's name and its rows, as :func:`_add_sheet` takes
            them.

    Returns:
        Where it was written.
    """
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
        # A date and time shows as one, not as the number Excel keeps it as.
        "default_date_format": "yyyy-mm-dd hh:mm:ss",
    }
    data = io.BytesIO()
    with xlsxwriter.Workbook(data, options) as workbook:
        for title, table_name, rows in sheets:
            _add_sheet(workbook, title, table_name, rows)
    path, file = _create(f"{stem}_{datetime.date.today().isoformat()}", ".xlsx")
    with file:
        file.write(data.getvalue())
    return path


def _create(stem: str, suffix: str) -> tuple[Path, BinaryIO]:
    """Create a new file in :data:`OUT_DIR`, never one that is already there.

    Each name is created only if it does not exist, so an earlier run's file is never
    replaced.

    Args:
        stem: The file name before its suffix, such as ``MailboxAuditGaps_summary_2026-10-29``.
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
