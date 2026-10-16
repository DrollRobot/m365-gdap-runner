# /// script
# requires-python = ">=3.14"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
# ]
# ///
"""Example script: mailboxes whose audit action sets lag Microsoft's defaults.

A port of a PowerShell report. The script connected, collected, analysed, exported an
``.xlsx`` and read it back; here the runner connects and writes, so the script is only the
middle two steps and returns findings as flat rows. A tenant with nothing wrong returns no
rows, which is what makes the merged CSV a report of gaps only::

    uv run mailbox_audit_gaps.py --group pilot --format csv

or against one tenant::

    uv run mailbox_audit_gaps.py --tenant Contoso

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

import logging
from collections.abc import Iterable, Mapping
from typing import Any

from m365_gdap_runner import TenantContext

logger = logging.getLogger(__name__)

# Exchange only: no Graph, so no SCOPES.
CLIENTS = ["exchange"]

# The logon types a mailbox audits, each with its own action set and default-set flag.
LOGON_TYPES = ("Admin", "Delegate", "Owner")

# Mailbox kinds the report covers.
RECIPIENT_TYPES = ["UserMailbox", "SharedMailbox", "RoomMailbox", "EquipmentMailbox"]

# Properties read beside the Audit property set, which carries each logon type's actions.
AUDIT_PROPERTIES = ["DefaultAuditSet", "UserPrincipalName", "RecipientTypeDetails"]

# What the ``finding`` column holds. Every row has the same columns so the CSV stays flat;
# ``recipientTypeDetails`` and ``missingActions`` are blank except for ``audit_gap``.
AUDIT_DISABLED = "audit_disabled"
UNIFIED_AUDIT_LOG_OFF = "unified_audit_log_off"
AUDIT_GAP = "audit_gap"
AUDIT_BYPASS = "audit_bypass"


def run(ctx: TenantContext) -> list[dict[str, str]]:
    """Return every audit finding in the tenant.

    Args:
        ctx: The tenant context supplied by the runner.

    Returns:
        One row per finding with ``finding``, ``identity``, ``recipientTypeDetails`` and
        ``missingActions`` (comma-separated). Empty when the tenant is clean.
    """
    exchange = ctx.exchange
    tenant = _initial_domain(exchange.run("Get-AcceptedDomain"))
    rows: list[dict[str, str]] = []

    (org_config,) = exchange.run("Get-OrganizationConfig")
    tenant = tenant or str(org_config.get("Name") or ctx.tenant.display_name)
    if org_config.get("AuditDisabled"):
        logger.warning("AuditDisabled is True: mailbox auditing is off tenant-wide.")
        rows.append(_row(AUDIT_DISABLED, tenant))

    (audit_config,) = exchange.run("Get-AdminAuditLogConfig")
    if not audit_config.get("UnifiedAuditLogIngestionEnabled"):
        logger.warning("UnifiedAuditLogIngestionEnabled is False: nothing reaches the audit log.")
        rows.append(_row(UNIFIED_AUDIT_LOG_OFF, tenant))

    mailboxes = exchange.mailboxes(
        PropertySet="Audit", RecipientTypeDetails=RECIPIENT_TYPES, select=AUDIT_PROPERTIES
    )
    logger.info("%d mailboxes retrieved", len(mailboxes))
    rows.extend(audit_gaps(mailboxes))

    logger.info("Checking audit bypass associations (enumerates every principal; can be slow)")
    for item in exchange.run("Get-MailboxAuditBypassAssociation"):
        if item.get("AuditBypassEnabled"):
            rows.append(_row(AUDIT_BYPASS, str(item.get("Identity") or item.get("Name") or "")))
    return rows


def audit_gaps(mailboxes: Iterable[Mapping[str, Any]]) -> list[dict[str, str]]:
    """Find mailboxes missing audit actions, measured against the tenant's own defaults.

    Args:
        mailboxes: The audit read of every mailbox.

    Returns:
        One ``audit_gap`` row per mailbox with missing actions, sorted by identity.
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
                _row(
                    AUDIT_GAP,
                    str(mailbox.get("UserPrincipalName") or ""),
                    recipient_type=str(mailbox.get("RecipientTypeDetails") or ""),
                    missing_actions=",".join(sorted(missing)),
                )
            )
    return sorted(rows, key=lambda row: row["identity"].casefold())


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


def _row(
    finding: str, identity: str, *, recipient_type: str = "", missing_actions: str = ""
) -> dict[str, str]:
    """Build one output row.

    Args:
        finding: One of the finding kinds.
        identity: The tenant domain, mailbox UPN or bypassed principal.
        recipient_type: ``RecipientTypeDetails`` for a mailbox finding.
        missing_actions: Comma-separated actions for an ``audit_gap`` finding.

    Returns:
        The row.
    """
    return {
        "finding": finding,
        "identity": identity,
        "recipientTypeDetails": recipient_type,
        "missingActions": missing_actions,
    }


if __name__ == "__main__":
    from m365_gdap_runner import main

    raise SystemExit(main(__name__))
