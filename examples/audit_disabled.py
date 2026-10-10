"""Example payload: is mailbox auditing switched off tenant-wide?

Shows a payload that uses Exchange Online instead of Graph. Run it across tenants::

    uv run m365-gdap-runner examples/audit_disabled.py --format csv --output audit.csv

or directly against one tenant::

    uv run python examples/audit_disabled.py --tenant Contoso
"""

from __future__ import annotations

from m365_gdap_runner import TenantContext

# Exchange only: no Graph, so no SCOPES. Exchange Online PowerShell is pre-authorised for
# its own cmdlets; consent is not per scope there.
CLIENTS = ["exchange"]


def run(ctx: TenantContext) -> dict[str, object]:
    """Return the tenant's organisation-wide audit switch.

    Args:
        ctx: The tenant context supplied by the runner or ``standalone()``.

    Returns:
        One row with ``name`` (the organisation's name in Exchange) and
        ``auditDisabled``. ``True`` means mailbox auditing is off for every mailbox,
        whatever each mailbox's own settings say.
    """
    (config,) = ctx.exchange.run("Get-OrganizationConfig")
    return {
        "name": str(config.get("Name") or ""),
        "auditDisabled": bool(config.get("AuditDisabled")),
    }


if __name__ == "__main__":
    import json

    from m365_gdap_runner import standalone

    with standalone(clients=CLIENTS) as ctx:
        print(json.dumps(run(ctx), indent=2))
