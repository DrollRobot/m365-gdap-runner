# /// script
# requires-python = ">=3.14"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
# ]
# ///
"""Example script: is mailbox auditing switched off tenant-wide?

Shows a script that uses Exchange Online instead of Graph. Run it across tenants::

    uv run audit_disabled.py --all --format csv --output audit.csv

or against one tenant::

    uv run audit_disabled.py --tenant Contoso
"""

from __future__ import annotations

from m365_gdap_runner import TenantContext

# Exchange only: no Graph, so no SCOPES. Exchange Online PowerShell is pre-authorised for
# its own cmdlets; consent is not per scope there.
CLIENTS = ["exchange"]


def run(ctx: TenantContext) -> dict[str, object]:
    """Return the tenant's organisation-wide audit switch.

    Args:
        ctx: The tenant context supplied by the runner.

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
    from m365_gdap_runner import main

    raise SystemExit(main(__name__))
