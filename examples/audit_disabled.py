# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
# ]
# ///
"""Example script: is mailbox auditing switched off tenant-wide?

Shows a script that uses Exchange Online instead of Graph, and one with no ``report()``: each
tenant's log line is the report, so a tenant with auditing off stands out as a warning::

    uv run audit_disabled.py --all

or against one tenant::

    uv run audit_disabled.py --tenant Contoso
"""

from __future__ import annotations

import logging

from m365_gdap_runner import TenantContext

logger = logging.getLogger(__name__)

# Exchange only: no Graph, so no SCOPES. Exchange Online PowerShell is pre-authorised for
# its own cmdlets; consent is not per scope there.
CLIENTS = ["exchange"]

# A report only reads, so one tenant's error is no reason to stop the others.
CONTINUE_ON_ERROR = True


def run(ctx: TenantContext) -> None:
    """Log whether mailbox auditing is off for every mailbox in the tenant.

    ``AuditDisabled`` turns auditing off whatever each mailbox's own settings say.

    Args:
        ctx: The tenant context supplied by the runner.
    """
    (config,) = ctx.exchange.run("Get-OrganizationConfig")
    if config.get("AuditDisabled"):
        logger.warning("Mailbox auditing is off tenant-wide (AuditDisabled is True)")
    else:
        logger.info("Mailbox auditing is on")


if __name__ == "__main__":
    from m365_gdap_runner import main

    raise SystemExit(main(__name__))
