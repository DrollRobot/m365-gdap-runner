"""Run one Python payload across GDAP customer tenants from a single partner sign-in.

Authentication is azure-auth's job. This package signs in once per client the payload
uses, lists the customer tenants, and calls the payload's ``run(ctx)`` for each one
with error isolation, optional parallelism, and merged JSON or CSV output.

A payload is a module with ``run(ctx)``, the Graph ``SCOPES`` it needs, and which
``CLIENTS`` it uses (``graph`` by default; also ``exchange``, ``ipps``, ``azure``)::

    from m365_gdap_runner import TenantContext

    SCOPES = ["User.Read.All"]
    CLIENTS = ["graph", "exchange"]

    def run(ctx: TenantContext) -> dict[str, object]:
        users = ctx.graph.get_all("/users", params={"$select": "userPrincipalName"})
        (config,) = ctx.exchange.run("Get-OrganizationConfig")
        return {"users": len(users), "audit_disabled": config["AuditDisabled"]}

    if __name__ == "__main__":
        import json
        from m365_gdap_runner import standalone
        print(json.dumps(run(standalone(SCOPES, CLIENTS)), indent=2))

Run it with ``m365-gdap-runner my_payload.py --group pilot``, where ``pilot`` is a TRUE/FALSE
column of the tenants file (see ``tenants.py``); ``--tenant`` and ``--all`` are the other
ways to choose tenants.
"""

from __future__ import annotations

from m365_gdap_runner.context import CLIENT_KINDS, ClientNotDeclared, TenantContext
from m365_gdap_runner.session import PartnerSession
from m365_gdap_runner.standalone import standalone
from m365_gdap_runner.tenants import Tenant

# __all__ controls what `from m365_gdap_runner import *` exports and what static
# analysis tools treat as the public API.
__all__ = [
    "CLIENT_KINDS",
    "ClientNotDeclared",
    "PartnerSession",
    "Tenant",
    "TenantContext",
    "standalone",
]
