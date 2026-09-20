"""Run one Python payload across GDAP customer tenants from a single partner sign-in.

A payload is a module with ``SCOPES`` and ``run(ctx)``; it needs only
:class:`TenantContext` and, to run the file directly, :func:`standalone`::

    from m365_gdap_runner import TenantContext

    SCOPES = ["User.Read.All"]

    def run(ctx: TenantContext) -> list[dict[str, object]]:
        return ctx.graph.get_all("/users", params={"$select": "userPrincipalName"})

    if __name__ == "__main__":
        import json
        from m365_gdap_runner import standalone
        print(json.dumps(run(standalone(SCOPES)), indent=2))

Run it across tenants with ``m365-gdap-runner my_payload.py``.
"""

from __future__ import annotations

from m365_gdap_runner.auth import PartnerSession, TenantAuthError, TenantCredential
from m365_gdap_runner.graph import GraphClient, GraphError
from m365_gdap_runner.standalone import standalone
from m365_gdap_runner.tenants import Tenant, TenantContext

# __all__ controls what `from m365_gdap_runner import *` exports and what static
# analysis tools treat as the public API.
__all__ = [
    "GraphClient",
    "GraphError",
    "PartnerSession",
    "Tenant",
    "TenantAuthError",
    "TenantContext",
    "TenantCredential",
    "standalone",
]
