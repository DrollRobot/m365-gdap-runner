"""Run a Python script across GDAP customer tenants from a single partner sign-in.

Authentication is azure-auth's job. This library signs in once per client the script uses,
lists the customer tenants, and calls the script's ``run(ctx)`` for each one with error
isolation, optional parallelism, and merged JSON or CSV output.

A script is a ``.py`` file with ``run(ctx)``, the Graph ``SCOPES`` it needs, and which
``CLIENTS`` it uses (``graph`` by default; also ``exchange``, ``ipps``, ``azure``). Its
``__main__`` block hands it to :func:`main`::

    from m365_gdap_runner import TenantContext

    SCOPES = ["User.Read.All"]
    CLIENTS = ["graph", "exchange"]

    def run(ctx: TenantContext) -> dict[str, object]:
        users = ctx.graph.get_all("/users", params={"$select": "userPrincipalName"})
        (config,) = ctx.exchange.run("Get-OrganizationConfig")
        return {"users": len(users), "audit_disabled": config["AuditDisabled"]}

    if __name__ == "__main__":
        from m365_gdap_runner import main

        raise SystemExit(main(__name__))

Run it with ``uv run my_script.py --group pilot``, where ``pilot`` is a TRUE/FALSE column of
the tenants file (see ``tenants.py``); ``--tenant`` and ``--all`` are the other ways to choose
tenants. A script with flags of its own adds the library's to its parser with
:func:`add_arguments` and passes the parsed result to :func:`main`.
"""

from __future__ import annotations

from m365_gdap_runner.cli import add_arguments, main
from m365_gdap_runner.context import CLIENT_KINDS, ClientNotDeclared, TenantContext
from m365_gdap_runner.session import PartnerSession
from m365_gdap_runner.tenants import Tenant

# __all__ controls what `from m365_gdap_runner import *` exports and what static
# analysis tools treat as the public API.
__all__ = [
    "CLIENT_KINDS",
    "ClientNotDeclared",
    "PartnerSession",
    "Tenant",
    "TenantContext",
    "add_arguments",
    "main",
]
