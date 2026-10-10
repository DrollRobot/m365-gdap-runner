"""Run a payload file directly against one tenant.

Payload authors call :func:`standalone` in their ``__main__`` block::

    if __name__ == "__main__":
        import json
        from m365_gdap_runner import standalone
        print(json.dumps(run(standalone(SCOPES, CLIENTS)), indent=2))

Then ``python my_payload.py --tenant Contoso``. Recognized flags: ``--tenant``
(prompted for when absent) and the settings flags ``--profile``,
``--partner-tenant-id``, ``--username``, ``--client-id``, ``--cache`` and
``--cloud``. Other arguments are left for the payload.
"""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence

import httpx
from azure_auth import AuthError

from m365_gdap_runner.context import TenantContext
from m365_gdap_runner.loader import DEFAULT_CLIENTS, validate_clients, validate_scopes
from m365_gdap_runner.main import (
    EXIT_TENANT_FAILED,
    EXIT_USAGE,
    SETUP_ERRORS,
    add_settings_arguments,
    settings_from_args,
)
from m365_gdap_runner.session import PartnerSession
from m365_gdap_runner.tenants import select


def standalone(
    scopes: Sequence[str] = (),
    clients: Sequence[str] = DEFAULT_CLIENTS,
    argv: Sequence[str] | None = None,
    *,
    transport: httpx.BaseTransport | None = None,
) -> TenantContext:
    """Sign in, resolve one tenant, check every client there, and return its context.

    Args:
        scopes: The payload's ``SCOPES`` (Graph scopes; leave empty without Graph).
        clients: The payload's ``CLIENTS``.
        argv: Arguments to parse; None reads ``sys.argv``.
        transport: Optional httpx transport for every client (tests).

    Returns:
        The context to pass to the payload's ``run()``. Close it, or use it as a
        context manager, when done.

    Raises:
        SystemExit: With code 2 on a configuration, payload, sign-in, enumeration,
            or tenant-selection error, or 1 when the tenant refuses a token (for
            example, missing consent). The message goes to stderr.
    """
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--tenant", default=None)
    add_settings_arguments(parser)
    args, _payload_args = parser.parse_known_args(argv)
    try:
        payload_clients = validate_clients(list(clients), "standalone")
        payload_scopes = validate_scopes(
            list(scopes) or None, "standalone", graph="graph" in payload_clients
        )
        session = PartnerSession(settings_from_args(args), transport=transport)
        session.login(payload_clients)
        tenants = session.list_tenants()
        spec = args.tenant or input("Tenant name or ID: ").strip()
        tenant = select(tenants, include=[spec])[0]
    except SETUP_ERRORS as exc:
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(EXIT_USAGE) from exc

    context = session.context_for(tenant, scopes=payload_scopes, clients=payload_clients)
    try:
        context.probe()
    except AuthError as exc:
        context.close()
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(EXIT_TENANT_FAILED) from exc
    return context
