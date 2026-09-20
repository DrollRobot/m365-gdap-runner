"""Run a payload file directly against one tenant.

Payload authors call :func:`standalone` in their ``__main__`` block::

    if __name__ == "__main__":
        import json
        from m365_gdap_runner import standalone
        print(json.dumps(run(standalone(SCOPES)), indent=2))

Then ``python my_payload.py --tenant Contoso``. Recognized flags: ``--tenant``
(prompted for when absent), ``--profile``, ``--partner-tenant-id``,
``--client-id``, ``--persist-token-cache``. Other arguments are left for the
payload.
"""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence

import httpx

from m365_gdap_runner.auth import (
    GRAPH_RESOURCE,
    MsalAppFactory,
    PartnerSession,
    TenantAuthError,
    default_app_factory,
    group_by_resource,
)
from m365_gdap_runner.graph import GraphClient
from m365_gdap_runner.loader import validate_scopes
from m365_gdap_runner.main import (
    EXIT_TENANT_FAILED,
    EXIT_USAGE,
    SETUP_ERRORS,
    add_settings_arguments,
    settings_from_args,
)
from m365_gdap_runner.runner import probe
from m365_gdap_runner.tenants import TenantContext, list_gdap_tenants, select


def standalone(
    scopes: Sequence[str],
    argv: Sequence[str] | None = None,
    *,
    app_factory: MsalAppFactory = default_app_factory,
    transport: httpx.BaseTransport | None = None,
) -> TenantContext:
    """Sign in, resolve one tenant, check consent, and return its context.

    Args:
        scopes: The payload's ``SCOPES``.
        argv: Arguments to parse; None reads ``sys.argv``.
        app_factory: Builds MSAL applications; tests inject a fake.
        transport: Optional httpx transport for Graph clients (tests).

    Returns:
        The context to pass to the payload's ``run()``.

    Raises:
        SystemExit: With code 2 on a configuration, sign-in, enumeration, or
            tenant-selection error, or 1 when the tenant refuses a token (for
            example, missing consent). The message goes to stderr.
    """
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--tenant", default=None)
    add_settings_arguments(parser)
    args, _payload_args = parser.parse_known_args(argv)
    try:
        payload_scopes = validate_scopes(scopes, "standalone")
        session = PartnerSession(settings_from_args(args), app_factory=app_factory)
        session.login()
        tenants = list_gdap_tenants(session, transport=transport)
        spec = args.tenant or input("Tenant name or ID: ").strip()
        tenant = select(tenants, include=[spec])[0]
    except SETUP_ERRORS as exc:
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(EXIT_USAGE) from exc

    credential = session.credential_for(tenant.tenant_id)
    try:
        probe(credential, payload_scopes)
    except TenantAuthError as exc:
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(EXIT_TENANT_FAILED) from exc
    graph_scopes = group_by_resource(payload_scopes).get(GRAPH_RESOURCE, [])
    graph = GraphClient(credential, graph_scopes, transport=transport)
    return TenantContext(tenant=tenant, graph=graph, credential=credential)
