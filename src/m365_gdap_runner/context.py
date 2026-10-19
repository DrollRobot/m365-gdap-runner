"""The per-tenant context a payload's ``run()`` receives.

A :class:`TenantContext` is a thin bundle: which tenant this call is for, the azure-auth
sign-in already pointed at that tenant, and one ready-made client per kind the payload
declared in ``CLIENTS``. Clients the payload did not declare were never signed in to or
checked, so asking for one raises :class:`ClientNotDeclared` instead of failing later with a
confusing token error.

Each context is used by one thread for one tenant; nothing in it is shared.
"""

from __future__ import annotations

from collections.abc import Sequence
from types import TracebackType
from typing import Literal, Self

import httpx
from azure_auth import AuthContext
from azure_auth.sync import ArmClient, ExchangeClient, GraphClient, IppsClient, ResourceClient

from m365_gdap_runner.tenants import Tenant

ClientKind = Literal["graph", "exchange", "ipps", "arm"]

# Every client a payload can declare in ``CLIENTS``, in the order they are signed in.
CLIENT_KINDS: tuple[ClientKind, ...] = ("graph", "exchange", "ipps", "arm")


class ClientNotDeclared(RuntimeError):
    """Raised when a payload uses a client it did not list in ``CLIENTS``."""

    def __init__(self, kind: str) -> None:
        """Name the missing client.

        Args:
            kind: The client kind that was asked for.
        """
        super().__init__(
            f"ctx.{kind} is not available: {kind!r} is not in this script's CLIENTS, so it "
            "was never signed in to or checked. Add it to CLIENTS."
        )
        self.kind = kind


def build_client(
    kind: str,
    auth: AuthContext,
    *,
    scopes: Sequence[str] = (),
    transport: httpx.BaseTransport | None = None,
) -> ResourceClient:
    """Create the azure-auth client for one kind.

    Args:
        kind: One of :data:`CLIENT_KINDS`.
        auth: The context the client gets its tokens from.
        scopes: Delegated Graph scopes; used by ``graph`` only.
        transport: Optional httpx transport (tests).

    Returns:
        The client. No network call is made.

    Raises:
        ValueError: If *kind* is not a known client kind.
    """
    if kind == "graph":
        return GraphClient(auth, scopes=list(scopes), transport=transport)
    if kind == "exchange":
        return ExchangeClient(auth, transport=transport)
    if kind == "ipps":
        return IppsClient(auth, transport=transport)
    if kind == "arm":
        return ArmClient(auth, transport=transport)
    raise ValueError(f"Unknown client kind {kind!r}; expected one of {', '.join(CLIENT_KINDS)}.")


class TenantContext:
    """Everything a payload's ``run()`` receives for one tenant.

    Attributes:
        tenant: The tenant being processed.
        auth: The azure-auth context for this tenant. Pass it to any azure-auth client, or
            to an Azure SDK client as ``credential=``, for resources not covered below.
    """

    def __init__(
        self,
        tenant: Tenant,
        auth: AuthContext,
        *,
        scopes: Sequence[str] = (),
        clients: Sequence[str] = ("graph",),
        transport: httpx.BaseTransport | None = None,
    ) -> None:
        """Build one client per declared kind. No network call is made.

        Args:
            tenant: The tenant being processed.
            auth: The azure-auth context for this tenant.
            scopes: The payload's Graph scopes.
            clients: The payload's ``CLIENTS``.
            transport: Optional httpx transport for every client (tests).
        """
        self.tenant = tenant
        self.auth = auth
        self._clients: dict[str, ResourceClient] = {
            kind: build_client(kind, auth, scopes=scopes, transport=transport) for kind in clients
        }

    @property
    def clients(self) -> tuple[str, ...]:
        """The client kinds this context holds, as declared by the payload."""
        return tuple(self._clients)

    def _client(self, kind: str) -> ResourceClient:
        """Return a declared client.

        Args:
            kind: The client kind.

        Returns:
            The client.

        Raises:
            ClientNotDeclared: If the payload did not declare *kind*.
        """
        try:
            return self._clients[kind]
        except KeyError:
            raise ClientNotDeclared(kind) from None

    @property
    def graph(self) -> GraphClient:
        """Microsoft Graph client bound to this tenant and the payload's Graph scopes."""
        client = self._client("graph")
        assert isinstance(client, GraphClient)  # noqa: S101 (narrowing; build_client guarantees it)
        return client

    @property
    def exchange(self) -> ExchangeClient:
        """Exchange Online cmdlet client bound to this tenant."""
        client = self._client("exchange")
        assert isinstance(client, ExchangeClient)  # noqa: S101
        return client

    @property
    def ipps(self) -> IppsClient:
        """Security & Compliance cmdlet client bound to this tenant."""
        client = self._client("ipps")
        assert isinstance(client, IppsClient)  # noqa: S101
        return client

    @property
    def arm(self) -> ArmClient:
        """Azure Resource Manager client bound to this tenant."""
        client = self._client("arm")
        assert isinstance(client, ArmClient)  # noqa: S101
        return client

    def probe(self) -> None:
        """Get a token for every declared client without prompting.

        This is the check :meth:`~m365_gdap_runner.session.PartnerSession.check_access` makes
        before the run: a context from azure-auth's ``get_tenant_context()`` never opens a
        browser, so a scope that is not consented in this tenant, or a Conditional Access policy
        that wants a fresh sign-in, fails here instead of inside the script.

        Raises:
            azure_auth.ConsentRequired: If consent is missing in this tenant.
            azure_auth.InteractionRequired: If the tenant wants a sign-in.
            azure_auth.AuthError: For any other token failure.
        """
        for client in self._clients.values():
            client.login()

    def close(self) -> None:
        """Close every client's HTTP connections."""
        for client in self._clients.values():
            client.close()

    def __enter__(self) -> Self:
        """Enter the context manager.

        Returns:
            This context.
        """
        return self

    def __exit__(
        self,
        _exc_type: type[BaseException] | None,
        _exc_value: BaseException | None,
        _traceback: TracebackType | None,
    ) -> None:
        """Close the clients when leaving the context manager."""
        self.close()
