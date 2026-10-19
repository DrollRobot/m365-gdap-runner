"""Check a script module against the contract :func:`~m365_gdap_runner.main` runs it by.

A script defines ``run(ctx)`` and declares what it uses:

- ``CLIENTS``: which of ``ctx.graph``, ``ctx.exchange``, ``ctx.ipps`` and ``ctx.arm``
  it needs. Defaults to ``["graph"]``. The runner signs in once per client at the
  partner tenant and checks each one in every customer tenant before ``run()``.
- ``SCOPES``: the delegated Graph scopes it needs in each customer tenant. Required
  when ``graph`` is in ``CLIENTS``, not allowed otherwise. The runner requests exactly
  these scopes per tenant, so a scope missing consent fails the tenant before
  ``run()`` starts; ``.default`` would bypass that check and is rejected.

It may also define ``report(results)``, called once after the last tenant with every
tenant's :class:`~m365_gdap_runner.runner.TenantResult`, failed ones included. The library
writes nothing itself: ``run()`` writes any per-tenant report, ``report()`` the run's.

And it may set ``CONTINUE_ON_ERROR``. An error in ``run()`` stops the run by default, so a
script that changes things does not repeat a broken change in every tenant; a script that
only reads sets it True to keep going. ``--continue-on-error`` and
``--no-continue-on-error`` override it.
"""

from __future__ import annotations

import inspect
from collections.abc import Callable
from dataclasses import dataclass
from types import ModuleType
from typing import TYPE_CHECKING

from m365_gdap_runner.context import CLIENT_KINDS, TenantContext

if TYPE_CHECKING:
    from m365_gdap_runner.runner import TenantResult

# Scopes MSAL adds itself and rejects when passed in.
RESERVED_SCOPES = frozenset({"openid", "profile", "offline_access"})

# What a script gets when it declares no ``CLIENTS``.
DEFAULT_CLIENTS: tuple[str, ...] = ("graph",)


class PayloadInvalid(ValueError):
    """Raised when a script breaks the contract."""


@dataclass(frozen=True)
class Payload:
    """A validated script.

    Attributes:
        name: The script's file name without ``.py``.
        run: The script's ``run(ctx)`` function.
        scopes: Delegated Graph scopes the script needs in each tenant.
        clients: Client kinds the script uses (``graph``, ``exchange``, ``ipps``, ``arm``).
        report: The script's ``report(results)`` function, or None when it has none.
        continue_on_error: Whether the run goes on after an error in ``run()``.
    """

    name: str
    run: Callable[[TenantContext], object]
    scopes: tuple[str, ...]
    clients: tuple[str, ...] = DEFAULT_CLIENTS
    report: Callable[[list[TenantResult]], object] | None = None
    continue_on_error: bool = False


def _one_argument_function(
    function: object, name: str, attribute: str, argument: str
) -> Callable[..., object]:
    """Make sure a script function exists and can be called with one positional argument.

    Args:
        function: The module attribute.
        name: Script name for error messages.
        attribute: The function's name in the contract, such as ``run``.
        argument: What the contract calls its argument, such as ``ctx``.

    Returns:
        The function.

    Raises:
        PayloadInvalid: If *function* is not callable or cannot take one positional argument.
    """
    if not callable(function):
        raise PayloadInvalid(f"Script {name!r} must define a callable {attribute}({argument}).")
    try:
        inspect.signature(function).bind(None)
    except TypeError as exc:
        raise PayloadInvalid(
            f"Script {name!r}: {attribute}() must accept exactly one positional argument "
            f"({argument})."
        ) from exc
    return function


def payload_from_module(module: ModuleType, name: str) -> Payload:
    """Validate a script module against the contract.

    Args:
        module: The script's module.
        name: Name used in error messages and logs.

    Returns:
        The validated script.

    Raises:
        PayloadInvalid: If ``run`` is missing, ``run`` or ``report`` cannot take one
            positional argument, ``CLIENTS`` or ``SCOPES`` is invalid, or
            ``CONTINUE_ON_ERROR`` is not True or False.
    """
    run = _one_argument_function(getattr(module, "run", None), name, "run", "ctx")
    report = getattr(module, "report", None)
    if report is not None:
        report = _one_argument_function(report, name, "report", "results")
    clients = validate_clients(getattr(module, "CLIENTS", None), name)
    scopes = validate_scopes(getattr(module, "SCOPES", None), name, graph="graph" in clients)
    continue_on_error = getattr(module, "CONTINUE_ON_ERROR", False)
    if not isinstance(continue_on_error, bool):
        raise PayloadInvalid(
            f"Script {name!r}: CONTINUE_ON_ERROR must be True or False, not {continue_on_error!r}."
        )
    return Payload(
        name=name,
        run=run,
        scopes=scopes,
        clients=clients,
        report=report,
        continue_on_error=continue_on_error,
    )


def validate_clients(value: object, name: str) -> tuple[str, ...]:
    """Validate a script's ``CLIENTS``.

    Args:
        value: The ``CLIENTS`` value (None when undefined, which means ``graph``).
        name: Script name for error messages.

    Returns:
        The client kinds, deduplicated, in declaration order.

    Raises:
        PayloadInvalid: If ``CLIENTS`` is defined but empty, not a list, or names
            an unknown client.
    """
    if value is None:
        return DEFAULT_CLIENTS
    known = ", ".join(CLIENT_KINDS)
    if not isinstance(value, list | tuple) or not value:
        raise PayloadInvalid(
            f"Script {name!r}: CLIENTS must be a non-empty list drawn from {known}."
        )
    clients: list[str] = []
    for kind in value:
        if not isinstance(kind, str) or kind.strip().lower() not in CLIENT_KINDS:
            raise PayloadInvalid(f"Script {name!r}: unknown client {kind!r}; expected {known}.")
        if kind.strip().lower() not in clients:
            clients.append(kind.strip().lower())
    return tuple(clients)


def validate_scopes(value: object, name: str, *, graph: bool = True) -> tuple[str, ...]:
    """Validate a script's ``SCOPES``.

    Args:
        value: The ``SCOPES`` value (None when undefined).
        name: Script name for error messages.
        graph: Whether the script uses Graph. Scopes are required then, and not
            allowed otherwise, because they only mean something to Graph.

    Returns:
        The scopes, stripped, as a tuple; empty when *graph* is false.

    Raises:
        PayloadInvalid: If ``SCOPES`` is missing or empty, holds a non-string,
            uses ``.default``, lists a scope MSAL reserves, or is given by a
            script that does not use Graph.
    """
    if not graph:
        if value:
            raise PayloadInvalid(
                f"Script {name!r}: SCOPES are Graph scopes, but 'graph' is not in CLIENTS. "
                "Add it, or remove SCOPES."
            )
        return ()
    if not isinstance(value, list | tuple) or not value:
        raise PayloadInvalid(
            f"Script {name!r} must define SCOPES, a non-empty list of delegated Graph scopes."
        )
    scopes: list[str] = []
    for scope in value:
        if not isinstance(scope, str) or not scope.strip():
            raise PayloadInvalid(
                f"Script {name!r}: every SCOPES entry must be a non-empty string, got {scope!r}."
            )
        scope = scope.strip()
        if scope == ".default" or scope.endswith("/.default"):
            raise PayloadInvalid(
                f"Script {name!r}: SCOPES must not use {scope!r}; list specific scopes so "
                "consent is checked in each tenant."
            )
        if scope.lower() in RESERVED_SCOPES:
            raise PayloadInvalid(
                f"Script {name!r}: {scope!r} is added automatically; remove it from SCOPES."
            )
        scopes.append(scope)
    return tuple(scopes)
