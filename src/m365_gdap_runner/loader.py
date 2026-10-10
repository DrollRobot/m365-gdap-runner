"""Load a payload from a ``.py`` file or an installed entry point and validate it.

A payload module defines ``run(ctx)`` and declares what it uses:

- ``CLIENTS``: which of ``ctx.graph``, ``ctx.exchange``, ``ctx.ipps`` and ``ctx.azure``
  it needs. Defaults to ``["graph"]``. The runner signs in once per client at the
  partner tenant and checks each one in every customer tenant before ``run()``.
- ``SCOPES``: the delegated Graph scopes it needs in each customer tenant. Required
  when ``graph`` is in ``CLIENTS``, not allowed otherwise. The runner requests exactly
  these scopes per tenant, so a scope missing consent fails the tenant before
  ``run()`` starts; ``.default`` would bypass that check and is rejected.
"""

from __future__ import annotations

import importlib.metadata
import importlib.util
import inspect
import sys
import uuid
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from types import ModuleType

from m365_gdap_runner.context import CLIENT_KINDS, TenantContext

# Entry point group installed payload packages register under.
ENTRY_POINT_GROUP = "m365_gdap_runner.payloads"

# Scopes MSAL adds itself and rejects when passed in.
RESERVED_SCOPES = frozenset({"openid", "profile", "offline_access"})

# What a payload gets when it declares no ``CLIENTS``.
DEFAULT_CLIENTS: tuple[str, ...] = ("graph",)


class PayloadNotFound(LookupError):
    """Raised when a payload spec is neither an existing .py file nor an installed entry point."""


class PayloadInvalid(ValueError):
    """Raised when a payload fails to import or breaks the payload contract."""


@dataclass(frozen=True)
class Payload:
    """A validated payload.

    Attributes:
        name: File stem or entry point name.
        run: The payload's ``run(ctx)`` function.
        scopes: Delegated Graph scopes the payload needs in each tenant.
        clients: Client kinds the payload uses (``graph``, ``exchange``, ``ipps``, ``azure``).
    """

    name: str
    run: Callable[[TenantContext], object]
    scopes: tuple[str, ...]
    clients: tuple[str, ...] = DEFAULT_CLIENTS


def load_payload(spec: str) -> Payload:
    """Load and validate a payload.

    Args:
        spec: Path to a ``.py`` file, or the name of an entry point in
            :data:`ENTRY_POINT_GROUP`.

    Returns:
        The validated payload.

    Raises:
        PayloadNotFound: If the file does not exist or no entry point has that name.
        PayloadInvalid: If importing fails or the module breaks the contract.
    """
    path = Path(spec)
    if path.suffix.lower() == ".py":
        if not path.is_file():
            raise PayloadNotFound(f"Payload file not found: {spec}")
        return payload_from_module(_import_file(path), path.stem)
    return payload_from_module(_import_entry_point(spec), spec)


def _import_file(path: Path) -> ModuleType:
    """Import a payload file under a unique module name.

    Args:
        path: Existing ``.py`` file.

    Returns:
        The imported module (registered in ``sys.modules``).

    Raises:
        PayloadInvalid: If the file cannot be imported.
    """
    module_name = f"_m365_gdap_runner_payload_{uuid.uuid4().hex}"
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise PayloadInvalid(f"Cannot import payload file {path}.")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as exc:
        del sys.modules[module_name]
        raise PayloadInvalid(f"Importing payload {path} failed: {exc!r}") from exc
    return module


def _import_entry_point(name: str) -> ModuleType:
    """Import the module an installed payload entry point references.

    Args:
        name: Entry point name in :data:`ENTRY_POINT_GROUP`.

    Returns:
        The imported module.

    Raises:
        PayloadNotFound: If no entry point has that name.
        PayloadInvalid: If loading fails or the entry point is not a module.
    """
    matches = importlib.metadata.entry_points(group=ENTRY_POINT_GROUP, name=name)
    entry_point = next(iter(matches), None)
    if entry_point is None:
        raise PayloadNotFound(
            f"{name!r} is neither a .py file nor a payload installed under the "
            f"{ENTRY_POINT_GROUP!r} entry point group."
        )
    try:
        loaded = entry_point.load()
    except Exception as exc:
        raise PayloadInvalid(f"Loading payload entry point {name!r} failed: {exc!r}") from exc
    if not isinstance(loaded, ModuleType):
        raise PayloadInvalid(
            f"Payload entry point {name!r} must reference a module, not {type(loaded).__name__}."
        )
    return loaded


def payload_from_module(module: ModuleType, name: str) -> Payload:
    """Validate a payload module against the contract.

    Args:
        module: The imported payload module.
        name: Name used in error messages and results.

    Returns:
        The validated payload.

    Raises:
        PayloadInvalid: If ``run`` is missing or cannot take one positional
            argument, or ``CLIENTS`` or ``SCOPES`` is invalid.
    """
    run = getattr(module, "run", None)
    if not callable(run):
        raise PayloadInvalid(f"Payload {name!r} must define a callable run(ctx).")
    try:
        inspect.signature(run).bind(None)
    except TypeError as exc:
        raise PayloadInvalid(
            f"Payload {name!r}: run() must accept exactly one positional argument (ctx)."
        ) from exc
    clients = validate_clients(getattr(module, "CLIENTS", None), name)
    scopes = validate_scopes(getattr(module, "SCOPES", None), name, graph="graph" in clients)
    return Payload(name=name, run=run, scopes=scopes, clients=clients)


def validate_clients(value: object, name: str) -> tuple[str, ...]:
    """Validate a payload's ``CLIENTS``.

    Args:
        value: The ``CLIENTS`` value (None when undefined, which means ``graph``).
        name: Payload name for error messages.

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
            f"Payload {name!r}: CLIENTS must be a non-empty list drawn from {known}."
        )
    clients: list[str] = []
    for kind in value:
        if not isinstance(kind, str) or kind.strip().lower() not in CLIENT_KINDS:
            raise PayloadInvalid(f"Payload {name!r}: unknown client {kind!r}; expected {known}.")
        if kind.strip().lower() not in clients:
            clients.append(kind.strip().lower())
    return tuple(clients)


def validate_scopes(value: object, name: str, *, graph: bool = True) -> tuple[str, ...]:
    """Validate a payload's ``SCOPES``.

    Args:
        value: The ``SCOPES`` value (None when undefined).
        name: Payload name for error messages.
        graph: Whether the payload uses Graph. Scopes are required then, and not
            allowed otherwise, because they only mean something to Graph.

    Returns:
        The scopes, stripped, as a tuple; empty when *graph* is false.

    Raises:
        PayloadInvalid: If ``SCOPES`` is missing or empty, holds a non-string,
            uses ``.default``, lists a scope MSAL reserves, or is given by a
            payload that does not use Graph.
    """
    if not graph:
        if value:
            raise PayloadInvalid(
                f"Payload {name!r}: SCOPES are Graph scopes, but 'graph' is not in CLIENTS. "
                "Add it, or remove SCOPES."
            )
        return ()
    if not isinstance(value, list | tuple) or not value:
        raise PayloadInvalid(
            f"Payload {name!r} must define SCOPES, a non-empty list of delegated Graph scopes."
        )
    scopes: list[str] = []
    for scope in value:
        if not isinstance(scope, str) or not scope.strip():
            raise PayloadInvalid(
                f"Payload {name!r}: every SCOPES entry must be a non-empty string, got {scope!r}."
            )
        scope = scope.strip()
        if scope == ".default" or scope.endswith("/.default"):
            raise PayloadInvalid(
                f"Payload {name!r}: SCOPES must not use {scope!r}; list specific scopes so "
                "consent is checked in each tenant."
            )
        if scope.lower() in RESERVED_SCOPES:
            raise PayloadInvalid(
                f"Payload {name!r}: {scope!r} is added automatically; remove it from SCOPES."
            )
        scopes.append(scope)
    return tuple(scopes)
