"""Single source of truth for this package's configuration options.

The :class:`Settings` dataclass below is THE one place where configuration
option names, types, defaults, secret classification, and help text are
defined. Everything else derives from it by iterating
``dataclasses.fields(Settings)``:

- ``resolve.py`` resolves each field through the precedence layers.
- ``file.py`` validates config.toml keys against the field names.
- The config CLI prompts and validates per field.
- Tests build fake configurations from the same field list.

To add, rename, or remove an option: edit :class:`Settings`; nothing else
needs changing.

Field conventions -- declare every field with :func:`option` or
:func:`secret`, never with a bare ``field(default=...)``:

- ``option(...)``: a non-secret. ``default_value`` is the value used when
  no layer provides one; omit it to make the option required.
- ``secret(...)``: a secret. Secret values never appear in source, so
  secret fields cannot have defaults and are always required.
  ``default_secret_name`` is the backend storage NAME the config CLI
  offers at init and writes to config.toml as ``<field>_secret_name``;
  runtime reads the name from config.toml only.
- Both take ``help``: the description used in prompts and comments.
- All fields are keyword-only, so declaration order is free.
- A schema with no secret fields never touches the secret machinery, which
  can then be removed from the package entirely.
"""

from __future__ import annotations

from dataclasses import MISSING, Field, dataclass, field
from typing import Any

# The application name, used for the config directory, env-var prefix, CLI
# name, and keyring service name. Kept as a literal (not derived from
# __package__) so a project-wide find-replace of the package name updates it.
APP_NAME = "m365_gdap_runner"

# Prefix for every environment-variable override: the app name upper-cased,
# so the field 'tenant_id' is overridden by <ENV_PREFIX>TENANT_ID.
ENV_PREFIX = APP_NAME.upper() + "_"

# Name of the config CLI's console script, used in error messages.
CLI_NAME = APP_NAME.replace("_", "-") + "-config"

# Credential-backend policy: "none", this package stores no secrets. The
# secret-storage machinery is removed, and Settings below declares no secret
# fields (sign-in is interactive; the token cache is handled by azure-auth).
CREDENTIAL_BACKEND = "none"


class ConfigError(Exception):
    """Raised for any configuration problem: missing, malformed, or mistyped values."""


def option(
    *,
    help: str,
    default_value: Any = MISSING,
    default_factory: Any = MISSING,
) -> Any:
    """Declare a non-secret settings field.

    Args:
        help: Description used in prompts and generated comments.
        default_value: Value used when no layer provides one. Omit to make
            the option required.
        default_factory: Zero-argument callable producing the default, for
            mutable defaults. Mutually exclusive with ``default_value``.

    Returns:
        A dataclass field for the settings schema.

    Raises:
        ValueError: If both defaults are given.
    """
    if default_value is not MISSING and default_factory is not MISSING:
        raise ValueError("Pass default_value or default_factory, not both.")
    kwargs: dict[str, Any] = {"kw_only": True, "metadata": {"help": help}}
    if default_value is not MISSING:
        kwargs["default"] = default_value
    if default_factory is not MISSING:
        kwargs["default_factory"] = default_factory
    return field(**kwargs)


def secret(
    *,
    help: str,
    default_secret_name: str | None = None,
) -> Any:
    """Declare a secret settings field.

    Secret values never appear in source: the field cannot take a default
    and is always required. Values come from the credential backend, or an
    environment variable at runtime.

    Args:
        help: Description used in prompts and generated comments.
        default_secret_name: Backend storage NAME (a name, not a value) the
            config CLI offers at init and writes to config.toml as
            ``<field>_secret_name``. Omit to offer the field name.

    Returns:
        A dataclass field for the settings schema (``repr=False`` so the
        value cannot leak through ``repr()``/tracebacks).

    Raises:
        ValueError: If ``default_secret_name`` is empty.
    """
    metadata: dict[str, Any] = {"secret": True, "help": help}
    if default_secret_name is not None:
        if not isinstance(default_secret_name, str) or not default_secret_name:
            raise ValueError(
                f"default_secret_name must be a non-empty string, got {default_secret_name!r}."
            )
        metadata["default_secret_name"] = default_secret_name
    return field(repr=False, kw_only=True, metadata=metadata)


@dataclass(frozen=True)
class Settings:
    """Resolved configuration for one profile.

    Instances are produced by :func:`m365_gdap_runner.config.load_settings`
    after applying the full precedence chain (explicit overrides > env vars >
    secret backend > config.toml > field defaults). Application code should
    treat instances as immutable value objects.
    """

    partner_tenant_id: str = option(
        help="Tenant ID or verified domain name of the MSP partner tenant"
    )
    username: str = option(help="User principal name of the partner account that signs in")
    client_id: str = option(
        default_value="",
        help=(
            "Application (client) ID to sign in with for every client; empty uses "
            "Microsoft's first-party apps (Graph, Exchange Online and Azure PowerShell)"
        ),
    )
    cache: str = option(
        default_value="memory",
        help=(
            "Token cache: 'memory' lives for one run; 'disk' is an encrypted file that lets "
            "later runs skip the browser sign-in"
        ),
    )
    cloud: str = option(
        default_value="",
        help=(
            "Microsoft cloud of the partner tenant (Commercial, USGov, USGovDoD or China); "
            "empty looks it up"
        ),
    )
    tenants_file: str = option(
        default_value="",
        help=(
            "CSV of customer tenants with one TRUE/FALSE column per group, used by --group; "
            "empty uses tenants.csv in the config directory"
        ),
    )
    debug_log_path: str = option(
        default_value="",
        help=(
            "File every log record is appended to, debug and library ones included, for "
            "development; empty writes none"
        ),
    )


def is_secret(f: Field[Any]) -> bool:
    """Return True if *f* is marked as a secret in its metadata.

    Args:
        f: A dataclass field of the settings schema.

    Returns:
        True when the field's value must come from a secret backend and never
        from config.toml.
    """
    return bool(f.metadata.get("secret", False))


def field_help(f: Field[Any]) -> str:
    """Return the human-readable description of *f* (empty string if unset).

    Args:
        f: A dataclass field of the settings schema.

    Returns:
        The ``help`` metadata string, or ``""`` when none was provided.
    """
    return str(f.metadata.get("help", ""))


def is_required(f: Field[Any]) -> bool:
    """Return True if *f* has no default and must be provided by some layer.

    Args:
        f: A dataclass field of the settings schema.

    Returns:
        True when the field has neither a default value nor a default factory.
    """
    return f.default is MISSING and f.default_factory is MISSING


def default_secret_name(f: Field[Any]) -> str | None:
    """Return *f*'s schema-declared backend storage name, if any.

    Args:
        f: A dataclass field of the settings schema.

    Returns:
        The ``default_secret_name`` metadata string, or None (the config
        CLI then offers the field name).
    """
    name = f.metadata.get("default_secret_name")
    return str(name) if name is not None else None


def field_default(f: Field[Any]) -> Any:
    """Return the field's default, materializing a default factory if needed.

    Args:
        f: A dataclass field of the settings schema (must not be required).

    Returns:
        The default value.

    Raises:
        ValueError: If the field is required and therefore has no default.
    """
    if f.default is not MISSING:
        return f.default
    if f.default_factory is not MISSING:
        return f.default_factory()
    raise ValueError(f"Field {f.name!r} is required and has no default.")
