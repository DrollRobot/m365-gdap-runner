"""Running a script across tenants, and the ``m365-gdap-runner-customer-export`` command.

A script runs itself: its ``__main__`` block hands the module to :func:`main`, which reads
the command line, signs in once per client the script uses, lists GDAP customer tenants, and
calls the script's ``run(ctx)`` for each selected tenant::

    if __name__ == "__main__":
        from m365_gdap_runner import main

        raise SystemExit(main(__name__))

then ``uv run audit.py --include pilot``. A script with flags of its own builds the parser,
adds the library's flags with :func:`add_arguments`, and passes what it parsed to ``main()``.

The library writes no output of its own. ``run()`` writes any per-tenant report; after the
last tenant, the script's ``report(results)``, when it has one, gets every tenant's result and
writes the run's.

Tenants are chosen with ``--include``, less ``--exclude``: each takes a tenant, a group (a
TRUE/FALSE column of the tenants file, see ``tenants.py``), and ``--include`` also ``all``,
every GDAP tenant. A run that includes nothing is refused. A tenant or group may pick tenants
without a GDAP relationship, or in another cloud, from the tenants file; before its
``run()``, every tenant the partner sign-in cannot reach gets a sign-in of its own. Each
tenant's ``run()`` starts as soon as its access is settled, without waiting for the other
tenants'. A tenant whose ``m365_gdap`` is not TRUE and whose ``m365_username`` is blank is
left out. A run never writes the tenants file: it warns of the columns, GDAP tenants and
clouds the file lacks. The ``m365-gdap-runner-customer-export`` command
(:func:`list_tenants`) adds them, creating the file if need be, and prints the list.

Configuration comes from the config system (see ``src/m365_gdap_runner/config/``);
``--partner-tenant-id``, ``--username``, ``--client-id``, ``--cache``, ``--cloud``,
``--tenants-file`` and ``--jobs`` override it. Progress and the summary go to stderr; stdout
is the script's. The ``debug_log_path`` setting also appends every log record to a file.

Exit codes: 0 all tenants ok; 1 any tenant not ok; 2 usage, configuration, script, sign-in,
enumeration, tenants-file or tenant-selection error, or ``report()`` failed; 130 interrupted.
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from collections.abc import Sequence
from pathlib import Path
from types import ModuleType

import httpx
from azure_auth import AuthError, ResourceError

from m365_gdap_runner.config import ConfigError, Settings, load_settings
from m365_gdap_runner.loader import Payload, PayloadInvalid, payload_from_module
from m365_gdap_runner.runner import (
    RunInterrupted,
    TenantLogFilter,
    TenantResult,
    run_payload,
    summarize,
)
from m365_gdap_runner.session import CACHE_KINDS, PartnerSession
from m365_gdap_runner.tenants import (
    ALL,
    Tenant,
    TenantFile,
    TenantFileError,
    TenantSelectionError,
    select,
    tenants_file_path,
)

logger = logging.getLogger(__name__)

EXIT_OK = 0
EXIT_TENANT_FAILED = 1
EXIT_USAGE = 2
EXIT_INTERRUPTED = 130

# Errors that end a run before any tenant is processed.
SETUP_ERRORS = (
    ConfigError,
    PayloadInvalid,
    AuthError,
    TenantFileError,
    TenantSelectionError,
    ResourceError,
    httpx.HTTPError,
)

# Library loggers kept at WARNING on the terminal unless -vv is given.
_LIBRARY_LOGGERS = (
    "httpx",
    "httpcore",
    "msal",
    "msal_extensions",
    "urllib3",
    "azure",
    "azure_auth",
)

# Debug log lines: the full date, and the logger each record came from.
_DEBUG_LOG_FORMAT = "%(asctime)s %(levelname)s [%(tenant)s] %(name)s: %(message)s"

# The debug log's file name when ``debug_log_path`` names a folder.
_DEBUG_LOG_NAME = "m365_gdap_runner_debug.log"

# Settings a flag can override, with the flag's metavar.
_SETTING_FLAGS = (
    ("partner_tenant_id", "ID_OR_DOMAIN", "Partner tenant ID or domain name."),
    ("username", "UPN", "Partner account to sign in with."),
    (
        "client_id",
        "ID",
        "Application ID to sign in with for every client (default: Microsoft's first-party apps).",
    ),
    (
        "cloud",
        "NAME",
        "Cloud of the partner tenant: Commercial, USGov, USGovDoD or China (default: look it up).",
    ),
    (
        "tenants_file",
        "PATH",
        "Tenants CSV (default: the tenants_file setting, else tenants.csv in the config "
        "directory).",
    ),
)


def _positive_int(value: str) -> int:
    """Parse a whole number of at least 1, for ``--jobs``.

    Args:
        value: The flag's text.

    Returns:
        The number.

    Raises:
        argparse.ArgumentTypeError: If the text is not a whole number of at least 1.
    """
    try:
        number = int(value)
    except ValueError:
        raise argparse.ArgumentTypeError(f"{value!r} is not a whole number") from None
    if number < 1:
        raise argparse.ArgumentTypeError("must be at least 1")
    return number


def _add_settings_arguments(parser: argparse.ArgumentParser) -> None:
    """Add the flags that override settings, and ``-v``.

    Args:
        parser: Parser to extend.
    """
    group = parser.add_argument_group("m365-gdap-runner settings")
    group.add_argument(
        "--profile",
        default=None,
        help="Config profile to use (default: default_profile from config.toml).",
    )
    for name, metavar, help_text in _SETTING_FLAGS:
        group.add_argument(
            "--" + name.replace("_", "-"), default=None, metavar=metavar, help=help_text
        )
    group.add_argument(
        "--cache",
        choices=CACHE_KINDS,
        default=None,
        help="Token cache: memory (default) or disk, an encrypted file that lets later runs "
        "skip the sign-in.",
    )
    group.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="-v: debug logs from this tool and the script; -vv: also from libraries.",
    )


def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Add every flag :func:`main` reads to a script's own parser.

    Call it once, after adding the script's own flags; ``-h`` then lists both. Flag names
    the script uses for itself must not clash with these.

    Args:
        parser: The script's parser.
    """
    tenants = parser.add_argument_group("m365-gdap-runner tenants")
    tenants.add_argument(
        "--include",
        action="append",
        default=None,
        metavar="TENANT_OR_GROUP",
        help=(
            "Run this tenant (display name or ID), the tenants marked TRUE in this column of "
            "the tenants file, or all, every GDAP tenant (repeatable; they add up)."
        ),
    )
    tenants.add_argument(
        "--exclude",
        action="append",
        default=None,
        metavar="TENANT_OR_GROUP",
        help="Skip this tenant or group, even where --include names it (repeatable).",
    )
    run = parser.add_argument_group("m365-gdap-runner run")
    run.add_argument(
        "--continue-on-error",
        action=argparse.BooleanOptionalAction,
        default=None,
        help=(
            "Keep running other tenants after an error in run(); --no-continue-on-error "
            "stops at the first. Default: the script's CONTINUE_ON_ERROR, else stop."
        ),
    )
    run.add_argument(
        "--grant-when-possible",
        action=argparse.BooleanOptionalAction,
        default=None,
        help=(
            "Where a tenant's grant lacks some of the script's SCOPES and the app holds "
            "DelegatedPermissionGrant.ReadWrite.All there, add them to the app's grant. "
            "Default: the script's GRANT_WHEN_POSSIBLE, else off."
        ),
    )
    run.add_argument(
        "--jobs",
        type=_positive_int,
        default=None,
        metavar="N",
        help=(
            "Tenants to work on at once (default: the jobs setting, else 1). 5-8 suits "
            "read-only reports; keep 1 for scripts that write unless they are idempotent."
        ),
    )
    _add_settings_arguments(parser)


def settings_from_args(args: argparse.Namespace) -> Settings:
    """Resolve settings with the flags from :func:`add_arguments` as overrides.

    Args:
        args: Parsed arguments.

    Returns:
        The resolved settings.

    Raises:
        ConfigError: If a required value is missing or invalid.
    """
    overrides: dict[str, object] = {}
    for name, _metavar, _help in _SETTING_FLAGS:
        value = getattr(args, name)
        if value is not None:
            overrides[name] = value
    if args.cache is not None:
        overrides["cache"] = args.cache
    # add_arguments() adds --jobs; the m365-gdap-runner-customer-export command, which runs
    # nothing, does not.
    if getattr(args, "jobs", None) is not None:
        overrides["jobs"] = args.jobs
    return load_settings(profile=args.profile, overrides=overrides)


def _selection_error(args: argparse.Namespace) -> str | None:
    """Check that the flags choose tenants, before any sign-in.

    Args:
        args: Parsed arguments.

    Returns:
        What is wrong, or None when the selection is usable.
    """
    if not args.include:
        return "Choose tenants with --include: a tenant, a group, or all."
    if any(value.strip().casefold() == ALL for value in args.exclude or ()):
        return "--exclude all would leave no tenant to run."
    return None


def _quiet_libraries(record: logging.LogRecord) -> bool:
    """Drop a library's records below WARNING: the terminal's filter without ``-vv``.

    Args:
        record: The log record.

    Returns:
        False for a record below WARNING from a logger in ``_LIBRARY_LOGGERS`` or below one.
    """
    if record.levelno >= logging.WARNING:
        return True
    return not any(
        record.name == name or record.name.startswith(name + ".") for name in _LIBRARY_LOGGERS
    )


def configure_logging(verbosity: int) -> None:
    """Send logs to stderr, tagged with the tenant each record belongs to.

    The handler filters by level itself, as well as the loggers, so :func:`_start_debug_log`
    can lower the loggers' levels without changing what reaches the terminal.

    Args:
        verbosity: Count of ``-v`` flags.
    """
    handler = logging.StreamHandler(sys.stderr)
    handler.setLevel(logging.DEBUG if verbosity >= 1 else logging.INFO)
    if verbosity < 2:
        handler.addFilter(_quiet_libraries)
    handler.addFilter(TenantLogFilter())
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s [%(tenant)s] %(message)s", "%H:%M:%S")
    )
    root = logging.getLogger()
    root.addHandler(handler)
    root.setLevel(logging.DEBUG if verbosity >= 1 else logging.INFO)
    library_level = logging.DEBUG if verbosity >= 2 else logging.WARNING
    for name in _LIBRARY_LOGGERS:
        logging.getLogger(name).setLevel(library_level)


def _debug_log_file(setting: str) -> Path:
    """Return the file a ``debug_log_path`` setting names.

    A setting that names a folder, one that exists or one ending in a path separator, means
    :data:`_DEBUG_LOG_NAME` in that folder.

    Args:
        setting: The ``debug_log_path`` setting; not empty.

    Returns:
        The debug log's path, with ``~`` expanded.
    """
    path = Path(setting).expanduser()
    if path.is_dir() or setting.endswith(tuple(sep for sep in (os.sep, os.altsep) if sep)):
        return path / _DEBUG_LOG_NAME
    return path


def _start_debug_log(settings: Settings) -> None:
    """Append every log record, at every level, to the ``debug_log_path`` file, when set.

    The terminal shows what it did before: its handler from :func:`configure_logging`
    filters by level itself. Records logged before settings load reach only the terminal. A
    file that cannot be opened is reported, and the run goes on without it.

    Args:
        settings: Resolved settings.
    """
    if not settings.debug_log_path:
        return
    path = _debug_log_file(settings.debug_log_path)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        handler = logging.FileHandler(path, mode="a", encoding="utf-8")
    except OSError as exc:
        logger.warning("Could not open debug log %s: %s", path, exc)
        return
    handler.addFilter(TenantLogFilter())
    handler.setFormatter(logging.Formatter(_DEBUG_LOG_FORMAT))
    root = logging.getLogger()
    root.addHandler(handler)
    root.setLevel(logging.DEBUG)
    for name in _LIBRARY_LOGGERS:
        logging.getLogger(name).setLevel(logging.DEBUG)
    logger.debug("Debug log started: %s", " ".join(sys.argv))


def _print_tenants(tenants: Sequence[Tenant]) -> None:
    """Print one ``tenant_id<TAB>display_name`` line per tenant to stdout.

    Args:
        tenants: Tenants to print.
    """
    for tenant in tenants:
        print(f"{tenant.tenant_id}\t{tenant.display_name}")


def _load_tenant_file(
    settings: Settings, session: PartnerSession, tenants: Sequence[Tenant]
) -> TenantFile:
    """Load the tenants file for a run, which never writes it, and warn of what it lacks.

    A column this package owns that the file lacks, a GDAP tenant with no row, and a blank
    cloud, looked up by discovery, are filled in memory for this run (see
    :meth:`TenantFile.sync`), with a warning that the ``m365-gdap-runner-customer-export``
    command saves them. A missing file reads as one with no rows.

    Args:
        settings: Resolved settings, for the file's location.
        session: The partner session, for cloud lookups.
        tenants: The GDAP tenants.

    Returns:
        The file, with every GDAP tenant in it unless it is missing.

    Raises:
        TenantFileError: If the file exists but cannot be read.
    """
    path = tenants_file_path(settings)
    if not path.exists():
        logger.warning(
            "No tenants file at %s; run m365-gdap-runner-customer-export to create it", path
        )
        return TenantFile.empty(path)
    tenant_file = TenantFile.load(path)
    report = tenant_file.sync(tenants, session.cloud_of)
    if report.columns_added:
        logger.warning(
            "Tenants file %s lacks column(s) %s; run m365-gdap-runner-customer-export to add them",
            path,
            ", ".join(report.columns_added),
        )
    if report.added:
        names = ", ".join(tenant.display_name for tenant in report.added)
        logger.warning(
            "%d GDAP tenant(s) have no row in %s; run m365-gdap-runner-customer-export to add "
            "them: %s",
            len(report.added),
            path,
            names,
        )
    if report.clouds_filled:
        names = ", ".join(tenant.display_name for tenant in report.clouds_filled)
        logger.warning(
            "Looked up the blank m365_cloud of %d tenant(s) in %s; run "
            "m365-gdap-runner-customer-export to save them: %s",
            len(report.clouds_filled),
            path,
            names,
        )
    return tenant_file


def _update_tenant_file(
    settings: Settings, session: PartnerSession, tenants: Sequence[Tenant]
) -> None:
    """Bring the tenants file up to date with the GDAP tenants.

    For the ``m365-gdap-runner-customer-export`` command.

    A column this package owns that the file lacks is added. A GDAP tenant with no row is
    appended; one whose cloud is blank gets it looked up. A missing file is created. A file
    that cannot be written (a spreadsheet may hold it open) is reported and left as it was.

    Args:
        settings: Resolved settings, for the file's location.
        session: The partner session, for cloud lookups.
        tenants: The GDAP tenants.

    Raises:
        TenantFileError: If the file exists but cannot be read.
    """
    path = tenants_file_path(settings)
    existed = path.exists()
    tenant_file = TenantFile.load(path) if existed else TenantFile.empty(path)
    report = tenant_file.sync(tenants, session.cloud_of)
    if existed and not report.changed:
        logger.debug("Tenants file %s is up to date", path)
        return
    try:
        tenant_file.save()
    except OSError as exc:
        logger.warning("Could not write tenants file %s: %s", path, exc)
        return
    if not existed:
        logger.info(
            "Created tenants file %s with %d tenant(s); add a TRUE/FALSE column per group",
            path,
            len(report.added),
        )
        return
    if report.columns_added:
        logger.info("Added column(s) %s to %s", ", ".join(report.columns_added), path)
    if report.added:
        names = ", ".join(tenant.display_name for tenant in report.added)
        logger.warning(
            "Added %d tenant(s) to %s; they are in no group yet: %s",
            len(report.added),
            path,
            names,
        )
    if report.clouds_filled:
        names = ", ".join(tenant.display_name for tenant in report.clouds_filled)
        logger.info(
            "Filled in the cloud of %d tenant(s) in %s: %s", len(report.clouds_filled), path, names
        )


def _select_tenants(
    args: argparse.Namespace, tenants: Sequence[Tenant], tenant_file: TenantFile, cloud: str
) -> list[Tenant]:
    """Apply ``--include`` and ``--exclude`` (see :func:`~m365_gdap_runner.tenants.select`).

    ``all`` means every GDAP tenant. A tenant or group may also pick a row of the tenants file
    that GDAP does not list.

    Args:
        args: Parsed arguments that passed :func:`_selection_error`, so ``--include`` was
            given.
        tenants: The GDAP tenants.
        tenant_file: The tenants file.
        cloud: Name of the partner's cloud, for a tenant whose row names none.

    Returns:
        The tenants to run, GDAP tenants first, each with its cloud and sign-in account. A
        tenant whose ``m365_gdap`` is not TRUE and whose ``m365_username`` is blank has no
        way in, by design: it is left out, with a warning.

    Raises:
        TenantFileError: If a group named holds a cell that is not TRUE, FALSE or blank, or
            a selected tenant's cloud is not a known one.
        TenantSelectionError: If a value names nothing or could mean two things, or
            ``--include`` selects no tenant.
    """
    chosen = select(tenants, tenant_file, args.include, args.exclude or ())
    detailed = [tenant_file.details(tenant, default_cloud=cloud) for tenant in chosen]
    no_access = [tenant for tenant in detailed if not (tenant.gdap or tenant.username)]
    if no_access:
        logger.warning(
            "Skipping %d tenant(s) with neither m365_gdap TRUE nor an m365_username: %s",
            len(no_access),
            ", ".join(tenant.display_name for tenant in no_access),
        )
    return [tenant for tenant in detailed if tenant.gdap or tenant.username]


def _script_name(module: ModuleType) -> str:
    """Name a script for logs and error messages: its file name without ``.py``.

    Args:
        module: The script's module.

    Returns:
        The file's stem, or the module's name when it has no file.
    """
    file = getattr(module, "__file__", None)
    return Path(file).stem if file else module.__name__


def _parser(module: ModuleType) -> argparse.ArgumentParser:
    """Build the parser :func:`main` uses for a script that brings none.

    Args:
        module: The script's module; the first line of its docstring describes it.

    Returns:
        A parser with every flag from :func:`add_arguments`.
    """
    doc = (module.__doc__ or "").strip()
    parser = argparse.ArgumentParser(description=doc.splitlines()[0] if doc else None)
    add_arguments(parser)
    return parser


def _report(payload: Payload, results: list[TenantResult]) -> bool:
    """Hand every tenant's result to the script's ``report()``, when it has one.

    Args:
        payload: The validated script.
        results: One result per tenant, in input order.

    Returns:
        False when ``report()`` raised; the error is logged with its traceback.
    """
    if payload.report is None:
        return True
    logger.debug("Calling %s's report() with %d result(s)", payload.name, len(results))
    try:
        payload.report(results)
    except Exception:
        logger.exception("%s's report() failed", payload.name)
        return False
    return True


def _run(
    module: ModuleType, args: argparse.Namespace, *, transport: httpx.BaseTransport | None
) -> int:
    """Validate the script, sign in, pick tenants, run them, and call ``report()``.

    The script is checked before sign-in, so a mistake in it never costs a browser prompt.

    Args:
        module: The script's module.
        args: Parsed arguments.
        transport: Optional httpx transport for every client (tests).

    Returns:
        The process exit code.
    """
    problem = _selection_error(args)
    if problem:
        logger.error("%s", problem)
        return EXIT_USAGE
    try:
        payload = payload_from_module(module, _script_name(module))
        settings = settings_from_args(args)
        _start_debug_log(settings)
        if settings.jobs < 1:
            raise ConfigError(f"jobs must be at least 1, not {settings.jobs!r}.")
        session = PartnerSession(settings, transport=transport)
        session.login(payload.clients)
        tenants = session.list_tenants()
        tenant_file = _load_tenant_file(settings, session, tenants)
        selected = _select_tenants(args, tenants, tenant_file, session.cloud_name)
    except SETUP_ERRORS as exc:
        logger.error("%s", exc)
        return EXIT_USAGE

    # Every prompt happens in check_access(), one at a time, on this thread, and each tenant
    # starts as soon as it hands it over.
    ready = session.check_access(
        selected,
        scopes=payload.scopes,
        clients=payload.clients,
        jobs=settings.jobs,
        grant_when_possible=(
            payload.grant_when_possible
            if args.grant_when_possible is None
            else args.grant_when_possible
        ),
    )
    logger.info("Running %s against %d tenant(s)", payload.name, len(selected))
    interrupted = False
    try:
        results = run_payload(
            session,
            payload,
            selected,
            ready=ready,
            continue_on_error=(
                payload.continue_on_error
                if args.continue_on_error is None
                else args.continue_on_error
            ),
            jobs=settings.jobs,
        )
    except RunInterrupted as exc:
        results = exc.results
        interrupted = True

    # A report built from part of a run would look complete, so an interrupted run has none.
    reported = interrupted or _report(payload, results)
    print(f"Summary: {summarize(results)}", file=sys.stderr)
    if interrupted:
        return EXIT_INTERRUPTED
    if not reported:
        return EXIT_USAGE
    return EXIT_OK if all(r.status == "ok" for r in results) else EXIT_TENANT_FAILED


def main(
    module: str | ModuleType,
    args: argparse.Namespace | None = None,
    *,
    transport: httpx.BaseTransport | None = None,
) -> int:
    """Run a script's ``run(ctx)`` against the tenants its command line selects.

    Call it from the script's ``__main__`` block as ``raise SystemExit(main(__name__))``.

    Args:
        module: The script: its module, or its name in ``sys.modules`` (``__name__``).
        args: What the script's own parser returned, when it has flags of its own and added
            the library's with :func:`add_arguments`. None parses ``sys.argv`` with the
            library's flags only.
        transport: Optional httpx transport for every client (tests).

    Returns:
        The process exit code.
    """
    if isinstance(module, str):
        module = sys.modules[module]
    if args is None:
        args = _parser(module).parse_args()
    configure_logging(args.verbose)
    try:
        return _run(module, args, transport=transport)
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        return EXIT_INTERRUPTED


def list_tenants(
    argv: Sequence[str] | None = None, *, transport: httpx.BaseTransport | None = None
) -> int:
    """The ``m365-gdap-runner-customer-export`` command: update the tenants file, list tenants.

    The file gets the columns, a row for every GDAP tenant and the clouds it lacks. It is the
    only command that writes the file; runs only warn. Run it once to create the file, then
    add the group columns, and again whenever a run warns.

    Args:
        argv: Arguments without the program name; None reads ``sys.argv``.
        transport: Optional httpx transport for every client (tests).

    Returns:
        The process exit code.
    """
    parser = argparse.ArgumentParser(
        prog="m365-gdap-runner-customer-export",
        description="Sign in, add the GDAP tenants the tenants file lacks, and list them.",
    )
    _add_settings_arguments(parser)
    args = parser.parse_args(argv)
    configure_logging(args.verbose)
    try:
        settings = settings_from_args(args)
        _start_debug_log(settings)
        session = PartnerSession(settings, transport=transport)
        session.login()
        tenants = session.list_tenants()
        _update_tenant_file(settings, session, tenants)
    except SETUP_ERRORS as exc:
        logger.error("%s", exc)
        return EXIT_USAGE
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        return EXIT_INTERRUPTED
    _print_tenants(tenants)
    return EXIT_OK


# This block runs only when the file is executed directly, not when it is
# imported as a module.
if __name__ == "__main__":
    raise SystemExit(list_tenants())
