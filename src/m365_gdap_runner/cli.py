"""Running a script across tenants, and the ``m365-gdap-runner`` command.

A script runs itself: its ``__main__`` block hands the module to :func:`main`, which reads
the command line, signs in once per client the script uses, lists GDAP customer tenants, and
calls the script's ``run(ctx)`` for each selected tenant::

    if __name__ == "__main__":
        from m365_gdap_runner import main

        raise SystemExit(main(__name__))

then ``uv run audit.py --group pilot``. A script with flags of its own builds the parser,
adds the library's flags with :func:`add_arguments`, and passes what it parsed to ``main()``.

The library writes no output of its own. ``run()`` writes any per-tenant report; after the
last tenant, the script's ``report(results)``, when it has one, gets every tenant's result and
writes the run's.

Tenants are chosen with ``--group`` (a TRUE/FALSE column of the tenants file, see
``tenants.py``), ``--tenant`` or ``--all``, less ``--exclude``; a run that names none is
refused. ``--group`` and ``--tenant`` may pick tenants without a GDAP relationship, or in
another cloud, from the tenants file; before the loop, every tenant the partner sign-in
cannot reach gets a sign-in of its own, unless ``--gdap-only`` is given. Every run appends
GDAP tenants the file lacks. The ``m365-gdap-runner`` command (:func:`list_tenants`) does only
that, and prints the list.

Configuration comes from the config system (see ``src/m365_gdap_runner/config/``);
``--partner-tenant-id``, ``--username``, ``--client-id``, ``--cache``, ``--cloud`` and
``--tenants-file`` override it. Progress and the summary go to stderr; stdout is the
script's.

Exit codes: 0 all tenants ok; 1 any tenant not ok; 2 usage, configuration, script, sign-in,
enumeration, tenants-file or tenant-selection error, or ``report()`` failed; 130 interrupted.
"""

from __future__ import annotations

import argparse
import logging
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

# Library loggers kept at WARNING unless -vv is given.
_LIBRARY_LOGGERS = (
    "httpx",
    "httpcore",
    "msal",
    "msal_extensions",
    "urllib3",
    "azure",
    "azure_auth",
)

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
        "skip the browser sign-in.",
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
        "--group",
        action="append",
        default=None,
        metavar="NAME",
        help=(
            "Run the tenants marked TRUE in this column of the tenants file "
            "(repeatable; groups add up)."
        ),
    )
    tenants.add_argument(
        "--tenant",
        action="append",
        default=None,
        metavar="NAME_OR_ID",
        help=(
            "Run this tenant, by display name or ID: a GDAP customer or a row of the tenants "
            "file (repeatable)."
        ),
    )
    tenants.add_argument("--all", action="store_true", help="Run every GDAP tenant.")
    tenants.add_argument(
        "--exclude",
        action="append",
        default=None,
        metavar="NAME_OR_ID",
        help="Skip this tenant (repeatable).",
    )
    tenants.add_argument(
        "--gdap-only",
        action="store_true",
        help=(
            "Reach tenants through the partner sign-in only: a tenant it cannot reach is "
            "auth_failed instead of getting a sign-in of its own."
        ),
    )
    run = parser.add_argument_group("m365-gdap-runner run")
    run.add_argument(
        "--continue-on-error",
        action="store_true",
        help="Keep running other tenants after an error in run().",
    )
    run.add_argument(
        "--jobs",
        type=_positive_int,
        default=1,
        metavar="N",
        help=(
            "Tenants to run concurrently (default: 1). 5-8 suits read-only reports; "
            "keep 1 for scripts that write unless they are idempotent."
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
    return load_settings(profile=args.profile, overrides=overrides)


def _selection_error(args: argparse.Namespace) -> str | None:
    """Check that the flags choose tenants, and do so one way.

    Args:
        args: Parsed arguments.

    Returns:
        What is wrong, or None when the selection is usable.
    """
    picked = args.tenant or args.group
    if args.all and picked:
        return "--all cannot be combined with --tenant or --group."
    if not (args.all or picked):
        return "Choose tenants with --group, --tenant or --all."
    return None


def configure_logging(verbosity: int) -> None:
    """Send logs to stderr, tagged with the tenant each record belongs to.

    Args:
        verbosity: Count of ``-v`` flags.
    """
    handler = logging.StreamHandler(sys.stderr)
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


def _print_tenants(tenants: Sequence[Tenant]) -> None:
    """Print one ``tenant_id<TAB>display_name`` line per tenant to stdout.

    Args:
        tenants: Tenants to print.
    """
    for tenant in tenants:
        print(f"{tenant.tenant_id}\t{tenant.display_name}")


def _sync_tenant_file(
    settings: Settings, session: PartnerSession, tenants: Sequence[Tenant]
) -> TenantFile:
    """Load the tenants file and bring it up to date with the GDAP tenants.

    A GDAP tenant with no row is appended; one whose cloud is blank gets it looked up. A
    missing file is created. A file that cannot be written (a spreadsheet may hold it open)
    is reported and left as it was; the run goes on with the rows in memory.

    Args:
        settings: Resolved settings, for the file's location.
        session: The partner session, for cloud lookups.
        tenants: The GDAP tenants.

    Returns:
        The file, with every GDAP tenant in it.

    Raises:
        TenantFileError: If the file exists but cannot be read or lacks a required column.
    """
    path = tenants_file_path(settings)
    existed = path.exists()
    tenant_file = TenantFile.load(path) if existed else TenantFile.empty(path)
    report = tenant_file.sync(tenants, session.cloud_of)
    if existed and not report.changed:
        logger.debug("Tenants file %s is up to date", path)
        return tenant_file
    try:
        tenant_file.save()
    except OSError as exc:
        logger.warning("Could not write tenants file %s: %s", path, exc)
        return tenant_file
    if not existed:
        logger.info(
            "Created tenants file %s with %d tenant(s); add a TRUE/FALSE column per group",
            path,
            len(report.added),
        )
        return tenant_file
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
    return tenant_file


def _select_tenants(
    args: argparse.Namespace, tenants: Sequence[Tenant], tenant_file: TenantFile, cloud: str
) -> list[Tenant]:
    """Apply ``--all``, ``--group``, ``--tenant`` and ``--exclude``.

    ``--all`` means every GDAP tenant. ``--group`` and ``--tenant`` may also pick a row of the
    tenants file that GDAP does not list.

    Args:
        args: Parsed arguments that passed :func:`_selection_error`, so one of ``--all``,
            ``--group`` or ``--tenant`` was given.
        tenants: The GDAP tenants.
        tenant_file: The tenants file.
        cloud: Name of the partner's cloud, for a tenant whose row names none.

    Returns:
        The tenants to run, GDAP tenants first, each with its cloud and sign-in account.

    Raises:
        TenantFileError: If a group has no column, a cell in it is not TRUE, FALSE or blank,
            or a selected tenant's cloud is not a known one.
        TenantSelectionError: If a value names no tenant, or the groups select none.
    """
    known = [*tenants, *tenant_file.extra_tenants(tenants)]
    if args.all:
        gdap = {tenant.tenant_id for tenant in tenants}
        chosen = [t for t in select(known, exclude=args.exclude) if t.tenant_id in gdap]
    else:
        include = list(args.tenant or ())
        for group in args.group or ():
            include += [row.tenant_id for row in tenant_file.members(group)]
        if not include:
            groups = ", ".join(args.group or ())
            raise TenantSelectionError(f"No tenant is marked TRUE in group(s) {groups}.")
        chosen = select(known, include=include, exclude=args.exclude)
    return [tenant_file.details(tenant, default_cloud=cloud) for tenant in chosen]


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
        session = PartnerSession(settings, transport=transport)
        session.login(payload.clients)
        tenants = session.list_tenants()
        tenant_file = _sync_tenant_file(settings, session, tenants)
        selected = _select_tenants(args, tenants, tenant_file, session.cloud_name)
    except SETUP_ERRORS as exc:
        logger.error("%s", exc)
        return EXIT_USAGE

    # Every prompt happens here, one at a time, so the loop below can run in parallel.
    session.check_access(
        selected,
        scopes=payload.scopes,
        clients=payload.clients,
        direct=not args.gdap_only,
        jobs=args.jobs,
    )
    logger.info("Running %s against %d tenant(s)", payload.name, len(selected))
    interrupted = False
    try:
        results = run_payload(
            session,
            payload,
            selected,
            continue_on_error=args.continue_on_error,
            jobs=args.jobs,
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
    """The ``m365-gdap-runner`` command: sign in, update the tenants file, list the tenants.

    The file gets a row for every GDAP tenant it lacks, as on every run; run this once to
    create it, then add the group columns.

    Args:
        argv: Arguments without the program name; None reads ``sys.argv``.
        transport: Optional httpx transport for every client (tests).

    Returns:
        The process exit code.
    """
    parser = argparse.ArgumentParser(
        prog="m365-gdap-runner",
        description="Sign in, add the GDAP tenants the tenants file lacks, and list them.",
    )
    _add_settings_arguments(parser)
    args = parser.parse_args(argv)
    configure_logging(args.verbose)
    try:
        settings = settings_from_args(args)
        session = PartnerSession(settings, transport=transport)
        session.login()
        tenants = session.list_tenants()
        _sync_tenant_file(settings, session, tenants)
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
