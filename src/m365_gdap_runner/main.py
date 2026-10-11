"""Command-line entry point: ``m365-gdap-runner``.

Signs in once per client the payload uses, lists GDAP customer tenants, and runs
one payload against the selected tenants::

    uv run m365-gdap-runner examples/licensed_users.py --group pilot --format csv

Tenants are chosen with ``--group`` (a TRUE/FALSE column of the tenants file, see
``tenants.py``), ``--tenant`` or ``--all``, less ``--exclude``; a run that names none is
refused. Every run appends GDAP tenants the file lacks, and ``--list-tenants`` does only that.

Configuration comes from the config system (see ``src/m365_gdap_runner/config/``);
``--partner-tenant-id``, ``--username``, ``--client-id``, ``--cache``, ``--cloud`` and
``--tenants-file`` override it. Progress and the summary go to stderr; stdout carries only
output (``--output -``).

Exit codes: 0 all tenants ok; 1 any tenant not ok; 2 usage, configuration,
payload, sign-in, enumeration, tenants-file, or tenant-selection error; 130 interrupted.
"""

from __future__ import annotations

import argparse
import logging
import sys
from collections.abc import Sequence

import httpx
from azure_auth import AuthError, ResourceError

from m365_gdap_runner.config import ConfigError, Settings, load_settings
from m365_gdap_runner.loader import PayloadInvalid, PayloadNotFound, load_payload
from m365_gdap_runner.output import FORMATS, STDOUT, OutputFormatError, write_results
from m365_gdap_runner.runner import RunInterrupted, TenantLogFilter, run_payload, summarize
from m365_gdap_runner.session import CACHE_KINDS, PartnerSession
from m365_gdap_runner.tenants import (
    Tenant,
    TenantFile,
    TenantFileError,
    TenantSelectionError,
    group_members,
    select,
    tenants_file_path,
)

logger = logging.getLogger(__name__)

EXIT_OK = 0
EXIT_TENANT_FAILED = 1
EXIT_USAGE = 2
EXIT_INTERRUPTED = 130

# Errors that end a run before any payload executes.
SETUP_ERRORS = (
    ConfigError,
    PayloadNotFound,
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
)


def add_settings_arguments(parser: argparse.ArgumentParser) -> None:
    """Add the flags that override settings (shared with :func:`standalone`).

    Args:
        parser: Parser to extend.
    """
    parser.add_argument(
        "--profile",
        default=None,
        help="Config profile to use (default: default_profile from config.toml).",
    )
    for name, metavar, help_text in _SETTING_FLAGS:
        parser.add_argument(
            "--" + name.replace("_", "-"), default=None, metavar=metavar, help=help_text
        )
    parser.add_argument(
        "--cache",
        choices=CACHE_KINDS,
        default=None,
        help="Token cache: memory (default) or disk, an encrypted file that lets later runs "
        "skip the browser sign-in.",
    )


def settings_from_args(args: argparse.Namespace) -> Settings:
    """Resolve settings with the flags from :func:`add_settings_arguments` as overrides.

    ``--tenants-file`` belongs to the runner command only, so it is read when present.

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
    tenants_file = getattr(args, "tenants_file", None)
    if tenants_file is not None:
        overrides["tenants_file"] = tenants_file
    return load_settings(profile=args.profile, overrides=overrides)


def build_parser() -> argparse.ArgumentParser:
    """Build the command-line parser.

    Returns:
        The parser.
    """
    parser = argparse.ArgumentParser(
        prog="m365-gdap-runner",
        description="Run a payload's run(ctx) against GDAP customer tenants from one sign-in.",
    )
    parser.add_argument(
        "payload", nargs="?", metavar="PAYLOAD", help="Payload .py file or installed payload name."
    )
    parser.add_argument(
        "--group",
        action="append",
        default=None,
        metavar="NAME",
        help=(
            "Run the tenants marked TRUE in this column of the tenants file "
            "(repeatable; groups add up)."
        ),
    )
    parser.add_argument(
        "--tenant",
        action="append",
        default=None,
        metavar="NAME_OR_ID",
        help="Run this tenant, by display name or ID (repeatable).",
    )
    parser.add_argument("--all", action="store_true", help="Run every GDAP tenant.")
    parser.add_argument(
        "--exclude",
        action="append",
        default=None,
        metavar="NAME_OR_ID",
        help="Skip this tenant (repeatable).",
    )
    parser.add_argument(
        "--tenants-file",
        default=None,
        metavar="PATH",
        help=(
            "Tenants CSV (default: the tenants_file setting, else tenants.csv in the config "
            "directory)."
        ),
    )
    parser.add_argument(
        "--output", default=STDOUT, metavar="PATH", help="Output file, or - for stdout (default)."
    )
    parser.add_argument(
        "--format", choices=FORMATS, default="json", help="Output format (default: json)."
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="Keep running other tenants after a payload error.",
    )
    parser.add_argument(
        "--jobs",
        type=int,
        default=1,
        metavar="N",
        help=(
            "Tenants to run concurrently (default: 1). 5-8 suits read-only reports; "
            "keep 1 for payloads that write unless they are idempotent."
        ),
    )
    add_settings_arguments(parser)
    parser.add_argument(
        "--list-tenants", action="store_true", help="Sign in, list GDAP tenants, and exit."
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="-v: debug logs from this tool and payloads; -vv: also from libraries.",
    )
    return parser


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """Parse and validate command-line arguments.

    Args:
        argv: Arguments without the program name; None reads ``sys.argv``.

    Returns:
        The parsed arguments.

    Raises:
        SystemExit: With code 2 on invalid usage (argparse behavior).
    """
    parser = build_parser()
    args = parser.parse_args(argv)
    if not args.list_tenants:
        if args.payload is None:
            parser.error("PAYLOAD is required unless --list-tenants is given.")
        picked = args.tenant or args.group
        if args.all and picked:
            parser.error("--all cannot be combined with --tenant or --group.")
        if not (args.all or picked):
            parser.error("Choose tenants with --group, --tenant or --all.")
    if args.jobs < 1:
        parser.error("--jobs must be at least 1.")
    return args


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


def _sync_tenant_file(settings: Settings, tenants: Sequence[Tenant]) -> TenantFile:
    """Load the tenants file, appending a row for every GDAP tenant it lacks.

    A missing file is created. A file that cannot be written (a spreadsheet may hold it
    open) is reported and left as it was; the run goes on with the rows in memory.

    Args:
        settings: Resolved settings, for the file's location.
        tenants: The GDAP tenants.

    Returns:
        The file, with every GDAP tenant in it.

    Raises:
        TenantFileError: If the file exists but cannot be read or lacks a required column.
    """
    path = tenants_file_path(settings)
    existed = path.exists()
    tenant_file = TenantFile.load(path) if existed else TenantFile.empty(path)
    added = tenant_file.add_missing(tenants)
    if existed and not added:
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
            len(added),
        )
    else:
        names = ", ".join(tenant.display_name for tenant in added)
        logger.warning(
            "Added %d tenant(s) to %s; they are in no group yet: %s", len(added), path, names
        )
    return tenant_file


def _select_tenants(
    args: argparse.Namespace, tenants: Sequence[Tenant], tenant_file: TenantFile
) -> list[Tenant]:
    """Apply ``--all``, ``--group``, ``--tenant`` and ``--exclude``.

    Args:
        args: Arguments from :func:`parse_args`, which guarantees one of ``--all``,
            ``--group`` or ``--tenant`` was given.
        tenants: The GDAP tenants.
        tenant_file: The tenants file, for ``--group``.

    Returns:
        The tenants to run, in enumeration order.

    Raises:
        TenantFileError: If a group has no column, or a cell in it is not TRUE, FALSE or blank.
        TenantSelectionError: If a value names no usable tenant, or the groups select none.
    """
    if args.all:
        return select(tenants, exclude=args.exclude)
    include = list(args.tenant or ())
    include += [t.tenant_id for t in group_members(tenants, tenant_file, args.group or ())]
    if not include:
        groups = ", ".join(args.group or ())
        raise TenantSelectionError(f"No tenant is marked TRUE in group(s) {groups}.")
    return select(tenants, include=include, exclude=args.exclude)


def run(args: argparse.Namespace, *, transport: httpx.BaseTransport | None = None) -> int:
    """Execute a parsed command line.

    The payload is loaded before sign-in, so payload errors never cost a
    browser prompt.

    Args:
        args: Arguments from :func:`parse_args`.
        transport: Optional httpx transport for every client (tests).

    Returns:
        The process exit code.
    """
    try:
        settings = settings_from_args(args)
        payload = None if args.list_tenants else load_payload(args.payload)
        session = PartnerSession(settings, transport=transport)
        session.login(payload.clients if payload else ())
        tenants = session.list_tenants()
        tenant_file = _sync_tenant_file(settings, tenants)
        if payload is None:
            _print_tenants(tenants)
            return EXIT_OK
        selected = _select_tenants(args, tenants, tenant_file)
    except SETUP_ERRORS as exc:
        logger.error("%s", exc)
        return EXIT_USAGE

    logger.info("Running %s against %d tenant(s)", payload.name, len(selected))
    interrupted = False
    try:
        results = run_payload(
            session,
            payload,
            selected,
            continue_on_error=args.continue_on_error,
            jobs=args.jobs,
            output_format=args.format,
        )
    except RunInterrupted as exc:
        results = exc.results
        interrupted = True

    try:
        write_results(results, args.output, args.format)
    except (OSError, OutputFormatError) as exc:
        logger.error("Could not write output: %s", exc)
        return EXIT_USAGE
    print(f"Summary: {summarize(results)}", file=sys.stderr)
    if interrupted:
        return EXIT_INTERRUPTED
    return EXIT_OK if all(r.status == "ok" for r in results) else EXIT_TENANT_FAILED


def main(argv: Sequence[str] | None = None) -> int:
    """Console-script entry point.

    Args:
        argv: Arguments without the program name; None reads ``sys.argv``.

    Returns:
        The process exit code.
    """
    args = parse_args(argv)
    configure_logging(args.verbose)
    try:
        return run(args)
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        return EXIT_INTERRUPTED


# This block runs only when the file is executed directly, not when it is
# imported as a module.
if __name__ == "__main__":
    raise SystemExit(main())
