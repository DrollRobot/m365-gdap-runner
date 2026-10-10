"""Command-line entry point: ``m365-gdap-runner``.

Signs in once per client the payload uses, lists GDAP customer tenants, and runs
one payload against the selected tenants::

    uv run m365-gdap-runner examples/licensed_users.py --tenant Contoso --format csv

Configuration comes from the config system (see ``src/m365_gdap_runner/config/``);
``--partner-tenant-id``, ``--username``, ``--client-id``, ``--cache`` and ``--cloud``
override it. Progress and the summary go to stderr; stdout carries only output
(``--output -``).

Exit codes: 0 all tenants ok; 1 any tenant not ok; 2 usage, configuration,
payload, sign-in, enumeration, or tenant-selection error; 130 interrupted.
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
from m365_gdap_runner.tenants import Tenant, TenantSelectionError, select

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
        "--tenant",
        action="append",
        default=None,
        metavar="NAME_OR_ID",
        help="Run only this tenant (repeatable).",
    )
    parser.add_argument(
        "--exclude",
        action="append",
        default=None,
        metavar="NAME_OR_ID",
        help="Skip this tenant (repeatable).",
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
    if args.payload is None and not args.list_tenants:
        parser.error("PAYLOAD is required unless --list-tenants is given.")
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
        if payload is None:
            _print_tenants(tenants)
            return EXIT_OK
        selected = select(tenants, include=args.tenant, exclude=args.exclude)
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
