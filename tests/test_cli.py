"""Tests for main(), add_arguments() and the m365-gdap-runner command in cli.py.

End-to-end runs use the fake msal module and a fake service transport.
"""

from __future__ import annotations

import argparse
import contextlib
import json
import logging
import os
import sys
import textwrap
from collections.abc import Callable, Iterator
from pathlib import Path
from types import ModuleType

import httpx
import pytest

from m365_gdap_runner import cli
from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.config.schema import ENV_PREFIX
from m365_gdap_runner.runner import TenantResult
from m365_gdap_runner.tenants import TenantFile
from tests._gdap_fakes import (
    CONSENT_REQUIRED,
    GAMMA_ADMIN,
    NOT_IN_TENANT,
    PARTNER_TENANT,
    TENANT_A,
    TENANT_B,
    TENANT_C,
    USER,
    FakeMsal,
    bearer_tenant,
    fake_services,
    load_script,
)

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"
LICENSED_USERS = load_script(EXAMPLES / "licensed_users.py")
AUDIT_DISABLED = load_script(EXAMPLES / "audit_disabled.py")

# The real one: the fixture below replaces it so tests do not stack handlers on the root logger.
CONFIGURE_LOGGING = cli.configure_logging

# A script whose report() keeps what it is given, so a test can see every tenant's result.
RECORDER = """
    SCOPES = ["User.Read.All"]
    REPORTED = []

    def run(ctx):
        return [user["userPrincipalName"] for user in ctx.graph.get_all("/users")]

    def report(results):
        REPORTED.append(results)
"""

Handler = Callable[[httpx.Request], httpx.Response]


@pytest.fixture(autouse=True)
def _config(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Isolate config and supply the partner tenant and user through the environment."""
    for key in list(os.environ):
        if key.startswith(ENV_PREFIX):
            monkeypatch.delenv(key)
    monkeypatch.setenv(CONFIG_DIR_ENV, str(tmp_path / "config"))
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", PARTNER_TENANT)
    monkeypatch.setenv(ENV_PREFIX + "USERNAME", USER)
    monkeypatch.setattr(cli, "configure_logging", lambda verbosity: None)


@pytest.fixture
def recorder(tmp_path: Path) -> ModuleType:
    return _script(tmp_path, RECORDER)


def _args(*argv: str) -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    cli.add_arguments(parser)
    return parser.parse_args(list(argv))


def _main(script: ModuleType, *argv: str, handler: Handler = fake_services) -> int:
    return cli.main(script, _args(*argv), transport=httpx.MockTransport(handler))


def _list_tenants(*argv: str) -> int:
    return cli.list_tenants(list(argv), transport=httpx.MockTransport(fake_services))


def _script(tmp_path: Path, body: str, name: str = "script") -> ModuleType:
    path = tmp_path / f"{name}.py"
    path.write_text(textwrap.dedent(body), encoding="utf-8")
    return load_script(path)


def _reported(script: ModuleType) -> list[TenantResult]:
    """Return what the script's report() was given by the latest run."""
    reported: list[list[TenantResult]] = script.__dict__["REPORTED"]
    return reported[-1]


def _tenant_ids(script: ModuleType) -> list[str]:
    return [result.tenant.tenant_id for result in _reported(script)]


# --- arguments -------------------------------------------------------------------------


@pytest.mark.unit
def test_add_arguments_defaults() -> None:
    args = _args("--all")
    assert (args.jobs, args.continue_on_error) == (1, False)
    assert (args.tenant, args.group, args.exclude, args.all) == (None, None, None, True)
    assert (args.partner_tenant_id, args.username, args.client_id, args.cache, args.cloud) == (
        None,
        None,
        None,
        None,
        None,
    )
    assert (args.tenants_file, args.profile, args.verbose) == (None, None, 0)


@pytest.mark.unit
@pytest.mark.parametrize(
    "argv",
    [["--jobs", "0"], ["--jobs", "two"], ["--cache", "plaintext"], ["--format", "json"]],
    ids=str,
)
def test_the_parser_refuses_bad_values(argv: list[str]) -> None:
    with pytest.raises(SystemExit) as info:
        _args("--all", *argv)
    assert info.value.code == 2


@pytest.mark.unit
def test_the_parser_main_builds_is_described_by_the_scripts_docstring() -> None:
    parser = cli._parser(LICENSED_USERS)
    assert parser.description == "Example script: licensed users in each tenant."
    assert "--group" in parser.format_help()


@pytest.mark.unit
def test_main_returns_130_on_interrupt(monkeypatch: pytest.MonkeyPatch) -> None:
    def interrupted(*args: object, **kwargs: object) -> int:
        raise KeyboardInterrupt

    monkeypatch.setattr(cli, "_run", interrupted)
    assert cli.main(LICENSED_USERS, _args("--all")) == 130


@contextlib.contextmanager
def _restoring_loggers() -> Iterator[None]:
    """Undo what configure_logging() and the debug log do to the root and library loggers.

    Use it inside a test, not a fixture: pytest swaps its own root handlers between phases.
    """
    root = logging.getLogger()
    handlers = list(root.handlers)
    saved = {name: logging.getLogger(name).level for name in ("", *cli._LIBRARY_LOGGERS)}
    try:
        yield
    finally:
        for handler in root.handlers[len(handlers) :]:
            root.removeHandler(handler)
            handler.close()
        for name, level in saved.items():
            logging.getLogger(name).setLevel(level)


@pytest.mark.unit
def test_configure_logging_quiets_libraries_unless_very_verbose() -> None:
    root = logging.getLogger()
    with _restoring_loggers():
        CONFIGURE_LOGGING(0)
        assert root.level == logging.INFO
        assert logging.getLogger("azure_auth").level == logging.WARNING
        CONFIGURE_LOGGING(2)
        assert root.level == logging.DEBUG
        assert logging.getLogger("msal").level == logging.DEBUG


@pytest.mark.unit
@pytest.mark.parametrize(
    ("name", "level", "kept"),
    [
        ("httpx", logging.DEBUG, False),
        ("azure.identity", logging.INFO, False),
        ("httpx", logging.WARNING, True),
        ("httpx_extras", logging.INFO, True),
        ("m365_gdap_runner.session", logging.DEBUG, True),
    ],
)
def test_the_terminal_drops_library_records_below_warning(
    name: str, level: int, kept: bool
) -> None:
    record = logging.LogRecord(name, level, __file__, 1, "message", None, None)
    assert cli._quiet_libraries(record) is kept


@pytest.mark.integration
def test_settings_flags_override_config() -> None:
    args = _args(
        "--partner-tenant-id",
        "other.example",
        "--username",
        "someone@other.example",
        "--client-id",
        "custom",
        "--cache",
        "disk",
        "--cloud",
        "USGov",
        "--tenants-file",
        "C:/msp/tenants.csv",
        "--all",
    )
    settings = cli.settings_from_args(args)
    assert (
        settings.partner_tenant_id,
        settings.username,
        settings.client_id,
        settings.cache,
        settings.cloud,
        settings.tenants_file,
    ) == ("other.example", "someone@other.example", "custom", "disk", "USGov", "C:/msp/tenants.csv")


# --- how a script calls main() ---------------------------------------------------------


@pytest.mark.integration
@pytest.mark.acceptance
def test_main_reads_the_command_line_when_the_script_brings_no_parser(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, recorder: ModuleType
) -> None:
    monkeypatch.setattr(sys, "argv", ["script.py", "--tenant", "Alpha"])
    assert cli.main(recorder, transport=httpx.MockTransport(fake_services)) == 0
    assert _tenant_ids(recorder) == [TENANT_A]


@pytest.mark.integration
def test_main_finds_the_script_by_module_name(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, recorder: ModuleType
) -> None:
    monkeypatch.setitem(sys.modules, "__main__", recorder)
    assert cli.main("__main__", _args("--all"), transport=httpx.MockTransport(fake_services)) == 0
    assert _tenant_ids(recorder) == [TENANT_A, TENANT_B]


@pytest.mark.integration
@pytest.mark.acceptance
def test_a_script_adds_flags_of_its_own(fake_msal: FakeMsal, tmp_path: Path) -> None:
    script = _script(
        tmp_path,
        """
        CLIENTS = ["exchange"]
        LABEL = "unset"
        REPORTED = []

        def run(ctx):
            return {"label": LABEL}

        def report(results):
            REPORTED.append(results)
        """,
    )
    parser = argparse.ArgumentParser()
    parser.add_argument("--label", default="none")
    cli.add_arguments(parser)
    assert "--label" in parser.format_help()
    assert "--group" in parser.format_help()

    args = parser.parse_args(["--label", "weekly", "--tenant", "Alpha"])
    script.__dict__["LABEL"] = args.label  # what the script's __main__ block does with it
    assert cli.main(script, args, transport=httpx.MockTransport(fake_services)) == 0
    assert [result.result for result in _reported(script)] == [{"label": "weekly"}]


@pytest.mark.integration
@pytest.mark.acceptance
def test_main_requires_a_tenant_selection(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture
) -> None:
    assert _main(LICENSED_USERS) == 2
    assert "Choose tenants with --group, --tenant or --all." in caplog.text
    assert fake_msal.interactive_calls == 0


@pytest.mark.integration
@pytest.mark.parametrize("other", [["--tenant", "Alpha"], ["--group", "pilot"]])
def test_main_refuses_all_with_a_narrower_selection(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture, other: list[str]
) -> None:
    assert _main(LICENSED_USERS, "--all", *other) == 2
    assert "--all cannot be combined with --tenant or --group." in caplog.text


# --- report() --------------------------------------------------------------------------


@pytest.mark.integration
@pytest.mark.acceptance
def test_report_gets_every_tenant_in_order_failures_included(
    fake_msal: FakeMsal, recorder: ModuleType, capsys: pytest.CaptureFixture[str]
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    assert _main(recorder, "--all", "--gdap-only") == 1
    (alpha, beta) = _reported(recorder)
    assert (alpha.tenant.tenant_id, alpha.status, alpha.result) == (
        TENANT_A,
        "ok",
        [f"admin@{TENANT_A}.example"],
    )
    assert (beta.tenant.tenant_id, beta.status, beta.result) == (TENANT_B, "auth_failed", None)
    assert beta.error is not None
    assert "AADSTS65001" in beta.error
    captured = capsys.readouterr()
    assert captured.out == ""  # the library writes nothing; stdout is the script's
    assert "Summary: ok=1, auth_failed=1, error=0, skipped=0" in captured.err


@pytest.mark.integration
def test_a_script_without_report_runs_and_writes_nothing(
    fake_msal: FakeMsal, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    script = _script(tmp_path, 'SCOPES = ["User.Read.All"]\n\ndef run(ctx):\n    return 1\n')
    assert _main(script, "--all") == 0
    captured = capsys.readouterr()
    assert captured.out == ""
    assert "Summary: ok=2, auth_failed=0, error=0, skipped=0" in captured.err


@pytest.mark.integration
@pytest.mark.acceptance
def test_exit_2_when_report_fails(
    fake_msal: FakeMsal,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    caplog: pytest.LogCaptureFixture,
) -> None:
    script = _script(
        tmp_path,
        """
        SCOPES = ["User.Read.All"]

        def run(ctx):
            return 1

        def report(results):
            raise RuntimeError("disk full")
        """,
        name="audit",
    )
    assert _main(script, "--all") == 2
    assert "audit's report() failed" in caplog.text
    assert "RuntimeError: disk full" in caplog.text
    assert "Summary: ok=2, auth_failed=0, error=0, skipped=0" in capsys.readouterr().err


@pytest.mark.integration
@pytest.mark.acceptance
def test_interrupt_skips_report(
    fake_msal: FakeMsal, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    script = _script(
        tmp_path,
        """
        SCOPES = ["User.Read.All"]
        REPORTED = []

        def run(ctx):
            if ctx.tenant.display_name == "beta corp":
                raise KeyboardInterrupt
            return 1

        def report(results):
            REPORTED.append(results)
        """,
    )
    assert _main(script, "--all") == 130
    assert script.__dict__["REPORTED"] == []
    assert "Summary: ok=1, auth_failed=0, error=1, skipped=0" in capsys.readouterr().err


# --- end to end ------------------------------------------------------------------------


@pytest.mark.integration
def test_runs_the_graph_example_and_its_report(
    fake_msal: FakeMsal, capsys: pytest.CaptureFixture[str]
) -> None:
    assert _main(LICENSED_USERS, "--all", "--jobs", "2") == 0
    rows = json.loads(capsys.readouterr().out)
    assert [(r["tenantId"], r["tenant"], r["status"]) for r in rows] == [
        (TENANT_A, "Alpha", "ok"),
        (TENANT_B, "beta corp", "ok"),
    ]
    assert rows[1]["users"] == [
        {
            "userPrincipalName": f"admin@{TENANT_B}.example",
            "displayName": "Admin",
            "skuIds": "sku-1;sku-2",
        }
    ]
    assert fake_msal.interactive_calls == 1


@pytest.mark.integration
def test_runs_the_exchange_example_whose_log_is_the_report(
    fake_msal: FakeMsal, capsys: pytest.CaptureFixture[str], caplog: pytest.LogCaptureFixture
) -> None:
    caplog.set_level(logging.INFO)
    assert _main(AUDIT_DISABLED, "--all") == 0
    assert capsys.readouterr().out == ""
    warnings = [r.getMessage() for r in caplog.records if r.levelno == logging.WARNING]
    assert warnings == ["Mailbox auditing is off tenant-wide (AuditDisabled is True)"]
    assert fake_msal.interactive_calls == 2


@pytest.mark.integration
def test_tenant_filters(fake_msal: FakeMsal, recorder: ModuleType) -> None:
    assert _main(recorder, "--all", "--exclude", "BETA CORP") == 0
    assert _tenant_ids(recorder) == [TENANT_A]


@pytest.mark.integration
def test_exit_2_on_unknown_tenant_before_running_the_script(fake_msal: FakeMsal) -> None:
    seen: list[str] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request.url.path)
        return fake_services(request)

    assert _main(LICENSED_USERS, "--tenant", "Contoso", handler=handler) == 2
    assert "/v1.0/users" not in seen


@pytest.mark.integration
def test_exit_2_on_invalid_script_before_sign_in(
    fake_msal: FakeMsal, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    script = _script(tmp_path, "def run(ctx):\n    return None\n")
    assert _main(script, "--all") == 2
    assert "Script 'script' must define SCOPES" in caplog.text
    assert fake_msal.interactive_calls == 0


@pytest.mark.integration
def test_exit_2_on_missing_config(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    monkeypatch.delenv(ENV_PREFIX + "USERNAME")
    assert _main(LICENSED_USERS, "--all") == 2
    assert "username" in caplog.text


@pytest.mark.integration
def test_exit_2_on_failed_sign_in(fake_msal: FakeMsal) -> None:
    fake_msal.interactive_result = {"error": "access_denied", "error_description": "cancelled"}
    assert _main(LICENSED_USERS, "--all") == 2


@pytest.mark.integration
def test_undeclared_client_is_a_tenant_error(fake_msal: FakeMsal, tmp_path: Path) -> None:
    script = _script(
        tmp_path,
        """
        SCOPES = ["User.Read.All"]
        REPORTED = []

        def run(ctx):
            return ctx.exchange.run("Get-Mailbox")

        def report(results):
            REPORTED.append(results)
        """,
    )
    assert _main(script, "--tenant", "Alpha") == 1
    (result,) = _reported(script)
    assert result.status == "error"
    assert result.error is not None
    assert "ClientNotDeclared" in result.error


# --- the m365-gdap-runner command ------------------------------------------------------


@pytest.mark.integration
@pytest.mark.acceptance
def test_list_tenants_creates_the_tenants_file_and_prints_the_tenants(
    fake_msal: FakeMsal,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    caplog: pytest.LogCaptureFixture,
) -> None:
    caplog.set_level(logging.INFO, logger=cli.__name__)
    assert _list_tenants() == 0
    assert capsys.readouterr().out.splitlines() == [f"{TENANT_A}\tAlpha", f"{TENANT_B}\tbeta corp"]
    path = tmp_path / "config" / "tenants.csv"
    assert path.read_text(encoding="utf-8-sig").splitlines() == [
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,m365_username",
        f"{TENANT_A},Alpha,,Commercial,",
        f"{TENANT_B},beta corp,,Commercial,",
    ]
    assert f"Created tenants file {path} with 2 tenant(s)" in caplog.text
    assert fake_msal.interactive_calls == 1


@pytest.mark.integration
def test_list_tenants_takes_the_settings_flags(fake_msal: FakeMsal, tmp_path: Path) -> None:
    path = tmp_path / "elsewhere" / "customers.csv"
    assert _list_tenants("--tenants-file", str(path)) == 0
    assert path.exists()
    assert not (tmp_path / "config" / "tenants.csv").exists()


@pytest.mark.integration
def test_list_tenants_exit_2_on_missing_config(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    monkeypatch.delenv(ENV_PREFIX + "USERNAME")
    assert _list_tenants() == 2
    assert "username" in caplog.text


@pytest.mark.integration
@pytest.mark.acceptance
def test_debug_log_gets_every_record_and_the_terminal_info_and_up(
    fake_msal: FakeMsal,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    path = tmp_path / "logs" / "debug.log"
    monkeypatch.setenv(ENV_PREFIX + "DEBUG_LOG_PATH", str(path))
    monkeypatch.setattr(cli, "configure_logging", CONFIGURE_LOGGING)
    for _run in range(2):  # each as its own process would, with handlers removed after
        with _restoring_loggers():
            assert _list_tenants() == 0
    log = path.read_text(encoding="utf-8")
    terminal = capsys.readouterr().err
    # This package's debug records, and a library's below WARNING, reach the file only.
    assert "DEBUG [-] m365_gdap_runner.session: Signing in to the partner tenant" in log
    assert "INFO [-] httpx: HTTP Request" in log
    assert "Signing in to the partner tenant" not in terminal
    assert "HTTP Request" not in terminal
    assert "Created tenants file" in terminal
    # The second run appends.
    assert log.count("Debug log started") == 2


@pytest.mark.integration
def test_a_debug_log_that_cannot_be_opened_is_a_warning_not_an_error(
    fake_msal: FakeMsal,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
) -> None:
    monkeypatch.setenv(ENV_PREFIX + "DEBUG_LOG_PATH", str(tmp_path))  # a folder, not a file
    assert _list_tenants() == 0
    assert f"Could not open debug log {tmp_path}" in caplog.text


@pytest.mark.unit
def test_list_tenants_takes_no_selection_flags(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as info:
        cli.list_tenants(["--all"])
    assert info.value.code == 2
    assert "unrecognized arguments: --all" in capsys.readouterr().err


# --- tenants file ----------------------------------------------------------------------


def _tenants_csv(path: Path, *lines: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


@pytest.mark.integration
@pytest.mark.acceptance
def test_run_appends_tenants_the_file_lacks_and_warns(
    fake_msal: FakeMsal, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,m365_username,pilot,notes",
        f"{TENANT_A},Alpha,alpha.example,Commercial,,TRUE,keep me",
    )
    assert _main(LICENSED_USERS, "--tenant", "Alpha", "--tenants-file", str(path)) == 0
    assert path.read_text(encoding="utf-8-sig").splitlines() == [
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,m365_username,pilot,notes",
        f"{TENANT_A},Alpha,alpha.example,Commercial,,TRUE,keep me",
        f"{TENANT_B},beta corp,,Commercial,,,",
    ]
    assert f"Added 1 tenant(s) to {path}; they are in no group yet: beta corp" in caplog.text


@pytest.mark.integration
@pytest.mark.acceptance
def test_group_runs_only_the_marked_tenants(
    fake_msal: FakeMsal, tmp_path: Path, recorder: ModuleType
) -> None:
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
        f"{TENANT_A},Alpha,,Commercial,FALSE",
        f"{TENANT_B},beta corp,,Commercial,TRUE",
        f"{TENANT_C},Gamma,,,FALSE",  # no GDAP relationship: harmless while unselected
    )
    assert _main(recorder, "--group", "pilot", "--tenants-file", str(path)) == 0
    assert _tenant_ids(recorder) == [TENANT_B]


@pytest.mark.integration
def test_groups_and_tenants_add_up_and_exclude_applies(
    fake_msal: FakeMsal, tmp_path: Path, recorder: ModuleType
) -> None:
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot,patching",
        f"{TENANT_A},Alpha,,Commercial,TRUE,FALSE",
        f"{TENANT_B},beta corp,,Commercial,FALSE,TRUE",
    )
    file_flag = ["--tenants-file", str(path)]
    assert _main(recorder, "--group", "pilot", "--group", "PATCHING", *file_flag) == 0
    assert _tenant_ids(recorder) == [TENANT_A, TENANT_B]
    assert _main(recorder, "--group", "pilot", "--tenant", "beta corp", *file_flag) == 0
    assert _tenant_ids(recorder) == [TENANT_A, TENANT_B]
    assert _main(recorder, "--group", "pilot", "--exclude", "Alpha", *file_flag) == 0
    assert _tenant_ids(recorder) == []


@pytest.mark.integration
def test_tenants_file_setting_locates_the_file(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, tmp_path: Path, recorder: ModuleType
) -> None:
    path = _tenants_csv(
        tmp_path / "elsewhere" / "customers.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
        f"{TENANT_A},Alpha,,Commercial,TRUE",
    )
    monkeypatch.setenv(ENV_PREFIX + "TENANTS_FILE", str(path))
    assert _main(recorder, "--group", "pilot") == 0
    assert _tenant_ids(recorder) == [TENANT_A]
    assert not (tmp_path / "config" / "tenants.csv").exists()


def _gamma_csv(tmp_path: Path, cloud: str = "Commercial", username: str = GAMMA_ADMIN) -> Path:
    """A tenants file whose pilot group holds Alpha and Gamma, which has no GDAP relationship."""
    return _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,m365_username,pilot",
        f"{TENANT_A},Alpha,,Commercial,,TRUE",
        f"{TENANT_C},Gamma,,{cloud},{username},TRUE",
    )


@pytest.mark.integration
@pytest.mark.acceptance
def test_a_group_member_without_gdap_gets_a_sign_in_of_its_own(
    fake_msal: FakeMsal, tmp_path: Path, recorder: ModuleType, caplog: pytest.LogCaptureFixture
) -> None:
    caplog.set_level(logging.INFO, logger="m365_gdap_runner.session")
    fake_msal.tenant_errors[TENANT_C] = NOT_IN_TENANT  # the partner account is not in Gamma
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return fake_services(request)

    path = _gamma_csv(tmp_path)
    assert _main(recorder, "--group", "pilot", "--tenants-file", str(path), handler=handler) == 0
    assert [(r.tenant.tenant_id, r.status) for r in _reported(recorder)] == [
        (TENANT_A, "ok"),
        (TENANT_C, "ok"),
    ]
    # One browser sign-in for the partner, one for Gamma as its own account; none for Alpha.
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_C]
    assert f"Signing in to Gamma directly as {GAMMA_ADMIN}" in caplog.text
    assert "AADSTS50020" in caplog.text
    users = [bearer_tenant(r) for r in seen if r.url.path == "/v1.0/users"]
    assert sorted(users) == [TENANT_A, TENANT_C]


@pytest.mark.integration
@pytest.mark.acceptance
def test_a_tenant_in_another_cloud_is_signed_in_to_there(
    fake_msal: FakeMsal, tmp_path: Path, recorder: ModuleType
) -> None:
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return fake_services(request)

    path = _gamma_csv(tmp_path, cloud="USGov")
    assert _main(recorder, "--tenant", "Gamma", "--tenants-file", str(path), handler=handler) == 0
    (gamma,) = _reported(recorder)
    assert (gamma.tenant.tenant_id, gamma.tenant.cloud, gamma.status) == (TENANT_C, "USGov", "ok")
    # Straight to the US Government sign-in: the commercial partner sign-in was never tried.
    assert fake_msal.silent_calls(TENANT_C) == []
    assert fake_msal.authority(TENANT_C) == f"https://login.microsoftonline.us/{TENANT_C}"
    (users,) = [r for r in seen if r.url.path == "/v1.0/users"]
    assert users.url.host == "graph.microsoft.us"


@pytest.mark.integration
@pytest.mark.acceptance
def test_a_gdap_tenant_missing_consent_gets_a_sign_in_of_its_own(
    fake_msal: FakeMsal, tmp_path: Path, recorder: ModuleType
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,m365_username",
        f"{TENANT_B},beta corp,,Commercial,admin@beta.example",
    )
    assert _main(recorder, "--all", "--tenants-file", str(path)) == 0
    assert [r.status for r in _reported(recorder)] == ["ok", "ok"]
    # The browser opens at beta corp itself, as its own account, where consent is given.
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT, TENANT_B]


@pytest.mark.integration
@pytest.mark.regression
def test_a_blank_username_fails_the_tenant_without_a_sign_in_as_the_partner_account(
    fake_msal: FakeMsal, tmp_path: Path, recorder: ModuleType
) -> None:
    fake_msal.tenant_errors[TENANT_C] = NOT_IN_TENANT  # the partner account is not in Gamma
    path = _gamma_csv(tmp_path, username="")
    assert _main(recorder, "--group", "pilot", "--tenants-file", str(path)) == 1
    alpha, gamma = _reported(recorder)
    assert (alpha.status, gamma.status) == ("ok", "auth_failed")
    assert gamma.error is not None
    assert "AADSTS50020" in gamma.error
    assert "No direct sign-in to Gamma: its m365_username is blank." in gamma.error
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]


@pytest.mark.integration
@pytest.mark.acceptance
@pytest.mark.parametrize(
    ("cloud", "reason"),
    [("Commercial", "AADSTS50020"), ("USGov", "Gamma is in the USGov cloud")],
)
def test_gdap_only_leaves_unreachable_tenants_failed(
    fake_msal: FakeMsal, tmp_path: Path, recorder: ModuleType, cloud: str, reason: str
) -> None:
    fake_msal.tenant_errors[TENANT_C] = NOT_IN_TENANT
    path = _gamma_csv(tmp_path, cloud=cloud)
    assert _main(recorder, "--group", "pilot", "--gdap-only", "--tenants-file", str(path)) == 1
    alpha, gamma = _reported(recorder)
    assert (alpha.status, gamma.status) == ("ok", "auth_failed")
    assert gamma.error is not None
    assert reason in gamma.error
    assert fake_msal.interactive_tenants() == [PARTNER_TENANT]


@pytest.mark.integration
def test_a_failed_direct_sign_in_fails_that_tenant_only(
    fake_msal: FakeMsal, tmp_path: Path, recorder: ModuleType
) -> None:
    fake_msal.tenant_errors[TENANT_C] = NOT_IN_TENANT
    fake_msal.interactive_errors[TENANT_C] = {"error": "access_denied", "error_description": "no"}
    path = _gamma_csv(tmp_path)
    assert _main(recorder, "--group", "pilot", "--tenants-file", str(path)) == 1
    assert [r.status for r in _reported(recorder)] == ["ok", "auth_failed"]


@pytest.mark.integration
def test_all_runs_gdap_tenants_only(
    fake_msal: FakeMsal, tmp_path: Path, recorder: ModuleType
) -> None:
    path = _gamma_csv(tmp_path)
    assert _main(recorder, "--all", "--tenants-file", str(path)) == 0
    assert _tenant_ids(recorder) == [TENANT_A, TENANT_B]


@pytest.mark.integration
def test_exit_2_when_the_groups_select_nothing(
    fake_msal: FakeMsal, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
        f"{TENANT_A},Alpha,,Commercial,FALSE",
    )
    assert _main(LICENSED_USERS, "--group", "pilot", "--tenants-file", str(path)) == 2
    assert "No tenant is marked TRUE in group(s) pilot." in caplog.text


@pytest.mark.integration
def test_run_fills_in_blank_clouds_by_discovery(
    fake_msal: FakeMsal, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    caplog.set_level(logging.INFO, logger=cli.__name__)
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,m365_username,pilot",
        f"{TENANT_A},Alpha,,,,TRUE",
        f"{TENANT_B},beta corp,,,,FALSE",
        f"{TENANT_C},Gamma,,,,FALSE",  # no GDAP relationship, but a tenant all the same
    )
    assert _main(LICENSED_USERS, "--group", "pilot", "--tenants-file", str(path)) == 0
    assert path.read_text(encoding="utf-8-sig").splitlines() == [
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,m365_username,pilot",
        f"{TENANT_A},Alpha,,Commercial,,TRUE",
        f"{TENANT_B},beta corp,,Commercial,,FALSE",
        f"{TENANT_C},Gamma,,Commercial,,FALSE",
    ]
    assert f"Filled in the cloud of 3 tenant(s) in {path}: Alpha, beta corp, Gamma" in caplog.text


@pytest.mark.integration
@pytest.mark.acceptance
def test_run_adds_the_columns_the_file_lacks(
    fake_msal: FakeMsal, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    caplog.set_level(logging.INFO, logger=cli.__name__)
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,pilot",
        f"{TENANT_A},Alpha,TRUE",
        f"{TENANT_B},beta corp,FALSE",
    )
    assert _main(LICENSED_USERS, "--group", "pilot", "--tenants-file", str(path)) == 0
    assert path.read_text(encoding="utf-8-sig").splitlines() == [
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,m365_username,pilot",
        f"{TENANT_A},Alpha,,Commercial,,TRUE",
        f"{TENANT_B},beta corp,,Commercial,,FALSE",
    ]
    assert f"Added column(s) m365_domain, m365_cloud, m365_username to {path}" in caplog.text


@pytest.mark.integration
@pytest.mark.parametrize(
    ("lines", "message"),
    [
        (
            ("m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud", f"{TENANT_A},Alpha,,"),
            "has no group column 'pilot'",
        ),
        (
            (
                "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
                f"{TENANT_A},Alpha,,,maybe",
            ),
            "row 2: 'maybe' in column 'pilot' is not TRUE, FALSE or blank",
        ),
        (
            (
                "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
                f"{TENANT_A},Alpha,,Mars,TRUE",
            ),
            "row 2: 'Mars' in column 'm365_cloud' is not a cloud",
        ),
    ],
)
def test_exit_2_on_a_bad_tenants_file(
    fake_msal: FakeMsal,
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
    lines: tuple[str, ...],
    message: str,
) -> None:
    path = _tenants_csv(tmp_path / "tenants.csv", *lines)
    assert _main(LICENSED_USERS, "--group", "pilot", "--tenants-file", str(path)) == 2
    assert message in caplog.text


@pytest.mark.integration
def test_unwritable_tenants_file_is_a_warning_not_an_error(
    fake_msal: FakeMsal,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
) -> None:
    def refuse(self: TenantFile) -> None:
        raise PermissionError("held open by a spreadsheet")

    monkeypatch.setattr(TenantFile, "save", refuse)
    assert _main(LICENSED_USERS, "--all") == 0
    assert "Could not write tenants file" in caplog.text
    assert "held open by a spreadsheet" in caplog.text
    assert not (tmp_path / "config" / "tenants.csv").exists()
