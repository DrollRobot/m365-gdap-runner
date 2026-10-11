"""Tests for the m365-gdap-runner command line in main.py.

End-to-end runs use the fake msal module and a fake service transport.
"""

from __future__ import annotations

import json
import logging
import os
from collections.abc import Callable
from pathlib import Path

import httpx
import pytest

from m365_gdap_runner import main as cli
from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.config.schema import ENV_PREFIX
from m365_gdap_runner.tenants import TenantFile
from tests._gdap_fakes import (
    CONSENT_REQUIRED,
    PARTNER_TENANT,
    TENANT_A,
    TENANT_B,
    TENANT_C,
    USER,
    FakeMsal,
    fake_services,
)

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"
EXAMPLE = str(EXAMPLES / "licensed_users.py")
AUDIT_EXAMPLE = str(EXAMPLES / "audit_disabled.py")

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


def _run(*argv: str, handler: Handler = fake_services) -> int:
    return cli.run(cli.parse_args(list(argv)), transport=httpx.MockTransport(handler))


def _write_payload(tmp_path: Path, body: str) -> str:
    path = tmp_path / "payload.py"
    path.write_text(body, encoding="utf-8")
    return str(path)


# --- argument parsing ------------------------------------------------------------------


@pytest.mark.unit
def test_parse_args_defaults() -> None:
    args = cli.parse_args([EXAMPLE, "--all"])
    assert (args.output, args.format, args.jobs, args.continue_on_error) == ("-", "json", 1, False)
    assert (args.tenant, args.group, args.exclude, args.all) == (None, None, None, True)
    assert (args.partner_tenant_id, args.username, args.client_id, args.cache, args.cloud) == (
        None,
        None,
        None,
        None,
        None,
    )
    assert args.tenants_file is None


@pytest.mark.unit
def test_parse_args_requires_payload_unless_listing_tenants() -> None:
    with pytest.raises(SystemExit) as info:
        cli.parse_args([])
    assert info.value.code == 2
    assert cli.parse_args(["--list-tenants"]).payload is None


@pytest.mark.unit
@pytest.mark.acceptance
def test_parse_args_requires_a_tenant_selection(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as info:
        cli.parse_args([EXAMPLE])
    assert info.value.code == 2
    assert "Choose tenants with --group, --tenant or --all" in capsys.readouterr().err
    for selection in (["--group", "pilot"], ["--tenant", "Alpha"], ["--all"]):
        assert cli.parse_args([EXAMPLE, *selection]).payload == EXAMPLE
    assert cli.parse_args(["--list-tenants"]).all is False


@pytest.mark.unit
@pytest.mark.parametrize("other", [["--tenant", "Alpha"], ["--group", "pilot"]])
def test_parse_args_rejects_all_with_a_narrower_selection(other: list[str]) -> None:
    with pytest.raises(SystemExit) as info:
        cli.parse_args([EXAMPLE, "--all", *other])
    assert info.value.code == 2


@pytest.mark.unit
def test_parse_args_rejects_jobs_below_one() -> None:
    with pytest.raises(SystemExit) as info:
        cli.parse_args([EXAMPLE, "--jobs", "0"])
    assert info.value.code == 2


@pytest.mark.unit
def test_parse_args_rejects_unknown_cache() -> None:
    with pytest.raises(SystemExit) as info:
        cli.parse_args([EXAMPLE, "--cache", "plaintext"])
    assert info.value.code == 2


@pytest.mark.unit
def test_main_returns_130_on_interrupt(monkeypatch: pytest.MonkeyPatch) -> None:
    def interrupted(args: object) -> int:
        raise KeyboardInterrupt

    monkeypatch.setattr(cli, "configure_logging", lambda verbosity: None)
    monkeypatch.setattr(cli, "run", interrupted)
    assert cli.main([EXAMPLE, "--all"]) == 130


@pytest.mark.unit
def test_configure_logging_quiets_libraries_unless_very_verbose() -> None:
    root = logging.getLogger()
    handlers = list(root.handlers)
    saved = {name: logging.getLogger(name).level for name in ("", *cli._LIBRARY_LOGGERS)}
    try:
        cli.configure_logging(0)
        assert root.level == logging.INFO
        assert logging.getLogger("azure_auth").level == logging.WARNING
        cli.configure_logging(2)
        assert root.level == logging.DEBUG
        assert logging.getLogger("msal").level == logging.DEBUG
    finally:
        for handler in root.handlers[len(handlers) :]:
            root.removeHandler(handler)
        for name, level in saved.items():
            logging.getLogger(name).setLevel(level)


@pytest.mark.integration
def test_settings_flags_override_config() -> None:
    args = cli.parse_args(
        [
            EXAMPLE,
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
        ]
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


# --- end to end ------------------------------------------------------------------------


@pytest.mark.integration
def test_runs_example_payload_to_json_on_stdout(
    fake_msal: FakeMsal, capsys: pytest.CaptureFixture[str]
) -> None:
    assert _run(EXAMPLE, "--all", "--jobs", "2") == 0
    captured = capsys.readouterr()
    records = json.loads(captured.out)
    assert [(r["tenant_id"], r["status"]) for r in records] == [(TENANT_A, "ok"), (TENANT_B, "ok")]
    assert records[1]["result"][0]["userPrincipalName"] == f"admin@{TENANT_B}.example"
    assert "Summary: ok=2, auth_failed=0, error=0, skipped=0" in captured.err
    assert fake_msal.interactive_calls == 1


@pytest.mark.integration
def test_runs_exchange_example(fake_msal: FakeMsal, capsys: pytest.CaptureFixture[str]) -> None:
    assert _run(AUDIT_EXAMPLE, "--tenant", "beta corp") == 0
    records = json.loads(capsys.readouterr().out)
    assert [(r["status"], r["result"]["auditDisabled"]) for r in records] == [("ok", True)]
    assert fake_msal.interactive_calls == 2


@pytest.mark.integration
def test_writes_csv_file(fake_msal: FakeMsal, tmp_path: Path) -> None:
    output = tmp_path / "users.csv"
    assert _run(EXAMPLE, "--all", "--format", "csv", "--output", str(output)) == 0
    assert output.read_text(encoding="utf-8").splitlines() == [
        "tenant_id,display_name,userPrincipalName,displayName,skuIds",
        f"{TENANT_A},Alpha,admin@{TENANT_A}.example,Admin,sku-1;sku-2",
        f"{TENANT_B},beta corp,admin@{TENANT_B}.example,Admin,sku-1;sku-2",
    ]


@pytest.mark.integration
def test_exit_1_when_a_tenant_refuses_consent(
    fake_msal: FakeMsal, capsys: pytest.CaptureFixture[str]
) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    assert _run(EXAMPLE, "--all") == 1
    records = json.loads(capsys.readouterr().out)
    assert [r["status"] for r in records] == ["ok", "auth_failed"]


@pytest.mark.integration
def test_tenant_filters(fake_msal: FakeMsal, capsys: pytest.CaptureFixture[str]) -> None:
    assert _run(EXAMPLE, "--all", "--exclude", "BETA CORP") == 0
    assert [r["tenant_id"] for r in json.loads(capsys.readouterr().out)] == [TENANT_A]


@pytest.mark.integration
def test_list_tenants(fake_msal: FakeMsal, capsys: pytest.CaptureFixture[str]) -> None:
    assert _run("--list-tenants") == 0
    assert capsys.readouterr().out.splitlines() == [f"{TENANT_A}\tAlpha", f"{TENANT_B}\tbeta corp"]
    assert fake_msal.interactive_calls == 1


@pytest.mark.integration
def test_exit_2_on_unknown_tenant_before_running_payload(fake_msal: FakeMsal) -> None:
    seen: list[str] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request.url.path)
        return fake_services(request)

    assert _run(EXAMPLE, "--tenant", "Contoso", handler=handler) == 2
    assert "/v1.0/users" not in seen


@pytest.mark.integration
def test_exit_2_on_invalid_payload_before_sign_in(fake_msal: FakeMsal, tmp_path: Path) -> None:
    payload = _write_payload(tmp_path, "def run(ctx):\n    return None\n")
    assert _run(payload, "--all") == 2
    assert fake_msal.interactive_calls == 0


@pytest.mark.integration
def test_exit_2_on_missing_config(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    monkeypatch.delenv(ENV_PREFIX + "USERNAME")
    assert _run(EXAMPLE, "--all") == 2
    assert "username" in caplog.text


@pytest.mark.integration
def test_exit_2_on_failed_sign_in(fake_msal: FakeMsal) -> None:
    fake_msal.interactive_result = {"error": "access_denied", "error_description": "cancelled"}
    assert _run(EXAMPLE, "--all") == 2


@pytest.mark.integration
def test_exit_2_on_undeclared_client_is_a_tenant_error(
    fake_msal: FakeMsal, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    payload = _write_payload(
        tmp_path,
        'SCOPES = ["User.Read.All"]\n\ndef run(ctx):\n    return ctx.exchange.run("Get-Mailbox")\n',
    )
    assert _run(payload, "--tenant", "Alpha") == 1
    (record,) = json.loads(capsys.readouterr().out)
    assert record["status"] == "error"
    assert "ClientNotDeclared" in record["error"]


@pytest.mark.integration
def test_interrupt_writes_partial_output(fake_msal: FakeMsal, tmp_path: Path) -> None:
    payload = _write_payload(
        tmp_path,
        'SCOPES = ["User.Read.All"]\n\n'
        "def run(ctx):\n"
        '    if ctx.tenant.display_name == "beta corp":\n'
        "        raise KeyboardInterrupt\n"
        "    return {'tenant': ctx.tenant.display_name}\n",
    )
    output = tmp_path / "out.json"
    assert _run(payload, "--all", "--output", str(output)) == 130
    records = json.loads(output.read_text(encoding="utf-8"))
    assert [(r["status"], r["error"]) for r in records] == [("ok", None), ("error", "interrupted")]


# --- tenants file ----------------------------------------------------------------------


def _tenants_csv(path: Path, *lines: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


@pytest.mark.integration
@pytest.mark.acceptance
def test_list_tenants_creates_the_tenants_file(
    fake_msal: FakeMsal, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    caplog.set_level(logging.INFO, logger=cli.__name__)
    assert _run("--list-tenants") == 0
    path = tmp_path / "config" / "tenants.csv"
    assert path.read_text(encoding="utf-8-sig").splitlines() == [
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud",
        f"{TENANT_A},Alpha,,Commercial",
        f"{TENANT_B},beta corp,,Commercial",
    ]
    assert f"Created tenants file {path} with 2 tenant(s)" in caplog.text


@pytest.mark.integration
@pytest.mark.acceptance
def test_run_appends_tenants_the_file_lacks_and_warns(
    fake_msal: FakeMsal, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot,notes",
        f"{TENANT_A},Alpha,alpha.example,Commercial,TRUE,keep me",
    )
    assert _run(EXAMPLE, "--tenant", "Alpha", "--tenants-file", str(path)) == 0
    assert path.read_text(encoding="utf-8-sig").splitlines() == [
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot,notes",
        f"{TENANT_A},Alpha,alpha.example,Commercial,TRUE,keep me",
        f"{TENANT_B},beta corp,,Commercial,,",
    ]
    assert f"Added 1 tenant(s) to {path}; they are in no group yet: beta corp" in caplog.text


@pytest.mark.integration
@pytest.mark.acceptance
def test_group_runs_only_the_marked_tenants(
    fake_msal: FakeMsal, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
        f"{TENANT_A},Alpha,,Commercial,FALSE",
        f"{TENANT_B},beta corp,,Commercial,TRUE",
        f"{TENANT_C},Gamma,,,FALSE",  # no GDAP relationship: harmless while unselected
    )
    assert _run(EXAMPLE, "--group", "pilot", "--tenants-file", str(path)) == 0
    assert [r["tenant_id"] for r in json.loads(capsys.readouterr().out)] == [TENANT_B]


@pytest.mark.integration
def test_groups_and_tenants_add_up_and_exclude_applies(
    fake_msal: FakeMsal, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot,patching",
        f"{TENANT_A},Alpha,,Commercial,TRUE,FALSE",
        f"{TENANT_B},beta corp,,Commercial,FALSE,TRUE",
    )
    file_flag = ["--tenants-file", str(path)]
    assert _run(EXAMPLE, "--group", "pilot", "--group", "PATCHING", *file_flag) == 0
    assert [r["tenant_id"] for r in json.loads(capsys.readouterr().out)] == [TENANT_A, TENANT_B]
    assert _run(EXAMPLE, "--group", "pilot", "--tenant", "beta corp", *file_flag) == 0
    assert [r["tenant_id"] for r in json.loads(capsys.readouterr().out)] == [TENANT_A, TENANT_B]
    assert _run(EXAMPLE, "--group", "pilot", "--exclude", "Alpha", *file_flag) == 0
    assert json.loads(capsys.readouterr().out) == []


@pytest.mark.integration
def test_tenants_file_setting_locates_the_file(
    fake_msal: FakeMsal,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    path = _tenants_csv(
        tmp_path / "elsewhere" / "customers.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
        f"{TENANT_A},Alpha,,Commercial,TRUE",
    )
    monkeypatch.setenv(ENV_PREFIX + "TENANTS_FILE", str(path))
    assert _run(EXAMPLE, "--group", "pilot") == 0
    assert [r["tenant_id"] for r in json.loads(capsys.readouterr().out)] == [TENANT_A]
    assert not (tmp_path / "config" / "tenants.csv").exists()


@pytest.mark.integration
@pytest.mark.acceptance
def test_exit_2_when_a_group_member_has_no_gdap_relationship(
    fake_msal: FakeMsal, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
        f"{TENANT_A},Alpha,,Commercial,TRUE",
        f"{TENANT_C},Gamma,,,TRUE",
    )
    seen: list[str] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request.url.path)
        return fake_services(request)

    assert _run(EXAMPLE, "--group", "pilot", "--tenants-file", str(path), handler=handler) == 2
    assert f"'Gamma' ({TENANT_C}) is in group 'pilot' but has no GDAP relationship" in caplog.text
    assert "/v1.0/users" not in seen


@pytest.mark.integration
def test_exit_2_when_the_groups_select_nothing(
    fake_msal: FakeMsal, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
        f"{TENANT_A},Alpha,,Commercial,FALSE",
    )
    assert _run(EXAMPLE, "--group", "pilot", "--tenants-file", str(path)) == 2
    assert "No tenant is marked TRUE in group(s) pilot." in caplog.text


@pytest.mark.integration
@pytest.mark.acceptance
def test_group_skips_members_in_another_cloud(
    fake_msal: FakeMsal,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    caplog: pytest.LogCaptureFixture,
) -> None:
    caplog.set_level(logging.INFO, logger="m365_gdap_runner.tenants")
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
        f"{TENANT_A},Alpha,,Commercial,TRUE",
        f"{TENANT_C},Gamma,,USGov,TRUE",  # another cloud: out of this sign-in's reach, no error
    )
    assert _run(EXAMPLE, "--group", "pilot", "--tenants-file", str(path)) == 0
    assert [r["tenant_id"] for r in json.loads(capsys.readouterr().out)] == [TENANT_A]
    assert "Skipped 1 tenant(s) outside the Commercial cloud: Gamma (USGov)" in caplog.text


@pytest.mark.integration
def test_run_fills_in_blank_clouds_by_discovery(
    fake_msal: FakeMsal, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    caplog.set_level(logging.INFO, logger=cli.__name__)
    path = _tenants_csv(
        tmp_path / "tenants.csv",
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
        f"{TENANT_A},Alpha,,,TRUE",
        f"{TENANT_B},beta corp,,,FALSE",
    )
    assert _run(EXAMPLE, "--group", "pilot", "--tenants-file", str(path)) == 0
    assert path.read_text(encoding="utf-8-sig").splitlines() == [
        "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot",
        f"{TENANT_A},Alpha,,Commercial,TRUE",
        f"{TENANT_B},beta corp,,Commercial,FALSE",
    ]
    assert f"Filled in the cloud of 2 tenant(s) in {path}: Alpha, beta corp" in caplog.text


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
        (
            ("m365_tenant_id,pilot", f"{TENANT_A},TRUE"),
            "lacks the column(s) m365_tenant_name, m365_domain, m365_cloud",
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
    assert _run(EXAMPLE, "--group", "pilot", "--tenants-file", str(path)) == 2
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
    assert _run(EXAMPLE, "--all") == 0
    assert "Could not write tenants file" in caplog.text
    assert "held open by a spreadsheet" in caplog.text
    assert not (tmp_path / "config" / "tenants.csv").exists()
