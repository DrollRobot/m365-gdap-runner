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
from tests._gdap_fakes import (
    CONSENT_REQUIRED,
    PARTNER_TENANT,
    TENANT_A,
    TENANT_B,
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
    args = cli.parse_args([EXAMPLE])
    assert (args.output, args.format, args.jobs, args.continue_on_error) == ("-", "json", 1, False)
    assert args.tenant is None
    assert (args.partner_tenant_id, args.username, args.client_id, args.cache, args.cloud) == (
        None,
        None,
        None,
        None,
        None,
    )


@pytest.mark.unit
def test_parse_args_requires_payload_unless_listing_tenants() -> None:
    with pytest.raises(SystemExit) as info:
        cli.parse_args([])
    assert info.value.code == 2
    assert cli.parse_args(["--list-tenants"]).payload is None


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
    assert cli.main([EXAMPLE]) == 130


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
        ]
    )
    settings = cli.settings_from_args(args)
    assert (
        settings.partner_tenant_id,
        settings.username,
        settings.client_id,
        settings.cache,
        settings.cloud,
    ) == ("other.example", "someone@other.example", "custom", "disk", "USGov")


# --- end to end ------------------------------------------------------------------------


@pytest.mark.integration
def test_runs_example_payload_to_json_on_stdout(
    fake_msal: FakeMsal, capsys: pytest.CaptureFixture[str]
) -> None:
    assert _run(EXAMPLE, "--jobs", "2") == 0
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
    assert _run(EXAMPLE, "--format", "csv", "--output", str(output)) == 0
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
    assert _run(EXAMPLE) == 1
    records = json.loads(capsys.readouterr().out)
    assert [r["status"] for r in records] == ["ok", "auth_failed"]


@pytest.mark.integration
def test_tenant_filters(fake_msal: FakeMsal, capsys: pytest.CaptureFixture[str]) -> None:
    assert _run(EXAMPLE, "--exclude", "BETA CORP") == 0
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
    assert _run(payload) == 2
    assert fake_msal.interactive_calls == 0


@pytest.mark.integration
def test_exit_2_on_missing_config(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    monkeypatch.delenv(ENV_PREFIX + "USERNAME")
    assert _run(EXAMPLE) == 2
    assert "username" in caplog.text


@pytest.mark.integration
def test_exit_2_on_failed_sign_in(fake_msal: FakeMsal) -> None:
    fake_msal.interactive_result = {"error": "access_denied", "error_description": "cancelled"}
    assert _run(EXAMPLE) == 2


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
    assert _run(payload, "--output", str(output)) == 130
    records = json.loads(output.read_text(encoding="utf-8"))
    assert [(r["status"], r["error"]) for r in records] == [("ok", None), ("error", "interrupted")]
