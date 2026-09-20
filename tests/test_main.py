"""Tests for the m365-gdap-runner command line in main.py.

End-to-end runs use the fake MSAL factory and a fake Graph transport.
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
    FakeMsal,
    fake_graph,
)

EXAMPLE = str(Path(__file__).resolve().parents[1] / "examples" / "licensed_users.py")

Handler = Callable[[httpx.Request], httpx.Response]


@pytest.fixture(autouse=True)
def _config(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Isolate config and supply the partner tenant through the environment."""
    for key in list(os.environ):
        if key.startswith(ENV_PREFIX):
            monkeypatch.delenv(key)
    monkeypatch.setenv(CONFIG_DIR_ENV, str(tmp_path / "config"))
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", PARTNER_TENANT)


def _run(*argv: str, fake: FakeMsal | None = None, handler: Handler = fake_graph) -> int:
    return cli.run(
        cli.parse_args(list(argv)),
        app_factory=fake or FakeMsal(),
        transport=httpx.MockTransport(handler),
    )


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
    assert args.persist_token_cache is None


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
        assert logging.getLogger("httpx").level == logging.WARNING
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
        [EXAMPLE, "--partner-tenant-id", "other", "--client-id", "custom", "--persist-token-cache"]
    )
    settings = cli.settings_from_args(args)
    assert (settings.partner_tenant_id, settings.client_id, settings.persist_token_cache) == (
        "other",
        "custom",
        True,
    )


# --- end to end ------------------------------------------------------------------------


@pytest.mark.integration
def test_runs_example_payload_to_json_on_stdout(capsys: pytest.CaptureFixture[str]) -> None:
    assert _run(EXAMPLE, "--jobs", "2") == 0
    captured = capsys.readouterr()
    records = json.loads(captured.out)
    assert [(r["tenant_id"], r["status"]) for r in records] == [(TENANT_A, "ok"), (TENANT_B, "ok")]
    assert records[1]["result"][0]["userPrincipalName"] == f"admin@{TENANT_B}.example"
    assert "Summary: ok=2, auth_failed=0, error=0, skipped=0" in captured.err


@pytest.mark.integration
def test_writes_csv_file(tmp_path: Path) -> None:
    output = tmp_path / "users.csv"
    assert _run(EXAMPLE, "--format", "csv", "--output", str(output)) == 0
    assert output.read_text(encoding="utf-8").splitlines() == [
        "tenant_id,display_name,userPrincipalName,displayName,skuIds",
        f"{TENANT_A},Alpha,admin@{TENANT_A}.example,Admin,sku-1;sku-2",
        f"{TENANT_B},beta corp,admin@{TENANT_B}.example,Admin,sku-1;sku-2",
    ]


@pytest.mark.integration
def test_exit_1_when_a_tenant_refuses_consent(capsys: pytest.CaptureFixture[str]) -> None:
    fake = FakeMsal(tenant_errors={TENANT_B: CONSENT_REQUIRED})
    assert _run(EXAMPLE, fake=fake) == 1
    records = json.loads(capsys.readouterr().out)
    assert [r["status"] for r in records] == ["ok", "auth_failed"]


@pytest.mark.integration
def test_tenant_filters(capsys: pytest.CaptureFixture[str]) -> None:
    assert _run(EXAMPLE, "--exclude", "BETA CORP") == 0
    assert [r["tenant_id"] for r in json.loads(capsys.readouterr().out)] == [TENANT_A]


@pytest.mark.integration
def test_list_tenants(capsys: pytest.CaptureFixture[str]) -> None:
    assert _run("--list-tenants") == 0
    assert capsys.readouterr().out.splitlines() == [f"{TENANT_A}\tAlpha", f"{TENANT_B}\tbeta corp"]


@pytest.mark.integration
def test_exit_2_on_unknown_tenant_before_running_payload() -> None:
    seen: list[str] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request.url.path)
        return fake_graph(request)

    assert _run(EXAMPLE, "--tenant", "Contoso", handler=handler) == 2
    assert "/v1.0/users" not in seen


@pytest.mark.integration
def test_exit_2_on_invalid_payload_before_sign_in(tmp_path: Path) -> None:
    fake = FakeMsal()
    payload = _write_payload(tmp_path, "def run(ctx):\n    return None\n")
    assert _run(payload, fake=fake) == 2
    assert fake.interactive_calls == 0


@pytest.mark.integration
def test_exit_2_on_missing_config(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    monkeypatch.delenv(ENV_PREFIX + "PARTNER_TENANT_ID")
    assert _run(EXAMPLE) == 2
    assert "partner_tenant_id" in caplog.text


@pytest.mark.integration
def test_exit_2_on_failed_sign_in() -> None:
    fake = FakeMsal(interactive_result={"error": "access_denied", "error_description": "cancelled"})
    assert _run(EXAMPLE, fake=fake) == 2


@pytest.mark.integration
def test_interrupt_writes_partial_output(tmp_path: Path) -> None:
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
