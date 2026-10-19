"""Integration tests: runner with the real session, azure-auth clients and the examples."""

from __future__ import annotations

import argparse
import datetime
import json
import logging
import os
from pathlib import Path
from typing import Any

import httpx
import pytest
from azure_auth.constants import EXCHANGE_POWERSHELL_CLIENT_ID, GRAPH_CLI_CLIENT_ID
from openpyxl import load_workbook

from m365_gdap_runner import cli
from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.config.schema import ENV_PREFIX
from m365_gdap_runner.loader import Payload, payload_from_module
from m365_gdap_runner.runner import TenantResult, run_payload
from m365_gdap_runner.session import PartnerSession
from m365_gdap_runner.tenants import Tenant
from tests._gdap_fakes import (
    CONSENT_REQUIRED,
    PARTNER_TENANT,
    TENANT_A,
    TENANT_B,
    USER,
    FakeMsal,
    bearer_tenant,
    fake_services,
    load_script,
    make_settings,
)

pytestmark = pytest.mark.integration

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"
LICENSED_USERS = EXAMPLES / "licensed_users.py"
AUDIT_DISABLED = EXAMPLES / "audit_disabled.py"
MAILBOX_AUDIT_GAPS = EXAMPLES / "mailbox_audit_gaps.py"
UNIFIED_AUDIT_LOG = EXAMPLES / "unified_audit_log.py"
SERVICE_PRINCIPAL_ENABLED = EXAMPLES / "service_principal_enabled.py"


def _example(path: Path) -> Payload:
    return payload_from_module(load_script(path), path.stem)


@pytest.fixture(autouse=True)
def _isolated_config(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Strip this package's env vars and point the config dir at a temp dir."""
    for key in list(os.environ):
        if key.startswith(ENV_PREFIX):
            monkeypatch.delenv(key)
    monkeypatch.setenv(CONFIG_DIR_ENV, str(tmp_path))


def test_graph_example_runs_through_session_and_azure_auth(fake_msal: FakeMsal) -> None:
    fake_msal.tenant_errors[TENANT_B] = CONSENT_REQUIRED
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return fake_services(request)

    payload = _example(LICENSED_USERS)
    session = PartnerSession(make_settings(), transport=httpx.MockTransport(handler))
    session.login(payload.clients)
    tenants = session.list_tenants()
    ready = session.check_access(tenants, scopes=payload.scopes, clients=payload.clients)
    results = run_payload(session, payload, tenants, ready=ready)

    assert [t.tenant_id for t in tenants] == [TENANT_A, TENANT_B]
    assert [r.status for r in results] == ["ok", "auth_failed"]
    assert results[0].result == [
        {
            "userPrincipalName": f"admin@{TENANT_A}.example",
            "displayName": "Admin",
            "skuIds": "sku-1;sku-2",
        }
    ]
    assert results[1].error is not None
    # Consent was checked against the script's own scopes: the refusal names them.
    assert "User.Read.All cannot be issued" in results[1].error
    # The script never ran for the tenant that refused the token.
    users_requests = [r for r in seen if r.url.path == "/v1.0/users"]
    assert [bearer_tenant(r) for r in users_requests] == [TENANT_A]
    assert users_requests[0].headers["ConsistencyLevel"] == "eventual"
    assert fake_msal.interactive_calls == 1


def test_exchange_example_signs_in_to_exchange_and_runs_the_cmdlet(
    fake_msal: FakeMsal, caplog: pytest.LogCaptureFixture
) -> None:
    caplog.set_level(logging.INFO)
    payload = _example(AUDIT_DISABLED)
    assert (payload.clients, payload.scopes, payload.report) == (("exchange",), (), None)

    session = PartnerSession(make_settings(), transport=httpx.MockTransport(fake_services))
    session.login(payload.clients)
    tenants = session.list_tenants()
    ready = session.check_access(tenants, clients=payload.clients)
    results = run_payload(session, payload, tenants, ready=ready)

    assert [r.status for r in results] == ["ok", "ok"]
    assert [r.getMessage() for r in caplog.records if r.filename == "audit_disabled.py"] == [
        "Mailbox auditing is on",
        "Mailbox auditing is off tenant-wide (AuditDisabled is True)",
    ]
    # Graph for the tenant list, Exchange for the cmdlet: two home sign-ins.
    assert [c.client_id for c in fake_msal.calls if c.method == "interactive"] == [
        GRAPH_CLI_CLIENT_ID,
        EXCHANGE_POWERSHELL_CLIENT_ID,
    ]
    # Each customer was checked with the Exchange client id before the cmdlet ran.
    assert fake_msal.silent_calls(TENANT_A)[0] == (
        EXCHANGE_POWERSHELL_CLIENT_ID,
        ("https://outlook.office365.com/.default",),
    )


# --- mailbox_audit_gaps.py -------------------------------------------------------------


def _mailbox(upn: str, kind: str, defaults: list[str], **actions: Any) -> dict[str, Any]:
    """Build one mailbox as the Audit read returns it; actions default to the tenant's own."""
    sets = {"Admin": ["Update", "Move"], "Delegate": ["SendAs"], "Owner": ["MailboxLogin"]}
    return {
        "UserPrincipalName": upn,
        "RecipientTypeDetails": kind,
        "DefaultAuditSet": defaults,
        **{f"Audit{logon_type}": actions.get(logon_type, sets[logon_type]) for logon_type in sets},
    }


def _audit_services(request: httpx.Request) -> httpx.Response:
    """Fake Exchange for the audit example: tenant A is clean, tenant B has every finding."""
    tenant = bearer_tenant(request)
    gaps = tenant == TENANT_B
    if request.url.path == f"/adminapi/beta/{tenant}/Mailbox":
        mailboxes = [_mailbox(f"enrolled@{tenant}", "UserMailbox", ["Admin", "Delegate", "Owner"])]
        if gaps:
            mailboxes += [
                _mailbox("Zed@b", "SharedMailbox", ["Admin", "Delegate"], Owner=[]),
                _mailbox("cee@b", "UserMailbox", [], Admin=["Update"], Owner="MailboxLogin"),
            ]
        return httpx.Response(200, json={"value": mailboxes})
    if request.url.path != f"/adminapi/beta/{tenant}/InvokeCommand":
        return fake_services(request)
    cmdlet = json.loads(request.content)["CmdletInput"]["CmdletName"]
    answers: dict[str, list[dict[str, Any]]] = {
        "Get-AcceptedDomain": [
            {"DomainName": "contoso.example", "InitialDomain": False},
            {"DomainName": f"{tenant}.onmicrosoft.com", "InitialDomain": True},
        ],
        "Get-OrganizationConfig": [{"AuditDisabled": gaps}],
        "Get-AdminAuditLogConfig": [{"UnifiedAuditLogIngestionEnabled": not gaps}],
        "Get-MailboxAuditBypassAssociation": [
            {"Name": "svc-backup", "Identity": "svc-backup", "AuditBypassEnabled": gaps},
            {"Name": "svc-other", "Identity": "svc-other", "AuditBypassEnabled": False},
        ],
    }
    return httpx.Response(200, json={"value": answers[cmdlet]})


def _sheets(path: Path) -> dict[str, list[tuple[Any, ...]]]:
    """Read every sheet of a workbook as rows of values, the header row first."""
    workbook = load_workbook(path)
    return {sheet.title: list(sheet.iter_rows(values_only=True)) for sheet in workbook.worksheets}


@pytest.fixture
def mailbox_audit_gaps(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Any:
    """The example, writing its workbooks to a temp dir, as ``--out`` makes it."""
    module = load_script(MAILBOX_AUDIT_GAPS)
    monkeypatch.setattr(module, "OUT_DIR", tmp_path / "reports")
    return module


def test_mailbox_audit_example_writes_a_workbook_per_tenant_and_a_summary(
    fake_msal: FakeMsal,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    mailbox_audit_gaps: Any,
    capsys: pytest.CaptureFixture[str],
) -> None:
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", PARTNER_TENANT)
    monkeypatch.setenv(ENV_PREFIX + "USERNAME", USER)
    monkeypatch.setattr(cli, "configure_logging", lambda verbosity: None)
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return _audit_services(request)

    parser = argparse.ArgumentParser()
    cli.add_arguments(parser)
    args = parser.parse_args(["--include", "all", "--jobs", "2"])
    assert cli.main(mailbox_audit_gaps, args, transport=httpx.MockTransport(handler)) == 0

    today = datetime.date.today().isoformat()
    reports = tmp_path / "reports"
    assert sorted(p.name for p in reports.iterdir()) == [
        f"MailboxAuditGaps_{TENANT_A}_{today}.xlsx",
        f"MailboxAuditGaps_{TENANT_B}_{today}.xlsx",
        f"MailboxAuditGaps_summary_{today}.xlsx",
    ]
    summary_path = reports / f"MailboxAuditGaps_summary_{today}.xlsx"
    assert capsys.readouterr().out == f"Wrote {summary_path}\n"

    clean = _sheets(reports / f"MailboxAuditGaps_{TENANT_A}_{today}.xlsx")
    assert list(clean) == ["Tenant"]
    header, settings = clean["Tenant"]
    assert header == ("Tenant", "Collected", "AuditDisabled", "UnifiedAuditLogIngestionEnabled")
    assert (settings[0], settings[2], settings[3]) == (f"{TENANT_A}.onmicrosoft.com", False, True)
    assert isinstance(settings[1], datetime.datetime)

    findings = _sheets(reports / f"MailboxAuditGaps_{TENANT_B}_{today}.xlsx")
    assert list(findings) == ["Tenant", "Audit Gaps", "Audit Bypass"]
    assert findings["Audit Gaps"] == [
        ("UserPrincipalName", "RecipientTypeDetails", "MissingActions"),
        ("cee@b", "UserMailbox", "Move"),
        ("Zed@b", "SharedMailbox", "MailboxLogin"),
    ]
    assert findings["Audit Bypass"] == [("Name", "Identity"), ("svc-backup", "svc-backup")]

    summary = _sheets(summary_path)
    assert summary["Tenants"] == [
        (
            "Tenant",
            "TenantId",
            "Status",
            "AuditDisabled",
            "UnifiedAuditLogIngestionEnabled",
            "Mailboxes",
            "WithGaps",
            "Bypassed",
            "Workbook",
            "Error",
        ),
        (
            "Alpha",
            TENANT_A,
            "ok",
            False,
            True,
            1,
            0,
            0,
            f"MailboxAuditGaps_{TENANT_A}_{today}.xlsx",
            None,
        ),
        (
            "beta corp",
            TENANT_B,
            "ok",
            True,
            False,
            3,
            2,
            1,
            f"MailboxAuditGaps_{TENANT_B}_{today}.xlsx",
            None,
        ),
    ]
    assert summary["Audit Gaps"] == [
        ("Tenant", "UserPrincipalName", "RecipientTypeDetails", "MissingActions"),
        ("beta corp", "cee@b", "UserMailbox", "Move"),
        ("beta corp", "Zed@b", "SharedMailbox", "MailboxLogin"),
    ]

    # Get-EXOMailbox -PropertySets Audit -RecipientTypeDetails ... -Properties ...
    (read,) = [r for r in seen if r.url.path == f"/adminapi/beta/{TENANT_B}/Mailbox"]
    assert read.method == "GET"
    assert dict(read.url.params) == {
        "PropertySet": "Audit",
        "RecipientTypeDetails": "UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox",
        "$select": "DefaultAuditSet,UserPrincipalName,RecipientTypeDetails",
    }


def test_mailbox_audit_summary_lists_the_tenants_that_failed(
    tmp_path: Path, mailbox_audit_gaps: Any
) -> None:
    failed = TenantResult(Tenant(TENANT_A, "Alpha"), "auth_failed", error="AADSTS65001")
    mailbox_audit_gaps.report([failed])
    (path,) = (tmp_path / "reports").iterdir()
    workbook = load_workbook(path)
    sheet = workbook["Tenants"]
    assert list(sheet.iter_rows(min_row=2, values_only=True)) == [
        ("Alpha", TENANT_A, "auth_failed", None, None, None, None, None, None, "AADSTS65001")
    ]
    assert workbook.sheetnames == ["Tenants"]  # no gaps from a tenant that never ran
    assert sheet.freeze_panes == "A2"
    (table,) = sheet.tables.values()
    assert table.tableStyleInfo is not None
    assert table.tableStyleInfo.name == "TableStyleMedium8"


# --- unified_audit_log.py --------------------------------------------------------------


def _ual_services(request: httpx.Request) -> httpx.Response:
    """Fake Exchange for the unified audit log example: on in tenant A, off in tenant B."""
    tenant = bearer_tenant(request)
    if request.url.path != f"/adminapi/beta/{tenant}/InvokeCommand":
        return fake_services(request)
    cmdlet = json.loads(request.content)["CmdletInput"]["CmdletName"]
    assert cmdlet == "Get-AdminAuditLogConfig"
    config = {"UnifiedAuditLogIngestionEnabled": tenant == TENANT_A}
    return httpx.Response(200, json={"value": [config]})


@pytest.fixture
def unified_audit_log(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Any:
    """The example, writing its workbook to a temp dir, as ``--out`` makes it."""
    module = load_script(UNIFIED_AUDIT_LOG)
    monkeypatch.setattr(module, "OUT_DIR", tmp_path / "reports")
    return module


def test_unified_audit_log_example_writes_one_workbook_for_every_tenant(
    fake_msal: FakeMsal,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    unified_audit_log: Any,
    capsys: pytest.CaptureFixture[str],
    caplog: pytest.LogCaptureFixture,
) -> None:
    caplog.set_level(logging.INFO)
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", PARTNER_TENANT)
    monkeypatch.setenv(ENV_PREFIX + "USERNAME", USER)
    monkeypatch.setattr(cli, "configure_logging", lambda verbosity: None)

    parser = argparse.ArgumentParser()
    cli.add_arguments(parser)
    args = parser.parse_args(["--include", "all"])
    assert cli.main(unified_audit_log, args, transport=httpx.MockTransport(_ual_services)) == 0

    (path,) = (tmp_path / "reports").iterdir()
    assert path.name == f"UnifiedAuditLog_{datetime.date.today().isoformat()}.xlsx"
    assert capsys.readouterr().out == f"Wrote {path}\n"
    assert _sheets(path) == {
        "Tenants": [
            ("Tenant", "TenantId", "Status", "UnifiedAuditLogIngestionEnabled", "Error"),
            ("Alpha", TENANT_A, "ok", True, None),
            ("beta corp", TENANT_B, "ok", False, None),
        ]
    }
    assert [
        (r.levelname, r.getMessage())
        for r in caplog.records
        if r.filename == UNIFIED_AUDIT_LOG.name
    ] == [
        ("INFO", "Unified audit log is on"),
        ("WARNING", "Unified audit log is off (UnifiedAuditLogIngestionEnabled is False)"),
    ]


def test_unified_audit_log_workbook_lists_the_tenants_that_failed(
    tmp_path: Path, unified_audit_log: Any
) -> None:
    unified_audit_log.report(
        [
            TenantResult(Tenant(TENANT_A, "Alpha"), "ok", False),
            TenantResult(Tenant(TENANT_B, "beta corp"), "auth_failed", error="AADSTS65001"),
        ]
    )
    (path,) = (tmp_path / "reports").iterdir()
    sheet = load_workbook(path)["Tenants"]
    assert list(sheet.iter_rows(min_row=2, values_only=True)) == [
        ("Alpha", TENANT_A, "ok", False, None),
        ("beta corp", TENANT_B, "auth_failed", None, "AADSTS65001"),
    ]
    assert sheet.freeze_panes == "A2"
    (table,) = sheet.tables.values()
    assert table.ref == "A1:E3"
    assert table.tableStyleInfo is not None
    assert table.tableStyleInfo.name == "TableStyleMedium8"


def test_unified_audit_log_writes_no_workbook_when_no_tenant_ran(
    tmp_path: Path, unified_audit_log: Any, capsys: pytest.CaptureFixture[str]
) -> None:
    unified_audit_log.report([])
    assert not (tmp_path / "reports").exists()
    assert capsys.readouterr().out == "No tenants ran; no workbook written.\n"


# --- service_principal_enabled.py --------------------------------------------------------

BACKUP_APP_ID = "cw000000-0000-0000-0000-000000000001"
REPORTER_APP_ID = "rp000000-0000-0000-0000-000000000002"
MAIL_APP_ID = "ml000000-0000-0000-0000-000000000003"


def _sp(object_id: str, name: str, app_id: str, *, enabled: bool) -> dict[str, Any]:
    """Build a service principal as the list read returns it."""
    return {"id": object_id, "appId": app_id, "displayName": name, "accountEnabled": enabled}


def _sp_services(
    apps: dict[str, list[dict[str, Any]]], patches: list[httpx.Request]
) -> httpx.MockTransport:
    """Fake Graph holding *apps* per tenant, recording every PATCH it is sent.

    A read of one service principal returns it whole: as listed, plus ``tags``, which the list
    read leaves out. A PATCH changes nothing, so a later read still shows the old state.
    """

    def handler(request: httpx.Request) -> httpx.Response:
        held = apps.get(bearer_tenant(request), [])
        path = request.url.path
        if path == "/v1.0/servicePrincipals":
            assert request.url.params["$select"] == "id,appId,displayName,accountEnabled"
            assert request.url.params["$top"] == "999"
            return httpx.Response(200, json={"value": held})
        if path.startswith("/v1.0/servicePrincipals/"):
            if request.method == "PATCH":
                patches.append(request)
                return httpx.Response(204)
            (app,) = [app for app in held if app["id"] == path.rsplit("/", 1)[1]]
            return httpx.Response(200, json={**app, "tags": ["WindowsAzureActiveDirectoryApp"]})
        return fake_services(request)

    return httpx.MockTransport(handler)


@pytest.fixture
def sp_enabled(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Any:
    """The example, run through ``main()`` with ``--include all``, writing to a temp dir.

    Its patterns are ``backup`` and the reporter app's ID, as the command line would set them.
    """
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", PARTNER_TENANT)
    monkeypatch.setenv(ENV_PREFIX + "USERNAME", USER)
    monkeypatch.setattr(cli, "configure_logging", lambda verbosity: None)
    module = load_script(SERVICE_PRINCIPAL_ENABLED)
    monkeypatch.setattr(module, "OUT_DIR", tmp_path / "reports")
    patterns = [module.pattern("backup"), module.pattern(REPORTER_APP_ID)]
    monkeypatch.setattr(module, "PATTERNS", patterns)
    return module


def _disable(monkeypatch: pytest.MonkeyPatch, module: Any) -> None:
    """Set what ``--disable`` sets."""
    monkeypatch.setattr(module, "DISABLE", True)
    monkeypatch.setattr(module, "SCOPES", ["Application.ReadWrite.All"])


def _run_all(module: Any, transport: httpx.MockTransport) -> int:
    parser = argparse.ArgumentParser()
    cli.add_arguments(parser)
    return cli.main(module, parser.parse_args(["--include", "all"]), transport=transport)


def _output(tmp_path: Path, suffix: str = "xlsx") -> Path:
    """Return where the example writes today's workbook, or its JSON with ``json``."""
    today = datetime.date.today().isoformat()
    return tmp_path / "reports" / f"ServicePrincipalEnabled_{today}.{suffix}"


def test_sp_enabled_reports_every_match_and_changes_nothing(
    fake_msal: FakeMsal,
    sp_enabled: Any,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    caplog: pytest.LogCaptureFixture,
) -> None:
    caplog.set_level(logging.INFO)
    patches: list[httpx.Request] = []
    apps = {
        TENANT_A: [
            _sp("sp-a1", "ConnectWise Cloud Backup", BACKUP_APP_ID, enabled=True),
            _sp("sp-a2", "Reporter", REPORTER_APP_ID, enabled=False),
            _sp("sp-a3", "Mail", MAIL_APP_ID, enabled=True),
        ],
        TENANT_B: [_sp("sp-b1", "ConnectWise Cloud Backup", BACKUP_APP_ID, enabled=False)],
    }

    assert _run_all(sp_enabled, _sp_services(apps, patches)) == 0

    assert patches == []
    path = _output(tmp_path)
    assert [p.name for p in path.parent.iterdir()] == [path.name]
    assert capsys.readouterr().out == f"Wrote {path}\n"
    assert _sheets(path) == {
        "Tenants": [
            ("Tenant", "Status", "TenantId", "ConnectWise Cloud Backup", "Reporter"),
            ("Alpha", "ok", TENANT_A, "enabled", "disabled"),
            ("beta corp", "ok", TENANT_B, "disabled", None),
        ]
    }
    # Nothing is written while tenants run.
    assert [r for r in caplog.records if r.filename == SERVICE_PRINCIPAL_ENABLED.name] == []
    # A report asks only to read.
    assert (
        GRAPH_CLI_CLIENT_ID,
        ("https://graph.microsoft.com/Application.Read.All",),
    ) in fake_msal.silent_calls(TENANT_A)


def test_sp_enabled_disable_turns_sign_in_off_where_it_is_on(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, sp_enabled: Any, tmp_path: Path
) -> None:
    _disable(monkeypatch, sp_enabled)
    patches: list[httpx.Request] = []
    apps = {
        TENANT_A: [
            _sp("sp-a1", "ConnectWise Cloud Backup", BACKUP_APP_ID, enabled=True),
            _sp("sp-a2", "Reporter", REPORTER_APP_ID, enabled=False),
        ],
        TENANT_B: [_sp("sp-b1", "ConnectWise Cloud Backup", BACKUP_APP_ID, enabled=True)],
    }

    assert _run_all(sp_enabled, _sp_services(apps, patches)) == 0

    assert [(bearer_tenant(p), p.url.path) for p in patches] == [
        (TENANT_A, "/v1.0/servicePrincipals/sp-a1"),
        (TENANT_B, "/v1.0/servicePrincipals/sp-b1"),
    ]
    assert all(json.loads(p.content) == {"accountEnabled": False} for p in patches)
    assert _sheets(_output(tmp_path))["Tenants"][1:] == [
        ("Alpha", "ok", TENANT_A, "disabled", "disabled"),
        ("beta corp", "ok", TENANT_B, "disabled", None),
    ]
    assert (
        GRAPH_CLI_CLIENT_ID,
        ("https://graph.microsoft.com/Application.ReadWrite.All",),
    ) in fake_msal.silent_calls(TENANT_A)


def test_sp_enabled_exports_each_match_whole_keyed_by_tenant_id(
    fake_msal: FakeMsal,
    monkeypatch: pytest.MonkeyPatch,
    sp_enabled: Any,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    monkeypatch.setattr(sp_enabled, "EXPORT_JSON", True)
    # With --disable, the export shows each app as the run left it.
    _disable(monkeypatch, sp_enabled)
    apps = {
        TENANT_A: [
            _sp("sp-a1", "ConnectWise Cloud Backup", BACKUP_APP_ID, enabled=True),
            _sp("sp-a3", "Mail", MAIL_APP_ID, enabled=True),
        ]
    }

    assert _run_all(sp_enabled, _sp_services(apps, [])) == 0

    path = _output(tmp_path, "json")
    assert capsys.readouterr().out == f"Wrote {_output(tmp_path)}\nWrote {path}\n"
    backup = _sp("sp-a1", "ConnectWise Cloud Backup", BACKUP_APP_ID, enabled=False)
    assert json.loads(path.read_text(encoding="utf-8")) == {
        TENANT_A: [{**backup, "tags": ["WindowsAzureActiveDirectoryApp"]}],
        TENANT_B: [],
    }


def test_sp_enabled_workbook_lists_failed_tenants_and_apps_that_share_a_name(
    sp_enabled: Any, tmp_path: Path
) -> None:
    sp_enabled.report(
        [
            TenantResult(
                Tenant(TENANT_A, "Alpha"),
                "ok",
                [
                    _sp("sp-1", "Backup", "app-1", enabled=True),
                    _sp("sp-2", "Backup", "app-2", enabled=False),
                    {"id": "sp-3", "appId": "app-3", "accountEnabled": True},
                ],
            ),
            TenantResult(Tenant(TENANT_B, "beta corp"), "auth_failed", error="AADSTS65001"),
        ]
    )
    sheet = load_workbook(_output(tmp_path))["Tenants"]
    assert list(sheet.iter_rows(values_only=True)) == [
        # An app without a display name is shown by its app ID.
        ("Tenant", "Status", "TenantId", "app-3", "Backup"),
        ("Alpha", "ok", TENANT_A, "enabled", "disabled, enabled"),
        ("beta corp", "auth_failed", TENANT_B, None, None),
    ]
    assert sheet.freeze_panes == "A2"
    (table,) = sheet.tables.values()
    assert table.ref == "A1:E3"


def test_sp_enabled_writes_nothing_when_no_tenant_ran(
    sp_enabled: Any, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    sp_enabled.report([])
    assert not (tmp_path / "reports").exists()
    assert capsys.readouterr().out == "No tenants ran; nothing written.\n"


def _refused_in_alpha(transport: httpx.MockTransport) -> httpx.MockTransport:
    """Wrap a fake Graph so Alpha refuses the app read, as without a role that may read it."""

    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path == "/v1.0/servicePrincipals" and bearer_tenant(request) == TENANT_A:
            error = {"code": "Authorization_RequestDenied", "message": "Insufficient privileges."}
            return httpx.Response(403, json={"error": error})
        return transport.handle_request(request)

    return httpx.MockTransport(handler)


def test_sp_enabled_report_goes_on_after_a_tenant_error(
    fake_msal: FakeMsal, sp_enabled: Any, tmp_path: Path
) -> None:
    apps = {TENANT_B: [_sp("sp-b1", "ConnectWise Cloud Backup", BACKUP_APP_ID, enabled=True)]}

    assert _run_all(sp_enabled, _refused_in_alpha(_sp_services(apps, []))) == 1

    alpha, beta = _sheets(_output(tmp_path))["Tenants"][1:]
    assert alpha == ("Alpha", "error", TENANT_A, None)
    assert beta == ("beta corp", "ok", TENANT_B, "enabled")


def test_sp_enabled_disable_goes_on_after_a_tenant_error(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, sp_enabled: Any, tmp_path: Path
) -> None:
    _disable(monkeypatch, sp_enabled)
    patches: list[httpx.Request] = []
    apps = {TENANT_B: [_sp("sp-b1", "ConnectWise Cloud Backup", BACKUP_APP_ID, enabled=True)]}

    assert _run_all(sp_enabled, _refused_in_alpha(_sp_services(apps, patches))) == 1

    rows = _sheets(_output(tmp_path))["Tenants"][1:]
    assert [(row[0], row[1]) for row in rows] == [("Alpha", "error"), ("beta corp", "ok")]
    assert rows[1][3] == "disabled"
    assert [(bearer_tenant(p), p.url.path) for p in patches] == [
        (TENANT_B, "/v1.0/servicePrincipals/sp-b1")
    ]
