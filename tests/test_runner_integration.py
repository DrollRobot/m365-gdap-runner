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
ENTRA_SIGN_IN_LOGS = EXAMPLES / "entra_sign_in_logs.py"
ENTRA_AUDIT_LOGS = EXAMPLES / "entra_audit_logs.py"
APP_SCOPES = EXAMPLES / "app_scopes.py"


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
    # Consent was checked in the tenant, which has none, and beta corp has no account to try.
    assert results[1].error == "partner: no Graph consent; direct: no m365_username"
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


def _assert_consolas(path: Path) -> None:
    """Check a workbook is in Consolas throughout: every cell, and its default font."""
    for sheet in load_workbook(path).worksheets:
        cells = [cell for row in sheet.iter_rows() for cell in row]
        assert {cell.font.name for cell in cells} == {"Consolas"}
        # A theme scheme would make Excel show the theme's font, Calibri, in Consolas's place.
        assert {cell.font.scheme for cell in cells} == {None}
        # A cell nothing was written to has the workbook's default font.
        assert sheet["Z99"].font.name == "Consolas"


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


def test_mailbox_audit_keeps_an_earlier_runs_workbooks(
    fake_msal: FakeMsal,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    mailbox_audit_gaps: Any,
    capsys: pytest.CaptureFixture[str],
) -> None:
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", PARTNER_TENANT)
    monkeypatch.setenv(ENV_PREFIX + "USERNAME", USER)
    monkeypatch.setattr(cli, "configure_logging", lambda verbosity: None)
    today = datetime.date.today().isoformat()
    reports = tmp_path / "reports"
    reports.mkdir()
    earlier = [
        reports / f"MailboxAuditGaps_{TENANT_A}_{today}.xlsx",
        reports / f"MailboxAuditGaps_summary_{today}.xlsx",
    ]
    for path in earlier:
        path.write_bytes(path.name.encode())

    parser = argparse.ArgumentParser()
    cli.add_arguments(parser)
    args = parser.parse_args(["--include", "all"])
    assert cli.main(mailbox_audit_gaps, args, transport=httpx.MockTransport(_audit_services)) == 0

    for path in earlier:
        assert path.read_bytes() == path.name.encode()
    alpha, summary = (path.with_name(f"{path.stem}_2.xlsx") for path in earlier)
    assert capsys.readouterr().out == f"Wrote {summary}\n"
    # The summary names the workbook the tenant got.
    assert _sheets(summary)["Tenants"][1][8] == alpha.name
    _assert_consolas(alpha)
    _assert_consolas(summary)


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


def test_unified_audit_log_keeps_an_earlier_workbook(
    tmp_path: Path, unified_audit_log: Any, capsys: pytest.CaptureFixture[str]
) -> None:
    results = [TenantResult(Tenant(TENANT_A, "Alpha"), "ok", True)]
    unified_audit_log.report(results)
    (first,) = (tmp_path / "reports").iterdir()
    earlier = first.read_bytes()

    unified_audit_log.report(results)

    second = first.with_name(f"{first.stem}_2.xlsx")
    assert sorted(p.name for p in (tmp_path / "reports").iterdir()) == [first.name, second.name]
    assert first.read_bytes() == earlier
    assert capsys.readouterr().out == f"Wrote {first}\nWrote {second}\n"
    _assert_consolas(second)


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
            _sp("sp-a1", "Contoso Cloud Backup", BACKUP_APP_ID, enabled=True),
            _sp("sp-a2", "Reporter", REPORTER_APP_ID, enabled=False),
            _sp("sp-a3", "Mail", MAIL_APP_ID, enabled=True),
        ],
        TENANT_B: [_sp("sp-b1", "Contoso Cloud Backup", BACKUP_APP_ID, enabled=False)],
    }

    assert _run_all(sp_enabled, _sp_services(apps, patches)) == 0

    assert patches == []
    path = _output(tmp_path)
    assert [p.name for p in path.parent.iterdir()] == [path.name]
    assert capsys.readouterr().out == f"Wrote {path}\n"
    assert _sheets(path) == {
        "Tenants": [
            ("Tenant", "Status", "TenantId", "Contoso Cloud Backup", "Reporter"),
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
            _sp("sp-a1", "Contoso Cloud Backup", BACKUP_APP_ID, enabled=True),
            _sp("sp-a2", "Reporter", REPORTER_APP_ID, enabled=False),
        ],
        TENANT_B: [_sp("sp-b1", "Contoso Cloud Backup", BACKUP_APP_ID, enabled=True)],
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
            _sp("sp-a1", "Contoso Cloud Backup", BACKUP_APP_ID, enabled=True),
            _sp("sp-a3", "Mail", MAIL_APP_ID, enabled=True),
        ]
    }

    assert _run_all(sp_enabled, _sp_services(apps, [])) == 0

    path = _output(tmp_path, "json")
    assert capsys.readouterr().out == f"Wrote {_output(tmp_path)}\nWrote {path}\n"
    backup = _sp("sp-a1", "Contoso Cloud Backup", BACKUP_APP_ID, enabled=False)
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


def test_sp_enabled_keeps_an_earlier_runs_files(
    monkeypatch: pytest.MonkeyPatch,
    sp_enabled: Any,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    monkeypatch.setattr(sp_enabled, "EXPORT_JSON", True)
    earlier = [_output(tmp_path), _output(tmp_path, "json")]
    earlier[0].parent.mkdir(parents=True)
    for path in earlier:
        path.write_bytes(path.name.encode())
    app = _sp("sp-1", "Backup", BACKUP_APP_ID, enabled=True)

    sp_enabled.report([TenantResult(Tenant(TENANT_A, "Alpha"), "ok", [app])])

    for path in earlier:
        assert path.read_bytes() == path.name.encode()
    workbook, exported = (path.with_name(f"{path.stem}_2{path.suffix}") for path in earlier)
    assert capsys.readouterr().out == f"Wrote {workbook}\nWrote {exported}\n"
    assert json.loads(exported.read_text(encoding="utf-8")) == {TENANT_A: [app]}
    assert _sheets(workbook)["Tenants"][1] == ("Alpha", "ok", TENANT_A, "enabled")
    _assert_consolas(workbook)


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
    apps = {TENANT_B: [_sp("sp-b1", "Contoso Cloud Backup", BACKUP_APP_ID, enabled=True)]}

    assert _run_all(sp_enabled, _refused_in_alpha(_sp_services(apps, []))) == 1

    alpha, beta = _sheets(_output(tmp_path))["Tenants"][1:]
    assert alpha == ("Alpha", "error", TENANT_A, None)
    assert beta == ("beta corp", "ok", TENANT_B, "enabled")


def test_sp_enabled_disable_goes_on_after_a_tenant_error(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, sp_enabled: Any, tmp_path: Path
) -> None:
    _disable(monkeypatch, sp_enabled)
    patches: list[httpx.Request] = []
    apps = {TENANT_B: [_sp("sp-b1", "Contoso Cloud Backup", BACKUP_APP_ID, enabled=True)]}

    assert _run_all(sp_enabled, _refused_in_alpha(_sp_services(apps, patches))) == 1

    rows = _sheets(_output(tmp_path))["Tenants"][1:]
    assert [(row[0], row[1]) for row in rows] == [("Alpha", "error"), ("beta corp", "ok")]
    assert rows[1][3] == "disabled"
    assert [(bearer_tenant(p), p.url.path) for p in patches] == [
        (TENANT_B, "/v1.0/servicePrincipals/sp-b1")
    ]


# --- entra_sign_in_logs.py ---------------------------------------------------------------

SIGN_IN_FILTER = "userPrincipalName eq 'admin@contoso.com'"


def _sign_in(number: int) -> dict[str, Any]:
    """Build a sign-in as Graph returns it, in part."""
    return {
        "id": f"si-{number}",
        "userPrincipalName": "admin@contoso.com",
        "status": {"errorCode": 0},
    }


def _sign_in_services(
    sign_ins: dict[str, list[dict[str, Any]]], reads: list[httpx.Request], version: str = "v1.0"
) -> httpx.MockTransport:
    """Fake Graph holding *sign_ins* per tenant under *version*, recording every read of them.

    A tenant's sign-ins come one per page, each page but the last with a next link, so a read
    of more than one follows the links.
    """

    def handler(request: httpx.Request) -> httpx.Response:
        held = sign_ins.get(bearer_tenant(request), [])
        if request.url.path == f"/{version}/auditLogs/signIns":
            reads.append(request)
            index = int(request.url.params.get("$skiptoken", "0"))
            page: dict[str, Any] = {"value": held[index : index + 1]}
            if index + 1 < len(held):
                page["@odata.nextLink"] = (
                    f"https://graph.microsoft.com/{version}/auditLogs/signIns"
                    f"?$skiptoken={index + 1}"
                )
            return httpx.Response(200, json=page)
        return fake_services(request)

    return httpx.MockTransport(handler)


@pytest.fixture
def sign_in_logs(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Any:
    """The example, run through ``main()`` with ``--include all``, writing to a temp dir.

    Its filter is :data:`SIGN_IN_FILTER`, as the command line would set it.
    """
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", PARTNER_TENANT)
    monkeypatch.setenv(ENV_PREFIX + "USERNAME", USER)
    monkeypatch.setattr(cli, "configure_logging", lambda verbosity: None)
    module = load_script(ENTRA_SIGN_IN_LOGS)
    monkeypatch.setattr(module, "OUT_DIR", tmp_path / "reports")
    monkeypatch.setattr(module, "FILTER", SIGN_IN_FILTER)
    return module


def _sign_in_output(tmp_path: Path, tenant: str | None = None) -> Path:
    """Return where the example writes today's workbook, or with *tenant* that tenant's JSON."""
    today = datetime.date.today().isoformat()
    if tenant is None:
        return tmp_path / "reports" / f"EntraSignInLogs_{today}.xlsx"
    return tmp_path / "reports" / f"EntraSignInLogs_{tenant}_{today}.json"


def test_sign_in_logs_writes_json_per_tenant_with_matches_and_a_count_workbook(
    fake_msal: FakeMsal,
    sign_in_logs: Any,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    caplog: pytest.LogCaptureFixture,
) -> None:
    caplog.set_level(logging.INFO)
    reads: list[httpx.Request] = []
    held = [_sign_in(1), _sign_in(2), _sign_in(3)]

    assert _run_all(sign_in_logs, _sign_in_services({TENANT_A: held}, reads)) == 0

    alpha_json = _sign_in_output(tmp_path, "Alpha")
    workbook = _sign_in_output(tmp_path)
    assert sorted(p.name for p in (tmp_path / "reports").iterdir()) == sorted(
        [alpha_json.name, workbook.name]
    )
    # Every page, whole, as Graph returned it.
    assert json.loads(alpha_json.read_text(encoding="utf-8")) == held
    assert capsys.readouterr().out == f"Wrote {workbook}\n"
    assert _sheets(workbook) == {
        "Tenants": [
            (SIGN_IN_FILTER, None, None, None, None),
            ("Tenant", "TenantId", "Status", "SignIns", "Error"),
            ("Alpha", TENANT_A, "ok", 3, None),
            ("beta corp", TENANT_B, "ok", 0, None),
        ]
    }
    # The filter goes on the first page; next links carry their own query.
    firsts = [r for r in reads if "$skiptoken" not in r.url.params]
    assert [(bearer_tenant(r), r.url.params["$filter"]) for r in firsts] == [
        (TENANT_A, SIGN_IN_FILTER),
        (TENANT_B, SIGN_IN_FILTER),
    ]
    assert len(reads) == 4
    assert sorted(
        r.getMessage() for r in caplog.records if r.filename == ENTRA_SIGN_IN_LOGS.name
    ) == [f"3 sign-ins match; wrote {alpha_json}", "No sign-ins match"]


def test_sign_in_logs_beta_reads_the_beta_endpoint(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, sign_in_logs: Any, tmp_path: Path
) -> None:
    monkeypatch.setattr(sign_in_logs, "BETA", True)
    reads: list[httpx.Request] = []
    transport = _sign_in_services({TENANT_B: [_sign_in(1)]}, reads, version="beta")

    assert _run_all(sign_in_logs, transport) == 0

    assert [(bearer_tenant(r), r.url.host, r.url.path) for r in reads] == [
        (TENANT_A, "graph.microsoft.com", "/beta/auditLogs/signIns"),
        (TENANT_B, "graph.microsoft.com", "/beta/auditLogs/signIns"),
    ]
    beta_json = _sign_in_output(tmp_path, "beta_corp")
    assert json.loads(beta_json.read_text(encoding="utf-8")) == [_sign_in(1)]
    rows = _sheets(_sign_in_output(tmp_path))["Tenants"]
    assert rows[0][0] == SIGN_IN_FILTER
    assert [row[3] for row in rows[2:]] == [0, 1]


def test_sign_in_logs_report_goes_on_after_a_tenant_error(
    fake_msal: FakeMsal, sign_in_logs: Any, tmp_path: Path
) -> None:
    reads: list[httpx.Request] = []
    held = _sign_in_services({TENANT_B: [_sign_in(1)]}, reads)

    def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path == "/v1.0/auditLogs/signIns" and bearer_tenant(request) == TENANT_A:
            error = {
                "code": "Authentication_RequestFromNonPremiumTenantOrB2CTenant",
                "message": "Neither tenant is B2C or tenant doesn't have premium license",
            }
            return httpx.Response(403, json={"error": error})
        return held.handle_request(request)

    assert _run_all(sign_in_logs, httpx.MockTransport(handler)) == 1

    alpha, beta = _sheets(_sign_in_output(tmp_path))["Tenants"][2:]
    assert alpha[:4] == ("Alpha", TENANT_A, "error", None)
    assert "premium license" in alpha[4]
    assert beta == ("beta corp", TENANT_B, "ok", 1, None)


def test_sign_in_logs_workbook_has_the_filter_title_and_a_table_below_it(
    sign_in_logs: Any, tmp_path: Path
) -> None:
    sign_in_logs.report(
        [
            TenantResult(Tenant(TENANT_A, "Alpha"), "ok", 2),
            TenantResult(Tenant(TENANT_B, "beta corp"), "auth_failed", error="AADSTS65001"),
        ]
    )
    sheet = load_workbook(_sign_in_output(tmp_path))["Tenants"]
    assert sheet["A1"].value == SIGN_IN_FILTER
    assert sheet["A1"].font.b
    assert (sheet["A1"].font.sz, sheet["B3"].font.sz) == (20, 11)
    assert {cell.font.name for row in sheet.iter_rows() for cell in row} == {"Consolas"}
    # A theme scheme would make Excel show the theme's font, Calibri, in Consolas's place.
    assert {cell.font.scheme for row in sheet.iter_rows() for cell in row} == {None}
    assert sheet.column_dimensions["B"].width >= len(TENANT_A) + 2
    assert [str(cells) for cells in sheet.merged_cells.ranges] == ["A1:E1"]
    assert list(sheet.iter_rows(min_row=3, values_only=True)) == [
        ("Alpha", TENANT_A, "ok", 2, None),
        ("beta corp", TENANT_B, "auth_failed", None, "AADSTS65001"),
    ]
    assert sheet.freeze_panes == "A3"
    (table,) = sheet.tables.values()
    assert table.ref == "A2:E4"
    assert table.tableStyleInfo is not None
    assert table.tableStyleInfo.name == "TableStyleMedium8"
    # The workbook's default font: a cell nothing was written to has it too. Reading it
    # creates the cell, so this comes last.
    assert sheet["Z99"].font.name == "Consolas"


def test_sign_in_logs_writes_nothing_when_no_tenant_ran(
    sign_in_logs: Any, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    sign_in_logs.report([])
    assert not (tmp_path / "reports").exists()
    assert capsys.readouterr().out == "No tenants ran; no workbook written.\n"


def test_sign_in_logs_keeps_an_earlier_runs_files(
    fake_msal: FakeMsal,
    sign_in_logs: Any,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    caplog: pytest.LogCaptureFixture,
) -> None:
    caplog.set_level(logging.INFO)
    earlier_json = _sign_in_output(tmp_path, "Alpha")
    earlier_workbook = _sign_in_output(tmp_path)
    earlier_json.parent.mkdir(parents=True)
    earlier_json.write_bytes(b"earlier sign-ins")
    earlier_workbook.write_bytes(b"earlier workbook")

    transport = _sign_in_services({TENANT_A: [_sign_in(1)]}, [])
    assert _run_all(sign_in_logs, transport) == 0

    assert earlier_json.read_bytes() == b"earlier sign-ins"
    assert earlier_workbook.read_bytes() == b"earlier workbook"
    new_json = earlier_json.with_name(f"{earlier_json.stem}_2.json")
    new_workbook = earlier_workbook.with_name(f"{earlier_workbook.stem}_2.xlsx")
    assert json.loads(new_json.read_text(encoding="utf-8")) == [_sign_in(1)]
    assert _sheets(new_workbook)["Tenants"][0][0] == SIGN_IN_FILTER
    assert capsys.readouterr().out == f"Wrote {new_workbook}\n"
    assert f"1 sign-ins match; wrote {new_json}" in caplog.messages


def test_sign_in_logs_numbers_a_file_past_every_name_taken(
    sign_in_logs: Any, tmp_path: Path
) -> None:
    out = tmp_path / "reports"
    out.mkdir()
    for name in ("Alpha.json", "Alpha_2.json", "Alpha_4.json"):
        (out / name).write_bytes(name.encode())

    paths = []
    for _ in range(3):
        path, file = sign_in_logs._create("Alpha", ".json")
        with file:
            file.write(b"new")
        paths.append(path.name)

    assert paths == ["Alpha_3.json", "Alpha_5.json", "Alpha_6.json"]
    for name in ("Alpha.json", "Alpha_2.json", "Alpha_4.json"):
        assert (out / name).read_bytes() == name.encode()


# --- entra_audit_logs.py -----------------------------------------------------------------

AUDIT_FILTER = "activityDisplayName eq 'Add member to role'"


def _directory_audit(number: int) -> dict[str, Any]:
    """Build an audit event as Graph returns it, in part."""
    return {
        "id": f"da-{number}",
        "activityDisplayName": "Add member to role",
        "result": "success",
    }


def _directory_audit_services(
    events: dict[str, list[dict[str, Any]]], reads: list[httpx.Request], version: str = "v1.0"
) -> httpx.MockTransport:
    """Fake Graph holding *events* per tenant under *version*, recording every read of them.

    A tenant's audit events come one per page, each page but the last with a next link, so a
    read of more than one follows the links.
    """

    def handler(request: httpx.Request) -> httpx.Response:
        held = events.get(bearer_tenant(request), [])
        if request.url.path == f"/{version}/auditLogs/directoryAudits":
            reads.append(request)
            index = int(request.url.params.get("$skiptoken", "0"))
            page: dict[str, Any] = {"value": held[index : index + 1]}
            if index + 1 < len(held):
                page["@odata.nextLink"] = (
                    f"https://graph.microsoft.com/{version}/auditLogs/directoryAudits"
                    f"?$skiptoken={index + 1}"
                )
            return httpx.Response(200, json=page)
        return fake_services(request)

    return httpx.MockTransport(handler)


@pytest.fixture
def audit_logs(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Any:
    """The example, run through ``main()`` with ``--include all``, writing to a temp dir.

    Its filter is :data:`AUDIT_FILTER`, as the command line would set it.
    """
    monkeypatch.setenv(ENV_PREFIX + "PARTNER_TENANT_ID", PARTNER_TENANT)
    monkeypatch.setenv(ENV_PREFIX + "USERNAME", USER)
    monkeypatch.setattr(cli, "configure_logging", lambda verbosity: None)
    module = load_script(ENTRA_AUDIT_LOGS)
    monkeypatch.setattr(module, "OUT_DIR", tmp_path / "reports")
    monkeypatch.setattr(module, "FILTER", AUDIT_FILTER)
    return module


def _audit_logs_output(tmp_path: Path, tenant: str | None = None) -> Path:
    """Return where the example writes today's workbook, or with *tenant* that tenant's JSON."""
    today = datetime.date.today().isoformat()
    if tenant is None:
        return tmp_path / "reports" / f"EntraAuditLogs_{today}.xlsx"
    return tmp_path / "reports" / f"EntraAuditLogs_{tenant}_{today}.json"


def test_audit_logs_writes_json_per_tenant_with_matches_and_a_count_workbook(
    fake_msal: FakeMsal,
    audit_logs: Any,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    caplog: pytest.LogCaptureFixture,
) -> None:
    caplog.set_level(logging.INFO)
    reads: list[httpx.Request] = []
    held = [_directory_audit(1), _directory_audit(2), _directory_audit(3)]

    assert _run_all(audit_logs, _directory_audit_services({TENANT_A: held}, reads)) == 0

    alpha_json = _audit_logs_output(tmp_path, "Alpha")
    workbook = _audit_logs_output(tmp_path)
    assert sorted(p.name for p in (tmp_path / "reports").iterdir()) == sorted(
        [alpha_json.name, workbook.name]
    )
    # Every page, whole, as Graph returned it.
    assert json.loads(alpha_json.read_text(encoding="utf-8")) == held
    assert capsys.readouterr().out == f"Wrote {workbook}\n"
    assert _sheets(workbook) == {
        "Tenants": [
            (AUDIT_FILTER, None, None, None, None),
            ("Tenant", "TenantId", "Status", "AuditEvents", "Error"),
            ("Alpha", TENANT_A, "ok", 3, None),
            ("beta corp", TENANT_B, "ok", 0, None),
        ]
    }
    # The filter goes on the first page; next links carry their own query.
    firsts = [r for r in reads if "$skiptoken" not in r.url.params]
    assert [(bearer_tenant(r), r.url.params["$filter"]) for r in firsts] == [
        (TENANT_A, AUDIT_FILTER),
        (TENANT_B, AUDIT_FILTER),
    ]
    assert len(reads) == 4
    assert sorted(
        r.getMessage() for r in caplog.records if r.filename == ENTRA_AUDIT_LOGS.name
    ) == [f"3 audit events match; wrote {alpha_json}", "No audit events match"]


def test_audit_logs_beta_reads_the_beta_endpoint(
    fake_msal: FakeMsal, monkeypatch: pytest.MonkeyPatch, audit_logs: Any, tmp_path: Path
) -> None:
    monkeypatch.setattr(audit_logs, "BETA", True)
    reads: list[httpx.Request] = []
    transport = _directory_audit_services({TENANT_B: [_directory_audit(1)]}, reads, "beta")

    assert _run_all(audit_logs, transport) == 0

    assert [(bearer_tenant(r), r.url.host, r.url.path) for r in reads] == [
        (TENANT_A, "graph.microsoft.com", "/beta/auditLogs/directoryAudits"),
        (TENANT_B, "graph.microsoft.com", "/beta/auditLogs/directoryAudits"),
    ]
    beta_json = _audit_logs_output(tmp_path, "beta_corp")
    assert json.loads(beta_json.read_text(encoding="utf-8")) == [_directory_audit(1)]
    rows = _sheets(_audit_logs_output(tmp_path))["Tenants"]
    assert rows[0][0] == AUDIT_FILTER
    assert [row[3] for row in rows[2:]] == [0, 1]


def test_audit_logs_report_goes_on_after_a_tenant_error(
    fake_msal: FakeMsal, audit_logs: Any, tmp_path: Path
) -> None:
    reads: list[httpx.Request] = []
    held = _directory_audit_services({TENANT_B: [_directory_audit(1)]}, reads)

    def handler(request: httpx.Request) -> httpx.Response:
        path = "/v1.0/auditLogs/directoryAudits"
        if request.url.path == path and bearer_tenant(request) == TENANT_A:
            error = {
                "code": "Authentication_RequestFromUnsupportedUserRole",
                "message": "User is not in the allowed roles",
            }
            return httpx.Response(403, json={"error": error})
        return held.handle_request(request)

    assert _run_all(audit_logs, httpx.MockTransport(handler)) == 1

    alpha, beta = _sheets(_audit_logs_output(tmp_path))["Tenants"][2:]
    assert alpha[:4] == ("Alpha", TENANT_A, "error", None)
    assert "allowed roles" in alpha[4]
    assert beta == ("beta corp", TENANT_B, "ok", 1, None)


def test_audit_logs_workbook_has_the_filter_title_and_a_table_below_it(
    audit_logs: Any, tmp_path: Path
) -> None:
    audit_logs.report(
        [
            TenantResult(Tenant(TENANT_A, "Alpha"), "ok", 2),
            TenantResult(Tenant(TENANT_B, "beta corp"), "auth_failed", error="AADSTS65001"),
        ]
    )
    sheet = load_workbook(_audit_logs_output(tmp_path))["Tenants"]
    assert sheet["A1"].value == AUDIT_FILTER
    assert sheet["A1"].font.b
    assert [str(cells) for cells in sheet.merged_cells.ranges] == ["A1:E1"]
    assert list(sheet.iter_rows(min_row=3, values_only=True)) == [
        ("Alpha", TENANT_A, "ok", 2, None),
        ("beta corp", TENANT_B, "auth_failed", None, "AADSTS65001"),
    ]
    assert sheet.freeze_panes == "A3"
    (table,) = sheet.tables.values()
    assert table.ref == "A2:E4"
    assert table.tableStyleInfo is not None
    assert table.tableStyleInfo.name == "TableStyleMedium8"


def test_audit_logs_writes_nothing_when_no_tenant_ran(
    audit_logs: Any, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    audit_logs.report([])
    assert not (tmp_path / "reports").exists()
    assert capsys.readouterr().out == "No tenants ran; no workbook written.\n"


# --- app_scopes.py -----------------------------------------------------------------------


@pytest.fixture
def app_scopes(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Any:
    """The example, checking for two scopes and writing to a temp dir."""
    module = load_script(APP_SCOPES)
    monkeypatch.setattr(module, "OUT_DIR", tmp_path / "reports")
    monkeypatch.setattr(module, "REQUESTED_SCOPES", ["User.Read.All", "Mail.Read"])
    return module


def _app_scopes_results() -> list[TenantResult]:
    """Return one tenant whose app lacks Mail.Read, and one that failed."""
    held = {"appId": "app-1", "scopes": ["Policy.Read.All", "User.Read.All"]}
    return [
        TenantResult(Tenant(TENANT_A, "Alpha"), "ok", held),
        TenantResult(Tenant(TENANT_B, "beta corp"), "auth_failed", error="AADSTS65001"),
    ]


def test_app_scopes_workbook_lists_every_tenant(
    app_scopes: Any, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    app_scopes.report(_app_scopes_results())

    (path,) = (tmp_path / "reports").iterdir()
    assert path.name == f"AppScopes_{datetime.date.today().isoformat()}.xlsx"
    assert capsys.readouterr().out == f"Wrote {path}\n"
    sheet = load_workbook(path)["Tenants"]
    assert list(sheet.iter_rows(values_only=True)) == [
        ("Tenant", "Status", "TenantId", "HasScopes", "Missing", "Scopes", "Error"),
        ("Alpha", "ok", TENANT_A, False, "Mail.Read", "Policy.Read.All User.Read.All", None),
        ("beta corp", "auth_failed", TENANT_B, None, None, None, "AADSTS65001"),
    ]
    assert sheet.freeze_panes == "A2"
    (table,) = sheet.tables.values()
    assert table.ref == "A1:G3"
    _assert_consolas(path)


def test_app_scopes_keeps_an_earlier_runs_files(
    app_scopes: Any,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    monkeypatch.setattr(app_scopes, "EXPORT_JSON", True)
    today = datetime.date.today().isoformat()
    reports = tmp_path / "reports"
    reports.mkdir()
    earlier = [reports / f"AppScopes_{today}.xlsx", reports / f"AppScopes_{today}.json"]
    for path in earlier:
        path.write_bytes(path.name.encode())

    app_scopes.report(_app_scopes_results())

    for path in earlier:
        assert path.read_bytes() == path.name.encode()
    workbook, exported = (path.with_name(f"{path.stem}_2{path.suffix}") for path in earlier)
    assert capsys.readouterr().out == f"Wrote {workbook}\nWrote {exported}\n"
    assert app_scopes.load_results(exported) == _app_scopes_results()
    _assert_consolas(workbook)
