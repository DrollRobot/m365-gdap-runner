"""Integration tests: runner with the real session, azure-auth clients and the examples."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import httpx
import pytest
from azure_auth.constants import EXCHANGE_POWERSHELL_CLIENT_ID, GRAPH_CLI_CLIENT_ID

from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.config.schema import ENV_PREFIX
from m365_gdap_runner.loader import Payload, payload_from_module
from m365_gdap_runner.runner import run_payload
from m365_gdap_runner.session import PartnerSession
from tests._gdap_fakes import (
    CONSENT_REQUIRED,
    TENANT_A,
    TENANT_B,
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
    results = run_payload(session, payload, tenants, output_format="csv")

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
    assert "AADSTS65001" in results[1].error

    # Consent was checked with the payload's own scopes, not .default.
    assert fake_msal.silent_calls(TENANT_B) == [
        (GRAPH_CLI_CLIENT_ID, ("https://graph.microsoft.com/User.Read.All",))
    ]
    # The payload never ran for the tenant that refused the token.
    users_requests = [r for r in seen if r.url.path == "/v1.0/users"]
    assert [bearer_tenant(r) for r in users_requests] == [TENANT_A]
    assert users_requests[0].headers["ConsistencyLevel"] == "eventual"
    assert fake_msal.interactive_calls == 1


def test_exchange_example_signs_in_to_exchange_and_runs_the_cmdlet(fake_msal: FakeMsal) -> None:
    payload = _example(AUDIT_DISABLED)
    assert (payload.clients, payload.scopes) == (("exchange",), ())

    session = PartnerSession(make_settings(), transport=httpx.MockTransport(fake_services))
    session.login(payload.clients)
    results = run_payload(session, payload, session.list_tenants())

    assert [r.status for r in results] == ["ok", "ok"]
    assert [r.result for r in results] == [
        {"name": f"{TENANT_A}.onmicrosoft.com", "auditDisabled": False},
        {"name": f"{TENANT_B}.onmicrosoft.com", "auditDisabled": True},
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
            {"Identity": "svc-backup", "AuditBypassEnabled": gaps},
            {"Identity": "svc-other", "AuditBypassEnabled": False},
        ],
    }
    return httpx.Response(200, json={"value": answers[cmdlet]})


def test_mailbox_audit_example_reads_mailboxes_through_the_rest_endpoint(
    fake_msal: FakeMsal,
) -> None:
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return _audit_services(request)

    payload = _example(MAILBOX_AUDIT_GAPS)
    session = PartnerSession(make_settings(), transport=httpx.MockTransport(handler))
    session.login(payload.clients)
    results = run_payload(session, payload, session.list_tenants(), output_format="csv")

    assert [r.status for r in results] == ["ok", "ok"]
    assert results[0].result == []
    domain = f"{TENANT_B}.onmicrosoft.com"
    columns = ("finding", "identity", "recipientTypeDetails", "missingActions")
    assert results[1].result == [
        dict(zip(columns, row, strict=True))
        for row in [
            ("audit_disabled", domain, "", ""),
            ("unified_audit_log_off", domain, "", ""),
            ("audit_gap", "cee@b", "UserMailbox", "Move"),
            ("audit_gap", "Zed@b", "SharedMailbox", "MailboxLogin"),
            ("audit_bypass", "svc-backup", "", ""),
        ]
    ]
    # Get-EXOMailbox -PropertySets Audit -RecipientTypeDetails ... -Properties ...
    (read,) = [r for r in seen if r.url.path == f"/adminapi/beta/{TENANT_B}/Mailbox"]
    assert read.method == "GET"
    assert dict(read.url.params) == {
        "PropertySet": "Audit",
        "RecipientTypeDetails": "UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox",
        "$select": "DefaultAuditSet,UserPrincipalName,RecipientTypeDetails",
    }
