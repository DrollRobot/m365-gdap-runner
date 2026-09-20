"""Unit tests for GDAP tenant enumeration and selection in tenants.py."""

from __future__ import annotations

import httpx
import pytest

from m365_gdap_runner.tenants import Tenant, TenantSelectionError, list_gdap_tenants, select
from tests._gdap_fakes import (
    PARTNER_TENANT,
    TENANT_A,
    TENANT_B,
    TENANT_C,
    FakeSession,
    bearer_tenant,
    relationships_body,
)

pytestmark = pytest.mark.unit

ALPHA = Tenant(TENANT_A, "Alpha")
BETA = Tenant(TENANT_B, "beta corp")
GAMMA = Tenant(TENANT_C, "Gamma")


def test_list_gdap_tenants_dedupes_sorts_and_filters_active() -> None:
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        body = relationships_body(
            (TENANT_B, "beta corp"), (TENANT_A, "Alpha"), (TENANT_B.upper(), "beta corp")
        )
        body["value"].append({"id": "no-customer", "status": "active"})
        body["value"].append({"id": "no-name", "customer": {"tenantId": TENANT_C}})
        return httpx.Response(200, json=body)

    session = FakeSession()
    tenants = list_gdap_tenants(session, transport=httpx.MockTransport(handler))

    assert tenants == [ALPHA, BETA, Tenant(TENANT_C, TENANT_C)]
    (request,) = seen
    assert request.url.path == "/v1.0/tenantRelationships/delegatedAdminRelationships"
    assert dict(request.url.params) == {"$filter": "status eq 'active'"}
    assert bearer_tenant(request) == PARTNER_TENANT
    assert session.credentials[PARTNER_TENANT].calls == [("DelegatedAdminRelationship.Read.All",)]


def test_select_without_filters_keeps_everything() -> None:
    assert select([ALPHA, BETA, GAMMA]) == [ALPHA, BETA, GAMMA]


def test_select_include_by_id_or_name_case_insensitive_keeps_order() -> None:
    chosen = select([ALPHA, BETA, GAMMA], include=["gamma", TENANT_A.upper()])
    assert chosen == [ALPHA, GAMMA]


def test_select_exclude() -> None:
    assert select([ALPHA, BETA, GAMMA], exclude=["BETA CORP"]) == [ALPHA, GAMMA]


def test_select_include_and_exclude() -> None:
    assert select([ALPHA, BETA, GAMMA], include=["Alpha", "Gamma"], exclude=[TENANT_C]) == [ALPHA]


@pytest.mark.parametrize("option", ["include", "exclude"])
def test_select_unknown_name_is_an_error(option: str) -> None:
    with pytest.raises(TenantSelectionError, match="No GDAP tenant matches 'Contoso'"):
        select([ALPHA, BETA], **{option: ["Contoso"]})


def test_select_ambiguous_name_is_an_error() -> None:
    twin = Tenant(TENANT_C, "Alpha")
    with pytest.raises(TenantSelectionError, match="matches 2 tenants") as info:
        select([ALPHA, twin], include=["alpha"])
    assert TENANT_A in str(info.value)
    assert TENANT_C in str(info.value)


def test_select_ambiguous_name_resolves_by_id() -> None:
    twin = Tenant(TENANT_C, "Alpha")
    assert select([ALPHA, twin], include=[TENANT_C]) == [twin]
