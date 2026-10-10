"""Unit tests for tenant selection in tenants.py."""

from __future__ import annotations

import pytest

from m365_gdap_runner.tenants import Tenant, TenantSelectionError, select
from tests._gdap_fakes import TENANT_A, TENANT_B, TENANT_C

pytestmark = pytest.mark.unit

ALPHA = Tenant(TENANT_A, "Alpha")
BETA = Tenant(TENANT_B, "beta corp")
GAMMA = Tenant(TENANT_C, "Gamma")


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
