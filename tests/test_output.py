"""Unit tests for result checks and JSON / CSV output in output.py."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from m365_gdap_runner.output import (
    OutputFormatError,
    check_result,
    render_csv,
    render_json,
    write_results,
)
from m365_gdap_runner.runner import Status, TenantResult
from m365_gdap_runner.tenants import Tenant
from tests._gdap_fakes import TENANT_A, TENANT_B, TENANT_C

pytestmark = pytest.mark.unit

ALPHA = Tenant(TENANT_A, "Alpha")
BETA = Tenant(TENANT_B, "Beta")
GAMMA = Tenant(TENANT_C, "Gamma")


def _result(
    tenant: Tenant, status: Status = "ok", result: object = None, error: str | None = None
) -> TenantResult:
    return TenantResult(tenant, status, result, error)


def test_check_result_accepts_json_values() -> None:
    check_result({"users": [{"upn": "a", "count": 1, "enabled": True, "note": None}]}, "json")
    check_result(None, "json")


@pytest.mark.parametrize("value", [object(), {"when": {1, 2}}, float("nan")])
def test_check_result_rejects_values_json_cannot_hold(value: object) -> None:
    with pytest.raises(OutputFormatError, match="not JSON-serializable"):
        check_result(value, "json")


@pytest.mark.parametrize("value", [{"a": 1, "b": "x"}, [{"a": 1}, {"b": None}], [], None])
def test_csv_check_accepts_flat_rows(value: object) -> None:
    check_result(value, "csv")


@pytest.mark.parametrize("value", [{"a": {"b": 1}}, [{"a": [1]}], [1, 2], "text"])
def test_csv_check_rejects_nested_or_non_dict_results(value: object) -> None:
    with pytest.raises(OutputFormatError, match="CSV output needs"):
        check_result(value, "csv")


def test_render_json_has_one_record_per_tenant() -> None:
    results = [
        _result(ALPHA, result=[{"upn": "a@x"}]),
        _result(BETA, "auth_failed", error="consent_required"),
        _result(GAMMA, "skipped"),
    ]
    assert json.loads(render_json(results)) == [
        {
            "tenant_id": TENANT_A,
            "display_name": "Alpha",
            "status": "ok",
            "error": None,
            "result": [{"upn": "a@x"}],
        },
        {
            "tenant_id": TENANT_B,
            "display_name": "Beta",
            "status": "auth_failed",
            "error": "consent_required",
            "result": None,
        },
        {
            "tenant_id": TENANT_C,
            "display_name": "Gamma",
            "status": "skipped",
            "error": None,
            "result": None,
        },
    ]


def test_render_csv_prepends_tenant_columns_and_skips_non_ok() -> None:
    results = [
        _result(ALPHA, result=[{"upn": "a@x", "count": 1}, {"upn": "b@x", "count": 2}]),
        _result(BETA, result={"upn": "c@x", "enabled": True, "tenant_id": "spoofed"}),
        _result(GAMMA, "error", error="boom"),
    ]
    assert render_csv(results).splitlines() == [
        "tenant_id,display_name,upn,count,enabled",
        f"{TENANT_A},Alpha,a@x,1,",
        f"{TENANT_A},Alpha,b@x,2,",
        f"{TENANT_B},Beta,c@x,,True",
    ]


def test_write_results_to_file_and_stdout(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    results = [_result(ALPHA, result={"upn": "a@x"})]
    path = tmp_path / "out.json"
    write_results(results, str(path), "json")
    assert json.loads(path.read_text(encoding="utf-8"))[0]["result"] == {"upn": "a@x"}

    write_results(results, "-", "csv")
    assert capsys.readouterr().out.splitlines() == [
        "tenant_id,display_name,upn",
        f"{TENANT_A},Alpha,a@x",
    ]


def test_write_results_rejects_unknown_format(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="Unsupported output format"):
        write_results([], str(tmp_path / "out.xml"), "xml")
