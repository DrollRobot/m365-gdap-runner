"""Unit tests for tenant selection and the tenants file in tenants.py."""

from __future__ import annotations

from pathlib import Path

import pytest

from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.tenants import (
    REQUIRED_COLUMNS,
    Tenant,
    TenantFile,
    TenantFileError,
    TenantRow,
    TenantSelectionError,
    group_members,
    select,
    tenants_file_path,
)
from tests._gdap_fakes import TENANT_A, TENANT_B, TENANT_C, make_settings

pytestmark = pytest.mark.unit

ALPHA = Tenant(TENANT_A, "Alpha")
BETA = Tenant(TENANT_B, "beta corp")
GAMMA = Tenant(TENANT_C, "Gamma")

M365 = "m365_tenant_id,m365_tenant_name,m365_domain"
# A file shared with other tools: their columns sit beside this package's and the groups.
HEADER = f"{M365},patching,pilot,cw_company_id,notes"


def _write(path: Path, *lines: str, encoding: str = "utf-8") -> Path:
    path.write_text("\r\n".join(lines) + "\r\n", encoding=encoding)
    return path


@pytest.fixture
def csv_path(tmp_path: Path) -> Path:
    return _write(
        tmp_path / "tenants.csv",
        HEADER,
        f"{TENANT_A},Alpha,alpha.example,TRUE,FALSE,1001,first one",
        f"{TENANT_B},beta corp,beta.example,TRUE,TRUE,1002,",
        f"{TENANT_C},Gamma,,FALSE,,1003,gone from GDAP",
        ",,,TRUE,TRUE,1004,no M365 tenant",
    )


# --- select ----------------------------------------------------------------------------


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


# --- tenants_file_path -----------------------------------------------------------------


def test_tenants_file_path_defaults_to_config_dir(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setenv(CONFIG_DIR_ENV, str(tmp_path))
    assert tenants_file_path(make_settings()) == tmp_path / "tenants.csv"


def test_tenants_file_path_honours_setting_and_expands_home() -> None:
    assert tenants_file_path(make_settings(tenants_file="~/msp/tenants.csv")) == (
        Path.home() / "msp" / "tenants.csv"
    )


# --- TenantFile.load -------------------------------------------------------------------


def test_load_keeps_every_column_and_cell(csv_path: Path) -> None:
    tenant_file = TenantFile.load(csv_path)
    assert tenant_file.path == csv_path
    assert tenant_file.columns == HEADER.split(",")
    assert tenant_file.rows[0] == {
        "m365_tenant_id": TENANT_A,
        "m365_tenant_name": "Alpha",
        "m365_domain": "alpha.example",
        "patching": "TRUE",
        "pilot": "FALSE",
        "cw_company_id": "1001",
        "notes": "first one",
    }
    assert tenant_file.rows[2]["pilot"] == ""


def test_groups_are_the_columns_holding_only_true_false_or_blank(csv_path: Path) -> None:
    assert TenantFile.load(csv_path).groups == ["patching", "pilot"]


def test_load_accepts_a_byte_order_mark_and_short_rows(tmp_path: Path) -> None:
    path = _write(tmp_path / "t.csv", HEADER, f"{TENANT_A},Alpha", encoding="utf-8-sig")
    tenant_file = TenantFile.load(path)
    assert tenant_file.columns[0] == "m365_tenant_id"
    assert tenant_file.rows == [
        {
            "m365_tenant_id": TENANT_A,
            "m365_tenant_name": "Alpha",
            "m365_domain": "",
            "patching": "",
            "pilot": "",
            "cw_company_id": "",
            "notes": "",
        }
    ]


def test_load_of_an_empty_file_is_a_fresh_file(tmp_path: Path) -> None:
    path = tmp_path / "t.csv"
    path.write_bytes(b"")
    tenant_file = TenantFile.load(path)
    assert (tenant_file.columns, tenant_file.rows) == (list(REQUIRED_COLUMNS), [])


def test_load_requires_the_three_m365_columns(tmp_path: Path) -> None:
    path = _write(tmp_path / "t.csv", "m365_tenant_id,patching", f"{TENANT_A},TRUE")
    with pytest.raises(
        TenantFileError, match=r"lacks the column\(s\) m365_tenant_name, m365_domain"
    ):
        TenantFile.load(path)


def test_load_rejects_non_utf8(tmp_path: Path) -> None:
    path = tmp_path / "t.csv"
    path.write_bytes(b"\xff\xfet\x00e\x00")
    with pytest.raises(TenantFileError, match="not UTF-8"):
        TenantFile.load(path)


def test_load_reports_an_unreadable_path(tmp_path: Path) -> None:
    with pytest.raises(TenantFileError, match="Cannot read"):
        TenantFile.load(tmp_path)  # a directory


# --- TenantFile.members ----------------------------------------------------------------


def test_members_returns_marked_rows_in_file_order(csv_path: Path) -> None:
    tenant_file = TenantFile.load(csv_path)
    assert tenant_file.members("patching") == [
        TenantRow(TENANT_A, "Alpha", "alpha.example"),
        TenantRow(TENANT_B, "beta corp", "beta.example"),
    ]
    assert tenant_file.members("pilot") == [TenantRow(TENANT_B, "beta corp", "beta.example")]


def test_members_skips_a_marked_row_with_no_m365_tenant(tmp_path: Path) -> None:
    # A customer another tool knows about, with no M365 tenant: not this package's business.
    path = _write(tmp_path / "t.csv", f"{M365},g", ",Acme,,TRUE", f"{TENANT_A},Alpha,,TRUE")
    assert [row.tenant_id for row in TenantFile.load(path).members("g")] == [TENANT_A]


@pytest.mark.parametrize("cell", ["TRUE", "true", " True "])
def test_members_accepts_true_in_any_case(tmp_path: Path, cell: str) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},g", f"{TENANT_A},Alpha,,{cell}")
    assert [row.tenant_id for row in TenantFile.load(path).members("g")] == [TENANT_A]


@pytest.mark.parametrize("cell", ["FALSE", "false", "", "  "])
def test_members_skips_false_and_blank(tmp_path: Path, cell: str) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},g", f"{TENANT_A},Alpha,,{cell}")
    assert TenantFile.load(path).members("g") == []


@pytest.mark.parametrize("cell", ["maybe", "yes", "no", "1", "0", "x"])
def test_members_rejects_everything_else_naming_the_row(tmp_path: Path, cell: str) -> None:
    path = _write(
        tmp_path / "t.csv", f"{M365},g", f"{TENANT_A},Alpha,,TRUE", f"{TENANT_B},beta corp,,{cell}"
    )
    with pytest.raises(TenantFileError, match=f"row 3: '{cell}' in column 'g' is not TRUE, FALSE"):
        TenantFile.load(path).members("g")


def test_members_leaves_other_tools_columns_unchecked(csv_path: Path) -> None:
    # 'cw_company_id' and 'notes' are another tool's; only asking for one as a group checks it.
    assert len(TenantFile.load(csv_path).members("patching")) == 2
    with pytest.raises(TenantFileError, match="'1001' in column 'cw_company_id'"):
        TenantFile.load(csv_path).members("cw_company_id")


def test_members_matches_the_group_name_case_insensitively(csv_path: Path) -> None:
    assert len(TenantFile.load(csv_path).members("PATCHING")) == 2


def test_members_unknown_group_lists_only_the_boolean_columns(csv_path: Path) -> None:
    with pytest.raises(
        TenantFileError, match=r"no group column 'm365_domain'\. Groups: patching, pilot\."
    ):
        TenantFile.load(csv_path).members("m365_domain")


def test_members_ambiguous_group_name_is_an_error(tmp_path: Path) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},Pilot,pilot", f"{TENANT_A},Alpha,,TRUE,TRUE")
    with pytest.raises(TenantFileError, match="'pilot' matches 2 columns"):
        TenantFile.load(path).members("pilot")


# --- TenantFile.add_missing / save -----------------------------------------------------


def test_add_missing_appends_blank_rows_for_unknown_ids_only(csv_path: Path) -> None:
    tenant_file = TenantFile.load(csv_path)
    delta = Tenant("dddddddd-0000-0000-0000-00000000000d", "Delta")
    added = tenant_file.add_missing([Tenant(TENANT_A.upper(), "Alpha renamed"), delta])
    assert added == [delta]
    assert tenant_file.rows[-1] == {
        "m365_tenant_id": delta.tenant_id,
        "m365_tenant_name": "Delta",
        "m365_domain": "",
        "patching": "",
        "pilot": "",
        "cw_company_id": "",
        "notes": "",
    }
    assert tenant_file.rows[0]["m365_tenant_name"] == "Alpha"  # existing rows are left alone


def test_save_round_trips_other_tools_data_and_appended_rows(csv_path: Path) -> None:
    tenant_file = TenantFile.load(csv_path)
    delta = Tenant("dddddddd-0000-0000-0000-00000000000d", "Delta")
    tenant_file.add_missing([delta])
    tenant_file.save()
    assert csv_path.read_text(encoding="utf-8-sig").splitlines() == [
        HEADER,
        f"{TENANT_A},Alpha,alpha.example,TRUE,FALSE,1001,first one",
        f"{TENANT_B},beta corp,beta.example,TRUE,TRUE,1002,",
        f"{TENANT_C},Gamma,,FALSE,,1003,gone from GDAP",
        ",,,TRUE,TRUE,1004,no M365 tenant",
        f"{delta.tenant_id},Delta,,,,,",
    ]
    assert csv_path.read_bytes().startswith(b"\xef\xbb\xbf")  # BOM for spreadsheets
    assert [p.name for p in csv_path.parent.iterdir()] == [csv_path.name]  # no temp file left


def test_save_creates_the_file_and_its_directory(tmp_path: Path) -> None:
    path = tmp_path / "deeper" / "tenants.csv"
    tenant_file = TenantFile.empty(path)
    tenant_file.add_missing([ALPHA])
    tenant_file.save()
    assert path.read_text(encoding="utf-8-sig").splitlines() == [M365, f"{TENANT_A},Alpha,"]


# --- group_members ---------------------------------------------------------------------


def test_group_members_unions_groups_in_gdap_order(csv_path: Path) -> None:
    tenant_file = TenantFile.load(csv_path)
    tenants = [BETA, ALPHA]  # GDAP order differs from file order
    assert group_members(tenants, tenant_file, ["pilot"]) == [BETA]
    assert group_members(tenants, tenant_file, ["pilot", "patching"]) == [BETA, ALPHA]
    assert group_members(tenants, tenant_file, []) == []


def test_group_members_without_gdap_is_an_error(csv_path: Path) -> None:
    # Gamma has a row but no GDAP relationship: harmless until a group selects it.
    tenant_file = TenantFile.load(csv_path)
    assert group_members([ALPHA, BETA], tenant_file, ["pilot"]) == [BETA]
    path = _write(csv_path, f"{M365},nope", f"{TENANT_C},Gamma,,TRUE")
    with pytest.raises(TenantSelectionError, match=rf"'Gamma' \({TENANT_C}\) is in group 'nope'"):
        group_members([ALPHA, BETA], TenantFile.load(path), ["nope"])


def test_group_members_propagates_file_errors(csv_path: Path) -> None:
    with pytest.raises(TenantFileError, match="no group column 'missing'"):
        group_members([ALPHA, BETA], TenantFile.load(csv_path), ["missing"])
