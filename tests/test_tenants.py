"""Unit tests for tenant selection and the tenants file in tenants.py."""

from __future__ import annotations

from pathlib import Path

import pytest

from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.tenants import (
    OWN_COLUMNS,
    Tenant,
    TenantFile,
    TenantFileError,
    TenantRow,
    TenantSelectionError,
    select,
    tenants_file_path,
)
from tests._gdap_fakes import TENANT_A, TENANT_B, TENANT_C, make_settings

pytestmark = pytest.mark.unit

ALPHA = Tenant(TENANT_A, "Alpha")
BETA = Tenant(TENANT_B, "beta corp")
GAMMA = Tenant(TENANT_C, "Gamma")
DELTA = Tenant("dddddddd-0000-0000-0000-00000000000d", "Delta")

M365 = "m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud"
# A file shared with other tools: their columns sit beside this package's and the groups.
HEADER = f"{M365},patching,pilot,cw_company_id,notes"


def _write(path: Path, *lines: str, encoding: str = "utf-8") -> Path:
    path.write_text("\r\n".join(lines) + "\r\n", encoding=encoding)
    return path


def commercial(tenant_id: str) -> str:
    """A cloud lookup that finds every tenant in the commercial cloud."""
    return "Commercial"


def nowhere(tenant_id: str) -> str:
    """A cloud lookup that finds nothing."""
    return ""


@pytest.fixture
def csv_path(tmp_path: Path) -> Path:
    return _write(
        tmp_path / "tenants.csv",
        HEADER,
        f"{TENANT_A},Alpha,alpha.example,Commercial,TRUE,FALSE,1001,first one",
        f"{TENANT_B},beta corp,beta.example,,TRUE,TRUE,1002,",
        f"{TENANT_C},Gamma,,,FALSE,,1003,gone from GDAP",
        ",,,,TRUE,TRUE,1004,no M365 tenant",
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
    with pytest.raises(
        TenantSelectionError,
        match="No tenant matches 'Contoso': it is neither a GDAP customer nor in the tenants file",
    ):
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
        "m365_cloud": "Commercial",
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
            "m365_cloud": "",
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
    assert (tenant_file.columns, tenant_file.rows) == (list(OWN_COLUMNS), [])


def test_load_requires_the_four_m365_columns(tmp_path: Path) -> None:
    path = _write(tmp_path / "t.csv", "m365_tenant_id,patching", f"{TENANT_A},TRUE")
    with pytest.raises(
        TenantFileError, match=r"lacks the column\(s\) m365_tenant_name, m365_domain, m365_cloud"
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
        TenantRow(TENANT_A, "Alpha", "Commercial"),
        TenantRow(TENANT_B, "beta corp", ""),
    ]
    assert tenant_file.members("pilot") == [TenantRow(TENANT_B, "beta corp", "")]


def test_members_reads_the_optional_username_column(tmp_path: Path) -> None:
    path = _write(
        tmp_path / "t.csv",
        f"{M365},m365_username,g",
        f"{TENANT_A},Alpha,,Commercial, admin@alpha.example ,TRUE",
        f"{TENANT_B},beta corp,,,,TRUE",
    )
    assert TenantFile.load(path).members("g") == [
        TenantRow(TENANT_A, "Alpha", "Commercial", "admin@alpha.example"),
        TenantRow(TENANT_B, "beta corp", "", ""),
    ]


def test_the_username_column_is_not_a_group(tmp_path: Path) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},m365_username,g", f"{TENANT_A},Alpha,,,,TRUE")
    tenant_file = TenantFile.load(path)
    assert tenant_file.groups == ["g"]
    with pytest.raises(TenantFileError, match="no group column 'm365_username'"):
        tenant_file.members("m365_username")


def test_members_skips_a_marked_row_with_no_m365_tenant(tmp_path: Path) -> None:
    # A customer another tool knows about, with no M365 tenant: not this package's business.
    path = _write(tmp_path / "t.csv", f"{M365},g", ",Acme,,,TRUE", f"{TENANT_A},Alpha,,,TRUE")
    assert [row.tenant_id for row in TenantFile.load(path).members("g")] == [TENANT_A]


@pytest.mark.parametrize("cell", ["TRUE", "true", " True "])
def test_members_accepts_true_in_any_case(tmp_path: Path, cell: str) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},g", f"{TENANT_A},Alpha,,,{cell}")
    assert [row.tenant_id for row in TenantFile.load(path).members("g")] == [TENANT_A]


@pytest.mark.parametrize("cell", ["FALSE", "false", "", "  "])
def test_members_skips_false_and_blank(tmp_path: Path, cell: str) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},g", f"{TENANT_A},Alpha,,,{cell}")
    assert TenantFile.load(path).members("g") == []


@pytest.mark.parametrize("cell", ["maybe", "yes", "no", "1", "0", "x"])
def test_members_rejects_everything_else_naming_the_row(tmp_path: Path, cell: str) -> None:
    path = _write(
        tmp_path / "t.csv",
        f"{M365},g",
        f"{TENANT_A},Alpha,,,TRUE",
        f"{TENANT_B},beta corp,,,{cell}",
    )
    with pytest.raises(TenantFileError, match=f"row 3: '{cell}' in column 'g' is not TRUE, FALSE"):
        TenantFile.load(path).members("g")


@pytest.mark.parametrize(("cell", "cloud"), [("usgov", "USGov"), ("CHINA", "China"), ("", "")])
def test_members_spells_the_cloud_as_azure_auth_does(tmp_path: Path, cell: str, cloud: str) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},g", f"{TENANT_A},Alpha,,{cell},TRUE")
    assert TenantFile.load(path).members("g")[0].cloud == cloud


def test_members_rejects_an_unknown_cloud_on_a_marked_row_only(tmp_path: Path) -> None:
    path = _write(
        tmp_path / "t.csv",
        f"{M365},g,h",
        f"{TENANT_A},Alpha,,Commercial,TRUE,FALSE",
        f"{TENANT_B},beta corp,,Mars,FALSE,TRUE",
    )
    tenant_file = TenantFile.load(path)
    assert len(tenant_file.members("g")) == 1  # the bad cloud is on a row 'g' does not select
    with pytest.raises(
        TenantFileError,
        match=r"row 3: 'Mars' in column 'm365_cloud' is not a cloud\. Use Commercial, USGov, "
        r"USGovDoD, China, or blank\.",
    ):
        tenant_file.members("h")


def test_members_leaves_other_tools_columns_unchecked(csv_path: Path) -> None:
    # 'cw_company_id' and 'notes' are another tool's; only asking for one as a group checks it.
    assert len(TenantFile.load(csv_path).members("patching")) == 2
    with pytest.raises(TenantFileError, match="'1001' in column 'cw_company_id'"):
        TenantFile.load(csv_path).members("cw_company_id")


def test_members_matches_the_group_name_case_insensitively(csv_path: Path) -> None:
    assert len(TenantFile.load(csv_path).members("PATCHING")) == 2


def test_members_unknown_group_lists_only_the_boolean_columns(csv_path: Path) -> None:
    with pytest.raises(
        TenantFileError, match=r"no group column 'm365_cloud'\. Groups: patching, pilot\."
    ):
        TenantFile.load(csv_path).members("m365_cloud")


def test_members_ambiguous_group_name_is_an_error(tmp_path: Path) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},Pilot,pilot", f"{TENANT_A},Alpha,,,TRUE,TRUE")
    with pytest.raises(TenantFileError, match="'pilot' matches 2 columns"):
        TenantFile.load(path).members("pilot")


# --- TenantFile.sync / save ------------------------------------------------------------


def test_sync_appends_unknown_tenants_and_fills_blank_clouds(csv_path: Path) -> None:
    looked_up: list[str] = []

    def lookup(tenant_id: str) -> str:
        looked_up.append(tenant_id)
        return "USGov"

    tenant_file = TenantFile.load(csv_path)
    report = tenant_file.sync([Tenant(TENANT_A.upper(), "Alpha renamed"), BETA, DELTA], lookup)
    # Gamma has no GDAP relationship but a row, so its blank cloud is filled in too.
    assert (report.added, report.clouds_filled, report.changed) == ([DELTA], [BETA, GAMMA], True)
    assert looked_up == [DELTA.tenant_id, TENANT_B, TENANT_C]  # Alpha's cloud is already known
    assert tenant_file.rows[0]["m365_tenant_name"] == "Alpha"  # existing rows are left alone
    assert tenant_file.rows[1]["m365_cloud"] == "USGov"
    assert tenant_file.rows[2]["m365_cloud"] == "USGov"
    assert tenant_file.rows[-1] == {
        "m365_tenant_id": DELTA.tenant_id,
        "m365_tenant_name": "Delta",
        "m365_domain": "",
        "m365_cloud": "USGov",
        "patching": "",
        "pilot": "",
        "cw_company_id": "",
        "notes": "",
    }


def test_sync_leaves_the_cloud_blank_when_the_lookup_finds_nothing(csv_path: Path) -> None:
    tenant_file = TenantFile.load(csv_path)
    report = tenant_file.sync([BETA, DELTA], nowhere)
    assert (report.added, report.clouds_filled) == ([DELTA], [])
    assert tenant_file.rows[1]["m365_cloud"] == ""
    assert tenant_file.rows[-1]["m365_cloud"] == ""


def test_sync_with_nothing_to_do_changes_nothing_and_looks_nothing_up(tmp_path: Path) -> None:
    def never(tenant_id: str) -> str:
        raise AssertionError(f"looked up {tenant_id}")

    path = _write(
        tmp_path / "t.csv",
        f"{M365},g",
        f"{TENANT_A},Alpha,,Commercial,TRUE",
        f"{TENANT_C},Gamma,,USGov,TRUE",
        ",Acme,,,TRUE",
    )
    tenant_file = TenantFile.load(path)
    before = [dict(row) for row in tenant_file.rows]
    report = tenant_file.sync([ALPHA], never)
    assert report.changed is False
    assert tenant_file.rows == before


def test_save_round_trips_other_tools_data_and_appended_rows(csv_path: Path) -> None:
    tenant_file = TenantFile.load(csv_path)
    tenant_file.sync([ALPHA, BETA, DELTA], commercial)
    tenant_file.save()
    assert csv_path.read_text(encoding="utf-8-sig").splitlines() == [
        HEADER,
        f"{TENANT_A},Alpha,alpha.example,Commercial,TRUE,FALSE,1001,first one",
        f"{TENANT_B},beta corp,beta.example,Commercial,TRUE,TRUE,1002,",
        f"{TENANT_C},Gamma,,Commercial,FALSE,,1003,gone from GDAP",
        ",,,,TRUE,TRUE,1004,no M365 tenant",
        f"{DELTA.tenant_id},Delta,,Commercial,,,,",
    ]
    assert csv_path.read_bytes().startswith(b"\xef\xbb\xbf")  # BOM for spreadsheets
    assert [p.name for p in csv_path.parent.iterdir()] == [csv_path.name]  # no temp file left


def test_save_creates_the_file_and_its_directory(tmp_path: Path) -> None:
    path = tmp_path / "deeper" / "tenants.csv"
    tenant_file = TenantFile.empty(path)
    tenant_file.sync([ALPHA], commercial)
    tenant_file.save()
    assert path.read_text(encoding="utf-8-sig").splitlines() == [
        f"{M365},m365_username",
        f"{TENANT_A},Alpha,,Commercial,",
    ]


# --- extra_tenants / details -----------------------------------------------------------


def test_extra_tenants_are_the_rows_gdap_does_not_list(tmp_path: Path) -> None:
    path = _write(
        tmp_path / "t.csv",
        M365,
        f"{TENANT_A},Alpha renamed,,",
        f"{TENANT_C},Gamma,,USGov",
        ",Acme,,",  # no M365 tenant
        f"{TENANT_C.upper()},Gamma again,,",  # the same tenant twice: the first row wins
        f"{DELTA.tenant_id},,,",  # unnamed: named by its ID
    )
    assert TenantFile.load(path).extra_tenants([ALPHA, BETA]) == [
        GAMMA,
        Tenant(DELTA.tenant_id, DELTA.tenant_id),
    ]


def test_details_take_the_cloud_and_username_from_the_row(tmp_path: Path) -> None:
    path = _write(
        tmp_path / "t.csv",
        f"{M365},m365_username",
        f"{TENANT_A},Alpha,,,",
        f"{TENANT_C},Gamma,,usgov,admin@gamma.example",
    )
    tenant_file = TenantFile.load(path)
    assert tenant_file.details(GAMMA, default_cloud="Commercial") == Tenant(
        TENANT_C, "Gamma", "USGov", "admin@gamma.example"
    )
    # A blank cloud, or no row at all, is the partner's cloud and the partner account.
    assert tenant_file.details(ALPHA, default_cloud="Commercial") == Tenant(
        TENANT_A, "Alpha", "Commercial"
    )
    assert tenant_file.details(BETA, default_cloud="Commercial") == Tenant(
        TENANT_B, "beta corp", "Commercial"
    )


def test_details_reject_an_unknown_cloud_naming_the_row(tmp_path: Path) -> None:
    path = _write(
        tmp_path / "t.csv", M365, f"{TENANT_A},Alpha,,Commercial", f"{TENANT_C},Gamma,,Mars"
    )
    with pytest.raises(TenantFileError, match="row 3: 'Mars' in column 'm365_cloud'"):
        TenantFile.load(path).details(GAMMA, default_cloud="Commercial")
