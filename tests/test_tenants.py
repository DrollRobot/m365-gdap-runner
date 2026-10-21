"""Unit tests for tenant selection and the tenants file in tenants.py."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet

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


GDAP = [ALPHA, BETA]


@pytest.fixture
def groups_file(tmp_path: Path) -> TenantFile:
    """Alpha and beta corp, the GDAP tenants, and Gamma, a row GDAP does not list."""
    return TenantFile.load(
        _write(
            tmp_path / "t.csv",
            f"{M365},pilot,patching",
            f"{TENANT_A},Alpha,,,TRUE,FALSE",
            f"{TENANT_B},beta corp,,,FALSE,TRUE",
            f"{TENANT_C},Gamma,,,TRUE,TRUE",
        )
    )


def test_select_all_is_every_gdap_tenant(groups_file: TenantFile) -> None:
    assert select(GDAP, groups_file, ["ALL"]) == [ALPHA, BETA]


def test_select_tenants_by_id_or_name_in_any_case_keeps_order(groups_file: TenantFile) -> None:
    assert select(GDAP, groups_file, ["gamma", TENANT_A.upper()]) == [ALPHA, GAMMA]


def test_select_a_group_is_its_marked_rows_gdap_tenants_first(groups_file: TenantFile) -> None:
    assert select(GDAP, groups_file, ["PILOT"]) == [ALPHA, GAMMA]


def test_select_inclusions_add_up(groups_file: TenantFile) -> None:
    assert select(GDAP, groups_file, ["pilot", "beta corp"]) == [ALPHA, BETA, GAMMA]
    assert select(GDAP, groups_file, ["all", "Gamma"]) == [ALPHA, BETA, GAMMA]


def test_select_an_exclusion_beats_an_inclusion(groups_file: TenantFile) -> None:
    assert select(GDAP, groups_file, ["all", "Gamma"], ["patching"]) == [ALPHA]
    assert select(GDAP, groups_file, ["pilot"], [TENANT_C]) == [ALPHA]
    # Excluding what was never included is fine.
    assert select(GDAP, groups_file, ["Alpha"], ["BETA CORP"]) == [ALPHA]


def test_select_m365_gdap_is_a_group(tmp_path: Path) -> None:
    path = _write(
        tmp_path / "t.csv",
        f"{M365},m365_gdap",
        f"{TENANT_A},Alpha,,,TRUE",
        f"{TENANT_B},beta corp,,,FALSE",
        f"{TENANT_C},Gamma,,,",
    )
    assert select(GDAP, TenantFile.load(path), ["m365_gdap"]) == [ALPHA]


@pytest.mark.parametrize("option", ["include", "exclude"])
def test_select_a_name_that_matches_nothing_is_an_error(
    groups_file: TenantFile, option: str
) -> None:
    values = {"include": ["all"], "exclude": []} | {option: ["Contoso"]}
    with pytest.raises(
        TenantSelectionError,
        match=r"No tenant or group matches 'Contoso'\. Groups in .*t\.csv: pilot, patching\.",
    ):
        select(GDAP, groups_file, values["include"], values["exclude"])


def test_select_an_include_that_selects_no_tenant_is_an_error(tmp_path: Path) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},pilot", f"{TENANT_A},Alpha,,,FALSE")
    with pytest.raises(TenantSelectionError, match="--include pilot selects no tenant"):
        select([ALPHA], TenantFile.load(path), ["pilot"])


def test_select_a_name_two_tenants_share_is_an_error_unless_named_by_id(tmp_path: Path) -> None:
    twin = Tenant(TENANT_C, "Alpha")
    tenant_file = TenantFile.empty(tmp_path / "t.csv")
    with pytest.raises(
        TenantSelectionError,
        match=rf"'alpha' could mean tenant Alpha \({TENANT_A}\) or tenant Alpha \({TENANT_C}\); "
        r"name the tenant by its ID\.",
    ):
        select([ALPHA, twin], tenant_file, ["alpha"])
    assert select([ALPHA, twin], tenant_file, [TENANT_C]) == [twin]


@pytest.mark.parametrize("option", ["include", "exclude"])
def test_select_a_name_a_tenant_and_a_group_share_is_an_error_unless_named_by_id(
    tmp_path: Path, option: str
) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},Alpha", f"{TENANT_A},Alpha,,,TRUE")
    tenant_file = TenantFile.load(path)
    values = {"include": [TENANT_A], "exclude": []} | {option: ["alpha"]}
    with pytest.raises(
        TenantSelectionError,
        match=rf"'alpha' could mean tenant Alpha \({TENANT_A}\) or group column 'Alpha' of "
        r".*t\.csv; name the tenant by its ID or rename the column\.",
    ):
        select([ALPHA], tenant_file, values["include"], values["exclude"])
    assert select([ALPHA], tenant_file, [TENANT_A]) == [ALPHA]


def test_select_all_as_a_group_name_is_an_error(tmp_path: Path) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},All", f"{TENANT_A},Alpha,,,TRUE")
    with pytest.raises(
        TenantSelectionError,
        match=r"'all' could mean every GDAP tenant or group column 'All' of .*t\.csv; rename the "
        r"column\.",
    ):
        select([ALPHA], TenantFile.load(path), ["all"])


def test_select_all_as_a_tenant_name_is_an_error_unless_named_by_id(tmp_path: Path) -> None:
    named_all = Tenant(TENANT_B, "All")
    tenant_file = TenantFile.empty(tmp_path / "t.csv")
    with pytest.raises(
        TenantSelectionError,
        match=rf"'all' could mean every GDAP tenant or tenant All \({TENANT_B}\); name the "
        r"tenant by its ID\.",
    ):
        select([ALPHA, named_all], tenant_file, ["all"])
    assert select([ALPHA, named_all], tenant_file, [TENANT_B]) == [named_all]


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


def test_load_accepts_a_header_missing_m365_columns(tmp_path: Path) -> None:
    path = _write(tmp_path / "t.csv", "m365_tenant_id,patching", f"{TENANT_A},TRUE")
    tenant_file = TenantFile.load(path)
    assert tenant_file.columns == ["m365_tenant_id", "patching"]
    assert tenant_file.members("patching") == [TenantRow(TENANT_A, "", "")]


def test_load_rejects_non_utf8(tmp_path: Path) -> None:
    path = tmp_path / "t.csv"
    path.write_bytes(b"\xff\xfet\x00e\x00")
    with pytest.raises(TenantFileError, match="not UTF-8"):
        TenantFile.load(path)


def test_load_reports_an_unreadable_path(tmp_path: Path) -> None:
    with pytest.raises(TenantFileError, match="Cannot read"):
        TenantFile.load(tmp_path)  # a directory


# A row of workbook cells, as openpyxl writes them.
XlsxRow = list[str | bool | int | None]


def _write_xlsx(path: Path, *tables: tuple[str, list[XlsxRow]], totals_row: bool = False) -> Path:
    """Write a one-sheet workbook under a title in A1, with a table per (top-left cell, rows).

    The first of each table's rows is its header; with *totals_row*, its last is a totals row.
    """
    workbook = Workbook()
    sheet = workbook.active
    assert isinstance(sheet, Worksheet)
    sheet["A1"] = "Customers"
    for number, (corner, rows) in enumerate(tables, start=1):
        top, left = coordinate_to_tuple(corner)
        for row_offset, row in enumerate(rows):
            for column_offset, value in enumerate(row):
                sheet.cell(top + row_offset, left + column_offset, value)
        last = f"{get_column_letter(left + len(rows[0]) - 1)}{top + len(rows) - 1}"
        table = Table(displayName=f"Tenants{number}", ref=f"{corner}:{last}")
        if totals_row:
            table.totalsRowCount = 1
        sheet.add_table(table)
    workbook.save(path)
    return path


XLSX_HEADER: XlsxRow = [*M365.split(","), "pilot", "cw_company_id"]


def test_load_reads_the_one_table_of_a_workbook_as_text(tmp_path: Path) -> None:
    path = _write_xlsx(
        tmp_path / "tenants.XLSX",  # any case
        (
            "B3",
            [
                XLSX_HEADER,
                [TENANT_A, "Alpha", None, "Commercial", True, 1001],
                [TENANT_B, "beta corp", "beta.example", None, False, 1002],
                [TENANT_C, "Gamma", None, None, "TRUE", None],  # TRUE as text counts too
                ["Total", None, None, None, 2, 3105],
            ],
        ),
        totals_row=True,
    )
    tenant_file = TenantFile.load(path)
    assert tenant_file.columns == XLSX_HEADER
    assert len(tenant_file.rows) == 3  # the totals row is left out
    assert tenant_file.rows[0] == {
        "m365_tenant_id": TENANT_A,
        "m365_tenant_name": "Alpha",
        "m365_domain": "",
        "m365_cloud": "Commercial",
        "pilot": "TRUE",
        "cw_company_id": "1001",
    }
    assert [row.tenant_id for row in tenant_file.members("pilot")] == [TENANT_A, TENANT_C]
    assert tenant_file.first_row == 4


def test_load_of_a_workbook_numbers_rows_as_the_sheet_does(tmp_path: Path) -> None:
    path = _write_xlsx(
        tmp_path / "t.xlsx",
        (
            "A3",
            [
                XLSX_HEADER,
                [TENANT_A, "Alpha", None, None, True, None],
                [TENANT_B, "beta corp", None, None, "maybe", None],
            ],
        ),
    )
    with pytest.raises(TenantFileError, match="row 5: 'maybe' in column 'pilot'"):
        TenantFile.load(path).members("pilot")


@pytest.mark.parametrize("corners", [(), ("A3", "J3")])
def test_load_of_a_workbook_needs_exactly_one_table(
    tmp_path: Path, corners: tuple[str, ...]
) -> None:
    rows: list[XlsxRow] = [XLSX_HEADER, [TENANT_A]]
    path = _write_xlsx(tmp_path / "t.xlsx", *((corner, rows) for corner in corners))
    with pytest.raises(
        TenantFileError,
        match=f"has {len(corners)} tables; a tenants workbook must have exactly one",
    ):
        TenantFile.load(path)


def test_load_reports_an_xlsx_that_is_not_a_workbook(tmp_path: Path) -> None:
    path = _write(tmp_path / "t.xlsx", M365)
    with pytest.raises(TenantFileError, match="Cannot read"):
        TenantFile.load(path)


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


def test_the_gdap_column_is_a_group(tmp_path: Path) -> None:
    path = _write(
        tmp_path / "t.csv",
        f"{M365},m365_gdap,g",
        f"{TENANT_A},Alpha,,,TRUE,TRUE",
        f"{TENANT_B},beta corp,,,,TRUE",
    )
    tenant_file = TenantFile.load(path)
    assert tenant_file.groups == ["m365_gdap", "g"]
    assert [row.tenant_id for row in tenant_file.members("M365_GDAP")] == [TENANT_A]


def test_members_reads_the_optional_gdap_column(tmp_path: Path) -> None:
    # Only TRUE turns the partner sign-in on; blank, like a missing column, leaves it off.
    path = _write(
        tmp_path / "t.csv",
        f"{M365},m365_gdap,g",
        f"{TENANT_A},Alpha,,, true ,TRUE",
        f"{TENANT_B},beta corp,,,FALSE,TRUE",
        f"{TENANT_C},Gamma,,,,TRUE",
    )
    assert [row.gdap for row in TenantFile.load(path).members("g")] == [True, False, False]


@pytest.mark.parametrize("cell", ["no", "0", "maybe"])
def test_members_rejects_a_gdap_cell_that_is_not_true_or_false(tmp_path: Path, cell: str) -> None:
    path = _write(tmp_path / "t.csv", f"{M365},m365_gdap,g", f"{TENANT_A},Alpha,,,{cell},TRUE")
    with pytest.raises(
        TenantFileError,
        match=f"row 2: '{cell}' in column 'm365_gdap' is not TRUE, FALSE or blank",
    ):
        TenantFile.load(path).members("g")


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
    # The file lacked m365_gdap: it is TRUE for the GDAP tenants only.
    assert [row["m365_gdap"] for row in tenant_file.rows[:4]] == ["TRUE", "TRUE", "", ""]
    assert tenant_file.rows[-1] == {
        "m365_tenant_id": DELTA.tenant_id,
        "m365_tenant_name": "Delta",
        "m365_domain": "",
        "m365_cloud": "USGov",
        "m365_username": "",
        "m365_gdap": "TRUE",
        "m365_pw_url": "",
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
        f"{M365},m365_username,m365_gdap,m365_pw_url,g",
        f"{TENANT_A},Alpha,,Commercial,,,,TRUE",  # a blank m365_gdap is left blank
        f"{TENANT_C},Gamma,,USGov,,,,TRUE",
        ",Acme,,,,,,TRUE",
    )
    tenant_file = TenantFile.load(path)
    before = [dict(row) for row in tenant_file.rows]
    report = tenant_file.sync([ALPHA], never)
    assert report.changed is False
    assert tenant_file.rows == before


def test_sync_fills_the_blank_cloud_of_every_row_of_a_tenant_looking_it_up_once(
    tmp_path: Path,
) -> None:
    looked_up: list[str] = []

    def lookup(tenant_id: str) -> str:
        looked_up.append(tenant_id)
        return "USGov"

    path = _write(
        tmp_path / "t.csv",
        M365,
        f"{TENANT_C},Gamma,,",
        f"{TENANT_A},Alpha,,Commercial",
        f"{TENANT_C},Gamma,,",
    )
    tenant_file = TenantFile.load(path)
    report = tenant_file.sync([ALPHA], lookup)
    assert (report.clouds_filled, looked_up) == ([GAMMA], [TENANT_C])
    assert [row["m365_cloud"] for row in tenant_file.rows] == ["USGov", "Commercial", "USGov"]
    # Gamma's rows still agree.
    assert tenant_file.details(GAMMA, default_cloud="Commercial").cloud == "USGov"


def test_save_round_trips_other_tools_data_and_appended_rows(csv_path: Path) -> None:
    tenant_file = TenantFile.load(csv_path)
    tenant_file.sync([ALPHA, BETA, DELTA], commercial)
    tenant_file.save()
    # The file predates m365_username, m365_gdap and m365_pw_url, so they are added after
    # m365_cloud.
    assert csv_path.read_text(encoding="utf-8-sig").splitlines() == [
        f"{M365},m365_username,m365_gdap,m365_pw_url,patching,pilot,cw_company_id,notes",
        f"{TENANT_A},Alpha,alpha.example,Commercial,,TRUE,,TRUE,FALSE,1001,first one",
        f"{TENANT_B},beta corp,beta.example,Commercial,,TRUE,,TRUE,TRUE,1002,",
        f"{TENANT_C},Gamma,,Commercial,,,,FALSE,,1003,gone from GDAP",
        ",,,,,,,TRUE,TRUE,1004,no M365 tenant",
        f"{DELTA.tenant_id},Delta,,Commercial,,TRUE,,,,,",
    ]
    assert csv_path.read_bytes().startswith(b"\xef\xbb\xbf")  # BOM for spreadsheets
    assert [p.name for p in csv_path.parent.iterdir()] == [csv_path.name]  # no temp file left


def test_sync_adds_missing_columns_after_the_last_m365_column(tmp_path: Path) -> None:
    # The m365_ columns out of order, among another tool's, and without m365_username,
    # m365_gdap or m365_pw_url.
    path = _write(
        tmp_path / "t.csv",
        "my_name,m365_tenant_name,m365_domain,m365_tenant_id,m365_cloud,exclude",
        f"Alpha Inc,Alpha,,{TENANT_A},Commercial,TRUE",
    )
    tenant_file = TenantFile.load(path)
    report = tenant_file.sync([ALPHA], commercial)
    assert (report.columns_added, report.added, report.changed) == (
        ["m365_username", "m365_gdap", "m365_pw_url"],
        [],
        True,
    )
    tenant_file.save()
    assert path.read_text(encoding="utf-8-sig").splitlines() == [
        "my_name,m365_tenant_name,m365_domain,m365_tenant_id,m365_cloud,m365_username,"
        "m365_gdap,m365_pw_url,exclude",
        f"Alpha Inc,Alpha,,{TENANT_A},Commercial,,TRUE,,TRUE",
    ]


def test_sync_adds_every_m365_column_at_the_end_of_a_file_without_them(tmp_path: Path) -> None:
    path = _write(tmp_path / "t.csv", "my_name,pilot", "Acme,TRUE")
    tenant_file = TenantFile.load(path)
    report = tenant_file.sync([ALPHA], commercial)
    assert (report.columns_added, report.added) == (list(OWN_COLUMNS), [ALPHA])
    tenant_file.save()
    assert path.read_text(encoding="utf-8-sig").splitlines() == [
        f"my_name,pilot,{','.join(OWN_COLUMNS)}",
        "Acme,TRUE,,,,,,,",
        f",,{TENANT_A},Alpha,,Commercial,,TRUE,",
    ]


def test_save_creates_the_file_and_its_directory(tmp_path: Path) -> None:
    path = tmp_path / "deeper" / "tenants.csv"
    tenant_file = TenantFile.empty(path)
    tenant_file.sync([ALPHA], commercial)
    tenant_file.save()
    assert path.read_text(encoding="utf-8-sig").splitlines() == [
        f"{M365},m365_username,m365_gdap,m365_pw_url",
        f"{TENANT_A},Alpha,,Commercial,,TRUE,",
    ]


def test_save_creates_a_workbook_with_one_table(tmp_path: Path) -> None:
    path = tmp_path / "deeper" / "tenants.xlsx"
    tenant_file = TenantFile.empty(path)
    tenant_file.sync([ALPHA], commercial)
    tenant_file.save()
    sheet = load_workbook(path)["Tenants"]
    assert sheet.tables["Tenants"].ref == "A1:G2"
    assert [cell.value for cell in sheet[2]] == [
        TENANT_A,
        "Alpha",
        None,
        "Commercial",
        None,
        True,  # Excel's own TRUE
        None,
    ]
    assert TenantFile.load(path).rows == tenant_file.rows
    assert [p.name for p in path.parent.iterdir()] == [path.name]  # no temp file left


def _shared_workbook(path: Path, *, totals_row: bool = False) -> Path:
    """A workbook another tool keeps: a table at B3 with a formula, lacking some m365_ columns."""
    totals: list[XlsxRow] = [["Total", None, None, None, None]] if totals_row else []
    return _write_xlsx(
        path,
        (
            "B3",
            [
                ["m365_tenant_id", "m365_tenant_name", "m365_cloud", "pilot", "notes"],
                [TENANT_A, "Alpha", None, True, "=1+1"],
                [TENANT_C, "Gamma", "USGov", False, 1003],
                *totals,
            ],
        ),
        totals_row=totals_row,
    )


def test_save_writes_into_a_workbook_only_what_changed(tmp_path: Path) -> None:
    path = _shared_workbook(tmp_path / "t.xlsx")
    tenant_file = TenantFile.load(path)
    report = tenant_file.sync([ALPHA, BETA], commercial)
    assert (report.added, report.clouds_filled) == ([BETA], [ALPHA])
    tenant_file.save()
    workbook = load_workbook(path)
    sheet = workbook.active
    assert isinstance(sheet, Worksheet)
    # New columns go right of the table, the new row below it, and the table grows over both.
    table = sheet.tables["Tenants1"]
    assert table.ref == "B3:J6"
    assert [column.name for column in table.tableColumns] == [
        "m365_tenant_id",
        "m365_tenant_name",
        "m365_cloud",
        "pilot",
        "notes",
        "m365_domain",
        "m365_username",
        "m365_gdap",
        "m365_pw_url",
    ]
    assert [[cell.value for cell in row] for row in sheet["B3:J6"]] == [
        [column.name for column in table.tableColumns],
        [TENANT_A, "Alpha", "Commercial", True, "=1+1", None, None, True, None],
        [TENANT_C, "Gamma", "USGov", False, 1003, None, None, None, None],
        [TENANT_B, "beta corp", "Commercial", None, None, None, None, True, None],
    ]
    assert sheet["A1"].value == "Customers"  # outside the table, left alone


@pytest.mark.parametrize("in_the_way", ["B6", "G4"])  # below the table, and right of it
def test_save_will_not_grow_a_workbook_table_over_cells_in_use(
    tmp_path: Path, in_the_way: str
) -> None:
    path = _shared_workbook(tmp_path / "t.xlsx")
    workbook = load_workbook(path)
    workbook["Sheet"][in_the_way] = "mine"
    workbook.save(path)
    before = path.read_bytes()
    tenant_file = TenantFile.load(path)
    tenant_file.sync([ALPHA, BETA], commercial)
    with pytest.raises(TenantFileError, match="the cells it would grow into"):
        tenant_file.save()
    assert path.read_bytes() == before


def test_save_will_not_grow_a_workbook_table_with_a_totals_row(tmp_path: Path) -> None:
    path = _shared_workbook(tmp_path / "t.xlsx", totals_row=True)
    tenant_file = TenantFile.load(path)
    tenant_file.sync([ALPHA, BETA], commercial)
    with pytest.raises(TenantFileError, match="it has a totals row"):
        tenant_file.save()


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


def test_details_take_the_cloud_username_and_gdap_from_the_row(tmp_path: Path) -> None:
    path = _write(
        tmp_path / "t.csv",
        f"{M365},m365_username,m365_gdap",
        f"{TENANT_A},Alpha,,,,TRUE",
        f"{TENANT_C},Gamma,,usgov,admin@gamma.example,",
    )
    tenant_file = TenantFile.load(path)
    # A blank m365_gdap is FALSE.
    assert tenant_file.details(GAMMA, default_cloud="Commercial") == Tenant(
        TENANT_C, "Gamma", "USGov", "admin@gamma.example", gdap=False
    )
    # A blank cloud, or no row at all, is the partner's cloud, and no username.
    assert tenant_file.details(ALPHA, default_cloud="Commercial") == Tenant(
        TENANT_A, "Alpha", "Commercial"
    )
    assert tenant_file.details(BETA, default_cloud="Commercial") == Tenant(
        TENANT_B, "beta corp", "Commercial"
    )


def test_details_take_the_password_page_from_the_row(tmp_path: Path) -> None:
    path = _write(
        tmp_path / "t.csv",
        f"{M365},m365_username,m365_pw_url",
        f"{TENANT_A},Alpha,,,,",
        f"{TENANT_C},Gamma,,USGov,admin@gamma.example, HTTPS://passwords.example/gamma ",
    )
    tenant_file = TenantFile.load(path)
    assert tenant_file.details(GAMMA, default_cloud="Commercial").pw_url == (
        "HTTPS://passwords.example/gamma"
    )
    # A blank cell is no password page.
    assert tenant_file.details(ALPHA, default_cloud="Commercial").pw_url == ""


@pytest.mark.parametrize(
    "pw_url",
    [r"C:\tools\setup.exe", "file:///C:/tools/setup.exe", "passwords.example/gamma", "https://"],
    ids=["windows-path", "file-url", "no-scheme", "no-host"],
)
def test_a_password_page_that_is_not_a_web_address_is_an_error(tmp_path: Path, pw_url: str) -> None:
    # Opened the way a web address is, a path would run the program it names.
    path = _write(
        tmp_path / "t.csv", f"{M365},m365_pw_url,g", f"{TENANT_C},Gamma,,USGov,{pw_url},TRUE"
    )
    tenant_file = TenantFile.load(path)
    error = "row 2: .* in column 'm365_pw_url' is not a web address"
    with pytest.raises(TenantFileError, match=error):
        tenant_file.details(GAMMA, default_cloud="Commercial")
    with pytest.raises(TenantFileError, match=error):
        tenant_file.members("g")


def test_details_reject_an_unknown_cloud_naming_the_row(tmp_path: Path) -> None:
    path = _write(
        tmp_path / "t.csv", M365, f"{TENANT_A},Alpha,,Commercial", f"{TENANT_C},Gamma,,Mars"
    )
    with pytest.raises(TenantFileError, match="row 3: 'Mars' in column 'm365_cloud'"):
        TenantFile.load(path).details(GAMMA, default_cloud="Commercial")


def test_details_accept_a_tenant_whose_rows_agree(tmp_path: Path) -> None:
    path = _write(
        tmp_path / "t.csv",
        f"{M365},m365_username,m365_gdap,pilot",
        f"{TENANT_C},Gamma,gamma.example,USGov,admin@gamma.example,TRUE,TRUE",
        # Case and surrounding spaces aside, the m365_ columns match; a group may differ.
        f"{TENANT_C.upper()}, GAMMA ,Gamma.Example,usgov,Admin@gamma.example,true,FALSE",
    )
    assert TenantFile.load(path).details(GAMMA, default_cloud="Commercial") == Tenant(
        TENANT_C, "Gamma", "USGov", "admin@gamma.example"
    )


@pytest.mark.parametrize("column", OWN_COLUMNS[1:])  # every m365_ column but the ID
def test_details_reject_a_tenant_whose_rows_disagree_naming_both_rows(
    tmp_path: Path, column: str
) -> None:
    row = dict.fromkeys(OWN_COLUMNS, "") | {"m365_tenant_id": TENANT_C, "m365_tenant_name": "Gamma"}
    path = _write(
        tmp_path / "t.csv",
        ",".join(OWN_COLUMNS),
        ",".join(row.values()),
        f"{TENANT_A},Alpha,,,,,",
        ",".join((row | {column: "changed"}).values()),
    )
    tenant_file = TenantFile.load(path)
    error = f"rows 2 and 4 are both tenant {TENANT_C} but differ in column '{column}': .* 'changed'"
    with pytest.raises(TenantFileError, match=error):
        tenant_file.details(GAMMA, default_cloud="Commercial")
    # Only the tenants a run selects are checked.
    assert tenant_file.details(ALPHA, default_cloud="Commercial").cloud == "Commercial"
