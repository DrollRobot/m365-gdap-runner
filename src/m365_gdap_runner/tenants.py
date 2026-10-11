"""GDAP customer tenants, the tenants file, and how the command line picks tenants.

Graph says which customer tenants exist (see ``session.py``). The tenants file says which of
them belong to which group: a CSV edited in a spreadsheet, one row per customer, with the
columns ``m365_tenant_id``, ``m365_tenant_name`` and ``m365_domain`` and one column per group
holding TRUE or FALSE. ``--group NAME`` runs the tenants marked TRUE in column ``NAME``.

The file may be shared with tools that describe the same customers on other platforms, so
the columns this package owns carry the ``m365_`` prefix, every other column is left alone
unless a run names it as a group, and a row with no ``m365_tenant_id`` is a customer without
an M365 tenant: never selected, never an error.

Tenants are matched by ID only; the name and domain columns are for the person editing the
file. A GDAP tenant with no row is appended with every other column blank
(:meth:`TenantFile.add_missing`). A row whose tenant has no GDAP relationship does no harm
until a group selects it; then :func:`group_members` raises, since the runner cannot reach
that tenant.
"""

from __future__ import annotations

import contextlib
import csv
import os
import tempfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

from m365_gdap_runner.config import Settings
from m365_gdap_runner.config.paths import config_dir

# The columns this package owns. A new file gets them in this order; other tools sharing the
# file add their own, and any column outside these may serve as a group.
ID_COLUMN = "m365_tenant_id"
NAME_COLUMN = "m365_tenant_name"
DOMAIN_COLUMN = "m365_domain"
REQUIRED_COLUMNS = (ID_COLUMN, NAME_COLUMN, DOMAIN_COLUMN)

# File name in the config directory when the ``tenants_file`` setting is empty.
TENANTS_FILE_NAME = "tenants.csv"

# Cell values that put a tenant in a group, or keep it out, compared case-insensitively:
# the spreadsheet booleans only. A blank cell, which every appended row has, keeps it out.
TRUE_VALUES = frozenset({"true"})
FALSE_VALUES = frozenset({"false", ""})


class TenantSelectionError(ValueError):
    """Raised when a selection names a tenant the runner cannot use, or selects nothing."""


class TenantFileError(ValueError):
    """Raised when the tenants file cannot be read or is not laid out as expected."""


@dataclass(frozen=True)
class Tenant:
    """A GDAP customer tenant.

    Attributes:
        tenant_id: Entra tenant ID.
        display_name: Tenant display name as set in Entra ID.
    """

    tenant_id: str
    display_name: str


@dataclass(frozen=True)
class TenantRow:
    """The M365 columns of one row of the tenants file, stripped of surrounding whitespace.

    Attributes:
        tenant_id: The ``m365_tenant_id`` cell.
        display_name: The ``m365_tenant_name`` cell; may be empty.
        domain: The ``m365_domain`` cell; may be empty.
    """

    tenant_id: str
    display_name: str
    domain: str


def tenants_file_path(settings: Settings) -> Path:
    """Return where the tenants file lives.

    Args:
        settings: Resolved settings.

    Returns:
        The ``tenants_file`` setting with ``~`` expanded, or ``tenants.csv`` in the config
        directory when the setting is empty.
    """
    if settings.tenants_file:
        return Path(settings.tenants_file).expanduser()
    return config_dir() / TENANTS_FILE_NAME


class TenantFile:
    """The tenants CSV: one row per customer, one TRUE/FALSE column per group.

    Rows keep every cell as the text the file holds, so :meth:`save` writes back what was
    read plus any appended rows. Only the group column a caller asks about is checked for
    valid values; every other column may hold anything.

    Attributes:
        path: Where the file is, or will be, on disk.
        columns: Column names in file order.
        rows: One mapping per row, column name to cell text.
    """

    def __init__(
        self, path: Path, columns: Sequence[str], rows: Sequence[dict[str, str]] = ()
    ) -> None:
        """Hold a file's contents in memory.

        Args:
            path: Where the file is, or will be, on disk.
            columns: Column names in file order; must include :data:`REQUIRED_COLUMNS`.
            rows: One mapping per row, column name to cell text.
        """
        self.path = path
        self.columns = list(columns)
        self.rows = [dict(row) for row in rows]

    @classmethod
    def empty(cls, path: Path) -> TenantFile:
        """Return a file with the required columns and no rows, not yet on disk.

        Args:
            path: Where :meth:`save` will write it.

        Returns:
            The empty file.
        """
        return cls(path, REQUIRED_COLUMNS)

    @classmethod
    def load(cls, path: Path) -> TenantFile:
        """Read the file.

        A file with nothing in it, not even a header, counts as :meth:`empty`.

        Args:
            path: The CSV to read. UTF-8, with or without a byte-order mark.

        Returns:
            The file's contents.

        Raises:
            TenantFileError: If the file cannot be read, is not UTF-8, or its header lacks a
                required column.
        """
        try:
            with path.open(newline="", encoding="utf-8-sig") as handle:
                reader = csv.DictReader(handle, restval="")
                columns = reader.fieldnames
                # A row longer than the header puts the extra cells under the key None.
                rows = [
                    {key: value for key, value in row.items() if key is not None} for row in reader
                ]
        except UnicodeDecodeError as exc:
            raise TenantFileError(f"{path} is not UTF-8 text: {exc}") from exc
        except OSError as exc:
            raise TenantFileError(f"Cannot read {path}: {exc}") from exc
        if columns is None:
            return cls.empty(path)
        missing = [column for column in REQUIRED_COLUMNS if column not in columns]
        if missing:
            raise TenantFileError(
                f"{path} lacks the column(s) {', '.join(missing)}; the header must include "
                f"{', '.join(REQUIRED_COLUMNS)}."
            )
        return cls(path, columns, rows)

    @property
    def groups(self) -> list[str]:
        """Columns that look like groups, in file order.

        A column this package does not own whose filled cells are all TRUE or FALSE. Columns
        other tools keep in a shared file are left out. This is the hint in the unknown-group
        error; ``--group`` itself may name any column this package does not own.
        """
        return [column for column in self._other_columns() if self._is_boolean(column)]

    def _other_columns(self) -> list[str]:
        """Return every named column this package does not own, in file order."""
        return [column for column in self.columns if column and column not in REQUIRED_COLUMNS]

    def _is_boolean(self, column: str) -> bool:
        """Return True if every cell of *column* is TRUE, FALSE or blank.

        Args:
            column: A column name.

        Returns:
            Whether the column could be a group.
        """
        allowed = TRUE_VALUES | FALSE_VALUES
        return all(row.get(column, "").strip().casefold() in allowed for row in self.rows)

    def _group_column(self, group: str) -> str:
        """Resolve a group name to its column, case-insensitively.

        Args:
            group: The name given on the command line.

        Returns:
            The column name as the file spells it.

        Raises:
            TenantFileError: If no column, or more than one, matches.
        """
        needle = group.strip().casefold()
        matches = [column for column in self._other_columns() if column.casefold() == needle]
        if len(matches) == 1:
            return matches[0]
        if not matches:
            known = ", ".join(self.groups) or "none"
            raise TenantFileError(f"{self.path} has no group column {group!r}. Groups: {known}.")
        raise TenantFileError(
            f"{group!r} matches {len(matches)} columns of {self.path}: {', '.join(matches)}."
        )

    def members(self, group: str) -> list[TenantRow]:
        """Return the rows marked TRUE in a group column, in file order.

        A marked row with no ``m365_tenant_id`` is a customer without an M365 tenant and is
        left out.

        Args:
            group: The group's column name, in any case.

        Returns:
            The member rows.

        Raises:
            TenantFileError: If there is no such column, or a cell in it is not TRUE, FALSE or
                blank.
        """
        column = self._group_column(group)
        chosen: list[TenantRow] = []
        for number, row in enumerate(self.rows, start=2):
            cell = row.get(column, "")
            value = cell.strip().casefold()
            if value in FALSE_VALUES:
                continue
            if value not in TRUE_VALUES:
                raise TenantFileError(
                    f"{self.path} row {number}: {cell!r} in column {column!r} is not TRUE, "
                    "FALSE or blank."
                )
            member = self._row(row)
            if member.tenant_id:
                chosen.append(member)
        return chosen

    @staticmethod
    def _row(row: dict[str, str]) -> TenantRow:
        """Pick the M365 columns out of a row.

        Args:
            row: Column name to cell text.

        Returns:
            The row's tenant ID, display name and domain, stripped.
        """
        return TenantRow(
            row.get(ID_COLUMN, "").strip(),
            row.get(NAME_COLUMN, "").strip(),
            row.get(DOMAIN_COLUMN, "").strip(),
        )

    def add_missing(self, tenants: Sequence[Tenant]) -> list[Tenant]:
        """Append a row, with every other column blank, for each tenant the file does not know.

        Nothing is written; call :meth:`save`.

        Args:
            tenants: The GDAP tenants.

        Returns:
            The tenants appended, in input order.
        """
        known = {row.get(ID_COLUMN, "").strip().lower() for row in self.rows}
        added: list[Tenant] = []
        for tenant in tenants:
            key = tenant.tenant_id.lower()
            if key in known:
                continue
            row = dict.fromkeys(self.columns, "")
            row[ID_COLUMN] = tenant.tenant_id
            row[NAME_COLUMN] = tenant.display_name
            self.rows.append(row)
            known.add(key)
            added.append(tenant)
        return added

    def save(self) -> None:
        """Write the file, replacing it in one step so a failure leaves the old one intact.

        The parent directory is created if needed. The file is written with a UTF-8
        byte-order mark so spreadsheets open it correctly.

        Raises:
            OSError: If the file or its directory cannot be written, for example while a
                spreadsheet holds the file open.
        """
        self.path.parent.mkdir(parents=True, exist_ok=True)
        descriptor, temp = tempfile.mkstemp(
            dir=self.path.parent, prefix=self.path.name + ".", suffix=".tmp"
        )
        try:
            with os.fdopen(descriptor, "w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.DictWriter(handle, fieldnames=self.columns, extrasaction="ignore")
                writer.writeheader()
                writer.writerows(self.rows)
            os.replace(temp, self.path)
        except BaseException:
            with contextlib.suppress(OSError):
                os.unlink(temp)
            raise


def _match(tenants: Sequence[Tenant], spec: str) -> Tenant:
    """Resolve one tenant by ID or display name (case-insensitive).

    Args:
        tenants: Candidate tenants.
        spec: Tenant ID or display name.

    Returns:
        The single matching tenant. An ID match wins over a name match.

    Raises:
        TenantSelectionError: If nothing matches, or the name matches several tenants.
    """
    needle = spec.strip().casefold()
    for tenant in tenants:
        if tenant.tenant_id.casefold() == needle:
            return tenant
    by_name = [t for t in tenants if t.display_name.casefold() == needle]
    if not by_name:
        raise TenantSelectionError(f"No GDAP tenant matches {spec!r}.")
    if len(by_name) > 1:
        ids = ", ".join(t.tenant_id for t in by_name)
        raise TenantSelectionError(
            f"{spec!r} matches {len(by_name)} tenants by display name; use a tenant ID ({ids})."
        )
    return by_name[0]


def select(
    tenants: Sequence[Tenant],
    include: Sequence[str] | None = None,
    exclude: Sequence[str] | None = None,
) -> list[Tenant]:
    """Filter tenants by ID or display name, preserving order.

    Every ``include`` and ``exclude`` value must match exactly one tenant: a
    typo is an error, never a silent no-op.

    Args:
        tenants: All enumerated tenants.
        include: Keep only these tenants. None or empty keeps all.
        exclude: Drop these tenants.

    Returns:
        The selected tenants, in input order.

    Raises:
        TenantSelectionError: If a value matches no tenant or several.
    """
    chosen = list(tenants)
    if include:
        wanted = {_match(tenants, spec).tenant_id for spec in include}
        chosen = [t for t in chosen if t.tenant_id in wanted]
    if exclude:
        unwanted = {_match(tenants, spec).tenant_id for spec in exclude}
        chosen = [t for t in chosen if t.tenant_id not in unwanted]
    return chosen


def group_members(
    tenants: Sequence[Tenant], tenant_file: TenantFile, groups: Sequence[str]
) -> list[Tenant]:
    """Resolve the members of *groups* to GDAP tenants.

    Args:
        tenants: All enumerated tenants.
        tenant_file: The tenants file.
        groups: Group names; a tenant marked in any of them is selected.

    Returns:
        The selected tenants, each once, in *tenants* order.

    Raises:
        TenantFileError: If a group has no column, or a cell in it is not TRUE, FALSE or blank.
        TenantSelectionError: If a member has no GDAP relationship with the partner tenant.
    """
    by_id = {t.tenant_id.lower(): t for t in tenants}
    wanted: set[str] = set()
    for group in groups:
        for row in tenant_file.members(group):
            key = row.tenant_id.lower()
            if key not in by_id:
                name = row.display_name or row.tenant_id
                raise TenantSelectionError(
                    f"{name!r} ({row.tenant_id}) is in group {group!r} but has no GDAP "
                    "relationship with the partner tenant."
                )
            wanted.add(key)
    return [t for t in tenants if t.tenant_id.lower() in wanted]
