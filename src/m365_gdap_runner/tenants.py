"""GDAP customer tenants, the tenants file, and how the command line picks tenants.

Graph says which customer tenants have a GDAP relationship (see ``session.py``). The tenants
file says which tenants belong to which group: a CSV edited in a spreadsheet, one row per
customer, with the columns ``m365_tenant_id``, ``m365_tenant_name``, ``m365_domain``,
``m365_cloud`` and ``m365_username``, and one column per group holding TRUE or FALSE.
``--group NAME`` runs the tenants marked TRUE in column ``NAME``.

The file may be shared with tools that describe the same customers on other platforms, so
the columns this package owns carry the ``m365_`` prefix, every other column is left alone
unless a run names it as a group, and a row with no ``m365_tenant_id`` is a customer without
an M365 tenant: never selected, never an error.

Tenants are matched by ID only; the name and domain columns are for the person editing the
file. A column the file lacks is added, a GDAP tenant with no row is appended, and a blank
cloud is filled in by OpenID Connect discovery (:meth:`TenantFile.sync`). A row whose tenant
has no GDAP relationship, or lives in another Microsoft cloud, can still be selected: the
runner signs in to it directly, as the account in ``m365_username``. A blank cell means no
direct sign-in.
"""

from __future__ import annotations

import contextlib
import csv
import os
import tempfile
from collections.abc import Callable, Sequence
from dataclasses import dataclass, replace
from pathlib import Path

from azure_auth.clouds import CLOUDS, get_cloud

from m365_gdap_runner.config import Settings
from m365_gdap_runner.config.paths import config_dir

# The columns this package owns. A new file gets them in this order, and an existing one gets
# any it lacks; other tools sharing the file add their own, and any column outside these may
# serve as a group.
ID_COLUMN = "m365_tenant_id"
NAME_COLUMN = "m365_tenant_name"
DOMAIN_COLUMN = "m365_domain"
CLOUD_COLUMN = "m365_cloud"
USERNAME_COLUMN = "m365_username"
OWN_COLUMNS = (ID_COLUMN, NAME_COLUMN, DOMAIN_COLUMN, CLOUD_COLUMN, USERNAME_COLUMN)

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
    """A customer tenant: a GDAP customer, or a row of the tenants file.

    Attributes:
        tenant_id: Entra tenant ID.
        display_name: Tenant display name as set in Entra ID, or as the tenants file names it.
        cloud: The cloud it lives in, as azure-auth spells it. Filled in for a tenant a run
            selects; empty for one only listed.
        username: The account to sign in with when the partner sign-in cannot reach it; empty
            for no direct sign-in.
    """

    tenant_id: str
    display_name: str
    cloud: str = ""
    username: str = ""


@dataclass(frozen=True)
class TenantRow:
    """The M365 columns of one row of the tenants file, stripped of surrounding whitespace.

    Attributes:
        tenant_id: The ``m365_tenant_id`` cell.
        display_name: The ``m365_tenant_name`` cell; may be empty.
        cloud: The ``m365_cloud`` cell as azure-auth spells the cloud's name, or empty.
        username: The ``m365_username`` cell, or empty when blank or the column is absent.
    """

    tenant_id: str
    display_name: str
    cloud: str
    username: str = ""


@dataclass(frozen=True)
class SyncReport:
    """What :meth:`TenantFile.sync` changed.

    Attributes:
        columns_added: Columns this package owns that the file lacked, in the order added.
        added: Tenants that got a new row.
        clouds_filled: Tenants whose existing row got its ``m365_cloud`` filled in.
    """

    columns_added: list[str]
    added: list[Tenant]
    clouds_filled: list[Tenant]

    @property
    def changed(self) -> bool:
        """Whether the file differs from what was loaded."""
        return bool(self.columns_added or self.added or self.clouds_filled)


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
    read plus any added columns and rows. Only the group column a caller asks about, and the
    cloud of the rows marked in it, are checked; every other cell may hold anything.

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
            columns: Column names in file order. :meth:`sync` adds any of
                :data:`OWN_COLUMNS` missing.
            rows: One mapping per row, column name to cell text.
        """
        self.path = path
        self.columns = list(columns)
        self.rows = [dict(row) for row in rows]

    @classmethod
    def empty(cls, path: Path) -> TenantFile:
        """Return a file with this package's columns and no rows, not yet on disk.

        Args:
            path: Where :meth:`save` will write it.

        Returns:
            The empty file.
        """
        return cls(path, OWN_COLUMNS)

    @classmethod
    def load(cls, path: Path) -> TenantFile:
        """Read the file.

        A file with nothing in it, not even a header, counts as :meth:`empty`. A header that
        lacks some of this package's columns is fine; :meth:`sync` adds them.

        Args:
            path: The CSV to read. UTF-8, with or without a byte-order mark.

        Returns:
            The file's contents.

        Raises:
            TenantFileError: If the file cannot be read or is not UTF-8.
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
        return [column for column in self.columns if column and column not in OWN_COLUMNS]

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
            TenantFileError: If there is no such column, a cell in it is not TRUE, FALSE or
                blank, or a marked row's ``m365_cloud`` names no known cloud.
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
            member = self._row(row, number)
            if member.tenant_id:
                chosen.append(member)
        return chosen

    def _row(self, row: dict[str, str], number: int) -> TenantRow:
        """Pick the M365 columns out of a row.

        Args:
            row: Column name to cell text.
            number: The row's number in the spreadsheet, for error messages.

        Returns:
            The row's tenant ID, display name, cloud and username, stripped; the cloud
            spelled as azure-auth names it.

        Raises:
            TenantFileError: If ``m365_cloud`` is filled but names no known cloud.
        """
        cell = row.get(CLOUD_COLUMN, "").strip()
        cloud = ""
        if cell:
            try:
                cloud = get_cloud(cell).name
            except ValueError:
                names = ", ".join(known.name for known in CLOUDS)
                raise TenantFileError(
                    f"{self.path} row {number}: {cell!r} in column {CLOUD_COLUMN!r} is not a "
                    f"cloud. Use {names}, or blank."
                ) from None
        return TenantRow(
            row.get(ID_COLUMN, "").strip(),
            row.get(NAME_COLUMN, "").strip(),
            cloud,
            row.get(USERNAME_COLUMN, "").strip(),
        )

    def extra_tenants(self, gdap: Sequence[Tenant]) -> list[Tenant]:
        """Return the tenants the file has a row for but GDAP does not list, in file order.

        Args:
            gdap: The GDAP tenants.

        Returns:
            One tenant per ID, named as the file names it, or by its ID when unnamed.
        """
        seen = {tenant.tenant_id.lower() for tenant in gdap}
        extra: list[Tenant] = []
        for row in self.rows:
            tenant_id = row.get(ID_COLUMN, "").strip()
            if tenant_id and tenant_id.lower() not in seen:
                seen.add(tenant_id.lower())
                extra.append(Tenant(tenant_id, row.get(NAME_COLUMN, "").strip() or tenant_id))
        return extra

    def details(self, tenant: Tenant, *, default_cloud: str) -> Tenant:
        """Return *tenant* with the cloud and sign-in account its row gives.

        Args:
            tenant: A selected tenant.
            default_cloud: The cloud when the row has none, or there is no row: the partner's.

        Returns:
            The tenant, with ``cloud`` always filled in.

        Raises:
            TenantFileError: If the row's ``m365_cloud`` names no known cloud.
        """
        key = tenant.tenant_id.lower()
        for number, row in enumerate(self.rows, start=2):
            if row.get(ID_COLUMN, "").strip().lower() == key:
                found = self._row(row, number)
                return replace(tenant, cloud=found.cloud or default_cloud, username=found.username)
        return replace(tenant, cloud=default_cloud)

    def _add_missing_columns(self) -> list[str]:
        """Add the columns of :data:`OWN_COLUMNS` the file lacks, blank in every row.

        They go right after the last column this package owns, so its columns stay together,
        or at the end when the file has none of them.

        Returns:
            The columns added, in :data:`OWN_COLUMNS` order.
        """
        missing = [column for column in OWN_COLUMNS if column not in self.columns]
        if not missing:
            return []
        own = [index for index, column in enumerate(self.columns) if column in OWN_COLUMNS]
        at = own[-1] + 1 if own else len(self.columns)
        self.columns[at:at] = missing
        for row in self.rows:
            for column in missing:
                row.setdefault(column, "")
        return missing

    def sync(self, tenants: Sequence[Tenant], cloud_of: Callable[[str], str]) -> SyncReport:
        """Add missing columns, append a row for each tenant the file lacks, and fill in clouds.

        A column of :data:`OWN_COLUMNS` the file lacks is added, blank in every row. An
        appended row gets the tenant's ID and name, its cloud from *cloud_of*, and every
        other column blank. Any other row with a tenant ID and a blank ``m365_cloud``, a GDAP
        tenant or not, gets it from *cloud_of* too, so a file another tool wrote, or that
        predates the column, fills in over time. Nothing is written; call :meth:`save`.

        Args:
            tenants: The GDAP tenants.
            cloud_of: Returns the name of the cloud a tenant ID lives in, or ``""`` when it
                cannot be found. Called once per tenant that needs it.

        Returns:
            What changed.
        """
        columns_added = self._add_missing_columns()
        by_id: dict[str, dict[str, str]] = {}
        for existing in self.rows:
            by_id.setdefault(existing.get(ID_COLUMN, "").strip().lower(), existing)
        known = dict(by_id)
        added: list[Tenant] = []
        for tenant in tenants:
            key = tenant.tenant_id.lower()
            if key not in by_id:
                row = dict.fromkeys(self.columns, "")
                row[ID_COLUMN] = tenant.tenant_id
                row[NAME_COLUMN] = tenant.display_name
                row[CLOUD_COLUMN] = cloud_of(tenant.tenant_id)
                self.rows.append(row)
                by_id[key] = row
                added.append(tenant)
        gdap = {tenant.tenant_id.lower(): tenant for tenant in tenants}
        filled: list[Tenant] = []
        for key, row in known.items():
            if not key or row.get(CLOUD_COLUMN, "").strip():
                continue
            tenant_id = row[ID_COLUMN].strip()
            cloud = cloud_of(tenant_id)
            if cloud:
                row[CLOUD_COLUMN] = cloud
                named = row.get(NAME_COLUMN, "").strip() or tenant_id
                filled.append(gdap.get(key) or Tenant(tenant_id, named))
        return SyncReport(columns_added, added, filled)

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
        raise TenantSelectionError(
            f"No tenant matches {spec!r}: it is neither a GDAP customer nor in the tenants file."
        )
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
