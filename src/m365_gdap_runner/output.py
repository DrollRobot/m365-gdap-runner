"""Result output: per-tenant shape checks and JSON / CSV writers.

The runner calls :func:`check_result` on every tenant's return value, so a
result that cannot be written fails that tenant instead of the whole write at
the end of the run.
"""

from __future__ import annotations

import csv
import io
import json
import sys
from collections.abc import Sequence
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from m365_gdap_runner.runner import TenantResult

# Supported --format values.
FORMATS = ("json", "csv")

# Destination meaning standard output.
STDOUT = "-"

# Value types allowed in a CSV row.
_Scalar = str | int | float | bool | None


class OutputFormatError(ValueError):
    """Raised when a result cannot be written in the requested format."""


def check_result(result: object, fmt: str) -> None:
    """Verify that one tenant's result can be written as *fmt*.

    Args:
        result: The payload's return value.
        fmt: ``"json"`` or ``"csv"``.

    Raises:
        OutputFormatError: If the result is not JSON-serializable, or *fmt* is
            ``"csv"`` and the result is not a flat dict or list of flat dicts.
    """
    try:
        json.dumps(result, allow_nan=False)
    except (TypeError, ValueError) as exc:
        raise OutputFormatError(f"result is not JSON-serializable: {exc}") from exc
    if fmt == "csv":
        csv_rows(result)


def csv_rows(result: object) -> list[dict[str, object]]:
    """Return the CSV rows a result produces.

    Args:
        result: A flat dict, a list of flat dicts, or None (no rows).

    Returns:
        The rows.

    Raises:
        OutputFormatError: If the result is not a flat dict or a list of flat dicts.
    """
    if result is None:
        return []
    items = result if isinstance(result, list) else [result]
    rows: list[dict[str, object]] = []
    for item in items:
        if not isinstance(item, dict):
            raise OutputFormatError(
                f"CSV output needs a flat dict or a list of flat dicts, got {type(item).__name__}."
            )
        for key, value in item.items():
            if not isinstance(key, str) or not isinstance(value, _Scalar):
                raise OutputFormatError(
                    f"CSV output needs flat dicts; key {key!r} holds {type(value).__name__}."
                )
        rows.append(item)
    return rows


def to_records(results: Sequence[TenantResult]) -> list[dict[str, object]]:
    """Convert results to JSON-ready records.

    Args:
        results: Per-tenant results.

    Returns:
        One ``{tenant_id, display_name, status, error, result}`` dict per tenant.
    """
    return [
        {
            "tenant_id": r.tenant.tenant_id,
            "display_name": r.tenant.display_name,
            "status": r.status,
            "error": r.error,
            "result": r.result,
        }
        for r in results
    ]


def render_json(results: Sequence[TenantResult]) -> str:
    """Render results as a JSON array.

    Args:
        results: Per-tenant results.

    Returns:
        Indented JSON text ending in a newline.
    """
    return json.dumps(to_records(results), indent=2, allow_nan=False) + "\n"


def render_csv(results: Sequence[TenantResult]) -> str:
    """Render the rows of every ``ok`` result as CSV.

    Non-ok tenants are omitted; the run summary reports them. Columns are
    ``tenant_id``, ``display_name``, then every other key in first-seen order.

    Args:
        results: Per-tenant results.

    Returns:
        CSV text with a header row.

    Raises:
        OutputFormatError: If an ok result is not flat.
    """
    columns: dict[str, None] = {"tenant_id": None, "display_name": None}
    rows: list[dict[str, object]] = []
    for r in results:
        if r.status != "ok":
            continue
        for row in csv_rows(r.result):
            columns.update(dict.fromkeys(row))
            rows.append(
                row | {"tenant_id": r.tenant.tenant_id, "display_name": r.tenant.display_name}
            )
    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=list(columns), lineterminator="\n")
    writer.writeheader()
    writer.writerows(rows)
    return buffer.getvalue()


def _write_text(text: str, dest: str) -> None:
    """Write text to a file, or to stdout when *dest* is ``-``.

    Args:
        text: Content to write.
        dest: File path, or :data:`STDOUT`.
    """
    if dest == STDOUT:
        sys.stdout.write(text)
        sys.stdout.flush()
    else:
        Path(dest).write_text(text, encoding="utf-8")


def write_json(results: Sequence[TenantResult], dest: str) -> None:
    """Write results as JSON.

    Args:
        results: Per-tenant results.
        dest: File path, or ``-`` for stdout.
    """
    _write_text(render_json(results), dest)


def write_csv(results: Sequence[TenantResult], dest: str) -> None:
    """Write ok results as CSV.

    Args:
        results: Per-tenant results.
        dest: File path, or ``-`` for stdout.

    Raises:
        OutputFormatError: If an ok result is not flat.
    """
    _write_text(render_csv(results), dest)


def write_results(results: Sequence[TenantResult], dest: str, fmt: str) -> None:
    """Write results in the requested format.

    Args:
        results: Per-tenant results.
        dest: File path, or ``-`` for stdout.
        fmt: ``"json"`` or ``"csv"``.

    Raises:
        OutputFormatError: If *fmt* is ``"csv"`` and an ok result is not flat.
        ValueError: If *fmt* is not supported.
    """
    if fmt == "json":
        write_json(results, dest)
    elif fmt == "csv":
        write_csv(results, dest)
    else:
        raise ValueError(f"Unsupported output format {fmt!r}; use one of {', '.join(FORMATS)}.")
