"""Unit tests for payload loading and validation in loader.py."""

from __future__ import annotations

import importlib.metadata
import textwrap
import uuid
from pathlib import Path
from typing import Any

import pytest

from m365_gdap_runner.loader import (
    ENTRY_POINT_GROUP,
    PayloadInvalid,
    PayloadNotFound,
    load_payload,
)

pytestmark = pytest.mark.unit

VALID = """
    SCOPES = ["User.Read.All"]

    def run(ctx):
        return {"tenant": ctx}
"""


def _write(tmp_path: Path, body: str, name: str = "payload.py") -> Path:
    path = tmp_path / name
    path.write_text(textwrap.dedent(body), encoding="utf-8")
    return path


def _fake_entry_points(
    monkeypatch: pytest.MonkeyPatch, *points: importlib.metadata.EntryPoint
) -> None:
    def entry_points(**params: Any) -> importlib.metadata.EntryPoints:
        return importlib.metadata.EntryPoints(points).select(**params)

    monkeypatch.setattr(importlib.metadata, "entry_points", entry_points)


def test_load_from_path(tmp_path: Path) -> None:
    payload = load_payload(str(_write(tmp_path, VALID)))
    assert payload.name == "payload"
    assert payload.scopes == ("User.Read.All",)
    assert payload.clients == ("graph",)
    assert callable(payload.run)


def test_clients_are_normalized_and_deduplicated(tmp_path: Path) -> None:
    body = """
        CLIENTS = ["Exchange", "graph", " exchange ", "ipps"]
        SCOPES = ["User.Read.All"]

        def run(ctx):
            return None
    """
    assert load_payload(str(_write(tmp_path, body))).clients == ("exchange", "graph", "ipps")


def test_payload_without_graph_needs_no_scopes(tmp_path: Path) -> None:
    body = """
        CLIENTS = ["exchange"]

        def run(ctx):
            return None
    """
    payload = load_payload(str(_write(tmp_path, body)))
    assert (payload.clients, payload.scopes) == (("exchange",), ())


def test_two_files_with_the_same_name_load_independently(tmp_path: Path) -> None:
    (tmp_path / "one").mkdir()
    (tmp_path / "two").mkdir()
    first = _write(tmp_path / "one", VALID)
    second = _write(tmp_path / "two", VALID.replace("User.Read.All", "Group.Read.All"))
    assert load_payload(str(first)).scopes == ("User.Read.All",)
    assert load_payload(str(second)).scopes == ("Group.Read.All",)


def test_load_from_entry_point(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    module_name = f"gdap_test_payload_{uuid.uuid4().hex}"
    _write(tmp_path, VALID, f"{module_name}.py")
    monkeypatch.syspath_prepend(str(tmp_path))
    _fake_entry_points(
        monkeypatch,
        importlib.metadata.EntryPoint(name="licensed", value=module_name, group=ENTRY_POINT_GROUP),
    )
    payload = load_payload("licensed")
    assert payload.name == "licensed"
    assert payload.scopes == ("User.Read.All",)


def test_unknown_entry_point_is_not_found(monkeypatch: pytest.MonkeyPatch) -> None:
    _fake_entry_points(monkeypatch)
    with pytest.raises(PayloadNotFound, match=r"neither a \.py file nor a payload"):
        load_payload("licensed")


def test_entry_point_must_reference_a_module(monkeypatch: pytest.MonkeyPatch) -> None:
    _fake_entry_points(
        monkeypatch,
        importlib.metadata.EntryPoint(name="bad", value="json:dumps", group=ENTRY_POINT_GROUP),
    )
    with pytest.raises(PayloadInvalid, match="must reference a module"):
        load_payload("bad")


def test_missing_file_is_not_found(tmp_path: Path) -> None:
    with pytest.raises(PayloadNotFound, match="not found"):
        load_payload(str(tmp_path / "missing.py"))


def test_run_with_optional_parameters_is_accepted(tmp_path: Path) -> None:
    body = 'SCOPES = ("User.Read.All",)\n\ndef run(ctx, verbose=False):\n    return None\n'
    assert load_payload(str(_write(tmp_path, body))).scopes == ("User.Read.All",)


@pytest.mark.parametrize(
    ("body", "message"),
    [
        ('SCOPES = ["User.Read.All"]\n', "callable run"),
        ('SCOPES = ["User.Read.All"]\nrun = 1\n', "callable run"),
        ('SCOPES = ["User.Read.All"]\ndef run():\n    pass\n', "one positional argument"),
        ('SCOPES = ["User.Read.All"]\ndef run(ctx, other):\n    pass\n', "one positional argument"),
        ("def run(ctx):\n    pass\n", "must define SCOPES"),
        ("SCOPES = []\ndef run(ctx):\n    pass\n", "must define SCOPES"),
        ('SCOPES = "User.Read.All"\ndef run(ctx):\n    pass\n', "must define SCOPES"),
        ('SCOPES = ["User.Read.All", 1]\ndef run(ctx):\n    pass\n', "non-empty string"),
        ('SCOPES = [" "]\ndef run(ctx):\n    pass\n', "non-empty string"),
        (
            'SCOPES = ["https://graph.microsoft.com/.default"]\ndef run(ctx):\n    pass\n',
            "must not use",
        ),
        (
            'SCOPES = ["User.Read.All", "offline_access"]\ndef run(ctx):\n    pass\n',
            "automatically",
        ),
        ('CLIENTS = []\nSCOPES = ["User.Read.All"]\ndef run(ctx):\n    pass\n', "non-empty list"),
        (
            'CLIENTS = "graph"\nSCOPES = ["User.Read.All"]\ndef run(ctx):\n    pass\n',
            "non-empty list",
        ),
        ('CLIENTS = ["teams"]\ndef run(ctx):\n    pass\n', "unknown client 'teams'"),
        ("CLIENTS = [1]\ndef run(ctx):\n    pass\n", "unknown client 1"),
        (
            'CLIENTS = ["exchange"]\nSCOPES = ["User.Read.All"]\ndef run(ctx):\n    pass\n',
            "'graph' is not in CLIENTS",
        ),
        ('CLIENTS = ["graph", "exchange"]\ndef run(ctx):\n    pass\n', "must define SCOPES"),
        ('raise RuntimeError("boom")\n', "Importing payload .* failed: RuntimeError"),
    ],
)
def test_invalid_payloads_are_rejected(tmp_path: Path, body: str, message: str) -> None:
    with pytest.raises(PayloadInvalid, match=message):
        load_payload(str(_write(tmp_path, body)))
