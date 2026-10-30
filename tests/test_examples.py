"""Unit tests for the examples: the uv script headers they carry, and their helpers."""

from __future__ import annotations

import argparse
import re
import runpy
import sys
import tomllib
from pathlib import Path
from typing import Any

import pytest

import m365_gdap_runner
from tests._gdap_fakes import load_script

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[1]
EXAMPLES = sorted((ROOT / "examples").glob("*.py"))
SERVICE_PRINCIPAL_ENABLED = ROOT / "examples" / "service_principal_enabled.py"
MESSAGE_TRACE = ROOT / "examples" / "exchange_message_trace.py"

# The block a script's header lives in, as PEP 723 gives the pattern.
BLOCK = re.compile(r"(?m)^# /// (?P<type>[a-zA-Z0-9-]+)$\s(?P<content>(^#(| .*)$\s)+)^# ///$")

RUNNER = re.compile(
    r"m365-gdap-runner @ git\+https://github\.com/DrollRobot/m365-gdap-runner\.git@v\d+\.\d+\.\d+"
)


def _header(path: Path) -> dict[str, Any]:
    blocks = [m for m in BLOCK.finditer(path.read_text(encoding="utf-8")) if m["type"] == "script"]
    assert len(blocks) == 1, f"{path.name} needs exactly one '# /// script' block"
    lines = blocks[0]["content"].splitlines(keepends=True)
    return tomllib.loads("".join(line[2:] if line.startswith("# ") else line[1:] for line in lines))


@pytest.mark.parametrize("path", EXAMPLES, ids=lambda path: path.name)
def test_example_header_installs_the_runner_from_a_release(path: Path) -> None:
    header = _header(path)
    pyproject = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert header["requires-python"] == pyproject["project"]["requires-python"]
    runner = [dep for dep in header["dependencies"] if dep.startswith("m365-gdap-runner")]
    assert len(runner) == 1
    assert RUNNER.fullmatch(runner[0]), runner[0]


@pytest.mark.parametrize(
    "path",
    [path for path in EXAMPLES if "OUT_DIR" in path.read_text(encoding="utf-8")],
    ids=lambda path: path.name,
)
def test_example_that_writes_files_takes_its_folder_as_o_or_output(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str], path: Path
) -> None:
    monkeypatch.setattr(sys, "argv", [path.name, "--help"])
    with pytest.raises(SystemExit) as exc:
        runpy.run_path(str(path), run_name="__main__")
    assert exc.value.code == 0
    assert "-o DIR, --output DIR" in capsys.readouterr().out


@pytest.mark.parametrize(
    "path",
    [path for path in EXAMPLES if "OUT_DIR" in path.read_text(encoding="utf-8")],
    ids=lambda path: path.name,
)
def test_example_that_writes_files_defaults_to_the_current_folder(path: Path) -> None:
    assert Path() == load_script(path).OUT_DIR


def test_service_principal_pattern_is_case_insensitive_and_rejects_a_bad_regex() -> None:
    module = load_script(ROOT / "examples" / "service_principal_enabled.py")
    assert module.pattern("BACKUP$").search("Contoso Cloud Backup")
    with pytest.raises(argparse.ArgumentTypeError, match=r"not a valid regular expression: '\('"):
        module.pattern("(")


def test_service_principal_matches_display_name_or_app_id(monkeypatch: pytest.MonkeyPatch) -> None:
    module = load_script(ROOT / "examples" / "service_principal_enabled.py")
    monkeypatch.setattr(module, "PATTERNS", [module.pattern("^backup"), module.pattern("^cw00")])
    assert module.is_match({"displayName": "Backup Tool", "appId": "x"})
    assert module.is_match({"displayName": "Other", "appId": "cw000000-0000"})
    assert not module.is_match({"displayName": "Cloud Backup", "appId": "x"})
    assert not module.is_match({})


@pytest.mark.parametrize(
    ("flags", "expected"),
    [
        ([], ([("Application.Read.All", "Application.ReadWrite.All")], False, False)),
        (["--disable"], (["Application.ReadWrite.All"], True, False)),
        (["--delete"], (["Application.ReadWrite.All"], False, True)),
    ],
    ids=["read", "disable", "delete"],
)
def test_service_principal_flags_set_the_scopes_and_the_change(
    monkeypatch: pytest.MonkeyPatch, flags: list[str], expected: tuple[Any, bool, bool]
) -> None:
    seen: list[tuple[Any, bool, bool]] = []

    def fake_main(name: str, args: argparse.Namespace) -> int:
        script = sys.modules[name]
        seen.append((script.SCOPES, script.DISABLE, script.DELETE))
        return 0

    monkeypatch.setattr(m365_gdap_runner, "main", fake_main)
    monkeypatch.setattr(sys, "argv", ["service_principal_enabled.py", "Backup", *flags])
    with pytest.raises(SystemExit) as exc:
        runpy.run_path(str(SERVICE_PRINCIPAL_ENABLED), run_name="__main__")
    assert exc.value.code == 0
    assert seen == [expected]


def test_service_principal_refuses_disable_with_delete(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    monkeypatch.setattr(
        sys, "argv", ["service_principal_enabled.py", "Backup", "--disable", "--delete"]
    )
    with pytest.raises(SystemExit) as exc:
        runpy.run_path(str(SERVICE_PRINCIPAL_ENABLED), run_name="__main__")
    assert exc.value.code == 2
    assert "argument --delete: not allowed with argument --disable" in capsys.readouterr().err


@pytest.mark.parametrize(
    ("flags", "size"),
    [(["--output", "x"], 5000), (["-o", "x", "--size", "100"], 100)],
    ids=["--output, default size", "-o, --size"],
)
def test_message_trace_command_line_takes_the_filter_as_given(
    monkeypatch: pytest.MonkeyPatch, flags: list[str], size: int
) -> None:
    given = '{"SenderAddress": "admin@contoso.com"}'
    seen: list[tuple[str, int, Path]] = []

    def fake_main(name: str, args: argparse.Namespace) -> int:
        script = sys.modules[name]
        seen.append((script.FILTER, script.SIZE, script.OUT_DIR))
        return 0

    monkeypatch.setattr(m365_gdap_runner, "main", fake_main)
    monkeypatch.setattr(sys, "argv", ["exchange_message_trace.py", given, *flags])
    with pytest.raises(SystemExit) as exc:
        runpy.run_path(str(MESSAGE_TRACE), run_name="__main__")
    assert exc.value.code == 0
    assert seen == [(given, size, Path("x"))]
