"""Unit tests for the examples: the uv script headers they carry, and their helpers."""

from __future__ import annotations

import argparse
import re
import tomllib
from pathlib import Path
from typing import Any

import pytest

from tests._gdap_fakes import load_script

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[1]
EXAMPLES = sorted((ROOT / "examples").glob("*.py"))

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


def test_service_principal_pattern_is_case_insensitive_and_rejects_a_bad_regex() -> None:
    module = load_script(ROOT / "examples" / "service_principal_enabled.py")
    assert module.pattern("BACKUP$").search("ConnectWise Cloud Backup")
    with pytest.raises(argparse.ArgumentTypeError, match=r"not a valid regular expression: '\('"):
        module.pattern("(")


def test_service_principal_matches_display_name_or_app_id(monkeypatch: pytest.MonkeyPatch) -> None:
    module = load_script(ROOT / "examples" / "service_principal_enabled.py")
    monkeypatch.setattr(module, "PATTERNS", [module.pattern("^backup"), module.pattern("^cw00")])
    assert module.is_match({"displayName": "Backup Tool", "appId": "x"})
    assert module.is_match({"displayName": "Other", "appId": "cw000000-0000"})
    assert not module.is_match({"displayName": "Cloud Backup", "appId": "x"})
    assert not module.is_match({})
