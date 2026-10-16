"""Unit tests for the uv script headers the examples carry."""

from __future__ import annotations

import re
import tomllib
from pathlib import Path
from typing import Any

import pytest

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
