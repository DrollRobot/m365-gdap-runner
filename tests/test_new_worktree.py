"""Unit tests for the pure helpers in scripts/new_worktree.py.

The scripts folder is not a package, so the module is imported by adding the
folder to sys.path, mirroring how the scripts import their shared _cli module.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from new_worktree import (
    find_source_workspace,
    link_or_copy,
    local_dir_links,
    parse_args,
    slug_arg,
)

pytestmark = pytest.mark.unit

# --- parse_args ----------------------------------------------------------------------


def test_parse_args_defaults_to_remote_enabled(monkeypatch: pytest.MonkeyPatch) -> None:
    """Without --no-remote the worktree forks from origin/<base>."""
    monkeypatch.setattr(sys, "argv", ["new_worktree.py", "issue-42"])
    args = parse_args()
    assert args.slug == "issue-42"
    assert args.base == "develop"
    assert args.no_remote is False


def test_parse_args_accepts_no_remote_with_a_base(monkeypatch: pytest.MonkeyPatch) -> None:
    """--no-remote forks from the local base branch given positionally."""
    monkeypatch.setattr(sys, "argv", ["new_worktree.py", "issue-42", "trunk", "--no-remote"])
    args = parse_args()
    assert args.base == "trunk"
    assert args.no_remote is True


def test_parse_args_no_remote_is_independent_of_bootstrap(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The two skip flags do not imply one another."""
    monkeypatch.setattr(sys, "argv", ["new_worktree.py", "issue-42", "--no-remote"])
    args = parse_args()
    assert args.no_remote is True
    assert args.no_bootstrap is False


def test_parse_args_rejects_a_bad_slug(monkeypatch: pytest.MonkeyPatch) -> None:
    """Slug validation still applies with the new flag present."""
    monkeypatch.setattr(sys, "argv", ["new_worktree.py", "bad slug", "--no-remote"])
    with pytest.raises(SystemExit):
        parse_args()


# --- slug_arg ------------------------------------------------------------------------


def test_slug_arg_accepts_plain_slug() -> None:
    """A simple slug passes through unchanged."""
    assert slug_arg("issue-42") == "issue-42"


def test_slug_arg_accepts_internal_slash() -> None:
    """Slashes are allowed inside the slug (e.g. fix/login)."""
    assert slug_arg("fix/login") == "fix/login"


@pytest.mark.parametrize("bad", ["a b", "..", "a..b", "/lead", "trail/", "a//b"])
def test_slug_arg_rejects_bad_slugs(bad: str) -> None:
    """Characters and shapes that would break a branch or path name are refused."""
    with pytest.raises(argparse.ArgumentTypeError):
        slug_arg(bad)


# --- find_source_workspace -----------------------------------------------------------


def test_find_source_workspace_prefers_the_first_directory(tmp_path: Path) -> None:
    """Search order decides: the worktree's own workspace wins over the repo's."""
    first = tmp_path / "wt"
    second = tmp_path / "repo"
    first.mkdir()
    second.mkdir()
    (first / "a.code-workspace").write_text("{}", encoding="utf-8")
    (second / "b.code-workspace").write_text("{}", encoding="utf-8")
    assert find_source_workspace([first, second], "target.code-workspace") == (
        first / "a.code-workspace"
    )


def test_find_source_workspace_skips_the_target_file(tmp_path: Path) -> None:
    """The file being generated is never used as its own template."""
    (tmp_path / "target.code-workspace").write_text("{}", encoding="utf-8")
    assert find_source_workspace([tmp_path], "target.code-workspace") is None


def test_find_source_workspace_none_when_empty(tmp_path: Path) -> None:
    """No workspace anywhere means the caller generates a minimal one."""
    assert find_source_workspace([tmp_path], "target.code-workspace") is None


# --- link_or_copy --------------------------------------------------------------------


def _can_symlink(tmp_path: Path) -> bool:
    """Whether this environment can create symlinks (Windows needs Developer Mode)."""
    probe = tmp_path / "probe-link"
    try:
        probe.symlink_to(tmp_path / "probe-target")
    except OSError:
        return False
    probe.unlink()
    return True


def _link_fixture(tmp_path: Path) -> tuple[Path, Path]:
    """Create a source file in a fake repo and return it with its worktree destination."""
    src = tmp_path / "repo" / ".env"
    src.parent.mkdir()
    src.write_text("A=1\n", encoding="utf-8")
    dst = tmp_path / "wt" / ".env"
    dst.parent.mkdir()
    return src, dst


def _refuse(*_args: object, **_kwargs: object) -> None:
    """Stand in for a link call the platform refuses."""
    raise OSError("not permitted")


def test_link_or_copy_prefers_a_symlink(tmp_path: Path) -> None:
    """A symlink is created wherever the platform allows one."""
    if not _can_symlink(tmp_path):
        pytest.skip("symlinks unavailable in this environment")
    src, dst = _link_fixture(tmp_path)
    link_or_copy(src, dst, ".env")
    assert dst.is_symlink()
    assert dst.read_text(encoding="utf-8") == "A=1\n"


def test_link_or_copy_hard_links_without_symlinks(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    """Without symlinks the file is hard-linked, so both paths are one file."""
    monkeypatch.setattr(Path, "symlink_to", _refuse)
    src, dst = _link_fixture(tmp_path)
    link_or_copy(src, dst, ".env")
    assert not dst.is_symlink()
    assert os.path.samefile(src, dst)
    captured = capsys.readouterr()
    assert "hard-linked .env" in captured.out + captured.err


def test_link_or_copy_copies_without_any_link(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    """With neither kind of link available the file is copied."""
    monkeypatch.setattr(Path, "symlink_to", _refuse)
    monkeypatch.setattr(os, "link", _refuse)
    src, dst = _link_fixture(tmp_path)
    link_or_copy(src, dst, ".env")
    assert not dst.is_symlink()
    assert not os.path.samefile(src, dst)
    assert dst.read_text(encoding="utf-8") == "A=1\n"
    captured = capsys.readouterr()
    assert "copied .env" in captured.out + captured.err


# --- local_dir_links -----------------------------------------------------------------


def test_local_dir_links_missing_dir_is_empty(tmp_path: Path) -> None:
    """A repo without .local/ has nothing to link."""
    assert local_dir_links(tmp_path / "repo", tmp_path / "wt") == []


def test_local_dir_links_pairs_every_file_recursively(tmp_path: Path) -> None:
    """Each file under .local/, at any depth, maps to the same path in the worktree."""
    repo = tmp_path / "repo"
    wt = tmp_path / "wt"
    (repo / ".local" / "notes").mkdir(parents=True)
    (repo / ".local" / "next_commit.md").write_text("", encoding="utf-8")
    (repo / ".local" / "notes" / "plan.md").write_text("x\n", encoding="utf-8")
    assert local_dir_links(repo, wt) == [
        (
            repo / ".local" / "next_commit.md",
            wt / ".local" / "next_commit.md",
            ".local/next_commit.md",
        ),
        (
            repo / ".local" / "notes" / "plan.md",
            wt / ".local" / "notes" / "plan.md",
            ".local/notes/plan.md",
        ),
    ]


def test_local_dir_links_skips_files_already_in_the_worktree(tmp_path: Path) -> None:
    """A file the worktree already has is left alone, never clobbered."""
    repo = tmp_path / "repo"
    wt = tmp_path / "wt"
    (repo / ".local").mkdir(parents=True)
    (wt / ".local").mkdir(parents=True)
    (repo / ".local" / "PR.md").write_text("old\n", encoding="utf-8")
    (wt / ".local" / "PR.md").write_text("mine\n", encoding="utf-8")
    assert local_dir_links(repo, wt) == []
