"""Leave a visible reminder, where a person will see it, that a token cache sits on disk.

With ``cache = "disk"`` the sign-ins are kept in an encrypted cache file that outlives the
run, so later runs need no sign-in. That file holds refresh tokens for the partner account
and every tenant's own account, and it is easy to forget in a cache directory nobody opens.
When the ``cache_reminder_path`` setting names a folder, every run puts
``M365-GDAP-RUNNER_MSAL_CACHE_DELETEME`` there, pointing at the folder that holds the cache:
on Windows an Internet Shortcut (a ``.url`` text file Explorer opens as the folder),
elsewhere a symbolic link. Either holds nothing but a path, so the reminder may sit in a
synced folder such as a Desktop. The cache itself never should.

Adapted from azure-auth's ``tests/cache_reminder.py``, which does the same for its live
tests.
"""

from __future__ import annotations

import os
from pathlib import Path

REMINDER_NAME = "M365-GDAP-RUNNER_MSAL_CACHE_DELETEME"


def reminder_path(folder: Path, *, windows: bool = os.name == "nt") -> Path:
    """Return where the reminder goes in ``folder``.

    Args:
        folder: The folder the reminder is placed in.
        windows: Name a ``.url`` file, as on Windows, instead of a symbolic link.

    Returns:
        The reminder's path.
    """
    return folder / (f"{REMINDER_NAME}.url" if windows else REMINDER_NAME)


def place_reminder(cache_file: Path, folder: Path, *, windows: bool = os.name == "nt") -> Path:
    """Put the reminder in ``folder``, pointing at the folder that holds ``cache_file``.

    An existing reminder is replaced, so it always points at the current cache. A real
    directory under the reminder's name is left alone and reported, since removing it could
    take someone's files.

    Args:
        cache_file: The token cache file; it need not exist yet.
        folder: Where the reminder goes; created when missing.
        windows: Write a ``.url`` file, as on Windows, instead of a symbolic link.

    Returns:
        The reminder's path.

    Raises:
        FileExistsError: If a directory, not a reminder, has the reminder's name.
        OSError: If the folder or the reminder cannot be written.
    """
    target = cache_file.resolve().parent
    folder.mkdir(parents=True, exist_ok=True)
    link = reminder_path(folder, windows=windows)
    if link.is_dir() and not link.is_symlink():
        raise FileExistsError(f"{link} is a directory, not a reminder; not replacing it")
    if link.is_symlink() or link.exists():
        link.unlink()
    if windows:
        link.write_text(f"[InternetShortcut]\nURL={target.as_uri()}/\n", encoding="utf-8")
    else:
        link.symlink_to(target, target_is_directory=True)
    return link
