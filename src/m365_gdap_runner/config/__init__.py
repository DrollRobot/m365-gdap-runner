"""Configuration for m365_gdap_runner.

Public API::

    from m365_gdap_runner.config import Settings, load_settings

    settings = load_settings(profile="contoso")
    settings.api_url

Values live in a per-user ``config.toml`` (see ``paths.py`` for the
OS-specific location). This package stores no secrets: the secret-storage
machinery has been removed, so ``credential_backend`` and the other reserved
backend keys are no longer legal in the file. The config CLI creates and
edits it; no hand-editing is required.

Module map (each module's docstring carries its own removal instructions
where applicable):

- ``schema.py``  — the ``Settings`` dataclass: single source of truth for
  option names, types, defaults, secret classification, and help text.
- ``paths.py``   — config-directory resolution.
- ``file.py``    — config.toml reading and validation.
- ``resolve.py`` — the precedence engine behind :func:`load_settings`.
"""

from __future__ import annotations

from m365_gdap_runner.config.resolve import load_settings, resolve_settings
from m365_gdap_runner.config.schema import APP_NAME, CLI_NAME, ConfigError, Settings

__all__ = [
    "APP_NAME",
    "CLI_NAME",
    "ConfigError",
    "Settings",
    "load_settings",
    "resolve_settings",
]
