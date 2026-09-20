"""Module entry point: ``python -m m365_gdap_runner.config``."""

from __future__ import annotations

import sys

from m365_gdap_runner.config.cli import main

if __name__ == "__main__":
    sys.exit(main())
