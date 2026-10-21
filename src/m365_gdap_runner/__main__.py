"""Module entry point: ``python -m m365_gdap_runner``, as ``m365-gdap-runner-customer-export``."""

from __future__ import annotations

from m365_gdap_runner.cli import list_tenants

if __name__ == "__main__":
    raise SystemExit(list_tenants())
