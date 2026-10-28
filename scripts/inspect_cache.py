"""What does the disk token cache hold for one account, and does a silent sign-in work?

Run it on the machine and Windows account that ran the script, from the repo folder::

    uv run python scripts/inspect_cache.py admin@contoso.com --tenant contoso.com

It prints the cached accounts, refresh tokens and access tokens for that username (never
the tokens themselves), then, with ``--tenant``, asks for a Graph token the way a direct
sign-in does, without ever opening a browser, and prints MSAL's own reasons if it fails.
"""

from __future__ import annotations

import argparse
import datetime
import logging
from pathlib import Path
from typing import Any

import msal_extensions
from azure_auth import AuthContext
from azure_auth.auth.cache import default_cache_path
from azure_auth.sync import GraphClient
from msal import TokenCache


def _when(unix: Any) -> str:
    """Format a cache timestamp as local time.

    Args:
        unix: Seconds since the epoch, as the cache stores them (often a string).

    Returns:
        ``YYYY-MM-DD HH:MM:SS``, or the raw value if it is not a number.
    """
    try:
        return datetime.datetime.fromtimestamp(int(unix)).strftime("%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError):
        return str(unix)


def show_cache(path: Path, username: str) -> None:
    """Print what the cache file holds for a username, leaving out every secret.

    Args:
        path: The cache file.
        username: The account to look for, compared case-insensitively.
    """
    modified = _when(path.stat().st_mtime)
    print(f"Cache: {path} ({path.stat().st_size:,} bytes, last written {modified})")
    cache = msal_extensions.PersistedTokenCache(
        msal_extensions.build_encrypted_persistence(str(path))
    )
    accounts = list(cache.search(TokenCache.CredentialType.ACCOUNT))
    mine = [a for a in accounts if str(a.get("username", "")).lower() == username.lower()]
    print(f"Accounts in the cache: {len(accounts)}; named {username}: {len(mine)}")
    if not mine:
        names = sorted({str(a.get("username")) for a in accounts}, key=str.lower)
        print("Usernames in the cache:\n  " + "\n  ".join(names))
        return
    for account in mine:
        home = account.get("home_account_id")
        print(f"\nAccount entry: home {home}, tenant {account.get('realm')}")
        for rt in cache.search(TokenCache.CredentialType.REFRESH_TOKEN):
            if rt.get("home_account_id") == home:
                print(
                    f"  refresh token: client {rt.get('client_id')}, family {rt.get('family_id')}"
                )
        for at in cache.search(TokenCache.CredentialType.ACCESS_TOKEN):
            if at.get("home_account_id") == home:
                print(
                    f"  access token: client {at.get('client_id')}, tenant {at.get('realm')}, "
                    f"cached {_when(at.get('cached_at'))}, expires {_when(at.get('expires_on'))}"
                    f"\n    scopes: {at.get('target')}"
                )


def try_silent(path: Path, username: str, tenant: str) -> None:
    """Ask for a Graph token as a direct sign-in would, but never open a browser.

    Args:
        path: The cache file.
        username: The tenant's ``m365_username``.
        tenant: The tenant's GUID or a verified domain.
    """
    print(f"\nSilent Graph sign-in to {tenant} as {username}:")
    auth = AuthContext(tenant, username=username, cache="disk", cache_path=path)
    with GraphClient(auth, scopes=["User.Read"]) as graph:
        client_id, scopes = graph.client_id, graph.scopes
    try:
        token = auth.acquire_token_silently(scopes, client_id=client_id)
    except Exception as exc:
        print(f"  FAILED: {type(exc).__name__}: {exc}")
        return
    print(f"  OK, token expires {_when(token.expires_on)}")


def main() -> None:
    """Parse the arguments and run the checks."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("username", help="The tenant's m365_username.")
    parser.add_argument("--tenant", help="Tenant GUID or domain; also try a silent sign-in.")
    parser.add_argument(
        "--cache-path", type=Path, default=default_cache_path(), help="The cache file."
    )
    args = parser.parse_args()
    show_cache(args.cache_path, args.username)
    if args.tenant:
        # MSAL's own debug lines say whether it found the account and why a refresh failed.
        logging.basicConfig(format="  msal: %(message)s")
        logging.getLogger("msal.application").setLevel(logging.DEBUG)
        try_silent(args.cache_path, args.username, args.tenant)


if __name__ == "__main__":
    main()
