"""Unit tests for partner sign-in, per-tenant tokens, and the token cache in auth.py."""

from __future__ import annotations

import os
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import msal
import msal_extensions
import pytest

from m365_gdap_runner.auth import (
    AUTHORITY_HOST,
    GRAPH_RESOURCE,
    LoginError,
    PartnerSession,
    TenantAuthError,
    TokenCacheError,
    build_token_cache,
    group_by_resource,
    token_cache_path,
)
from m365_gdap_runner.config.paths import CONFIG_DIR_ENV
from m365_gdap_runner.config.schema import ENV_PREFIX
from tests._gdap_fakes import (
    CONSENT_REQUIRED,
    PARTNER_TENANT,
    TENANT_A,
    TENANT_B,
    FakeMsal,
    fake_access,
    make_settings,
)

pytestmark = pytest.mark.unit


@pytest.fixture(autouse=True)
def _isolated_config(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Strip this package's env vars and point the config dir at a temp dir."""
    for key in list(os.environ):
        if key.startswith(ENV_PREFIX):
            monkeypatch.delenv(key)
    monkeypatch.setenv(CONFIG_DIR_ENV, str(tmp_path))


def _session(fake: FakeMsal, cache: object | None = None) -> PartnerSession:
    return PartnerSession(make_settings(), app_factory=fake, token_cache=cache or object())


# --- login ---------------------------------------------------------------------------


def test_login_is_interactive_when_cache_is_empty() -> None:
    fake = FakeMsal()
    session = _session(fake)
    session.login()
    assert fake.interactive_calls == 1
    assert session.signed_in
    assert fake.created[0][1] == f"{AUTHORITY_HOST}/{PARTNER_TENANT}"


def test_login_is_silent_with_cached_partner_account() -> None:
    fake = FakeMsal()
    fake.seed_cached_account()
    session = _session(fake)
    session.login()
    assert fake.interactive_calls == 0
    assert session.signed_in


def test_login_ignores_cached_account_from_another_tenant() -> None:
    fake = FakeMsal()
    fake.accounts.append({"home_account_id": "someone.other-tenant"})
    fake.has_refresh_token = True
    _session(fake).login()
    assert fake.interactive_calls == 1


def test_login_failure_raises_login_error() -> None:
    fake = FakeMsal(
        interactive_result={
            "error": "access_denied",
            "error_description": "AADSTS50126: Sign-in denied.\r\nTrace ID: 0",
        }
    )
    session = _session(fake)
    with pytest.raises(LoginError, match=r"access_denied: AADSTS50126: Sign-in denied\.$"):
        session.login()
    assert not session.signed_in


def test_login_rejects_account_outside_partner_tenant() -> None:
    fake = FakeMsal(
        interactive_result={
            "access_token": "x",
            "expires_in": 60,
            "id_token_claims": {"oid": "someone", "tid": "other-tenant"},
        }
    )
    with pytest.raises(LoginError, match="not a member of partner tenant"):
        _session(fake).login()


# --- token_for -------------------------------------------------------------------------


def test_token_for_returns_token_with_absolute_expiry() -> None:
    fake = FakeMsal()
    session = _session(fake)
    session.login()
    before = int(time.time())
    token = session.token_for(TENANT_A, ["User.Read.All"])
    after = int(time.time())
    assert token.token == fake_access(TENANT_A)
    assert before + 3600 <= token.expires_on <= after + 3600
    assert fake.silent_calls[-1] == (TENANT_A, ("User.Read.All",), None)


def test_token_for_error_raises_tenant_auth_error_with_details() -> None:
    fake = FakeMsal(tenant_errors={TENANT_A: CONSENT_REQUIRED})
    session = _session(fake)
    session.login()
    with pytest.raises(TenantAuthError) as info:
        session.token_for(TENANT_A, ["User.Read.All"])
    error = info.value
    assert error.tenant_id == TENANT_A
    assert error.error == "invalid_grant"
    assert error.suberror == "consent_required"
    message = str(error)
    assert "AADSTS65001" in message
    assert "Trace ID" not in message
    assert "not consented in this tenant" in message


def test_token_for_without_cached_token_raises_tenant_auth_error() -> None:
    fake = FakeMsal()
    session = _session(fake)
    session.login()
    fake.has_refresh_token = False
    with pytest.raises(TenantAuthError) as info:
        session.token_for(TENANT_A, ["User.Read.All"])
    assert info.value.error == "no_token"


def test_token_for_requires_login() -> None:
    with pytest.raises(RuntimeError, match="login"):
        _session(FakeMsal()).token_for(TENANT_A, ["User.Read.All"])


def test_interaction_required_message_has_hint() -> None:
    error = TenantAuthError(TENANT_A, "interaction_required", "AADSTS50076: MFA required.")
    assert str(error) == (
        f"Tenant {TENANT_A}: interaction_required: AADSTS50076: MFA required. "
        "Hint: Conditional Access or MFA in this tenant requires interaction."
    )


# --- application per authority ---------------------------------------------------------


def test_one_app_per_authority_all_sharing_one_cache() -> None:
    fake = FakeMsal()
    cache = object()
    session = _session(fake, cache)
    session.login()
    session.token_for(TENANT_A, ["User.Read.All"])
    session.token_for(TENANT_B, ["User.Read.All"])
    session.token_for(TENANT_A.upper(), ["User.Read.All"])
    authorities = [authority for _, authority, _ in fake.created]
    assert authorities == [
        f"{AUTHORITY_HOST}/{PARTNER_TENANT}",
        f"{AUTHORITY_HOST}/{TENANT_A}",
        f"{AUTHORITY_HOST}/{TENANT_B}",
    ]
    assert all(token_cache is cache for _, _, token_cache in fake.created)
    assert all(client_id == make_settings().client_id for client_id, _, _ in fake.created)


def test_app_for_is_memoized_under_concurrency() -> None:
    session = _session(FakeMsal())
    with ThreadPoolExecutor(max_workers=8) as pool:
        apps = list(pool.map(lambda _: session.app_for(TENANT_A), range(32)))
    assert len({id(app) for app in apps}) == 1


# --- TenantCredential ----------------------------------------------------------------


def test_credential_delegates_to_bound_tenant_and_passes_claims() -> None:
    fake = FakeMsal()
    session = _session(fake)
    session.login()
    credential = session.credential_for(TENANT_A)
    token = credential.get_token("User.Read.All", claims="challenge", tenant_id=TENANT_A.upper())
    assert token.token == fake_access(TENANT_A)
    assert fake.silent_calls[-1] == (TENANT_A, ("User.Read.All",), "challenge")


def test_credential_rejects_empty_scopes() -> None:
    credential = _session(FakeMsal()).credential_for(TENANT_A)
    with pytest.raises(ValueError, match="at least one scope"):
        credential.get_token()


def test_credential_rejects_other_tenant() -> None:
    credential = _session(FakeMsal()).credential_for(TENANT_A)
    with pytest.raises(ValueError, match="bound to tenant"):
        credential.get_token("User.Read.All", tenant_id=TENANT_B)


# --- scope grouping --------------------------------------------------------------------


def test_group_by_resource() -> None:
    groups = group_by_resource(
        [
            "User.Read.All",
            "https://graph.microsoft.com/Group.Read.All",
            "https://outlook.office365.com/EWS.AccessAsUser.All",
            "api://00000000-0000-0000-0000-000000000001/access_as_user",
        ]
    )
    assert groups == {
        GRAPH_RESOURCE: ["User.Read.All", "https://graph.microsoft.com/Group.Read.All"],
        "https://outlook.office365.com": ["https://outlook.office365.com/EWS.AccessAsUser.All"],
        "api://00000000-0000-0000-0000-000000000001": [
            "api://00000000-0000-0000-0000-000000000001/access_as_user"
        ],
    }


# --- token cache -----------------------------------------------------------------------


def test_memory_cache_is_the_default() -> None:
    cache = build_token_cache(False)
    assert isinstance(cache, msal.TokenCache)
    assert not isinstance(cache, msal_extensions.PersistedTokenCache)
    assert not token_cache_path().exists()


def test_session_builds_memory_cache_from_default_settings() -> None:
    session = PartnerSession(make_settings(), app_factory=FakeMsal())
    session.app_for(TENANT_A)
    assert not token_cache_path().exists()


def test_persisted_cache_uses_encrypted_persistence(monkeypatch: pytest.MonkeyPatch) -> None:
    locations: list[str] = []
    persistence = object()

    def fake_build(location: str) -> object:
        locations.append(location)
        return persistence

    class FakePersistedCache:
        def __init__(self, given: object) -> None:
            self.persistence = given

    monkeypatch.setattr(msal_extensions, "build_encrypted_persistence", fake_build)
    monkeypatch.setattr(msal_extensions, "PersistedTokenCache", FakePersistedCache)
    cache = build_token_cache(True)
    assert isinstance(cache, FakePersistedCache)
    assert cache.persistence is persistence
    assert locations == [str(token_cache_path())]


def test_persisted_cache_failure_raises_and_writes_no_file(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def unavailable(location: str) -> object:
        raise RuntimeError("libsecret is not available")

    monkeypatch.setattr(msal_extensions, "build_encrypted_persistence", unavailable)
    with pytest.raises(TokenCacheError, match="libsecret is not available"):
        PartnerSession(make_settings(persist_token_cache=True), app_factory=FakeMsal())
    assert not token_cache_path().exists()
