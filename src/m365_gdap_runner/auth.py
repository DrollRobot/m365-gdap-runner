"""Partner sign-in and per-tenant delegated tokens.

One interactive sign-in to the partner tenant yields a refresh token that MSAL
stores without a tenant key, so the same token cache can mint access tokens for
any GDAP customer tenant silently. :class:`PartnerSession` keeps one MSAL
``PublicClientApplication`` per tenant authority, all sharing one token cache;
:class:`TenantCredential` exposes one tenant as an
``azure.core.credentials.TokenCredential``.

The per-tenant authority is set on each application rather than passed to
``acquire_token_silent*``: MSAL warns that it has not decided whether to support
a per-call ``authority``.

The token cache lives in memory unless ``persist_token_cache`` is set, in which
case it is an encrypted file (DPAPI on Windows, Keychain on macOS, libsecret on
Linux). There is no plaintext fallback: if encryption is unavailable,
:func:`build_token_cache` raises :class:`TokenCacheError`.

msal and msal-extensions are untyped, so they are imported lazily inside the
functions that use them and wrapped behind the :class:`MsalApp` protocol.
"""

from __future__ import annotations

import threading
import time
from collections.abc import Iterable, Mapping, Sequence
from pathlib import Path
from typing import Any, Protocol

from azure.core.credentials import AccessToken, TokenCredential

from m365_gdap_runner.config import Settings, paths

# Entra ID sign-in host; each application's authority is this plus a tenant ID.
AUTHORITY_HOST = "https://login.microsoftonline.com"

# Resource that short scope names (e.g. "User.Read.All") belong to.
GRAPH_RESOURCE = "https://graph.microsoft.com"

# Scopes requested from the partner tenant: enough to enumerate GDAP relationships.
PARTNER_SCOPES: tuple[str, ...] = ("DelegatedAdminRelationship.Read.All",)

# Scopes MSAL adds itself and rejects when passed explicitly.
RESERVED_SCOPES = frozenset({"openid", "profile", "offline_access"})

# File name of the encrypted token cache inside the config directory.
CACHE_FILE_NAME = "msal_token_cache.bin"

# Readable hints appended to TenantAuthError messages, keyed by suberror or error.
_AUTH_HINTS = {
    "consent_required": "a requested scope is not consented in this tenant.",
    "interaction_required": "Conditional Access or MFA in this tenant requires interaction.",
}


class TokenCacheError(Exception):
    """Raised when the encrypted on-disk token cache cannot be used."""


class LoginError(Exception):
    """Raised when the partner sign-in fails, is cancelled, or yields no usable account."""


class TenantAuthError(Exception):
    """Raised when a token for one tenant cannot be acquired silently.

    This is the per-tenant failure path: missing consent, Conditional Access,
    lapsed GDAP access, and similar.

    Attributes:
        tenant_id: The tenant the token was requested for.
        error: OAuth error code (e.g. ``invalid_grant``).
        description: Full error description from Entra ID.
        suberror: Entra ID suberror (e.g. ``consent_required``), if any.
    """

    def __init__(
        self,
        tenant_id: str,
        error: str,
        description: str = "",
        suberror: str | None = None,
    ) -> None:
        """Store the error details and build a one-line message.

        Args:
            tenant_id: The tenant the token was requested for.
            error: OAuth error code.
            description: Error description; only its first line goes in the message.
            suberror: Entra ID suberror, if any.
        """
        self.tenant_id = tenant_id
        self.error = error
        self.description = description
        self.suberror = suberror
        super().__init__(_auth_message(tenant_id, error, description, suberror))


def _first_line(text: str) -> str:
    """Return the first non-blank line of *text*.

    Entra ID descriptions append trace and correlation IDs on later lines.

    Args:
        text: Possibly multi-line text.

    Returns:
        The first non-blank line, stripped, or ``""``.
    """
    for line in text.splitlines():
        if line.strip():
            return line.strip()
    return ""


def _auth_message(tenant_id: str, error: str, description: str, suberror: str | None) -> str:
    """Build the one-line message for a :class:`TenantAuthError`.

    Args:
        tenant_id: The tenant the token was requested for.
        error: OAuth error code.
        description: Error description.
        suberror: Entra ID suberror, if any.

    Returns:
        ``Tenant <id>: <error> (<suberror>): <description line> Hint: ...``.
    """
    message = f"Tenant {tenant_id}: {error}"
    if suberror:
        message += f" ({suberror})"
    summary = _first_line(description)
    if summary:
        message += f": {summary}"
    hint = _AUTH_HINTS.get(suberror or "") or _AUTH_HINTS.get(error)
    if hint:
        message += f" Hint: {hint}"
    return message


class MsalApp(Protocol):
    """The subset of ``msal.PublicClientApplication`` this module uses."""

    def get_accounts(self) -> list[dict[str, Any]]:
        """Return the accounts present in the token cache."""
        ...

    def acquire_token_silent_with_error(
        self,
        scopes: list[str],
        account: Mapping[str, Any] | None,
        *,
        claims_challenge: str | None = None,
    ) -> dict[str, Any] | None:
        """Return a token result, an error result, or None when nothing is cached."""
        ...

    def acquire_token_interactive(self, scopes: list[str]) -> dict[str, Any]:
        """Run the browser sign-in and return a token or error result."""
        ...


class MsalAppFactory(Protocol):
    """Callable that builds one :class:`MsalApp` for an authority."""

    def __call__(self, client_id: str, *, authority: str, token_cache: object) -> MsalApp:
        """Build an application bound to *authority* that uses *token_cache*."""
        ...


class TokenSession(Protocol):
    """What the runner needs from a session: the partner tenant and per-tenant credentials."""

    @property
    def partner_tenant_id(self) -> str:
        """The partner tenant ID."""
        ...

    def credential_for(self, tenant_id: str) -> TokenCredential:
        """Return a credential bound to *tenant_id*."""
        ...


def default_app_factory(client_id: str, *, authority: str, token_cache: object) -> MsalApp:
    """Build a real ``msal.PublicClientApplication``.

    Args:
        client_id: Public client application ID.
        authority: Authority URL including the tenant ID.
        token_cache: MSAL token cache shared by every application.

    Returns:
        The MSAL application.
    """
    import msal

    app: MsalApp = msal.PublicClientApplication(
        client_id, authority=authority, token_cache=token_cache
    )
    return app


def token_cache_path() -> Path:
    """Return where the encrypted token cache is stored.

    Returns:
        ``<config dir>/msal_token_cache.bin``.
    """
    return paths.config_dir() / CACHE_FILE_NAME


def build_token_cache(persist: bool) -> object:
    """Build the MSAL token cache.

    Args:
        persist: False for an in-memory cache; True for an encrypted file at
            :func:`token_cache_path`.

    Returns:
        An ``msal.TokenCache`` or ``msal_extensions.PersistedTokenCache``.

    Raises:
        TokenCacheError: If *persist* is True and encryption is unavailable on
            this system. Never falls back to plaintext.
    """
    if not persist:
        import msal

        cache: object = msal.TokenCache()
        return cache

    import msal_extensions

    paths.ensure_config_dir()
    location = token_cache_path()
    try:
        persistence = msal_extensions.build_encrypted_persistence(str(location))
    except Exception as exc:  # platform-specific: ImportError, keychain/libsecret errors
        raise TokenCacheError(
            f"Encrypted token cache is unavailable on this system ({exc!r}). "
            "Run without persist_token_cache to use an in-memory cache."
        ) from exc
    persisted: object = msal_extensions.PersistedTokenCache(persistence)
    return persisted


def group_by_resource(scopes: Iterable[str]) -> dict[str, list[str]]:
    """Group scopes by the resource they belong to.

    One token request cannot span resources. Short names (``User.Read.All``)
    belong to Graph; URI scopes belong to everything before their last ``/``.

    Args:
        scopes: Delegated scopes.

    Returns:
        Resource mapped to its scopes, in first-seen order.
    """
    groups: dict[str, list[str]] = {}
    for scope in scopes:
        resource = scope.rpartition("/")[0] if "/" in scope else GRAPH_RESOURCE
        groups.setdefault(resource, []).append(scope)
    return groups


def _home_tenant(account: Mapping[str, Any]) -> str:
    """Return the lower-cased home tenant ID of an MSAL account.

    Args:
        account: MSAL account dict (``home_account_id`` is ``<oid>.<tid>``).

    Returns:
        The tenant part of ``home_account_id``, or ``""``.
    """
    return str(account.get("home_account_id", "")).rpartition(".")[2].lower()


def _optional_str(value: object) -> str | None:
    """Return *value* as a string, or None when it is None or empty.

    Args:
        value: A value from an MSAL result dict.

    Returns:
        The string, or None.
    """
    return str(value) if value else None


class PartnerSession:
    """One partner sign-in, shared by every tenant in a run.

    Thread-safe: applications are memoized under a lock, and MSAL's token cache
    (plus msal-extensions' file lock when persisted) guards the cache itself.
    """

    def __init__(
        self,
        settings: Settings,
        *,
        app_factory: MsalAppFactory = default_app_factory,
        token_cache: object | None = None,
    ) -> None:
        """Create a session; nothing is requested until :meth:`login`.

        Args:
            settings: Resolved settings (partner tenant, client ID, cache mode).
            app_factory: Builds MSAL applications; tests inject a fake.
            token_cache: Token cache to use instead of building one from settings.

        Raises:
            TokenCacheError: If the persisted cache is requested but unavailable.
        """
        self._settings = settings
        self._app_factory = app_factory
        self._token_cache = (
            token_cache
            if token_cache is not None
            else build_token_cache(settings.persist_token_cache)
        )
        self._apps: dict[str, MsalApp] = {}
        self._apps_lock = threading.Lock()
        self._account: dict[str, Any] | None = None

    @property
    def partner_tenant_id(self) -> str:
        """The partner tenant ID from settings."""
        return self._settings.partner_tenant_id

    @property
    def signed_in(self) -> bool:
        """Whether :meth:`login` has succeeded."""
        return self._account is not None

    def app_for(self, tenant_id: str) -> MsalApp:
        """Return the memoized MSAL application for *tenant_id*'s authority.

        The application is built outside the lock (MSAL may fetch authority
        metadata while constructing it); a concurrent duplicate is discarded.

        Args:
            tenant_id: Tenant ID.

        Returns:
            The application bound to ``<AUTHORITY_HOST>/<tenant_id>``.
        """
        key = tenant_id.lower()
        with self._apps_lock:
            existing = self._apps.get(key)
        if existing is not None:
            return existing
        created = self._app_factory(
            self._settings.client_id,
            authority=f"{AUTHORITY_HOST}/{tenant_id}",
            token_cache=self._token_cache,
        )
        with self._apps_lock:
            return self._apps.setdefault(key, created)

    def login(self) -> None:
        """Sign in to the partner tenant. The only interactive point.

        Tries a cached account from the partner tenant silently first (useful
        with a persisted cache), then falls back to the browser sign-in.

        Raises:
            LoginError: If the sign-in fails or the account cannot be found
                in the cache afterward.
        """
        app = self.app_for(self.partner_tenant_id)
        scopes = list(PARTNER_SCOPES)
        partner = self.partner_tenant_id.lower()
        for account in app.get_accounts():
            if _home_tenant(account) != partner:
                continue
            result = app.acquire_token_silent_with_error(scopes, account)
            if result is not None and "access_token" in result:
                self._account = account
                return
        result = app.acquire_token_interactive(scopes)
        if "access_token" not in result:
            error = result.get("error", "unknown_error")
            summary = _first_line(str(result.get("error_description", "")))
            raise LoginError(f"Sign-in failed: {error}" + (f": {summary}" if summary else ""))
        self._account = self._signed_in_account(app, result)

    def _signed_in_account(self, app: MsalApp, result: Mapping[str, Any]) -> dict[str, Any]:
        """Find the cached account that an interactive sign-in produced.

        Args:
            app: The partner tenant's application.
            result: The successful interactive result (carries ``id_token_claims``).

        Returns:
            The MSAL account dict.

        Raises:
            LoginError: If no cached account matches the signed-in user.
        """
        claims = result.get("id_token_claims") or {}
        home_account_id = f"{claims.get('oid')}.{claims.get('tid')}".lower()
        for account in app.get_accounts():
            if str(account.get("home_account_id", "")).lower() == home_account_id:
                return account
        raise LoginError(
            f"Signed in, but the account is not a member of partner tenant "
            f"{self.partner_tenant_id}. Sign in with a partner tenant account."
        )

    def token_for(
        self, tenant_id: str, scopes: Sequence[str], *, claims: str | None = None
    ) -> AccessToken:
        """Acquire an access token for *tenant_id* silently.

        Uses ``acquire_token_silent_with_error``: plain ``acquire_token_silent``
        returns None on failure and discards the reason.

        Args:
            tenant_id: Tenant to mint the token in.
            scopes: Scopes for a single resource. The request fails if any is
                not consented in that tenant.
            claims: Claims challenge to pass through, if any.

        Returns:
            The access token with an absolute ``expires_on`` timestamp.

        Raises:
            RuntimeError: If :meth:`login` has not succeeded.
            TenantAuthError: If no token can be acquired for this tenant.
        """
        if self._account is None:
            raise RuntimeError("PartnerSession.login() must succeed before requesting tokens.")
        app = self.app_for(tenant_id)
        result = app.acquire_token_silent_with_error(
            list(scopes), self._account, claims_challenge=claims
        )
        if result is None:
            raise TenantAuthError(
                tenant_id, "no_token", "The token cache has no refresh token for this account."
            )
        if "access_token" not in result:
            raise TenantAuthError(
                tenant_id,
                str(result.get("error", "unknown_error")),
                str(result.get("error_description", "")),
                _optional_str(result.get("suberror")),
            )
        # MSAL reports seconds remaining, not an absolute timestamp.
        expires_on = int(time.time()) + int(result["expires_in"])
        return AccessToken(str(result["access_token"]), expires_on)

    def credential_for(self, tenant_id: str) -> TenantCredential:
        """Return a credential bound to *tenant_id*.

        Args:
            tenant_id: Tenant ID.

        Returns:
            A :class:`TenantCredential` backed by this session.
        """
        return TenantCredential(self, tenant_id)


class TenantCredential(TokenCredential):
    """``TokenCredential`` for one tenant, backed by a :class:`PartnerSession`.

    Holds no token state; MSAL's cache handles access-token reuse and expiry.
    """

    def __init__(self, session: PartnerSession, tenant_id: str) -> None:
        """Bind the credential to a tenant.

        Args:
            session: The signed-in partner session.
            tenant_id: Tenant this credential issues tokens for.
        """
        self._session = session
        self._tenant_id = tenant_id

    @property
    def tenant_id(self) -> str:
        """The tenant this credential issues tokens for."""
        return self._tenant_id

    def get_token(
        self,
        *scopes: str,
        claims: str | None = None,
        tenant_id: str | None = None,
        enable_cae: bool = False,
        **kwargs: Any,
    ) -> AccessToken:
        """Acquire an access token for the bound tenant.

        Args:
            *scopes: Scopes for a single resource; at least one.
            claims: Claims challenge to pass through, if any.
            tenant_id: Optional tenant; must match the bound tenant.
            enable_cae: Accepted for protocol compatibility; ignored.
            **kwargs: Accepted for protocol compatibility; ignored.

        Returns:
            The access token.

        Raises:
            ValueError: If no scopes are given or *tenant_id* differs from the
                bound tenant.
            TenantAuthError: If the token cannot be acquired.
        """
        if not scopes:
            raise ValueError("get_token() requires at least one scope.")
        if tenant_id is not None and tenant_id.lower() != self._tenant_id.lower():
            raise ValueError(
                f"This credential is bound to tenant {self._tenant_id}, not {tenant_id}."
            )
        return self._session.token_for(self._tenant_id, scopes, claims=claims)
