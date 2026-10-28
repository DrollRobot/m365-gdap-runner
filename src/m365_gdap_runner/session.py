"""The partner sign-in and what hangs off it: tenant listing and per-tenant contexts.

azure-auth does the authentication. This module only decides *what* to sign in to: one
sign-in per client id the script uses, all at the partner (home) tenant, before the loop
starts. Customer tenants are reached through contexts from azure-auth's ``get_tenant_context()``,
which never prompt; they redeem the home sign-in's refresh token in the customer tenant.

Every sign-in opens the browser: the system default, or the ``browser_path`` setting's, with
``browser_args``. The Windows broker is not offered: GDAP does not work through it.

A customer whose Conditional Access policy finds the partner sign-in too old, or its
multi-factor authentication (AADSTS70043, AADSTS50078), gets a new partner sign-in that asks
for the password again, once a run, and is checked through it again.

A tenant the partner sign-in cannot reach (no GDAP relationship, missing consent, a policy
that wants a fresh sign-in, or another Microsoft cloud) gets a direct sign-in instead: a
context of its own, in its own cloud, as the tenant's own account. So does a tenant whose
``m365_gdap`` cell in the tenants file is not TRUE, without trying the partner sign-in. A
tenant with no account of its own gets none. Where the partner sign-in lacks consent, the
direct sign-in opens Entra's consent screen whatever the cache holds, so an administrator can
tick "Consent on behalf of your organization": a tenant whose consent is one user's own never
shows the screen again by itself, and the partner sign-in needs the tenant-wide grant. When a
direct sign-in opens a window, the tenant's ``m365_pw_url``, if any, opens with it, once, in
the ``pw_browser_path`` setting's browser, with ``pw_browser_args``, or the system default.
Ctrl+C during a direct sign-in asks whether to skip that tenant or stop the run. All of that
happens in
:meth:`PartnerSession.check_access`, which hands each tenant to the run as soon as its way in
is settled and opens every sign-in window itself, one at a time. Every sign-in, the partner's
and each tenant's own, asks Entra for a new token for each client rather than trusting one
saved earlier, so a sign-in that a policy has started refusing fails there, and a tenant's
``run()`` opens a window only when a sign-in that passed the check is refused later.

Why sign in per client: Microsoft's first-party apps differ per resource (Graph PowerShell,
Exchange Online PowerShell, Azure PowerShell) and a refresh token belongs to one app, so a
script that uses Graph and Exchange needs two home sign-ins. With the encrypted disk cache,
which is opt-in (``cache = "disk"``), later runs are silent. It is kept in azure-auth's file,
or the ``cache_path`` setting's; ``cache_reminder_path`` leaves a link to it in plain sight
(see ``cache_reminder.py``).
"""

from __future__ import annotations

import logging
import sys
import webbrowser
from collections import deque
from collections.abc import Callable, Generator, Iterable, Sequence
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from pathlib import Path
from time import monotonic
from typing import Literal, cast

import azure_auth
import httpx
from azure_auth import AuthContext, AuthError, ConsentRequired
from azure_auth.auth.cache import default_cache_path
from azure_auth.clouds import Cloud, get_cloud
from azure_auth.sync import GraphClient

from m365_gdap_runner.cache_reminder import place_reminder
from m365_gdap_runner.config import ConfigError, Settings
from m365_gdap_runner.context import CLIENT_KINDS, TenantContext, build_client
from m365_gdap_runner.error_text import brief
from m365_gdap_runner.loader import Scope, choose_scopes
from m365_gdap_runner.runner import tag_logs
from m365_gdap_runner.tenants import Tenant

logger = logging.getLogger(__name__)

# Delegated scope the partner sign-in needs to list GDAP customers. Payload scopes are never
# added here: consent in the partner tenant says nothing about the customer tenants.
PARTNER_SCOPES = ("DelegatedAdminRelationship.Read.All",)

CACHE_KINDS = ("memory", "disk")

# Seconds between progress lines while check_access() checks tenants.
_PROGRESS_EVERY = 30.0

# Entra's codes for a customer whose Conditional Access policy finds the partner's home sign-in
# too old: its multi-factor authentication (50078) or the sign-in itself (70043). Both seen
# live 2026-10-22. A new home sign-in that asks for the password again answers both.
_STALE_SIGN_IN_CODES = frozenset({50078, 70043})


def _place_cache_reminder(cache_file: Path, folder: str) -> None:
    """Put the DELETEME reminder in a folder, pointing at the disk cache's folder.

    A reminder that cannot be placed is reported, and the run goes on without it.

    Args:
        cache_file: The disk cache file.
        folder: The ``cache_reminder_path`` setting; not empty.
    """
    try:
        link = place_reminder(cache_file, Path(folder).expanduser())
    except OSError as exc:
        logger.warning("Could not put the token cache reminder in %s: %s", folder, exc)
        return
    logger.debug("Token cache reminder %s points at %s", link, cache_file.parent)


def _browser(path: str, args: Sequence[str], *, prefix: str = "") -> list[str] | None:
    """Build a browser command from a pair of browser settings.

    Args:
        path: The ``browser_path`` setting, or ``pw_browser_path``: a path, or a program on
            ``PATH``.
        args: The matching ``browser_args`` setting.
        prefix: ``pw_`` for the password page's settings, which the error message names.

    Returns:
        The program followed by its arguments, or None for the system default.

    Raises:
        ConfigError: If the args setting is set without the path setting.
    """
    if not path:
        if args:
            raise ConfigError(
                f"{prefix}browser_args needs {prefix}browser_path, the browser to pass them to."
            )
        return None
    return [str(Path(path).expanduser()), *args]


def _open_pw_page(tenant: Tenant, browser: Sequence[str] | None) -> None:
    """Open a tenant's password page, for the sign-in window about to open.

    A page that cannot be opened is reported, and the sign-in goes on without it.

    Args:
        tenant: The tenant, with ``pw_url`` filled in.
        browser: The ``pw_browser_path`` program followed by ``pw_browser_args``, or None for
            the system default.
    """
    logger.debug("Opening the password page %s", tenant.pw_url)
    if browser is None:
        opened = webbrowser.open(tenant.pw_url)
    else:
        # "%s" is where webbrowser puts the address.
        opened = webbrowser.BackgroundBrowser([*browser, "%s"]).open(tenant.pw_url)
    if not opened:
        logger.warning("Could not open the password page %s", tenant.pw_url)


def _ask_to_skip(name: str) -> None:
    """Ask, after Ctrl+C during a sign-in, whether to skip it or stop the run.

    Enter skips the sign-in. The question goes to stderr, with the logs, since stdout is the
    script's.

    Args:
        name: What the sign-in was to: a tenant's display name, or the partner tenant.

    Raises:
        KeyboardInterrupt: On Ctrl+C at the question, or when nobody can answer it (stdin is
            closed), to stop the run.
    """
    print(
        f"Stopped the sign-in to {name}. Press Enter to skip it, or Ctrl+C to stop the run.",
        file=sys.stderr,
        flush=True,
    )
    try:
        input()
    except EOFError:
        raise KeyboardInterrupt from None


def _time_left(seconds: float) -> str:
    """Say roughly how long the rest of the access check will take.

    Args:
        seconds: The estimate.

    Returns:
        ``about N min left``, or ``under a minute left``.
    """
    if seconds < 60:
        return "under a minute left"
    return f"about {round(seconds / 60)} min left"


def _partner_tenant(name: str, cloud: Cloud) -> azure_auth.Tenant:
    """Build the partner tenant in a configured cloud, looking it up only when needed.

    A GUID with its cloud is all azure-auth needs, so no request is made. A domain name is
    still looked up, for its GUID, in that cloud only.

    Args:
        name: The ``partner_tenant_id`` setting: a GUID or a domain name.
        cloud: The ``cloud`` setting.

    Returns:
        The tenant, for :class:`~azure_auth.AuthContext`.

    Raises:
        azure_auth.AuthError: If a domain name cannot be looked up in that cloud.
    """
    try:
        return azure_auth.Tenant(name, cloud)
    except ValueError:
        return azure_auth.Tenant.lookup(name, cloud)


class PartnerSession:
    """One partner sign-in, and the customer tenants reachable from it.

    Attributes:
        auth: The home azure-auth context, for the partner tenant.
    """

    def __init__(self, settings: Settings, *, transport: httpx.BaseTransport | None = None) -> None:
        """Build the home context from settings. No sign-in happens yet.

        Args:
            settings: Resolved settings.
            transport: Optional httpx transport for every client this session builds (tests).

        Raises:
            ConfigError: If ``cache`` or ``cloud`` holds a value azure-auth does not know,
                ``browser_args`` is set without ``browser_path``, ``pw_browser_args`` without
                ``pw_browser_path``, or ``browser_timeout`` is not more than 0.
            azure_auth.AuthError: If the partner tenant cannot be looked up, or the disk
                cache cannot be encrypted on this platform.
        """
        if settings.cache not in CACHE_KINDS:
            raise ConfigError(
                f"cache must be one of {', '.join(CACHE_KINDS)}, not {settings.cache!r}."
            )
        if settings.browser_timeout <= 0:
            raise ConfigError(
                f"browser_timeout must be more than 0 seconds, not {settings.browser_timeout!r}."
            )
        tenant: str | azure_auth.Tenant = settings.partner_tenant_id
        if settings.cloud:
            try:
                cloud = get_cloud(settings.cloud)
            except ValueError as exc:
                raise ConfigError(str(exc)) from exc
            tenant = _partner_tenant(settings.partner_tenant_id, cloud)
        self._transport = transport
        self._settings = settings
        # None keeps the disk cache in azure-auth's default file.
        self._cache_path = Path(settings.cache_path).expanduser() if settings.cache_path else None
        self._browser = _browser(settings.browser_path, settings.browser_args)
        self._pw_browser = _browser(
            settings.pw_browser_path, settings.pw_browser_args, prefix="pw_"
        )
        # How check_access() found each tenant can be reached, by lower-cased tenant ID. It
        # writes and context_for() reads, on different threads, but never for the same tenant
        # at once: check_access() hands a tenant out only after writing its route.
        self._routes: dict[str, AuthContext | Exception] = {}
        # The Graph scopes each tenant is asked for, by lower-cased tenant ID: the script's,
        # with one scope chosen from each tuple (:func:`choose_scopes`). Written and read as
        # ``_routes`` is.
        self._scopes: dict[str, tuple[str, ...]] = {}
        # Whether the partner tenant was signed in to again this run, for a customer that found
        # the sign-in too old: None until one does, then whether it worked.
        self._signed_in_again: bool | None = None
        self.auth = AuthContext(
            tenant,
            username=settings.username,
            client_id=settings.client_id or None,
            cache=cast(Literal["memory", "disk"], settings.cache),
            cache_path=self._cache_path,
            browser_timeout=settings.browser_timeout,
            browser=self._browser,
        )
        if settings.cache == "disk" and settings.cache_reminder_path:
            _place_cache_reminder(
                self._cache_path or default_cache_path(), settings.cache_reminder_path
            )

    @property
    def partner_tenant_id(self) -> str:
        """The partner tenant's GUID."""
        return self.auth.tenant.id

    @property
    def cloud_name(self) -> str:
        """Name of the cloud this session signs in to, as azure-auth spells it."""
        return self.auth.tenant.cloud.name

    def cloud_of(self, tenant_id: str) -> str:
        """Find the cloud a tenant lives in, by OpenID Connect discovery.

        Discovery is unauthenticated and asks each public sign-in host once, so it costs
        three requests and works for any tenant. It fills the tenants file's ``m365_cloud``
        column.

        Args:
            tenant_id: The tenant's GUID.

        Returns:
            The cloud's name (``Commercial``, ``USGov``, ``USGovDoD`` or ``China``), or ``""``
            when no sign-in host knows the tenant or none can be reached.
        """
        try:
            return azure_auth.Tenant.lookup(tenant_id, transport=self._transport).cloud.name
        except (AuthError, httpx.HTTPError, ValueError) as exc:
            logger.debug("Could not look up the cloud of %s: %s", tenant_id, exc)
            return ""

    def login(self, clients: Iterable[str] = (), *, reauthenticate: bool = False) -> None:
        """Sign in at the partner tenant, once per client id the payload needs.

        Graph is always signed in, for listing customers. Each client asks Entra for a new
        token, so a sign-in window opens only for a client id with no usable cached sign-in,
        or one Entra will not renew. Each client's sign-in is written to the disk cache at
        once, so a run killed later keeps it.

        Args:
            clients: The payload's ``CLIENTS``.
            reauthenticate: Open a sign-in window for every client instead, and have Entra ask
                for the password again, and for multi-factor authentication where the partner
                tenant wants it, whatever the cache holds (azure-auth's ``login()`` option of
                that name).

        Raises:
            azure_auth.AuthError: If a sign-in fails or is cancelled.
        """
        kinds = ["graph", *(kind for kind in clients if kind != "graph")]
        for kind in kinds:
            if kind not in CLIENT_KINDS:
                raise ValueError(f"Unknown client kind {kind!r}.")
            logger.debug("Signing in to the partner tenant for %s", kind)
            scopes = PARTNER_SCOPES if kind == "graph" else ()
            with build_client(kind, self.auth, scopes=scopes, transport=self._transport) as client:
                # A sign-in that asked for the password got its token from Entra already.
                client.login(fresh=not reauthenticate, reauthenticate=reauthenticate)
            self.auth.flush_cache()

    def list_tenants(self) -> list[Tenant]:
        """List the customer tenants with a GDAP relationship to the partner.

        Returns:
            Tenants deduplicated by ID, sorted by display name. A customer without a display
            name is named by its ID.

        Raises:
            azure_auth.AuthError: If no partner token can be obtained.
            azure_auth.GraphError: If Graph rejects the request.
        """
        with GraphClient(
            self.auth, scopes=list(PARTNER_SCOPES), transport=self._transport
        ) as graph:
            customers = graph.list_customers()
        tenants: dict[str, Tenant] = {}
        for customer in customers:
            tenant_id = str(customer["tenantId"])
            display_name = str(customer.get("displayName") or tenant_id)
            tenants.setdefault(tenant_id.lower(), Tenant(tenant_id, display_name))
        return sorted(
            tenants.values(), key=lambda t: (t.display_name.casefold(), t.tenant_id.lower())
        )

    def check_access(
        self,
        tenants: Sequence[Tenant],
        *,
        scopes: Sequence[Scope] = (),
        clients: Sequence[str] = ("graph",),
        jobs: int = 1,
        grant_when_possible: bool = False,
    ) -> Generator[Tenant, None, None]:
        """Find a way into each tenant, and hand it out as soon as one is found.

        A tenant in the partner's cloud is checked first through the partner sign-in, silently
        and the way the run will use it: every client the script uses must get a new token
        from Entra (:meth:`TenantContext.probe`). Where *scopes* has a tuple, the check reads
        what the tenant has granted first and asks for the first scope in the tuple it holds,
        or the first scope in the tuple where it holds none (:func:`choose_scopes`); a direct
        sign-in asks for the first scope in each tuple. With *grant_when_possible*, a tenant whose
        grant lacks some of *scopes* has them added to the app's grant there, when the app holds
        ``DelegatedPermissionGrant.ReadWrite.All``, and passes. A tenant whose policy finds the
        partner sign-in too old (AADSTS70043, AADSTS50078) has the partner tenant signed in to
        again, asking for the password, once a run (:meth:`_sign_in_again`), and is checked
        again. A tenant that fails that check, for whatever reason, a tenant in another cloud,
        and a tenant whose ``gdap`` is False, which is never checked, then get a direct
        sign-in: a context of their own, in their own cloud, as their ``username``, signed in
        once per client. One with no ``username`` gets none, and no prompt. A sign-in window
        opens only where no cached sign-in will do, except for a tenant whose check failed for
        missing consent: its Graph sign-in opens Entra's consent screen whatever the cache
        holds, for the tenant-wide grant (:meth:`_sign_in_directly`).

        The silent checks run in the background, *jobs* at a time. Direct sign-ins run on the
        thread that iterates, one at a time, as soon as a tenant needs one: the tenants never
        checked first, then each one whose check fails. Before each one, every tenant whose
        check has passed is handed out, so a sign-in window holds up no tenant that is ready.
        A log line every 30 seconds says how many checks are done and about how long the rest
        will take. Ctrl+C during a direct sign-in asks whether to skip that tenant, which
        leaves it unreachable, or stop the run, which raises KeyboardInterrupt from here.
        Stopping the iteration (``close()``) cancels the checks not yet started and opens no
        more sign-in windows.

        :meth:`context_for` then hands out what was found, and raises the reason for a tenant
        that could not be reached.

        Args:
            tenants: The selected tenants, with their cloud filled in.
            scopes: The script's Graph scopes.
            clients: The script's ``CLIENTS``.
            jobs: How many silent checks run at once.
            grant_when_possible: Add the scopes a tenant lacks to the app's grant where the
                app can, through the partner sign-in only (azure-auth's ``GraphClient`` option
                of that name). Each grant waits for a token that carries the new scopes,
                about a minute and three at most.

        Yields:
            Each tenant once, as soon as :meth:`context_for` can hand out its context or the
            reason it could not be reached.
        """
        partner_cloud = self.cloud_name.lower()
        nearby = [
            t for t in tenants if t.gdap and (t.cloud or self.cloud_name).lower() == partner_cloud
        ]
        checked_ids = {t.tenant_id.lower() for t in nearby}
        # Tenants waiting for a direct sign-in, each with why the partner sign-in failed, or
        # None when it was never tried.
        direct: deque[tuple[Tenant, Exception | None]] = deque(
            (t, None) for t in tenants if t.tenant_id.lower() not in checked_ids
        )
        logger.info("Checking access to %d tenant(s)", len(tenants))
        done = 0
        started = last = monotonic()
        with ThreadPoolExecutor(
            max_workers=max(1, min(jobs, len(nearby))), thread_name_prefix="check"
        ) as pool:
            checks = {
                pool.submit(self._through_partner, t, scopes, clients, grant_when_possible): t
                for t in nearby
            }
            position = {future: index for index, future in enumerate(checks)}
            waiting = set(checks)
            try:
                while waiting or direct:
                    # Wait for a check only when no sign-in is waiting to be done.
                    finished, waiting = wait(
                        waiting, timeout=0 if direct else None, return_when=FIRST_COMPLETED
                    )
                    for future in sorted(finished, key=position.__getitem__):
                        tenant, route = checks[future], future.result()
                        done += 1
                        now = monotonic()
                        if done < len(nearby) and now - last >= _PROGRESS_EVERY:
                            last = now
                            left = (now - started) / done * (len(nearby) - done)
                            logger.info(
                                "Checked %d of %d tenant(s), %s",
                                done,
                                len(nearby),
                                _time_left(left),
                            )
                        if not isinstance(route, AuthContext) and self._sign_in_again(
                            tenant, route, clients, grant_when_possible
                        ):
                            route = self._through_partner(
                                tenant, scopes, clients, grant_when_possible
                            )
                        if isinstance(route, AuthContext):
                            self._routes[tenant.tenant_id.lower()] = route
                            yield tenant
                        else:
                            direct.append((tenant, route))
                    if direct:
                        tenant, failure = direct.popleft()
                        self._routes[tenant.tenant_id.lower()] = self._reach_directly(
                            tenant, failure, scopes, clients, grant_when_possible
                        )
                        yield tenant
            finally:
                # On Ctrl+C, or when the run stops early, the checks not yet started never start.
                for future in checks:
                    future.cancel()

    def _sign_in_again(
        self,
        tenant: Tenant,
        failure: Exception,
        clients: Sequence[str],
        grant_when_possible: bool,
    ) -> bool:
        """Sign in at the partner tenant again, for a tenant that finds the sign-in too old.

        Renewing the partner sign-in keeps the time of the first one, and of its multi-factor
        authentication, so a customer whose policy allows less time refuses it however often
        it is renewed. A new sign-in that asks for the password again resets both. It opens
        one window per client, once a run, for the first tenant that needs it; a tenant that
        needs it after that is checked again without one. Ctrl+C during it asks whether to
        skip it, which sends the tenants that needed it to their own sign-in, or stop the run.

        Args:
            tenant: The tenant whose check through the partner sign-in failed.
            failure: Why it failed.
            clients: The script's ``CLIENTS``.
            grant_when_possible: As for :meth:`check_access`, for the reason logged.

        Returns:
            Whether to check the tenant through the partner sign-in again: the failure is one a
            new sign-in answers, and the partner tenant has been signed in to again.

        Raises:
            KeyboardInterrupt: When the person chooses to stop the run.
        """
        if not (isinstance(failure, AuthError) and failure.error_codes & _STALE_SIGN_IN_CODES):
            return False
        if self._signed_in_again is None:
            with tag_logs(tenant):
                logger.info(
                    "partner: %s; signing in to the partner tenant again",
                    brief(failure, grant_when_possible=grant_when_possible),
                )
            try:
                self.login(clients, reauthenticate=True)
            except KeyboardInterrupt:
                _ask_to_skip("the partner tenant")
                self._signed_in_again = False
            except Exception as exc:
                logger.debug("Signing in to the partner tenant again failed", exc_info=True)
                logger.warning("Could not sign in to the partner tenant again: %s", brief(exc))
                self._signed_in_again = False
            else:
                self._signed_in_again = True
        return self._signed_in_again

    def _reach_directly(
        self,
        tenant: Tenant,
        failure: Exception | None,
        scopes: Sequence[Scope],
        clients: Sequence[str],
        grant_when_possible: bool,
    ) -> AuthContext | Exception:
        """Give a tenant the partner sign-in does not reach a sign-in of its own, if it can.

        The sign-in asks for the first scope in each tuple in *scopes*: what the tenant has
        granted cannot be read before it.

        Args:
            tenant: The tenant.
            failure: Why its check through the partner sign-in failed, or None when it had none:
                its ``gdap`` is False, or it is in another cloud.
            scopes: The script's Graph scopes.
            clients: The script's ``CLIENTS``.
            grant_when_possible: Whether that check could add missing scopes to the app's
                grant, for what the failure says of it.

        Returns:
            The tenant's own context, else why it cannot be reached, in short
            (:func:`~m365_gdap_runner.error_text.brief`): ``partner: <why>; direct: <why>``. A
            tenant with no ``username`` gets no sign-in, and no prompt.
        """
        if failure is not None:
            partner = brief(failure, grant_when_possible=grant_when_possible)
        elif not tenant.gdap:
            partner = "m365_gdap is FALSE"
        else:
            partner = f"tenant is in the {tenant.cloud} cloud"
        if not tenant.username:
            return AuthError(f"partner: {partner}; direct: no m365_username")
        chosen = choose_scopes(scopes, None)
        self._scopes[tenant.tenant_id.lower()] = chosen
        consent = isinstance(failure, ConsentRequired)
        with tag_logs(tenant):
            if consent:
                logger.info(
                    'partner: %s; direct: signing in as %s; tick "Consent on behalf of your '
                    'organization" in the consent screen',
                    partner,
                    tenant.username,
                )
            else:
                logger.info("partner: %s; direct: signing in as %s", partner, tenant.username)
            route = self._sign_in_directly(tenant, chosen, clients, consent=consent)
        if isinstance(route, AuthContext):
            return route
        return AuthError(f"partner: {partner}; direct: {route}")

    def _through_partner(
        self,
        tenant: Tenant,
        scopes: Sequence[Scope],
        clients: Sequence[str],
        grant_when_possible: bool,
    ) -> AuthContext | Exception:
        """Check a tenant through the partner sign-in, silently.

        Args:
            tenant: The tenant, in the partner's cloud.
            scopes: The script's Graph scopes.
            clients: The script's ``CLIENTS``.
            grant_when_possible: Add missing scopes to the app's grant where it can.

        Returns:
            The customer tenant's context when every client got a token, else what went wrong.
        """
        with tag_logs(tenant):
            try:
                auth = self.auth.get_tenant_context(tenant.tenant_id)
                chosen = self._choose_scopes(auth, scopes, clients)
                self._scopes[tenant.tenant_id.lower()] = chosen
                with TenantContext(
                    tenant,
                    auth,
                    scopes=chosen,
                    clients=clients,
                    grant_when_possible=grant_when_possible,
                    transport=self._transport,
                ) as context:
                    context.probe()
            except Exception as exc:
                logger.debug("The partner sign-in cannot reach the tenant", exc_info=True)
                return exc
            return auth

    def _choose_scopes(
        self, auth: AuthContext, scopes: Sequence[Scope], clients: Sequence[str]
    ) -> tuple[str, ...]:
        """Pick the scope to ask a tenant for from each tuple in the script's scopes.

        Where *scopes* has a tuple and the script uses Graph, the tenant's grant is read
        through *auth* (azure-auth's ``granted_scopes()``: one silent request to Entra, no
        prompt). A grant that cannot be read leaves each tuple its first scope, and the check
        that follows raises the reason, if it is one.

        Args:
            auth: The tenant's context, through the partner sign-in.
            scopes: The script's Graph scopes.
            clients: The script's ``CLIENTS``.

        Returns:
            One scope per entry (:func:`choose_scopes`).
        """
        if "graph" not in clients or all(isinstance(entry, str) for entry in scopes):
            return choose_scopes(scopes, None)
        granted: frozenset[str] | None = None
        try:
            with build_client("graph", auth, transport=self._transport) as client:
                assert isinstance(client, GraphClient)  # noqa: S101
                granted = client.granted_scopes()
        except AuthError:
            logger.debug("The tenant's grant cannot be read", exc_info=True)
        chosen = choose_scopes(scopes, granted)
        logger.debug("Asking for Graph scopes %s", ", ".join(chosen))
        return chosen

    def _sign_in_directly(
        self,
        tenant: Tenant,
        scopes: Sequence[str],
        clients: Sequence[str],
        *,
        consent: bool = False,
    ) -> AuthContext | str:
        """Sign in to a tenant on its own, as its own account, once per client the script uses.

        Each client asks Entra for a new token, as :meth:`TenantContext.probe` does, so a
        cached sign-in that Entra has stopped renewing opens the window here, where it may.
        With *consent*, the Graph sign-in opens the window whatever the cache holds and has
        Entra show its consent screen for the script's scopes (azure-auth's ``login()`` option
        of that name): a tenant whose consent is one user's own never shows it again by
        itself, and an administrator can tick "Consent on behalf of your organization" there.
        Ctrl+C while it waits asks whether to skip the tenant or stop the run
        (:func:`_ask_to_skip`).

        Args:
            tenant: The tenant, with its cloud and username filled in.
            scopes: The script's Graph scopes.
            clients: The script's ``CLIENTS``.
            consent: Whether the partner sign-in failed for missing consent, so the Graph
                sign-in should show the consent screen.

        Returns:
            The tenant's own context, signed in to every client, else what went wrong, in
            short.

        Raises:
            KeyboardInterrupt: When the person chooses to stop the run.
        """
        try:
            cloud = get_cloud(tenant.cloud or self.cloud_name)
            auth = AuthContext(
                azure_auth.Tenant(tenant.tenant_id, cloud),
                username=tenant.username,
                client_id=self._settings.client_id or None,
                cache=cast(Literal["memory", "disk"], self._settings.cache),
                cache_path=self._cache_path,
                browser_timeout=self._settings.browser_timeout,
                browser=self._browser,
                before_browser_opens=self._pw_page_opener(tenant),
            )
            for kind in clients:
                with build_client(kind, auth, scopes=scopes, transport=self._transport) as client:
                    # The consent screen is Graph's; a sign-in that showed it got its token
                    # from Entra already.
                    wants_consent = consent and kind == "graph"
                    client.login(consent=wants_consent, fresh=not wants_consent)
            # Written now, so a run killed later keeps the sign-in it may have asked for.
            auth.flush_cache()
        except KeyboardInterrupt:
            _ask_to_skip(tenant.display_name)
            return "skipped with Ctrl+C"
        except Exception as exc:
            logger.debug("The direct sign-in failed", exc_info=True)
            why = brief(exc)
            logger.warning("Could not sign in as %s: %s", tenant.username, why)
            return why
        return auth

    def _pw_page_opener(self, tenant: Tenant) -> Callable[[], None] | None:
        """Build what opens a tenant's password page with its first sign-in window.

        Args:
            tenant: The tenant, about to sign in directly.

        Returns:
            A function for azure-auth's ``before_browser_opens``, which opens the page the
            first time a window opens and does nothing after, so a script with several clients
            opens it once; or None when the tenant has no ``pw_url``.
        """
        if not tenant.pw_url:
            return None
        opened = False

        def open_once() -> None:
            nonlocal opened
            if not opened:
                opened = True
                _open_pw_page(tenant, self._pw_browser)

        return open_once

    def context_for(
        self, tenant: Tenant, *, scopes: Sequence[Scope] = (), clients: Sequence[str] = ("graph",)
    ) -> TenantContext:
        """Build the context a script receives for one customer tenant.

        Args:
            tenant: The customer tenant.
            scopes: The script's Graph scopes.
            clients: The script's ``CLIENTS``.

        Returns:
            The context, with one client per declared kind: on the route
            :meth:`check_access` found, asking for the scopes it chose there, or through the
            partner sign-in for a tenant it did not check, asking for the first scope in each
            tuple. No network call is made.

        Raises:
            Exception: Why :meth:`check_access` could not reach the tenant.
        """
        route = self._routes.get(tenant.tenant_id.lower())
        if isinstance(route, Exception):
            raise route
        chosen = self._scopes.get(tenant.tenant_id.lower()) or choose_scopes(scopes, None)
        return TenantContext(
            tenant,
            route or self.auth.get_tenant_context(tenant.tenant_id),
            scopes=chosen,
            clients=clients,
            transport=self._transport,
        )
