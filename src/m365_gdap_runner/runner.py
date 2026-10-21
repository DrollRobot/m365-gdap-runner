"""Run one script across many tenants.

Access is settled per tenant before its ``run()``
(:meth:`~m365_gdap_runner.session.PartnerSession.check_access`): every declared client already
holds a token for the tenant, if it could be reached. ``check_access()`` hands the tenants
over as ``ready``, one by one, and each starts as soon as it is handed over, while this thread
goes back for the next. For each tenant the runner builds a
:class:`~m365_gdap_runner.context.TenantContext` and calls ``run()``; a tenant that could not
be reached becomes an ``auth_failed`` result and ``run()`` never starts.

Consent is not authorization: a consented scope can still get 403 if the GDAP
relationship lacks the needed role. That surfaces inside ``run()`` as a resource
error and becomes an ``error`` result.

Tenants run serially or on a thread pool (``jobs``); with ``ready``, always on the pool, so
nothing ``ready`` does to find the next tenant, such as a sign-in, waits for a ``run()``.
Every input tenant gets exactly one :class:`TenantResult`, in input order. After an error
stops the run, no tenant starts, but tenants already running finish and are recorded.
Thread safety relies on azure-auth's locks around its token cache and on one
context (with its own clients) per tenant; nothing else is shared.
"""

from __future__ import annotations

import logging
import threading
import time
from collections import Counter
from collections.abc import Generator, Iterable, Sequence
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from contextvars import ContextVar
from dataclasses import dataclass
from typing import Literal, Protocol

from azure_auth import AuthError

from m365_gdap_runner.context import TenantContext
from m365_gdap_runner.loader import Payload
from m365_gdap_runner.tenants import Tenant

logger = logging.getLogger(__name__)

Status = Literal["ok", "auth_failed", "error", "skipped"]

# Every status, in summary order.
STATUSES: tuple[Status, ...] = ("ok", "auth_failed", "error", "skipped")

# Display name of the tenant the current thread is processing, for log records.
_current_tenant: ContextVar[str | None] = ContextVar("m365_gdap_runner_tenant", default=None)


class ContextFactory(Protocol):
    """What the runner needs from a session: a context per tenant."""

    def context_for(
        self, tenant: Tenant, *, scopes: Sequence[str] = (), clients: Sequence[str] = ("graph",)
    ) -> TenantContext:
        """Build the context a payload receives for one tenant."""
        ...


class TenantLogFilter(logging.Filter):
    """Logging filter that adds ``%(tenant)s`` to every record.

    Attach it to a handler so interleaved records from parallel tenants,
    including payload and library records, name their tenant (``-`` outside one).
    """

    def filter(self, record: logging.LogRecord) -> bool:
        """Set ``record.tenant`` and keep the record.

        Args:
            record: The log record.

        Returns:
            Always True.
        """
        record.__dict__["tenant"] = _current_tenant.get() or "-"
        return True


@dataclass(frozen=True)
class TenantResult:
    """Outcome of running the payload against one tenant.

    Attributes:
        tenant: The tenant.
        status: ``ok``, ``auth_failed``, ``error``, or ``skipped``.
        result: What the script's ``run()`` returned when ``ok``, as it is; else None.
        error: Failure description when not ``ok`` or ``skipped``, else None.
    """

    tenant: Tenant
    status: Status
    result: object = None
    error: str | None = None


class RunInterrupted(Exception):
    """Raised when Ctrl+C stops a run; carries the partial results.

    Attributes:
        results: One result per tenant; tenants that never ran are ``skipped``.
    """

    def __init__(self, results: list[TenantResult]) -> None:
        """Store the partial results.

        Args:
            results: One result per tenant, in input order.
        """
        self.results = results
        super().__init__("Run interrupted.")


def summarize(results: Sequence[TenantResult]) -> str:
    """Return a one-line count of results per status.

    Args:
        results: Per-tenant results.

    Returns:
        ``ok=N, auth_failed=N, error=N, skipped=N``.
    """
    counts = Counter(r.status for r in results)
    return ", ".join(f"{status}={counts[status]}" for status in STATUSES)


def run_payload(
    session: ContextFactory,
    payload: Payload,
    tenants: Sequence[Tenant],
    *,
    ready: Iterable[Tenant] | None = None,
    continue_on_error: bool = False,
    jobs: int = 1,
) -> list[TenantResult]:
    """Run *payload* against every tenant.

    Args:
        session: A signed-in session.
        payload: The validated payload.
        tenants: Tenants to run against, in output order.
        ready: The same tenants, in the order they may start, each handed over when its
            access is settled, as ``PartnerSession.check_access()`` yields them. None starts
            them in input order. A generator is closed when the run stops early.
        continue_on_error: Keep going after a payload error. Auth failures
            never stop the run.
        jobs: Maximum tenants processed concurrently.

    Returns:
        One result per tenant, in input order.

    Raises:
        ValueError: If *jobs* is less than 1.
        RunInterrupted: On Ctrl+C, carrying the partial results.
    """
    if jobs < 1:
        raise ValueError("jobs must be at least 1.")
    if ready is None and (jobs == 1 or len(tenants) <= 1):
        return _run_serial(session, payload, tenants, continue_on_error=continue_on_error)
    return _run_parallel(
        session,
        payload,
        tenants,
        ready=tenants if ready is None else ready,
        jobs=max(1, min(jobs, len(tenants))),
        continue_on_error=continue_on_error,
    )


def _run_serial(
    session: ContextFactory,
    payload: Payload,
    tenants: Sequence[Tenant],
    *,
    continue_on_error: bool,
) -> list[TenantResult]:
    """Run tenants one at a time.

    Args:
        session: A signed-in session.
        payload: The validated payload.
        tenants: Tenants in output order.
        continue_on_error: Keep going after a payload error.

    Returns:
        One result per tenant; tenants after a stopping error are ``skipped``.

    Raises:
        RunInterrupted: On Ctrl+C; the interrupted tenant is recorded as ``error``.
    """
    results: list[TenantResult] = []
    for tenant in tenants:
        try:
            tenant_result = _run_one(session, payload, tenant)
        except KeyboardInterrupt:
            results.append(TenantResult(tenant, "error", error="interrupted"))
            raise RunInterrupted(_fill_skipped(results, tenants)) from None
        results.append(tenant_result)
        if tenant_result.status == "error" and not continue_on_error:
            logger.warning("Stopping after an error; remaining tenants are skipped.")
            break
    return _fill_skipped(results, tenants)


def _run_parallel(
    session: ContextFactory,
    payload: Payload,
    tenants: Sequence[Tenant],
    *,
    ready: Iterable[Tenant],
    jobs: int,
    continue_on_error: bool,
) -> list[TenantResult]:
    """Run tenants on a thread pool, each as soon as *ready* hands it over.

    This thread takes the tenants from *ready* while the pool runs the ones it has.

    Args:
        session: A signed-in session.
        payload: The validated payload.
        tenants: Tenants in output order.
        ready: The same tenants, in the order they may start.
        jobs: Worker threads.
        continue_on_error: Keep going after a payload error.

    Returns:
        One result per tenant, in input order; tenants that never started are ``skipped``.

    Raises:
        RunInterrupted: On Ctrl+C, after in-flight tenants finish.
    """
    position = {tenant.tenant_id.lower(): index for index, tenant in enumerate(tenants)}
    done: dict[int, TenantResult] = {}
    # Set by the worker whose tenant's error stops the run, before it takes another tenant.
    stopped = threading.Event()
    executor = ThreadPoolExecutor(max_workers=jobs, thread_name_prefix="tenant")
    futures: dict[Future[TenantResult], int] = {}
    handing_over = iter(ready)
    interrupted = False
    try:
        for tenant in handing_over:
            if stopped.is_set():
                break
            future = executor.submit(
                _run_unless_stopped, session, payload, tenant, stopped, continue_on_error
            )
            futures[future] = position[tenant.tenant_id.lower()]
        for future in as_completed(futures):
            index = futures[future]
            done[index] = _result_or_interrupted(future, tenants[index])
            if done[index].error == "interrupted":
                raise KeyboardInterrupt
            if done[index].status == "error" and not continue_on_error:
                logger.warning(
                    "Stopping after an error; running tenants finish, pending ones are skipped."
                )
                break
    except KeyboardInterrupt:
        interrupted = True
    finally:
        # A generator, such as check_access(), stops: no more checks or sign-in windows.
        if isinstance(handing_over, Generator):
            handing_over.close()
        # Cancel tenants that have not started. as_completed() is never notified about
        # futures cancelled this way, so wait on the running ones directly.
        executor.shutdown(wait=False, cancel_futures=True)
        for future, index in futures.items():
            if index not in done and not future.cancelled():
                done[index] = _result_or_interrupted(future, tenants[index])
        executor.shutdown(wait=True)
    results = _ordered(done, tenants)
    if interrupted:
        raise RunInterrupted(results)
    return results


def _result_or_interrupted(future: Future[TenantResult], tenant: Tenant) -> TenantResult:
    """Wait for a tenant's future, recording a KeyboardInterrupt as an ``error``.

    Args:
        future: The tenant's future.
        tenant: The tenant, for the interrupted result.

    Returns:
        The tenant's result, or an ``error`` result with error ``interrupted``.
    """
    try:
        return future.result()
    except KeyboardInterrupt:
        return TenantResult(tenant, "error", error="interrupted")


def _fill_skipped(results: list[TenantResult], tenants: Sequence[Tenant]) -> list[TenantResult]:
    """Append ``skipped`` results for tenants after the recorded ones.

    Args:
        results: Results for a prefix of *tenants*.
        tenants: All tenants, in order.

    Returns:
        One result per tenant.
    """
    return results + [TenantResult(t, "skipped") for t in tenants[len(results) :]]


def _ordered(done: dict[int, TenantResult], tenants: Sequence[Tenant]) -> list[TenantResult]:
    """Order results by tenant index, marking tenants without a result ``skipped``.

    Args:
        done: Results keyed by tenant index.
        tenants: All tenants, in order.

    Returns:
        One result per tenant.
    """
    return [
        done[index] if index in done else TenantResult(tenant, "skipped")
        for index, tenant in enumerate(tenants)
    ]


def _run_unless_stopped(
    session: ContextFactory,
    payload: Payload,
    tenant: Tenant,
    stopped: threading.Event,
    continue_on_error: bool,
) -> TenantResult:
    """Process one tenant unless an error has stopped the run, and stop it on an error.

    Args:
        session: A signed-in session.
        payload: The validated payload.
        tenant: The tenant.
        stopped: Set once an error stops the run.
        continue_on_error: Keep going after a payload error.

    Returns:
        The tenant's result; ``skipped`` when the run stopped before it started.
    """
    if stopped.is_set():
        return TenantResult(tenant, "skipped")
    result = _run_one(session, payload, tenant)
    if result.status == "error" and not continue_on_error:
        stopped.set()
    return result


def _run_one(session: ContextFactory, payload: Payload, tenant: Tenant) -> TenantResult:
    """Process one tenant and log its outcome: an error when the tenant failed.

    Args:
        session: A signed-in session.
        payload: The validated payload.
        tenant: The tenant.

    Returns:
        The tenant's result.
    """
    context_token = _current_tenant.set(tenant.display_name)
    start = time.monotonic()
    try:
        status, result, error = _execute(session, payload, tenant)
        duration = time.monotonic() - start
        if error:
            logger.error("%s in %.1fs: %s", status, duration, error)
        else:
            logger.info("%s in %.1fs", status, duration)
        return TenantResult(tenant, status, result, error)
    finally:
        _current_tenant.reset(context_token)


def _execute(
    session: ContextFactory, payload: Payload, tenant: Tenant
) -> tuple[Status, object, str | None]:
    """Build the context and call ``run()``.

    Args:
        session: A signed-in session.
        payload: The validated payload.
        tenant: The tenant.

    Returns:
        ``(status, result, error)``.
    """
    try:
        context = session.context_for(tenant, scopes=payload.scopes, clients=payload.clients)
    except AuthError as exc:
        return "auth_failed", None, str(exc)
    except Exception as exc:
        logger.debug("Building the tenant context failed", exc_info=True)
        return "auth_failed", None, repr(exc)

    with context:
        try:
            result = payload.run(context)
        except AuthError as exc:
            return "auth_failed", None, str(exc)
        except Exception as exc:
            logger.debug("Payload raised", exc_info=True)
            return "error", None, repr(exc)
    return "ok", result, None
