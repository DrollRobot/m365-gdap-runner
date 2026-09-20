"""Run one payload across many tenants.

For each tenant the runner requests a token for exactly the payload's
``SCOPES`` before calling ``run()``. The refresh-token grant refuses to mint a
token when any requested scope is not consented in that tenant, so missing
consent becomes an ``auth_failed`` result and ``run()`` never starts.
(``GraphClient`` fetches tokens lazily, which would be inside ``run()``.)

Consent is not authorization: a consented scope can still get 403 if the GDAP
relationship lacks the needed role. That surfaces inside ``run()`` as a
``GraphError`` and becomes an ``error`` result.

Tenants run serially or on a thread pool (``jobs``). Every input tenant gets
exactly one :class:`TenantResult`, in input order. Stop-on-first-error is
best-effort under parallelism: tenants already running finish and are recorded.
Thread safety relies on MSAL's cache lock, the session's application lock, and
one credential plus one ``GraphClient`` per tenant; nothing else is shared.
"""

from __future__ import annotations

import logging
import time
from collections import Counter
from collections.abc import Sequence
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from contextvars import ContextVar
from dataclasses import dataclass
from typing import Literal

import httpx
from azure.core.credentials import TokenCredential

from m365_gdap_runner.auth import GRAPH_RESOURCE, TenantAuthError, TokenSession, group_by_resource
from m365_gdap_runner.graph import GraphClient
from m365_gdap_runner.loader import Payload
from m365_gdap_runner.output import OutputFormatError, check_result
from m365_gdap_runner.tenants import Tenant, TenantContext

logger = logging.getLogger(__name__)

Status = Literal["ok", "auth_failed", "error", "skipped"]

# Every status, in summary order.
STATUSES: tuple[Status, ...] = ("ok", "auth_failed", "error", "skipped")

# Display name of the tenant the current thread is processing, for log records.
_current_tenant: ContextVar[str | None] = ContextVar("m365_gdap_runner_tenant", default=None)


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
        result: The payload's return value when ``ok``, else None.
        error: Failure description when not ``ok`` or ``skipped``, else None.
        duration_s: Wall time spent on this tenant, in seconds.
    """

    tenant: Tenant
    status: Status
    result: object = None
    error: str | None = None
    duration_s: float = 0.0


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


def probe(credential: TokenCredential, scopes: Sequence[str]) -> None:
    """Request a token for *scopes*, once per resource, to check consent up front.

    Args:
        credential: Credential bound to the tenant.
        scopes: The payload's scopes.

    Raises:
        TenantAuthError: If any scope group cannot be granted in the tenant.
    """
    for group in group_by_resource(scopes).values():
        credential.get_token(*group)


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
    session: TokenSession,
    payload: Payload,
    tenants: Sequence[Tenant],
    *,
    continue_on_error: bool = False,
    jobs: int = 1,
    output_format: str = "json",
    transport: httpx.BaseTransport | None = None,
) -> list[TenantResult]:
    """Run *payload* against every tenant.

    Args:
        session: A signed-in session.
        payload: The validated payload.
        tenants: Tenants to run against, in output order.
        continue_on_error: Keep going after a payload error. Auth failures
            never stop the run.
        jobs: Maximum tenants processed concurrently.
        output_format: Format results must fit (``json`` or ``csv``).
        transport: Optional httpx transport for every Graph client (tests).

    Returns:
        One result per tenant, in input order.

    Raises:
        ValueError: If *jobs* is less than 1.
        RunInterrupted: On Ctrl+C, carrying the partial results.
    """
    if jobs < 1:
        raise ValueError("jobs must be at least 1.")
    if jobs == 1 or len(tenants) <= 1:
        return _run_serial(
            session,
            payload,
            tenants,
            continue_on_error=continue_on_error,
            output_format=output_format,
            transport=transport,
        )
    return _run_parallel(
        session,
        payload,
        tenants,
        jobs=min(jobs, len(tenants)),
        continue_on_error=continue_on_error,
        output_format=output_format,
        transport=transport,
    )


def _run_serial(
    session: TokenSession,
    payload: Payload,
    tenants: Sequence[Tenant],
    *,
    continue_on_error: bool,
    output_format: str,
    transport: httpx.BaseTransport | None,
) -> list[TenantResult]:
    """Run tenants one at a time.

    Args:
        session: A signed-in session.
        payload: The validated payload.
        tenants: Tenants in output order.
        continue_on_error: Keep going after a payload error.
        output_format: Format results must fit.
        transport: Optional httpx transport (tests).

    Returns:
        One result per tenant; tenants after a stopping error are ``skipped``.

    Raises:
        RunInterrupted: On Ctrl+C; the interrupted tenant is recorded as ``error``.
    """
    results: list[TenantResult] = []
    for tenant in tenants:
        try:
            tenant_result = _run_one(
                session, payload, tenant, output_format=output_format, transport=transport
            )
        except KeyboardInterrupt:
            results.append(TenantResult(tenant, "error", error="interrupted"))
            raise RunInterrupted(_fill_skipped(results, tenants)) from None
        results.append(tenant_result)
        if tenant_result.status == "error" and not continue_on_error:
            logger.warning("Stopping after an error; remaining tenants are skipped.")
            break
    return _fill_skipped(results, tenants)


def _run_parallel(
    session: TokenSession,
    payload: Payload,
    tenants: Sequence[Tenant],
    *,
    jobs: int,
    continue_on_error: bool,
    output_format: str,
    transport: httpx.BaseTransport | None,
) -> list[TenantResult]:
    """Run tenants on a thread pool.

    Args:
        session: A signed-in session.
        payload: The validated payload.
        tenants: Tenants in output order.
        jobs: Worker threads.
        continue_on_error: Keep going after a payload error.
        output_format: Format results must fit.
        transport: Optional httpx transport (tests).

    Returns:
        One result per tenant, in input order; cancelled tenants are ``skipped``.

    Raises:
        RunInterrupted: On Ctrl+C, after in-flight tenants finish.
    """
    done: dict[int, TenantResult] = {}
    executor = ThreadPoolExecutor(max_workers=jobs, thread_name_prefix="tenant")
    futures: dict[Future[TenantResult], int] = {
        executor.submit(
            _run_one, session, payload, tenant, output_format=output_format, transport=transport
        ): index
        for index, tenant in enumerate(tenants)
    }
    interrupted = False
    try:
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


def _run_one(
    session: TokenSession,
    payload: Payload,
    tenant: Tenant,
    *,
    output_format: str,
    transport: httpx.BaseTransport | None,
) -> TenantResult:
    """Process one tenant and log its outcome.

    Args:
        session: A signed-in session.
        payload: The validated payload.
        tenant: The tenant.
        output_format: Format the result must fit.
        transport: Optional httpx transport (tests).

    Returns:
        The tenant's result.
    """
    context_token = _current_tenant.set(tenant.display_name)
    start = time.monotonic()
    try:
        status, result, error = _execute(session, payload, tenant, output_format, transport)
        duration = time.monotonic() - start
        if error:
            logger.info("%s in %.1fs: %s", status, duration, error)
        else:
            logger.info("%s in %.1fs", status, duration)
        return TenantResult(tenant, status, result, error, duration)
    finally:
        _current_tenant.reset(context_token)


def _execute(
    session: TokenSession,
    payload: Payload,
    tenant: Tenant,
    output_format: str,
    transport: httpx.BaseTransport | None,
) -> tuple[Status, object, str | None]:
    """Probe consent, call ``run()``, and check the result for one tenant.

    Args:
        session: A signed-in session.
        payload: The validated payload.
        tenant: The tenant.
        output_format: Format the result must fit.
        transport: Optional httpx transport (tests).

    Returns:
        ``(status, result, error)``.
    """
    try:
        credential = session.credential_for(tenant.tenant_id)
        probe(credential, payload.scopes)
    except TenantAuthError as exc:
        return "auth_failed", None, str(exc)
    except Exception as exc:
        logger.debug("Token request failed", exc_info=True)
        return "auth_failed", None, repr(exc)

    graph_scopes = group_by_resource(payload.scopes).get(GRAPH_RESOURCE, [])
    try:
        with GraphClient(credential, graph_scopes, transport=transport) as graph:
            result = payload.run(TenantContext(tenant=tenant, graph=graph, credential=credential))
    except TenantAuthError as exc:
        return "auth_failed", None, str(exc)
    except Exception as exc:
        logger.debug("Payload raised", exc_info=True)
        return "error", None, repr(exc)

    try:
        check_result(result, output_format)
    except OutputFormatError as exc:
        return "error", None, f"invalid result: {exc}"
    return "ok", result, None
