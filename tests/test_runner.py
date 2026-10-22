"""Unit tests for runner.py with a fake session and fake payloads."""

from __future__ import annotations

import io
import logging
import re
import threading
from collections.abc import Callable, Iterator

import pytest
from azure_auth import AuthError, ConsentRequired, InteractionRequired

from m365_gdap_runner.context import TenantContext
from m365_gdap_runner.loader import Payload
from m365_gdap_runner.runner import (
    RunInterrupted,
    TenantLogFilter,
    TenantResult,
    run_payload,
    summarize,
)
from m365_gdap_runner.tenants import Tenant
from tests._gdap_fakes import TENANT_A, TENANT_B, TENANT_C, FakeSession

pytestmark = pytest.mark.unit

ALPHA = Tenant(TENANT_A, "Alpha")
BETA = Tenant(TENANT_B, "Beta")
GAMMA = Tenant(TENANT_C, "Gamma")
DELTA = Tenant("dddddddd-0000-0000-0000-00000000000d", "Delta")
TENANTS = [ALPHA, BETA, GAMMA]

Run = Callable[[TenantContext], object]


def _payload(
    run: Run, scopes: tuple[str, ...] = ("User.Read.All",), clients: tuple[str, ...] = ("graph",)
) -> Payload:
    return Payload(name="test", run=run, scopes=scopes, clients=clients)


def _statuses(results: list[TenantResult]) -> list[str]:
    return [r.status for r in results]


def _consent_error(tenant_id: str) -> ConsentRequired:
    return ConsentRequired(
        "invalid_grant: AADSTS65001 consent_required", tenant_id=tenant_id, scopes=["User.Read.All"]
    )


class Recorder:
    """Payload that records which tenants it ran for and fails for chosen ones."""

    def __init__(self, failures: dict[str, BaseException] | None = None) -> None:
        self.failures = failures or {}
        self.called: list[str] = []
        self.lock = threading.Lock()

    def __call__(self, ctx: TenantContext) -> object:
        with self.lock:
            self.called.append(ctx.tenant.tenant_id)
        failure = self.failures.get(ctx.tenant.tenant_id)
        if failure is not None:
            raise failure
        return {"tenant": ctx.tenant.display_name}


@pytest.fixture
def release() -> Iterator[threading.Event]:
    """Event set by a timer shortly after the test starts, to unblock waiting payloads."""
    event = threading.Event()
    timer = threading.Timer(0.3, event.set)
    timer.start()
    yield event
    timer.cancel()
    event.set()


def test_results_cover_ok_auth_failed_and_error_in_order() -> None:
    session = FakeSession(auth_errors={TENANT_B: _consent_error(TENANT_B)})
    payload = Recorder(failures={TENANT_C: RuntimeError("boom")})
    results = run_payload(session, _payload(payload), TENANTS, continue_on_error=True)
    assert [r.tenant for r in results] == TENANTS
    assert _statuses(results) == ["ok", "auth_failed", "error"]
    assert results[0].result == {"tenant": "Alpha"}
    assert results[0].error is None
    assert results[1].error == "invalid_grant: AADSTS65001 consent_required"
    assert results[2].error == "RuntimeError: boom"


def test_a_tenant_that_failed_logs_an_error(caplog: pytest.LogCaptureFixture) -> None:
    caplog.set_level(logging.INFO, logger="m365_gdap_runner.runner")
    session = FakeSession(auth_errors={TENANT_B: _consent_error(TENANT_B)})
    payload = Recorder(failures={TENANT_C: RuntimeError("boom")})
    run_payload(session, _payload(payload), TENANTS, continue_on_error=True)
    outcomes = [(r.levelno, r.getMessage().split(" ")[0]) for r in caplog.records]
    assert outcomes == [
        (logging.INFO, "ok"),
        (logging.ERROR, "auth_failed:"),
        (logging.ERROR, "error"),
    ]


def test_only_a_tenant_whose_run_started_logs_a_time(caplog: pytest.LogCaptureFixture) -> None:
    caplog.set_level(logging.INFO, logger="m365_gdap_runner.runner")
    unreachable = AuthError("partner: no Graph consent; direct: no m365_username")
    session = FakeSession(auth_errors={TENANT_B: unreachable})
    payload = Recorder(failures={TENANT_C: RuntimeError("boom")})
    run_payload(session, _payload(payload), TENANTS, continue_on_error=True)
    beta, gamma = [r.getMessage() for r in caplog.records if r.levelno == logging.ERROR]
    assert beta == "auth_failed: partner: no Graph consent; direct: no m365_username"
    assert re.fullmatch(r"error in \d+\.\ds: RuntimeError: boom", gamma)


def test_an_auth_error_inside_run_is_short_and_kept_in_full_for_the_debug_log(
    caplog: pytest.LogCaptureFixture,
) -> None:
    caplog.set_level(logging.DEBUG, logger="m365_gdap_runner.runner")
    refusal = ConsentRequired(
        f"Tenant {TENANT_A} has not granted Mail.Read to application ...",
        tenant_id=TENANT_A,
        scopes=[
            "https://graph.microsoft.com/User.Read.All",
            "https://graph.microsoft.com/Mail.Read",
        ],
        granted=["User.Read.All"],
        missing=["Mail.Read"],
    )
    results = run_payload(FakeSession(), _payload(Recorder(failures={TENANT_A: refusal})), [ALPHA])
    assert results[0].error == "missing 1 of 2 Graph scopes"
    (debug,) = [r for r in caplog.records if r.levelno == logging.DEBUG]
    assert debug.exc_info is not None
    assert debug.exc_info[1] is refusal


def test_context_is_built_with_the_payloads_scopes_and_clients_and_closed() -> None:
    session = FakeSession(auth_errors={TENANT_B: _consent_error(TENANT_B)})
    payload = _payload(Recorder(), scopes=("Mail.Read",), clients=("graph", "exchange"))
    run_payload(session, payload, [ALPHA, BETA])
    assert session.requests == [
        (TENANT_A, ("Mail.Read",), ("graph", "exchange")),
        (TENANT_B, ("Mail.Read",), ("graph", "exchange")),
    ]
    assert all(context.closed for context in session.contexts.values())


def test_run_is_not_called_after_auth_failure() -> None:
    session = FakeSession(auth_errors={TENANT_B: _consent_error(TENANT_B)})
    payload = Recorder()
    results = run_payload(session, _payload(payload), TENANTS)
    assert payload.called == [TENANT_A, TENANT_C]
    assert _statuses(results) == ["ok", "auth_failed", "ok"]


def test_interaction_required_is_auth_failed() -> None:
    error = InteractionRequired("MFA wanted", tenant_id=TENANT_A, scopes=["User.Read.All"])
    results = run_payload(FakeSession(auth_errors={TENANT_A: error}), _payload(Recorder()), [ALPHA])
    assert _statuses(results) == ["auth_failed"]
    assert results[0].error == "MFA wanted"


def test_unexpected_token_failure_is_auth_failed() -> None:
    session = FakeSession(auth_errors={TENANT_A: ConnectionError("login host unreachable")})
    results = run_payload(session, _payload(Recorder()), [ALPHA])
    assert _statuses(results) == ["auth_failed"]
    assert results[0].error == "ConnectionError: login host unreachable"


def test_failure_to_build_the_context_is_auth_failed() -> None:
    class BrokenSession(FakeSession):
        def context_for(self, tenant: Tenant, **kwargs: object) -> TenantContext:
            raise AuthError(f"no such tenant {tenant.display_name}")

    results = run_payload(BrokenSession(), _payload(Recorder()), [ALPHA])
    assert _statuses(results) == ["auth_failed"]
    assert results[0].error == "no such tenant Alpha"


def test_auth_error_inside_run_is_auth_failed() -> None:
    payload = Recorder(failures={TENANT_A: _consent_error(TENANT_A)})
    assert _statuses(run_payload(FakeSession(), _payload(payload), [ALPHA])) == ["auth_failed"]


def test_serial_stop_on_error_skips_remaining_tenants() -> None:
    payload = Recorder(failures={TENANT_A: RuntimeError("boom")})
    results = run_payload(FakeSession(), _payload(payload), TENANTS)
    assert _statuses(results) == ["error", "skipped", "skipped"]
    assert payload.called == [TENANT_A]


def test_auth_failures_never_stop_the_run() -> None:
    session = FakeSession(auth_errors={TENANT_A: _consent_error(TENANT_A)})
    results = run_payload(session, _payload(Recorder()), TENANTS)
    assert _statuses(results) == ["auth_failed", "ok", "ok"]


def test_any_result_is_kept_as_it_is() -> None:
    found = object()
    results = run_payload(FakeSession(), _payload(lambda ctx: found), [ALPHA])
    assert _statuses(results) == ["ok"]
    assert results[0].result is found


def test_parallel_runs_tenants_concurrently_in_input_order() -> None:
    barrier = threading.Barrier(4, timeout=5)

    def run(ctx: TenantContext) -> object:
        barrier.wait()
        return ctx.tenant.display_name

    tenants = [ALPHA, BETA, GAMMA, DELTA]
    results = run_payload(FakeSession(), _payload(run), tenants, jobs=4)
    assert _statuses(results) == ["ok"] * 4
    assert [r.result for r in results] == ["Alpha", "Beta", "Gamma", "Delta"]


def test_serial_run_is_not_concurrent() -> None:
    barrier = threading.Barrier(2, timeout=0.2)

    def run(ctx: TenantContext) -> object:
        barrier.wait()
        return None

    results = run_payload(FakeSession(), _payload(run), [ALPHA, BETA], continue_on_error=True)
    assert _statuses(results) == ["error", "error"]


def test_parallel_stop_on_error_lets_running_tenants_finish_and_starts_no_more(
    release: threading.Event,
) -> None:
    beta_started = threading.Event()

    def run(ctx: TenantContext) -> object:
        if ctx.tenant is ALPHA:
            assert beta_started.wait(5)
            raise RuntimeError("boom")
        if ctx.tenant is BETA:
            beta_started.set()
            release.wait(5)
        return None

    results = run_payload(FakeSession(), _payload(run), [ALPHA, BETA, GAMMA, DELTA], jobs=2)
    assert _statuses(results) == ["error", "ok", "skipped", "skipped"]


def test_tenants_start_in_the_order_ready_hands_them_over() -> None:
    payload = Recorder()
    results = run_payload(FakeSession(), _payload(payload), TENANTS, ready=[GAMMA, ALPHA, BETA])
    assert payload.called == [TENANT_C, TENANT_A, TENANT_B]
    assert [r.tenant for r in results] == TENANTS


@pytest.mark.acceptance
def test_a_tenant_runs_while_ready_finds_the_next(release: threading.Event) -> None:
    asked_for_beta = threading.Event()

    def ready() -> Iterator[Tenant]:
        yield ALPHA
        # A sign-in for the next tenant, say: Alpha's run() must not wait for it.
        asked_for_beta.set()
        yield BETA

    def run(ctx: TenantContext) -> object:
        if ctx.tenant is ALPHA:
            assert asked_for_beta.wait(5)
        return None

    results = run_payload(FakeSession(), _payload(run), [ALPHA, BETA], ready=ready())
    assert _statuses(results) == ["ok", "ok"]


def test_no_tenant_ready_hands_over_after_an_error_starts() -> None:
    payload = Recorder(failures={TENANT_A: RuntimeError("boom")})
    alpha_done = threading.Event()

    def ready() -> Iterator[Tenant]:
        for tenant in TENANTS:
            yield tenant
            if tenant is ALPHA:
                # Alpha's run() has failed by the time Beta is handed over.
                assert alpha_done.wait(5)

    def run(ctx: TenantContext) -> object:
        try:
            return payload(ctx)
        finally:
            alpha_done.set()

    results = run_payload(FakeSession(), _payload(run), TENANTS, ready=ready())
    assert _statuses(results) == ["error", "skipped", "skipped"]
    assert payload.called == [TENANT_A]


def test_an_interrupt_while_ready_finds_the_next_waits_for_running_tenants(
    release: threading.Event,
) -> None:
    alpha_started = threading.Event()
    closed = threading.Event()

    def ready() -> Iterator[Tenant]:
        try:
            yield ALPHA
            assert alpha_started.wait(5)
            raise KeyboardInterrupt  # Ctrl+C in a sign-in window, say
        finally:
            closed.set()

    def run(ctx: TenantContext) -> object:
        alpha_started.set()
        release.wait(5)  # still running when Ctrl+C comes
        return None

    with pytest.raises(RunInterrupted) as info:
        run_payload(FakeSession(), _payload(run), TENANTS, ready=ready(), jobs=2)
    assert _statuses(info.value.results) == ["ok", "skipped", "skipped"]
    assert closed.is_set()


def test_serial_interrupt_returns_partial_results() -> None:
    payload = Recorder(failures={TENANT_B: KeyboardInterrupt()})
    with pytest.raises(RunInterrupted) as info:
        run_payload(FakeSession(), _payload(payload), TENANTS)
    results = info.value.results
    assert _statuses(results) == ["ok", "error", "skipped"]
    assert results[1].error == "interrupted"


def test_parallel_interrupt_waits_for_running_tenants(release: threading.Event) -> None:
    def run(ctx: TenantContext) -> object:
        if ctx.tenant is ALPHA:
            raise KeyboardInterrupt
        release.wait(5)
        return None

    with pytest.raises(RunInterrupted) as info:
        run_payload(FakeSession(), _payload(run), [ALPHA, BETA, GAMMA, DELTA], jobs=2)
    results = info.value.results
    assert results[0].status == "error"
    assert results[0].error == "interrupted"
    assert results[1].status == "ok"
    assert results[3].status == "skipped"


def test_jobs_must_be_positive() -> None:
    with pytest.raises(ValueError, match="at least 1"):
        run_payload(FakeSession(), _payload(Recorder()), TENANTS, jobs=0)


def test_summarize_counts_every_status() -> None:
    results = [
        TenantResult(ALPHA, "ok"),
        TenantResult(BETA, "ok"),
        TenantResult(GAMMA, "skipped"),
    ]
    assert summarize(results) == "ok=2, auth_failed=0, error=0, skipped=1"


def test_log_records_name_their_tenant() -> None:
    stream = io.StringIO()
    handler = logging.StreamHandler(stream)
    handler.addFilter(TenantLogFilter())
    handler.setFormatter(logging.Formatter("[%(tenant)s] %(message)s"))
    root = logging.getLogger()
    previous_level = root.level
    root.addHandler(handler)
    root.setLevel(logging.INFO)
    try:

        def run(ctx: TenantContext) -> object:
            logging.getLogger("some.payload").info("hello")
            return None

        run_payload(FakeSession(), _payload(run), [ALPHA])
        logging.getLogger("outside").info("after")
    finally:
        root.removeHandler(handler)
        root.setLevel(previous_level)
    lines = stream.getvalue().splitlines()
    assert "[Alpha] hello" in lines
    assert any(line.startswith("[Alpha] ok in ") for line in lines)
    assert "[-] after" in lines
