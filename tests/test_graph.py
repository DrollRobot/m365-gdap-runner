"""Unit tests for GraphClient: versions, paging, retries, and errors (httpx.MockTransport)."""

from __future__ import annotations

from collections.abc import Callable, Sequence

import httpx
import pytest

from m365_gdap_runner import graph
from m365_gdap_runner.graph import MAX_RETRIES, GraphClient, GraphError
from tests._gdap_fakes import TENANT_A, FakeCredential, fake_access

pytestmark = pytest.mark.unit

Handler = Callable[[httpx.Request], httpx.Response]


@pytest.fixture(autouse=True)
def sleeps(monkeypatch: pytest.MonkeyPatch) -> list[float]:
    """Record retry delays instead of sleeping."""
    delays: list[float] = []
    monkeypatch.setattr(graph, "_sleep", delays.append)
    return delays


def _client(handler: Handler, scopes: Sequence[str] = ("User.Read.All",)) -> GraphClient:
    return GraphClient(FakeCredential(TENANT_A), scopes, transport=httpx.MockTransport(handler))


def _sequence(*responses: httpx.Response) -> tuple[Handler, list[httpx.Request]]:
    """Return a handler that replays *responses* in order, and the requests it saw."""
    seen: list[httpx.Request] = []
    remaining = list(responses)

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return remaining.pop(0)

    return handler, seen


def test_version_is_v1_by_default_and_beta_per_call() -> None:
    handler, seen = _sequence(httpx.Response(200, json={}), httpx.Response(200, json={}))
    credential = FakeCredential(TENANT_A)
    with GraphClient(credential, ["User.Read.All"], transport=httpx.MockTransport(handler)) as c:
        c.get("/users")
        c.get("users", beta=True)
    assert [r.url.path for r in seen] == ["/v1.0/users", "/beta/users"]
    assert seen[0].headers["Authorization"] == f"Bearer {fake_access(TENANT_A)}"
    assert credential.calls == [("User.Read.All",), ("User.Read.All",)]


def test_get_all_follows_next_links_without_resending_params() -> None:
    next_link = "https://graph.microsoft.com/v1.0/users?$skiptoken=page2"
    handler, seen = _sequence(
        httpx.Response(
            200, json={"value": [{"id": "1"}, {"id": "2"}], "@odata.nextLink": next_link}
        ),
        httpx.Response(200, json={"value": [{"id": "3"}]}),
    )
    items = _client(handler).get_all(
        "/users",
        params={"$filter": "accountEnabled eq true"},
        headers={"ConsistencyLevel": "eventual"},
    )
    assert [item["id"] for item in items] == ["1", "2", "3"]
    assert dict(seen[0].url.params) == {"$filter": "accountEnabled eq true"}
    assert dict(seen[1].url.params) == {"$skiptoken": "page2"}
    assert all(r.headers["ConsistencyLevel"] == "eventual" for r in seen)


def test_get_all_requires_a_collection() -> None:
    handler, _ = _sequence(httpx.Response(200, json={"id": "not-a-collection"}))
    with pytest.raises(ValueError, match="did not return a collection"):
        _client(handler).get_all("/me")


def test_429_waits_retry_after_then_succeeds(sleeps: list[float]) -> None:
    handler, seen = _sequence(
        httpx.Response(429, headers={"Retry-After": "7"}),
        httpx.Response(200, json={"ok": True}),
    )
    assert _client(handler).get("/users") == {"ok": True}
    assert len(seen) == 2
    assert sleeps == [7.0]


def test_get_retries_503_with_exponential_backoff(sleeps: list[float]) -> None:
    handler, _ = _sequence(
        httpx.Response(503), httpx.Response(503), httpx.Response(200, json={"ok": True})
    )
    assert _client(handler).get("/users") == {"ok": True}
    assert sleeps == [1.0, 2.0]


def test_post_is_not_retried_on_503(sleeps: list[float]) -> None:
    handler, seen = _sequence(httpx.Response(503))
    with pytest.raises(GraphError) as info:
        _client(handler).post("/groups", json={"displayName": "x"})
    assert info.value.status == 503
    assert len(seen) == 1
    assert sleeps == []


def test_post_is_retried_on_429() -> None:
    handler, seen = _sequence(httpx.Response(429), httpx.Response(201, json={"id": "new"}))
    assert _client(handler).post("/groups", json={"displayName": "x"}) == {"id": "new"}
    assert len(seen) == 2
    assert seen[1].content == b'{"displayName":"x"}'


def test_retries_stop_after_max_retries(sleeps: list[float]) -> None:
    handler, seen = _sequence(*[httpx.Response(429) for _ in range(MAX_RETRIES + 1)])
    with pytest.raises(GraphError) as info:
        _client(handler).get("/users")
    assert info.value.status == 429
    assert len(seen) == MAX_RETRIES + 1
    assert len(sleeps) == MAX_RETRIES


def test_transport_error_is_retried_for_get_but_not_post() -> None:
    calls: list[str] = []

    def flaky(request: httpx.Request) -> httpx.Response:
        calls.append(request.method)
        if len(calls) == 1 or request.method == "POST":
            raise httpx.ConnectError("connection reset", request=request)
        return httpx.Response(200, json={"ok": True})

    client = _client(flaky)
    assert client.get("/users") == {"ok": True}
    with pytest.raises(httpx.ConnectError):
        client.post("/groups", json={})
    assert calls == ["GET", "GET", "POST"]


def test_graph_error_carries_graph_details() -> None:
    body = {"error": {"code": "Authorization_RequestDenied", "message": "Insufficient privileges"}}
    handler, _ = _sequence(httpx.Response(403, json=body, headers={"request-id": "rid-1"}))
    with pytest.raises(GraphError) as info:
        _client(handler).get("/users")
    error = info.value
    assert (error.status, error.code, error.message, error.request_id) == (
        403,
        "Authorization_RequestDenied",
        "Insufficient privileges",
        "rid-1",
    )
    assert str(error) == (
        "Graph 403 Authorization_RequestDenied: Insufficient privileges (request-id rid-1)"
    )


def test_refuses_to_send_token_to_non_graph_url() -> None:
    handler, seen = _sequence()
    with pytest.raises(ValueError, match="non-Graph URL"):
        _client(handler).get("https://attacker.example/v1.0/users")
    assert seen == []


def test_client_without_scopes_explains_the_fix() -> None:
    handler, seen = _sequence()
    with pytest.raises(ValueError, match="no Graph scopes"):
        _client(handler, scopes=()).get("/users")
    assert seen == []


def test_patch_and_delete_with_empty_bodies() -> None:
    handler, seen = _sequence(httpx.Response(204), httpx.Response(204))
    client = _client(handler)
    assert client.patch("/users/1", json={"jobTitle": "x"}, beta=True) is None
    client.delete("/users/1")
    assert [(r.method, r.url.path) for r in seen] == [
        ("PATCH", "/beta/users/1"),
        ("DELETE", "/v1.0/users/1"),
    ]


def test_context_manager_closes_the_client() -> None:
    handler, _ = _sequence()
    with _client(handler) as client:
        pass
    with pytest.raises(RuntimeError, match="closed"):
        client.get("/users")
