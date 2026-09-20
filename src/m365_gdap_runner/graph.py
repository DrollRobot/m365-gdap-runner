"""Small synchronous Microsoft Graph client.

A convenience for payloads, not an SDK: JSON in, JSON out, with paging,
throttling retries, and bearer tokens from any
``azure.core.credentials.TokenCredential``. Payloads that want typed models
can build their own SDK client from ``ctx.credential``.

The API version is chosen per call (``beta=True``), so one client serves v1.0
and beta in the same payload.
"""

from __future__ import annotations

import logging
import time
from collections.abc import Generator, Mapping, Sequence
from types import TracebackType
from typing import Any, Self

import httpx
from azure.core.credentials import TokenCredential

logger = logging.getLogger(__name__)

# Graph host; request paths are relative to a version under it.
GRAPH_ROOT = "https://graph.microsoft.com"

# Per-request timeout, in seconds.
DEFAULT_TIMEOUT_S = 60.0

# Retries after the first attempt, and the cap on computed backoff.
MAX_RETRIES = 5
MAX_BACKOFF_S = 60.0

# 429 means the request was not processed, so any method may retry.
_ALWAYS_RETRY_STATUSES = frozenset({429})

# 503/504 may follow a request that was applied; only idempotent methods retry.
_IDEMPOTENT_RETRY_STATUSES = frozenset({503, 504})
_IDEMPOTENT_METHODS = frozenset({"GET", "DELETE"})

# Indirection so tests can skip real waits.
_sleep = time.sleep


class GraphError(Exception):
    """Raised for a non-2xx Graph response.

    Attributes:
        status: HTTP status code.
        code: Graph error code (e.g. ``Authorization_RequestDenied``), or ``""``.
        message: Graph error message, or the HTTP reason phrase.
        request_id: The ``request-id`` response header, if present.
    """

    def __init__(self, status: int, code: str, message: str, request_id: str | None) -> None:
        """Store the error details and build a one-line message.

        Args:
            status: HTTP status code.
            code: Graph error code, or ``""``.
            message: Error message.
            request_id: The ``request-id`` response header, if present.
        """
        self.status = status
        self.code = code
        self.message = message
        self.request_id = request_id
        text = f"Graph {status}"
        if code:
            text += f" {code}"
        text += f": {message}"
        if request_id:
            text += f" (request-id {request_id})"
        super().__init__(text)


class _BearerAuth(httpx.Auth):
    """httpx auth flow that sets a fresh bearer token on every request."""

    def __init__(self, credential: TokenCredential, scopes: tuple[str, ...]) -> None:
        """Store the credential and scopes.

        Args:
            credential: Token source.
            scopes: Graph scopes to request.
        """
        self._credential = credential
        self._scopes = scopes

    def auth_flow(self, request: httpx.Request) -> Generator[httpx.Request, httpx.Response]:
        """Attach ``Authorization: Bearer`` and send the request.

        Args:
            request: The outgoing request.

        Yields:
            The request with the header set.

        Raises:
            ValueError: If the client was built without scopes.
        """
        if not self._scopes:
            raise ValueError(
                "GraphClient has no Graph scopes; declare them in the payload's SCOPES."
            )
        token = self._credential.get_token(*self._scopes)
        request.headers["Authorization"] = f"Bearer {token.token}"
        yield request


def _url(path: str, beta: bool) -> str:
    """Resolve a request path to a URL relative to :data:`GRAPH_ROOT`.

    Args:
        path: Path relative to the API version (``"/users"``), or an absolute
            Graph URL such as an ``@odata.nextLink``.
        beta: Use the beta endpoint instead of v1.0. Ignored for absolute URLs.

    Returns:
        ``/v1.0/<path>``, ``/beta/<path>``, or the absolute URL unchanged.

    Raises:
        ValueError: If an absolute URL is not on the Graph host, which would
            leak the bearer token to another host.
    """
    if "://" in path:
        if not path.startswith(GRAPH_ROOT + "/"):
            raise ValueError(f"Refusing to send a Graph token to a non-Graph URL: {path}")
        return path
    version = "beta" if beta else "v1.0"
    return f"/{version}/{path.lstrip('/')}"


def _retryable(method: str, status: int) -> bool:
    """Return whether a response status should be retried for *method*.

    Args:
        method: HTTP method, upper-case.
        status: HTTP status code.

    Returns:
        True for 429 on any method, or 503/504 on GET and DELETE.
    """
    if status in _ALWAYS_RETRY_STATUSES:
        return True
    return status in _IDEMPOTENT_RETRY_STATUSES and method in _IDEMPOTENT_METHODS


def _retry_after(response: httpx.Response) -> float | None:
    """Return the ``Retry-After`` delay in seconds, if present and numeric.

    Args:
        response: The throttled response.

    Returns:
        The delay, or None when the header is missing or not a number.
    """
    value = response.headers.get("Retry-After")
    if value is None:
        return None
    try:
        delay = float(value)
    except ValueError:
        return None
    return delay if delay >= 0 else None


def _backoff(attempt: int) -> float:
    """Return the exponential backoff delay for a zero-based retry *attempt*.

    Args:
        attempt: Number of retries already made.

    Returns:
        ``2 ** attempt`` seconds, capped at :data:`MAX_BACKOFF_S`.
    """
    return min(2.0**attempt, MAX_BACKOFF_S)


def _graph_error(response: httpx.Response) -> GraphError:
    """Build a :class:`GraphError` from an error response.

    Args:
        response: A non-2xx response.

    Returns:
        The error, with Graph's ``error.code``/``error.message`` when the body has them.
    """
    code = ""
    message = response.reason_phrase
    try:
        body = response.json()
    except ValueError:
        body = None
    if isinstance(body, dict) and isinstance(body.get("error"), dict):
        code = str(body["error"].get("code", ""))
        message = str(body["error"].get("message", message))
    return GraphError(response.status_code, code, message, response.headers.get("request-id"))


def _json_object(response: httpx.Response) -> dict[str, Any]:
    """Parse a response body as a JSON object.

    Args:
        response: A successful response.

    Returns:
        The parsed object, or ``{}`` for an empty body.

    Raises:
        GraphError: If the body is not a JSON object.
    """
    if not response.content:
        return {}
    data = response.json()
    if not isinstance(data, dict):
        raise GraphError(
            response.status_code,
            "",
            "Expected a JSON object in the response body.",
            response.headers.get("request-id"),
        )
    return data


class GraphClient:
    """Microsoft Graph client bound to one credential and one set of scopes.

    Not shared across tenants. Use as a context manager, or call :meth:`close`.
    """

    def __init__(
        self,
        credential: TokenCredential,
        scopes: Sequence[str],
        *,
        transport: httpx.BaseTransport | None = None,
        timeout: float = DEFAULT_TIMEOUT_S,
    ) -> None:
        """Create the client.

        Args:
            credential: Token source for this tenant.
            scopes: Graph scopes requested for every call.
            transport: Optional httpx transport; tests pass ``httpx.MockTransport``.
            timeout: Per-request timeout in seconds.
        """
        self._client = httpx.Client(
            base_url=GRAPH_ROOT,
            timeout=timeout,
            auth=_BearerAuth(credential, tuple(scopes)),
            transport=transport,
        )

    def __enter__(self) -> Self:
        """Return the client for use in a ``with`` block."""
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        """Close the client when the ``with`` block exits."""
        self.close()

    def close(self) -> None:
        """Close the underlying HTTP connections."""
        self._client.close()

    def get(
        self,
        path: str,
        params: Mapping[str, str] | None = None,
        headers: Mapping[str, str] | None = None,
        *,
        beta: bool = False,
    ) -> dict[str, Any]:
        """Send a GET request.

        Args:
            path: Path relative to the API version, or an absolute Graph URL.
            params: Query parameters.
            headers: Extra request headers (e.g. ``ConsistencyLevel``).
            beta: Call the beta endpoint instead of v1.0.

        Returns:
            The response body as a JSON object.

        Raises:
            GraphError: On a non-2xx response after retries.
        """
        return _json_object(self._request("GET", path, beta, params=params, headers=headers))

    def get_all(
        self,
        path: str,
        params: Mapping[str, str] | None = None,
        headers: Mapping[str, str] | None = None,
        *,
        beta: bool = False,
    ) -> list[dict[str, Any]]:
        """GET a collection and follow ``@odata.nextLink`` until exhausted.

        Args:
            path: Collection path relative to the API version.
            params: Query parameters for the first request only; next links
                already carry the query.
            headers: Extra request headers, sent on every page.
            beta: Call the beta endpoint instead of v1.0.

        Returns:
            Every item from every page's ``value``.

        Raises:
            GraphError: On a non-2xx response after retries.
            ValueError: If a page has no ``value`` list.
        """
        items: list[dict[str, Any]] = []
        page = self.get(path, params, headers, beta=beta)
        while True:
            value = page.get("value")
            if not isinstance(value, list):
                raise ValueError(f"{path} did not return a collection (no 'value' list).")
            items.extend(value)
            next_link = page.get("@odata.nextLink")
            if not next_link:
                return items
            page = self.get(str(next_link), headers=headers)

    def post(
        self,
        path: str,
        json: object = None,
        headers: Mapping[str, str] | None = None,
        *,
        beta: bool = False,
    ) -> dict[str, Any] | None:
        """Send a POST request. Not retried on 503/504.

        Args:
            path: Path relative to the API version.
            json: JSON-serializable request body.
            headers: Extra request headers.
            beta: Call the beta endpoint instead of v1.0.

        Returns:
            The response body as a JSON object, or None for an empty body.

        Raises:
            GraphError: On a non-2xx response after retries.
        """
        response = self._request("POST", path, beta, json=json, headers=headers)
        return _json_object(response) if response.content else None

    def patch(
        self,
        path: str,
        json: object = None,
        headers: Mapping[str, str] | None = None,
        *,
        beta: bool = False,
    ) -> dict[str, Any] | None:
        """Send a PATCH request. Not retried on 503/504.

        Args:
            path: Path relative to the API version.
            json: JSON-serializable request body.
            headers: Extra request headers.
            beta: Call the beta endpoint instead of v1.0.

        Returns:
            The response body as a JSON object, or None for an empty body.

        Raises:
            GraphError: On a non-2xx response after retries.
        """
        response = self._request("PATCH", path, beta, json=json, headers=headers)
        return _json_object(response) if response.content else None

    def delete(
        self,
        path: str,
        headers: Mapping[str, str] | None = None,
        *,
        beta: bool = False,
    ) -> None:
        """Send a DELETE request.

        Args:
            path: Path relative to the API version.
            headers: Extra request headers.
            beta: Call the beta endpoint instead of v1.0.

        Raises:
            GraphError: On a non-2xx response after retries.
        """
        self._request("DELETE", path, beta, headers=headers)

    def _request(
        self,
        method: str,
        path: str,
        beta: bool,
        *,
        params: Mapping[str, str] | None = None,
        headers: Mapping[str, str] | None = None,
        json: object = None,
    ) -> httpx.Response:
        """Send one request, retrying throttling and transient failures.

        Args:
            method: HTTP method, upper-case.
            path: Path relative to the API version, or an absolute Graph URL.
            beta: Call the beta endpoint instead of v1.0.
            params: Query parameters.
            headers: Extra request headers.
            json: JSON request body, or None for no body.

        Returns:
            The successful response.

        Raises:
            GraphError: On a non-2xx response that is not retried or runs out
                of retries.
            httpx.TransportError: On a network failure that is not retried or
                runs out of retries.
        """
        url = _url(path, beta)
        attempt = 0
        while True:
            try:
                response = self._client.request(
                    method, url, params=params, headers=headers, json=json
                )
            except httpx.TransportError as exc:
                if method not in _IDEMPOTENT_METHODS or attempt >= MAX_RETRIES:
                    raise
                delay = _backoff(attempt)
                logger.warning(
                    "Graph %s %s failed (%r); retry %d in %.1fs",
                    method,
                    url,
                    exc,
                    attempt + 1,
                    delay,
                )
            else:
                if attempt >= MAX_RETRIES or not _retryable(method, response.status_code):
                    if response.is_error:
                        raise _graph_error(response)
                    return response
                retry_after = _retry_after(response)
                delay = retry_after if retry_after is not None else _backoff(attempt)
                logger.warning(
                    "Graph %s %s returned %d; retry %d in %.1fs",
                    method,
                    url,
                    response.status_code,
                    attempt + 1,
                    delay,
                )
            attempt += 1
            _sleep(delay)
