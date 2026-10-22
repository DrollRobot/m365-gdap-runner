"""Short error text: why a tenant failed, in a line, for its result and the terminal.

The full error still reaches the debug log, with its traceback. The short text differs from
it in three ways:

- A consent error from azure-auth's Graph client counts the scopes the tenant lacks instead
  of listing them. Its fields say which ones they are, so no text is read.
- Microsoft's ``Trace ID: … Correlation ID: … Timestamp: …``, which ends every sign-in error
  and only matters to Microsoft support, is dropped.
- A message still longer than :data:`MAX_LENGTH` characters is cut to its first sentence.

An azure-auth error is its message; any other exception is ``Type: message``. Neither ends
with a full stop, like the log lines they go in.
"""

from __future__ import annotations

import re

from azure_auth import AuthError, ConsentRequired
from azure_auth.clients.graph import GRANT_SCOPE

# Past this many characters, an error message is cut to its first sentence.
MAX_LENGTH = 300

# The IDs Microsoft ends a sign-in error with.
_TRACE_IDS = re.compile(r"\s*Trace ID: \S+(\s+Correlation ID: \S+)?(\s+Timestamp: \S+ \S+)?")


def brief(exc: BaseException, *, grant_when_possible: bool | None = None) -> str:
    """Say in a line why something failed.

    Args:
        exc: The error.
        grant_when_possible: Whether the Graph client that raised *exc* was allowed to add
            the scopes a tenant lacks to the app's grant. When given, a consent error says
            whether that failed, could not be tried, or would help. None says nothing of it.

    Returns:
        The text.
    """
    if isinstance(exc, ConsentRequired) and exc.missing:
        return _consent(exc, grant_when_possible)
    text = _TRACE_IDS.sub("", str(exc)).strip()
    if len(text) > MAX_LENGTH:
        text = _first_sentence(text)
    text = text.removesuffix(".")
    if isinstance(exc, AuthError):
        return text
    name = type(exc).__name__
    return f"{name}: {text}" if text else name


def _consent(exc: ConsentRequired, grant_when_possible: bool | None) -> str:
    """Say how much of the requested Graph consent a tenant lacks, from the error's fields.

    Args:
        exc: The Graph client's refusal, with the scopes it found granted and missing.
        grant_when_possible: As for :func:`brief`.

    Returns:
        The text.
    """
    if not exc.granted:
        return "no Graph consent"
    if isinstance(exc.__cause__, ConsentRequired):
        # azure-auth added the scopes to the grant, but Entra had not put them in a token yet.
        return f"added {len(exc.missing)} Graph scopes, no token had them yet"
    text = f"missing {len(exc.missing)} of {len(exc.scopes)} Graph scopes"
    if grant_when_possible is None:
        return text
    if GRANT_SCOPE.lower() not in {scope.lower() for scope in exc.granted}:
        return f"{text}, can't self-grant"
    if grant_when_possible:
        return f"{text}, self-grant failed"
    return f"{text}, --grant-when-possible would add them"


def _first_sentence(text: str) -> str:
    """Cut a message after its first sentence.

    Args:
        text: The message.

    Returns:
        Everything up to the first full stop followed by a space, or all of *text* when it
        has none.
    """
    end = text.find(". ")
    return text if end < 0 else text[: end + 1]
