"""Unit tests for error_text.py: the short text a tenant's failure gets."""

from __future__ import annotations

import pytest
from azure_auth import AuthError, ConsentRequired, InteractionRequired

from m365_gdap_runner.error_text import MAX_LENGTH, brief
from tests._gdap_fakes import TENANT_A

pytestmark = pytest.mark.unit

GRANT_SCOPE = "DelegatedPermissionGrant.ReadWrite.All"
REQUESTED = ("User.Read.All", "Mail.Read", "Policy.Read.All")

# How Microsoft ends a sign-in error.
TRACE = (
    " Trace ID: 00000000-0000-0000-0000-000000000001 Correlation ID: "
    "00000000-0000-0000-0000-000000000002 Timestamp: 2026-10-22 01:50:37Z"
)

# Entra's answer when sign-in frequency wants a fresh sign-in: long, but not too long.
SIGN_IN_FREQUENCY = (
    "invalid_grant: AADSTS70043: The refresh token has expired or is invalid due to sign-in "
    "frequency checks by conditional access. The token was issued on "
    "2026-10-19T02:37:16.9962506Z and the maximum allowed lifetime for this request is 7200."
)

# Entra's answer when the app needs assignment: far too long.
NOT_ASSIGNED = (
    "invalid_grant: AADSTS50105: Your administrator has configured the application Microsoft "
    "Graph Command Line Tools ('14d82eec-204b-4c2f-b7e8-296a70dab67e') to block users unless "
    "they are specifically granted ('assigned') access to the application.  The signed in user "
    "'{EUII Hidden}' is blocked because they are not a direct member of a group with access, "
    "nor had access directly assigned by an administrator. Please contact your administrator "
    "to assign access to this application."
)


def _refusal(granted: tuple[str, ...], missing: tuple[str, ...]) -> ConsentRequired:
    """A Graph client's refusal, as azure-auth builds it: a long message, and fields."""
    return ConsentRequired(
        f"Tenant {TENANT_A} has not granted {', '.join(missing)} to application ...",
        tenant_id=TENANT_A,
        scopes=[f"https://graph.microsoft.com/{scope}" for scope in REQUESTED],
        granted=granted,
        missing=missing,
    )


@pytest.mark.acceptance
def test_no_consent_at_all_lists_no_scope() -> None:
    assert brief(_refusal((), REQUESTED)) == "no Graph consent"


@pytest.mark.acceptance
def test_missing_scopes_are_counted_not_listed() -> None:
    refusal = _refusal(("User.Read.All",), ("Mail.Read", "Policy.Read.All"))
    assert brief(refusal) == "missing 2 of 3 Graph scopes"


@pytest.mark.parametrize(
    ("granted", "grant_when_possible", "ending"),
    [
        (("User.Read.All",), True, "can't self-grant"),
        (("User.Read.All",), False, "can't self-grant"),
        (("User.Read.All", GRANT_SCOPE), True, "self-grant failed"),
        (("User.Read.All", GRANT_SCOPE.lower()), False, "--grant-when-possible would add them"),
    ],
)
def test_missing_scopes_say_what_a_self_grant_can_do(
    granted: tuple[str, ...], grant_when_possible: bool, ending: str
) -> None:
    refusal = _refusal(granted, ("Mail.Read", "Policy.Read.All"))
    assert brief(refusal, grant_when_possible=grant_when_possible) == (
        f"missing 2 of 3 Graph scopes, {ending}"
    )


def test_a_grant_no_token_carried_yet_says_it_was_made() -> None:
    refusal = _refusal((GRANT_SCOPE,), ("Mail.Read", "Policy.Read.All"))
    waited = ConsentRequired(
        "Mail.Read, Policy.Read.All added to the grant ...",
        tenant_id=TENANT_A,
        scopes=refusal.scopes,
        granted=refusal.granted,
        missing=refusal.missing,
    )
    # azure-auth raises it from the last refusal, when no token carried the scopes in time.
    waited.__cause__ = refusal
    assert brief(waited, grant_when_possible=True) == "added 2 Graph scopes, no token had them yet"


def test_a_consent_error_from_entra_is_its_message() -> None:
    # A sign-in's own consent error carries no fields to count.
    error = ConsentRequired("invalid_grant: AADSTS65001: Not consented." + TRACE, tenant_id="t")
    assert brief(error) == "invalid_grant: AADSTS65001: Not consented"


@pytest.mark.acceptance
def test_microsofts_ids_are_dropped_and_the_rest_kept() -> None:
    error = InteractionRequired(SIGN_IN_FREQUENCY + TRACE, tenant_id=TENANT_A)
    assert brief(error) == SIGN_IN_FREQUENCY.removesuffix(".")


def test_ids_on_their_own_lines_are_dropped() -> None:
    error = AuthError("AADSTS65001: Not consented.\r\nTrace ID: 0\r\nCorrelation ID: 0")
    assert brief(error) == "AADSTS65001: Not consented"


@pytest.mark.acceptance
def test_an_excessively_long_message_is_cut_to_its_first_sentence() -> None:
    assert len(NOT_ASSIGNED) > MAX_LENGTH
    assert brief(AuthError(NOT_ASSIGNED + TRACE)) == (
        "invalid_grant: AADSTS50105: Your administrator has configured the application Microsoft "
        "Graph Command Line Tools ('14d82eec-204b-4c2f-b7e8-296a70dab67e') to block users unless "
        "they are specifically granted ('assigned') access to the application"
    )


def test_a_long_message_with_no_sentence_to_cut_at_is_kept() -> None:
    text = "x" * (MAX_LENGTH + 1)
    assert brief(AuthError(text)) == text


@pytest.mark.parametrize(
    ("error", "text"),
    [
        (RuntimeError("boom"), "RuntimeError: boom"),
        (KeyError("id"), "KeyError: 'id'"),
        (RuntimeError(), "RuntimeError"),
        (ValueError("Not a tenant."), "ValueError: Not a tenant"),
    ],
)
def test_any_other_exception_is_its_type_and_message(error: Exception, text: str) -> None:
    assert brief(error) == text


def test_another_exceptions_long_message_is_cut_too() -> None:
    error = RuntimeError("First sentence. " + "y" * MAX_LENGTH)
    assert brief(error) == "RuntimeError: First sentence"
