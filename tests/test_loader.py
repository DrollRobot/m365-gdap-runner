"""Unit tests for checking a script against the contract in loader.py."""

from __future__ import annotations

import textwrap
from types import ModuleType

import pytest

from m365_gdap_runner.loader import PayloadInvalid, Scope, choose_scopes, payload_from_module

pytestmark = pytest.mark.unit

VALID = """
    SCOPES = ["User.Read.All"]

    def run(ctx):
        return {"tenant": ctx}
"""


def _module(body: str) -> ModuleType:
    module = ModuleType("script")
    exec(textwrap.dedent(body), module.__dict__)  # noqa: S102 (the test's own source)
    return module


def _payload(body: str) -> object:
    return payload_from_module(_module(body), "script")


def test_valid_script() -> None:
    payload = payload_from_module(_module(VALID), "audit")
    assert payload.name == "audit"
    assert payload.scopes == ("User.Read.All",)
    assert payload.clients == ("graph",)
    assert callable(payload.run)
    assert payload.report is None
    assert payload.continue_on_error is False
    assert payload.grant_when_possible is False


def test_report_is_picked_up() -> None:
    module = _module(VALID + "\n    def report(results, title='Report'):\n        pass\n")
    assert payload_from_module(module, "script").report is module.report


def test_continue_on_error_is_picked_up() -> None:
    module = _module(VALID + "\n    CONTINUE_ON_ERROR = True\n")
    assert payload_from_module(module, "script").continue_on_error is True


def test_grant_when_possible_is_picked_up() -> None:
    module = _module(VALID + "\n    GRANT_WHEN_POSSIBLE = True\n")
    assert payload_from_module(module, "script").grant_when_possible is True


def test_clients_are_normalized_and_deduplicated() -> None:
    body = """
        CLIENTS = ["Exchange", "graph", " exchange ", "ipps"]
        SCOPES = ["User.Read.All"]

        def run(ctx):
            return None
    """
    assert payload_from_module(_module(body), "script").clients == ("exchange", "graph", "ipps")


def test_script_without_graph_needs_no_scopes() -> None:
    body = """
        CLIENTS = ["exchange"]

        def run(ctx):
            return None
    """
    payload = payload_from_module(_module(body), "script")
    assert (payload.clients, payload.scopes) == (("exchange",), ())


def test_empty_scopes_need_none_in_particular() -> None:
    body = "SCOPES = []\n\ndef run(ctx):\n    return None\n"
    assert payload_from_module(_module(body), "script").scopes == ()


def test_run_with_optional_parameters_is_accepted() -> None:
    body = 'SCOPES = ("User.Read.All",)\n\ndef run(ctx, verbose=False):\n    return None\n'
    assert payload_from_module(_module(body), "script").scopes == ("User.Read.All",)


@pytest.mark.parametrize(
    ("body", "message"),
    [
        ('SCOPES = ["User.Read.All"]\n', "Script 'script' must define a callable run"),
        ('SCOPES = ["User.Read.All"]\nrun = 1\n', "callable run"),
        ('SCOPES = ["User.Read.All"]\ndef run():\n    pass\n', "one positional argument"),
        ('SCOPES = ["User.Read.All"]\ndef run(ctx, other):\n    pass\n', "one positional argument"),
        ("def run(ctx):\n    pass\n", "must define SCOPES"),
        ('SCOPES = "User.Read.All"\ndef run(ctx):\n    pass\n', "must define SCOPES"),
        ('SCOPES = ["User.Read.All", 1]\ndef run(ctx):\n    pass\n', "non-empty string"),
        ('SCOPES = [" "]\ndef run(ctx):\n    pass\n', "non-empty string"),
        (
            'SCOPES = ["https://graph.microsoft.com/.default"]\ndef run(ctx):\n    pass\n',
            "must not use",
        ),
        (
            'SCOPES = ["User.Read.All", "offline_access"]\ndef run(ctx):\n    pass\n',
            "automatically",
        ),
        ('CLIENTS = []\nSCOPES = ["User.Read.All"]\ndef run(ctx):\n    pass\n', "non-empty list"),
        (
            'CLIENTS = "graph"\nSCOPES = ["User.Read.All"]\ndef run(ctx):\n    pass\n',
            "non-empty list",
        ),
        ('CLIENTS = ["teams"]\ndef run(ctx):\n    pass\n', "unknown client 'teams'"),
        ("CLIENTS = [1]\ndef run(ctx):\n    pass\n", "unknown client 1"),
        (
            'CLIENTS = ["exchange"]\nSCOPES = ["User.Read.All"]\ndef run(ctx):\n    pass\n',
            "'graph' is not in CLIENTS",
        ),
        ('CLIENTS = ["graph", "exchange"]\ndef run(ctx):\n    pass\n', "must define SCOPES"),
        (
            'SCOPES = ["User.Read.All"]\ndef run(ctx):\n    pass\nreport = "summary.xlsx"\n',
            r"must define a callable report\(results\)",
        ),
        (
            'SCOPES = ["User.Read.All"]\ndef run(ctx):\n    pass\ndef report():\n    pass\n',
            r"report\(\) must accept exactly one positional argument \(results\)",
        ),
        (
            'SCOPES = ["User.Read.All"]\ndef run(ctx):\n    pass\nCONTINUE_ON_ERROR = "yes"\n',
            "CONTINUE_ON_ERROR must be True or False, not 'yes'",
        ),
        (
            'SCOPES = ["User.Read.All"]\ndef run(ctx):\n    pass\nGRANT_WHEN_POSSIBLE = 1\n',
            "GRANT_WHEN_POSSIBLE must be True or False, not 1",
        ),
    ],
)
def test_invalid_scripts_are_rejected(body: str, message: str) -> None:
    with pytest.raises(PayloadInvalid, match=message):
        _payload(body)


def test_a_tuple_entry_lists_scopes_any_of_which_will_do() -> None:
    body = """
        SCOPES = [
            "User.Read.All",
            (" Application.Read.All", "Application.ReadWrite.All "),
            ["Mail.Read"],
        ]

        def run(ctx):
            return None
    """
    assert payload_from_module(_module(body), "script").scopes == (
        "User.Read.All",
        ("Application.Read.All", "Application.ReadWrite.All"),
        ("Mail.Read",),
    )


@pytest.mark.parametrize(
    ("body", "message"),
    [
        ("SCOPES = [()]\ndef run(ctx):\n    pass\n", "at least one scope"),
        ('SCOPES = [("User.Read.All", 1)]\ndef run(ctx):\n    pass\n', "non-empty string"),
        (
            'SCOPES = [("User.Read.All", "https://graph.microsoft.com/.default")]\n'
            "def run(ctx):\n    pass\n",
            "must not use",
        ),
        ('SCOPES = [("User.Read.All", "openid")]\ndef run(ctx):\n    pass\n', "automatically"),
    ],
)
def test_invalid_tuple_entries_are_rejected(body: str, message: str) -> None:
    with pytest.raises(PayloadInvalid, match=message):
        _payload(body)


EITHER = ("Application.Read.All", "Application.ReadWrite.All")


@pytest.mark.parametrize(
    ("scopes", "granted", "chosen"),
    [
        (["User.Read.All", EITHER], None, ("User.Read.All", "Application.Read.All")),
        ([EITHER], {"User.Read.All"}, ("Application.Read.All",)),
        ([EITHER], {"Application.ReadWrite.All"}, ("Application.ReadWrite.All",)),
        (
            [EITHER],
            {"Application.ReadWrite.All", "Application.Read.All"},
            ("Application.Read.All",),
        ),
        ([EITHER], {"application.readwrite.all"}, ("Application.ReadWrite.All",)),
        (
            [tuple(f"https://graph.microsoft.com/{scope}" for scope in EITHER)],
            {"Application.ReadWrite.All"},
            ("https://graph.microsoft.com/Application.ReadWrite.All",),
        ),
        (["User.Read.All"], {"Application.ReadWrite.All"}, ("User.Read.All",)),
    ],
    ids=["unknown", "neither", "second", "both", "case", "qualified", "plain"],
)
def test_choose_scopes_takes_the_first_scope_the_tenant_holds(
    scopes: list[Scope], granted: set[str] | None, chosen: tuple[str, ...]
) -> None:
    assert choose_scopes(scopes, granted) == chosen
