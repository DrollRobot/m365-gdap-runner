"""Unit tests for checking a script against the contract in loader.py."""

from __future__ import annotations

import textwrap
from types import ModuleType

import pytest

from m365_gdap_runner.loader import PayloadInvalid, payload_from_module

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
        ("SCOPES = []\ndef run(ctx):\n    pass\n", "must define SCOPES"),
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
    ],
)
def test_invalid_scripts_are_rejected(body: str, message: str) -> None:
    with pytest.raises(PayloadInvalid, match=message):
        _payload(body)
