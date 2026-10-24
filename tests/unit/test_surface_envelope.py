from __future__ import annotations

from xml.etree.ElementTree import ParseError

from kitchensink4ppt.core.errors import PptMcpError
from kitchensink4ppt.surface import copy, envelope


def test_engine_error_keeps_closed_code_hint_and_details() -> None:
    exc = PptMcpError("the batch refused")
    exc.code = "STALE_ANCHOR"
    exc.hint = "refresh the deck view"
    exc.detail = {"failures": [{"index": 2}]}

    payload = envelope.refusal(exc)

    assert payload == {
        "ok": False,
        "error": {
            "code": "STALE_ANCHOR",
            "message": "the batch refused",
            "hint": "refresh the deck view",
            "details": {"failures": [{"index": 2}]},
        },
    }


def test_unknown_declared_code_is_clamped_to_engine_classification() -> None:
    exc = PptMcpError("bad operation")
    exc.code = "MADE_UP"

    assert envelope.refusal(exc)["error"]["code"] == "BAD_PARAMS"


def test_hostile_parser_is_bad_params_but_attribute_failure_is_internal() -> None:
    assert envelope.refusal(ParseError("broken xml"))["error"]["code"] == "BAD_PARAMS"
    assert envelope.refusal(AttributeError("missing"))["error"]["code"] == "INTERNAL_ERROR"


def test_unexpected_failure_does_not_echo_private_exception_text() -> None:
    payload = envelope.refusal(RuntimeError("secret implementation path"))

    assert payload["error"]["code"] == "INTERNAL_ERROR"
    assert payload["error"]["message"] == copy.INTERNAL_ERROR_MESSAGE
    assert "secret implementation path" not in str(payload)
