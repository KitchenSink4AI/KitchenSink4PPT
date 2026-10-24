"""ppt_help apply_edits names the edits apply_edits can batch.

Field report: ppt_help apply_edits listed no edit ops; the list only
appeared in a refusal, and ppt_help set_text answered NOT_FOUND (set_text is
an edit apply_edits batches, not an operation of its own). The help now
lists every batchable op with its fields and how it names its target,
generated from the table apply_edits enforces, so the two cannot drift.
"""

from __future__ import annotations

import pytest

from kitchensink4ppt import server  # noqa: F401 - populates the registry
from kitchensink4ppt.ops import batch
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.help import help_payload
from kitchensink4ppt.surface.plan import build


def _plan():
    return build(Settings(), public_ops())


def test_apply_edits_help_lists_every_batchable_op_with_its_fields():
    payload = help_payload(_plan(), "apply_edits")
    rows = payload["batchable_ops"]
    assert [row["op"] for row in rows] == list(batch._OPS)
    for row in rows:
        allowed, required = batch._OPS[row["op"]]
        assert row["fields"] == sorted(allowed)
        assert row["required"] == sorted(required)
        assert row["address"], row
        assert row["example"]["op"] == row["op"]


def test_the_edits_argument_names_them_too():
    payload = help_payload(_plan(), "apply_edits")
    text = payload["args_schema"]["properties"]["edits"]["description"]
    for name in batch._OPS:
        assert name in text, (name, text)


def test_the_list_comes_from_the_table_apply_edits_enforces(monkeypatch):
    """Generated, not copied: an op added to the batch table shows up in
    the help with nothing else edited."""
    monkeypatch.setitem(batch._OPS, "set_notes_line", ({"text"}, {"text"}))
    monkeypatch.setitem(batch._ADDRESS, "set_notes_line", "slide")
    monkeypatch.setitem(batch._EXAMPLE, "set_notes_line", {"slide": 0, "text": "x"})
    payload = help_payload(_plan(), "apply_edits")
    assert "set_notes_line" in [row["op"] for row in payload["batchable_ops"]]
    assert "set_notes_line" in payload["args_schema"]["properties"]["edits"]["description"]


def test_every_batchable_op_has_an_address_and_a_valid_example():
    assert set(batch._ADDRESS) == set(batch._OPS)
    assert set(batch._EXAMPLE) == set(batch._OPS)
    for name, edit in batch._EXAMPLE.items():
        allowed, required = batch._OPS[name]
        keys = set(edit)
        assert required <= keys, name
        assert keys <= allowed | batch._LOCATION_KEYS | {"slide_number"}, name


def test_help_for_set_text_points_to_apply_edits():
    payload = help_payload(_plan(), "set_text")
    assert payload["batch_only"] is True
    assert payload["use"] == "apply_edits"
    assert payload["fields"] == ["text"]
    assert payload["example"]["op"] == "apply_edits"
    assert payload["example"]["args"]["edits"][0]["op"] == "set_text"


@pytest.mark.parametrize("op", ["format_text", "set_shape", "delete_shape",
                                "set_table_cells", "set_placeholder_text",
                                "search_and_replace"])
def test_an_op_that_also_batches_says_so(op):
    assert help_payload(_plan(), op)["batchable_in"] == "apply_edits"


def test_an_op_that_does_not_batch_says_nothing_about_it():
    assert "batchable_in" not in help_payload(_plan(), "insert_textbox")
