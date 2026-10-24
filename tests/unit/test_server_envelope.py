"""Operation envelopes and the fixed-surface refusal vocabulary."""

from __future__ import annotations

import pytest

from kitchensink4ppt import server
from kitchensink4ppt.core.errors import CODES
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface import envelope
from kitchensink4ppt.surface.plan import MAX_DESCRIPTION_CHARS, build

_CLOSED_CODES = CODES


def _fn(name: str):
    fn = getattr(server, name)

    def call(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except BaseException as exc:
            return envelope.refusal(exc)

    return call


def test_mutation_success_envelope(make_deck):
    deck = make_deck("env.pptx")
    out = _fn("set_placeholder_text")(
        file_path=str(deck), slide=0, placeholder="title", text="New Title"
    )
    assert out["ok"] is True
    assert out["file"] == str(deck)
    assert out["saved"]
    assert out["backup"] is True
    assert isinstance(out["changed"], dict)
    assert isinstance(out["warnings"], list)


def test_refusal_envelope_not_found(make_deck):
    deck = make_deck("env2.pptx")
    out = _fn("delete_slide")(file_path=str(deck), slide=99)
    assert out["ok"] is False
    err = out["error"]
    assert err["code"] == "NOT_FOUND"
    assert err["code"] in _CLOSED_CODES
    assert "99" in err["message"]  # actionable, names the bad index
    assert err["hint"]


def test_refusal_envelope_bad_params(make_deck):
    deck = make_deck("env3.pptx")
    out = _fn("list_elements")(file_path=str(deck), kind="widgets")
    assert out["ok"] is False
    assert out["error"]["code"] == "BAD_PARAMS"
    assert "widgets" in out["error"]["message"]


def test_missing_file_refusal():
    out = _fn("get_presentation_info")(file_path="Z:/nope/missing.pptx")
    assert out["ok"] is False
    assert out["error"]["code"] in _CLOSED_CODES


def test_all_mapped_codes_are_closed_vocabulary():
    for _etype, code in envelope._CODE_MAP:
        assert code in _CLOSED_CODES, f"{code} is outside the closed vocabulary"


def test_internal_hint_tools_do_not_reintroduce_runtime_surface_advice(make_deck):
    """Legacy hint metadata cannot tell callers to mutate the fixed surface."""
    from kitchensink4ppt.core.errors import UnsupportedStructure

    exc = UnsupportedStructure(
        "table cell text goes through set_table_cells, not format_text"
    )
    exc.hint_tools = ["set_table_cells"]
    out = envelope.refusal(exc)
    assert out["ok"] is False
    hint = out["error"]["hint"]
    assert hint
    assert "enable_tools" not in hint
    assert "pack" not in hint.lower()


def test_declared_hint_overrides_generic_fixed_surface_hint():
    from kitchensink4ppt.core.errors import UnsupportedStructure

    exc = UnsupportedStructure(
        "table cell text goes through set_table_cells"
    )
    exc.hint = "Use set_table_cells for table-cell content."
    out = envelope.refusal(exc)
    assert out["error"]["hint"] == exc.hint


def test_format_text_cell_refusal_names_the_correct_operation(make_deck):
    """The real refusal path points at the operation that can do the job."""
    deck = make_deck("cellref.pptx")
    from kitchensink4ppt.core.package import PptxPackage
    from kitchensink4ppt.ops.read import list_elements

    tables = list_elements(PptxPackage(deck), "tables")["items"]
    assert tables
    t = tables[0]
    out = _fn("format_text")(
        file_path=str(deck), slide=t["slide_index"], shape=t["id"],
        bold=True,
    )
    assert out["ok"] is False
    assert "set_table_cells" in out["error"]["message"]
    assert "enable_tools" not in out["error"]["hint"]


def test_fixed_tool_description_budget_and_no_em_dashes():
    plan = build(Settings())
    for name, tool in plan.tools.items():
        desc = tool.description
        assert desc, f"{name} has no description"
        assert "\u2014" not in desc, f"{name} description has an em dash"
        assert len(desc) <= MAX_DESCRIPTION_CHARS


def test_mutating_tools_carry_contract_sentence():
    contract = "Saves atomically with two-slot backup"
    for name, spec in public_ops().items():
        if "backup" in spec.args.model_fields:
            desc = " ".join((getattr(server, name).__doc__ or "").split())
            assert contract in desc, (
                f"{name} takes backup= but lacks the mutation contract "
                "sentence"
            )


def test_fixed_surface_copy_does_not_advertise_retired_controls():
    for name, tool in build(Settings()).tools.items():
        desc = tool.description.lower()
        assert "enable_tools" not in desc, name
        assert "disable_tools" not in desc, name
        assert "tools/list_changed" not in desc, name
