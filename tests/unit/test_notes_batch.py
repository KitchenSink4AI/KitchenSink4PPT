"""Speaker notes for many slides in one call: one backup, one save.

Field report: a deck-wide notes pass took one set_notes call per slide, each
carrying its whole note and each making its own backup and save, because
apply_edits could not batch notes. Notes now batch two ways: a set_notes
edit inside apply_edits, and set_notes with a list (notes). Both name every
slide by slide_id or slide_number, never by where the entry sits in the
list and never by a bare 0-based index, so a long pass cannot shift onto
the wrong slides in silence. Every address is resolved before anything is
written; one bad entry refuses the whole call and the file is unchanged.
"""

from __future__ import annotations

import asyncio
import hashlib
from pathlib import Path

import pytest

from kitchensink4ppt import server  # noqa: F401  (the import registers ops)
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import notes as _nt
from kitchensink4ppt.ops.read import slide_table
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.help import help_payload
from kitchensink4ppt.surface.plan import build


def _plan():
    return build(Settings(), public_ops())


def _call(tool: str, op: str, args: dict):
    return asyncio.run(dispatch(_plan(), tool, {"op": op, "args": args}))


def _ok(tool: str, op: str, args: dict) -> dict:
    payload, is_error = _call(tool, op, args)
    assert is_error is False, payload
    return payload


def _refused(tool: str, op: str, args: dict) -> dict:
    payload, is_error = _call(tool, op, args)
    assert is_error is True, payload
    return payload["error"]


def _md5(path) -> str:
    return hashlib.md5(Path(path).read_bytes()).hexdigest()


def _ids(deck) -> list[int]:
    return [rec["slide_id"] for rec in slide_table(PptxPackage(deck))]


def _notes(deck, index: int):
    return _nt.get_notes(PptxPackage(deck), index, rich=False)["text"]


@pytest.fixture()
def counted_saves(monkeypatch):
    """How many times a package was saved during the test."""
    count = {"saves": 0}
    real = PptxPackage.save

    def save(self, *args, **kwargs):
        count["saves"] += 1
        return real(self, *args, **kwargs)

    monkeypatch.setattr(PptxPackage, "save", save)
    return count


# ------------------------------------------------------------ apply_edits


def test_apply_edits_writes_notes_on_many_slides_in_one_save(make_deck, counted_saves):
    deck = make_deck("notes_batch.pptx")
    ids = _ids(deck)
    payload = _ok("ppt_edit", "apply_edits", {"file_path": str(deck), "edits": [
        {"op": "set_notes", "slide_id": ids[1], "text": "Second slide notes."},
        {"op": "set_notes", "slide_number": 3, "text": "Third slide notes."},
        {"op": "set_notes", "slide": {"slide_id": ids[5]}, "text": "Sixth\nslide."},
    ]})
    assert counted_saves["saves"] == 1
    assert payload["applied"] == 3
    changed = payload["changed"]
    assert changed["0"]["slide_number"] == 2 and changed["0"]["slide_id"] == ids[1]
    assert changed["1"]["slide_number"] == 3
    assert changed["2"]["slide_number"] == 6
    assert _notes(deck, 1) == "Second slide notes."
    assert _notes(deck, 2) == "Third slide notes."
    assert _notes(deck, 5) == "Sixth\nslide."


def test_one_backup_for_the_whole_batch(make_deck):
    """The prev slot holds the deck as it was before the batch: one
    rotation, not one per slide."""
    deck = make_deck("notes_backup.pptx")
    before = _md5(deck)
    ids = _ids(deck)
    _ok("ppt_edit", "apply_edits", {"file_path": str(deck), "edits": [
        {"op": "set_notes", "slide_id": sid, "text": f"Note {n}."}
        for n, sid in enumerate(ids, 1)
    ]})
    listing = _ok("ppt_deck", "manage_backups", {"file_path": str(deck), "action": "list"})
    listing = listing.get("changed", listing)
    prev = [b for b in listing["slots"] if b.get("slot") == "prev"]
    assert prev, listing
    assert _md5(prev[0]["path"]) == before


def test_structured_notes_batch_too(make_deck):
    deck = make_deck("notes_struct.pptx")
    _ok("ppt_edit", "apply_edits", {"file_path": str(deck), "edits": [
        {"op": "set_notes", "slide_number": 2, "paragraphs": [
            {"runs": [{"text": "Bold lead", "bold": True}, {"text": " then plain."}]},
        ]},
    ]})
    rich = _nt.get_notes(PptxPackage(deck), 1)
    assert rich["text"] == "Bold lead then plain."
    assert rich["paragraphs"][0]["runs"][0]["bold"] is True


def test_a_bad_entry_refuses_the_whole_batch_and_writes_nothing(make_deck):
    deck = make_deck("notes_atomic.pptx")
    before = _md5(deck)
    error = _refused("ppt_edit", "apply_edits", {"file_path": str(deck), "edits": [
        {"op": "set_notes", "slide_number": 2, "text": "Fine."},
        {"op": "set_notes", "slide_id": 99999, "text": "No such slide."},
    ]})
    assert "99999" in error["message"]
    assert _md5(deck) == before


def test_a_bare_slide_index_is_refused_for_notes(make_deck):
    deck = make_deck("notes_index.pptx")
    before = _md5(deck)
    error = _refused("ppt_edit", "apply_edits", {"file_path": str(deck), "edits": [
        {"op": "set_notes", "slide": 2, "text": "Which slide?"},
    ]})
    assert "slide_id" in error["message"] and "slide_number" in error["message"]
    assert _md5(deck) == before


def test_a_notes_edit_with_no_slide_is_refused(make_deck):
    deck = make_deck("notes_noslide.pptx")
    error = _refused("ppt_edit", "apply_edits", {"file_path": str(deck), "edits": [
        {"op": "set_notes", "text": "Nowhere."},
    ]})
    assert "slide_id" in error["message"]


def test_two_notes_edits_for_one_slide_are_refused(make_deck):
    deck = make_deck("notes_twice.pptx")
    before = _md5(deck)
    ids = _ids(deck)
    error = _refused("ppt_edit", "apply_edits", {"file_path": str(deck), "edits": [
        {"op": "set_notes", "slide_number": 2, "text": "First."},
        {"op": "set_notes", "slide_id": ids[1], "text": "Second."},
    ]})
    assert "slide_number 2" in error["message"]
    assert _md5(deck) == before


def test_text_and_paragraphs_together_are_refused(make_deck):
    deck = make_deck("notes_both.pptx")
    error = _refused("ppt_edit", "apply_edits", {"file_path": str(deck), "edits": [
        {"op": "set_notes", "slide_number": 1, "text": "x",
         "paragraphs": [{"runs": [{"text": "y"}]}]},
    ]})
    assert "text" in error["message"] and "paragraphs" in error["message"]


def test_a_flattening_write_abandons_the_batch(make_deck):
    """Notes with bold refuse a plain overwrite until flatten=true, inside a
    batch as alone, and the batch then leaves the file unchanged."""
    deck = make_deck("notes_flatten.pptx")
    _ok("ppt_edit", "set_notes", {"file_path": str(deck), "slide": 1, "live": "off",
                                  "paragraphs": [{"runs": [{"text": "Bold", "bold": True}]}]})
    before = _md5(deck)
    error = _refused("ppt_edit", "apply_edits", {"file_path": str(deck), "edits": [
        {"op": "set_notes", "slide_number": 1, "text": "Plain one."},
        {"op": "set_notes", "slide_number": 2, "text": "Plain two."},
    ]})
    assert "flatten" in error["message"]
    assert _md5(deck) == before
    _ok("ppt_edit", "apply_edits", {"file_path": str(deck), "edits": [
        {"op": "set_notes", "slide_number": 2, "text": "Plain two.", "flatten": True},
    ]})
    assert _notes(deck, 1) == "Plain two."


# ------------------------------------------------------- set_notes(notes=)


def test_set_notes_takes_a_list(make_deck, counted_saves):
    deck = make_deck("notes_list.pptx")
    ids = _ids(deck)
    payload = _ok("ppt_edit", "set_notes", {"file_path": str(deck), "live": "off", "notes": [
        {"slide_id": ids[4], "text": "Fifth."},
        {"slide_number": 2, "text": "Second."},
    ]})
    assert counted_saves["saves"] == 1
    rows = payload["changed"]["notes"]
    assert [row["slide_number"] for row in rows] == [5, 2]
    assert payload["changed"]["applied"] == 2
    assert _notes(deck, 4) == "Fifth." and _notes(deck, 1) == "Second."


def test_the_list_is_never_matched_by_position(make_deck):
    deck = make_deck("notes_pos.pptx")
    before = _md5(deck)
    error = _refused("ppt_edit", "set_notes", {"file_path": str(deck), "live": "off", "notes": [
        {"text": "No slide named."},
    ]})
    assert "slide_id" in error["message"] and "slide_number" in error["message"]
    error = _refused("ppt_edit", "set_notes", {"file_path": str(deck), "live": "off", "notes": [
        {"slide": 0, "text": "A bare index."},
    ]})
    assert "slide_number" in error["message"]
    assert _md5(deck) == before


def test_the_list_refuses_a_single_slide_beside_it(make_deck):
    deck = make_deck("notes_mixed.pptx")
    error = _refused("ppt_edit", "set_notes", {"file_path": str(deck), "live": "off", "slide": 0,
                                               "notes": [{"slide_number": 1, "text": "x"}]})
    assert error["code"] == "BAD_PARAMS"
    assert "notes alone" in error["message"]


def test_a_single_note_still_needs_its_slide(make_deck):
    deck = make_deck("notes_single.pptx")
    error = _refused("ppt_edit", "set_notes", {"file_path": str(deck), "live": "off",
                                               "text": "Where?"})
    assert "slide" in error["message"]


def test_the_list_has_no_live_route_and_says_so_before_powerpoint(make_deck, monkeypatch):
    from kitchensink4ppt.com import live_ops

    def boom(*_a, **_k):  # pragma: no cover - reaching PowerPoint is the failure
        raise AssertionError("the list form reached the live layer")

    monkeypatch.setattr(live_ops, "live_set_notes", boom)
    deck = make_deck("notes_live.pptx")
    error = _refused("ppt_edit", "set_notes", {"file_path": str(deck), "live": "force",
                                               "notes": [{"slide_number": 1, "text": "x"}]})
    assert "notes" in error["message"] and "live" in error["message"]


# ------------------------------------------------------------------ help


def test_help_lists_the_list_form_and_the_batch_edit():
    payload = help_payload(_plan(), "set_notes")
    props = payload["args_schema"]["properties"]
    assert "notes" in props
    assert "slide_id" in props["notes"]["description"]
    assert "slide_number" in props["notes"]["description"]
    assert payload["batchable_in"] == "apply_edits"
    rows = help_payload(_plan(), "apply_edits")["batchable_ops"]
    row = next(row for row in rows if row["op"] == "set_notes")
    assert "slide_id" in row["address"] and "slide_number" in row["address"]
