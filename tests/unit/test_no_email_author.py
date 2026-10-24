"""An email address is never written into a deck as an author's name.

Comments and replies already keep to this (test_comment_authorship.py). The
other places a name lands in a deck keep to it too: anonymize_deck's
replacement, which becomes every comment author and the deck's creator and
last-modified-by. An address sent there is set aside, the op's own default
stands in, and the answer says so in one sentence that never repeats the
address.
"""

from __future__ import annotations

import asyncio
import importlib

import pytest

from kitchensink4ppt import server  # noqa: F401  (registers the operations)
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import comments as cm
from kitchensink4ppt.ops.assembly import ANONYMIZE_DEFAULT
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.plan import build

CORE = "docProps/core.xml"
DC = "{http://purl.org/dc/elements/1.1/}"
CP = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
EMAILS = ["jane.doe@example.com", "Jane Doe <jane.doe@example.com>",
          "mailto:jane.doe@example.com", "  jd@example.com  "]


@pytest.fixture
def office(monkeypatch):
    for var in ("KS4P_AUTHOR", "KS4P_COMMENT_AUTHOR"):
        monkeypatch.delenv(var, raising=False)
    monkeypatch.setenv("USERNAME", "jdoe")
    values = {"UserName": "Jane Doe", "UserInitials": "JD"}
    module = importlib.import_module("kitchensink4ppt.core.authorship")
    monkeypatch.setattr(module, "_office_value", lambda name: values.get(name))
    return values


def _answer(op: str, args: dict) -> dict:
    plan = build(Settings())
    payload, is_error = asyncio.run(
        dispatch(plan, plan.specs[op].group, {"op": op, "args": args}))
    assert is_error is False, payload
    return payload


def _email_notes(answer: dict) -> list[str]:
    return [w for w in answer.get("warnings") or [] if "email address" in w]


def _names_in(deck) -> list[str]:
    pkg = PptxPackage(deck)
    names = [e.get("name") for e in
             pkg.root("ppt/authors.xml").findall(f"{{{cm.NS_P188}}}author")]
    core = pkg.root(CORE)
    for tag in (f"{DC}creator", f"{CP}lastModifiedBy"):
        el = core.find(tag)
        if el is not None:
            names.append(el.text or "")
    return names


@pytest.mark.parametrize("sent", EMAILS)
def test_anonymize_never_writes_an_email_address(make_deck, office, sent):
    """anonymize_deck(replacement="jane.doe@example.com") wrote
    "jane.doe@example.com-1" as every comment author and the address as the
    deck's creator and last-modified-by."""
    deck = make_deck("anon_email.pptx")
    _answer("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this",
                            "author": "Alex Kim"})
    answer = _answer("anonymize_deck", {"file_path": str(deck), "replacement": sent})
    names = _names_in(deck)
    assert names and all("@" not in n for n in names), names
    assert f"{ANONYMIZE_DEFAULT}-1" in names
    assert answer["changed"]["replacement"] == ANONYMIZE_DEFAULT
    notes = _email_notes(answer)
    assert len(notes) == 1, answer.get("warnings")
    assert f'"{ANONYMIZE_DEFAULT}"' in notes[0] and "@" not in notes[0]


def test_anonymize_with_a_name_says_nothing_about_email(make_deck, office):
    deck = make_deck("anon_name.pptx")
    _answer("add_comment", {"file_path": str(deck), "slide": 0, "text": "Check this",
                            "author": "Alex Kim"})
    answer = _answer("anonymize_deck", {"file_path": str(deck), "replacement": "Panel"})
    assert "Panel-1" in _names_in(deck)
    assert answer["changed"]["replacement"] == "Panel"
    assert _email_notes(answer) == []


# The deck's own author properties: set_document_properties writes the
# creator (author) and cp:lastModifiedBy (last_modified_by). An email address
# sent for either is set aside; the name a comment would be signed with
# stands in, and the answer says so in one sentence.


@pytest.mark.parametrize("field, tag", [("author", f"{DC}creator"),
                                        ("last_modified_by", f"{CP}lastModifiedBy")])
@pytest.mark.parametrize("sent", EMAILS)
def test_the_document_author_is_never_an_email_address(make_deck, office, field, tag, sent):
    deck = make_deck("props_email.pptx")
    answer = _answer("set_document_properties", {
        "file_path": str(deck), "properties": {field: sent, "title": "Quarterly plan"}})
    core = PptxPackage(deck).root(CORE)
    assert core.find(tag).text == "Jane Doe"
    assert core.find(f"{DC}title").text == "Quarterly plan"
    notes = _email_notes(answer)
    assert len(notes) == 1, answer.get("warnings")
    assert '"Jane Doe"' in notes[0] and "@" not in notes[0]
    row = next(c for c in answer["changed"]["changed"] if c["property"] == field)
    assert row["new"] == "Jane Doe" and row["author_source"] == "office_username"


def test_with_the_assistant_setting_the_label_stands_in(make_deck, office, monkeypatch):
    monkeypatch.setenv("KS4P_AUTHOR", "assistant")
    deck = make_deck("props_email_assistant.pptx")
    answer = _answer("set_document_properties", {
        "file_path": str(deck), "properties": {"author": "jane.doe@example.com"}})
    assert PptxPackage(deck).root(CORE).find(f"{DC}creator").text == "AI Assistant via KS4PPT"
    assert len(_email_notes(answer)) == 1


def test_a_name_or_the_word_assistant_is_written_as_given(make_deck, office):
    """The document author is a property, not a signature: like
    KitchenSink4XL's workbook author, "assistant" is written as given."""
    deck = make_deck("props_name.pptx")
    answer = _answer("set_document_properties", {
        "file_path": str(deck), "properties": {"author": "Sam Lee",
                                               "last_modified_by": "assistant"}})
    core = PptxPackage(deck).root(CORE)
    assert core.find(f"{DC}creator").text == "Sam Lee"
    assert core.find(f"{CP}lastModifiedBy").text == "assistant"
    assert _email_notes(answer) == []
