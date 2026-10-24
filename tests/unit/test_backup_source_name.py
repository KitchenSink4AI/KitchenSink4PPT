"""manage_backups names the backup to restore `source`, and says so everywhere.

Field report: manage_backups restore {"slot": "anchor"} was refused (the
argument is `source`), while the list answer calls prev and anchor
"slots". `source` stays the one documented name, because it also takes a
snapshot's path, which is not a slot. `slot` is read as `source`, and the
answer says so; both sent with different values is refused, naming both.
Every listed slot and snapshot carries the `source` value that restores it.
"""

from __future__ import annotations

import asyncio
from pathlib import Path

from lxml import etree

from kitchensink4ppt import server  # noqa: F401 - populates the registry
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import backups as bk
from kitchensink4ppt.registry import public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.help import help_payload
from kitchensink4ppt.surface.plan import build

_DC_TITLE = "{http://purl.org/dc/elements/1.1/}title"
_CORE = "docProps/core.xml"


def _call(op: str, args: dict):
    plan = build(Settings(), public_ops())
    return asyncio.run(dispatch(plan, "ppt_deck", {"op": op, "args": args}))


def _mutate(doc: Path, title: str) -> None:
    pkg = PptxPackage(doc)
    root = pkg.root(_CORE)
    t = root.find(_DC_TITLE)
    if t is None:
        t = etree.SubElement(root, _DC_TITLE)
    t.text = title
    pkg.mark_dirty(_CORE)
    pkg.save()


def _deck_with_backups(make_deck) -> tuple[Path, bytes]:
    doc = make_deck()
    original = doc.read_bytes()
    _mutate(doc, "one")  # anchor and prev now hold the original
    _mutate(doc, "two")
    return doc, original


def test_slot_is_read_as_source_and_the_answer_says_so(make_deck):
    doc, original = _deck_with_backups(make_deck)
    payload, is_error = _call("manage_backups", {
        "action": "restore", "file_path": str(doc), "slot": "anchor"})
    assert is_error is False, payload
    assert doc.read_bytes() == original
    seen = payload.get("normalized_from") or {}
    assert seen.get("args", {}).get("slot") == "anchor", payload
    assert "source" in seen.get("hint", ""), payload


def test_slot_and_source_naming_the_same_backup_are_fine(make_deck):
    doc, original = _deck_with_backups(make_deck)
    payload, is_error = _call("manage_backups", {
        "action": "restore", "file_path": str(doc),
        "slot": "anchor", "source": "anchor"})
    assert is_error is False, payload
    assert doc.read_bytes() == original


def test_slot_and_source_that_disagree_are_refused_naming_both(make_deck):
    doc, _original = _deck_with_backups(make_deck)
    before = doc.read_bytes()
    payload, is_error = _call("manage_backups", {
        "action": "restore", "file_path": str(doc),
        "slot": "anchor", "source": "prev"})
    assert is_error is True, payload
    error = payload["error"]
    assert error["code"] == "BAD_ARGS"
    assert "slot" in error["message"] and "source" in error["message"]
    assert "source" in error.get("hint", "")
    assert doc.read_bytes() == before


def test_every_listed_backup_carries_the_source_that_restores_it(make_deck):
    doc, _original = _deck_with_backups(make_deck)
    snap = bk.create_snapshot(str(doc))["snapshot"]
    listing = bk.manage_backups("list", file_path=str(doc))
    assert {item["slot"]: item["source"] for item in listing["slots"]} == {
        "prev": "prev", "anchor": "anchor"}
    assert [item["source"] for item in listing["snapshots"]] == [snap]
    folder = bk.manage_backups("list", directory=str(doc.parent))
    entry = folder["presentations"][0]
    assert {item["slot"]: item["source"] for item in entry["slots"]} == {
        "prev": "prev", "anchor": "anchor"}


def test_a_slot_name_in_capitals_still_names_the_slot(make_deck):
    doc, original = _deck_with_backups(make_deck)
    bk.manage_backups("restore", file_path=str(doc), source="Anchor")
    assert doc.read_bytes() == original


def test_a_restore_with_no_source_says_what_to_send():
    import pytest

    from kitchensink4ppt.core.errors import PptMcpError

    with pytest.raises(PptMcpError) as caught:
        bk.manage_backups("restore", file_path="deck.pptx")
    message = str(caught.value)
    assert "source" in message and "action='list'" in message


def test_help_names_source_and_the_slot_alias():
    plan = build(Settings(), public_ops())
    source = help_payload(plan, "manage_backups")["args_schema"]["properties"]["source"]
    text = source.get("description", "")
    for word in ("prev", "anchor", "snapshot", "slot"):
        assert word in text, (word, text)
