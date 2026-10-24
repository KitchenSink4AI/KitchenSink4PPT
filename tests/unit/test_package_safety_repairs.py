"""Regression tests for package safety: protected decks, anonymize
snapshots and public snapshots."""

from __future__ import annotations

import asyncio
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt import server
from kitchensink4ppt.core import errors, safesave
from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import backups, comments
from kitchensink4ppt.ops.assembly import (
    get_document_properties,
    set_document_properties,
)
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface import envelope
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.plan import build


def _add_modify_verifier(path: Path) -> None:
    pkg = PptxPackage(path)
    verifier = etree.Element(qn("p:modifyVerifier"))
    verifier.set("cryptProviderType", "rsaAES")
    verifier.set("cryptAlgorithmClass", "hash")
    verifier.set("cryptAlgorithmType", "typeAny")
    verifier.set("cryptAlgorithmSid", "14")
    verifier.set("spinCount", "100000")
    verifier.set("hashData", "AA==")
    verifier.set("saltData", "AA==")
    pkg._insert_presentation_child(verifier)
    pkg.mark_dirty()
    pkg.save(do_backup=False)


def test_file_path_mutation_refuses_modify_verifier_with_protected_code(
    make_deck, monkeypatch,
):
    deck = make_deck("protected.pptx")
    _add_modify_verifier(deck)
    before = deck.read_bytes()

    def must_not_run(*args, **kwargs):
        raise AssertionError("protected mutation reached its operation body")

    monkeypatch.setattr(server._asm, "set_document_properties", must_not_run)

    with pytest.raises(errors.DocumentEditProtected) as caught:
        server.set_document_properties(
            str(deck), {"title": "must not land"}, backup=False
        )

    refusal = envelope.refusal(caught.value)
    assert refusal["error"]["code"] == "PROTECTED"

    payload, is_error = asyncio.run(
        dispatch(
            build(Settings()),
            "ppt_deck",
            {
                "op": "set_document_properties",
                "args": {
                    "file_path": str(deck),
                    "properties": {"title": "must not land"},
                    "backup": False,
                },
            },
        )
    )
    assert is_error is True
    assert payload["error"]["code"] == "PROTECTED"
    assert deck.read_bytes() == before


def test_anonymize_deck_manages_private_snapshot_and_can_restore_identities(
    make_deck,
):
    deck = make_deck("identified.pptx")
    pkg = PptxPackage(deck)
    set_document_properties(
        pkg,
        author="Jane Reviewer",
        last_modified_by="Jane Reviewer",
        company="Example Org",
    )
    comments.add_comment(pkg, 0, "Identity-bearing note", author="Jane Reviewer")
    pkg.save(do_backup=False)
    before = deck.read_bytes()

    result = server.anonymize_deck(
        str(deck), replacement="Reviewer", backup=False
    )

    snapshot = Path(result["snapshot"])
    assert snapshot.is_file()
    assert snapshot.parent == safesave.slot_dir(deck)
    assert snapshot.parent.parent.name == safesave.BACKUP_DIR_NAME
    assert snapshot.read_bytes() == before
    assert "before anonymize" in snapshot.name
    assert not list(deck.parent.glob("*before anonymize*.pptx"))
    assert deck.read_bytes() != before
    assert result["backup"] is False

    anonymized = PptxPackage(deck)
    assert get_document_properties(anonymized)["core"]["author"] == "Reviewer"
    anon_comment = comments.list_comments(anonymized)["slides"][0]["comments"][0]
    assert anon_comment["author"] == "Reviewer-1"

    listing = server.manage_backups("list", file_path=str(deck))
    assert listing["slots"] == []
    assert [Path(item["source"]) for item in listing["snapshots"]] == [snapshot]
    directory_listing = server.manage_backups("list", directory=str(deck.parent))
    item = next(
        item
        for item in directory_listing["presentations"]
        if Path(item["presentation"]) == deck
    )
    assert [Path(entry["source"]) for entry in item["snapshots"]] == [snapshot]

    restored = server.manage_backups(
        "restore", file_path=str(deck), source=listing["snapshots"][0]["source"]
    )
    assert restored["from"] == f"file {snapshot.name}"
    assert deck.read_bytes() == before
    original = PptxPackage(deck)
    assert get_document_properties(original)["core"]["author"] == "Jane Reviewer"
    original_comment = comments.list_comments(original)["slides"][0]["comments"][0]
    assert original_comment["author"] == "Jane Reviewer"
    assert original_comment["author_initials"] == "JR"

    server.manage_backups(
        "purge", file_path=str(deck), scope="slots", dry_run=False
    )
    assert snapshot.is_file()
    assert server.manage_backups("list", file_path=str(deck))["snapshots"]


def test_public_snapshot_is_hidden_and_in_the_managed_listing(make_deck):
    """A snapshot goes into the deck's hidden backup folder by default, where
    manage_backups lists it; only a dest_dir copy is visible, and that one is
    the caller's to keep, outside the listing."""
    deck = make_deck("public.pptx")
    result = backups.create_snapshot(str(deck), label="keeper")
    snapshot = Path(result["snapshot"])

    assert snapshot.parent == safesave.slot_dir(deck)
    assert result["managed"] is True and result["hidden"] is True
    listing = backups.list_backups(file_path=str(deck))["snapshots"]
    assert [Path(item["source"]) for item in listing] == [snapshot]

    keep = deck.parent / "keep"
    keep.mkdir()
    visible = backups.create_snapshot(str(deck), label="keeper", dest_dir=str(keep))
    assert Path(visible["snapshot"]).parent == keep
    assert "managed" not in visible and visible["hidden"] is False
    listing = backups.list_backups(file_path=str(deck))["snapshots"]
    assert [Path(item["source"]) for item in listing] == [snapshot]
