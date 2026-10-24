"""Ordinary file read-only attributes are not process locks."""

from __future__ import annotations

import os
import stat

import pytest

from kitchensink4ppt import server
from kitchensink4ppt.core import safesave
from kitchensink4ppt.core.errors import DocumentReadOnly
from kitchensink4ppt.core.package import PptxPackage, has_read_only_attribute
from kitchensink4ppt.surface import envelope


def _mark_read_only(path):
    os.chmod(path, stat.S_IREAD)
    assert has_read_only_attribute(path)


def _restore_writable(path):
    if path.exists():
        os.chmod(path, stat.S_IREAD | stat.S_IWRITE)


def test_file_mode_readers_accept_a_read_only_presentation(
    make_deck, monkeypatch,
):
    deck = make_deck(extra_slides=1)
    before = deck.read_bytes()
    _mark_read_only(deck)
    monkeypatch.setattr(
        server._lo,
        "live_get_text",
        lambda *args, **kwargs: pytest.fail(
            "a read-only attribute was misrouted to the live reader"
        ),
    )
    try:
        text = server.get_text(str(deck))
        view = server.get_presentation_view(str(deck))
        audit = server.audit_accessibility(str(deck))

        assert text["slide_count"] >= 1
        assert view["slide_count"] == text["slide_count"]
        assert audit["slides_checked"] == text["slide_count"]
        assert deck.read_bytes() == before
        assert has_read_only_attribute(deck)
    finally:
        _restore_writable(deck)


def test_dual_mode_mutation_does_not_misroute_read_only_as_live(
    make_deck, monkeypatch,
):
    deck = make_deck(extra_slides=0)
    _mark_read_only(deck)
    monkeypatch.setattr(
        server._lo,
        "live_set_notes",
        lambda *args, **kwargs: pytest.fail(
            "a read-only attribute was misdiagnosed as a PowerPoint lock"
        ),
    )
    try:
        with pytest.raises(DocumentReadOnly, match="read-only"):
            server.set_notes(str(deck), 0, text="cannot save this")
    finally:
        _restore_writable(deck)


def test_mutation_refuses_with_read_only_reason_before_sidecars(make_deck):
    deck = make_deck(extra_slides=0)
    before = deck.read_bytes()
    backup_root = safesave.slot_dir(deck).parent
    _mark_read_only(deck)
    try:
        with pytest.raises(DocumentReadOnly) as caught:
            server.set_document_properties(
                str(deck), {"title": "must not be written"}
            )

        refusal = envelope.refusal(caught.value)
        assert refusal["error"]["code"] == "PROTECTED"
        message = refusal["error"]["message"].lower()
        assert "read-only" in message
        assert "open in powerpoint" not in message
        assert "locked by another process" not in message
        assert "copy_presentation" in refusal["error"]["hint"]
        assert deck.read_bytes() == before
        assert not backup_root.exists()
    finally:
        _restore_writable(deck)


def test_package_save_backstop_refuses_a_read_only_target(make_deck):
    deck = make_deck(extra_slides=0)
    pkg = PptxPackage(deck)
    _mark_read_only(deck)
    try:
        with pytest.raises(DocumentReadOnly, match="read-only"):
            pkg.save(do_backup=False)
    finally:
        _restore_writable(deck)


def test_copy_presentation_clears_the_sources_read_only_attribute(make_deck):
    source = make_deck("readonly-source.pptx", extra_slides=0)
    destination = source.with_name("writable-copy.pptx")
    before = source.read_bytes()
    _mark_read_only(source)
    try:
        result = server.copy_presentation(str(source), str(destination))

        assert result["ok"] is True
        assert destination.read_bytes() == before
        assert has_read_only_attribute(source)
        assert not has_read_only_attribute(destination)

        # Prove the advertised workaround reaches a normal mutation, rather
        # than merely changing the copied mode bit in isolation.
        changed = server.set_document_properties(
            str(destination), {"title": "Writable working copy"},
            backup=False,
        )
        assert changed["ok"] is True
    finally:
        _restore_writable(source)
        _restore_writable(destination)
