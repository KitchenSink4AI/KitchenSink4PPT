"""OS write refusals are typed and leave the target intact."""

from __future__ import annotations

import os
import stat
import sys
from pathlib import Path

import pytest

from kitchensink4ppt import server
from kitchensink4ppt.core import safesave
from kitchensink4ppt.core.package import has_read_only_attribute
from kitchensink4ppt.surface import envelope


def _call(fn, *args, **kwargs):
    try:
        return fn(*args, **kwargs)
    except BaseException as exc:
        return envelope.refusal(exc)


def _sharing_violation() -> PermissionError:
    exc = PermissionError(13, "sharing violation")
    exc.winerror = 32
    return exc


def _mark_read_only(path: Path) -> None:
    os.chmod(path, stat.S_IREAD)
    assert has_read_only_attribute(path)


def _restore_writable(path: Path) -> None:
    if path.exists():
        os.chmod(path, stat.S_IREAD | stat.S_IWRITE)


def _temporary_debris(directory: Path) -> list[Path]:
    patterns = (
        "*.ppt-mcp-tmp",
        "*.ks4p-copy-tmp",
        ".slot-*.tmp",
        ".restore-*.tmp",
        ".lock-*.tmp",
        ".ks4p-replaced-*",
    )
    found: list[Path] = []
    for pattern in patterns:
        found.extend(directory.rglob(pattern))
    found.extend(directory.rglob(safesave.LOCK_FILE_NAME))
    return found


@pytest.mark.skipif(sys.platform != "win32", reason="Windows sharing semantics")
def test_plain_read_handle_edit_is_document_locked_without_mutation(make_deck):
    deck = make_deck("held-open.pptx", extra_slides=0)
    before = deck.read_bytes()

    # Python's ordinary Windows read handle does not share delete access, so
    # the final atomic replace reproduces Explorer/AV/sync-client readers.
    with deck.open("rb"):
        out = _call(
            server.set_document_properties,
            str(deck),
            {"title": "must not publish"},
        )

    assert out["ok"] is False
    assert out["error"]["code"] == "DOCUMENT_LOCKED"
    assert str(deck) in out["error"]["message"]
    assert deck.read_bytes() == before
    assert _temporary_debris(deck.parent) == []


def test_backup_rotation_sharing_failure_is_document_locked_and_clean(
    make_deck, monkeypatch,
):
    deck = make_deck("backup-held.pptx", extra_slides=0)
    before = deck.read_bytes()

    def refuse_rotation(*_args, **_kwargs):
        raise _sharing_violation()

    monkeypatch.setattr(safesave, "prepare_rotation", refuse_rotation)
    out = _call(
        server.set_document_properties,
        str(deck),
        {"title": "must not publish"},
    )

    assert out["ok"] is False
    assert out["error"]["code"] == "DOCUMENT_LOCKED"
    assert str(deck) in out["error"]["message"]
    assert deck.read_bytes() == before
    assert _temporary_debris(deck.parent) == []


def test_copy_overwrite_read_only_destination_is_protected_before_sidecars(
    make_deck,
):
    source = make_deck("source.pptx", seed=1, extra_slides=0)
    destination = make_deck("read-only-destination.pptx", seed=2, extra_slides=0)
    before = destination.read_bytes()
    backup_root = safesave.slot_dir(destination).parent
    _mark_read_only(destination)
    try:
        out = _call(
            server.copy_presentation,
            str(source),
            str(destination),
            overwrite=True,
        )

        assert out["ok"] is False
        assert out["error"]["code"] == "PROTECTED"
        assert str(destination) in out["error"]["message"]
        assert destination.read_bytes() == before
        assert not backup_root.exists()
        assert _temporary_debris(destination.parent) == []
    finally:
        _restore_writable(destination)


def test_restore_read_only_destination_is_protected_before_sidecars(make_deck):
    destination = make_deck("restore-read-only.pptx", seed=3, extra_slides=0)
    source = make_deck("restore-source.pptx", seed=4, extra_slides=0)
    before = destination.read_bytes()
    backup_root = safesave.slot_dir(destination).parent
    _mark_read_only(destination)
    try:
        out = _call(
            server.manage_backups,
            "restore",
            file_path=str(destination),
            source=str(source),
        )

        assert out["ok"] is False
        assert out["error"]["code"] == "PROTECTED"
        assert str(destination) in out["error"]["message"]
        assert destination.read_bytes() == before
        assert not backup_root.exists()
        assert _temporary_debris(destination.parent) == []
    finally:
        _restore_writable(destination)


def test_unrelated_publish_defect_stays_internal_error_and_cleans_temp(
    make_deck, monkeypatch,
):
    deck = make_deck("internal-defect.pptx", extra_slides=0)
    before = deck.read_bytes()

    def broken_publish(_src, _dst, **_kwargs):
        raise RuntimeError("injected implementation defect")

    monkeypatch.setattr(safesave, "promote", broken_publish)
    out = _call(
        server.set_document_properties,
        str(deck),
        {"title": "must not publish"},
        backup=False,
    )

    assert out["ok"] is False
    assert out["error"]["code"] == "INTERNAL_ERROR"
    assert deck.read_bytes() == before
    assert _temporary_debris(deck.parent) == []


def test_permission_failure_for_new_path_is_protected_and_names_path(tmp_path):
    destination = tmp_path / "permission-denied.pptx"
    refusal = safesave.os_write_refusal(
        destination,
        PermissionError(13, "permission denied"),
    )

    assert refusal is not None
    out = envelope.refusal(refusal)
    assert out["error"]["code"] == "PROTECTED"
    assert str(destination) in out["error"]["message"]


def test_posix_permission_failure_for_existing_path_is_protected(tmp_path):
    destination = tmp_path / "acl-denied.pptx"
    destination.write_bytes(b"unchanged")
    refusal = safesave.os_write_refusal(
        destination,
        PermissionError(13, "permission denied"),
    )

    assert refusal is not None
    out = envelope.refusal(refusal)
    assert out["error"]["code"] == "PROTECTED"
    assert str(destination) in out["error"]["message"]
