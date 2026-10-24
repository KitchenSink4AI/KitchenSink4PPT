"""Encrypted inputs are refused before any Office engine is called."""

from __future__ import annotations

import pytest

from kitchensink4ppt import server
from kitchensink4ppt.com import bridge
from kitchensink4ppt.core.errors import DocumentProtected
from kitchensink4ppt.ops import export as export_ops


OLE_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")


@pytest.fixture
def encrypted_deck(tmp_path):
    deck = tmp_path / "password.pptx"
    deck.write_bytes(OLE_MAGIC + (b"\x00" * 4088))
    return deck


def _must_not_call(*_args, **_kwargs):
    raise AssertionError("encrypted input reached an Office export/COM route")


def test_validate_refuses_encrypted_input_before_com(encrypted_deck, monkeypatch):
    monkeypatch.setattr(bridge, "com_validate_opens_clean", _must_not_call)

    with pytest.raises(DocumentProtected, match="password-protected"):
        server.validate(str(encrypted_deck))


@pytest.mark.parametrize("engine", ["com", "auto"])
def test_export_pdf_refuses_encrypted_input_before_com(
    encrypted_deck, tmp_path, monkeypatch, engine
):
    monkeypatch.setattr(export_ops, "_com_available", _must_not_call)
    monkeypatch.setattr(export_ops, "_pick_engine_name", _must_not_call)
    monkeypatch.setattr(bridge, "com_export_pdf", _must_not_call)

    with pytest.raises(DocumentProtected, match="password-protected"):
        export_ops.export_pdf(
            str(encrypted_deck),
            output=str(tmp_path / "out.pdf"),
            engine=engine,
        )


@pytest.mark.parametrize("engine", ["com", "auto"])
def test_export_slide_images_refuses_encrypted_input_before_com(
    encrypted_deck, tmp_path, monkeypatch, engine
):
    monkeypatch.setattr(export_ops, "_com_available", _must_not_call)
    monkeypatch.setattr(export_ops, "_pick_image_engine_name", _must_not_call)
    monkeypatch.setattr(bridge, "com_export_slide_images", _must_not_call)

    with pytest.raises(DocumentProtected, match="password-protected"):
        export_ops.export_slide_images(
            str(encrypted_deck),
            output_dir=str(tmp_path / "slides"),
            engine=engine,
        )


def test_export_handout_refuses_encrypted_input_before_com(
    encrypted_deck, tmp_path, monkeypatch
):
    monkeypatch.setattr(export_ops, "_com_available", _must_not_call)
    monkeypatch.setattr(bridge, "com_export_handout", _must_not_call)

    with pytest.raises(DocumentProtected, match="password-protected"):
        export_ops.export_handout(
            str(encrypted_deck),
            output=str(tmp_path / "handout.pdf"),
        )
