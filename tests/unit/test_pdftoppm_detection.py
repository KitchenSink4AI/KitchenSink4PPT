"""pdftoppm detection and the remedy the engine report carries.

Field report 2026-09-21 (P4): diagnose said "needs pdftoppm (poppler): NOT
found" and stopped there, so the one line that settles the render-engine
question did not say what to do about it. The remedy sentence already
existed a few lines away in the refusal; it is one constant now, used by
both. Detection also matched _find_soffice's shape: an env override and the
well-known install locations, because a tool that is installed but not on
PATH reads to the caller as a tool that is missing.
"""

from __future__ import annotations

from pathlib import Path

from kitchensink4ppt.ops import diagnostics as diag
from kitchensink4ppt.ops import export as export_ops


def test_env_override_wins(monkeypatch, tmp_path):
    fake = tmp_path / "pdftoppm.exe"
    fake.write_bytes(b"")
    monkeypatch.setenv("KS4P_PDFTOPPM", str(fake))
    assert export_ops._find_pdftoppm() == fake


def test_env_override_pointing_nowhere_is_ignored(monkeypatch, tmp_path):
    monkeypatch.setenv("KS4P_PDFTOPPM", str(tmp_path / "nope.exe"))
    monkeypatch.setattr(export_ops.shutil, "which", lambda _n: None)
    monkeypatch.setattr(export_ops, "PDFTOPPM_WELL_KNOWN", ())
    assert export_ops._find_pdftoppm() is None


def test_well_known_paths_are_probed(monkeypatch, tmp_path):
    monkeypatch.delenv("KS4P_PDFTOPPM", raising=False)
    monkeypatch.setattr(export_ops.shutil, "which", lambda _n: None)
    fake = tmp_path / "poppler" / "bin" / "pdftoppm.exe"
    fake.parent.mkdir(parents=True)
    fake.write_bytes(b"")
    monkeypatch.setattr(
        export_ops, "PDFTOPPM_WELL_KNOWN", (Path(tmp_path / "no"), fake)
    )
    assert export_ops._find_pdftoppm() == fake


def test_engine_note_carries_the_remedy_when_poppler_is_missing(monkeypatch):
    monkeypatch.setattr(
        export_ops, "_find_soffice", lambda: Path("/opt/libreoffice/soffice")
    )
    monkeypatch.setattr(export_ops, "_find_pdftoppm", lambda: None)
    note = export_ops.get_export_engines()["engines"]["libreoffice"]["note"]
    assert "NOT found" in note
    assert export_ops.PDFTOPPM_REMEDY.strip() in note


def test_the_remedy_is_one_constant_used_by_both(monkeypatch, make_deck):
    """The refusal and the engine report must not drift apart."""
    import pytest

    from kitchensink4ppt.core.errors import PptMcpError

    monkeypatch.setattr(
        export_ops, "_find_soffice", lambda: Path("/opt/libreoffice/soffice")
    )
    monkeypatch.setattr(export_ops, "_find_pdftoppm", lambda: None)
    monkeypatch.setattr(export_ops, "_com_available", lambda: False)
    deck = make_deck("poppler.pptx")
    with pytest.raises(PptMcpError) as exc_info:
        export_ops.export_slide_images(str(deck), engine="auto")
    assert export_ops.PDFTOPPM_REMEDY.strip() in str(exc_info.value)


def test_no_remedy_when_libreoffice_is_not_installed_either(monkeypatch):
    monkeypatch.setattr(export_ops, "_find_soffice", lambda: None)
    monkeypatch.setattr(export_ops, "_find_pdftoppm", lambda: None)
    note = export_ops.get_export_engines()["engines"]["libreoffice"]["note"]
    assert export_ops.PDFTOPPM_REMEDY.strip() not in note


def test_diagnose_carries_the_remedy_and_stays_paste_safe(monkeypatch):
    monkeypatch.setattr(
        export_ops, "_find_soffice", lambda: Path("/opt/libreoffice/soffice")
    )
    monkeypatch.setattr(export_ops, "_find_pdftoppm", lambda: None)
    out = diag.diagnose()
    note = out["engines"]["engines"]["libreoffice"]["note"]
    assert export_ops.PDFTOPPM_REMEDY.strip() in note
    assert chr(0x2014) not in note  # the repo bans em dashes in runtime text
    assert "/opt/libreoffice" not in note  # paste-safe: no install locations
