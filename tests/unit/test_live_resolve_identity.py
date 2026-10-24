"""A deck PowerPoint opened through a junction or a mapped path is the same deck.

Live edits used to match an open presentation to the requested file by comparing
PowerPoint's FullName with the resolved path as text. A deck opened through a
directory junction (or a SUBST drive) has a FullName that names the junction,
while the resolved path names the real folder, so live edits said the deck
was not open, and the bridge's rule that it never opens the user's own copy
a second time could not see it either. Now the file's identity (volume
serial and file index) decides when the text differs. The files are made up;
no PowerPoint is started here.
"""

from __future__ import annotations

import sys

import pytest

from kitchensink4ppt.com import live
from kitchensink4ppt.core.errors import DocumentNotOpenInPowerPoint

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="junctions are Windows-only")


class _Pres:
    def __init__(self, full_name: str):
        self.FullName = full_name


class _PV:
    Count = 0


class _App:
    def __init__(self, names):
        self.Presentations = [_Pres(n) for n in names]
        self.ProtectedViewWindows = _PV()


class _PW:
    class com_error(Exception):
        pass


def _junction(tmp_path):
    import _winapi

    real = tmp_path / "real"
    real.mkdir()
    deck = real / "plan.pptx"
    deck.write_bytes(b"made-up bytes; never opened here")
    via = tmp_path / "via"
    _winapi.CreateJunction(str(real), str(via))
    return deck, via / "plan.pptx"


def test_a_deck_opened_through_a_junction_is_found(tmp_path, monkeypatch):
    deck, through = _junction(tmp_path)
    monkeypatch.setattr(live, "_find_pres_via_rot", lambda *a, **k: (None, None))
    app = _App([str(tmp_path / "other.pptx"), str(through)])
    got_app, pres = live._resolve_presentation(None, _PW, None, app, str(deck))
    assert got_app is app and pres.FullName == str(through)


def test_another_file_with_the_same_name_is_not_taken(tmp_path, monkeypatch):
    deck, _through = _junction(tmp_path)
    elsewhere = tmp_path / "elsewhere"
    elsewhere.mkdir()
    (elsewhere / "plan.pptx").write_bytes(b"another made-up file")
    monkeypatch.setattr(live, "_find_pres_via_rot", lambda *a, **k: (None, None))
    app = _App([str(elsewhere / "plan.pptx"), "Presentation1", "https://example.com/plan.pptx"])
    with pytest.raises(DocumentNotOpenInPowerPoint):
        live._resolve_presentation(None, _PW, None, app, str(deck))


def test_same_deck_helper(tmp_path):
    deck, through = _junction(tmp_path)
    assert live.same_deck(str(through), str(deck))
    assert live.same_deck(str(deck).upper(), str(deck))
    assert not live.same_deck("Presentation1", str(deck))
    assert not live.same_deck(str(tmp_path / "missing.pptx"), str(deck))
