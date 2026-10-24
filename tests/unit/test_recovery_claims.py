"""The changelog and README say only what the save and its recovery do.

Each phrase below was published about saves and their recovery and was not
true in every case: a comparison with other software that nothing here
measured, a time limit the start did not keep, and promises of a notice or a
put-back that some cases did not give. A hit means one came back.
"""

from __future__ import annotations

from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
SURFACES = ("CHANGELOG.md", "README.md")

#: (phrase, why it is not said)
RETIRED = (
    ("as with Office's own saves", "a comparison nothing here measured"),
    ("missing for a few milliseconds during a save",
     "usually a few milliseconds; a slow disk or scan can make it longer"),
    ("holds the server up for at most 1.5 seconds",
     "the start no longer waits for a share; one share used to cost 20 seconds"),
    ("staged backups that a stopped save left are cleared after an hour",
     "staged backups a record of a save still needs are kept"),
    ("lists it under `notices`, once.", "a kept record is named again in each run"),
    ("the next start of the server or the next change to the deck puts it back",
     "a read-only server puts nothing back"),
    ("Now the old copy stays where it is",
     "only when the record of the save was kept"),
    ("Recovery finds a deck under any spelling of its path",
     "a save through a subst or mapped drive letter is matched only at the next "
     "start, and a UNC spelling of a local path was never proven"),
)


@pytest.mark.parametrize("surface", SURFACES)
@pytest.mark.parametrize("phrase, why", RETIRED, ids=[p for p, _w in RETIRED])
def test_a_retired_claim_about_saves_does_not_come_back(surface, phrase, why):
    text = (ROOT / surface).read_text(encoding="utf-8")
    assert phrase not in text, f"{surface} says {phrase!r} ({why})"


def test_the_public_text_about_saves_has_no_em_dash():
    for surface in SURFACES:
        assert "—" not in (ROOT / surface).read_text(encoding="utf-8"), surface
