"""The test suite runs against a corpus of presentations in tests/corpus/
(real-world decks; drop your own files here to run the full suite).

Without the corpus, corpus-dependent tests are skipped with an explanation.
To run the full suite, drop your own .pptx/.potx files into tests/corpus/
using the expected names below; decks with real layouts, tables, pictures,
and speaker notes give the most meaningful coverage.
"""

import os
from pathlib import Path

import pytest

CORPUS = Path(__file__).parent / "corpus"

# name -> features the tests expect it to contain
EXPECTED = {
    "proposal_deck.pptx": "26-slide templated deck (the key deck)",
    "conference_talk.pptx": "conference talk deck",
    "image_deck.pptx": "conference deck with images",
    "conference_template.potx": "PowerPoint template package (.potx)",
    "heavy_deck.pptx": "heavy deck with many layouts and shapes",
    "table_deck.pptx": "deck with tables",
}


def _present() -> set[str]:
    return {p.name for p in CORPUS.glob("*.pptx")} | {
        p.name for p in CORPUS.glob("*.potx")
    }


def pytest_configure(config):
    """Missing corpus files are GENERATED as structural stand-ins (see
    tests/make_corpus.py), so the full suite runs anywhere, CI included.
    Local decks, when present, always take precedence."""
    # NO NETWORK IN TESTS. diagnose runs the on-demand update check, so the
    # suite turns it off for every test by default; the update tests switch
    # it back on for themselves and mock the fetch. setdefault, so a
    # developer who exports the variable still wins.
    os.environ.setdefault("KS4P_UPDATE_CHECK", "off")
    if set(EXPECTED) - _present():
        import make_corpus  # noqa: F401  (lives beside this file)

        made = make_corpus.generate_missing(verbose=True)
        if made:
            print(f"conftest: generated {len(made)} synthetic corpus file(s)")


def pytest_collection_modifyitems(config, items):
    missing = set(EXPECTED) - _present()
    if not missing:
        return
    skip = pytest.mark.skip(
        reason=(
            f"test corpus incomplete (missing: {sorted(missing)}) and "
            "generation failed; run python tests/make_corpus.py for details."
        )
    )
    uses_corpus: dict[str, bool] = {}
    for item in items:
        fname = str(item.fspath)
        if fname not in uses_corpus:
            uses_corpus[fname] = "corpus" in Path(fname).read_text(
                encoding="utf-8", errors="ignore"
            )
        # Only corpus-dependent tests are skipped; corpus-free tests still run.
        if uses_corpus[fname]:
            item.add_marker(skip)


@pytest.fixture()
def make_deck(tmp_path):
    """Build a small synthetic deck in tmp_path and return its Path."""

    def _make(name: str = "deck.pptx", *, seed: int = 0, extra_slides: int = 2):
        import make_corpus

        return make_corpus.build_deck(
            tmp_path / name, seed=seed, extra_slides=extra_slides
        )

    return _make
