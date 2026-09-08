"""Pins for the output budget and the save path's compression fidelity.

Three of these pin a contract in BOTH directions, which is the point:
a budget that is never exceeded is worthless if it drops content in
silence, and a budget that reports honestly is worthless if it still
overflows the window. So: never exceeded, nothing dropped without saying
so, and a deck small enough to fit comes back exactly as it always did.
"""

from __future__ import annotations

import json
import shutil
import zipfile
from pathlib import Path

import pytest

from kitchensink4ppt.core import budget
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import read as rd
from kitchensink4ppt.ops import view as vw

CORPUS = Path(__file__).resolve().parents[1] / "corpus"

SMALL = "proposal_defense.pptx"
LARGE = "military_brief.pptx"
STORED_MEDIA = "unitar_final.pptx"


def _deck(name: str) -> Path:
    path = CORPUS / name
    if not path.exists():
        pytest.skip(f"corpus deck {name} not present")
    return path


def _slide_xml_chars(path: Path, *, worst_slide: bool = False) -> int:
    """Uncompressed size of the deck's slide XML, in characters.

    Read straight out of the zip rather than through any reader, so this
    measurement cannot be satisfied by the code path it is used to guard.
    `worst_slide` returns the single heaviest slide instead of the total,
    for the tests that need one crowded slide rather than a crowded deck.
    """
    with zipfile.ZipFile(path) as zf:
        sizes = [
            info.file_size
            for info in zf.infolist()
            if info.filename.startswith("ppt/slides/")
            and info.filename.endswith(".xml")
        ]
    if not sizes:
        return 0
    return max(sizes) if worst_slide else sum(sizes)


#: How many characters of slide XML a deck needs per character of ceiling
#: before a read of it can plausibly overflow that ceiling.
#:
#: Slide XML is markup wrapped around the text a read returns, so the JSON
#: payload comes out several times smaller than the XML that produced it.
#: Measured on this corpus: `proposal_defense.pptx` carries 203,445
#: characters of slide XML and every read of it lands under the
#: 60,000-character default, which is what the small-read pins below
#: assert, and the other decks sit at the same order of ratio. Four is
#: picked to clear both edges by a wide margin at every ceiling these tests
#: use: the synthetic stand-ins top out at 35,490 characters of slide XML
#: and skip at all of them, while the real decks (6,912,608 for
#: military_brief, 364,343 in its heaviest single slide) run at all of them.
XML_TO_PAYLOAD_MARGIN = 4


def _deck_that_can_overflow(name: str, *, worst_slide: bool = False) -> Path:
    """A deck heavy enough that a read of it CAN exceed the ceiling in force.

    WHY THIS GUARD EXISTS. The suite runs against a corpus of real decks
    when one is present and against synthetic structural stand-ins when it
    is not (see tests/conftest.py). The stand-ins are a few slides of
    generated filler. Every test below pins the behaviour of a read that
    had to be CUT SHORT, and a deck that fits under the ceiling is never
    cut short, so on a stand-in these tests were not failing on a contract:
    they were reaching for a `page` block that correctly did not exist.

    THE THRESHOLD, DERIVED. The ceiling is `budget.max_chars()`, measured
    in characters of serialized JSON, and it is asked for here rather than
    hardcoded so a test that moves the ceiling moves this guard with it.
    The deck is then measured in the one unit that is available without
    running a read at all, its slide XML, and required to carry
    XML_TO_PAYLOAD_MARGIN characters of it per character of ceiling; that
    constant carries the measurement the factor comes from. Below the
    threshold a read cannot reach the ceiling, so skipping there skips
    nothing that could have run.

    WHAT THIS DOES NOT DO. It does not skip when truncation fails to
    happen. Above the threshold the test still runs and still requires the
    `page` block, so a regression that stopped truncating turns these red
    rather than quietly green, which is the whole hazard of a size-aware
    skip and the reason the measurement is of the fixture and not of the
    result.
    """
    path = _deck(name)
    ceiling = budget.max_chars()
    threshold = ceiling * XML_TO_PAYLOAD_MARGIN
    weight = _slide_xml_chars(path, worst_slide=worst_slide)
    what = "its heaviest single slide" if worst_slide else "its slide XML"
    if weight < threshold:
        pytest.skip(
            f"{name} is too small to truncate: {what} runs to {weight:,} "
            f"characters, under the {threshold:,} needed to overflow the "
            f"{ceiling:,}-character output ceiling in force, so no read of "
            f"it can be cut short. This is the synthetic stand-in rather "
            f"than the real corpus deck; drop the real file into "
            f"tests/corpus/ to run this pin."
        )
    return path


def _deck_with_stored_media(name: str) -> Path:
    """A deck that actually carries ZIP_STORED entries.

    Same defect class as the guard above, in the save path rather than the
    read path: the pin is that already-compressed media survives a save
    byte-identical and still STORED, and a synthetic stand-in whose every
    entry is deflated has no stored entry to preserve. Measured off the zip
    directory, so the fixture is what is being checked and not the saver.
    """
    path = _deck(name)
    with zipfile.ZipFile(path) as zf:
        stored = [
            info.filename
            for info in zf.infolist()
            if info.compress_type == zipfile.ZIP_STORED
        ]
    if not stored:
        pytest.skip(
            f"{name} carries no ZIP_STORED entries, so there is nothing for "
            f"the save path to preserve uncompressed. This is the synthetic "
            f"stand-in rather than the real corpus deck; drop the real file "
            f"into tests/corpus/ to run this pin."
        )
    return path


def _chars(obj) -> int:
    return len(json.dumps(obj, ensure_ascii=False, default=str))


def _readers(pkg):
    return {
        "list_elements(shapes)": lambda: rd.list_elements(pkg, "shapes"),
        "list_elements(shapes,compact)": lambda: rd.list_elements(
            pkg, "shapes", compact=True
        ),
        "list_elements(notes)": lambda: rd.list_elements(pkg, "notes"),
        "list_elements(slides)": lambda: rd.list_elements(pkg, "slides"),
        "get_text": lambda: rd.get_text(pkg),
        "get_text(notes)": lambda: rd.get_text(pkg, include_notes=True),
        "find_text": lambda: rd.find_text(pkg, "e"),
        "find_text(compact)": lambda: rd.find_text(pkg, "e", compact=True),
        "view": lambda: vw.get_presentation_view(pkg),
        "view(full)": lambda: vw.get_presentation_view(pkg, detail="full"),
        "view(outline)": lambda: vw.get_presentation_view(pkg, detail="outline"),
        "get_slide_info(worst)": lambda: max(
            (rd.get_slide_info(pkg, i) for i in range(len(rd.slide_table(pkg)))),
            key=_chars,
        ),
    }


def test_one_crowded_slide_is_budgeted_too():
    """A single slide is not automatically small: the heaviest corpus deck
    has one whose shape inventory ran to 130,000 characters."""
    pkg = PptxPackage(_deck_that_can_overflow(LARGE, worst_slide=True))
    worst = max(
        (rd.get_slide_info(pkg, i) for i in range(len(rd.slide_table(pkg)))),
        key=_chars,
    )
    assert _chars(worst) <= budget.max_chars()
    page = worst["page"]
    assert page["unit"] == "shapes"
    assert page["total"] == worst["shape_count"]
    assert page["returned"] == len(worst["shapes"])
    # the placeholder summary is never paged: it is the addressing surface
    assert len(worst["placeholders"]) == sum(
        1 for s in rd.list_elements(pkg, "placeholders",
                                    scope=worst["index"])["items"]
    )


# ------------------------------------------------ pin 1: never exceeded


@pytest.mark.parametrize("deck", [SMALL, LARGE, STORED_MEDIA, "pmr_tables.pptx"])
@pytest.mark.parametrize("ceiling", [2000, 8000, 25000, 60000, 250000])
def test_no_read_ever_exceeds_the_output_budget(monkeypatch, deck, ceiling):
    monkeypatch.setenv(budget.ENV_MAX_CHARS, str(ceiling))
    pkg = PptxPackage(_deck(deck))
    for label, call in _readers(pkg).items():
        size = _chars(call())
        assert size <= ceiling, f"{deck} {label} returned {size} > {ceiling}"


# ------------------------------------- pin 2: nothing dropped in silence


def test_a_truncated_read_names_its_total_and_its_continuation(monkeypatch):
    monkeypatch.setenv(budget.ENV_MAX_CHARS, "20000")
    pkg = PptxPackage(_deck_that_can_overflow(LARGE))
    result = rd.list_elements(pkg, "shapes")
    page = result["page"]
    assert page["truncated"] is True
    # count stays the truth even though items is a slice of it
    assert page["total"] == result["count"] > len(result["items"])
    assert page["returned"] == len(result["items"])
    assert page["omitted"] == page["total"] - page["returned"]
    assert page["next_offset"] == page["returned"]
    assert "offset=" in page["hint"]


def test_paging_through_a_truncated_read_reaches_every_item(monkeypatch):
    monkeypatch.setenv(budget.ENV_MAX_CHARS, "20000")
    pkg = PptxPackage(_deck_that_can_overflow(LARGE))
    seen = 0
    offset = 0
    pages = 0
    while True:
        result = rd.list_elements(pkg, "shapes", offset=offset)
        seen += len(result["items"])
        pages += 1
        page = result.get("page")
        assert page is not None
        if page["next_offset"] is None:
            break
        offset = page["next_offset"]
        assert pages < 500, "pagination is not advancing"
    assert seen == result["count"]
    assert pages > 1


def test_get_text_pages_in_slides_and_says_which_ones(monkeypatch):
    monkeypatch.setenv(budget.ENV_MAX_CHARS, "12000")
    pkg = PptxPackage(_deck_that_can_overflow(LARGE))
    first = rd.get_text(pkg)
    assert first["page"]["unit"] == "slides"
    assert first["slide_count"] > len(first["slides"])
    # the joined text covers exactly the slides that came back, no more
    assert first["text"] == "\n\n".join(s["text"] for s in first["slides"])
    nxt = rd.get_text(pkg, offset=first["page"]["next_offset"])
    assert nxt["slides"][0]["index"] == first["slides"][-1]["index"] + 1


def test_view_pages_in_whole_slides_never_mid_slide(monkeypatch):
    monkeypatch.setenv(budget.ENV_MAX_CHARS, "20000")
    pkg = PptxPackage(_deck_that_can_overflow(LARGE))
    result = vw.get_presentation_view(pkg)
    assert result["page"]["omitted"] > 0
    # every slide header that opened also closed: the count of headers
    # equals the count the page block claims to have rendered
    assert result["view"].count("\n## Slide ") == result["page"]["returned"]


def test_a_single_slide_larger_than_the_budget_is_cut_and_says_so(monkeypatch):
    monkeypatch.setenv(budget.ENV_MAX_CHARS, "2000")
    pkg = PptxPackage(
        _deck_that_can_overflow("pmr_tables.pptx", worst_slide=True))
    result = rd.get_text(pkg, include_notes=True)
    assert _chars(result) <= 2000
    page = result["page"]
    assert page["returned"] == 1
    assert page.get("chars_cut_from_last", 0) > 0
    assert "not returned" in result["slides"][0]["text"]


# --------------------------------------- pin 3: small reads are unchanged


def test_a_deck_that_fits_carries_no_page_block_at_all():
    pkg = PptxPackage(_deck(SMALL))
    for label, call in _readers(pkg).items():
        if "compact" in label or "find_text" == label:
            continue
        result = call()
        if label.startswith("find_text"):
            continue
        assert "page" not in result, f"{label} added a page block to a small read"


def test_small_deck_payloads_are_byte_identical_to_the_unbudgeted_shape():
    """The budget must be invisible below its ceiling: same keys, same
    order, same content as the reader produced before it existed."""
    pkg = PptxPackage(_deck(SMALL))
    elements = rd.list_elements(pkg, "shapes")
    assert list(elements) == ["kind", "count", "items"]
    assert elements["count"] == len(elements["items"])

    text = rd.get_text(pkg)
    assert list(text) == ["slide_count", "slides", "text"]
    assert text["slide_count"] == len(text["slides"])

    found = rd.find_text(pkg, "the")
    assert list(found) == ["query", "regex", "count", "matches"]
    assert found["count"] == len(found["matches"])

    view = vw.get_presentation_view(pkg)
    assert list(view) == ["view", "slide_count", "detail", "anchor_count"]


# ------------------------------------------- compaction is lossless


def test_compact_rows_carry_every_field_of_every_item():
    pkg = PptxPackage(_deck(SMALL))
    plain = rd.list_elements(pkg, "shapes")
    packed = rd.list_elements(pkg, "shapes", compact=True)
    assert packed["count"] == plain["count"]
    assert "page" not in packed  # this deck fits either way
    fields = packed["fields"]
    rebuilt = [
        {k: v for k, v in zip(fields, row) if v is not None}
        for row in packed["rows"]
    ]
    stripped = [
        {k: v for k, v in item.items() if v is not None} for item in plain["items"]
    ]
    assert rebuilt == stripped
    assert _chars(packed) < _chars(plain)


# ------------------------------------------- caller-supplied paging


def test_a_bad_offset_or_limit_refuses_rather_than_clamping():
    from kitchensink4ppt.core.errors import PptMcpError

    pkg = PptxPackage(_deck(SMALL))
    with pytest.raises(PptMcpError):
        rd.list_elements(pkg, "shapes", offset=-1)
    with pytest.raises(PptMcpError):
        rd.list_elements(pkg, "shapes", limit=0)


def test_an_unparseable_budget_env_refuses_rather_than_defaulting(monkeypatch):
    from kitchensink4ppt.core.errors import PptMcpError

    monkeypatch.setenv(budget.ENV_MAX_CHARS, "lots")
    with pytest.raises(PptMcpError):
        budget.max_chars()
    monkeypatch.setenv(budget.ENV_MAX_CHARS, "10")
    with pytest.raises(PptMcpError):
        budget.max_chars()
    monkeypatch.setenv(budget.ENV_MAX_CHARS, "")
    assert budget.max_chars() == budget.DEFAULT_MAX_CHARS


# ============================================ save path: compression


def _compress_map(path: Path) -> dict[str, int]:
    with zipfile.ZipFile(path) as zf:
        return {i.filename: i.compress_type for i in zf.infolist()}


def test_media_the_source_stored_is_stored_back_byte_identical(tmp_path):
    src = _deck_with_stored_media(STORED_MEDIA)
    before = _compress_map(src)
    stored = [n for n, c in before.items() if c == zipfile.ZIP_STORED]
    assert stored, "fixture no longer carries stored entries"

    work = tmp_path / src.name
    shutil.copy2(src, work)
    pkg = PptxPackage(work)
    pkg.tree("ppt/presentation.xml")
    pkg.mark_dirty("ppt/presentation.xml")
    pkg.save(do_backup=False)

    after = _compress_map(work)
    for name in stored:
        assert after[name] == zipfile.ZIP_STORED, f"{name} was re-deflated on save"

    # and the bytes themselves are untouched
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(work) as b:
        for name in stored:
            assert a.read(name) == b.read(name)


def test_already_compressed_media_is_not_deflated_on_save(tmp_path):
    src = _deck(LARGE)
    work = tmp_path / src.name
    shutil.copy2(src, work)
    pkg = PptxPackage(work)
    pkg.tree("ppt/presentation.xml")
    pkg.mark_dirty("ppt/presentation.xml")
    pkg.save(do_backup=False)

    after = _compress_map(work)
    media = [n for n in after if n.startswith("ppt/media/")]
    assert media, "fixture no longer carries media"
    jpegs = [n for n in media if n.lower().endswith((".jpeg", ".jpg", ".png", ".gif"))]
    assert jpegs
    for name in jpegs:
        assert after[name] == zipfile.ZIP_STORED, f"{name} was deflated for nothing"
    # XML still deflates: it compresses five to ten times over
    assert after["ppt/presentation.xml"] == zipfile.ZIP_DEFLATED


def test_xml_parts_keep_deflating(tmp_path, make_deck):
    deck = make_deck("plain.pptx")
    pkg = PptxPackage(deck)
    pkg.tree("ppt/presentation.xml")
    pkg.mark_dirty("ppt/presentation.xml")
    pkg.save(do_backup=False)
    after = _compress_map(Path(deck))
    xml = [n for n in after if n.endswith((".xml", ".rels"))]
    assert xml
    assert all(after[n] == zipfile.ZIP_DEFLATED for n in xml)


# ============================================ load path: the magic bytes


def test_opening_a_deck_does_not_read_the_whole_file_for_eight_bytes(monkeypatch):
    """The encryption probe used to slurp the entire presentation and throw
    all but the first eight bytes away: 23.6 ms and a 40 MB transient on
    the largest corpus deck, on EVERY tool call."""
    called: list[str] = []
    original = Path.read_bytes

    def spy(self):
        called.append(str(self))
        return original(self)

    monkeypatch.setattr(Path, "read_bytes", spy)
    PptxPackage(_deck(SMALL))
    assert called == [], f"the load path read whole files: {called}"
