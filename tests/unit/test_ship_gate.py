"""THE PRE-SHIP GREP GATE: a check that reads the artifact.

A publishing defect is usually catchable by one grep. It ships when what gets
checked is the report about the file rather than the file. This is the check
that opens the manifest and the other published surfaces and reads them.

Five patterns, all of them things that mean "somebody was still working
here" and none of them things a stranger should ever read:

- ``COPY PENDING`` and ``FACTS TO CONVEY``, the build's own instructions to
  a future writer. Six of these once shipped on another product's install
  screen.
- ``PLACEHOLDER``, a marker this product once printed at users out of its
  update-check strings.
- The em dash, banned across every public surface in every language.
- A ``user_config`` entry typed anything but ``boolean``. The install screen
  is checkboxes; a free-text box is a typo waiting to break somebody's
  install, and the manifest schema has no enum type that would make one
  safe. ``EXEMPT_TYPED_FIELDS`` names any field exempted from that rule, and
  naming it here is what keeps the exception from spreading.
- An absolute ``C:\\Users\\`` path, which is a build that only runs on the
  machine that built it.

SCOPE. Everything in ``ENFORCED`` is clean of all five, so every hit is a
regression. ``docs/index.html`` is on that list because landing pages once
shipped figures their own READMEs contradicted; the published-figure guard at
the bottom of this file covers the same class from the other direction. The
``src/`` trees are not scanned here: no em dash belongs in any served text,
and ``test_copy_guards.py`` holds that tree to the same rule at the source
level. This is the KitchenSink4Web gate (``tests/unit/test_ship_gate.py``
there), ported so the whole suite is checked by one rule rather than by one
product's discipline.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]

#: Text surfaces scanned. Every one is clean. A surface a stranger reads
#: belongs here; add it rather than trusting that nobody edits it.
ENFORCED = (
    "bundle/manifest.json",
    "README.md",
    "docs/llms.txt",
    "docs/QUICKSTART.md",
    "docs/index.html",
)

#: (needle, what it means) for the plain-text patterns. Written as literals
#: rather than a regex so the failure message can quote the needle back.
FORBIDDEN_TEXT = (
    ("COPY PENDING", "a copy marker aimed at a future writer"),
    ("FACTS TO CONVEY", "build instructions, not product copy"),
    ("PLACEHOLDER", "an unfilled string"),
    ("\u2014", "an em dash, banned on every public surface"),
)

#: The absolute-path pattern, in both the spellings a JSON file can hold it.
HOME_PATHS = ("C:\\Users\\", "C:\\\\Users\\\\")

#: The typed install-screen fields that are exempt, with the type each is
#: allowed to be. Empty for every product but KS4XL, whose folder picker is
#: allowed by name because a directory chooser is a picker rather than a
#: free-text box.
EXEMPT_TYPED_FIELDS: dict[str, str] = {}


def _read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8")


def _sources() -> dict[str, str]:
    return {rel: _read(rel) for rel in ENFORCED}


@pytest.mark.parametrize("needle,why", FORBIDDEN_TEXT)
def test_no_working_state_reaches_a_shipping_surface(needle, why):
    for where, text in _sources().items():
        assert needle not in text, (
            f"{where} carries {needle!r}: {why}. This pattern means "
            f"working state reached a shipping surface; fix "
            f"the surface, do not loosen the gate.")


def test_no_absolute_home_path_reaches_a_shipping_surface():
    """A hardcoded user directory means the artifact runs on exactly one
    machine. Another product's first build launched
    ``C:\\Users\\...\\venv\\Scripts\\...exe`` and could not have installed
    for anybody else."""
    for where, text in _sources().items():
        for pattern in HOME_PATHS:
            assert pattern not in text, (
                f"{where} hardcodes {pattern!r}. The bundle has to launch on "
                f"a machine nobody here owns.")


def test_every_install_screen_field_is_a_checkbox():
    """No typed field on the install screen except the ones allowed by
    name. The rule is stated as a scan rather than as a parity assertion,
    so it holds for the manifest a human installs from."""
    for where in ENFORCED:
        if not where.endswith(".json"):
            continue
        config = json.loads(_read(where)).get("user_config", {})
        typed = sorted(
            key for key, entry in config.items()
            if entry.get("type") != "boolean"
            and entry.get("type") != EXEMPT_TYPED_FIELDS.get(key))
        assert not typed, (
            f"{where} ships typed install-screen fields {typed}. The install "
            f"screen is checkboxes; a value a user has to spell correctly "
            f"belongs in a launch file, not on this screen. An exception is "
            f"added to EXEMPT_TYPED_FIELDS on purpose, never because the "
            f"gate was inconvenient.")


def test_every_enforced_surface_is_still_a_real_file():
    """A row pointing at a renamed file would quietly turn into no gate at
    all, which is how a total scope stops being total."""
    for rel in ENFORCED:
        assert (ROOT / rel).is_file(), rel


#: Surfaces cut from a measured source. docs/index.html is a redirect stub to
#: the community site and carries no figures. Every figure below is published
#: on both of these, so a change that restamps one and forgets another turns this
#: red. Stale landing pages once lived in exactly that gap: the README said
#: one number and docs/index.html said another, and nothing checked either
#: against the other.
FIGURE_SURFACES = (
    "README.md",
    "docs/llms.txt",
)

#: (figure, what it measures, the script that produces it). Values are the
#: current measurement. Change one only together with a re-run of the named
#: script.
PUBLISHED_FIGURES = (
    ('9 tools', 'tools', 'scripts/measure_surface.py'),
    ('145 operations', 'operations', 'scripts/measure_surface.py'),
    ('193', 'actions, counting each action an operation can take', 'scripts/count_operations.py'),
    ('3,689', 'tests', 'scripts/stamp_figures.py'),
    ('7.3k tokens', 'tokens for the whole surface', 'scripts/measure_surface.py'),
)

#: (string, what it used to mean). Absent from every FIGURE_SURFACES file.
#: A hit means a surface went backwards.
SUPERSEDED_FIGURES = (
    ('885 tests', 'the pre-1.2.0 test count'),
    ('1,238 tests', 'the test count before the gate gained its figure guard'),
    ('1,241 tests', 'the pre-1.2.2 test count'),
    ('885 Tests', 'the same count in german'),
    ('885 pruebas', 'the same count in spanish'),
    ('138 tools', 'the pre-1.2.0 tool count'),
    ('138-tool', 'the same count in the meta description'),
    ('5.5k', 'the lite bill from the description-plus-schema estimator'),
    ('29.9k', 'the full-surface bill from the same old estimator'),
    ('5,500 tokens', 'the same lite bill written out'),
    ('29,900', 'the same full bill written out'),
    ('№ 138', 'the same count as a catalog number'),
    ('141 tools', 'the pre-1.2.2 tool count'),
    ('141-tool', 'the same count in the meta description'),
    ('№ 141', 'the same count as a catalog number'),
    ('4.7k', 'the pre-1.2.0 lite core cost'),
    ('5.3k', 'the pre-1.2.2 lite core cost'),
    ('~7.1k', 'the pre-1.3.1 lite core cost, as the pack list wrote it'),
    ('roughly 7.1k tokens', 'the same cost in the README prose'),
    ('| 7.1k |', 'the same cost in the pack table'),
    ('Tokens, lite core up front', 'the landing-page label for that cost'),
    ('1,251 tests', 'the pre-1.3.1 test count'),
    ('1,640 tests', 'the 1.3.1 test count'),
    ('~7.2k', 'the 1.3.1 lite core cost, as the pack list wrote it'),
    ('roughly 7.2k tokens', 'the same cost in the README prose'),
    ('| 7.2k |', 'the same cost in the pack table'),
    ('about 7,200 tokens', 'the same cost written out in llms.txt'),
    ('39.7k', 'the 1.3.1 full surface cost'),
    ('142 tools with', 'the 1.3.1 tool total stated as the current surface'),
    ('6.8k', 'the 2.0.0rc1 whole-surface cost, before the slide rule joined the args description'),
    ('2,371', 'the 2.0.0rc1 test count'),
    ('2,386', 'the 2.0.0rc2 test count'),
    ('2,412', 'the 2.0.0rc3 test count'),
    ('2,467', 'the 2.0.0rc4 test count'),
    ('2,510', 'the 2.0.0rc5 test count'),
    ('2,536', 'the 2.0.0rc6 test count'),
    ('2,794', 'the 2.0.0rc7 test count'),
    ('3,075', 'the 2.0.0rc8 test count'),
    ('3,385', 'the 2.0.0rc9 test count'),
    ('3,414', 'the 2.0.0rc10 test count'),
    ('3,458', 'the 2.0.0rc12 test count'),
    ('3,474', 'the 2.0.0rc13 test count'),
    ('3,480', 'the 2.0.0rc14 test count'),
    ('3,488', 'the 2.0.0rc15 test count'),
    ('3,518', 'the 2.0.0rc16 test count'),
    ('3,546', 'the 2.0.0 test count before the update notice tests'),
    ('3,609', 'the 2.0.0 test count before the consent-after-refresh test'),
    ('3,610', 'the test count before the dialog-step version test'),
    ('3,611', 'the test count before the old-settings tests'),
    ('7.2k tokens', 'the whole-surface cost through 2.0.0rc16, cut instead of rounded (7,262 tokens is 7.3k)'),
    ('3,438', 'the 2.0.0rc11 test count'),
    ('140 operations', 'the 2.0.0rc6 operation count'),
    ('№ 140', 'the same count as a catalog number'),
    ('7.0k tokens', 'the 2.0.0rc6 whole-surface cost'),
    ('7,013', 'the 2.0.0rc6 whole-surface estimate written out'),
    ('7.1k tokens', 'the 2.0.0rc7 whole-surface cost'),
    ('7,148', 'the 2.0.0rc7 whole-surface estimate written out'),
    ('7,239', 'the 2.0.0rc8 whole-surface estimate written out'),
    ('143 operations', 'the 2.0.0 operation count before update_now and update_notices'),
    ('№ 143', 'the same count as a catalog number'),
    ('7,262', 'the whole-surface estimate before update_now and update_notices, written out'),
)


def test_every_published_figure_reads_the_same_on_every_surface():
    """A figure published on three surfaces has to be the same figure on all
    three. Checking presence rather than equality keeps the guard honest
    about locale number formatting, which differs by design.

    Looped rather than parametrized on purpose: this repo's collected test
    count is one of the figures below, so a parametrized table would move
    the number it is guarding every time somebody added a row to it."""
    missing = []
    for figure, what, source in PUBLISHED_FIGURES:
        for rel in FIGURE_SURFACES:
            if figure not in _read(rel):
                missing.append(f"{rel} does not carry {figure!r} ({what}, "
                               f"measured by {source})")
    assert not missing, (
        "\n".join(missing) + "\nEither a surface was missed by a restamp or "
        "the figure moved and this table was not updated. Re-run the script "
        "that produces it; do not guess the number.")


#: Lines that carry a PER-PACK bill rather than a headline figure. The
#: superseded list retires lite and full bills by bare number, and a pack bill
#: is free to land on a number a headline once used: tables-charts measures
#: ~5.3k on the current estimator, which is also the lite bill this repo
#: retired before 1.2.2. Pack bills move with every re-measure and are read
#: off packs.pack_cost, so excluding them here loses no coverage.
_PACK_BILL_LINE = re.compile(
    r"^\|\s*(lite core|graphics|tables-charts|design|assembly-export|"
    r"review-sweeps|com)\b|^Six packs plus the lite core"
)


def _without_pack_bills(text: str) -> str:
    return "\n".join(
        line for line in text.splitlines() if not _PACK_BILL_LINE.match(line)
    )


def test_no_superseded_figure_survives_on_a_published_surface():
    """The regression guard for the defect this file was widened over."""
    found = []
    for stale, what in SUPERSEDED_FIGURES:
        for rel in FIGURE_SURFACES:
            if stale in _without_pack_bills(_read(rel)):
                found.append(f"{rel} still carries {stale!r}, {what}")
    assert not found, (
        "\n".join(found) + "\nRestamp from the measuring script rather than "
        "deleting the row.")


def test_every_figure_surface_is_still_a_real_file():
    for rel in FIGURE_SURFACES:
        assert (ROOT / rel).is_file(), rel


def test_the_frozen_released_surface_reports_the_published_legacy_total():
    """The gate above holds three text surfaces to each other. This holds
    the RUNNING SERVER to them, which is the surface an agent actually
    reads: get_server_info counts the live registry, so a restamp that
    forgets a page turns this red rather than shipping a server that
    contradicts its own README."""
    released = json.loads(_read("migration/released_tools_1.3.1.json"))
    legacy = len(released["tools"])
    # 2.0 publishes the nine-tool surface; the frozen 1.3.1 total survives
    # only as the "from" side of the upgrade notes on the text surfaces.
    for rel in ("README.md", "docs/llms.txt"):
        assert f"{legacy} flat tools" in _read(rel), (
            f"the frozen release has {legacy} tools; {rel} must state that "
            f"count in its upgrade notes. Re-run scripts/measure_surface.py "
            f"and restamp {', '.join(FIGURE_SURFACES)} plus PUBLISHED_FIGURES.")


def test_the_server_reports_the_same_test_count_it_publishes():
    """shipped.TESTS is what get_server_info hands an agent. It cannot be
    derived at runtime, so this is the join that keeps it from drifting
    away from the number on the page. Re-stamp with
    scripts/stamp_figures.py --write, never by hand."""
    import sys

    sys.path.insert(0, str(ROOT / "src"))
    from kitchensink4ppt import shipped

    published = next(f for f, what, _ in PUBLISHED_FIGURES if what == 'tests')
    assert f"{shipped.TESTS:,}" == published, (
        f"the server reports {shipped.TESTS:,} tests, the published figure "
        f"is {published}. Run scripts/stamp_figures.py --write, then "
        f"restamp {', '.join(FIGURE_SURFACES)} and PUBLISHED_FIGURES.")
