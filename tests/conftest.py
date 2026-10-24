"""Shared test setup.

The suite runs on generated decks only. tests/make_corpus.py builds every
deck in tests/corpus/ from code (made-up filler text, solid-colour pictures, a
table, speaker notes) at session start, and no test reads any other deck. A
file under one of the corpus names that the generator did not write stops the
run instead of being read; move it out of tests/corpus/ first.

If generation itself fails (python-pptx missing), corpus-dependent tests are
skipped with an explanation and the rest still run.
"""

import hashlib
import os
import sys
import tempfile
from pathlib import Path

import pytest

CORPUS = Path(__file__).parent / "corpus"
#: The Business edition's own sources: its launcher package, its licence
#: notice and its signing script. The development tree carries them; the
#: Community source tree does not, so the tests that read them skip there.
BUSINESS_SOURCES = all(
    (Path(__file__).resolve().parents[1] / rel).exists() for rel in (
        "packages/kitchensink4ppt-business/pyproject.toml",
        "business/COMMERCIAL_LICENSE_NOTICE.md",
        "scripts/sign_business_license.py"))
SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"

# The core does not choose an edition and does not run without a policy, so
# the suite plays the part a launcher plays in a real installation: it injects
# one explicitly, before any test module imports kitchensink4ppt.server.
# Community is what a source checkout exercises by default; the Business tests
# build and inject their own policy per test.
#
# The two license policy adapters are ONE shared source serving every editor
# in the suite and are never vendored here. adapter_source materialises the
# pinned canonical revision into OS scratch and verifies its digest, so a
# forked or drifted adapter fails loudly rather than quietly becoming this
# product's private copy. See scripts/adapter_source.py.
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import adapter_source  # noqa: E402

adapter_source.install_import_path()

# EDITIONS. KS4_ADAPTER_EDITION=community is the public checkout: it carries the
# Community edition only (no Business adapter, Business launcher or business/
# tree). The modules and tests listed below exercise Business code (or, for a
# few tests, the commercial notice or an internal audit the public tree does
# not carry; each entry says which). They run wherever those sources exist
# (private staging and every Business build) and are not collected or are
# deselected on the Community path; the session header names them. tests/unit/test_business_only_manifest.py
# keeps the lists exact: every test module that imports Business code at top
# level is listed, and nothing else is.
EDITION = (os.environ.get("KS4_ADAPTER_EDITION") or "").strip().lower() or "all"
if EDITION not in ("all", "community"):
    raise RuntimeError(f"KS4_ADAPTER_EDITION={EDITION!r}: expected all or community")
BUSINESS_ONLY_MODULES = (
    "unit/test_chain_verification.py",
    "unit/test_editions.py",
    "unit/test_product_neutrality.py",
    "unit/test_receipt_compat.py",
    "unit/test_update_business.py",
)
BUSINESS_ONLY_TESTS = {
    "tests/unit/test_license_seam.py::test_setup_error_is_not_a_receipt_error":
        "Business adapter or launcher code",
    "tests/unit/test_license_seam.py::test_business_package_imports_no_community_module":
        "Business adapter source; on the Community path it would pass vacuously over a directory that does not exist",
    "tests/unit/test_license_seam.py::test_business_package_does_not_depend_on_the_editing_core":
        "Business adapter source; on the Community path it would pass vacuously over a directory that does not exist",
    "tests/unit/test_license_seam.py::test_business_package_imports_cleanly_without_community_installed":
        "Business adapter or launcher code",
    "tests/unit/test_license_seam.py::test_this_core_composes_the_powerpoint_order[business]":
        "Business adapter or launcher code",
    "tests/unit/test_license_seam.py::test_additional_tool_registration_is_optional_and_neither_policy_uses_it":
        "Business adapter or launcher code",
    "tests/unit/test_license_seam.py::test_both_shipped_policies_satisfy_the_tightened_validator":
        "Business adapter or launcher code",
    "tests/unit/test_license_seam.py::test_only_the_business_policy_declares_an_update_backend":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_the_allowlist_lets_the_commercial_notice_keep_its_agpl_sentences":
        "the commercial notice business/COMMERCIAL_LICENSE_NOTICE.md (not in the public tree)",
    "tests/unit/test_build_scan_gate.py::test_business_build_refuses_metadata_drift":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_a_real_business_build_ships_a_commercially_labelled_core":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_a_poisoned_core_label_fails_the_real_business_build":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_the_required_notices_survive_the_commercial_relabelling":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_a_real_business_build_passes_the_gate_and_ships_no_community_module":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_a_poisoned_build_input_fails_the_real_build":
        "Business adapter or launcher code",
    "tests/unit/test_dependency_ledger.py::test_ledger_exists":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
    "tests/unit/test_dependency_ledger.py::test_every_required_dependency_is_in_the_ledger":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
    "tests/unit/test_dependency_ledger.py::test_every_optional_dependency_is_in_the_ledger":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
    "tests/unit/test_dependency_ledger.py::test_no_copyleft_in_the_required_install":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
    "tests/unit/test_dependency_ledger.py::test_required_licenses_are_recognizably_permissive":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
    "tests/unit/test_dependency_ledger.py::test_the_ledger_records_the_font_attribution":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
}
collect_ignore = list(BUSINESS_ONLY_MODULES) if EDITION == "community" else []


def pytest_report_header(config):
    if EDITION != "community":
        return ["edition: all (Community and Business tests)"]
    return [f"edition: Community only. Not run here (Business code, run in private staging): "
            f"{len(BUSINESS_ONLY_MODULES)} modules {list(BUSINESS_ONLY_MODULES)}, "
            f"{len(BUSINESS_ONLY_TESTS)} tests {list(BUSINESS_ONLY_TESTS)}"]


def _deselect_business_only(config, items):
    if EDITION != "community" or not BUSINESS_ONLY_TESTS:
        return
    keep, drop = [], []
    for item in items:
        # an entry names one parametrized case exactly, or every case of a test
        hit = item.nodeid in BUSINESS_ONLY_TESTS or item.nodeid.split("[", 1)[0] in BUSINESS_ONLY_TESTS
        (drop if hit else keep).append(item)
    if drop:
        config.hook.pytest_deselected(items=drop)
        items[:] = keep


for _pkg in (("kitchensink4ppt-community",) if EDITION == "community"
             else ("kitchensink4ppt-community", "kitchensink4ppt-business")):
    _src = str(Path(__file__).resolve().parents[1] / "packages" / _pkg / "src")
    if _src not in sys.path:
        sys.path.insert(0, _src)

from ks4_license_community.policy import CommunityPolicy as _CommunityPolicy  # noqa: E402
from ks4_license_community.policy import ProductConfig as _CommunityConfig  # noqa: E402
from kitchensink4ppt import __version__ as _core_version  # noqa: E402
from kitchensink4ppt import license_policy as _license_policy  # noqa: E402

# Product identity comes from the launcher in a real installation; here the
# suite plays that part, naming the product explicitly rather than relying on
# a default the adapters deliberately no longer have.
_license_policy.install(_CommunityPolicy(
    _CommunityConfig(product_id="kitchensink4ppt",
                     display_name="KitchenSink4PPT"),
    _core_version))

# name -> what the tests use it for (every one generated by tests/make_corpus.py)
EXPECTED = {
    "proposal_deck.pptx": "26-slide templated deck (the key deck)",
    "conference_talk.pptx": "conference talk deck",
    "image_deck.pptx": "conference deck with images",
    "conference_template.potx": "PowerPoint template package (.potx)",
    "heavy_deck.pptx": "heavy deck with many layouts and shapes",
    "table_deck.pptx": "deck with tables",
}


#: Why the corpus could not be generated, when it could not.
_CORPUS_PROBLEM = None


def _present() -> set[str]:
    return {name for name in EXPECTED if (CORPUS / name).exists()}


def pytest_configure(config):
    """Every corpus deck is GENERATED (tests/make_corpus.py), so the full
    suite runs anywhere, CI included, and only ever on made-up content."""
    global _CORPUS_PROBLEM
    # NO NETWORK IN TESTS. diagnose runs the on-demand update check, so the
    # suite turns it off for every test by default; the update tests switch
    # it back on for themselves and mock the fetch. setdefault, so a
    # developer who exports the variable still wins.
    os.environ.setdefault("KS4P_UPDATE_CHECK", "off")
    # Saves write down each swap in a journal in the server's own folder; the
    # tests keep theirs apart from a real server's.
    os.environ.setdefault(
        "KS4P_SAVE_JOURNAL_DIR",
        os.path.join(tempfile.gettempdir(), f"ks4p-test-journal-{os.getpid()}"))
    config.addinivalue_line(
        "markers",
        "public_copy: release-only checks enabled by KS4P_RELEASE_GATE=1",
    )
    config.addinivalue_line(
        "markers",
        "business_sources: reads the Business edition sources, which the "
        "Community source tree does not carry",
    )
    import make_corpus  # (lives beside this file)

    assert set(EXPECTED) == set(make_corpus.GENERATORS), "corpus names out of step"
    try:
        made = make_corpus.ensure_corpus(verbose=True)
    except make_corpus.ForeignDeck as exc:
        raise pytest.UsageError(str(exc)) from None
    except Exception as exc:  # generation needs python-pptx (the dev extra)
        _CORPUS_PROBLEM = f"{type(exc).__name__}: {exc}"
    else:
        if made:
            print(f"conftest: generated {len(made)} corpus deck(s)")


def pytest_collection_modifyitems(config, items):
    _deselect_business_only(config, items)
    # The public-copy checks run only when KS4P_RELEASE_GATE=1; ordinary runs
    # collect them and skip them.
    if os.environ.get("KS4P_RELEASE_GATE") != "1":
        skip_public_copy = pytest.mark.skip(
            reason="public_copy release gate: set KS4P_RELEASE_GATE=1"
        )
        for item in items:
            if item.get_closest_marker("public_copy"):
                item.add_marker(skip_public_copy)
    if not BUSINESS_SOURCES:
        skip_business = pytest.mark.skip(
            reason="the Business edition sources are not in this copy of the source tree"
        )
        for item in items:
            if item.get_closest_marker("business_sources"):
                item.add_marker(skip_business)

    missing = set(EXPECTED) - _present()
    if not missing:
        return
    skip = pytest.mark.skip(
        reason=(
            f"test corpus incomplete (missing: {sorted(missing)}): generation "
            f"failed ({_CORPUS_PROBLEM or 'no reason recorded'}); run "
            "python tests/make_corpus.py for details."
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


@pytest.fixture(autouse=True)
def own_save_journal(request, monkeypatch, tmp_path_factory):
    """Every test writes the swaps of its saves down in a journal folder of
    its own, never the user's and never another test's. A process the test
    starts inherits it through the environment. What a recovery left for
    the next answer does not carry over into the next test."""
    digest = hashlib.sha1(
        request.node.nodeid.encode("utf-8", "surrogatepass")).hexdigest()[:16]
    monkeypatch.setenv(
        "KS4P_SAVE_JOURNAL_DIR",
        str(tmp_path_factory.getbasetemp() / "save-journals" / digest))
    yield
    take = getattr(sys.modules.get("kitchensink4ppt.core.safesave"),
                   "take_notices", None)
    if take is not None:
        take()


@pytest.fixture()
def make_deck(tmp_path):
    """Build a small synthetic deck in tmp_path and return its Path."""

    def _make(name: str = "deck.pptx", *, seed: int = 0, extra_slides: int = 2):
        import make_corpus

        return make_corpus.build_deck(
            tmp_path / name, seed=seed, extra_slides=extra_slides
        )

    return _make
