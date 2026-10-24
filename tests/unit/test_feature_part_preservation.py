"""Feature-part preservation: every corpus deck validates, and a benign text
edit keeps a deck's feature parts and relationships byte-identical.

Every deck here is generated. The corpus decks come from tests/make_corpus.py;
the feature decks are built per test from the generator's base deck plus a
made-up feature payload (a custom XML data part with its properties part, and
an embedded font part with its presentation entry), stored the way Office
stores them.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import make_corpus
import pytest

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops.read import _cnvpr, iter_shapes, shape_text, slide_table
from kitchensink4ppt.ops.text import format_text


CORPUS = make_corpus.CORPUS

#: feature -> (adds it to a generated deck, selects its parts from the package)
FEATURES = {
    "custom_xml": (
        make_corpus.add_custom_xml,
        lambda parts: {name for name in parts if name.startswith("customXml/")},
    ),
    "embedded_font": (
        make_corpus.add_embedded_font,
        lambda parts: {
            name
            for name in parts
            if name.startswith("ppt/fonts/") or name.lower().endswith(".fntdata")
        },
    ),
}


def _archive(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


def _first_text_shape(pkg: PptxPackage) -> tuple[int, int]:
    for slide in slide_table(pkg):
        tree = pkg.root(slide["part"]).find(f"{qn('p:cSld')}/{qn('p:spTree')}")
        if tree is None:
            continue
        for elem, kind, _z, _parent in iter_shapes(tree):
            if kind in {"group", "picture", "chart", "diagram", "ole", "graphicFrame", "table"}:
                continue
            cnvpr = _cnvpr(elem)
            if cnvpr is not None and shape_text(elem).strip():
                return slide["index"], int(cnvpr.get("id"))
    raise AssertionError("the generated deck has no ordinary text shape for a benign edit")


@pytest.mark.parametrize("name", sorted(make_corpus.GENERATORS))
def test_corpus_deck_relationship_graph_validates(name):
    source = CORPUS / name
    if not source.exists():
        pytest.skip(f"corpus deck missing: {name}")
    PptxPackage._validate_payload(source.read_bytes())


@pytest.mark.parametrize("feature", sorted(FEATURES))
def test_text_edit_preserves_feature_parts_and_relationships(feature, tmp_path):
    add_feature, feature_parts = FEATURES[feature]
    work = make_corpus.build_deck(tmp_path / f"{feature}.pptx", seed=5, extra_slides=1)
    add_feature(work)
    PptxPackage._validate_payload(work.read_bytes())
    before = _archive(work)
    protected = feature_parts(before)
    assert protected, f"the generator added no {feature} parts"

    pkg = PptxPackage(work)
    slide, shape_id = _first_text_shape(pkg)
    changed = format_text(pkg, slide, shape_id, bold=True)
    assert changed["runs_formatted"] > 0
    pkg.save(do_backup=False)

    after = _archive(work)
    assert set(after) == set(before), "a benign text edit added or dropped package parts"
    for part in protected:
        assert after[part] == before[part], f"feature payload changed: {part}"
    for part in before:
        if part.endswith(".rels"):
            assert after[part] == before[part], f"relationship graph changed: {part}"
    assert after["[Content_Types].xml"] == before["[Content_Types].xml"]

    PptxPackage._validate_payload(work.read_bytes())
    reopened = PptxPackage(work)
    assert reopened.slide_parts()
