"""anonymize_deck and the identity left on classic comment authors.

PowerPoint 2013 and later add p15:presenceInfo to a classic comment author:
a userId (often a sign-in address) and a providerId. anonymize_deck renamed
the author but left both. They are now cleared, as the modern authors' are,
and the answer names the identity data the operation leaves as it is when
the deck holds it. Every name and address here is made up.
"""

from __future__ import annotations

from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import assembly, slides

P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main"
RT_AUTHORS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors")
AUTHORS = (
    '<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    f'xmlns:p15="{P15}"><p:cmAuthor id="1" name="Pat Example" initials="PE" lastIdx="1" '
    'clrIdx="0"><p:extLst><p:ext uri="{19B8F6BF-5375-455C-9EA6-DF929625EA0E}">'
    '<p15:presenceInfo userId="pat@example.com" providerId="AD"/></p:ext></p:extLst>'
    "</p:cmAuthor></p:cmAuthorLst>"
)


def _deck(path, embedded=False):
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    pkg.add_part_with_content_type(
        "ppt/commentAuthors.xml", AUTHORS.encode(),
        "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml")
    pkg.add_relationship("ppt/presentation.xml", RT_AUTHORS, "commentAuthors.xml")
    if embedded:
        pkg.add_part_with_content_type(
            "ppt/embeddings/Made_Up_Sheet1.xlsx", b"PK made-up bytes",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    pkg.save(do_backup=False)
    return path


def test_classic_author_presence_ids_are_cleared(tmp_path):
    path = _deck(tmp_path / "authors.pptx")
    pkg = PptxPackage(path)
    out = assembly.anonymize_deck(pkg)
    pkg.save(do_backup=False)
    xml = PptxPackage(path).part_bytes("ppt/commentAuthors.xml").decode()
    assert "pat@example.com" not in xml and 'providerId="AD"' not in xml
    assert "Pat Example" not in xml
    record = next(r for r in out["changed"] if r["where"].startswith("classic author"))
    assert record.get("userId_cleared") is True
    assert record.get("providerId_cleared") is True


def test_the_answer_names_identity_it_leaves(tmp_path):
    out = assembly.anonymize_deck(PptxPackage(_deck(tmp_path / "e.pptx", embedded=True)))
    assert any("embedded files" in w for w in out["warnings"])
    out = assembly.anonymize_deck(PptxPackage(_deck(tmp_path / "p.pptx")))
    assert not any("may still name people" in w for w in out["warnings"])
