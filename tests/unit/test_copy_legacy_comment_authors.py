"""copy_slide_between keeps classic comments with the people who wrote them.

A classic comment names its author by number (p:cm authorId), and the number
belongs to the source deck's commentAuthors part. The comment part was copied
as it was, so in the destination each comment showed under whoever had that
number there (or under no one, when nobody did). Now each author is matched
by name and initials in the destination, or added with a fresh number, and
the copied comments are renumbered to match. Every name here is made up.
"""

from __future__ import annotations

from kitchensink4ppt.core.package import PptxPackage, qn, rels_name, resolve_target
from kitchensink4ppt.ops import interdeck, slides

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT_AUTHORS = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"


def _authors(*people):
    rows = "".join(
        f'<p:cmAuthor id="{i}" name="{n}" initials="{ini}" lastIdx="{last}" clrIdx="{i}"/>'
        for i, n, ini, last in people)
    return f'<p:cmAuthorLst xmlns:p="{P}">{rows}</p:cmAuthorLst>'.encode()


def _comments(*rows):
    body = "".join(
        f'<p:cm authorId="{a}" dt="2026-01-01T00:00:00.000" idx="{i}"><p:pos x="10" y="10"/>'
        f"<p:text>{t}</p:text></p:cm>" for a, i, t in rows)
    return f'<p:cmLst xmlns:p="{P}">{body}</p:cmLst>'.encode()


def _deck(path, authors=None, comments=None):
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    if authors is not None:
        pkg.add_part_with_content_type("ppt/commentAuthors.xml", authors, CT_AUTHORS)
        pkg.add_relationship("ppt/presentation.xml", RT + "commentAuthors", "commentAuthors.xml")
    if comments is not None:
        slide = pkg.slide_parts()[0]
        pkg.add_part_with_content_type("ppt/comments/comment1.xml", comments, CT_COMMENTS)
        pkg.add_relationship(slide, RT + "comments", "../comments/comment1.xml")
    pkg.save(do_backup=False)
    return str(path)


def _copied_comments(path):
    """[(author name, idx, text)] of the last slide's classic comments."""
    pkg = PptxPackage(path)
    slide = pkg.slide_parts()[-1]
    part = next(resolve_target(slide, r.get("Target")) for r in pkg.root(rels_name(slide))
                if r.get("Type") == RT + "comments")
    names = {a.get("id"): a.get("name")
             for a in pkg.root("ppt/commentAuthors.xml").findall(qn("p:cmAuthor"))}
    return [(names.get(cm.get("authorId")), int(cm.get("idx")), cm.findtext(qn("p:text")))
            for cm in pkg.root(part).findall(qn("p:cm"))]


def _copy(dest, source):
    pkg = PptxPackage(dest)
    interdeck.copy_slide_between(pkg, source, 0)
    pkg.save(do_backup=False)


SOURCE_AUTHORS = _authors((0, "Alex Sample", "AS", 2))
SOURCE_COMMENTS = _comments((0, 1, "Made-up note one"), (0, 2, "Made-up note two"))


def test_comments_keep_their_author_when_the_number_means_someone_else(tmp_path):
    source = _deck(tmp_path / "src.pptx", SOURCE_AUTHORS, SOURCE_COMMENTS)
    dest = _deck(tmp_path / "dst.pptx", _authors((0, "Sam Placeholder", "SP", 1)),
                 _comments((0, 1, "Made-up note here")))
    _copy(dest, source)
    assert [c[0] for c in _copied_comments(dest)] == ["Alex Sample", "Alex Sample"]


def test_an_author_already_in_the_destination_is_reused_and_numbered_on(tmp_path):
    source = _deck(tmp_path / "src.pptx", SOURCE_AUTHORS, SOURCE_COMMENTS)
    dest = _deck(tmp_path / "dst.pptx", _authors((0, "Sam Placeholder", "SP", 1),
                                                 (3, "Alex Sample", "AS", 5)))
    _copy(dest, source)
    assert _copied_comments(dest) == [("Alex Sample", 6, "Made-up note one"),
                                      ("Alex Sample", 7, "Made-up note two")]
    authors = PptxPackage(dest).root("ppt/commentAuthors.xml").findall(qn("p:cmAuthor"))
    assert len(authors) == 2
    assert next(a for a in authors if a.get("id") == "3").get("lastIdx") == "7"


def test_a_destination_without_comment_authors_gets_them(tmp_path):
    source = _deck(tmp_path / "src.pptx", SOURCE_AUTHORS, SOURCE_COMMENTS)
    dest = _deck(tmp_path / "dst.pptx")
    _copy(dest, source)
    assert [c[0] for c in _copied_comments(dest)] == ["Alex Sample", "Alex Sample"]


def test_merge_decks_keeps_classic_comment_authors_too(tmp_path):
    from kitchensink4ppt.ops import assembly

    source = _deck(tmp_path / "src.pptx", SOURCE_AUTHORS, SOURCE_COMMENTS)
    dest = _deck(tmp_path / "dst.pptx", _authors((0, "Sam Placeholder", "SP", 1)),
                 _comments((0, 1, "Made-up note here")))
    pkg = PptxPackage(dest)
    assembly.merge_decks(pkg, [source])
    pkg.save(do_backup=False)
    assert [c[0] for c in _copied_comments(dest)] == ["Alex Sample", "Alex Sample"]
