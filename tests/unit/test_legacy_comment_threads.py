"""Classic comments that PowerPoint threaded are listed as threads.

Since PowerPoint 2013 a reply to a classic comment is itself a classic
comment whose p15:threadingInfo names its parent (p15:parentCm authorId and
idx). list_comments and comment_report listed every classic comment as a
thread of its own, so a reply read as a new thread and total_replies stayed
0. Replies now sit under their thread's first comment, as PowerPoint shows
them; a reply whose parent is missing stays a thread so nothing is hidden. A
copied slide's replies keep pointing at their thread when the copy renumbers
the comments.

Every name here is made up.
"""

from __future__ import annotations

from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.ops import comments, interdeck, slides

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main"
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT_AUTHORS = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"
THREAD_EXT = "{C676402C-5697-4E1C-873F-D02D1690AC5C}"


def _authors(*people):
    rows = "".join(
        f'<p:cmAuthor id="{i}" name="{n}" initials="{ini}" lastIdx="{last}" clrIdx="{i}"/>'
        for i, n, ini, last in people)
    return f'<p:cmAuthorLst xmlns:p="{P}">{rows}</p:cmAuthorLst>'.encode()


def _comments(*rows):
    """rows: (authorId, idx, day, text, parent) with parent (authorId, idx) or None."""
    body = []
    for a, i, day, text, parent in rows:
        inner = f'<p15:parentCm authorId="{parent[0]}" idx="{parent[1]}"/>' if parent else ""
        body.append(
            f'<p:cm authorId="{a}" dt="2026-03-{day:02d}T09:00:00.000" idx="{i}">'
            f'<p:pos x="10" y="10"/><p:text>{text}</p:text><p:extLst><p:ext uri="{THREAD_EXT}">'
            f'<p15:threadingInfo xmlns:p15="{P15}" timeZoneBias="0">{inner}</p15:threadingInfo>'
            "</p:ext></p:extLst></p:cm>")
    return f'<p:cmLst xmlns:p="{P}">{"".join(body)}</p:cmLst>'.encode()


def _deck(path, authors=None, cms=None):
    slides.create_presentation(path)
    pkg = PptxPackage(path)
    slides.insert_slide(pkg, 6)
    if authors is not None:
        pkg.add_part_with_content_type("ppt/commentAuthors.xml", authors, CT_AUTHORS)
        pkg.add_relationship("ppt/presentation.xml", RT + "commentAuthors", "commentAuthors.xml")
    if cms is not None:
        slide = pkg.slide_parts()[0]
        pkg.add_part_with_content_type("ppt/comments/comment1.xml", cms, CT_COMMENTS)
        pkg.add_relationship(slide, RT + "comments", "../comments/comment1.xml")
    pkg.save(do_backup=False)
    return str(path)


AUTHORS = _authors((0, "Alex Sample", "AS", 2), (1, "Rowan Example", "RE", 2))
THREAD = _comments(
    (0, 1, 1, "Made-up thread start", None),
    (1, 1, 2, "Made-up reply", (0, 1)),
    (0, 2, 3, "Made-up reply to the reply", (1, 1)),
    (1, 2, 4, "Made-up reply to a comment that is gone", (0, 9)),
)


def test_replies_are_nested_under_their_thread(tmp_path):
    deck = _deck(tmp_path / "threads.pptx", AUTHORS, THREAD)
    out = comments.list_comments(PptxPackage(deck))
    assert out["total_comments"] == 2 and out["total_replies"] == 2
    first, orphan = out["slides"][0]["comments"]
    assert first["comment_id"] == "legacy-0-1" and first["text"] == "Made-up thread start"
    assert [(r["reply_id"], r["author"], r["text"]) for r in first["replies"]] == [
        ("legacy-1-1", "Rowan Example", "Made-up reply"),
        ("legacy-0-2", "Alex Sample", "Made-up reply to the reply"),
    ]
    assert first["replies"][0]["created"] == "2026-03-02T09:00:00.000"
    # A reply whose parent is not in the deck stays a thread of its own.
    assert orphan["comment_id"] == "legacy-1-2" and orphan["replies"] == []


def test_comment_report_counts_and_indents_the_replies(tmp_path):
    deck = _deck(tmp_path / "threads.pptx", AUTHORS, THREAD)
    out = comments.comment_report(PptxPackage(deck))
    assert out["total_comments"] == 2 and out["total_replies"] == 2
    assert "    - **Rowan Example** (2026-03-02T09:00:00.000): Made-up reply" in out["markdown"]


def test_a_copied_thread_keeps_its_replies_when_renumbered(tmp_path):
    source = _deck(tmp_path / "src.pptx", AUTHORS, _comments(
        (0, 1, 1, "Made-up thread start", None),
        (1, 1, 2, "Made-up reply", (0, 1)),
    ))
    # Alex already has comments in the destination, so the copied ones are
    # numbered after them, and Rowan is added with a new number.
    dest = _deck(tmp_path / "dst.pptx", _authors((0, "Sam Placeholder", "SP", 1),
                                                 (3, "Alex Sample", "AS", 5)))
    pkg = PptxPackage(dest)
    interdeck.copy_slide_between(pkg, source, 0)
    pkg.save(do_backup=False)
    out = comments.list_comments(PptxPackage(dest))
    copied = out["slides"][-1]["comments"]
    assert len(copied) == 1
    assert copied[0]["comment_id"] == "legacy-3-6"
    assert [(r["author"], r["text"]) for r in copied[0]["replies"]] == [("Rowan Example", "Made-up reply")]
