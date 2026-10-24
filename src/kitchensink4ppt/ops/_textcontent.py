"""Shared encoding for plain text inside one DrawingML paragraph.

The public package-mode convention matches PowerPoint TextRange and
python-pptx: ``\v`` is an ``a:br`` soft line break inside a paragraph.
Callers split ``\n`` into separate ``a:p`` elements before using this
module.
"""

from __future__ import annotations

from collections.abc import Callable

from lxml import etree

from ..core.package import qn


SOFT_BREAK = "\v"


def append_text(
    paragraph: etree._Element,
    text: str,
    configure_rpr: Callable[[etree._Element], None] | None = None,
) -> bool:
    """Append runs and ``a:br`` nodes for one paragraph's plain text.

    ``configure_rpr`` receives a new ``a:rPr`` on every run and break.
    Returns whether any text/run-family element was appended.
    """
    chunks = str(text).split(SOFT_BREAK)
    appended = False
    for index, chunk in enumerate(chunks):
        if chunk:
            run = etree.SubElement(paragraph, qn("a:r"))
            if configure_rpr is not None:
                configure_rpr(etree.SubElement(run, qn("a:rPr")))
            etree.SubElement(run, qn("a:t")).text = chunk
            appended = True
        if index < len(chunks) - 1:
            br = etree.SubElement(paragraph, qn("a:br"))
            if configure_rpr is not None:
                configure_rpr(etree.SubElement(br, qn("a:rPr")))
            appended = True
    return appended
