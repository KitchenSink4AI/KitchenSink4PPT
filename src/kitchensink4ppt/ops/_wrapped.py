"""Text PowerPoint shows from a compatibility wrapper inside a paragraph.

PowerPoint can store part of a paragraph in mc:AlternateContent: an inline
equation (an mc:Choice that Requires a14 and holds a14:m math, with an
mc:Fallback run for older readers), or a run written for a newer version (an
mc:Choice that Requires a namespace an older PowerPoint does not know, with a
Fallback run). Markup compatibility has a reader take the first Choice whose
Requires it understands, else the Fallback, and PowerPoint does so: measured
on made-up decks, a Choice that Requires an unknown namespace shows its
Fallback's words, and one that Requires a14 shows the equation.

Text reads report that branch, so a word PowerPoint shows is never missing
from them: an equation reads as the characters of its math runs (as
list_equations reports it; PowerPoint's own text adds spacing for the math
layout), and other content as its runs' text. Edits never rewrite a wrapper:
the run map treats one as a single unit that edits go around.
"""

from __future__ import annotations

from lxml import etree

MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
AC_TAG = f"{{{MC_NS}}}AlternateContent"
_CHOICE = f"{{{MC_NS}}}Choice"
_FALLBACK = f"{{{MC_NS}}}Fallback"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_A_T = f"{{{_A_NS}}}t"
_M_T = "{http://schemas.openxmlformats.org/officeDocument/2006/math}t"
_A14_M = "{http://schemas.microsoft.com/office/drawing/2010/main}m"
#: The tag of math written straight into a paragraph (ops/_runmap.build_map).
A14_M_TAG = _A14_M

#: Namespaces PowerPoint reads, beyond the standard ones (any namespace under
#: schemas.openxmlformats.org): the Office extensions found in PowerPoint's
#: own files. A Choice that Requires anything else is not shown.
_MS = "http://schemas.microsoft.com/office/"
UNDERSTOOD = frozenset({
    "urn:schemas-microsoft-com:vml",
    "urn:schemas-microsoft-com:office:office",
    "urn:schemas-microsoft-com:office:powerpoint",
    _MS + "drawing/2007/8/2/chart",
    _MS + "drawing/2008/diagram",
    _MS + "drawing/2010/main",
    _MS + "drawing/2010/chartDrawing",
    _MS + "drawing/2010/diagram",
    _MS + "drawing/2010/slicer",
    _MS + "drawing/2012/main",
    _MS + "drawing/2012/chart",
    _MS + "drawing/2012/chartStyle",
    _MS + "drawing/2013/main/command",
    _MS + "drawing/2014/main",
    _MS + "drawing/2014/chart",
    _MS + "drawing/2014/chartex",
    _MS + "drawing/2015/06/chart",
    _MS + "drawing/2015/9/8/chartex",
    _MS + "drawing/2015/10/21/chartex",
    _MS + "drawing/2016/5/9/chartex",
    _MS + "drawing/2016/5/10/chartex",
    _MS + "drawing/2016/5/11/chartex",
    _MS + "drawing/2016/5/12/chartex",
    _MS + "drawing/2016/5/13/chartex",
    _MS + "drawing/2016/5/14/chartex",
    _MS + "drawing/2016/11/main",
    _MS + "drawing/2016/11/diagram",
    _MS + "drawing/2016/12/diagram",
    _MS + "drawing/2016/SVG/main",
    _MS + "drawing/2017/03/chart",
    _MS + "drawing/2017/decorative",
    _MS + "drawing/2017/model3d",
    _MS + "drawing/2018/animation",
    _MS + "drawing/2018/animation/model3d",
    _MS + "drawing/2018/hyperlinkcolor",
    _MS + "drawing/2018/sketchyshapes",
    _MS + "drawing/2020/classificationShape",
    _MS + "powerpoint/2010/main",
    _MS + "powerpoint/2012/main",
    _MS + "powerpoint/2013/main/command",
    _MS + "powerpoint/2015/09/main",
    _MS + "powerpoint/2015/10/main",
    _MS + "powerpoint/2016/slidezoom",
    _MS + "powerpoint/2016/sectionzoom",
    _MS + "powerpoint/2016/summaryzoom",
    _MS + "powerpoint/2017/10/main",
    _MS + "powerpoint/2018/4/main",
    _MS + "powerpoint/2018/8/main",
    _MS + "powerpoint/2019/12/main",
    _MS + "powerpoint/2020/02/main",
    _MS + "powerpoint/2021/06/main",
    _MS + "powerpoint/2022/03/main",
    _MS + "powerpoint/2022/08/main",
    _MS + "thememl/2012/main",
    _MS + "webextensions/webextension/2010/11",
})


def understood(namespace: str | None) -> bool:
    """Whether PowerPoint reads elements of `namespace`."""
    if not namespace:
        return False
    return namespace in UNDERSTOOD or namespace.startswith("http://schemas.openxmlformats.org/")


def shown_branch(ac: etree._Element) -> etree._Element | None:
    """The branch of one mc:AlternateContent that PowerPoint shows: the first
    Choice whose Requires prefixes all name namespaces it reads, else the
    Fallback, else None."""
    for branch in ac:
        if branch.tag == _CHOICE:
            needs = (branch.get("Requires") or "").split()
            if needs and all(understood(branch.nsmap.get(p)) for p in needs):
                return branch
    return ac.find(_FALLBACK)


def wrapped_text(ac: etree._Element) -> str:
    """The text PowerPoint shows for one paragraph-level wrapper, in the
    read layer's convention: runs and fields by their text, a line break as
    \\v, an equation as the characters of its math runs."""
    if ac.tag == _A14_M:
        # Math written straight into a paragraph (in a box PowerPoint's a14
        # branch holds) reads as its characters, as in a wrapper.
        return "".join(t.text or "" for t in ac.iter(_M_T))
    branch = shown_branch(ac)
    if branch is None:
        return ""
    parts: list[str] = []
    for child in branch:
        if not isinstance(child.tag, str):
            continue
        if child.tag == AC_TAG:
            parts.append(wrapped_text(child))
            continue
        local = etree.QName(child).localname
        namespace = etree.QName(child).namespace
        if namespace == _A_NS and local in ("r", "fld"):
            t = child.find(_A_T)
            if t is not None and t.text:
                parts.append(t.text)
        elif namespace == _A_NS and local == "br":
            parts.append("\v")
        elif child.tag == _A14_M:
            parts.append("".join(t.text or "" for t in child.iter(_M_T)))
        else:
            parts.append("".join(
                t.text or "" for t in child.iter(_A_T, _M_T)
            ))
    return "".join(parts)

