"""One hardened way to parse XML in this server.

Every XML part this server reads goes through `parser()` or `fromstring()`:
no DTD loading, no network access, entities never substituted, and
libxml2's default size limits (huge_tree off).

A part that carries a document type declaration (<!DOCTYPE ...>) is refused
with a plain sentence, never parsed for its content. The Open Packaging
Conventions do not allow a DTD in a package's XML (ECMA-376 Part 2), so a
file PowerPoint wrote never has one. Refusing is the only safe reading:
- expanding the entities a DTD declares is how a "billion laughs" part
  makes a reader build gigabytes, and how an external entity reads another
  file or a network address;
- leaving them unexpanded keeps "&name;" references in the tree, which
  every text read skips, so the text would read wrong and a save would
  write the references back.

The check runs twice: a walk of each XML part's prolog (the bytes before
its root element, the only place XML allows the declaration) when a
package is opened, so the refusal comes before any parsing whichever
operation asked; and a look at the parsed document's own DOCTYPE after
every parse (a part added later, or an encoding the walk could not read).

The two outside libraries that parse XML here get the same guarantee
before they are called: svgelements reads the caller's SVG only after
`svg_without_doctype`, and mathml2omml reads converted math only after
`declares_dtd` says it has no declaration.
"""

from __future__ import annotations

import io
import re

from lxml import etree

#: The parser options every parse takes. dtd_validation and
#: attribute_defaults would read the DTD; all three entity switches are off.
_FLAGS = dict(
    resolve_entities=False,
    no_network=True,
    load_dtd=False,
    dtd_validation=False,
    attribute_defaults=False,
    huge_tree=False,
)

_FIXED = ("resolve_entities", "no_network", "load_dtd", "dtd_validation", "attribute_defaults")

#: "<!DOCTYPE" as it reads in the encodings an XML part may use: any
#: ASCII-compatible one (UTF-8 and the legacy single-byte and CJK
#: encodings), and UTF-16 and UTF-32 once decoded.
_MARK = "<!DOCTYPE"


def _checked(overrides: dict) -> dict:
    for key in _FIXED:
        if key in overrides:
            raise ValueError(f"xmlsafe: {key} is fixed")
    return {**_FLAGS, **overrides}


def parser(**overrides) -> etree.XMLParser:
    """A new hardened parser (lxml parsers are not shared between threads).
    `overrides` may set layout options (remove_blank_text, strip_cdata,
    remove_comments, remove_pis); never an entity or DTD switch."""
    return etree.XMLParser(**_checked(overrides))


def byte_encoding(data: bytes) -> str:
    """The encoding family of an XML part's bytes, from its byte order mark
    or its first character ("<" or the BOM-less first byte pair):
    "utf-32-le", "utf-32-be", "utf-16-le", "utf-16-be", or "utf-8" for
    every ASCII-compatible encoding."""
    head = data[:4]
    if head.startswith(b"\xff\xfe\x00\x00") or head == b"<\x00\x00\x00":
        return "utf-32-le"
    if head.startswith(b"\x00\x00\xfe\xff") or head == b"\x00\x00\x00<":
        return "utf-32-be"
    if head.startswith(b"\xff\xfe") or head.startswith(b"<\x00"):
        return "utf-16-le"
    if head.startswith(b"\xfe\xff") or head.startswith(b"\x00<"):
        return "utf-16-be"
    return "utf-8"


def _decoded(data: bytes) -> str:
    """A UTF-16 or UTF-32 part's characters, its byte order mark dropped
    (undecodable bytes replaced): what the prolog walk reads."""
    enc = byte_encoding(data)
    for bom in (b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff", b"\xff\xfe", b"\xfe\xff"):
        if data.startswith(bom):
            data = data[len(bom):]
            break
    return data.decode(enc, errors="replace")


def _prolog_doctype_at(data, lt, qm, pi_end, com_start, com_end, mark) -> int:
    """Walk the prolog (XML declaration, comments, processing instructions,
    white space) up to the root element's start tag; the index where a
    "<!DOCTYPE" comes first, else -1. XML allows the declaration nowhere
    else, so the walk reads only the few bytes before the root, however
    large the part."""
    i = 0
    while True:
        i = data.find(lt, i)
        if i < 0:
            return -1
        if data.startswith(mark, i):
            return i
        if data.startswith(com_start, i):
            j = data.find(com_end, i + len(com_start))
        elif data.startswith(lt + qm, i):
            j = data.find(pi_end, i + len(lt + qm))
        else:
            return -1  # the root element: no declaration before it
        if j < 0:
            return -1
        i = j + 1


def _doctype_index(text: str) -> int:
    return _prolog_doctype_at(text, "<", "?", "?>", "<!--", "-->", _MARK)


def declares_dtd(data: bytes | str) -> bool:
    """True when the part's prolog holds "<!DOCTYPE", read in the part's own
    encoding (UTF-16 and UTF-32 parts are decoded first; a str is read as
    it is)."""
    if isinstance(data, str):
        return _doctype_index(data) >= 0
    if byte_encoding(data) == "utf-8":
        return _prolog_doctype_at(data, b"<", b"?", b"?>", b"<!--", b"-->",
                                  _MARK.encode("ascii")) >= 0
    return _doctype_index(_decoded(data)) >= 0


def dtd_refusal(doc: str, part: str):
    """The UNSUPPORTED_CONTENT refusal for a part that declares a DTD
    (PackageRefused, a DocumentCorrupt that validate does not hand to
    PowerPoint)."""
    from .errors import PackageRefused

    return PackageRefused(
        f"{doc}: its part {part} contains a document type declaration "
        "(<!DOCTYPE ...>), which a PowerPoint presentation never has (the Open XML package "
        "rules do not allow one). It was not read: such a declaration can make an XML reader "
        "expand text without limit or read other files. Nothing was changed. If another "
        "program made this file, export it again from that program as a PowerPoint "
        "presentation.")


def has_doctype(root: etree._Element) -> bool:
    """True when the document that `root` belongs to declared a DTD."""
    info = root.getroottree().docinfo
    return bool(info.doctype) or info.internalDTD is not None or info.externalDTD is not None


def is_amplification(exc: Exception) -> bool:
    """libxml2's stop for an entity that expands past its amplification
    limit (measured with lxml 6.1.3 / libxml2 2.11.9: "Maximum entity
    amplification factor exceeded"), which only a DTD's entities reach."""
    text = str(exc).lower()
    return "entity" in text and "amplification" in text


def refuse_dtd(data: bytes, *, doc: str, part: str) -> None:
    """Raise the refusal when `data` declares a DTD. The prolog walk is
    exact for XML (a declaration can stand nowhere else), so a hit is
    refused at once, without parsing anything."""
    if declares_dtd(data):
        raise dtd_refusal(doc, part)


def fromstring(data: bytes | str, *, doc: str = "the presentation", part: str = "an XML part",
               **overrides) -> etree._Element:
    """Parse one part's bytes with the hardened parser. A part that declares
    a DTD is refused (DocumentCorrupt, before parsing when the prolog walk
    sees it); a part that does not parse raises lxml's XMLSyntaxError as
    before, except libxml2's entity-amplification stop, which is the
    refusal too. A str (markup this server builds) is parsed as UTF-8."""
    if isinstance(data, str):
        data = data.encode("utf-8")
    refuse_dtd(data, doc=doc, part=part)
    try:
        root = etree.fromstring(data, parser=parser(**overrides))
    except etree.XMLSyntaxError as exc:
        if is_amplification(exc):
            raise dtd_refusal(doc, part) from None
        raise
    if has_doctype(root):
        raise dtd_refusal(doc, part)
    return root



def root_namespace(data: bytes, *, doc: str = "the presentation",
                   part: str = "an XML part") -> str | None:
    """The namespace of a part's root element, read from its first start tag
    with the hardened flags, so a large part is never parsed whole; None
    when the part is not XML. A part that declares a DTD is refused."""
    refuse_dtd(data, doc=doc, part=part)
    try:
        for _event, element in etree.iterparse(io.BytesIO(data), events=("start",), **_FLAGS):
            return etree.QName(element).namespace
    except etree.XMLSyntaxError as exc:
        if is_amplification(exc):
            raise dtd_refusal(doc, part) from None
        return None
    return None

#: A DOCTYPE that only names a DTD (a public or system identifier, or
#: nothing) and declares nothing itself: no "[ ... ]" internal subset.
_BARE_DOCTYPE = re.compile(
    r"<!DOCTYPE\s+[^\s\[>]+"
    r"(?:\s+(?:SYSTEM\s+(?:\"[^\"]*\"|'[^']*')"
    r"|PUBLIC\s+(?:\"[^\"]*\"|'[^']*')\s+(?:\"[^\"]*\"|'[^']*')))?"
    r"\s*>")


def svg_without_doctype(svg_text: str) -> str:
    """The caller's SVG as every parser here reads it.

    A DOCTYPE that only names a DTD, such as the SVG 1.1 header many
    drawing programs still write, declares nothing: no parser here loads
    the DTD it names. It is dropped. A DOCTYPE with declarations of its own
    (a "[ ... ]" block, which is where entities live) is refused, before
    any parser reads the drawing."""
    from .errors import PptMcpError

    at = _doctype_index(svg_text)
    if at < 0:
        return svg_text
    bare = _BARE_DOCTYPE.match(svg_text, at)
    if bare is None:
        raise PptMcpError(
            "svg has a document type declaration with declarations of its own "
            "(<!DOCTYPE svg [ ... ]>), which this server does not read: such declarations can "
            "make an XML reader expand text without limit or read other files. Nothing was "
            "changed. Remove the bracketed block, replace any &name; reference that used it "
            "with its value, and send the SVG again.")
    return svg_text[:at] + svg_text[bare.end():]
