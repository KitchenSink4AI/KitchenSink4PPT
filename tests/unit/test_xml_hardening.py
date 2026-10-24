"""Every XML part of a deck is read with one hardened parser; a part that
declares a document type (a DTD) is refused plainly and fast; a ZIP package
is checked before any part of it is inflated.

A PowerPoint package's XML never carries a document type declaration (the
Open Packaging Conventions do not allow one). A DTD is how a part expands text
without limit ("billion laughs", quadratic blowup) or reads another file or a
network address (external and parameter entities). Parts were parsed with
lxml's defaults, which expanded a small internal entity silently in every
part kind; the attacks came back as the XML library's raw message; and a part
that an operation never parsed rode along into every save, DTD and all.

The same holds for the two outside libraries that parse XML for this server:
svgelements reads the caller's SVG and mathml2omml reads the converted math.
Neither ever sees a DTD.

Reading a package used to inflate every part in full: a few megabytes of
compressed zeros could ask for gigabytes. A package that names one part twice
was read from its last copy, whichever copy PowerPoint shows. Member names
never become file paths.

Every package, drawing and equation here is made up in code."""

from __future__ import annotations

import ast
import asyncio
import hashlib
import io
import json
import socket
import threading
import time
import tracemalloc
import warnings
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from kitchensink4ppt import server
from kitchensink4ppt.core.errors import DocumentCorrupt, ValidationFailed
from kitchensink4ppt.core.package import PptxPackage
from kitchensink4ppt.surface import envelope

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src" / "kitchensink4ppt"
MiB = 1 << 20

#: Words of the XML library's own messages; none may reach a caller.
RAW = ("amplification", "Entity '", "PEReference", "line 1, column", "not defined")

PARTS = {
    "presentation": "ppt/presentation.xml",
    "slide": "ppt/slides/slide1.xml",
    "layout": "ppt/slideLayouts/slideLayout1.xml",
    "master": "ppt/slideMasters/slideMaster1.xml",
    "notes": "ppt/notesSlides/notesSlide1.xml",
    "notes_master": "ppt/notesMasters/notesMaster1.xml",
    "chart": "ppt/charts/chart1.xml",
    "comments": "ppt/comments/modernComment_1.xml",
    "authors": "ppt/authors.xml",
    "theme": "ppt/theme/theme1.xml",
    "content_types": "[Content_Types].xml",
    "presentation_rels": "ppt/_rels/presentation.xml.rels",
    "core_properties": "docProps/core.xml",
}


def _call(name: str, **kw) -> dict:
    """One operation in process, answered in the envelope the tools use."""
    try:
        return envelope.success(getattr(server, name)(**kw))
    except BaseException as exc:  # noqa: BLE001 - the envelope maps it, as the tools do
        return envelope.refusal(exc)


def _md5(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


def _wire(*calls: tuple[str, dict]) -> list[tuple[dict, bool, float]]:
    """Each (op, args) through the public tool surface, one client session."""
    from fastmcp import Client

    from kitchensink4ppt.registry import OP_TO_GROUP

    async def run():
        out = []
        async with Client(server.create_server()) as client:
            for op, args in calls:
                t0 = time.perf_counter()
                res = await client.call_tool(OP_TO_GROUP[op], {"op": op, "args": args},
                                             raise_on_error=False)
                text = "".join(b.text for b in res.content if hasattr(b, "text"))
                out.append((json.loads(text), res.is_error, time.perf_counter() - t0))
        return out

    return asyncio.run(run())


@pytest.fixture(scope="module")
def base_deck(tmp_path_factory) -> Path:
    """A made-up deck holding every part kind a deck reads: slides, layouts,
    the master, speaker notes and the notes master, a chart with its
    embedded workbook, a modern comment and its authors part."""
    f = tmp_path_factory.mktemp("hardening") / "base.pptx"
    steps = [
        ("create_presentation", {"path": str(f)}),
        ("insert_slide", {"file_path": str(f), "layout": 0, "live": "off"}),
        ("set_placeholder_text", {"file_path": str(f), "slide": 0, "placeholder": "title",
                                  "text": "Made-up title", "live": "off"}),
        ("set_notes", {"file_path": str(f), "slide": 0, "text": "Made-up speaker note",
                       "live": "off"}),
        ("insert_slide", {"file_path": str(f), "layout": 5, "live": "off"}),
        ("create_chart", {"file_path": str(f), "slide": 1, "chart_type": "column",
                          "series": [{"name": "Made-up series", "values": [1, 2, 3]}],
                          "categories": ["Alpha", "Beta", "Gamma"], "x": 1, "y": 1, "w": 5,
                          "h": 3, "title": "Made-up chart title"}),
        ("add_comment", {"file_path": str(f), "slide": 0, "text": "Made-up comment text",
                         "author": "assistant"}),
    ]
    for name, kw in steps:
        out = _call(name, **kw)
        assert out.get("ok"), (name, out)
    with zipfile.ZipFile(f) as zf:
        assert set(PARTS.values()) <= set(zf.namelist())
    return f


def _part(deck: Path, name: str) -> bytes:
    with zipfile.ZipFile(deck) as zf:
        return zf.read(name)


def variant(base: Path, out: Path, replace: dict[str, bytes] | None = None,
            extra: list[tuple[str, bytes]] | None = None) -> Path:
    """A copy of `base` with parts replaced and members appended."""
    replace = replace or {}
    with zipfile.ZipFile(base) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            dst.writestr(info.filename, replace.get(info.filename, src.read(info.filename)))
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")  # a duplicate name is the point of some members
            for name, data in extra or []:
                dst.writestr(name, data)
    return out


# ------------------------------------------------------------- the attacks

class Listener:
    """A socket on 127.0.0.1 that counts connections. The external URLs of the
    attacks point here, never at an outside host; nothing may connect."""

    def __init__(self):
        self.hits = 0
        self.sock = socket.socket()
        self.sock.bind(("127.0.0.1", 0))
        self.sock.listen(8)
        self.sock.settimeout(0.2)
        self.port = self.sock.getsockname()[1]
        self.stop = threading.Event()
        threading.Thread(target=self._loop, daemon=True).start()

    def _loop(self):
        while not self.stop.is_set():
            try:
                conn, _ = self.sock.accept()
            except OSError:
                continue
            self.hits += 1
            conn.close()


@pytest.fixture(scope="module")
def listener():
    lst = Listener()
    yield lst
    lst.stop.set()
    lst.sock.close()


@pytest.fixture(scope="module")
def targets(tmp_path_factory) -> dict[str, str]:
    """Made-up local files an external or parameter entity points at; they
    must never be read."""
    d = tmp_path_factory.mktemp("targets")
    marker = d / "made_up_target.txt"
    marker.write_text("MADE-UP-MARKER-7f3a", encoding="utf-8")
    params = d / "made_up_params.dtd"
    params.write_text('<!ENTITY leak "MADE-UP-MARKER-param">', encoding="utf-8")
    return {"file": marker.resolve().as_uri(), "params": params.resolve().as_uri()}


def _laughs(root: str) -> str:
    ents = '<!ENTITY a "made up made up ">' + "".join(
        f'<!ENTITY {chr(98 + i)} "{("&" + chr(97 + i) + ";") * 10}">' for i in range(9))
    return f"<!DOCTYPE {root} [{ents}]>"


#: attack -> (doctype for a root element, the reference placed in the part)
ATTACKS = {
    "billion_laughs": lambda root, t, port: (_laughs(root), "&j;"),
    "quadratic_blowup": lambda root, t, port: (
        f'<!DOCTYPE {root} [<!ENTITY q "{"Q" * 40000}">]>', "&q;" * 40000),
    "small_internal_entity": lambda root, t, port: (
        f'<!DOCTYPE {root} [<!ENTITY s "made-up words">]>', "&s;"),
    "external_entity_file": lambda root, t, port: (
        f'<!DOCTYPE {root} [<!ENTITY x SYSTEM "{t["file"]}">]>', "&x;"),
    "external_entity_http": lambda root, t, port: (
        f'<!DOCTYPE {root} [<!ENTITY x SYSTEM "http://127.0.0.1:{port}/x.txt">]>', "&x;"),
    "parameter_entity_file": lambda root, t, port: (
        f'<!DOCTYPE {root} [<!ENTITY % p SYSTEM "{t["params"]}"> %p;]>', "&leak;"),
    "parameter_entity_http": lambda root, t, port: (
        f'<!DOCTYPE {root} [<!ENTITY % p SYSTEM "http://127.0.0.1:{port}/p.dtd"> %p;]>', "&leak;"),
    "dtd_without_entities": lambda root, t, port: (f"<!DOCTYPE {root}>", ""),
    "dtd_external_subset": lambda root, t, port: (
        f'<!DOCTYPE {root} SYSTEM "http://127.0.0.1:{port}/made_up.dtd">', ""),
}


def hostile(data: bytes, attack: str, targets: dict, port: int, encoding: str = "utf-8") -> bytes:
    """The part with a DOCTYPE after its XML declaration and the attack's
    reference in its first a:t text (or before its closing tag)."""
    import re

    text = data.decode("utf-8")
    m = re.match(r"\s*<\?xml[^>]*\?>", text)
    head = m.group(0) if m else ""
    rest = text[len(head):]
    root = re.match(r"\s*<([^\s>/!?]+)", rest).group(1)
    doctype, ref = ATTACKS[attack](root, targets, port)
    if ref:
        t = re.search(r"<a:t>([^<]*)</a:t>", rest)
        if t:
            rest = rest[:t.start(1)] + ref + rest[t.end(1):]
        else:
            close = rest.rfind("</")
            rest = rest[:close] + ref + rest[close:]
    xml = head + doctype + rest
    if encoding == "utf-8":
        return xml.encode("utf-8")
    xml = xml.replace('encoding="UTF-8"', 'encoding="UTF-16"', 1)
    return {"utf-16-le-bom": b"\xff\xfe" + xml.encode("utf-16-le"),
            "utf-16-be-bom": b"\xfe\xff" + xml.encode("utf-16-be"),
            "utf-16-le-nobom": xml.encode("utf-16-le")}[encoding]


def _assert_plain(message: str, part: str | None = None) -> None:
    assert "document type declaration" in message, message
    assert "Nothing was changed" in message, message
    if part is not None:
        assert part in message, message
    assert "MADE-UP-MARKER" not in message
    assert not any(raw in message for raw in RAW), message


@pytest.mark.parametrize("kind", sorted(PARTS))
@pytest.mark.parametrize("attack", sorted(ATTACKS))
def test_a_part_that_declares_a_dtd_is_refused_at_open_fast_and_plainly(
        attack, kind, base_deck, targets, listener, tmp_path):
    part = PARTS[kind]
    f = variant(base_deck, tmp_path / "attack.pptx",
                {part: hostile(_part(base_deck, part), attack, targets, listener.port)})
    hits = listener.hits
    t0 = time.perf_counter()
    with pytest.raises(DocumentCorrupt) as err:
        PptxPackage(f)
    assert time.perf_counter() - t0 < 1.0
    _assert_plain(str(err.value), part)
    assert listener.hits == hits


@pytest.mark.parametrize("encoding", ["utf-16-le-bom", "utf-16-be-bom", "utf-16-le-nobom"])
@pytest.mark.parametrize("kind", ["slide", "chart", "comments"])
@pytest.mark.parametrize("attack", ["small_internal_entity", "billion_laughs",
                                    "dtd_without_entities"])
def test_a_utf16_part_that_declares_a_dtd_is_refused(attack, kind, encoding, base_deck,
                                                    targets, listener, tmp_path):
    part = PARTS[kind]
    f = variant(base_deck, tmp_path / "u16.pptx",
                {part: hostile(_part(base_deck, part), attack, targets, listener.port, encoding)})
    with pytest.raises(DocumentCorrupt) as err:
        PptxPackage(f)
    _assert_plain(str(err.value), part)


@pytest.mark.parametrize("attack", sorted(ATTACKS))
def test_the_wire_refuses_every_attack_with_unsupported_content(
        attack, base_deck, targets, listener, tmp_path):
    part = PARTS["slide"]
    f = variant(base_deck, tmp_path / "wire.pptx",
                {part: hostile(_part(base_deck, part), attack, targets, listener.port)})
    hits = listener.hits
    [(payload, is_error, seconds)] = _wire(("get_text", {"file_path": str(f), "live": "off"}))
    assert is_error and payload["error"]["code"] == "UNSUPPORTED_CONTENT", payload
    _assert_plain(payload["error"]["message"], part)
    assert seconds < 5.0
    assert "MADE-UP-MARKER" not in json.dumps(payload) and "made-up words" not in json.dumps(payload)
    assert listener.hits == hits


#: the read that shows each part kind's own content
READS = {
    "slide": ("get_text", {"live": "off"}), "notes": ("get_notes", {"slide": 0}),
    "chart": ("get_chart_data", {"slide": 1}), "comments": ("list_comments", {}),
    "authors": ("list_comments", {}), "presentation": ("get_presentation_info", {}),
    "layout": ("list_elements", {"kind": "layouts"}), "theme": ("get_theme", {}),
    "content_types": ("get_presentation_info", {}),
    "core_properties": ("get_document_properties", {}),
}


@pytest.mark.parametrize("kind", sorted(READS))
def test_each_part_kinds_own_read_and_an_edit_refuse_and_change_nothing(
        kind, base_deck, targets, listener, tmp_path):
    """The small internal entity read as plain text before: the slide title,
    the notes, the chart, the comment all came back with the entity's text,
    and a part no read touched went into the next save with its DTD."""
    part = PARTS[kind]
    f = variant(base_deck, tmp_path / "kind.pptx",
                {part: hostile(_part(base_deck, part), "small_internal_entity", targets,
                               listener.port)})
    before = _md5(f)
    op, extra = READS[kind]
    answers = _wire((op, {"file_path": str(f), **extra}),
                    ("set_placeholder_text", {"file_path": str(f), "slide": 0,
                                              "placeholder": "title", "text": "Edited made-up title",
                                              "live": "off"}))
    for payload, is_error, _seconds in answers:
        assert is_error and payload["error"]["code"] == "UNSUPPORTED_CONTENT", payload
        _assert_plain(payload["error"]["message"], part)
        assert "made-up words" not in json.dumps(payload)
    assert _md5(f) == before


def test_a_declaration_after_a_long_prolog_is_still_refused_at_open(base_deck, tmp_path):
    part = PARTS["slide"]
    data = _part(base_deck, part).decode("utf-8")
    decl_end = data.index("?>") + 2
    pad = "<!-- " + "made-up padding " * 20000 + "--><?ks4 made-up?>"
    xml = (data[:decl_end] + pad + '<!DOCTYPE p:sld [<!ENTITY s "made-up words">]>'
           + data[decl_end:].replace("Made-up title", "&s;"))
    f = variant(base_deck, tmp_path / "pad.pptx", {part: xml.encode("utf-8")})
    with pytest.raises(DocumentCorrupt) as err:
        PptxPackage(f)
    _assert_plain(str(err.value), part)


def test_the_word_doctype_in_a_prolog_comment_is_not_a_declaration(base_deck, tmp_path):
    part = PARTS["slide"]
    data = _part(base_deck, part).decode("utf-8")
    decl_end = data.index("?>") + 2
    xml = data[:decl_end] + "<!-- made-up note: <!DOCTYPE is only text here -->" + data[decl_end:]
    f = variant(base_deck, tmp_path / "pc.pptx", {part: xml.encode("utf-8")})
    out = _call("get_text", file_path=str(f), live="off")
    assert out.get("ok"), out
    assert "Made-up title" in json.dumps(out)


def test_the_word_doctype_in_the_body_is_not_a_declaration(base_deck, tmp_path):
    part = PARTS["slide"]
    data = _part(base_deck, part).decode("utf-8")
    xml = data.replace("Made-up title", "&lt;!DOCTYPE made-up&gt; title").replace(
        "<p:cSld>", "<p:cSld><!-- made-up: <!DOCTYPE is not declared here -->", 1)
    f = variant(base_deck, tmp_path / "body.pptx", {part: xml.encode("utf-8")})
    out = _call("get_text", file_path=str(f), live="off")
    assert out.get("ok"), out
    assert "<!DOCTYPE made-up> title" in json.dumps(out)


def test_a_part_added_after_open_is_parsed_hardened(base_deck, targets, listener, tmp_path):
    f = variant(base_deck, tmp_path / "late.pptx")
    part = PARTS["slide"]
    for attack in ("small_internal_entity", "billion_laughs", "external_entity_file"):
        pkg = PptxPackage(f)
        pkg.set_raw_part(part, hostile(_part(base_deck, part), attack, targets, listener.port))
        with pytest.raises(DocumentCorrupt) as err:
            pkg.tree(part)
        _assert_plain(str(err.value), part)


def test_a_saved_part_that_declares_a_dtd_fails_validation(base_deck, targets, listener, tmp_path):
    f = variant(base_deck, tmp_path / "out.pptx")
    part = PARTS["theme"]
    pkg = PptxPackage(f)
    pkg.set_raw_part(part, hostile(_part(base_deck, part), "dtd_without_entities", targets,
                                   listener.port))
    with pytest.raises(ValidationFailed, match="document type declaration"):
        pkg.save(tmp_path / "out2.pptx")
    assert not (tmp_path / "out2.pptx").exists()


# ------------------------------------------------------------- the parser

def test_the_parser_never_expands_or_loads(base_deck, targets, listener):
    from kitchensink4ppt.core import xmlsafe

    p = xmlsafe.parser()
    for opt in ("resolve_entities", "no_network", "load_dtd", "dtd_validation",
                "attribute_defaults"):
        with pytest.raises(ValueError):
            xmlsafe.parser(**{opt: True})
    calls = []

    class Spy(etree.Resolver):
        def resolve(self, url, pubid, context):
            calls.append(url)

    p.resolvers.add(Spy())
    slide = _part(base_deck, PARTS["slide"])
    for attack in ("external_entity_file", "external_entity_http", "dtd_external_subset"):
        # Below the refusal: the raw parser itself reads nothing.
        try:
            root = etree.fromstring(hostile(slide, attack, targets, listener.port), parser=p)
        except etree.XMLSyntaxError:
            continue
        assert "MADE-UP-MARKER" not in "".join(root.itertext())
    assert calls == []
    assert listener.hits == 0


def test_fromstring_refuses_every_attack_in_memory(base_deck, targets, listener):
    from kitchensink4ppt.core import xmlsafe

    slide = _part(base_deck, PARTS["slide"])
    for attack in ATTACKS:
        t0 = time.perf_counter()
        with pytest.raises(DocumentCorrupt) as err:
            xmlsafe.fromstring(hostile(slide, attack, targets, listener.port),
                               doc="made-up.pptx", part="ppt/slides/slide1.xml")
        assert time.perf_counter() - t0 < 0.07
        _assert_plain(str(err.value), "ppt/slides/slide1.xml")


#: The XML entry points of lxml and the standard library. Outside the
#: hardened module every parse goes through it.
_RAW_CALLS = {("etree", "fromstring"), ("etree", "XML"), ("etree", "parse"),
              ("etree", "iterparse"), ("etree", "XMLParser"), ("etree", "XMLPullParser"),
              ("etree", "HTML"), ("etree", "HTMLParser"), ("lxml_etree", "fromstring"),
              ("lxml_etree", "XMLParser"), ("_et", "fromstring"), ("_et", "XML"),
              ("_et", "parse"), ("_et", "iterparse"), ("_et", "XMLParser"),
              ("_stdlib_et", "fromstring"), ("_stdlib_et", "XML"), ("_stdlib_et", "parse"),
              ("_stdlib_et", "iterparse"), ("_stdlib_et", "XMLParser"),
              ("ElementTree", "fromstring"), ("ElementTree", "parse"),
              ("ElementTree", "iterparse"), ("minidom", "parseString"), ("minidom", "parse"),
              ("sax", "parse"), ("sax", "parseString"), ("objectify", "fromstring"),
              ("expat", "ParserCreate")}


def test_no_module_parses_xml_outside_the_hardened_parser():
    roots = [SRC, ROOT / "scripts"]
    offenders = []
    for root in roots:
        for py in sorted(root.rglob("*.py")):
            if py.name == "xmlsafe.py" and py.parent.name == "core":
                continue
            tree = ast.parse(py.read_text(encoding="utf-8"))
            for node in ast.walk(tree):
                if (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                        and isinstance(node.func.value, ast.Name)
                        and (node.func.value.id, node.func.attr) in _RAW_CALLS):
                    offenders.append(f"{py.relative_to(ROOT)}:{node.lineno} "
                                     f"{node.func.value.id}.{node.func.attr}")
                if isinstance(node, ast.ImportFrom):
                    mod = node.module or ""
                    names = {a.name for a in node.names}
                    if (mod in ("lxml.etree", "xml.etree.ElementTree", "xml.etree.cElementTree",
                                "xml.dom.minidom", "xml.sax", "lxml.objectify")
                            and names & {"fromstring", "XML", "parse", "iterparse", "XMLParser",
                                         "parseString", "make_parser", "XMLPullParser"}):
                        offenders.append(f"{py.relative_to(ROOT)}:{node.lineno} from {mod}")
    assert offenders == []


# ------------------------------------------------- the outside XML libraries

SVG_NS = "http://www.w3.org/2000/svg"


def _svg(doctype: str, label: str, root_attrs: str = f'xmlns="{SVG_NS}"') -> str:
    return (f'<?xml version="1.0" encoding="UTF-8"?>{doctype}'
            f'<svg {root_attrs} width="100" height="60" viewBox="0 0 100 60">'
            '<rect x="5" y="5" width="40" height="30" fill="#336699"/>'
            f'<text x="50" y="30">{label}</text></svg>')


@pytest.fixture
def svgelements_spy(monkeypatch):
    import svgelements

    seen: list[str] = []
    original = svgelements.SVG.parse

    def spy(source, *a, **k):
        seen.append(source.getvalue() if hasattr(source, "getvalue") else str(source))
        return original(source, *a, **k)

    monkeypatch.setattr(svgelements.SVG, "parse", staticmethod(spy))
    return seen


SVG_ATTACKS = {
    "small_internal_entity": lambda t, port: _svg('<!DOCTYPE svg [<!ENTITY s "made-up words">]>',
                                                  "&s;"),
    "billion_laughs": lambda t, port: _svg(_laughs("svg"), "&j;"),
    "external_entity_file": lambda t, port: _svg(
        f'<!DOCTYPE svg [<!ENTITY x SYSTEM "{t["file"]}">]>', "&x;"),
    "external_entity_http": lambda t, port: _svg(
        f'<!DOCTYPE svg [<!ENTITY x SYSTEM "http://127.0.0.1:{port}/x">]>', "&x;"),
    "parameter_entity_http": lambda t, port: _svg(
        f'<!DOCTYPE svg [<!ENTITY % p SYSTEM "http://127.0.0.1:{port}/p"> %p;]>', "plain"),
    "namespace_entities": lambda t, port: _svg(
        f'<!DOCTYPE svg [<!ENTITY ns_svg "{SVG_NS}">]>', "Made-up label", 'xmlns="&ns_svg;"'),
}


@pytest.mark.parametrize("attack", sorted(SVG_ATTACKS))
def test_svg_with_its_own_declarations_is_refused_before_any_parser(
        attack, base_deck, targets, listener, svgelements_spy, tmp_path):
    f = variant(base_deck, tmp_path / "svg.pptx")
    before = _md5(f)
    hits = listener.hits
    out = _call("svg_to_shapes", file_path=str(f), svg=SVG_ATTACKS[attack](targets, listener.port),
                x=1, y=1, w=3, h=2, slide=0)
    assert out["ok"] is False and out["error"]["code"] == "BAD_PARAMS", out
    _assert_plain(out["error"]["message"])
    assert svgelements_spy == []
    assert listener.hits == hits
    assert _md5(f) == before


def test_a_plain_svg_doctype_is_dropped_and_the_drawing_still_converts(
        base_deck, listener, svgelements_spy, tmp_path):
    """The SVG 1.1 header (a public DTD name, no declarations of its own) is
    common in exported drawings and declares nothing: no parser here loads
    the DTD it names. It is dropped before any parser reads the drawing."""
    f = variant(base_deck, tmp_path / "svg11.pptx")
    hits = listener.hits
    svg = _svg('<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" '
               f'"http://127.0.0.1:{listener.port}/svg11.dtd">', "Made-up label")
    out = _call("svg_to_shapes", file_path=str(f), svg=svg, x=1, y=1, w=3, h=2, slide=0)
    assert out.get("ok"), out
    assert len(svgelements_spy) == 1 and "<!DOCTYPE" not in svgelements_spy[0]
    assert listener.hits == hits
    assert "Made-up label" in json.dumps(_call("get_text", file_path=str(f), live="off"))


def test_mathml2omml_never_sees_a_dtd(base_deck, monkeypatch, tmp_path):
    """mathml2omml parses with the standard library's SAX reader, which
    expands a DTD's internal entities. Whatever the LaTeX converter hands
    over, a DTD never reaches it."""
    import latex2mathml.converter as l2m
    import mathml2omml

    seen: list[str] = []
    original = mathml2omml.convert
    monkeypatch.setattr(mathml2omml, "convert", lambda xml, *a, **k: (seen.append(xml),
                                                                     original(xml, *a, **k))[1])
    monkeypatch.setattr(l2m, "convert", lambda latex, **k: (
        '<?xml version="1.0"?><!DOCTYPE math [<!ENTITY s "made-up words">]>'
        '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>&s;</mi></math>'))
    f = variant(base_deck, tmp_path / "eq.pptx")
    before = _md5(f)
    out = _call("insert_equation", file_path=str(f), slide=0, latex="x", x=1, y=1)
    assert out["ok"] is False, out
    _assert_plain(out["error"]["message"])
    assert seen == []
    assert _md5(f) == before


def test_latex_that_spells_a_doctype_never_declares_one(base_deck, tmp_path):
    """The LaTeX converter hands its text content over unescaped, so markup
    typed into \\text{} arrives as markup; a declaration there sits inside the
    math element, where XML allows none, and expands nothing."""
    f = variant(base_deck, tmp_path / "eqtext.pptx")
    out = _call("insert_equation", file_path=str(f), slide=0,
                latex=r"\text{<!DOCTYPE m [<!ENTITY s 'made-up words'>]>&s;}x", x=1, y=1)
    assert "made-up words" not in json.dumps(out) or out["ok"] is False
    if out.get("ok"):
        assert "made-up words" not in json.dumps(_call("list_equations", file_path=str(f)))


# ---------------------------------------------------------------- zip safety

def _zeros(zf: zipfile.ZipFile, name: str, mib: int) -> None:
    info = zipfile.ZipInfo(name, (2026, 1, 1, 0, 0, 0))
    info.compress_type = zipfile.ZIP_DEFLATED
    chunk = bytes(MiB)
    with zf.open(info, "w", force_zip64=True) as fh:
        for _ in range(mib):
            fh.write(chunk)


def bomb(base: Path, out: Path, *, members: int = 1, mib: int = 1024) -> Path:
    """`base` plus `members` parts of `mib` MiB of zeros each, deflated (about
    1,000 to 1 in the file)."""
    variant(base, out)
    with zipfile.ZipFile(out, "a", zipfile.ZIP_DEFLATED, allowZip64=True) as zf:
        for i in range(members):
            _zeros(zf, f"ppt/media/made_up_zeros_{i:04d}.bin", mib)
    return out


def _refused_cheaply(f: Path) -> DocumentCorrupt:
    tracemalloc.start()
    t0 = time.perf_counter()
    try:
        with pytest.raises(DocumentCorrupt) as err:
            PptxPackage(f)
        seconds = time.perf_counter() - t0
        _now, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    assert seconds < 1.0
    assert peak < 32 * MiB, peak
    message = str(err.value)
    assert "Nothing was changed" in message and "exhaust memory" in message, message
    return err.value


def test_a_decompression_bomb_part_is_refused_before_it_is_inflated(base_deck, tmp_path):
    """1 GiB of zeros in about 1 MB. Reading it asked for twice its size in
    memory; it is refused from the sizes the package declares, first."""
    f = bomb(base_deck, tmp_path / "bomb.pptx", mib=1024)
    assert f.stat().st_size < 4 * MiB
    err = _refused_cheaply(f)
    assert "ppt/media/made_up_zeros_0000.bin" in str(err)
    [(payload, is_error, seconds)] = _wire(("get_text", {"file_path": str(f), "live": "off"}))
    assert is_error and payload["error"]["code"] == "UNSUPPORTED_CONTENT", payload
    assert seconds < 5.0


def test_many_small_bomb_parts_are_refused_together(base_deck, tmp_path):
    f = bomb(base_deck, tmp_path / "many.pptx", members=300, mib=1)
    _refused_cheaply(f)


def test_a_bomb_in_the_charts_embedded_workbook_is_refused_before_it_is_inflated(
        base_deck, tmp_path):
    name = next(n for n in zipfile.ZipFile(base_deck).namelist()
                if n.startswith("ppt/embeddings/") and n.endswith(".xlsx"))
    nested = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(_part(base_deck, name))) as src, \
            zipfile.ZipFile(nested, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as dst:
        for info in src.infolist():
            dst.writestr(info.filename, src.read(info.filename))
        _zeros(dst, "xl/media/made_up_zeros.bin", 1024)
    f = variant(base_deck, tmp_path / "nested.pptx", {name: nested.getvalue()})
    before = _md5(f)
    tracemalloc.start()
    t0 = time.perf_counter()
    try:
        out = _call("update_chart_data", file_path=str(f), slide=1,
                    series=[{"name": "Made-up series", "values": [4, 5, 6]}],
                    categories=["Alpha", "Beta", "Gamma"])
        seconds = time.perf_counter() - t0
        _now, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    assert out["ok"] is False and out["error"]["code"] == "UNSUPPORTED_CONTENT", out
    assert "exhaust memory" in out["error"]["message"]
    assert seconds < 2.0 and peak < 64 * MiB, (seconds, peak)
    assert _md5(f) == before


def test_a_large_part_that_is_not_a_bomb_still_opens(base_deck, tmp_path):
    """The check is about what the package claims against what it stores:
    a part that stores what it holds reads as before, whatever its size."""
    from kitchensink4ppt.core import zipsafe

    stored = [zipfile.ZipInfo("ppt/media/made_up_video.mp4")]
    stored[0].file_size = stored[0].compress_size = 900 * MiB
    zipsafe.check_members(stored, doc="made-up.pptx")  # a large stored part: fine
    packed = [zipfile.ZipInfo("ppt/media/made_up_zeros.bin")]
    packed[0].file_size, packed[0].compress_size = 900 * MiB, 900 * 1024
    with pytest.raises(DocumentCorrupt, match="exhaust memory"):
        zipsafe.check_members(packed, doc="made-up.pptx")
    f = bomb(base_deck, tmp_path / "small.pptx", mib=16)
    assert len(PptxPackage(f).raw_part("ppt/media/made_up_zeros_0000.bin")) == 16 * MiB


def test_member_names_never_become_file_paths(base_deck, tmp_path):
    """Members named to climb out of a folder or to land on a drive are kept
    as names inside the package; no operation writes a file by them."""
    names = ["../made_up_escape_1.xml", "C:/made_up_escape_3.xml", "/made_up_escape_4.xml",
             "ppt/../../made_up_escape_5.xml", "ppt/media/../../../made_up_escape_7.bin"]
    work = tmp_path / "deep" / "er"
    work.mkdir(parents=True)
    f = variant(base_deck, work / "names.pptx",
                extra=[(n, f"<made-up n='{i}'/>".encode()) for i, n in enumerate(names)])
    raw = work / "names_raw.pptx"
    with zipfile.ZipFile(f) as src, zipfile.ZipFile(raw, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            dst.writestr(info, src.read(info))
        odd = zipfile.ZipInfo("placeholder")
        odd.filename = odd.orig_filename = "..\\made_up_escape_2.xml"
        dst.writestr(odd, b"<made-up/>")
    for deck in (f, raw):
        assert _call("get_text", file_path=str(deck), live="off").get("ok")
        assert _call("set_placeholder_text", file_path=str(deck), slide=0, placeholder="title",
                     text="Edited made-up title", live="off").get("ok")
        assert _call("copy_presentation", file_path=str(deck),
                     dest_path=str(work / f"copy_{deck.stem}.pptx")).get("ok")
        assert _call("split_deck", file_path=str(deck), output_dir=str(work / f"split_{deck.stem}"),
                     by="ranges", ranges=[{"start": 0, "end": 0, "name": "../../made_up_escape_s"},
                                          {"start": 1, "end": 1, "name": "C:\\made_up_escape_t"}]
                     ).get("ok")
        assert _call("create_snapshot", file_path=str(deck)).get("ok")
    with zipfile.ZipFile(f) as zf:
        assert set(names) <= set(zf.namelist())  # kept verbatim, as names
    places = {tmp_path, tmp_path.parent, work, work.parent, Path.cwd(), Path(tmp_path.anchor)}
    escaped = [p for d in places for p in d.glob("made_up_escape*")]
    escaped += [p for p in tmp_path.rglob("made_up_escape*")]
    assert escaped == []


@pytest.mark.parametrize("order", ["second_last", "second_first"])
def test_a_package_that_names_a_part_twice_is_refused(order, base_deck, tmp_path):
    """The reader kept whichever copy came last; PowerPoint repairs the file
    and may show the other one. Neither is read."""
    part = PARTS["slide"]
    original = _part(base_deck, part)
    second = original.replace(b"Made-up title", b"Made-up second copy")
    f = tmp_path / "twice.pptx"
    with zipfile.ZipFile(base_deck) as src, zipfile.ZipFile(f, "w", zipfile.ZIP_DEFLATED) as dst:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            for info in src.infolist():
                if info.filename == part and order == "second_first":
                    dst.writestr(part, second)
                dst.writestr(info.filename, src.read(info.filename))
            if order == "second_last":
                dst.writestr(part, second)
    before = _md5(f)
    with pytest.raises(DocumentCorrupt) as err:
        PptxPackage(f)
    message = str(err.value)
    assert part in message and "twice" in message and "Nothing was changed" in message
    [(payload, is_error, _s)] = _wire(("get_text", {"file_path": str(f), "live": "off"}))
    assert is_error and payload["error"]["code"] == "UNSUPPORTED_CONTENT", payload
    assert "Made-up second copy" not in json.dumps(payload)
    assert _md5(f) == before


def test_part_names_that_differ_only_in_capitals_are_refused(base_deck, tmp_path):
    """Package part names ignore case, so ppt/slides/Slide1.xml beside
    ppt/slides/slide1.xml names the same part twice."""
    data = _part(base_deck, PARTS["slide"]).replace(b"Made-up title", b"Made-up case twin")
    f = variant(base_deck, tmp_path / "case.pptx", extra=[("ppt/slides/Slide1.xml", data)])
    with pytest.raises(DocumentCorrupt) as err:
        PptxPackage(f)
    message = str(err.value)
    assert "ppt/slides/Slide1.xml" in message and "ppt/slides/slide1.xml" in message
    assert "Nothing was changed" in message
