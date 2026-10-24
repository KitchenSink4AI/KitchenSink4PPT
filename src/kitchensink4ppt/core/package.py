"""PptxPackage: safe load/save layer for .pptx files.

Design guarantees:
- Parts that were never touched are written back byte-for-byte identical.
  A deck has dozens of parts (slides, layouts, masters, themes, notesSlides,
  media); an edit to slide 7 must not re-serialize slide 3.
- Each entry's ZIP compression method is recorded at load and reused at save.
  A deck whose writer stored its JPEG/MP4 media uncompressed gets it stored
  back, not re-deflated: that is both faithful to the source package and, on
  media-heavy decks, most of the save time (deflating 39 MB of already
  compressed media bought 0.8% and cost 900 ms of a 1,520 ms save).
- Saves are atomic: temp file -> validation -> one swap (ReplaceFileW on
  Windows, which keeps the deck's identity; see core.safesave.promote). The
  original is never left half-written, and validation failure leaves it
  untouched.
- Auto-backup on by default: before each in-place save the current content is
  rotated into stable slots (prev.pptx / anchor.pptx) inside a hidden
  .ks4p-backups/ folder next to the presentation (see core.safesave).
- Files locked by PowerPoint are detected before any work happens.

PPTX structural rules baked in here (they differ from .docx and getting them
wrong is the classic corruption):
- Slide ORDER lives in presentation.xml's p:sldIdLst, resolved through the
  presentation rels. ZIP entry order and slideN.xml numbering mean nothing.
- EVERY part has its own .rels file; adding a part means adding its
  relationship AND its [Content_Types].xml override.
- Shape ids (p:cNvPr/@id) are unique PER SLIDE, never globally.
- p:sldId ids live in [256, 2147483647].
"""

from __future__ import annotations

import contextlib
import contextvars
import hashlib
import io
import os
import posixpath
import re
import stat
import time
import zipfile
from pathlib import Path
from urllib.parse import unquote

from lxml import etree

from .errors import (
    DocumentBusy,
    DocumentCorrupt,
    DocumentLocked,
    DocumentNotFound,
    DocumentEditProtected,
    DocumentReadOnly,
    DocumentProtected,
    OutputRefused,
    PptMcpError,
    SlideIdsExhausted,
    UnsupportedPackage,
    ValidationFailed,
)
from . import xmlsafe, zipsafe
from .sandbox import check_path
from .swapwait import deck_not_found, exists_after_swap


class DamagedPart(DocumentCorrupt, etree.XMLSyntaxError):
    """A part of the deck is not well-formed XML: damage in the file, not
    the caller's mistake, so it is reported as such and names the part. It
    is still an XMLSyntaxError, so code that reads around a part it cannot
    parse keeps doing so."""

    def __init__(self, deck: str, part: str, cause: etree.XMLSyntaxError):
        self.message = (
            f"{part} in {deck} is damaged: it is not well-formed XML ({cause}), "
            "so it cannot be read; nothing was changed. PowerPoint repairs what "
            "it can when it opens a deck: open this one in PowerPoint, save it, "
            "and try again."
        )
        etree.XMLSyntaxError.__init__(
            self, self.message, cause.code, cause.lineno, cause.offset, None
        )
        self.part = part

    def __str__(self) -> str:
        return self.message


NSMAP = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
    # The SVG layer of a dual blip (PowerPoint 2016+ SVG pictures).
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
}

PRESENTATION_PART = "ppt/presentation.xml"
#: The highest shape id PowerPoint reads (xsd:unsignedInt).
MAX_SHAPE_ID = 4294967295
#: The root relationship to the presentation, Transitional and Strict.
_RT_OFFICE_DOCUMENT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
_RT_OFFICE_DOCUMENT_STRICT = (
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument")
_STRICT_PML_NS = "http://purl.oclc.org/ooxml/presentationml/main"

#: Pauses before looking again when a read lands inside another program's
#: save (about 0.85 s in all); see PptxPackage._load.
_READ_RACE_DELAYS = (0.05, 0.1, 0.2, 0.5)

#: Set for the length of one call: whether a package opened during it is
#: opened to READ (dispatch sets it for the read-only operations), and the
#: decks it read from their saved file while PowerPoint had them open.
_READING: contextvars.ContextVar[bool] = contextvars.ContextVar(
    "ks4p_reading", default=False)
_HELD_READS: contextvars.ContextVar[list | None] = contextvars.ContextVar(
    "ks4p_held_reads", default=None)


@contextlib.contextmanager
def reading_call(read_only: bool):
    """For one call: open decks to read when the operation only reads, and
    collect every deck read from its saved file while PowerPoint had it
    open (open_deck_note turns them into the answer's note)."""
    held: list[dict] = []
    reading_token = _READING.set(bool(read_only))
    held_token = _HELD_READS.set(held)
    try:
        yield held
    finally:
        _READING.reset(reading_token)
        _HELD_READS.reset(held_token)


def _note_held(path, reason: str | None = None) -> None:
    held = _HELD_READS.get()
    if held is None:
        return
    key = os.path.normcase(os.path.abspath(str(path)))
    for record in held:
        if record["path"] == key:
            if reason and not record.get("reason"):
                record["reason"] = reason
            return
    held.append({"path": key, "name": Path(path).name, "reason": reason})


def note_live_unreachable(path, reason: str) -> None:
    """A read with a live route tried the open copy and could not reach it;
    the saved file is read instead, and the note says why."""
    _note_held(path, " ".join(str(reason).split())[:240])


def open_deck_note(held: list[dict]) -> str | None:
    """The answer's note for decks read from their saved file while
    PowerPoint had them open, or None."""
    from .messages import OPEN_DECK_LIVE_FAILED, OPEN_DECK_READ_NOTE

    parts = []
    for record in held:
        text = OPEN_DECK_READ_NOTE.format(name=record["name"])
        if record.get("reason"):
            text += OPEN_DECK_LIVE_FAILED.format(reason=record["reason"])
        parts.append(text)
    return " ".join(parts) or None


def _owner_file(path: Path) -> Path:
    """The owner file PowerPoint writes beside a deck it has open."""
    return path.with_name("~$" + path.name[-153:])


def powerpoint_holds(path) -> bool:
    """True when PowerPoint has the deck open: its owner file is there and
    the deck refuses write access. Only looks; changes nothing."""
    p = Path(path)
    if not _owner_file(p).exists():
        return False
    try:
        with open(p, "r+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


def has_read_only_attribute(path: str | os.PathLike) -> bool:
    """Whether the file is explicitly marked read-only.

    On Windows the owner-write mode bit mirrors FILE_ATTRIBUTE_READONLY; on
    POSIX it is the ordinary owner-write bit.  Inspecting the bit, rather
    than relying on ``os.access``, keeps the result stable when tests or the
    server run with elevated permissions.
    """
    try:
        return not bool(Path(path).stat().st_mode & stat.S_IWRITE)
    except OSError:
        return False


def read_only_error(path: str | os.PathLike) -> DocumentReadOnly:
    """Build the user-facing refusal for an in-place file mutation."""
    target = os.path.abspath(os.fspath(path))
    exc = DocumentReadOnly(
        f"{target} is read-only: its file read-only attribute is set. "
        "Reading is allowed, but file-mode changes cannot be saved."
    )
    exc.hint = (
        "Clear the file's read-only attribute and retry, or use "
        "copy_presentation to create a writable copy."
    )
    return exc


def make_file_writable(path: str | os.PathLike) -> None:
    """Clear a copied file's read-only bit while preserving other mode bits."""
    p = Path(path)
    p.chmod(p.stat().st_mode | stat.S_IWRITE)


#: Relationship type URIs used by the package layer.
RT_SLIDE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
)
RT_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)

#: Content types used by the package layer.
CT_RELATIONSHIPS = "application/vnd.openxmlformats-package.relationships+xml"
CT_SLIDE = (
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
)

#: p:sldId ids must live in this range (ECMA-376; python-pptx enforces the
#: same bounds).
SLIDE_ID_MIN = 256
SLIDE_ID_MAX = 2147483647
#: The refusal when a new slide would need an id past SLIDE_ID_MAX. Slide ids
#: are given out above the highest in use and never reused: a reused id
#: would make a view anchor or comment of a deleted slide name the new one.
SLIDE_IDS_EXHAUSTED = (
    "This deck's slide ids have reached 2147483647, the highest the file format allows, "
    "so a new slide cannot be numbered after them; nothing was changed. New slide ids are "
    "always numbered after the highest id in use. To add slides, copy this deck's slides "
    "into a new deck with merge_decks, where they are numbered afresh."
)

#: PowerPoint's practical drawing coordinate ceiling (32-bit signed EMU).
#: Payload validation refuses any xfrm/path coordinate past it: PowerPoint
#: cannot open such a deck, so payload_valid must never bless one.
MAX_DRAWING_EMU = 2147483647

#: The SmartArt cached-drawing namespace, whose dsp:txBody has the same
#: content model as p:txBody and a:txBody (kept here so the payload check
#: does not import ops/).
_DSP_NS = "http://schemas.microsoft.com/office/drawing/2008/diagram"

#: The chart-drawing namespace, whose cdr:txBody is the same text model.
_CDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"

#: Every text-body element that must carry at least one a:p. All of these
#: are CT_TextBody in the schema (bodyPr, lstStyle, then one or more a:p),
#: so one rule covers them: p:txBody on slides, layouts, masters and
#: notes; a:txBody in table cells; dsp:txBody in a diagram's cached
#: drawing; c:rich and c:txPr in a chart (titles, axis and data labels);
#: cdr:txBody in a chart's own drawing layer.
_TXBODY_TAGS = (
    f"{{{NSMAP['p']}}}txBody",
    f"{{{NSMAP['a']}}}txBody",
    f"{{{_DSP_NS}}}txBody",
    f"{{{NSMAP['c']}}}rich",
    f"{{{NSMAP['c']}}}txPr",
    f"{{{_CDR_NS}}}txBody",
)

#: Markup Compatibility. A body's paragraphs may be supplied through an
#: mc:AlternateContent wrapper rather than sitting as direct children, and
#: PowerPoint selects one branch at load time.
_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_MC_ALTERNATE = f"{{{_MC_NS}}}AlternateContent"
_MC_BRANCHES = (f"{{{_MC_NS}}}Choice", f"{{{_MC_NS}}}Fallback")

#: How deep AlternateContent may nest before the check stops looking and
#: accepts. Refusing a legitimate save is the worst thing this gate can
#: do, so pathological nesting gets the benefit of the doubt.
_MC_MAX_DEPTH = 8


def _supplies_paragraph(container: etree._Element, depth: int = 0) -> bool:
    """Does this text body have a paragraph, however it is supplied?

    A direct a:p is the ordinary case. The other legal shape is
    mc:AlternateContent: PowerPoint selects ONE branch, so the wrapper is
    a guarantee only when EVERY branch it offers supplies a paragraph. A
    blind descendant search would let one valid Fallback hide an empty
    Choice, which is the corruption this gate exists to catch.
    """
    if depth > _MC_MAX_DEPTH:
        return True
    if container.find(qn("a:p")) is not None:
        return True
    for alt in container.findall(_MC_ALTERNATE):
        branches = [b for tag in _MC_BRANCHES for b in alt.findall(tag)]
        if branches and all(
            _supplies_paragraph(b, depth + 1) for b in branches
        ):
            return True
    return False

#: Where to look for the name and id of the shape a text body belongs to,
#: relative to the body's PARENT (p:sp, dsp:sp) or its grandparent.
_CNVPR_PATHS = (
    # p:sp/p:nvSpPr/p:cNvPr on a slide, layout, master or notes slide.
    (f"{{{NSMAP['p']}}}nvSpPr", f"{{{NSMAP['p']}}}cNvPr"),
    # dsp:sp/dsp:nvSpPr/a:cNvPr in a cached SmartArt drawing.
    (f"{{{_DSP_NS}}}nvSpPr", f"{{{NSMAP['a']}}}cNvPr"),
)


def _txbody_owner(body: etree._Element) -> str:
    """The shape a text body hangs off, named as well as the XML allows:
    'shape 7 "Agenda Body"' for a p:sp, 'a table cell' for a:txBody, and
    the bare parent tag when there is no cNvPr to read."""
    parent = body.getparent()
    if parent is None:
        return "an unattached element"
    local = etree.QName(parent).localname
    if local == "tc":
        return "a table cell"
    for nv_tag, cnvpr_tag in _CNVPR_PATHS:
        nv = parent.find(nv_tag)
        if nv is None:
            continue
        cnvpr = nv.find(cnvpr_tag)
        if cnvpr is None:
            continue
        name = cnvpr.get("name")
        shape_id = cnvpr.get("id")
        named = f' "{name}"' if name else ""
        return f"shape {shape_id}{named}" if shape_id else f"a shape{named}"
    return f"a {local} element"


#: Schema-fixed order of p:presentation children. Inserting sldIdLst after
#: sldSz corrupts the file, so new children go in by rank, never appended.
_PRESENTATION_ORDER = (
    "p:sldMasterIdLst",
    "p:notesMasterIdLst",
    "p:handoutMasterIdLst",
    "p:sldIdLst",
    "p:sldSz",
    "p:notesSz",
    "p:smartTags",
    "p:embeddedFontLst",
    "p:custShowLst",
    "p:photoAlbum",
    "p:custDataLst",
    "p:kinsoku",
    "p:defaultTextStyle",
    "p:modifyVerifier",
    "p:extLst",
)


def qn(tag: str) -> str:
    """'p:sp' -> '{http://...}sp' (Clark notation)."""
    prefix, local = tag.split(":")
    return f"{{{NSMAP[prefix]}}}{local}"


def _is_ole_encrypted(head: bytes) -> bool:
    # OLE compound file magic = password-protected or legacy .ppt, not a ZIP.
    return head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")


def refuse_ole_encrypted(path: str | os.PathLike) -> None:
    """Refuse an OLE/encrypted presentation before an Office process sees it.

    Password-protected OOXML presentations are OLE compound files rather
    than ZIP packages.  PowerPoint raises a modal password prompt when one
    reaches ``Presentations.Open``, so every path-based COM/export caller
    uses this cheap package preflight before selecting or starting an engine.
    Legacy binary ``.ppt`` files have the same magic and remain unsupported.
    """
    p = Path(path)
    with open(p, "rb") as fh:
        head = fh.read(8)
    if _is_ole_encrypted(head):
        raise DocumentProtected(
            f"{p.name} is password-protected or a legacy binary .ppt; "
            "remove the password or convert to .pptx in PowerPoint first."
        )


#: Media extensions whose payload is ALREADY compressed. Deflating these is
#: the save path's largest cost and buys almost nothing: on the 40 MB corpus
#: deck, deflating 39.3 MB of JPEG/PNG/GIF/MP4 took 980 ms and shrank it by
#: 0.9%. Formats that genuinely do compress (BMP, WAV, TIFF, EMF/WMF, SVG)
#: are deliberately absent.
_PRECOMPRESSED_MEDIA_EXT = frozenset(
    {
        ".jpg", ".jpeg", ".jpe", ".jfif", ".png", ".gif", ".webp",
        ".jp2", ".heic", ".heif", ".avif",
        ".mp4", ".m4v", ".m4a", ".mp3", ".mpg", ".mpeg", ".mov",
        ".avi", ".wmv", ".wma", ".webm", ".ogg", ".oga", ".ogv", ".aac",
        ".zip", ".xlsx", ".docx", ".pptx",
    }
)


def _is_precompressed_media(name: str) -> bool:
    """True for a media part whose bytes are already a compressed format."""
    if not name.startswith("ppt/media/"):
        return False
    return posixpath.splitext(name)[1].lower() in _PRECOMPRESSED_MEDIA_EXT


def _default_compress_type(name: str) -> int:
    """Compression method for a part the source package did not carry (an
    image or video this session inserted). Already-compressed media is
    STORED, which is what PowerPoint's own writer does with it; everything
    else is XML and deflates five to ten times over."""
    return zipfile.ZIP_STORED if _is_precompressed_media(name) else zipfile.ZIP_DEFLATED


def rels_name(part: str) -> str:
    """'ppt/slides/slide1.xml' -> 'ppt/slides/_rels/slide1.xml.rels'."""
    d, fname = posixpath.split(part)
    return posixpath.join(d, "_rels", fname + ".rels") if d else posixpath.join("_rels", fname + ".rels")


def rels_source(rels_part: str) -> str:
    """'ppt/slides/_rels/slide1.xml.rels' -> 'ppt/slides/slide1.xml'.
    The package-root rels '_rels/.rels' maps to '' (the package itself)."""
    d, fname = posixpath.split(rels_part)
    parent = posixpath.dirname(d)
    source = fname[: -len(".rels")]
    if source == ".":  # '_rels/.rels'
        return ""
    return posixpath.join(parent, source) if parent else source


def resolve_target(source_part: str, target: str) -> str:
    """Resolve a relationship Target (part-relative, or absolute with a
    leading '/') against its source part's directory to a package part name."""
    if target.startswith("/"):
        return posixpath.normpath(target[1:])
    base = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(base, target) if base else target)


def spelled_target(source_part: str, target: str, names) -> str | None:
    """The text a relationship Target should be read as, when as written it
    names no part of the package but another spelling PowerPoint reads the
    same does: backslashes as slashes, a #fragment dropped, percent codes
    decoded (slide%31.xml is slide1.xml). None when the target already names
    a part, or when no spelling does. `names` is any container of part names."""
    if resolve_target(source_part, target) in names:
        return None
    plain = target.replace("\\", "/").split("#", 1)[0]
    if not plain:
        return None
    for candidate in (plain, unquote(plain)):
        if resolve_target(source_part, candidate) in names:
            return candidate
    return None


def seat_ahead_of_extlst(container: etree._Element, element: etree._Element) -> etree._Element:
    """Put a shape just added to the end of a shape tree (or group) ahead of
    the container's own p:extLst. The schema ends a p:spTree / p:grpSp with
    its extLst, and PowerPoint ignores a shape written after it: the file
    opens clean and the new shape is not there. Returns `element`."""
    ext = container.find("{%s}extLst" % NSMAP["p"])
    if ext is not None and element.getparent() is container:
        ext.addprevious(element)
    return element


_P14_CNVPR = "{http://schemas.microsoft.com/office/powerpoint/2010/main}cNvPr"


def known_problem_keys(path) -> frozenset:
    """The keys of the structural problems the deck at `path` already has
    (PptxPackage._payload_problems), so that a save or a restore refuses
    only the problems it would add. An empty set when the file is missing
    or cannot be read as a package: then every problem is refused."""
    try:
        data = Path(path).read_bytes()
        return frozenset(
            key for key, _m in PptxPackage._payload_problems(data, doc=Path(path).name))
    except Exception:  # noqa: BLE001 - unreadable means nothing is tolerated
        return frozenset()


def kept_problems_note(messages: list[str]) -> str:
    """One sentence for the answer's warnings: the problems a save or a
    restore kept because the deck already had them."""
    from .messages import KEPT_PROBLEM_ONE, KEPT_PROBLEMS_MANY

    first = messages[0].rstrip(".")
    if len(messages) == 1:
        return KEPT_PROBLEM_ONE.format(first=first)
    return KEPT_PROBLEMS_MANY.format(count=len(messages), first=first)


class PptxPackage:
    """One .pptx opened for inspection or editing.

    reading=True opens it to read: a deck PowerPoint has open is read from
    its saved file (PowerPoint shares read access) instead of refused, and
    held_by_powerpoint says so. None takes the call's setting
    (reading_call); an edit always opens with reading=False."""

    def __init__(self, path: str | os.PathLike, *, reading: bool | None = None):
        check_path(path, "open presentation")  # no-op unless KS4P_ALLOWED_ROOTS is set
        self.path = Path(path)
        self.reading = _READING.get() if reading is None else bool(reading)
        #: True when PowerPoint had the deck open and this package read the
        #: file as last saved.
        self.held_by_powerpoint = False
        self._raw: dict[str, bytes] = {}  # part name -> original bytes
        self._order: list[str] = []  # original entry order, preserved on save
        self._trees: dict[str, etree._ElementTree] = {}  # parsed parts
        self._dirty: set[str] = set()  # parts whose tree must be re-serialized
        #: part name -> the ZIP compression method the source package used
        #: (zipfile.ZIP_STORED / ZIP_DEFLATED / ...), replayed on save.
        self._compress: dict[str, int] = {}
        #: Warnings from the last save, for the caller's answer (a deck whose
        #: own permissions could not be kept, for one).
        self.save_warnings: list[str] = []
        self._load()

    # ---------- loading ----------

    def _comes_back(self) -> bool:
        """A save puts the new deck in place with ReplaceFileW, which moves
        the old file aside for an instant: the path is missing, then back.
        True once it is back; False at once when no save has the old file
        aside (core/swapwait.py, which only looks)."""
        return exists_after_swap(self.path)

    def _load(self) -> None:
        # os.path.exists and isdir, not pathlib's: a file in the middle of a
        # swap can make stat fail, and pathlib re-raises that on Python 3.12.
        if not os.path.exists(self.path) and not self._comes_back():
            raise deck_not_found(self.path, f"No file at {self.path}")
        if os.path.isdir(self.path):
            # Opening a directory raises PermissionError on Windows, which
            # the lock check would misread as a PowerPoint lock.
            raise PptMcpError(
                f"{self.path} is a directory, not a .pptx file; pass the "
                "full path to the presentation file"
            )
        # A save by another server replaces the file in place. A read that
        # lands inside that replace gets a bare PermissionError, from the lock
        # probe just before the open or from the open itself. Nothing typed
        # the first: the caller was told the server failed, or that the deck
        # was open in PowerPoint when no PowerPoint was involved. The replace
        # takes milliseconds: look again, probe included, and if it is still
        # being written, say so (DocumentBusy, never "open in PowerPoint").
        for delay in (*_READ_RACE_DELAYS, None):
            try:
                self._check_lock()
                self._read_package()
                break
            except (PermissionError, FileNotFoundError) as gone:
                self._raw.clear()
                self._order.clear()
                self._compress.clear()
                if isinstance(gone, FileNotFoundError):
                    # Missing mid-read: a save has the old deck aside for an
                    # instant, or the deck really went.
                    if not self._comes_back():
                        raise deck_not_found(
                            self.path, f"No file at {self.path}") from None
                if delay is None:
                    if self.reading and _owner_file(self.path).exists():
                        raise DocumentBusy(
                            f"{self.path.name} is open in PowerPoint, which is "
                            "not letting other programs read it right now (it "
                            "may be saving it), so it could not be read."
                        ) from None
                    raise DocumentBusy(
                        f"{self.path.name} is being written by another program "
                        "right now, so it could not be read."
                    ) from None
                if isinstance(gone, PermissionError):
                    time.sleep(delay)  # a deck that came back is read again at once
        self._check_main_part()

    def _check_main_part(self) -> None:
        """Refuse a package this server would misread: the root relationship
        names the presentation, and it must be ppt/presentation.xml in the
        Transitional namespaces PowerPoint saves by default. A Strict Open
        XML deck has the same part names and other namespaces, so it read as
        a deck with no slides; a package whose main part has another name was
        read from a ppt/presentation.xml PowerPoint never looks at, or
        refused as not a presentation."""
        from .messages import MAIN_PART_REFUSAL, STRICT_DECK_REFUSAL

        main, strict = None, False
        rels = self._raw.get("_rels/.rels")
        if rels is not None:
            try:
                root = xmlsafe.fromstring(rels, doc=self.path.name, part="_rels/.rels")
            except etree.XMLSyntaxError:
                root = None
            for rel in root if root is not None else ():
                kind = rel.get("Type", "")
                if kind in (_RT_OFFICE_DOCUMENT, _RT_OFFICE_DOCUMENT_STRICT):
                    main = resolve_target("", rel.get("Target", ""))
                    strict = kind == _RT_OFFICE_DOCUMENT_STRICT
                    break
        if not strict and PRESENTATION_PART in self._raw:
            strict = xmlsafe.root_namespace(
                self._raw[PRESENTATION_PART], doc=self.path.name,
                part=PRESENTATION_PART) == _STRICT_PML_NS
        if strict:
            raise UnsupportedPackage(STRICT_DECK_REFUSAL.format(name=self.path.name))
        if main is not None and main != PRESENTATION_PART:
            raise UnsupportedPackage(
                MAIN_PART_REFUSAL.format(name=self.path.name, part=main))
        if PRESENTATION_PART not in self._raw:
            raise DocumentCorrupt(
                f"{self.path.name} has no {PRESENTATION_PART}; not a "
                "PowerPoint presentation."
            )

    def _read_package(self) -> None:
        # Read 8 bytes, not the file. read_bytes()[:8] slurped the whole deck
        # (23.6 ms and a 40 MB transient allocation on the largest corpus
        # deck) to look at a magic number.
        refuse_ole_encrypted(self.path)
        try:
            with zipfile.ZipFile(self.path) as zf:
                # From the directory alone, before anything is inflated
                # (testzip inflates every part): a part named twice, and
                # parts that would inflate far past what they store.
                zipsafe.check_members(zf.infolist(), doc=self.path.name)
                bad = zf.testzip()
                if bad is not None:
                    raise DocumentCorrupt(
                        f"{self.path.name}: corrupt ZIP entry '{bad}'."
                    )
                for info in zf.infolist():
                    self._raw[info.filename] = zf.read(info)
                    self._order.append(info.filename)
                    self._compress[info.filename] = info.compress_type
        except zipfile.BadZipFile as exc:
            raise DocumentCorrupt(
                f"{self.path.name} is not a valid .pptx (bad ZIP): {exc}"
            ) from exc
        # A document type declaration in any XML part refuses the deck here,
        # whichever operation asked and whichever parts it would have read.
        for name in self._order:
            if name.lower().endswith((".xml", ".rels", ".vml")):
                xmlsafe.refuse_dtd(self._raw[name], doc=self.path.name, part=name)

    def _check_lock(self) -> None:
        """Distinguish a read-only attribute from a process-held lock, and
        PowerPoint's hold from another program's save.

        A read-only deck is still a valid input to every package reader.  The
        old unconditional ``r+b`` probe misdiagnosed it as open in
        PowerPoint and prevented even text extraction.  An exclusive holder
        still blocks the read probe.

        PowerPoint writes an owner file (``~$name``) beside a deck it has
        open, so a failed probe with that file present is PowerPoint's hold.
        An edit refuses it with DocumentLocked, at once, which is what sends
        a live='auto' edit to the open copy. A read does not: PowerPoint
        denies other programs write access to a deck it has open but shares
        read access, so a read opens the file as last saved and says so
        (held_by_powerpoint, and the answer's open_deck_note). Without the
        owner file the same failure is most likely another server's save
        replacing the deck, a matter of milliseconds, so the PermissionError
        goes to ``_load``, which looks again after the read-race pauses and
        answers DocumentBusy if the deck is still not free; that protection
        holds for reads and edits alike.
        """
        owner_file = _owner_file(self.path)
        mode = "rb" if has_read_only_attribute(self.path) else "r+b"
        try:
            with open(self.path, mode):
                pass
        except PermissionError:
            if not owner_file.exists():
                raise
            if self.reading:
                # A refusal of read access too (PowerPoint in the middle of
                # its own save) raises here and is looked at again in _load.
                with open(self.path, "rb"):
                    pass
                self.held_by_powerpoint = True
                _note_held(self.path)
                return
            raise DocumentLocked(
                f"{self.path.name} is open in PowerPoint or locked by another "
                "process (PowerPoint owner file present). Close the file in "
                "PowerPoint and retry."
            ) from None

    # ---------- part access ----------

    def has_part(self, name: str) -> bool:
        return name in self._raw

    def part_names(self) -> list[str]:
        return list(self._order)

    def tree(self, name: str = PRESENTATION_PART) -> etree._ElementTree:
        """Parsed XML for a part. Parsing is cached; call mark_dirty() after edits.
        A part that is not well-formed XML raises DamagedPart, naming it."""
        if name not in self._trees:
            if name not in self._raw:
                raise KeyError(f"part not in package: {name}")
            try:
                root = xmlsafe.fromstring(self._raw[name], doc=self.path.name, part=name,
                                          remove_blank_text=False, strip_cdata=False)
            except etree.XMLSyntaxError as exc:
                raise DamagedPart(self.path.name, name, exc) from exc
            if name.endswith(".rels"):
                self._read_targets_as_powerpoint_does(name, root)
            self._trees[name] = etree.ElementTree(root)
        return self._trees[name]

    def _read_targets_as_powerpoint_does(self, rels_part: str, root: etree._Element) -> None:
        """PowerPoint opens a deck whose relationship Target is spelled with a
        percent code, a backslash or a #fragment (slides/slide%31.xml). The
        reads here took the text as the part name and failed, and a write
        built a stray rels part for a slide that does not exist. The
        in-memory tree is given the spelling that names the real part; the
        file's bytes are written back unchanged unless this rels part is
        itself edited. Only the presentation's own relationships (its slides,
        masters and the like) are read this way: the targets of other parts
        keep the spelling they have, which a copy of a slide carries over as
        written."""
        source = rels_source(rels_part)
        if source != PRESENTATION_PART:
            return
        for rel in root:
            if not isinstance(rel.tag, str) or rel.get("TargetMode") == "External":
                continue
            spelling = spelled_target(source, rel.get("Target", ""), self._raw)
            if spelling is not None:
                rel.set("Target", spelling)

    def root(self, name: str = PRESENTATION_PART) -> etree._Element:
        return self.tree(name).getroot()

    def presentation(self) -> etree._Element:
        """Root element of ppt/presentation.xml (the spine)."""
        return self.root(PRESENTATION_PART)

    def assert_modifiable(self) -> None:
        """Refuse filesystem or package protection before an edit."""
        if has_read_only_attribute(self.path):
            raise read_only_error(self.path)
        if self.presentation().find(qn("p:modifyVerifier")) is not None:
            raise DocumentEditProtected(
                f"{self.path.name} has a password to modify, so this server will not edit it by "
                "file path; nothing was changed, and reading still works. A copy made "
                "with copy_presentation keeps the password."
            )

    def raw_part(self, name: str) -> bytes:
        return self._raw[name]

    def set_raw_part(self, name: str, data: bytes) -> None:
        """Add or replace a part with raw bytes (media, new XML parts)."""
        if name not in self._raw:
            self._order.append(name)
        self._raw[name] = data
        self._trees.pop(name, None)
        self._dirty.discard(name)  # raw bytes are authoritative now
        self._compress.setdefault(name, _default_compress_type(name))
        if (name.endswith(".rels") and name != "[Content_Types].xml"
                and "[Content_Types].xml" in self._raw
                and self.content_type(name) is None):
            # A deck that declares its relationship parts one by one (no
            # Default for "rels") gives a new rels part no type at all, and
            # PowerPoint repairs the file.
            self.add_content_type_override(name, CT_RELATIONSHIPS)

    def mark_dirty(self, name: str = PRESENTATION_PART) -> None:
        if name not in self._trees:
            raise RuntimeError(f"mark_dirty before tree() for {name}")
        self._dirty.add(name)

    # ---------- relationships ----------

    def rels_for(self, part: str, *, create: bool = False) -> etree._ElementTree:
        """The .rels tree for a part ('' = the package root rels). With
        create=True a missing rels part is created empty; edits to the
        returned tree still require mark_dirty(rels_name(part))."""
        name = rels_name(part) if part else "_rels/.rels"
        if name not in self._raw:
            if not create:
                raise KeyError(f"no rels part for {part or '<package root>'}")
            empty = (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                b'<Relationships xmlns='
                b'"http://schemas.openxmlformats.org/package/2006/relationships"/>'
            )
            self.set_raw_part(name, empty)
        return self.tree(name)

    def next_rid(self, part: str) -> str:
        """Next free rId in a part's rels namespace. rIds are per source part
        and never assumed contiguous: max numeric suffix + 1."""
        try:
            rels = self.rels_for(part)
        except KeyError:
            return "rId1"
        highest = 0
        for rel in rels.getroot():
            m = re.fullmatch(r"rId(\d+)", rel.get("Id", ""))
            if m:
                highest = max(highest, int(m.group(1)))
        return f"rId{highest + 1}"

    def add_relationship(
        self,
        source_part: str,
        rel_type: str,
        target: str,
        *,
        external: bool = False,
    ) -> str:
        """Add a relationship from source_part; returns the new rId. `target`
        is part-relative for internal rels (as PowerPoint writes them) and a
        URI for external ones."""
        rid = self.next_rid(source_part)
        rels = self.rels_for(source_part, create=True)
        rel = etree.SubElement(rels.getroot(), f"{{{NSMAP['rel']}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", rel_type)
        rel.set("Target", target)
        if external:
            rel.set("TargetMode", "External")
        self.mark_dirty(rels_name(source_part) if source_part else "_rels/.rels")
        return rid

    def relationship_target(self, source_part: str, rid: str) -> str:
        """Resolve one rId of source_part to a package part name."""
        rels = self.rels_for(source_part)
        for rel in rels.getroot():
            if rel.get("Id") == rid:
                if rel.get("TargetMode") == "External":
                    raise PptMcpError(
                        f"{rid} in {source_part} is an external relationship"
                    )
                return resolve_target(source_part, rel.get("Target", ""))
        raise KeyError(f"{source_part} has no relationship {rid}")

    # ---------- content types ----------

    def _content_type_tables(self) -> tuple[dict[str, str], dict[str, str]]:
        """({part name lowercased: Override type}, {extension lowercased:
        Default type}) from [Content_Types].xml; OPC compares both without
        case."""
        overrides: dict[str, str] = {}
        defaults: dict[str, str] = {}
        if "[Content_Types].xml" not in self._raw:
            return overrides, defaults
        for node in self.root("[Content_Types].xml"):
            local = etree.QName(node).localname if isinstance(node.tag, str) else ""
            if local == "Override":
                name = (node.get("PartName") or "").lstrip("/").lower()
                overrides[name] = node.get("ContentType", "")
            elif local == "Default":
                defaults[(node.get("Extension") or "").lower()] = node.get("ContentType", "")
        return overrides, defaults

    @staticmethod
    def _content_type_tables_of(root) -> tuple[dict[str, str], dict[str, str]]:
        """_content_type_tables for a parsed [Content_Types].xml root."""
        overrides: dict[str, str] = {}
        defaults: dict[str, str] = {}
        for node in root if root is not None else ():
            local = etree.QName(node).localname if isinstance(node.tag, str) else ""
            if local == "Override":
                overrides[(node.get("PartName") or "").lstrip("/").lower()] = node.get("ContentType", "")
            elif local == "Default":
                defaults[(node.get("Extension") or "").lower()] = node.get("ContentType", "")
        return overrides, defaults

    @staticmethod
    def _type_from(tables, part: str) -> str | None:
        overrides, defaults = tables
        found = overrides.get(part.lower())
        if found is not None:
            return found
        base = posixpath.basename(part)
        if "." not in base:
            return None
        return defaults.get(base.rsplit(".", 1)[-1].lower())

    def content_type(self, part: str) -> str | None:
        """The content type that covers one part: its Override, else the
        Default for its extension; None when nothing covers it."""
        return self._type_from(self._content_type_tables(), part)

    def content_type_map(self) -> dict[str, str]:
        """{part name: content type} for every covered part."""
        tables = self._content_type_tables()
        out: dict[str, str] = {}
        for name in self._order:
            found = self._type_from(tables, name)
            if found is not None:
                out[name] = found
        return out

    def add_content_type_override(self, part: str, content_type: str) -> None:
        """Ensure [Content_Types].xml carries an Override for `part`.
        Additive only: existing Overrides for other parts are never dropped."""
        ct_root = self.root("[Content_Types].xml")
        part_name = "/" + part
        for node in ct_root.findall(qn("ct:Override")):
            if node.get("PartName") == part_name:
                if node.get("ContentType") != content_type:
                    node.set("ContentType", content_type)
                    self.mark_dirty("[Content_Types].xml")
                return
        override = etree.SubElement(ct_root, qn("ct:Override"))
        override.set("PartName", part_name)
        override.set("ContentType", content_type)
        self.mark_dirty("[Content_Types].xml")

    def add_part_with_content_type(
        self, name: str, data: bytes, content_type: str
    ) -> None:
        """Add a part and its content-type override in one call. Forgetting
        the override is the #1 cause of PowerPoint repair prompts."""
        self.set_raw_part(name, data)
        self.add_content_type_override(name, content_type)

    # ---------- slides ----------

    def slide_target(self, rid: str) -> str:
        """The slide part one presentation relationship names. A slide
        relationship that names no part of the file is damage in the file:
        it is refused by name, not left to fail later as a missing key."""
        part = self.relationship_target(PRESENTATION_PART, rid)
        if part not in self._raw:
            raise DocumentCorrupt(
                f"{self.path.name} lists a slide ({rid}) whose relationship names "
                f"{part}, and the file has no such part, so its slides cannot be "
                "read; nothing was changed. PowerPoint repairs what it can when it "
                "opens a deck: open this one in PowerPoint, save it, and try again."
            )
        return part

    def slide_parts(self) -> list[str]:
        """Slide part names in PRESENTATION ORDER: p:sldIdLst entries resolved
        through the presentation rels. NEVER ZIP order or filename numbering;
        slide2.xml can be the fifth slide after reordering."""
        pres = self.presentation()
        sld_id_lst = pres.find(qn("p:sldIdLst"))
        if sld_id_lst is None:
            return []
        parts: list[str] = []
        for sld in sld_id_lst.findall(qn("p:sldId")):
            rid = sld.get(qn("r:id"))
            if rid is None:
                raise DocumentCorrupt("p:sldId entry without r:id")
            part = self.slide_target(rid)
            parts.append(part)
        return parts

    def next_shape_id(self, slide_part: str) -> int:
        """Next free shape id for one slide/layout/master/notesSlide part.
        Shape ids (p:cNvPr/@id) are unique PER PART, never package-global."""
        highest = 1  # the root spTree's own cNvPr is conventionally id=1
        used: set[int] = set()
        # Ink content parts carry their id on p14:cNvPr, often with no
        # picture fallback repeating it, so both spellings count.
        for node in self.root(slide_part).iter(qn("p:cNvPr"), _P14_CNVPR):
            try:
                value = int(node.get("id", "0"))
            except ValueError:
                continue
            used.add(value)
            highest = max(highest, value)
        if highest < MAX_SHAPE_ID:
            return highest + 1
        # A shape id is an unsigned 32-bit number: past it PowerPoint repairs
        # the deck. Take the lowest id nothing uses.
        candidate = 2
        while candidate in used:
            candidate += 1
        if candidate > MAX_SHAPE_ID:
            raise PptMcpError(
                f"{slide_part} has no free shape id left; nothing was changed")
        return candidate

    def _next_slide_id(self, sld_id_lst: etree._Element | None) -> int:
        highest = SLIDE_ID_MIN - 1
        if sld_id_lst is not None:
            for sld in sld_id_lst.findall(qn("p:sldId")):
                try:
                    highest = max(highest, int(sld.get("id", "0")))
                except ValueError:
                    continue
        if highest >= SLIDE_ID_MAX:
            raise SlideIdsExhausted(SLIDE_IDS_EXHAUSTED)
        return highest + 1

    def _insert_presentation_child(self, element: etree._Element) -> None:
        """Insert a child of p:presentation at its schema-fixed position."""
        pres = self.presentation()
        order = [qn(t) for t in _PRESENTATION_ORDER]
        try:
            rank = order.index(element.tag)
        except ValueError:
            raise PptMcpError(f"not a known p:presentation child: {element.tag}")
        for child in pres:
            if child.tag in order and order.index(child.tag) > rank:
                child.addprevious(element)
                return
        pres.append(element)

    def add_slide_part(
        self, slide_xml: bytes, *, layout_part: str
    ) -> dict:
        """Add a new slide to the deck ATOMICALLY in package terms: the part,
        its rels (with the layout relationship), the content-type override,
        the presentation relationship, and the p:sldId entry all land
        together, or an exception leaves the in-memory package unchanged.

        Adding a slide touches four registries; doing it half-way is the
        classic pptx corruption, which is why this is ONE method.
        """
        # Validate and precompute everything BEFORE mutating any structure.
        if layout_part not in self._raw:
            raise PptMcpError(f"layout part not in package: {layout_part}")
        try:
            xmlsafe.fromstring(slide_xml, doc=self.path.name, part="the new slide")
        except etree.XMLSyntaxError as exc:
            raise ValidationFailed(f"new slide XML is not well-formed: {exc}") from exc
        highest = 0
        for name in self._raw:
            m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name)
            if m:
                highest = max(highest, int(m.group(1)))
        new_part = f"ppt/slides/slide{highest + 1}.xml"
        pres = self.presentation()
        sld_id_lst = pres.find(qn("p:sldIdLst"))
        slide_id = self._next_slide_id(sld_id_lst)

        # Mutations begin; each step below is append-only and cannot fail
        # for structural reasons (all inputs were validated above).
        self.add_part_with_content_type(new_part, slide_xml, CT_SLIDE)
        layout_target = posixpath.relpath(layout_part, posixpath.dirname(new_part))
        layout_target = layout_target.replace(os.sep, "/")
        self.add_relationship(new_part, RT_SLIDE_LAYOUT, layout_target)
        pres_target = posixpath.relpath(new_part, posixpath.dirname(PRESENTATION_PART))
        pres_target = pres_target.replace(os.sep, "/")
        rid = self.add_relationship(PRESENTATION_PART, RT_SLIDE, pres_target)
        if sld_id_lst is None:
            sld_id_lst = etree.SubElement(pres, qn("p:sldIdLst"))
            pres.remove(sld_id_lst)
            self._insert_presentation_child(sld_id_lst)
        sld = etree.SubElement(sld_id_lst, qn("p:sldId"))
        sld.set("id", str(slide_id))
        sld.set(qn("r:id"), rid)
        self.mark_dirty(PRESENTATION_PART)
        return {"part": new_part, "slide_id": slide_id, "rid": rid}

    # ---------- saving ----------

    def _serialize(self, name: str) -> bytes:
        tree = self._trees[name]
        return etree.tostring(
            tree, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    def _compress_type_for(self, name: str) -> int:
        """How this part gets written into the output ZIP.

        Two rules, in order. A part the source package STORED is stored
        back: that is faithful to the writer that produced the deck, and
        before this the server silently re-deflated 26 of image_deck's 28
        media entries on every save. A part whose bytes are an
        already-compressed media format is stored too, whatever the source
        did, because deflating it is the single largest cost in the save
        path and buys under 1% (980 ms for 0.34 MB on the 40 MB deck).
        Everything else keeps the method the source used, which for XML is
        always deflate.

        Storing media that the source deflated grows the file by well under
        one percent and does not touch the part's bytes; the byte-for-byte
        writeback guarantee is about part CONTENT, and content is identical
        either way."""
        recorded = self._compress.get(name)
        if recorded == zipfile.ZIP_STORED:
            return zipfile.ZIP_STORED
        if _is_precompressed_media(name):
            return zipfile.ZIP_STORED
        if recorded is None:
            return _default_compress_type(name)
        return recorded

    def save(self, dest: str | os.PathLike | None = None, *, do_backup: bool = True) -> Path:
        """Atomic save. dest=None means save in place (with slot backup rotation
        by default: prev.pptx/anchor.pptx under .ks4p-backups/, see core.safesave).
        do_backup=False skips the rotation only; the atomic validated save always
        applies."""
        from . import safesave

        dest_path = Path(dest) if dest else self.path
        check_path(dest_path, "save presentation")
        in_place = dest_path.resolve() == self.path.resolve()
        if dest_path.exists() and has_read_only_attribute(dest_path):
            raise read_only_error(dest_path)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in self._order:
                data = self._serialize(name) if name in self._dirty else self._raw[name]
                zf.writestr(name, data, compress_type=self._compress_type_for(name))
        payload = buf.getvalue()

        # Validate the payload before touching the destination. A problem the
        # deck on disk already has is kept and reported, never refused.
        kept = self._check_payload(payload)

        # Linux filenames cap at 255 BYTES (not chars); a long CJK name plus
        # our suffix can exceed it, so hash-shorten while staying in the same
        # directory (atomic os.replace needs the same filesystem).
        tmp_name = dest_path.name + ".ppt-mcp-tmp"
        if len(tmp_name.encode("utf-8", "surrogatepass")) > 240:
            digest = hashlib.sha1(
                dest_path.name.encode("utf-8", "surrogatepass")
            ).hexdigest()[:16]
            tmp_name = f".{digest}.ppt-mcp-tmp"
        tmp = dest_path.with_name(tmp_name)
        ticket = None
        try:
            tmp.write_bytes(payload)
            safesave.claim_publish("the presentation save")
            # Stage the current (pre-mutation) content for the backup slots
            # by hard link. promote lands it on the slots only after the swap
            # went through, so a refused save leaves the slots as they were.
            if in_place and do_backup and dest_path.exists():
                ticket = safesave.prepare_rotation(dest_path)
            self.save_warnings = list(
                safesave.promote(tmp, dest_path, ticket=ticket) or [])
            if kept:
                self.save_warnings.append(kept_problems_note(kept))
        except BaseException as exc:
            if ticket is not None:
                ticket.abort()
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass
            if isinstance(exc, OSError):
                refusal = safesave.os_write_refusal(dest_path, exc)
                if refusal is not None:
                    raise refusal from exc
            raise
        # After a successful save the written bytes are the new baseline.
        if in_place:
            for name in self._dirty:
                self._raw[name] = self._serialize(name)
            self._dirty.clear()
        return dest_path

    @staticmethod
    def _validate_payload(
        payload: bytes, *, tolerate=frozenset(), doc: str = "the file"
    ) -> list[str]:
        """Structural sanity before the destination is touched: valid ZIP,
        required roots present, every XML part well-formed, every internal
        relationship target resolvable, every sldIdLst entry resolvable.
        Dangling rels are the pptx equivalent of Word's repair prompt.

        Raises ValidationFailed for the first problem whose key is not in
        `tolerate` (the problems the deck already had: see
        known_problem_keys); returns the messages of the tolerated ones,
        which the caller reports. With nothing tolerated (validate, a backup
        restored over a missing deck) every problem is refused, as before.
        A package built to exhaust memory, or naming a part twice, is
        refused (DocumentCorrupt, named by `doc`) before any part is read."""
        problems = PptxPackage._payload_problems(payload, doc=doc)
        for key, message in problems:
            if key not in tolerate:
                raise ValidationFailed(message)
        return [message for _key, message in problems]

    @staticmethod
    def _payload_problems(payload: bytes, *, doc: str = "the file") -> list[tuple[tuple, str]]:
        """Every structural problem of a payload as (key, message), in the
        order the checks run. The key names the problem by kind, part and
        id, so the same problem in the deck on disk and in a payload built
        from it compares equal. A payload that is not a ZIP raises at once,
        and so does one whose directory names a part twice or declares parts
        built to exhaust memory (core/zipsafe.py, as opening a deck checks
        it): a backup, a restore source and a file given to validate are
        read here without being opened first."""
        problems: list[tuple[tuple, str]] = []

        def found(key: tuple, message: str) -> None:
            problems.append((key, message))

        try:
            with zipfile.ZipFile(io.BytesIO(payload)) as zf:
                zipsafe.check_members(zf.infolist(), doc=doc)
                names = set(zf.namelist())
                if PRESENTATION_PART not in names:
                    found(("root", PRESENTATION_PART), f"output lost {PRESENTATION_PART}")
                    return problems
                if "[Content_Types].xml" not in names:
                    found(("root", "[Content_Types].xml"), "output lost [Content_Types].xml")
                    return problems
                roots: dict[str, etree._Element] = {}
                for name in sorted(names):
                    if name.endswith((".xml", ".rels")):
                        try:
                            roots[name] = xmlsafe.fromstring(zf.read(name), part=name)
                        except DocumentCorrupt:
                            raise OutputRefused(
                                f"output part {name} contains a document type declaration "
                                "(<!DOCTYPE ...>), which a PowerPoint presentation never "
                                "has; nothing was saved"
                            ) from None
                        except etree.XMLSyntaxError as exc:
                            found(("xml", name),
                                  f"output part {name} is not well-formed XML: {exc}")

                # No drawing coordinate may exceed PowerPoint's 2^31-1 EMU
                # ceiling: past it the deck fails to open (repair path), so
                # a payload carrying one is invalid even though the XML is
                # schema-clean.
                coord_tags = {
                    qn("a:off"), qn("a:ext"), qn("a:chOff"), qn("a:chExt"),
                    qn("a:pt"),
                }
                for name, root in roots.items():
                    if not name.startswith("ppt/") or name.endswith(".rels"):
                        continue
                    for el in root.iter(*coord_tags):
                        for attr in ("x", "y", "cx", "cy"):
                            raw = el.get(attr)
                            if raw is None:
                                continue
                            try:
                                value = int(raw)
                            except ValueError:
                                continue
                            if abs(value) > MAX_DRAWING_EMU:
                                found(
                                    ("coordinate", name, etree.QName(el).localname, attr, value),
                                    f"out-of-range coordinate in {name}: "
                                    f"{etree.QName(el).localname} {attr}="
                                    f"{value} EMU exceeds PowerPoint's "
                                    f"{MAX_DRAWING_EMU} EMU limit; the deck "
                                    "would not open cleanly",
                                )

                # Every text body needs at least one a:p. A p:txBody with
                # none is schema-clean and relationship-clean, so it sailed
                # through every check here, and PowerPoint then refused the
                # whole deck naming nothing: about 40 minutes of manual
                # bisection on a 45-slide build. The content model says a body
                # HAS paragraphs, so a payload carrying an empty one is
                # invalid.
                for name, root in roots.items():
                    if not name.startswith("ppt/") or name.endswith(".rels"):
                        continue
                    for body in root.iter(*_TXBODY_TAGS):
                        if _supplies_paragraph(body):
                            continue
                        owner = _txbody_owner(body)
                        found(
                            ("empty_text_body", name, owner),
                            f"empty text body in {name}: "
                            f"{etree.QName(body).localname} on "
                            f"{owner} has no a:p paragraph, "
                            "directly or through an AlternateContent "
                            "branch; PowerPoint refuses to open a deck "
                            "with one. Leave one empty paragraph instead.",
                        )

                # Every part needs a content type, by Override or by the
                # Default for its extension. A part with none is schema-clean
                # and relationship-clean, and PowerPoint repairs the deck.
                tables = PptxPackage._content_type_tables_of(roots.get("[Content_Types].xml"))
                for name in sorted(names if "[Content_Types].xml" in roots else ()):
                    if name == "[Content_Types].xml" or name.endswith("/"):
                        continue
                    if PptxPackage._type_from(tables, name) is None:
                        found(
                            ("no_content_type", name),
                            f"output part {name} has no content type: "
                            "[Content_Types].xml has no Override for it and no Default "
                            "for its extension, so PowerPoint would repair the deck",
                        )

                # Drawing ids are unsigned 32-bit numbers; a shape tree ends
                # with its own p:extLst; a timing target names a shape the
                # slide has. Each of these passes the checks above and makes
                # PowerPoint repair the deck (or ignore the edit).
                for name, root in roots.items():
                    if not name.startswith("ppt/") or name.endswith(".rels"):
                        continue
                    for el in root.iter(qn("p:cNvPr"), _P14_CNVPR):
                        raw = el.get("id")
                        if raw is not None and raw.isdigit() and int(raw) > MAX_SHAPE_ID:
                            found(
                                ("shape_id_range", name, int(raw)),
                                f"shape id {raw} in {name} is past {MAX_SHAPE_ID}, the "
                                "highest id PowerPoint reads; the deck would be repaired",
                            )
                    for tree_el in root.iter(qn("p:spTree"), qn("p:grpSp")):
                        kids = [c for c in tree_el if isinstance(c.tag, str)]
                        tags = [c.tag for c in kids]
                        if qn("p:extLst") in tags:
                            after = len(kids) - 1 - tags.index(qn("p:extLst"))
                            if after:
                                owner = tree_el.find(".//" + qn("p:cNvPr"))
                                found(
                                    ("shape_after_extlst", name,
                                     owner.get("id") if owner is not None else None, after),
                                    f"{after} shape(s) in {name} are written after the "
                                    "shape tree's p:extLst, where PowerPoint ignores them",
                                )
                    for authors in root.iter(qn("p:cmAuthorLst")):
                        seen: set[tuple[str, str]] = set()
                        for author in authors.findall(qn("p:cmAuthor")):
                            presence = next(
                                author.iter("{%s}presenceInfo" % NSMAP["p15"]), None)
                            if presence is None:
                                continue
                            who = (author.get("name", ""), presence.get("userId", ""))
                            if who in seen:
                                found(
                                    ("comment_authors_alike", name, who[0], who[1]),
                                    f"two comment authors in {name} share the name "
                                    f"{who[0]!r} and the same user id; PowerPoint would "
                                    "repair the deck",
                                )
                            seen.add(who)
                    if name.startswith("ppt/slides/slide") and root.tag == qn("p:sld"):
                        shape_ids = {
                            el.get("id") for el in root.iter(qn("p:cNvPr"), _P14_CNVPR)}
                        for tgt_el in root.iter(qn("p:tgtEl")):
                            tgt = next((c for c in tgt_el if isinstance(c.tag, str)), None)
                            if tgt is None:
                                found(
                                    ("timing_empty_target", name),
                                    f"an animation target in {name} names nothing "
                                    "(empty p:tgtEl); PowerPoint would repair the deck",
                                )
                            elif tgt.tag == qn("p:spTgt") and tgt.get("spid") not in shape_ids:
                                found(
                                    ("timing_target", name, tgt.get("spid")),
                                    f"an animation in {name} targets shape {tgt.get('spid')}, "
                                    "which is not on the slide; PowerPoint would repair the deck",
                                )

                # Every internal relationship target must exist in the package.
                # Keep the complete id set (including external relationships)
                # for the inverse check below: an XML r:id/r:embed/r:link is
                # just as broken when its .rels entry was dropped.
                rel_tag = f"{{{NSMAP['rel']}}}Relationship"
                pres_rid_targets: dict[str, str] = {}
                rel_ids_by_source: dict[str, set[str]] = {}
                for name, root in roots.items():
                    if not name.endswith(".rels"):
                        continue
                    source = rels_source(name)
                    for rel in root.iter(rel_tag):
                        rid = rel.get("Id", "")
                        rel_ids_by_source.setdefault(source, set()).add(rid)
                        if rel.get("TargetMode") == "External":
                            continue
                        target = rel.get("Target", "")
                        resolved = resolve_target(source, target)
                        if resolved not in names:
                            # Spellings PowerPoint reads as the part: escaped
                            # characters (%20), backslashes, a #fragment.
                            spelling = spelled_target(source, target, names)
                            if spelling is not None:
                                resolved = resolve_target(source, spelling)
                            elif unquote(resolved) in names:
                                resolved = unquote(resolved)
                            else:
                                found(
                                    ("dangling_relationship", name, rid, target),
                                    f"dangling relationship in {name}: "
                                    f"{rel.get('Id')} -> {target} (no part "
                                    f"{resolved})",
                                )
                                continue
                        if source == PRESENTATION_PART:
                            pres_rid_targets[rid] = resolved

                relationship_ns = NSMAP["r"]
                for name, root in roots.items():
                    if name.endswith(".rels"):
                        continue
                    known_ids = rel_ids_by_source.get(name, set())
                    for element in root.iter():
                        for attribute, rid in element.attrib.items():
                            qname = etree.QName(attribute)
                            if qname.namespace != relationship_ns:
                                continue
                            # PowerPoint's ground-truthed media click action
                            # uses a:hlinkClick r:id="" with
                            # action="ppaction://media". The empty value is a
                            # sentinel, not a relationship lookup.
                            if rid == "":
                                continue
                            if rid not in known_ids:
                                found(
                                    ("dangling_relationship_id", name, qname.localname, rid),
                                    f"dangling relationship id in {name}: "
                                    f"{qname.localname}={rid!r} has no entry in "
                                    f"{rels_name(name)}",
                                )

                # Every slide in sldIdLst must resolve through those rels.
                pres = roots.get(PRESENTATION_PART)
                sld_id_lst = pres.find(qn("p:sldIdLst")) if pres is not None else None
                if sld_id_lst is not None:
                    for sld in sld_id_lst.findall(qn("p:sldId")):
                        rid = sld.get(qn("r:id"))
                        if rid not in pres_rid_targets:
                            found(
                                ("slide_id", sld.get("id"), rid),
                                f"sldId {sld.get('id')} references {rid}, "
                                "which is not a presentation relationship",
                            )
        except zipfile.BadZipFile as exc:  # pragma: no cover
            raise ValidationFailed(f"output is not a valid ZIP: {exc}") from exc
        return problems

    def _check_payload(self, payload: bytes) -> list[str]:
        """The save's check: refuse any problem this change would add, and
        keep (and report) a problem the deck on disk already has. A deck
        that came in with an unused relationship to a missing picture, or a
        part nothing references, could otherwise never be changed at all."""
        problems = self._payload_problems(payload, doc=self.path.name)
        if not problems:
            return []
        return self._validate_payload(
            payload, tolerate=known_problem_keys(self.path), doc=self.path.name)

    # ---------- Package-level primitives for slide CRUD ----------

    def part_bytes(self, name: str) -> bytes:
        """Current effective bytes of a part: the serialized tree when the
        part is dirty, the original raw bytes otherwise. Use this instead of
        raw_part() when copying parts that may have in-memory edits."""
        if name in self._dirty:
            return self._serialize(name)
        return self._raw[name]

    def next_partname(self, template: str) -> str:
        """Collision-safe partname allocation. `template` carries one '{}'
        (e.g. 'ppt/slides/slide{}.xml'); scans the ACTUAL partnames in the
        package and returns template.format(max numeric suffix + 1). Never
        computed from a count: a deck that ever had deletions has holes, and
        len()+1 reuses an existing name (python-pptx's _next_slide_partname
        bug, corrupt package via duplicate ZIP entry names)."""
        prefix, _, suffix = template.partition("{}")
        pattern = re.compile(re.escape(prefix) + r"(\d+)" + re.escape(suffix) + r"\Z")
        highest = 0
        for name in self._raw:
            m = pattern.match(name)
            if m:
                highest = max(highest, int(m.group(1)))
        return template.format(highest + 1)

    def remove_part(self, name: str) -> None:
        """Drop a part from the in-memory package (raw bytes, cached tree,
        dirty flag, entry order). Content-type overrides and relationships
        pointing at the part are the caller's responsibility; the pre-save
        payload validation catches anything left dangling."""
        if name not in self._raw:
            raise KeyError(f"part not in package: {name}")
        del self._raw[name]
        self._order.remove(name)
        self._trees.pop(name, None)
        self._dirty.discard(name)
        self._compress.pop(name, None)

    def remove_content_type_override(self, part: str) -> bool:
        """Remove the [Content_Types].xml Override for `part` if present.
        Returns True when an Override was removed. Defaults are never touched
        (they cover other parts by extension)."""
        ct_root = self.root("[Content_Types].xml")
        part_name = "/" + part
        for node in ct_root.findall(qn("ct:Override")):
            if node.get("PartName") == part_name:
                ct_root.remove(node)
                self.mark_dirty("[Content_Types].xml")
                return True
        return False

    def register_slide_entry(
        self, slide_part: str, *, position: int | None = None
    ) -> dict:
        """Register an EXISTING slide part in the presentation spine: adds
        the presentation relationship and the p:sldId entry, at `position`
        (0-based index in deck order) or at the end. The slide part, its
        rels, and its content-type override must already be in the package;
        add_slide_part covers the build-from-XML path, this covers the
        clone-a-part path where the rels file is installed separately."""
        if slide_part not in self._raw:
            raise PptMcpError(f"slide part not in package: {slide_part}")
        pres = self.presentation()
        sld_id_lst = pres.find(qn("p:sldIdLst"))
        slide_id = self._next_slide_id(sld_id_lst)
        target = posixpath.relpath(slide_part, posixpath.dirname(PRESENTATION_PART))
        target = target.replace(os.sep, "/")
        rid = self.add_relationship(PRESENTATION_PART, RT_SLIDE, target)
        if sld_id_lst is None:
            sld_id_lst = etree.SubElement(pres, qn("p:sldIdLst"))
            pres.remove(sld_id_lst)
            self._insert_presentation_child(sld_id_lst)
        sld = etree.SubElement(sld_id_lst, qn("p:sldId"))
        sld.set("id", str(slide_id))
        sld.set(qn("r:id"), rid)
        entries = sld_id_lst.findall(qn("p:sldId"))
        index = len(entries) - 1
        if position is not None:
            if not 0 <= position < len(entries):
                raise PptMcpError(
                    f"position {position} out of range; the deck now has "
                    f"{len(entries)} slides (valid: 0..{len(entries) - 1})"
                )
            sld_id_lst.remove(sld)
            sld_id_lst.insert(position, sld)
            index = position
        self.mark_dirty(PRESENTATION_PART)
        return {"part": slide_part, "slide_id": slide_id, "rid": rid, "index": index}
