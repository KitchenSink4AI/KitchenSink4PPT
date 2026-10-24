"""Export layer: PDF and slide-image export with engine auto-routing.

Contract (differs from the pure-package ops modules by design):
- Functions here take FILE PATHS, not a PptxPackage; export is a
  representation transform of the saved file, never a package mutation, and
  the source file is never modified by any path through this module.
- Engine routing: "auto" prefers PowerPoint COM (Windows + PowerPoint
  installed + pywin32), falling back to LibreOffice headless
  (soffice --headless --convert-to). "com" and "libreoffice" force one
  engine. When neither engine exists the error names both options honestly.
- COM is ground truth for fidelity; LibreOffice output drifts on theme
  colors, some effects, and font substitution.
- Slide images via LibreOffice need a PDF rasterizer: soffice's own
  --convert-to png emits only the FIRST slide, so the LO image path converts
  to PDF then rasterizes with pdftoppm (poppler) when present, and refuses
  honestly when it is not. COM has no such dependency.
- Heavy imports (win32com via com.bridge) stay inside function bodies;
  this module imports cleanly on any platform.
- Every path parameter passes through the sandbox check_path at entry, and
  every produced artifact is verified to exist and be non-empty.
"""

from __future__ import annotations

import math
import os
import shutil
import subprocess
import zipfile
import sys
import tempfile
from pathlib import Path

from lxml import etree

from ..core import safesave, xmlsafe, zipsafe
from ..core.errors import PptMcpError
from ..core.package import qn, refuse_ole_encrypted
from ..core.sandbox import check_path, server_scratch
from ..core.swapwait import require_deck
from ..core.xmlbool import attr_bool

SOFFICE_WELL_KNOWN = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
SOFFICE_TIMEOUT = 300  # seconds; big decks convert slowly
PDFTOPPM_TIMEOUT = 300

#: Where poppler lands when a user installs it without putting it on PATH.
#: Same shape as SOFFICE_WELL_KNOWN, for the same reason: an installed tool
#: the server cannot see reads to the caller as a missing tool.
PDFTOPPM_WELL_KNOWN = (
    Path(r"C:\Program Files\poppler\Library\bin\pdftoppm.exe"),
    Path(r"C:\Program Files\poppler\bin\pdftoppm.exe"),
    Path(r"C:\Program Files (x86)\poppler\bin\pdftoppm.exe"),
    Path("/usr/bin/pdftoppm"),
    Path("/usr/local/bin/pdftoppm"),
    Path("/opt/homebrew/bin/pdftoppm"),
)

#: The remedy sentence for a LibreOffice install with no rasterizer. One
#: copy, used by the engine report and by the refusal, so the two cannot
#: drift apart.
PDFTOPPM_REMEDY = (
    " LibreOffice is installed but slide images also need "
    "pdftoppm (poppler) for PDF rasterization; install "
    "poppler or use PowerPoint COM."
)

DEFAULT_IMAGE_WIDTH = 1280


# ------------------------------------------------------------- detection


def _find_soffice() -> Path | None:
    env = os.environ.get("KS4P_SOFFICE")
    if env and Path(env).exists():
        return Path(env)
    if SOFFICE_WELL_KNOWN.exists():
        return SOFFICE_WELL_KNOWN
    hit = shutil.which("soffice")
    return Path(hit) if hit else None


def _runnable(path: Path) -> bool:
    """A usable executable: a FILE, and on POSIX one we may execute.

    Path.exists() alone accepted a DIRECTORY named pdftoppm, and on Unix a
    non-executable file. Either is reported as an available renderer and
    then fails at subprocess launch, which moves the error from detection
    (where it is actionable) to the middle of an export (where it is not)."""
    try:
        if not path.is_file():
            return False
    except OSError:
        return False
    if os.name == "posix":
        return os.access(path, os.X_OK)
    return True


def _find_pdftoppm() -> Path | None:
    env = os.environ.get("KS4P_PDFTOPPM")
    if env and _runnable(Path(env)):
        return Path(env)
    hit = shutil.which("pdftoppm")
    if hit:
        return Path(hit)
    for cand in PDFTOPPM_WELL_KNOWN:
        if _runnable(cand):
            return cand
    return None


def _com_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    from ..com.bridge import powerpoint_installed

    return powerpoint_installed()


def get_export_engines() -> dict:
    """What can this machine export with, right now?"""
    soffice = _find_soffice()
    pdftoppm = _find_pdftoppm()
    com_ok = _com_available()
    engines = {
        "powerpoint_com": {
            "available": com_ok,
            "pdf": com_ok,
            "slide_images": com_ok,
            "note": "ground truth for fidelity; Windows + PowerPoint + pywin32",
        },
        "libreoffice": {
            "available": soffice is not None,
            "path": str(soffice) if soffice else None,
            "pdf": soffice is not None,
            "slide_images": soffice is not None and pdftoppm is not None,
            "note": (
                "headless fallback; fidelity drifts on theme colors, effects, "
                "fonts. Slide images additionally need pdftoppm (poppler): "
                + ("found" if pdftoppm else "NOT found.")
                # The remedy belongs where the caller reads the verdict, not
                # only in the refusal they hit ten minutes later.
                + ("" if pdftoppm or soffice is None else PDFTOPPM_REMEDY)
            ),
        },
    }
    return {"engines": engines, "auto_pdf": _pick_engine_name(), "auto_images": _pick_image_engine_name()}


def _pick_engine_name() -> str | None:
    if _com_available():
        return "com"
    if _find_soffice() is not None:
        return "libreoffice"
    return None


def _pick_image_engine_name() -> str | None:
    if _com_available():
        return "com"
    if _find_soffice() is not None and _find_pdftoppm() is not None:
        return "libreoffice"
    return None


def _no_engine_error(what: str) -> PptMcpError:
    return PptMcpError(
        f"no {what} engine is available on this machine. Options: install "
        "Microsoft PowerPoint (COM path, Windows only, needs pywin32) or "
        "LibreOffice (headless soffice; set KS4P_SOFFICE if it is in a "
        "non-standard location)."
    )


# ------------------------------------------------------------- helpers


def _require_file(path: str, purpose: str) -> Path:
    path = check_path(path, purpose)
    require_deck(path, f"no file at {path}")
    return Path(path)


def _output_write_error(exc: OSError, what: str, target) -> PptMcpError:
    """Map a raw OSError/PermissionError on an export output path into the
    refusal envelope with the path and cause named: an
    unwritable output location is an ordinary user error, not a crash."""
    return PptMcpError(
        f"cannot write the {what} to {target}: {exc}. The output location "
        "must be a writable path (not read-only, locked by another "
        "process, or a directory name colliding with a file)."
    )


def _verify_output(out: Path, what: str, engine: str) -> None:
    if not out.exists() or out.stat().st_size == 0:
        raise PptMcpError(
            f"{engine} reported success but produced no {what} at {out}"
        )


def _run_soffice(soffice: Path, args: list[str], workdir: Path) -> None:
    """Run soffice headless with an ISOLATED user profile (a running
    LibreOffice GUI otherwise makes headless conversion silently no-op)."""
    profile = workdir / "lo_profile"
    profile.mkdir(parents=True, exist_ok=True)
    cmd = [
        str(soffice),
        "--headless",
        "--norestore",
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        *args,
    ]
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=SOFFICE_TIMEOUT,
            stdin=subprocess.DEVNULL,
        )
    except subprocess.TimeoutExpired as exc:
        raise PptMcpError(
            f"LibreOffice conversion timed out after {SOFFICE_TIMEOUT}s"
        ) from exc
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "").strip()[:500]
        raise PptMcpError(
            f"LibreOffice conversion failed (exit {result.returncode}): {detail}"
        )


def _libreoffice_pdf(p: Path, out: Path) -> None:
    soffice = _find_soffice()
    if soffice is None:  # callers pre-check; belt and braces
        raise _no_engine_error("PDF export")
    with tempfile.TemporaryDirectory(prefix="ks4p_lo_") as td:
        tdir = Path(td)
        _run_soffice(
            soffice,
            ["--convert-to", "pdf", "--outdir", str(tdir), str(p.resolve())],
            tdir,
        )
        produced = tdir / (p.stem + ".pdf")
        _verify_output(produced, "PDF", "LibreOffice")
        try:
            out.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(produced), str(out))
        except OSError as exc:
            raise _output_write_error(exc, "PDF", out) from exc
    _verify_output(out, "PDF", "LibreOffice")


# ------------------------------------------------------------- PDF


def _pdf_output(
    source: Path, output: str | None, *, purpose: str, overwrite: bool = False
) -> Path:
    """Return the exact PDF name PowerPoint/LibreOffice will receive.

    PowerPoint silently appends ``.pdf`` to an extensionless name, so all
    validation must happen after that extension is settled.  Exports are
    publish-new operations: an existing destination is replaced only when
    the call sends overwrite=true, and never when it is the deck itself or
    one of its backups.
    """
    raw = Path(output) if output else source.with_suffix(".pdf")
    # A trailing dot is an empty extension on every supported Python version.
    if raw.name.rstrip(".") and raw.name.endswith("."):
        raw = raw.with_name(raw.name.rstrip("."))
    out = raw.with_suffix(".pdf") if raw.suffix in ("", ".") else raw
    out = Path(check_path(str(out), f"{purpose} output"))
    safesave.refuse_protected_output(out, source, purpose=f"write the {purpose}")
    if out.is_dir():
        raise PptMcpError(
            f"refusing to write the {purpose}: {out} is a folder; name a .pdf file"
        )
    if out.exists() and not overwrite:
        raise PptMcpError(
            f"refusing to write the {purpose}: {out} already exists; send overwrite=true to replace it, or choose a new output path"
        )
    return out


def export_pdf(
    pkg_path: str, output: str | None = None, engine: str = "auto",
    overwrite: bool = False,
) -> dict:
    """Export a presentation to PDF. engine: "auto" | "com" | "libreoffice".

    Auto prefers PowerPoint COM (fidelity ground truth) and falls back to
    LibreOffice headless. The source file is never modified (COM opens
    ReadOnly and exports a copy; LibreOffice reads only). An existing PDF is
    replaced only with overwrite=True."""
    p = _require_file(pkg_path, "PDF export source")
    refuse_ole_encrypted(p)
    out = _pdf_output(p, output, purpose="PDF export", overwrite=overwrite)

    if engine not in ("auto", "com", "libreoffice"):
        raise PptMcpError(
            f"unknown engine '{engine}'; use auto, com, or libreoffice"
        )
    chosen = engine
    if engine == "auto":
        chosen = _pick_engine_name()
        if chosen is None:
            raise _no_engine_error("PDF export")

    if chosen == "com":
        if not _com_available():
            raise PptMcpError(
                "engine='com' but PowerPoint COM is not available here "
                "(needs Windows, PowerPoint, and pywin32); LibreOffice "
                + ("IS available, use engine='libreoffice'."
                   if _find_soffice() else "is not available either.")
            )
        from ..com.bridge import com_export_pdf

        try:
            safesave.claim_publish("the PDF export")
            return com_export_pdf(str(p), str(out), overwrite=overwrite)
        except OSError as exc:
            raise _output_write_error(exc, "PDF", out) from exc

    # libreoffice
    if _find_soffice() is None:
        raise PptMcpError(
            "engine='libreoffice' but soffice was not found (checked "
            f"KS4P_SOFFICE, {SOFFICE_WELL_KNOWN}, and PATH)."
        )
    safesave.claim_publish("the PDF export")
    _libreoffice_pdf(p, out)
    return {"pdf": str(out), "bytes": out.stat().st_size, "engine": "libreoffice"}


# ------------------------------------------------------------- handout PDF


def export_handout(
    pkg_path: str,
    output: str | None = None,
    slides_per_page: int = 3,
    include_notes: bool = False,
    overwrite: bool = False,
) -> dict:
    """Export a handout-layout PDF (N slides per page, or notes pages).

    COM-ONLY: handout and notes-page layouts are a PowerPoint print-pipeline
    feature (ExportAsFixedFormat ppPrintOutputType); LibreOffice's headless
    converter has no equivalent, so without PowerPoint this refuses honestly
    rather than emitting a plain one-slide-per-page PDF that is not a
    handout. The plain alternative is export_pdf (which does route through
    LibreOffice). slides_per_page: 1 | 2 | 3 | 4 | 6 | 9. include_notes=True
    exports notes pages (one slide + speaker notes per page) instead.
    """
    p = _require_file(pkg_path, "handout export source")
    refuse_ole_encrypted(p)
    default = p.with_name(
        p.stem + ("_notes.pdf" if include_notes else "_handout.pdf")
    )
    out = _pdf_output(
        p, output or str(default), purpose="handout export", overwrite=overwrite
    )
    if not _com_available():
        raise PptMcpError(
            "handout export needs PowerPoint COM (Windows + PowerPoint + "
            "pywin32): handout page layouts are a PowerPoint print feature "
            "with no LibreOffice equivalent. For a plain one-slide-per-page "
            "PDF use export_pdf"
            + (
                ", which IS available here via LibreOffice."
                if _find_soffice()
                else " once an engine is installed."
            )
        )
    from ..com.bridge import com_export_handout

    try:
        safesave.claim_publish("the handout export")
        return com_export_handout(
            str(p),
            str(out),
            slides_per_page=slides_per_page,
            include_notes=include_notes,
            overwrite=overwrite,
        )
    except OSError as exc:
        raise _output_write_error(
            exc, "handout PDF", out
        ) from exc


# ------------------------------------------------------------- slide images


def _numbered(result: dict) -> dict:
    """Each image leads with its slide counted from 1, the number its
    slideN.png file is named after, beside the 0-based index, so the first
    thing a reader sees is the slide the way people count it."""
    images = result.get("images")
    if not isinstance(images, list):
        return result
    for n, image in enumerate(images):
        index = image.get("slide") if isinstance(image, dict) else None
        if isinstance(index, int) and not isinstance(index, bool):
            number = image.get("slide_number", index + 1)
            images[n] = {"slide_number": number,
                         **{k: v for k, v in image.items() if k != "slide_number"}}
    return result


def _slide_count(p: Path) -> int | None:
    """The deck's slide count read straight from presentation.xml, without
    the edit lock (a deck open in PowerPoint can still be rendered), or
    None when it cannot be read."""
    try:
        with zipfile.ZipFile(p) as zf:
            info = zf.getinfo("ppt/presentation.xml")
            if zipsafe.too_large_to_read(info):
                return None
            root = xmlsafe.fromstring(zf.read(info), doc=p.name, part=info.filename)
    except Exception:  # noqa: BLE001 - unknown means "look at every file"
        return None
    lst = root.find(qn("p:sldIdLst"))
    return 0 if lst is None else len(lst.findall(qn("p:sldId")))


def existing_images(p: Path, out_dir: Path, slides) -> list[Path]:
    """The slideN.png files a render of `slides` (0-based; None for every
    slide) would replace in out_dir."""
    if not out_dir.is_dir():
        return []
    if slides is None:
        count = _slide_count(p)
        found = []
        for item in out_dir.glob("slide*.png"):
            digits = item.stem[len("slide"):]
            if digits.isdigit() and (count is None or 1 <= int(digits) <= count):
                found.append((int(digits), item))
        return [item for _n, item in sorted(found)]
    names = []
    for index in slides:
        if isinstance(index, int) and not isinstance(index, bool) and index >= 0:
            item = out_dir / f"slide{index + 1}.png"
            if item.exists() and item not in names:
                names.append(item)
    return names


def refuse_existing(found: list[Path], what: str, where,
                    elsewhere: str = "choose another output folder") -> None:
    """Refuse an export that would replace files, before any engine starts,
    naming them (at most ten) and the two ways on."""
    if not found:
        return
    shown = ", ".join(item.name for item in found[:10])
    more = f" and {len(found) - 10} more" if len(found) > 10 else ""
    one = len(found) == 1
    raise PptMcpError(
        f"refusing to write the {what}: {shown}{more} already "
        f"exist{'s' if one else ''} in {where}; send overwrite=true to replace "
        f"{'it' if one else 'them'}, or {elsewhere}"
    )


def export_slide_images(
    pkg_path: str,
    output_dir: str | None = None,
    slides: list[int] | None = None,
    width: int = DEFAULT_IMAGE_WIDTH,
    height: int | None = None,
    engine: str = "auto",
    overwrite: bool = False,
) -> dict:
    """Export slides as PNG images (the render-to-verify primitive).

    `slides` takes 0-based presentation-order indices; None = all slides.
    engine "auto" prefers PowerPoint COM; the LibreOffice path converts to
    PDF then rasterizes with pdftoppm (poppler) and refuses honestly when
    pdftoppm is missing (soffice alone can only render the first slide).
    An existing slideN.png is replaced only with overwrite=True: without it
    the call refuses, naming the files, before any engine starts."""
    p = _require_file(pkg_path, "slide image export source")
    refuse_ole_encrypted(p)
    if output_dir:
        output_dir = check_path(output_dir, "slide image output directory")
    out_dir = Path(output_dir) if output_dir else p.with_name(p.stem + "_slides")
    safesave.refuse_protected_output(
        out_dir, p, purpose="write the slide-image export"
    )
    if not overwrite:
        refuse_existing(existing_images(p, out_dir, slides), "slide images", out_dir)

    if engine not in ("auto", "com", "libreoffice"):
        raise PptMcpError(
            f"unknown engine '{engine}'; use auto, com, or libreoffice"
        )
    chosen = engine
    if engine == "auto":
        chosen = _pick_image_engine_name()
        if chosen is None:
            hint = ""
            if _find_soffice() is not None and _find_pdftoppm() is None:
                hint = PDFTOPPM_REMEDY
            raise PptMcpError(
                "no slide-image engine is available on this machine. Options: "
                "Microsoft PowerPoint (COM) or LibreOffice + poppler "
                "(pdftoppm)." + hint
            )

    if chosen == "com":
        if not _com_available():
            raise PptMcpError(
                "engine='com' but PowerPoint COM is not available here "
                "(needs Windows, PowerPoint, and pywin32)."
            )
        from ..com.bridge import com_export_slide_images

        try:
            safesave.claim_publish("the slide-image export")
            return _numbered(com_export_slide_images(
                str(p),
                str(out_dir),
                slides=slides,
                width=width,
                height=height if height is not None else None,
            ))
        except OSError as exc:
            raise _output_write_error(exc, "slide images", out_dir) from exc

    # libreoffice: pptx -> pdf -> pdftoppm PNGs
    if _find_soffice() is None:
        raise PptMcpError(
            "engine='libreoffice' but soffice was not found (checked "
            f"KS4P_SOFFICE, {SOFFICE_WELL_KNOWN}, and PATH)."
        )
    pdftoppm = _find_pdftoppm()
    if pdftoppm is None:
        raise PptMcpError(
            "the LibreOffice slide-image path needs pdftoppm (poppler) to "
            "rasterize the intermediate PDF; soffice alone renders only the "
            "first slide. Install poppler or use the PowerPoint COM engine."
        )
    if width <= 0:
        raise PptMcpError(f"width must be positive, got {width}")
    safesave.claim_publish("the slide-image export")
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise _output_write_error(exc, "slide images", out_dir) from exc
    images: list[dict] = []
    with tempfile.TemporaryDirectory(prefix="ks4p_lo_img_") as td:
        tdir = Path(td)
        pdf_tmp = tdir / (p.stem + ".pdf")
        _libreoffice_pdf(p, pdf_tmp)
        scale = ["-scale-to-x", str(width), "-scale-to-y", "-1"]
        if height is not None:
            scale = ["-scale-to-x", str(width), "-scale-to-y", str(height)]
        try:
            result = subprocess.run(
                [
                    str(pdftoppm), "-png", *scale,
                    str(pdf_tmp), str(tdir / "page"),
                ],
                capture_output=True,
                text=True,
                timeout=PDFTOPPM_TIMEOUT,
                stdin=subprocess.DEVNULL,
            )
        except subprocess.TimeoutExpired as exc:
            raise PptMcpError(
                f"pdftoppm timed out after {PDFTOPPM_TIMEOUT}s"
            ) from exc
        if result.returncode != 0:
            detail = (result.stderr or "").strip()[:500]
            raise PptMcpError(f"pdftoppm failed (exit {result.returncode}): {detail}")
        pages = sorted(tdir.glob("page-*.png"))
        if not pages:
            raise PptMcpError(
                "pdftoppm reported success but produced no PNG pages"
            )
        wanted = list(range(len(pages))) if slides is None else list(slides)
        for idx in wanted:
            if not isinstance(idx, int) or idx < 0 or idx >= len(pages):
                raise PptMcpError(
                    f"slide index {idx} out of range, presentation has "
                    f"{len(pages)} (indices are 0-based)"
                )
            dest = out_dir / f"slide{idx + 1}.png"
            try:
                shutil.move(str(pages[idx]), str(dest))
            except OSError as exc:
                raise _output_write_error(
                    exc, f"PNG for slide {idx + 1}", dest
                ) from exc
            _verify_output(dest, f"PNG for slide {idx + 1}", "LibreOffice")
            images.append({"slide_number": idx + 1, "slide": idx, "file": str(dest)})
    return {
        "images": images,
        "width": width,
        "height": height,
        "engine": "libreoffice",
        "fidelity_note": (
            "LibreOffice rendering; theme colors, some effects, and fonts "
            "may drift from PowerPoint output"
        ),
    }


# ------------------------------------------------------------- one shape


#: The largest slide render a shape image asks PowerPoint for, in pixels on
#: the long side. A shape too small on its slide for the size asked is cut
#: from a render this large and enlarged, and the answer says so.
MAX_RENDER_PX = 8192
DEFAULT_SHAPE_WIDTH = 1280


def _turn(points, cx: float, cy: float, degrees: float):
    """points turned clockwise (y runs down) by degrees about (cx, cy)."""
    if not degrees:
        return points
    rad = math.radians(degrees)
    cos, sin = math.cos(rad), math.sin(rad)
    return [(cx + (x - cx) * cos - (y - cy) * sin,
             cy + (x - cx) * sin + (y - cy) * cos) for x, y in points]


def _rot(xfrm) -> float:
    try:
        return int(xfrm.get("rot") or 0) / 60000.0
    except ValueError:
        return 0.0


def _outline_emu(pkg, part, elem) -> int:
    """The widest outline in elem (a shape or a group), in EMU: an explicit
    a:ln width, or the theme line style a shape's style points at."""
    from . import scaling as _scaling

    widest = 0
    theme = None
    for member in _scaling._members(elem):
        sppr = member.find(qn("p:spPr"))
        ln = sppr.find(qn("a:ln")) if sppr is not None else None
        if ln is not None and ln.find(qn("a:noFill")) is not None:
            continue
        if ln is not None and (ln.get("w") or "").isdigit():
            widest = max(widest, int(ln.get("w")))
            continue
        ref = member.find(f"{qn('p:style')}/{qn('a:lnRef')}")
        idx = ref.get("idx", "0") if ref is not None else "0"
        if idx.isdigit() and int(idx) > 0:
            if theme is None:
                theme = _scaling._theme_line_widths(pkg, part)
            width = theme[int(idx) - 1] if int(idx) <= len(theme) else None
            widest = max(widest, width or 0)
        elif ln is not None:
            widest = max(widest, 9525)  # an outline with no width is 0.75 pt
    return widest


def _slide_size(pkg) -> tuple[int, int]:
    """(width, height) of the deck's slides in EMU; 10 x 7.5 in when unset."""
    size = pkg.presentation().find(qn("p:sldSz"))
    try:
        return int(size.get("cx")), int(size.get("cy"))
    except (AttributeError, TypeError, ValueError):
        return 9144000, 6858000


def shape_area(pkg, part: str, elem, chain) -> tuple[tuple[float, ...], bool]:
    """((x, y, w, h) in EMU, clipped) for the area a shape is drawn in on
    its slide: its box through its groups, turned with any rotation (its
    own and its groups'), grown by half its widest outline, and clipped to
    the slide."""
    from . import inherit as _inh
    from . import shapes as _sh

    xfrm = _sh._xfrm_of(elem)
    if xfrm is not None and xfrm.find(qn("a:off")) is not None:
        x, y, cx, cy = _sh._xfrm_box(xfrm)
        own = _rot(xfrm)
    else:
        box = _inh.inherited_box(pkg, part, elem)
        if box is None:
            raise PptMcpError(
                "that shape has no box of its own or from its layout, so there is "
                "no area to render")
        x, y, cx, cy = box
        own = 0.0
    points = _turn([(x, y), (x + cx, y), (x + cx, y + cy), (x, y + cy)],
                   x + cx / 2, y + cy / 2, own)
    for grp in reversed(chain):  # innermost group first
        gx = grp.find(f"{qn('p:grpSpPr')}/{qn('a:xfrm')}")
        if gx is None:
            continue
        ox, oy, ow, oh = _sh._xfrm_box(gx)
        choff, chext = gx.find(qn("a:chOff")), gx.find(qn("a:chExt"))
        chx = int(choff.get("x")) if choff is not None else ox
        chy = int(choff.get("y")) if choff is not None else oy
        chw = int(chext.get("cx")) if chext is not None else ow
        chh = int(chext.get("cy")) if chext is not None else oh
        sx = ow / chw if chw else 1.0
        sy = oh / chh if chh else 1.0
        points = [(ox + (px - chx) * sx, oy + (py - chy) * sy) for px, py in points]
        mx, my = ox + ow / 2, oy + oh / 2
        if attr_bool(gx, "flipH"):
            points = [(2 * mx - px, py) for px, py in points]
        if attr_bool(gx, "flipV"):
            points = [(px, 2 * my - py) for px, py in points]
        points = _turn(points, mx, my, _rot(gx))
    half = _outline_emu(pkg, part, elem) / 2
    x0 = min(px for px, _py in points) - half
    y0 = min(py for _px, py in points) - half
    x1 = max(px for px, _py in points) + half
    y1 = max(py for _px, py in points) + half
    width, height = _slide_size(pkg)
    cx0, cy0 = max(0.0, x0), max(0.0, y0)
    cx1, cy1 = min(float(width), x1), min(float(height), y1)
    if cx1 <= cx0 or cy1 <= cy0:
        raise PptMcpError(
            "that shape is drawn entirely off the slide, so nothing of it would "
            "render; move it onto the slide first (set_shape)")
    clipped = (cx0, cy0, cx1, cy1) != (x0, y0, x1, y1)
    return (cx0, cy0, cx1 - cx0, cy1 - cy0), clipped


def _image_size(area, width, height) -> tuple[int, int]:
    _x, _y, w, h = area
    for value, name in ((width, "width"), (height, "height")):
        if value is not None and (isinstance(value, bool) or not isinstance(value, int)
                                  or not 1 <= value <= 20000):
            raise PptMcpError(f"{name} is a whole number of pixels from 1 to 20000")
    if width is None and height is None:
        width = DEFAULT_SHAPE_WIDTH
    if width is not None and height is not None:
        if abs(width / height - w / h) > 0.01 * (w / h):
            raise PptMcpError(
                f"width {width} and height {height} do not keep the shape's proportions "
                f"({w / h:.3f} wide to 1 high); send one of them")
        return width, height
    if width is not None:
        return width, max(1, round(width * h / w))
    return max(1, round(height * w / h)), height


def _safe_name(text: str) -> str:
    keep = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in text)
    return keep.strip("_")[:40] or "shape"


def export_shape_image(
    pkg_path: str,
    slide,
    shape,
    output: str | None = None,
    width: int | None = None,
    height: int | None = None,
    overwrite: bool = False,
) -> dict:
    """Render one shape or group to a PNG: PowerPoint renders the slide at
    the scale that gives the shape the size asked for, and the area the
    shape is drawn in is cut out (shape_area). Needs PowerPoint to render
    and Pillow (the optimize extra) to cut; both are checked before
    anything starts, as is the output file, which is replaced only with
    overwrite=True. The presentation is never modified."""
    from ..core.package import PptxPackage
    from . import shapes as _sh
    from .read import resolve_slide

    p = _require_file(pkg_path, "shape image export source")
    refuse_ole_encrypted(p)
    pkg = PptxPackage(p)
    rec = resolve_slide(pkg, slide)
    elem, chain = _sh._find_shape(pkg, rec["part"], shape)
    shape_id, name = _sh._shape_id(elem), _sh._shape_name(elem)
    area, clipped = shape_area(pkg, rec["part"], elem, chain)
    out_w, out_h = _image_size(area, width, height)

    raw = Path(output) if output else p.with_name(
        f"{p.stem}_slide{rec['index'] + 1}_{_safe_name(name or str(shape_id))}.png")
    if raw.suffix.lower() != ".png":
        raw = raw.with_name(raw.name + ".png")
    out = Path(check_path(str(raw), "shape image output"))
    safesave.refuse_protected_output(out, p, purpose="write the shape image")
    if out.is_dir():
        raise PptMcpError(f"refusing to write the shape image: {out} is a folder")
    if out.exists() and not overwrite:
        refuse_existing([out], "shape image", out.parent, "choose another output file")
    if not _com_available():
        raise PptMcpError(
            "export_shape_image needs PowerPoint (Windows, PowerPoint and pywin32) to "
            "render the slide; export_slide_image with engine='libreoffice' renders "
            "whole slides without it")
    try:
        from PIL import Image
    except ImportError:
        raise PptMcpError(
            "export_shape_image needs Pillow to cut the shape out of the slide render; "
            'install it with the optimize extra: pip install "kitchensink4ppt[optimize]"'
        ) from None

    slide_w, slide_h = _slide_size(pkg)
    per_emu = out_w / area[2]
    render_w, render_h = round(slide_w * per_emu), round(slide_h * per_emu)
    warnings: list[str] = []
    shrink = min(1.0, MAX_RENDER_PX / max(render_w, render_h))
    if shrink < 1.0:
        render_w = max(1, round(render_w * shrink))
        render_h = max(1, round(render_h * shrink))
    scale = render_w / slide_w  # pixels per EMU in the render actually made
    box = (round(area[0] * scale), round(area[1] * scale),
           round((area[0] + area[2]) * scale), round((area[1] + area[3]) * scale))

    from ..com.bridge import com_export_slide_images

    safesave.claim_publish("the shape image")
    # The slide render is the server's own intermediate file: it goes into a
    # scratch folder no caller names, so KS4P_ALLOWED_ROOTS, which governs
    # the deck and the output checked above, does not refuse it.
    with server_scratch("ks4p_shape_") as scratch:
        try:
            made = com_export_slide_images(
                str(p), output_dir=scratch, slides=[rec["index"]], width=render_w,
                height=render_h)
        except OSError as exc:
            raise _output_write_error(exc, "slide render", scratch.path) from exc
        rendered = Path(made["images"][0]["file"])
        with Image.open(rendered) as image:
            piece = image.crop(box)
            if piece.size != (out_w, out_h):
                if shrink < 1.0:
                    warnings.append(
                        f"the shape is small on its slide, so it was cut from a "
                        f"{render_w}-pixel-wide render ({piece.size[0]} pixels) and "
                        f"enlarged to {out_w}; ask for a smaller width for a sharper image")
                piece = piece.resize((out_w, out_h), Image.LANCZOS)
            target = out.with_name(f".~ks4p-{os.getpid()}-{out.name}") if out.exists() else out
            try:
                out.parent.mkdir(parents=True, exist_ok=True)
                piece.save(target, "PNG")
                if target != out:
                    os.replace(target, out)
            except OSError as exc:
                raise _output_write_error(exc, "shape image", out) from exc
            finally:
                if target != out and target.exists():
                    target.unlink()
    _verify_output(out, "shape image", "PowerPoint")
    inch = 914400.0
    per_px = out_w / area[2]
    result = {
        "file": str(out),
        "shape_id": shape_id,
        "name": name,
        "slide_index": rec["index"],
        "slide_id": rec["slide_id"],
        "image": {"width": out_w, "height": out_h},
        "bounds": {
            "inches": {k: round(v / inch, 4) for k, v in zip("xywh", area)},
            "pixels": {k: round(v * per_px) for k, v in zip("xywh", area)},
        },
        "clipped": clipped,
        "render": {"width": render_w, "height": render_h},
        "engine": "powerpoint-com",
    }
    if clipped:
        warnings.append(
            "part of the shape lies off the slide; the image holds only the part on it")
    if warnings:
        result["warnings"] = warnings
    return result

