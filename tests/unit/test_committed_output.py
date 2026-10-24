"""A change that was saved is never reported as a failure.

Regression: set_theme_fonts with a 200,000-character East Asian typeface
wrote the theme, saved the deck, and then answered LIMIT_EXCEEDED, because
the answer echoed the name four times and the final output cap gave up on
it. The client was told the call failed when the file had changed. Two
fixes are held here: the output cap always finds a way to shorten a
committed write's answer, and a typeface no installed font could carry is
refused before anything is written.
"""

from __future__ import annotations

import asyncio
import hashlib

from kitchensink4ppt import server  # noqa: F401 - populates the registry
from kitchensink4ppt.core import budget
from kitchensink4ppt.registry import OpSpec, StrictArgs, public_ops
from kitchensink4ppt.settings import Settings
from kitchensink4ppt.surface import envelope
from kitchensink4ppt.surface.dispatch import dispatch
from kitchensink4ppt.surface.plan import build


class _PathArgs(StrictArgs):
    file_path: str


def _sha(path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _call(tool: str, op: str, args: dict, *, cap: int | None = None):
    settings = Settings() if cap is None else Settings(max_output_chars=cap)
    plan = build(settings, public_ops())
    return asyncio.run(dispatch(plan, tool, {"op": op, "args": args}))


def _fake(name: str, group: str, result: dict, *, read_only: bool) -> tuple:
    spec = OpSpec(
        name, group, "Fake operation.", _PathArgs, lambda **_kwargs: result,
        not read_only, read_only, False, False,
    )
    plan = build(Settings(max_output_chars=4_000), {name: spec})
    return asyncio.run(
        dispatch(plan, group, {"op": name, "args": {"file_path": "deck.pptx"}})
    )


def test_the_output_cap_shortens_several_copies_of_one_long_echo():
    long = "E" * 200_000
    result = {
        "ok": True,
        "changed": {
            "set": {"major": {"ea": long}, "minor": {"ea": long}},
            "fonts": {"major": {"ea": long}, "minor": {"ea": long}},
        },
    }
    capped = budget.enforce(result, 52_000, wrap=envelope.success)
    assert len(envelope.compact_json(envelope.success(capped))) <= 52_000
    assert capped["page"]["truncated"] is True
    assert capped["ok"] is True


def test_a_committed_write_too_large_to_describe_still_reports_success():
    # Thousands of keys cannot be cut by shortening any one value, which is
    # the shape that made the cap give up. A read may refuse; a write that
    # already saved may not.
    result = {
        "ok": True,
        "file": "deck.pptx",
        "changed": {f"shape_{index}": index for index in range(5_000)},
        "saved": "deck.pptx",
        "backup": True,
    }
    payload, is_error = _fake("set_shape", "ppt_graphics", result, read_only=False)
    assert is_error is False, payload
    assert payload["ok"] is True
    assert payload["saved"] == "deck.pptx"
    assert payload["page"]["truncated"] is True
    assert "saved" in payload["page"]["note"] or "made" in payload["page"]["note"]
    assert len(envelope.compact_json(payload)) <= 4_000

    payload, is_error = _fake("get_text", "ppt_read", result, read_only=True)
    assert is_error is True
    assert payload["error"]["code"] == "LIMIT_EXCEEDED"


def test_an_absurd_theme_typeface_is_refused_before_anything_is_written(make_deck):
    deck = make_deck()
    before = _sha(deck)
    payload, is_error = _call(
        "ppt_design", "set_theme_fonts",
        {"file_path": str(deck), "ea": "E" * 200_000},
    )
    assert is_error is True, str(payload)[:400]
    assert payload["error"]["code"] == "BAD_ARGS"
    assert _sha(deck) == before
    assert "E" * 100 not in envelope.compact_json(payload)


def test_theme_typefaces_hold_the_windows_face_name_limit(make_deck):
    deck = make_deck()
    payload, is_error = _call(
        "ppt_design", "set_theme_fonts",
        {"file_path": str(deck), "major": "A" * 31},
    )
    assert is_error is False, payload
    before = _sha(deck)
    for args in (
        {"major": "B" * 32},
        {"minor": {"latin": "C" * 40}},
        {"major": {"cs": "D" * 33}, "minor": "Calibri"},
    ):
        payload, is_error = _call(
            "ppt_design", "set_theme_fonts", {"file_path": str(deck), **args}
        )
        assert is_error is True, args
        assert payload["error"]["code"] == "BAD_ARGS", payload
        assert _sha(deck) == before, args


def test_a_snapshot_of_an_empty_or_truncated_deck_is_refused(make_deck, tmp_path):
    """A permanent snapshot of a broken file keeps nothing worth keeping and
    reads as a safe restore point; refuse it and say why."""
    whole = make_deck()
    data = whole.read_bytes()
    for name, content in (("empty.pptx", b""), ("half.pptx", data[: len(data) // 2])):
        broken = tmp_path / name
        broken.write_bytes(content)
        before = sorted(p.name for p in tmp_path.iterdir())
        payload, is_error = _call(
            "ppt_deck", "create_snapshot", {"file_path": str(broken)}
        )
        assert is_error is True, (name, payload)
        assert payload["error"]["code"] == "UNSUPPORTED_CONTENT", payload
        assert sorted(p.name for p in tmp_path.iterdir()) == before, name

    payload, is_error = _call("ppt_deck", "create_snapshot", {"file_path": str(whole)})
    assert is_error is False, payload


_LONG_PATH = "\\".join((
    "C:", "Users", "someone", "OneDrive - Example Org", "Documents", "Quarterly reviews",
    "2026", "Board material, final versions agreed with the board",
    "Appendices and backup", "Q3 board review, appendix tables and the revenue bridge.pptx",
))


def _big_write(path: str) -> dict:
    return {
        "slide_number": 2, "slide_count": 5,
        "slide_title": "Revenue bridge, quarter three, with the one-off items separated",
        "slide_note": "You sent slide=1, which counts from 0, so this is the second slide.",
        "ok": True, "file": path,
        "changed": {str(i): {"shape_id": i, "text": "x" * 300} for i in range(1_500)},
        "saved": path, "backup": True, "warnings": ["w" * 400] * 40,
    }


def test_a_shortened_answer_keeps_where_it_saved_and_which_slide_whole():
    """An earlier version clipped every string to 60 characters in its last
    tiers, file and saved included, so a shortened answer named a mangled
    path: d2000.pptx came back as "...d2000.ppt" plus a truncation marker, a
    different file type. Paths over 200 characters were cut in the first
    tier too."""
    assert len(_LONG_PATH) > 200
    result = _big_write(_LONG_PATH)
    for cap in (52_000, 4_000, 2_000):
        out = budget.committed(result, cap, wrap=envelope.success, note="note")
        assert budget._wire_size(envelope.success(out)) <= cap, cap
        for key in ("ok", "file", "saved", "backup", "slide_number", "slide_count",
                    "slide_title", "slide_note"):
            assert out[key] == result[key], (cap, key)
        assert out["page"]["truncated"] is True


def test_a_field_that_cannot_fit_is_left_out_whole_never_cut():
    """At a cap too small for every whole field, each one is either exactly
    as written or absent: never a shortened path that names another file."""
    result = _big_write(_LONG_PATH)
    out = budget.committed(result, 600, wrap=envelope.success, note="note")
    assert budget._wire_size(envelope.success(out)) <= 600
    assert out["ok"] is True
    for key in budget.COMMITTED_WHOLE:
        assert key not in out or out[key] == result[key], key
    assert out.get("file") == _LONG_PATH  # the first field given room


def test_a_real_write_over_the_cap_reports_its_exact_path(make_deck, tmp_path):
    """End to end through dispatch at the minimum cap: the deck's real path
    (longer than 60 characters) comes back exactly in file and saved."""
    folder = tmp_path / ("a folder name long enough to pass sixty characters " * 2).strip()
    folder.mkdir()
    deck = folder / "board review.pptx"
    deck.write_bytes(make_deck(extra_slides=60).read_bytes())
    # Many small per-edit results: nothing one cut can shorten, which is
    # the shape that reaches the committed path (hundreds of edits in one call).
    edits = []
    for index in range(server.get_presentation_info(str(deck))["slide_count"]):
        info = server.get_slide_info(str(deck), index, live="off")
        edits += [
            {"op": "set_text", "slide": index, "shape": shape["id"], "text": f"edited {index}"}
            for shape in info["shapes"]
            if shape.get("type") in ("placeholder", "textbox") and "group_id" not in shape
        ]
    assert len(edits) > 100
    payload, is_error = _call(
        "ppt_edit", "apply_edits", {"file_path": str(deck), "edits": edits}, cap=2_000,
    )
    assert is_error is False, str(payload)[:400]
    assert payload["page"]["truncated"] is True and "stands" in payload["page"]["note"]
    assert payload["applied"] == len(edits)
    assert payload["file"] == str(deck)
    assert payload["saved"] == str(deck)
    assert len(envelope.compact_json(payload)) <= 2_000


def _stored_copy(src, dst):
    """The deck with every part stored uncompressed, so a test can change
    one byte of a part in place without touching its recorded checksum."""
    import zipfile

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_STORED) as zout:
        for info in zin.infolist():
            zout.writestr(info.filename, zin.read(info.filename))
    return dst


def _flip_inside(path, part: str, needle: bytes) -> None:
    """Change one byte of ``needle`` inside ``part``'s stored data."""
    import zipfile

    with zipfile.ZipFile(path) as package:
        info = package.getinfo(part)
    data = bytearray(path.read_bytes())
    name_len = int.from_bytes(data[info.header_offset + 26:info.header_offset + 28], "little")
    extra_len = int.from_bytes(data[info.header_offset + 28:info.header_offset + 30], "little")
    start = info.header_offset + 30 + name_len + extra_len
    at = bytes(data[start:start + info.compress_size]).index(needle) + start
    data[at] ^= 0x20  # a letter changes case; the XML stays well-formed
    path.write_bytes(bytes(data))


def test_a_snapshot_checks_what_it_says_it_checks(make_deck, tmp_path):
    """An earlier version checked only the zip directory and two part names,
    so a deck with a damaged part, a presentation.xml that was not XML, or a
    bare 8-byte OLE header was still kept as a restore point. Each is
    refused, with nothing written; a sound deck and a whole OLE frame are
    still kept."""
    import zipfile

    whole = make_deck()
    stored = _stored_copy(whole, tmp_path / "stored.pptx")
    payload, is_error = _call("ppt_deck", "create_snapshot", {"file_path": str(stored)})
    assert is_error is False, payload  # the control: stored parts are fine

    cases = {}
    flipped = _stored_copy(whole, tmp_path / "flipped.pptx")
    _flip_inside(flipped, "ppt/presentation.xml", b"presentationml")
    cases["flipped.pptx"] = "cut off or damaged"
    slide = _stored_copy(whole, tmp_path / "slide.pptx")
    _flip_inside(slide, "ppt/slides/slide1.xml", b"presentationml")
    cases["slide.pptx"] = "cut off or damaged"
    not_xml = tmp_path / "not_xml.pptx"
    with zipfile.ZipFile(not_xml, "w") as package:
        package.writestr("[Content_Types].xml", "<Types")
        package.writestr("ppt/presentation.xml", "not xml at all")
    cases["not_xml.pptx"] = "is not readable XML"
    wrong_root = tmp_path / "wrong_root.pptx"
    with zipfile.ZipFile(whole) as zin, zipfile.ZipFile(wrong_root, "w") as zout:
        for info in zin.infolist():
            body = zin.read(info.filename)
            if info.filename == "ppt/presentation.xml":
                body = b"<document/>"
            zout.writestr(info.filename, body)
    cases["wrong_root.pptx"] = "holds no PowerPoint presentation"
    bare_ole = tmp_path / "bare_ole.pptx"
    bare_ole.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
    cases["bare_ole.pptx"] = "cut off or damaged"

    for name, reason in cases.items():
        before = sorted(p.name for p in tmp_path.iterdir())
        payload, is_error = _call(
            "ppt_deck", "create_snapshot", {"file_path": str(tmp_path / name)}
        )
        assert is_error is True, (name, payload)
        assert payload["error"]["code"] == "UNSUPPORTED_CONTENT", payload
        assert reason in payload["error"]["message"], (name, payload["error"]["message"])
        assert sorted(p.name for p in tmp_path.iterdir()) == before, name

    # A whole OLE frame (header, byte-order mark, 512-byte sectors, one
    # sector) is a password-protected deck this server cannot open: kept.
    header = bytearray(512)
    header[:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    header[28:30] = b"\xfe\xff"
    header[30:32] = (9).to_bytes(2, "little")
    framed = tmp_path / "protected.pptx"
    framed.write_bytes(bytes(header) + bytes(512))
    payload, is_error = _call("ppt_deck", "create_snapshot", {"file_path": str(framed)})
    assert is_error is False, payload


def _record_size(path, part: str, size: int) -> None:
    """Change the uncompressed size a deck's zip directory records for ``part``,
    leaving the data and its checksum alone (what a damaged central directory
    looks like)."""
    import struct

    data = bytearray(path.read_bytes())
    at = 0
    while True:
        at = data.index(b"PK", at)
        name_len = struct.unpack_from("<H", data, at + 28)[0]
        if bytes(data[at + 46:at + 46 + name_len]) == part.encode():
            struct.pack_into("<I", data, at + 24, size)
            break
        at += 4
    path.write_bytes(bytes(data))


def test_a_snapshot_never_skips_a_required_part_whatever_size_it_claims(make_deck, tmp_path):
    """The check reads at most SNAPSHOT_CHECK_BYTES of part data and copies a
    larger package's remaining parts unchecked. That budget skipped the two
    required parts as well: a presentation.xml that was not XML, recorded as
    4 GiB by a damaged central directory, was kept as a restore point nobody
    could open. A required part that claims more than the budget is refused,
    and nothing is written."""
    import zipfile

    good = make_deck()
    bad = tmp_path / "bad.pptx"
    with zipfile.ZipFile(good) as zin, zipfile.ZipFile(bad, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            body = b"this is not xml" if info.filename == "ppt/presentation.xml" else zin.read(info.filename)
            zout.writestr(info.filename, body)
    _record_size(bad, "ppt/presentation.xml", 0xFFFFFFF0)
    dest = tmp_path / "snaps"
    dest.mkdir()
    payload, is_error = _call("ppt_deck", "create_snapshot", {"file_path": str(bad), "dest_dir": str(dest)})
    assert is_error is True, payload
    assert payload["error"]["code"] == "UNSUPPORTED_CONTENT"
    assert "ppt/presentation.xml" in payload["error"]["message"]
    assert "beyond any real deck" in payload["error"]["message"]
    assert not list(dest.iterdir())

    # The same for the content types.
    types_bad = tmp_path / "types_bad.pptx"
    with zipfile.ZipFile(good) as zin, zipfile.ZipFile(types_bad, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info.filename, zin.read(info.filename))
    _record_size(types_bad, "[Content_Types].xml", 0xFFFFFFF0)
    payload, is_error = _call("ppt_deck", "create_snapshot", {"file_path": str(types_bad), "dest_dir": str(dest)})
    assert is_error is True, payload
    assert "[Content_Types].xml" in payload["error"]["message"]
    assert not list(dest.iterdir())


def test_a_package_past_the_check_budget_is_still_kept_with_its_other_parts_unchecked(
        make_deck, tmp_path, monkeypatch):
    """What the budget is for: in a package whose parts add up to more than
    it, the parts after the two required ones are copied without a checksum
    test. The required parts are checked all the same."""
    import zipfile

    from kitchensink4ppt.ops import backups

    good = make_deck()
    with zipfile.ZipFile(good) as zin:
        required = sum(zin.getinfo(name).file_size
                       for name in ("[Content_Types].xml", "ppt/presentation.xml"))
    big = tmp_path / "big.pptx"
    with zipfile.ZipFile(good) as zin, zipfile.ZipFile(big, "w", zipfile.ZIP_STORED) as zout:
        for info in zin.infolist():
            zout.writestr(info.filename, zin.read(info.filename))
        zout.writestr("ppt/media/filler.bin", b"x" * 50_000)
    data = bytearray(big.read_bytes())
    data[data.index(b"x" * 100)] ^= 0x01  # a damaged byte in a part the budget skips
    big.write_bytes(bytes(data))
    monkeypatch.setattr(backups, "SNAPSHOT_CHECK_BYTES", required + 10_000)
    payload, is_error = _call("ppt_deck", "create_snapshot", {"file_path": str(big)})
    assert is_error is False, payload
    monkeypatch.setattr(backups, "SNAPSHOT_CHECK_BYTES", 10_000_000)
    payload, is_error = _call("ppt_deck", "create_snapshot", {"file_path": str(big), "label": "checked"})
    assert is_error is True, payload  # with room in the budget the damaged part is found
    assert "cut off or damaged" in payload["error"]["message"]
