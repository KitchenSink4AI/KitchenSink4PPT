"""Every PowerPoint COM route that can write a file is accounted for.

The Excel editor's COM tier once saved the user's workbook without taking the
cross-process write lock. The question for this editor is the same: which
COM calls can write a file, and which of them write the user's deck?

The answer, held here so a new route cannot appear unaudited:

- The four bridge routes (PDF, handout, slide images, opens-clean check) open
  the deck ReadOnly and never save it. They write only new output files:
  the PDF routes refuse an existing output unless the call sends
  overwrite=true, and always refuse the source itself and a backup folder;
  the image route writes slideN.png files in its own folder.
- ``live_save`` is the one route that writes the user's deck. It saves the
  deck that is open in the user's PowerPoint, under the process-wide COM lock
  and the cross-process live lock (com/xproc.py), and there is no copy to
  promote over a newer file: PowerPoint holds the deck open, so file-mode
  writers are refused while it is (core/package.py _check_lock).
- Every other live body edits the open presentation in memory and saves
  nothing.

When a COM call that can write appears anywhere else, or an existing route
changes how it opens a deck, these tests fail until the route is audited and
added below with its reason.
"""

from __future__ import annotations

import ast
import contextlib
import sys
from pathlib import Path

import pytest

SRC = Path(__file__).resolve().parents[2] / "src" / "kitchensink4ppt"

#: Calls on a PowerPoint object that put bytes in a file.
WRITING_CALLS = frozenset({
    "Save", "SaveAs", "SaveAs2", "SaveCopyAs", "SaveCopy",
    "Export", "ExportAsFixedFormat", "PrintOut", "ExecuteMso", "Publish",
})

#: (file, enclosing function, call) -> why that call cannot lose the user's work.
AUDITED = {
    ("com/bridge.py", "com_export_pdf", "SaveCopyAs"):
        "writes a new PDF, or replaces one through a temporary file only on overwrite=true; "
        "the deck is open ReadOnly; the output may not be the deck or sit in a backup folder",
    ("com/bridge.py", "com_export_pdf", "ExportAsFixedFormat"):
        "the same PDF, by the fallback call",
    ("com/bridge.py", "com_export_handout", "ExportAsFixedFormat"):
        "writes a new handout PDF; the deck is open ReadOnly; the same output rules",
    ("com/bridge.py", "com_export_slide_images", "Export"):
        "writes slideN.png files in an output folder, refusing existing ones unless "
        "overwrite=true; the deck is open ReadOnly",
    ("com/live_ops.py", "live_save.body", "Save"):
        "the one route that writes the user's deck, inside live.run_live: the COM lock and "
        "the cross-process live lock",
}

#: Where a COM call could be reached from: the packages that hold COM code.
SCANNED = ("com", "ops", "surface", "core")
SCANNED_FILES = ("server.py",)


def _sources():
    for folder in SCANNED:
        for path in sorted((SRC / folder).rglob("*.py")):
            yield path
    for name in SCANNED_FILES:
        yield SRC / name


class _Calls(ast.NodeVisitor):
    def __init__(self):
        self.stack: list[str] = []
        self.found: list[tuple[str, str, int]] = []

    def _scope(self, node):
        self.stack.append(node.name)
        self.generic_visit(node)
        self.stack.pop()

    visit_FunctionDef = visit_AsyncFunctionDef = visit_ClassDef = _scope

    def visit_Call(self, node: ast.Call):
        if isinstance(node.func, ast.Attribute) and node.func.attr in WRITING_CALLS:
            self.found.append((".".join(self.stack), node.func.attr, node.lineno))
        self.generic_visit(node)


def _found():
    out = {}
    for path in _sources():
        visitor = _Calls()
        visitor.visit(ast.parse(path.read_text(encoding="utf-8")))
        rel = path.relative_to(SRC).as_posix()
        for scope, call, line in visitor.found:
            out[(rel, scope, call)] = line
    return out


def test_every_com_call_that_can_write_a_file_is_audited():
    found = _found()
    unaudited = sorted(set(found) - set(AUDITED))
    assert not unaudited, (
        "a COM call that can write a file is not in AUDITED: "
        + ", ".join(f"{f}:{found[(f, s, c)]} {s} .{c}()" for f, s, c in unaudited)
        + ". Decide whether it can write the user's deck; if it can, it takes the deck's "
        "write lock or runs inside live.run_live, and then it is added here with its reason."
    )
    gone = sorted(set(AUDITED) - set(found))
    assert not gone, f"audited routes that no longer exist (remove them): {gone}"


def test_the_bridge_opens_every_deck_read_only():
    """open_presentation(read_only=False) is the only way the bridge could
    hold a deck for writing; nothing calls it that way."""
    callers = []
    for path in (SRC / "com").rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                    and node.func.id == "open_presentation"):
                flag = {kw.arg: kw.value for kw in node.keywords}.get("read_only")
                callers.append((path.name, node.lineno,
                                isinstance(flag, ast.Constant) and flag.value is True))
    assert len(callers) >= 4, callers
    assert all(read_only for _name, _line, read_only in callers), callers


def test_live_save_is_the_only_route_that_saves_and_it_goes_through_run_live():
    tree = ast.parse((SRC / "com" / "live_ops.py").read_text(encoding="utf-8"))
    saves = [fn.name for fn in ast.walk(tree)
             if isinstance(fn, ast.FunctionDef)
             and any(isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
                     and n.func.attr in {"Save", "SaveAs", "SaveAs2", "SaveCopyAs"}
                     for n in ast.walk(fn))]
    assert set(saves) == {"live_save", "body"}, saves  # body is live_save's nested function
    live_save = next(fn for fn in ast.walk(tree)
                     if isinstance(fn, ast.FunctionDef) and fn.name == "live_save")
    runs = [n for n in ast.walk(live_save)
            if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
            and n.func.attr == "run_live"]
    assert len(runs) == 1


class _FakePres:
    FullName = "C:/decks/open.pptx"
    Saved = True

    def __init__(self, observe):
        self._observe = observe

    def Save(self):
        self._observe()


class _FakeSession:
    state_restore_failed: list = []
    unsaved_before_our_edits = None
    had_unsaved_user_changes = False
    deck_key = None

    def __init__(self, observe):
        self.pres = _FakePres(observe)

    def result(self, body):
        return dict(body)


@pytest.fixture()
def live_lock_dir(tmp_path, monkeypatch):
    """Redirect the live lock so the machine-wide one is never touched."""
    from kitchensink4ppt.com import xproc

    folder = tmp_path / "live-locks"
    monkeypatch.setattr(xproc, "_lock_dir", lambda: folder)
    return folder


@pytest.mark.skipif(sys.platform != "win32", reason="the live layer is Windows-only")
def test_live_save_saves_only_while_it_holds_both_locks(live_lock_dir, monkeypatch):
    """The save happens inside the process-wide COM lock and the cross-process
    live lock, and both are let go afterwards."""
    from kitchensink4ppt.com import live, live_ops, xproc
    from kitchensink4ppt.com import serial as com_serial

    seen = []

    def observe():
        holder = xproc.holder_info()
        snapshot = com_serial.lock_snapshot()
        seen.append({
            "live_lock_ours": bool(holder and holder["ours"]),
            "com_lock_held": snapshot["held"],
            "com_op": (snapshot.get("current_op") or {}).get("name"),
        })

    @contextlib.contextmanager
    def fake_session(path, tool_name, mutating=True, edits=True):
        yield _FakeSession(observe)

    # live_session takes both locks and then enters _live_session_locked, which
    # attaches to PowerPoint; the attach is the only part replaced.
    monkeypatch.setattr(live, "_live_session_locked", fake_session)
    result = live_ops.live_save("C:/decks/open.pptx")
    assert result["saved"] == "C:/decks/open.pptx"
    assert seen == [{"live_lock_ours": True, "com_lock_held": True, "com_op": "live:live_save"}]
    assert xproc.holder_info() is None
    assert com_serial.lock_snapshot()["held"] is False
