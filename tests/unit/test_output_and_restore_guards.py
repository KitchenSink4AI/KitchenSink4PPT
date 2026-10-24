"""Regression tests for backup-store, export-output and exit-ownership guards.

Each test's docstring gives the behaviour it pins and why. The two
COM-gated tests skip unless KS4_VERIFY_PPT_COM=1 is set in a window where
no PowerPoint is running; no other test here reaches PowerPoint.
"""

from __future__ import annotations

import hashlib
import os
from pathlib import Path

import pytest

from kitchensink4ppt import server

ROOT = Path(__file__).resolve().parents[2]


def _deck(path: Path, slides: int = 3, table: bool = True) -> Path:
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    for i in range(slides):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"Slide {i}"
        slide.placeholders[1].text_frame.text = f"body {i}"
    if table:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tbl = slide.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(4), Inches(1)).table
        for r in range(2):
            for c in range(2):
                tbl.cell(r, c).text = f"cell{r}{c}"
    prs.save(path)
    return path


def _sha(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _table_id(deck: Path, slide: int) -> int:
    info = server.get_slide_info(str(deck), slide, live="off")
    return next(s["id"] for s in info["shapes"] if "Table" in (s.get("name") or ""))


@pytest.mark.parametrize("slot_name", ["prev.pptx", "anchor.pptx", "new.pptx"])
def test_r1_restore_never_writes_into_the_backup_store(tmp_path, slot_name):
    """Restore never writes into the backup store. manage_backups
    action='restore' with file_path inside .ks4p-backups replaced the undo
    slot (or created a new file there) and nested a second backup store
    inside the first, because restore_backup never called
    refuse_protected_output."""
    deck = _deck(tmp_path / "deck.pptx")
    server.set_slide_hidden(str(deck), 0, True)
    store = tmp_path / ".ks4p-backups"
    other = _deck(tmp_path / "other.pptx", slides=5, table=False)
    before = {p: _sha(p) for p in store.rglob("*.pptx") if p.is_file()}
    target = store / "deck.pptx" / slot_name
    try:
        server.manage_backups("restore", file_path=str(target), source=str(other))
    except Exception:
        pass
    after = {p: _sha(p) for p in store.rglob("*.pptx") if p.is_file()}
    assert after == before


def test_r2_an_export_never_writes_through_a_hardlink_of_the_deck(tmp_path):
    """An export never writes through a hardlink of the deck.
    refuse_protected_output must not compare normcase(realpath) only: an
    output path that is a hardlink alias of the source deck would pass, and
    export_table would truncate the shared data, turning the presentation
    into a CSV. os.path.samefile (same volume serial and file index) is the
    identity check that catches it."""
    deck = _deck(tmp_path / "deck.pptx")
    alias = tmp_path / "alias.csv"
    os.link(deck, alias)
    before = _sha(deck)
    try:
        server.export_table(str(deck), 3, _table_id(deck, 3), path=str(alias), format="csv")
    except Exception:
        pass
    assert _sha(deck) == before


def test_r3_the_op_sweep_gate_exercises_each_op_with_real_arguments():
    """The op sweep gate exercises each op with real arguments.
    test_family_op_sweep_gate.py must not dispatch every op with args={}
    and accept BAD_ARGS for every op, since then no op body ever runs;
    MUST_SUCCEED must include create_presentation and
    get_presentation_view; and the output-guard half must exercise each
    op's output site rather than call the guard helper on three strings
    (which is how the restore and hardlink holes above stayed open)."""
    text = (ROOT / "tests" / "unit" / "test_family_op_sweep_gate.py").read_text(encoding="utf-8")
    assert "create_presentation" in text and "get_presentation_view" in text
    assert '"args": {}' not in text


_SLOW_OPEN = """
import importlib.abc, importlib.util, json, os, sys, time
class F(importlib.abc.MetaPathFinder):
    def find_spec(self, name, path, target=None):
        if name != "kitchensink4ppt.com.bridge":
            return None
        sys.meta_path.remove(self)
        spec = importlib.util.find_spec(name)
        sys.meta_path.insert(0, self)
        loader = spec.loader
        class L(importlib.abc.Loader):
            def create_module(self, s): return loader.create_module(s)
            def exec_module(self, m):
                loader.exec_module(m)
                orig = m.open_presentation
                def slow(*a, _o=orig, _m=m, **k):
                    r = _o(*a, **k)
                    # The PowerPoint processes this server started, as its own
                    # ownership journal records them (pid: creation time).
                    owned = {str(pid): ct for d in list(_m._OWNED_JOURNAL.values())
                             for pid, ct in d.items()}
                    out = os.environ["KS4_R6_OWNED"]
                    with open(out + ".tmp", "w", encoding="utf-8") as f:
                        json.dump({"server_pid": os.getpid(), "owned": owned}, f)
                    os.replace(out + ".tmp", out)
                    time.sleep(400)
                    return r
                m.open_presentation = slow
        spec.loader = L()
        return spec
sys.meta_path.insert(0, F())
"""


@pytest.mark.timeout(600)
@pytest.mark.skipif(os.environ.get("KS4_VERIFY_PPT_COM") != "1",
                    reason="needs a window with no PowerPoint running: set KS4_VERIFY_PPT_COM=1")
def test_r6_exit_never_ends_an_owned_instance_that_now_holds_someone_elses_presentation(tmp_path):
    """Exit never ends an owned instance that now holds someone else's
    presentation. The server's hidden PowerPoint gains a presentation the
    server did not open (the user starting PowerPoint lands in the
    singleton; modelled windowless here to keep the desktop silent). The
    product's own rule (_still_ours: Visible == 0 and no presentations)
    says the instance is no longer ours, so the exit path must not end it
    on PID, creation time, image and /Automation alone: TerminateProcess
    would lose that presentation without a prompt.

    Only the PowerPoint this test's server started is watched, attached to
    and, in the cleanup, quit or ended: the server's own ownership journal
    names it (a hook in the server writes it out), matched by pid AND
    creation time. The test attaches only while that is the only
    PowerPoint, quits it only while it is still invisible with no
    presentations (the product's own rule), and ends a journalled process
    only through a handle that shows the recorded creation time. A
    PowerPoint that appeared for any other reason (the user's, another
    session's) is never touched."""
    import json
    import subprocess
    import sys
    import threading
    import time
    import queue

    import pythoncom
    import win32com.client

    from kitchensink4ppt.com import bridge
    from test_family_surface_guards import _end_if_same

    before = bridge.powerpnt_pids()
    if before is None or before:
        pytest.skip("a PowerPoint is already running; this test only runs on a clean machine")
    hook = tmp_path / "hook"
    hook.mkdir()
    (hook / "sitecustomize.py").write_text(_SLOW_OPEN, encoding="utf-8")
    owned_file = tmp_path / "owned.json"
    deck = _deck(tmp_path / "deck.pptx", table=False)
    users = _deck(tmp_path / "users.pptx", table=False)
    env = dict(os.environ, PYTHONUNBUFFERED="1", KS4P_UPDATE_CHECK="off", PYTHONPATH=str(hook),
               KS4_R6_OWNED=str(owned_file))
    proc = subprocess.Popen([sys.executable, "-X", "utf8", str(ROOT / "scripts" / "_bootstrap.py")],
                            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                            cwd=str(ROOT), env=env, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    q: queue.Queue = queue.Queue()
    threading.Thread(target=lambda: [q.put(line) for line in proc.stdout], daemon=True).start()

    def send(obj):
        proc.stdin.write((json.dumps(obj) + "\n").encode())
        proc.stdin.flush()

    ours: dict[int, int] = {}
    alive: set = set()
    app = pres = None
    attached = False
    may_end = True
    try:
        send({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {
            "protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "r6", "version": "0"}}})
        q.get(timeout=90)
        send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        send({"jsonrpc": "2.0", "id": 2, "method": "tools/call", "params": {"name": "ppt_deck", "arguments": {
            "op": "validate", "args": {"file_path": str(deck)}}}})
        end = time.monotonic() + 120
        while time.monotonic() < end and not owned_file.exists():
            time.sleep(1)
        assert owned_file.exists(), "the server never opened the deck in a PowerPoint"
        # Only this test's server has KS4_R6_OWNED, so the file is its own journal.
        record = json.loads(owned_file.read_text(encoding="utf-8"))
        ours = {int(pid): int(ct) for pid, ct in record["owned"].items()}
        if not ours:
            pytest.skip("another PowerPoint started first and the server attached to it; "
                        "this test needs the server to start its own")
        assert all(bridge._creation_time(pid) == ct for pid, ct in ours.items()), ours
        time.sleep(5)
        # PowerPoint is one instance per machine: attach only while the only
        # PowerPoint is the one the server started, so the presentation below
        # can land in no one else's.
        if (bridge.powerpnt_pids() or set()) != set(ours):
            pytest.skip("another PowerPoint is running beside the server's; this test opens nothing")
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("PowerPoint.Application")
        attached = True
        if bridge._powerpoint_app_pid(app) not in ours:
            pytest.skip("the PowerPoint this test reached is not the one its server started")
        pres = app.Presentations.Open(str(users), ReadOnly=False, Untitled=False, WithWindow=False)
        proc.stdin.close()
        proc.wait(300)
        time.sleep(10)
        alive = {pid for pid, ct in ours.items() if bridge._creation_time(pid) == ct}
    finally:
        try:
            if pres is not None:
                pres.Close()  # the presentation this test opened
        except Exception:
            pass
        pres = None
        if attached:
            # The instance may be quit, or ended below, only while it is still
            # what the server started: invisible, with no presentations.
            may_end = False
            try:
                if app is not None and int(app.Visible) == 0 and int(app.Presentations.Count) == 0:
                    may_end = True
                    app.Quit()
            except Exception:
                pass
        app = None
        if proc.poll() is None:
            proc.kill()
        deadline = time.monotonic() + 20
        while time.monotonic() < deadline and any(
                bridge._creation_time(pid) == ct for pid, ct in ours.items()):
            time.sleep(1)
        if may_end:
            for pid, ct in ours.items():
                _end_if_same(pid, ct)
    assert alive == set(ours), "exit ended a PowerPoint that held a presentation this server did not open"


@pytest.mark.timeout(300)
@pytest.mark.skipif(os.environ.get("KS4_VERIFY_PPT_COM") != "1",
                    reason="needs PowerPoint in a clean window: set KS4_VERIFY_PPT_COM=1")
def test_r7_an_extensionless_com_pdf_output_never_replaces_a_users_file(tmp_path):
    """An extensionless COM PDF output never replaces a user's file.
    export_pdf engine='com' output='report' passed every guard (the guard
    saw 'report'), PowerPoint's SaveAs appended .pdf and replaced the
    user's existing report.pdf with no backup, and the tool then answered
    BAD_PARAMS 'reported success but produced no PDF' (the caller was
    blamed for a file the tool overwrote)."""
    from kitchensink4ppt.com import bridge

    if bridge.powerpnt_pids():
        pytest.skip("a PowerPoint is already running")
    deck = _deck(tmp_path / "deck.pptx", table=False)
    user_pdf = tmp_path / "report.pdf"
    user_pdf.write_bytes(b"%PDF-1.4 the user's own report\n")
    before = _sha(user_pdf)
    try:
        server.export_pdf(str(deck), output=str(tmp_path / "report"), engine="com")
    except Exception:
        pass
    assert _sha(user_pdf) == before
