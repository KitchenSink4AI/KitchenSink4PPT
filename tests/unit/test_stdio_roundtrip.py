"""Raw stdio integration for the immutable grouped surface:

initialize -> exact tools/list -> grouped read -> grouped edit -> stable list.

No fastmcp client library on purpose: the bytes on the wire are the product.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parents[2]


class _Server:
    """Line-delimited JSON-RPC over the spawned server's stdio."""

    def __init__(
        self, env_extra: dict | None = None, *, drain_stderr: bool = True
    ):
        env = dict(os.environ)
        env["PYTHONUNBUFFERED"] = "1"
        env.update(env_extra or {})
        self.proc = subprocess.Popen(
            [sys.executable, "-X", "utf8", str(REPO / "scripts" / "_bootstrap.py")],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            cwd=str(REPO),
            env=env,
            text=True,
            encoding="utf-8",
        )
        self._q: queue.Queue = queue.Queue()
        self.notifications: list[dict] = []
        self.stderr_lines: list[str] = []
        self.stray_stdout: list[str] = []
        self.stderr_ready = threading.Event()
        self._reader = threading.Thread(target=self._pump, daemon=True)
        self._reader.start()
        self._stderr_reader = None
        if drain_stderr:
            self._stderr_reader = threading.Thread(
                target=self._pump_stderr, daemon=True
            )
            self._stderr_reader.start()
        self._id = 0

    def _pump(self):
        for line in self.proc.stdout:
            line = line.strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
            except json.JSONDecodeError:
                self.stray_stdout.append(line)  # non-protocol output
                continue
            self._q.put(msg)

    def _pump_stderr(self):
        for line in self.proc.stderr:
            self.stderr_lines.append(line.rstrip("\r\n"))
            self.stderr_ready.set()

    def _send(self, msg: dict):
        self.proc.stdin.write(json.dumps(msg) + "\n")
        self.proc.stdin.flush()

    def request(self, method: str, params: dict | None = None, timeout=30):
        self._id += 1
        rid = self._id
        self._send(
            {"jsonrpc": "2.0", "id": rid, "method": method,
             "params": params or {}}
        )
        while True:
            try:
                msg = self._q.get(timeout=timeout)
            except queue.Empty:
                raise AssertionError(
                    f"{method} produced no response within {timeout}s; "
                    f"exit={self.proc.poll()}; stderr={self.stderr_lines[-8:]}"
                ) from None
            if msg.get("id") == rid:
                assert "error" not in msg, f"{method} failed: {msg['error']}"
                return msg["result"]
            if "method" in msg and "id" not in msg:
                self.notifications.append(msg)

    def notify(self, method: str, params: dict | None = None):
        self._send({"jsonrpc": "2.0", "method": method, "params": params or {}})

    def drain_notifications(self, timeout=2.0):
        try:
            while True:
                msg = self._q.get(timeout=timeout)
                if "method" in msg and "id" not in msg:
                    self.notifications.append(msg)
        except queue.Empty:
            pass

    def close(self):
        try:
            self.proc.stdin.close()
        except OSError:
            pass
        try:
            self.proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            self.proc.kill()
            self.proc.wait(timeout=10)


def _handshake(srv: _Server):
    result = srv.request(
        "initialize",
        {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "stdio-roundtrip-test", "version": "0"},
        },
        timeout=60,
    )
    assert result["serverInfo"]["name"] == "kitchensink4ppt"
    srv.notify("notifications/initialized")
    return result


def _tool_names(srv: _Server) -> list[str]:
    result = srv.request("tools/list")
    return [t["name"] for t in result["tools"]]


def _tool_list_bytes(srv: _Server) -> bytes:
    result = srv.request("tools/list")
    return json.dumps(
        result["tools"], ensure_ascii=False, separators=(",", ":")
    ).encode("utf-8")


def _call(
    srv: _Server, name: str, arguments: dict, *, expect_error: bool = False
) -> dict:
    result = srv.request("tools/call", {"name": name, "arguments": arguments})
    # Structured refusals ALSO set the MCP-level isError flag (the JSON
    # envelope stays in the content).
    if expect_error:
        assert result.get("isError"), f"{name} should refuse: {result}"
    else:
        assert not result.get("isError"), f"{name} errored: {result}"
    if "structuredContent" in result and result["structuredContent"]:
        return result["structuredContent"]
    text = "".join(
        c.get("text", "") for c in result.get("content", [])
    )
    return json.loads(text) if text else {}


@pytest.mark.timeout(120)
def test_batch_stdio_client_receives_responses_before_clean_exit():
    """Closing the request pipe is ordinary EOF, not a forced disconnect.

    The Windows orphan-prevention sentinel may take over only while an
    operation worker is active; otherwise it must let FastMCP flush the
    initialize and tools/list responses used by one-shot clients.
    """
    spec = importlib.util.spec_from_file_location(
        "ppt_measure_surface", REPO / "scripts" / "measure_surface.py"
    )
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    report = module.measure_stdio()
    assert report["tools"] == 9
    assert report["ops"] == 145


@pytest.mark.timeout(120)
def test_stdio_roundtrip(make_deck, tmp_path):
    deck_src = make_deck("stdio_src.pptx")
    deck = tmp_path / "stdio_deck.pptx"
    shutil.copy2(deck_src, deck)
    before = hashlib.md5(deck.read_bytes()).hexdigest()

    srv = _Server()
    try:
        initialized = _handshake(srv)
        from kitchensink4ppt.surface.plan import ALL_TOOLS

        names = _tool_names(srv)
        assert names == list(ALL_TOOLS)
        assert not initialized.get("capabilities", {}).get("tools", {}).get(
            "listChanged", False
        )

        info = _call(
            srv, "ppt_read", {
                "op": "get_presentation_info",
                "args": {"file_path": str(deck)},
            },
        )
        assert info.get("slide_count", 0) >= 1

        shape = _call(
            srv, "ppt_graphics", {
                "op": "insert_shape",
                "args": {
                    "file_path": str(deck), "slide": 0,
                    "shape_type": "rect", "x": 1, "y": 1,
                    "w": 2, "h": 1, "text": "stdio proof",
                },
            },
        )
        assert shape["ok"] is True
        assert shape["changed"]["shape_id"]
        assert hashlib.md5(deck.read_bytes()).hexdigest() != before

        refusal = _call(
            srv, "ppt_edit", {
                "op": "delete_slide",
                "args": {"file_path": str(deck), "slide": 99},
            }, expect_error=True,
        )
        assert refusal["ok"] is False
        assert refusal["error"]["code"] == "NOT_FOUND"

        assert _tool_names(srv) == names
        srv.drain_notifications(timeout=0.2)
        assert not any(
            n.get("method") == "notifications/tools/list_changed"
            for n in srv.notifications
        )
    finally:
        srv.close()


@pytest.mark.parametrize("probe_count", [20, 120])
@pytest.mark.timeout(180)
def test_bad_preinitialize_requests_cannot_fill_undrained_stderr(probe_count):
    """20 discover probes and 120 mixed probes cannot wedge stdio."""
    srv = _Server(drain_stderr=False)
    try:
        wanted = {900 + i for i in range(probe_count)}
        for i in range(probe_count):
            discover = probe_count == 20 or i % 2 == 0
            message = {
                "jsonrpc": "2.0",
                "id": 900 + i,
                "method": "server/discover" if discover else "acme/whatever",
                "params": (
                    {
                        "_meta": {
                            "io.modelcontextprotocol/protocolVersion": "2026-07-28",
                            "io.modelcontextprotocol/clientCapabilities": {},
                            "io.modelcontextprotocol/clientInfo": {
                                "name": "probe",
                                "version": "0",
                            },
                        }
                    }
                    if discover
                    else {"n": i}
                ),
            }
            srv._send(message)

        replies = {}
        deadline = time.monotonic() + 90
        while wanted - replies.keys() and time.monotonic() < deadline:
            try:
                response = srv._q.get(timeout=0.25)
            except queue.Empty:
                continue
            if response.get("id") in wanted:
                replies[response["id"]] = response

        assert wanted <= replies.keys(), (
            f"{len(wanted - replies.keys())} of {probe_count} probes were "
            "left unanswered behind stderr backpressure"
        )
        assert all("error" in reply for reply in replies.values())
        initialized = srv.request(
            "initialize",
            {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "stderr-bound", "version": "0"},
            },
            timeout=90,
        )
        assert initialized["serverInfo"]["name"] == "kitchensink4ppt"
        srv.notify("notifications/initialized")

        # EOF must not wait for a daemon writer blocked in os.write to the
        # still-unread client pipe.
        srv.proc.stdin.close()
        assert srv.proc.wait(timeout=30) == 0
    finally:
        srv.close()


OLD_BUNDLE_DEFAULTS = {
    "KS4P_MODE": "lite",
    "KS4P_PACK_POLICY": "auto",
    "KS4P_ALL_TOOLS": "false",
    "KS4P_LOCK_TOOLS": "false",
}


@pytest.mark.timeout(60)
@pytest.mark.parametrize("env_extra,named", [
    ({"KS4P_MODE": "full", "KS4P_PACK_POLICY": "locked",
      "KS4P_ALL_TOOLS": "1", "KS4P_LOCK_TOOLS": "1"}, "KS4P_PACK_POLICY"),
    ({"KS4P_MODE": "lite", "KS4P_LOCK_TOOLS": "true"}, "KS4P_LOCK_TOOLS"),
    ({"KS4P_MODE": "lite,graphics"}, "KS4P_MODE"),
])
def test_stdio_old_settings_that_locked_the_tools_refuse_startup(env_extra, named):
    srv = _Server(env_extra={"KS4P_STAR_NUDGE": "off", **env_extra})
    try:
        assert srv.proc.wait(timeout=60) == 2
        srv._stderr_reader.join(timeout=5)
        lines = [line for line in srv.stderr_lines if line.strip()]
        assert len(lines) == 1
        assert named in lines[0]
        assert "KS4P_READ_ONLY=1" in lines[0] and "KS4P_FORBID=" in lines[0]
        assert srv.stray_stdout == []
    finally:
        srv.close()


@pytest.mark.timeout(120)
def test_stdio_old_defaults_start_normally_with_one_notice_and_the_full_tool_set():
    from kitchensink4ppt.settings import LEGACY_ENV_NAMES, legacy_notice
    from kitchensink4ppt.surface.plan import ALL_TOOLS

    srv = _Server(env_extra={"KS4P_STAR_NUDGE": "off", **OLD_BUNDLE_DEFAULTS})
    try:
        # The notice arrives on stderr before the client has said anything,
        # and stdout stays empty until the handshake.
        assert srv.stderr_ready.wait(timeout=60)
        assert srv._q.empty()
        assert srv.stray_stdout == []
        _handshake(srv)
        assert _tool_names(srv) == list(ALL_TOOLS)
        info = _call(srv, "ppt_read", {"op": "get_server_info", "args": {}})
        assert info["surface"]["legacy_settings_ignored"] == list(LEGACY_ENV_NAMES)
        assert info["surface"]["read_only"] is False
        srv.proc.stdin.close()
        assert srv.proc.wait(timeout=30) == 0
        srv._stderr_reader.join(timeout=5)
        lines = [line for line in srv.stderr_lines if line.strip()]
        assert lines == [legacy_notice(LEGACY_ENV_NAMES)]
        assert srv.stray_stdout == []
    finally:
        srv.close()


@pytest.mark.timeout(120)
def test_stdio_blank_old_settings_start_silently():
    srv = _Server(env_extra={
        "KS4P_STAR_NUDGE": "off", "KS4P_MODE": "", "KS4P_PACK_POLICY": "  ",
    })
    try:
        _handshake(srv)
        info = _call(srv, "ppt_read", {"op": "get_server_info", "args": {}})
        assert info["surface"]["legacy_settings_ignored"] == []
        srv.proc.stdin.close()
        assert srv.proc.wait(timeout=30) == 0
        srv._stderr_reader.join(timeout=5)
        assert [line for line in srv.stderr_lines if line.strip()] == []
    finally:
        srv.close()


@pytest.mark.timeout(120)
def test_stdio_mode_lite_does_not_undo_read_only_or_forbid():
    srv = _Server(env_extra={
        "KS4P_STAR_NUDGE": "off", "KS4P_MODE": "lite", "KS4P_READ_ONLY": "1",
    })
    try:
        _handshake(srv)
        assert _tool_names(srv) == ["ppt_read", "ppt_inspect", "ppt_help"]
    finally:
        srv.close()
    srv = _Server(env_extra={
        "KS4P_STAR_NUDGE": "off", "KS4P_MODE": "lite", "KS4P_FORBID": "ppt_graphics",
    })
    try:
        _handshake(srv)
        names = _tool_names(srv)
        assert "ppt_graphics" not in names and "ppt_edit" in names
    finally:
        srv.close()


@pytest.mark.timeout(60)
def test_stdio_ppt_app_forbid_removes_live_and_com_routes():
    srv = _Server(env_extra={"KS4P_FORBID": "ppt_app"})
    try:
        _handshake(srv)
        listed = srv.request("tools/list")["tools"]
        enums = {
            tool["name"]: set(tool["inputSchema"]["properties"]["op"].get("enum", []))
            for tool in listed
            if tool["name"] != "ppt_help"
        }
        visible = set().union(*enums.values())
        for name in (
            "get_text", "insert_shape", "live_save", "powerpoint_status",
            "zombie_check", "export_pdf", "export_handout",
            "export_slide_image", "validate",
        ):
            assert name not in visible
        assert "extract_text" in visible
        assert "create_presentation" in visible
        for tool, op in (("ppt_read", "get_text"), ("ppt_deck", "export_pdf")):
            refusal = _call(
                srv, tool, {"op": op, "args": {}}, expect_error=True
            )
            assert refusal["ok"] is False
            assert refusal["error"]["code"] == "NOT_AVAILABLE"
    finally:
        srv.close()


@pytest.mark.timeout(90)
def test_stdio_unknown_forbid_name_refuses_startup():
    srv = _Server(env_extra={"KS4P_FORBID": "not_a_real_tool"})
    try:
        assert srv.proc.wait(timeout=60) == 2
        srv._stderr_reader.join(timeout=5)
        stderr = "\n".join(srv.stderr_lines)
        assert "KS4P_FORBID" in stderr
        assert "not_a_real_tool" in stderr
        assert "did not start" in stderr
    finally:
        srv.close()


@pytest.mark.timeout(60)
def test_stdio_forbid_removes_tool_at_startup():
    srv = _Server(env_extra={"KS4P_FORBID": "ppt_graphics"})
    try:
        _handshake(srv)
        names = _tool_names(srv)
        assert "ppt_graphics" not in names
        assert "ppt_read" in names
    finally:
        srv.close()


@pytest.mark.parametrize(
    ("env_extra", "missing_tool", "hidden_op"),
    [
        ({"KS4P_FORBID": "ppt_graphics"}, "ppt_graphics", "insert_shape"),
        ({"KS4P_READ_ONLY": "1"}, "ppt_edit", "delete_slide"),
    ],
)
@pytest.mark.timeout(60)
def test_stdio_narrowing_is_enforced_inside_calls(
    env_extra, missing_tool, hidden_op
):
    """A stale client cannot bypass frozen launch policy by naming an op."""
    srv = _Server(env_extra=env_extra)
    try:
        _handshake(srv)
        names = _tool_names(srv)
        assert missing_tool not in names
        assert "ppt_read" in names

        refusal = _call(
            srv,
            "ppt_read",
            {"op": hidden_op, "args": {}},
            expect_error=True,
        )
        assert refusal["ok"] is False
        assert refusal["error"]["code"] == "NOT_AVAILABLE"
        assert refusal["error"]["details"]["op"] == hidden_op
        assert refusal["error"]["details"]["decided_by"]
        assert _tool_names(srv) == names
    finally:
        srv.close()


@pytest.mark.timeout(60)
def test_stdio_list_is_byte_identical_after_calling_each_fixed_tool_once():
    srv = _Server()
    try:
        initialized = _handshake(srv)
        assert not initialized.get("capabilities", {}).get("tools", {}).get(
            "listChanged", False
        )
        before = _tool_list_bytes(srv)
        names = json.loads(before)
        for item in names:
            name = item["name"]
            if name == "ppt_help":
                _call(srv, name, {})
            else:
                refusal = _call(
                    srv,
                    name,
                    {"op": "__fixed_list_probe__", "args": {}},
                    expect_error=True,
                )
                assert refusal["error"]["code"] == "NOT_FOUND", name
            assert _tool_list_bytes(srv) == before, name
        srv.drain_notifications(timeout=0.2)
        assert not any(
            notice.get("method") == "notifications/tools/list_changed"
            for notice in srv.notifications
        )
    finally:
        srv.close()
