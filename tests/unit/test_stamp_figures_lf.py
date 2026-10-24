"""scripts/stamp_figures.py writes shipped.py with LF line endings on every
platform, as the repository keeps it. On Windows the write used the
platform's line ending, so every stamp turned the whole file into CRLF and
the release had to put LF back by hand.

The count and the figures are stubbed: the test checks the write, not the
collection.
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def _stamp_module(monkeypatch):
    # The script puts scripts/ and src/ on sys.path; keep that to this test.
    monkeypatch.setattr(sys, "path", list(sys.path))
    spec = importlib.util.spec_from_file_location(
        "stamp_figures_under_test", ROOT / "scripts" / "stamp_figures.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_the_stamp_writes_shipped_py_with_lf(tmp_path, monkeypatch):
    module = _stamp_module(monkeypatch)
    shipped = tmp_path / "shipped.py"
    shipped.write_bytes(b'"""Made-up figures."""\n\nTESTS = 1\n')
    monkeypatch.setattr(module, "SHIPPED", shipped)
    monkeypatch.setattr(module, "tool_figures", lambda: {
        "tools": 9, "operations": 143, "approx_tokens": 1000, "instructions_chars": 100})
    monkeypatch.setattr(module, "test_count", lambda: 2)
    monkeypatch.setattr(sys, "argv", ["stamp_figures.py", "--write"])
    assert module.main() == 0
    data = shipped.read_bytes()
    assert b"TESTS = 2\n" in data
    assert b"\r\n" not in data
