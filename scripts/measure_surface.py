"""Measure the fixed PowerPoint family surface from its startup plan.

The script builds the same immutable plan passed to ``create_server`` and
reports each MCP tool's operation count, description size, schema size, and
approximate serialized token bill.

Run:  python -X utf8 scripts/measure_surface.py [--json]
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "scripts"))

import _bootstrap  # noqa: E402


# Ceiling for the whole tool list, in approximate tokens: the measured figure
# plus five percent. Raising it takes a new measurement.
SURFACE_TOKEN_CEILING = 9_022


def approx_tokens(tool: dict) -> int:
    return round(
        len(json.dumps(tool, ensure_ascii=False, separators=(",", ":"))) / 4
    )


def _report(tools: list[dict], *, instructions: str, version: str) -> dict:
    rows: list[dict[str, object]] = []
    for item in tools:
        schema = item.get("inputSchema") or item.get("input_schema") or {}
        wire = {
            "name": item["name"],
            "title": item.get("title"),
            "description": item.get("description") or "",
            "inputSchema": schema,
            "annotations": item.get("annotations") or {},
        }
        enum = schema.get("properties", {}).get("op", {}).get("enum", [])
        rows.append({
            "tool": item["name"],
            "description_chars": len(wire["description"]),
            "schema_bytes": len(json.dumps(
                schema, ensure_ascii=False, separators=(",", ":")
            ).encode("utf-8")),
            "approx_tokens": approx_tokens(wire),
            "ops": 0 if item["name"] == "ppt_help" else len(enum),
        })
    report = {
        "tools": len(rows),
        "ops": sum(int(row["ops"]) for row in rows),
        "approx_tokens": sum(int(row["approx_tokens"]) for row in rows),
        "approx_token_ceiling": SURFACE_TOKEN_CEILING,
        "instructions_chars": len(instructions),
        "server_version": version,
        "rows": rows,
    }
    if report["approx_tokens"] > SURFACE_TOKEN_CEILING:
        raise RuntimeError(
            f"surface load {report['approx_tokens']} exceeds the pinned "
            f"ceiling {SURFACE_TOKEN_CEILING}; record and re-measure the change"
        )
    return report


def measure(env: dict[str, str] | None = None) -> dict:
    _bootstrap.use_community()
    from kitchensink4ppt import registry, server
    from kitchensink4ppt.settings import Settings
    from kitchensink4ppt.surface import plan

    settings = Settings() if env is None else Settings.from_env(env)
    surface = plan.build(settings, registry.public_ops())
    mcp = server.create_server(settings)
    tools: list[dict] = []
    for item in surface.tools.values():
        tools.append({
            "name": item.name,
            "title": item.title,
            "description": item.description,
            "inputSchema": item.input_schema,
            "annotations": item.annotations,
        })
    return _report(
        tools, instructions=mcp.instructions or "", version=str(mcp.version)
    )


def measure_stdio() -> dict:
    """Measure the actual JSON-RPC initialize + tools/list wire response."""
    messages = [
        {
            "jsonrpc": "2.0", "id": 1, "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05", "capabilities": {},
                "clientInfo": {"name": "measure-surface", "version": "1"},
            },
        },
        {"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}},
        {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}},
    ]
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(ROOT / "scripts" / "_bootstrap.py")],
        input="".join(json.dumps(message) + "\n" for message in messages),
        capture_output=True,
        text=True,
        encoding="utf-8",
        cwd=ROOT,
        # Fresh imports on the certifying Windows host can take about 30 s
        # under antivirus/indexing load. This is a measurement harness bound,
        # not the per-operation deadline enforced by the server.
        timeout=90,
        creationflags=flags,
        check=False,
    )
    if proc.returncode:
        raise SystemExit(f"stdio server exited {proc.returncode}: {proc.stderr[:1000]}")
    responses = {}
    for line in proc.stdout.splitlines():
        try:
            message = json.loads(line)
        except json.JSONDecodeError:
            continue
        if "id" in message:
            responses[message["id"]] = message
    initialized = responses[1]["result"]
    listed = responses[2]["result"]["tools"]
    return _report(
        listed,
        instructions=initialized.get("instructions", ""),
        version=str(initialized["serverInfo"]["version"]),
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--stdio", action="store_true")
    args = parser.parse_args()
    report = measure_stdio() if args.stdio else measure()
    if args.json:
        print(json.dumps(report))
        return
    print(f"{'tool':<18} {'ops':>4} {'desc':>6} {'schema':>7} {'~tokens':>8}")
    for row in report["rows"]:
        print(
            f"{row['tool']:<18} {row['ops']:>4} "
            f"{row['description_chars']:>6} {row['schema_bytes']:>7} "
            f"{row['approx_tokens']:>8}"
        )
    print(
        f"TOTAL: {report['tools']} tools, {report['ops']} ops, "
        f"{report['approx_tokens']} approximate tokens; instructions "
        f"{report['instructions_chars']} chars"
    )


if __name__ == "__main__":
    main()
