"""Validate installed edition wheels, MCP wire identity and bundle contents.

Community ships app plus policy wheels; Business ships core, policy and launcher
rather than one. The properties being proved are the same as before the
split: the two editions expose an identical tool surface, the Business
artifacts carry no Community guidance, the console and stdio entry points
start, and nothing reaches the network. Two are new: the core payload is
identical under both grants, and the licensing text still lands ahead of the
product prose, which is PowerPoint's own composition order.
"""
from __future__ import annotations
import argparse
import asyncio
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

import adapter_source

CORE_VERSION = __import__("tomllib").loads(
    (ROOT / "pyproject.toml").read_text("utf-8"))["project"]["version"]

#: A distribution must not ship a wheel granted under the other license.
FORBIDDEN_LICENSE = {"community": "LicenseRef-KitchenSink-Commercial",
                     "business": "AGPL-3.0-only"}

TEMPLATE_MEMBER = "kitchensink4ppt/templates/default.pptx"
TEMPLATE_SHA256 = (
    "eb33715284d0a4838d3c1ffaa1245df153fef590ea5e77fa8f9d2fe4bc6867e6"
)

RETIRED_POLICY_ENV = (
    "KS4P_MODE",
    "KS4P_PACK_POLICY",
    "KS4P_ALL_TOOLS",
    "KS4P_LOCK_TOOLS",
)


def clean_runtime_env(source: dict[str, str] | None = None) -> dict[str, str]:
    """Return a launch environment without retired pre-2.0 policy controls."""
    clean = dict(os.environ if source is None else source)
    for name in RETIRED_POLICY_ENV:
        clean.pop(name, None)
    clean.pop("KS4P_LICENSE_FILE", None)
    clean.update({
        "KS4P_UPDATE_CHECK": "off",
        "PYTHONDONTWRITEBYTECODE": "1",
    })
    return clean

CHECK = r"""
import asyncio, builtins, json, os, sys, zipfile
from pathlib import Path
sys.path.insert(0, sys.argv[1])
edition = sys.argv[2]
root = Path(sys.argv[1]).resolve()

# The installed runtime must never reach the development-only oracle,
# even when the interpreter running this smoke happens to have it installed.
real_import = builtins.__import__
def no_python_pptx(name, *args, **kwargs):
    if name == "pptx" or name.startswith("pptx."):
        raise AssertionError("installed artifact imported python-pptx")
    return real_import(name, *args, **kwargs)
builtins.__import__ = no_python_pptx

# The launcher injects the policy; nothing else in this process chooses one.
if edition == "community":
    import ks4ppt_community_launch as launcher
else:
    import ks4ppt_business_launch as launcher
build_policy = launcher.build_policy
from kitchensink4ppt import license_policy
license_policy.install(build_policy())

from kitchensink4ppt import server
from fastmcp import Client

assert Path(server.__file__).resolve().is_relative_to(root), server.__file__
os.environ["KS4P_UPDATE_CHECK"] = "off"
# Failure is loud if either installed edition attempts a network update.
import urllib.request
def blocked(*a, **k):
    raise AssertionError("Unexpected runtime network request")
urllib.request.urlopen = blocked
path = root / "smoke.pptx"

if edition == "community":
    from ks4_license_community.notices import WARNING_13
else:
    WARNING_13 = None

text = license_policy.active().initialization_text()

def body(response):
    assert not response.is_error, response.content
    payload = json.loads("".join(
        item.text for item in response.content if hasattr(item, "text")
    ))
    assert payload.get("ok") is True, payload
    return payload

async def check():
    async with Client(server.create_server()) as client:
        init = client.initialize_result
        assert not (init.capabilities.tools and init.capabilities.tools.listChanged)
        info = body(await client.call_tool(
            "ppt_read", {"op": "get_server_info", "args": {}}
        ))
        assert info["edition"] in {"community", "business"}, info
        tools = await client.list_tools()
        created = body(await client.call_tool(
            "ppt_deck", {"op": "create_presentation", "args": {"path": str(path)}}
        ))
        assert created["slides_kept"] == 0
        body(await client.call_tool("ppt_edit", {
            "op": "insert_slide", "args": {
                "file_path": str(path), "layout": 1, "live": "off",
            },
        }))
        body(await client.call_tool("ppt_edit", {
            "op": "set_placeholder_text", "args": {
                "file_path": str(path), "slide": 0, "placeholder": "title",
                "text": "original", "live": "off",
            },
        }))
        body(await client.call_tool("ppt_edit", {
            "op": "search_and_replace",
            "args": {
                "file_path": str(path), "find": "original", "replace": "changed",
                "live": "off",
            },
        }))
        read = body(await client.call_tool("ppt_read", {
            "op": "get_text", "args": {"file_path": str(path), "live": "off"},
        }))
        assert "changed" in read["text"], read
        with zipfile.ZipFile(path) as package:
            assert package.testzip() is None
            assert "ppt/presentation.xml" in package.namelist()
        instructions = init.instructions
        prose_at = instructions.index("PowerPoint (.pptx) editor:")
        update = info["update_check"]
        # The public PyPI report predates the authenticated Business channel:
        # it names its discriminator ``state``; the Business channel's richer
        # report names it ``status``.  Validate both documented wire shapes
        # rather than making the artifact smoke itself edition-specific by
        # assuming one key for both.
        update_state = update.get("state", update.get("status"))
        if edition == "community":
            assert update_state == "disabled", update
        else:
            # A local Business build has no release stamp yet.  It therefore
            # reports that exact closed failure before an update policy can
            # take effect; a stamped artifact with the off switch set reports
            # check_disabled instead.  Neither path may touch the network
            # (the blocked urlopen above proves that independently).
            assert update_state in {"check_disabled", "check_failed"}, update
            if update_state == "check_failed":
                assert update.get("reason") == "no_stamp", update
        return {"name": init.serverInfo.name, "version": init.serverInfo.version,
                "core_version": info["core_version"],
                "product": info["product"],
                "edition": info["edition"], "license": info["licensing"]["license"],
                "status": info["licensing"]["status"],
                "warning_in_instructions": bool(WARNING_13) and WARNING_13 in instructions,
                # PowerPoint's composition order: licensing text FIRST.
                "licensing_before_prose": (
                    instructions.index(text["instructions"]) < prose_at),
                "starts_with_introduction": instructions.startswith(
                    text["introduction"]),
                "tools": [[t.name, t.annotations.model_dump(exclude_none=True)]
                          for t in tools],
                "update_state": update_state,
                "tools_registered": info["tools_registered"],
                "ppt_edit": "passed", "python_pptx_import": "blocked"}

async def served_strings():
    # Owner-test mode: read what the installed server actually serves,
    # because copy helpers can assemble a marker that is absent as a source
    # literal. Source comments are intentionally not part of this scan.
    from kitchensink4ppt.surface import copy
    served = []
    async with Client(server.create_server()) as client:
        served.append(client.initialize_result.instructions or "")
        for tool in await client.list_tools():
            served.append(json.dumps(
                tool.model_dump(exclude_none=True), ensure_ascii=False
            ))
        listing = body(await client.call_tool("ppt_help", {}))["tools"]
        for operations in listing.values():
            for op in operations:
                served.append((await client.call_tool(
                    "ppt_help", {"op": op}
                )).content[0].text)
    served.extend(copy.HINTS.values())
    home = str(Path.home())
    defects = []
    for value in served:
        for marker in (
            "COPY" "_SLOT:", "COPY PENDING", "FACTS TO CONVEY",
            "PLACEHOLDER", chr(0x2014),
        ):
            if marker in value:
                defects.append(marker)
        if (
            home.lower() in value.lower()
            or home.replace(chr(92), "/").lower() in value.lower()
        ):
            defects.append("home folder")
    return sorted(set(defects))

result = asyncio.run(check())
if len(sys.argv) > 3 and sys.argv[3] == "owner-test":
    result["served_string_defects"] = asyncio.run(served_strings())
print(json.dumps(result))
"""


def community_needles() -> tuple[str, ...]:
    """The forbidden sentences, from the canonical shared adapter source."""
    return adapter_source.notices_module().guidance_needles()


def payload_digest(wheel: Path, prefix: str = "") -> str:
    """Content hash of a wheel's SOURCE members, ignoring zip timestamps.

    Wheels are not byte-reproducible across builds, so comparing wheel bytes
    would be a claim the format cannot support. dist-info is excluded too,
    because the two core wheels are deliberately labelled differently: one
    AGPL at X.Y.Z, one commercial at X.Y.Z+commercial. What must be identical
    is the code, and that is exactly what this digest covers.
    """
    with zipfile.ZipFile(wheel) as archive:
        parts = sorted((name, hashlib.sha256(archive.read(name)).hexdigest())
                       for name in archive.namelist()
                       if ".dist-info/" not in name and name.startswith(prefix))
    assert parts, f"{wheel.name} has no payload outside dist-info"
    return hashlib.sha256(repr(parts).encode()).hexdigest()


def wheels_for(folder: Path) -> dict:
    found = {}
    for wheel in folder.glob("*.whl"):
        if wheel.name.startswith("kitchensink4ppt_core-"):
            found["core"] = wheel
        elif wheel.name.startswith("kitchensink4_license_"):
            found["policy"] = wheel
        else:
            found["launcher"] = wheel
    community = "launcher" in found and found["launcher"].name.startswith("kitchensink4ppt-")
    required = {"policy", "launcher"} if community else {"core", "policy", "launcher"}
    assert not (community and "core" in found), "Community must not ship a separate core"
    missing = required - set(found)
    assert not missing, f"{folder} is missing {sorted(missing)}"
    return found


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("artifacts", type=Path)
    parser.add_argument(
        "--owner-test",
        action="store_true",
        help=(
            "Owner-test mode: also scan every served string for copy slots, em dashes "
            "and home folders"
        ),
    )
    args = parser.parse_args()
    needles = community_needles()
    reports = {}
    for name in ("community", "business"):
        folder = args.artifacts.resolve() / name
        wheels = wheels_for(folder)
        bundle, = folder.glob("*.mcpb")
        licenses = {}
        for role, wheel in wheels.items():
            with zipfile.ZipFile(wheel) as z:
                names = z.namelist()
                pptx_members = [
                    member for member in names if member.endswith(".pptx")
                ]
                if role == "core" or (name == "community" and role == "launcher"):
                    assert pptx_members == [TEMPLATE_MEMBER], pptx_members
                    assert (
                        hashlib.sha256(z.read(TEMPLATE_MEMBER)).hexdigest()
                        == TEMPLATE_SHA256
                    )
                else:
                    assert pptx_members == [], pptx_members
                assert not any(
                    n.endswith((".pem", ".env")) or "/tests/" in n
                    for n in names
                )
                meta, = [n for n in names if n.endswith(".dist-info/METADATA")]
                declared = ""
                for line in z.read(meta).decode("utf-8").splitlines():
                    if line.startswith("License-Expression:"):
                        declared = line.split(":", 1)[1].strip()
                        break
                # No conflicting grant anywhere in a shipped distribution.
                assert declared, f"{wheel.name} declares no license"
                assert FORBIDDEN_LICENSE[name] not in declared, (
                    f"{wheel.name} declares {declared} in a {name} distribution")
                licenses[role] = declared
                if name == "business":
                    for member in names:
                        if Path(member).suffix in {".png", ".pptx", ".pyc"}:
                            continue
                        text = z.read(member).decode("utf-8", errors="replace")
                        for needle in needles:
                            assert needle not in text, f"{wheel.name}::{member}"
        with zipfile.ZipFile(wheels["policy"]) as z:
            names = z.namelist()
            if name == "community":
                source = z.read("ks4_license_community/notices.py")
                assert bytes.fromhex("c2a7") + b"13" in source
                assert b"?13" not in source
            # The shared adapter wheel is product-neutral: no generated
            # metadata, no product name, identical for every editor.
            assert not [n for n in names if n.endswith("_edition.json")]
            for member in names:
                if not member.endswith(".py"):
                    continue
                body = z.read(member).decode("utf-8", errors="replace")
                for forbidden in ("kitchensink4ppt", "KitchenSink4PPT", "KS4P_"):
                    assert forbidden not in body, f"{wheel.name}::{member}"
            meta, = [n for n in names if n.endswith(".dist-info/METADATA")]
            if name == "business":
                assert "Requires-Dist: cryptography>=42" in z.read(meta).decode()
        with zipfile.ZipFile(wheels["launcher"]) as z:
            names = z.namelist()
            meta, = [n for n in names if n.endswith(".dist-info/METADATA")]
            metadata = z.read(meta).decode()
            expected_license = ("LicenseRef-KitchenSink-Commercial" if name == "business"
                                else "AGPL-3.0-only")
            assert "License-Expression: " + expected_license in metadata
            if name == "business":
                assert "Requires-Dist: kitchensink4ppt-core==" + CORE_VERSION + "+commercial" in metadata
            else:
                assert "Requires-Dist: kitchensink4ppt-core" not in metadata
                assert "Requires-Dist: kitchensink4-license-community==0.2.0" in metadata
            ep, = [n for n in names if n.endswith("/entry_points.txt")]
            entry_points = z.read(ep).decode()
            if name == "business":
                assert "kitchensink4ppt-business" in entry_points
                assert "ks4ppt_business_launch/_edition.json" in names
            else:
                assert "ppt-mcp" in entry_points and "kitchensink4ppt" in entry_points
        module = f"ks4ppt_{name}_launch"
        with zipfile.ZipFile(bundle) as z:
            guide = z.read("INSTALL.md").decode("utf-8")
            if name == "community":
                assert "§13" in guide
            else:
                for needle in needles:
                    assert needle not in guide
            manifest = json.loads(z.read("manifest.json"))
            assert manifest["display_name"].endswith(name.title())
            assert manifest["server"]["mcp_config"]["args"] == ["-m", module]
            assert manifest["server"]["mcp_config"]["command"] == "${user_config.python_path}"
            for wheel in wheels.values():
                assert z.read(wheel.name) == wheel.read_bytes()
            assert {"LICENSE", "INSTALL.md", "icon.png"} <= set(z.namelist())
        with tempfile.TemporaryDirectory(prefix="ks4p-installed-smoke-") as scratch:
            subprocess.run([sys.executable, "-m", "pip", "install", "--no-deps",
                            "--no-compile", "--target", scratch,
                            *[str(w) for w in wheels.values()]],
                           check=True, capture_output=True, text=True,
                           stdin=subprocess.DEVNULL)
            child_env = clean_runtime_env()
            result = subprocess.run([sys.executable, "-I", "-c", CHECK, scratch,
                                     name, "owner-test" if args.owner_test else ""],
                                    check=False, capture_output=True, text=True,
                                    encoding="utf-8", env=child_env,
                                    stdin=subprocess.DEVNULL)
            if result.returncode:
                raise RuntimeError(result.stderr)
            info = json.loads(result.stdout)
            if args.owner_test:
                assert info["served_string_defects"] == [], (
                    f"{name}: served strings carry "
                    f"{info['served_string_defects']}"
                )
            from fastmcp import Client
            from fastmcp.client.transports import StdioTransport

            async def stdio_check():
                transport = StdioTransport(
                    command=sys.executable,
                    args=["-m", module],
                    env={"PYTHONPATH": scratch,
                         "PYTHONDONTWRITEBYTECODE": "1",
                         "KS4P_UPDATE_CHECK": "off"},
                    cwd=scratch, keep_alive=False)
                async with Client(transport) as client:
                    init = client.initialize_result
                    assert init.serverInfo.name == info["name"]
                    assert init.serverInfo.version == info["version"]
                    assert not (
                        init.capabilities.tools
                        and init.capabilities.tools.listChanged
                    )
                    response = await client.call_tool(
                        "ppt_read", {"op": "get_server_info", "args": {}}
                    )
                    payload = json.loads(response.content[0].text)
                    assert payload["edition"] == name
                    return [
                        [tool.name, tool.annotations.model_dump(exclude_none=True)]
                        for tool in await client.list_tools()
                    ]

            stdio_tools = asyncio.run(
                asyncio.wait_for(stdio_check(), timeout=60))
            assert stdio_tools == info["tools"]
            info["stdio_startup_tools"] = len(stdio_tools)
            assert info["stdio_startup_tools"] == info["tools_registered"] == 9
            assert info["edition"] == name
            assert info["warning_in_instructions"] == (name == "community")
            assert info["licensing_before_prose"] is True
            assert info["starts_with_introduction"] is True
            assert ("+business" in info["version"]) == (name == "business")
            info["wheel_sha256"] = {role: hashlib.sha256(w.read_bytes()).hexdigest()
                                    for role, w in wheels.items()}
            core_wheel = wheels["core"] if name == "business" else wheels["launcher"]
            info["core_payload_digest"] = payload_digest(core_wheel, "kitchensink4ppt/")
            info["policy_payload_digest"] = payload_digest(wheels["policy"])
            info["licenses"] = licenses
            info["core_wheel"] = core_wheel.name
            reports[name] = info
    comparable = {}
    for edition, report in reports.items():
        tools = json.loads(json.dumps(report["tools"]))
        for tool, annotations in tools:
            if tool == "ppt_review":
                # Community's public update action reaches the network;
                # the Business edition keeps its offline channel contract.
                assert annotations.pop("openWorldHint") == (edition == "community")
        comparable[edition] = tools
    assert comparable["community"] == comparable["business"]
    # One shared core under both editions: same members, same content.
    assert (reports["community"]["core_payload_digest"]
            == reports["business"]["core_payload_digest"])
    reports["adapter_source"] = {
        "pinned_commit": adapter_source.PINNED_COMMIT,
        "digest": adapter_source.tree_digest(adapter_source.ensure()),
    }
    for name in ("community", "business"):
        reports[name]["tool_count"] = len(reports[name].pop("tools"))
    print(json.dumps(reports, indent=2))


if __name__ == "__main__":
    main()
