"""Build the split edition distributions and MCP bundles, without publishing.

One shared editing core, two license policy adapters, two launcher facades.
A Community distribution is one app wheel plus the shared policy adapter; a
Business distribution is core + Business policy + Business launcher. Nothing
selects an edition at runtime: the wheels in the bundle are the choice.

The adapters are NOT in this repository. They are one shared source serving
every editor in the suite, materialised here from the pinned canonical
revision and digest-verified before anything is built (adapter_source.py).
Copying them in would be the defect that constraint forbids.

The anchored fail-closed rewrite discipline is unchanged. Every textual
rewrite goes through replace_once(), which raises when an anchor is missing
or ambiguous, and the staged metadata is parsed before it is built so a
malformed or duplicated table fails here rather than in a customer's
environment. A Business build additionally passes a deterministic scan gate
that proves no Community guidance string reached any shipped artifact while
the legally required notices survived.
"""
from __future__ import annotations
import argparse
import ast
import base64
import json
import re
import shutil
import subprocess
import sys
import tempfile
import tomllib
import zipfile
from datetime import date
from pathlib import Path

if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

import adapter_source

ROOT = Path(__file__).resolve().parents[1]
PACKAGES = ROOT / "packages"
ARCHIVE_REVISION_FILE = ROOT / "scripts" / "ARCHIVE_REVISION.txt"

CORE_DIST = "kitchensink4ppt-core"
COMMUNITY_POLICY_DIST = adapter_source.COMMUNITY_POLICY_DIST
BUSINESS_POLICY_DIST = adapter_source.BUSINESS_POLICY_DIST

#: License expressions, and the rule about which may appear where. One shared
#: source revision produces two legally labelled core artifacts: the public
#: one at X.Y.Z under AGPL, and the commercial one at X.Y.Z+commercial under
#: the commercial grant. Different artifact identities, so nothing ever ships
#: two wheels with the same name and version, different contents and
#: different grants.
AGPL = "AGPL-3.0-only"
COMMERCIAL = "LicenseRef-KitchenSink-Commercial"
COMMERCIAL_CORE_SUFFIX = "+commercial"
#: A distribution must not contain a wheel granted under the other license.
FORBIDDEN_LICENSE = {"community": COMMERCIAL, "business": AGPL}

#: Notices that MUST survive in a Business artifact. "No Community guidance"
#: never means deleting a required copyright, trademark or dependency notice,
#: so the gate asserts these are present as well as asserting the guidance is
#: absent.
REQUIRED_NOTICES = (
    "Copyright (c) 2026 Alvut Consulting, LLC",
    "Not affiliated with, endorsed by, or sponsored by Microsoft Corporation.",
)

#: Stronger than the set-level check above: every wheel in a distribution has
#: to carry this line ITSELF. A set-level check lets a wheel whose notices
#: went missing be excused by a sibling wheel that still has them, which is
#: the hole a shared-notice defect sat in.
REQUIRED_IN_EVERY_WHEEL = "Copyright (c) 2026 Alvut Consulting, LLC"

#: Artifact members whose bytes are not text and are never scanned.
_BINARY_SUFFIXES = {
    ".png", ".ico", ".pptx", ".pyc", ".whl", ".mcpb", ".so", ".pyd",
}

# ------------------------------------------------------- the release channel

#: Release channels a build may declare, matching the policy adapter's
#: ALLOWED_CHANNELS. The adapter is the enforcing copy; this one exists so a
#: bad --channel is rejected at the command line rather than at a customer's
#: first startup.
CHANNELS = ("beta", "stable")
DEFAULT_CHANNEL = "beta"

#: Words a Business customer must never read in a shipped artifact. The
#: product is a beta and "beta" is the word for it. The other five each claim
#: something different and none of them is true here: "preview" and "RC"
#: promise a finished build behind them on a schedule nobody has published,
#: "trial" and "demo" say the licence expires or the software is crippled, and
#: a customer who reads any of them has been told the wrong thing about what
#: they paid for. The gate is a spelling check on a commercial promise.
_PRERELEASE_WORDS = re.compile(
    r"\b(?:previews?|previewing|release[\s-]+candidates?|trials?|demos?)\b",
    re.IGNORECASE)
#: "RC" only as a standalone uppercase token, so ordinary words that contain
#: the letters are not swept up. Case-sensitive on purpose: the marketing term
#: is capitalised, and lowercase "rc" appears in ordinary technical text.
_PRERELEASE_RC = re.compile(r"(?<![A-Za-z0-9])RC(?![A-Za-z0-9])")

#: Members of a Business artifact set that the language gate does NOT read,
#: each with the reason it is out of scope. This list is the gate's honesty:
#: everything named here still carries pre-release language today, and every
#: entry is a document whose wording belongs to somebody other than the build
#: script. Nothing is allowlisted for being inconvenient.
PRERELEASE_ALLOWLIST = {
    "NOTICE.md":
        "third-party and trademark notices; obligations, not product copy",
    "CHANGELOG.md":
        "release history, which must be free to say what a past release was",
    "LICENSE-BETA.txt":
        "the beta licence terms, shipped verbatim with every bundle and "
        "every receipt envelope; its wording is fixed and must be "
        "byte-identical wherever it appears",
}

#: The members the gate DOES read, by suffix or exact name, out of the
#: Business artifact set. Narrow on purpose: these are the surfaces the build
#: script itself composes, plus the generated metadata, and a customer reads
#: every one of them.
def _prerelease_targets(label: str) -> bool:
    """True when this artifact member is customer-facing build-authored text."""
    name = label.rsplit("::", 1)[-1]
    base = name.rsplit("/", 1)[-1]
    if base in PRERELEASE_ALLOWLIST:
        return False
    if base in ("LICENSE", "INSTALL.md"):
        # The commercial licence notice and the filtered install text. Both
        # were allowlisted while they still called this a development
        # preview; both now say beta release, so both are read.
        return True
    if base == "manifest.json":
        return True
    if base == "_edition.json":
        return True
    if base == "METADATA" and name.endswith(".dist-info/METADATA"):
        return True
    return False


def find_prerelease_language(text: str) -> list[str]:
    """Every pre-release word in one string, in the order they appear."""
    return ([match.group(0) for match in _PRERELEASE_WORDS.finditer(text)]
            + [match.group(0) for match in _PRERELEASE_RC.finditer(text)])


def scan_prerelease_language(paths, exempt_text=()) -> dict:
    """Prove no shipped Business artifact calls this build anything but a beta.

    exempt_text holds documents that are embedded verbatim inside a scanned
    member: the commercial notice inside the generated launcher README, and
    the repository README inside every wheel's METADATA. Subtracting the
    embedded document keeps the gate on the build script's own words without
    allowlisting the whole member, so new pre-release language in the
    generated half is still caught.

    Line endings are normalised on both sides before the subtraction. A
    document read through Path.read_text arrives with LF and the same
    document read out of a zip member keeps whatever it was written with, and
    a gate that silently failed to subtract would report the embedded
    document's own words as a leak.
    """
    exempt_text = tuple(e.replace("\r\n", "\n") for e in exempt_text if e)
    leaks: list[str] = []
    scanned = 0
    for path in paths:
        path = Path(path)
        for label, text in _members(path):
            if not _prerelease_targets(label):
                continue
            scanned += 1
            text = text.replace("\r\n", "\n")
            for embedded in exempt_text:
                text = text.replace(embedded, "")
            for word in find_prerelease_language(text):
                leaks.append(f"{label}: {word!r}")
    if not scanned:
        raise BuildScanError(
            "The pre-release language gate read no customer-facing members, "
            "which means it would pass anything.")
    if leaks:
        raise BuildScanError(
            "A shipped Business artifact calls this build something other "
            "than a beta:\n  " + "\n  ".join(leaks))
    return {"members_scanned": scanned,
            "allowlisted": sorted(PRERELEASE_ALLOWLIST)}




class BuildScanError(RuntimeError):
    """A shipped Business artifact failed the deterministic content gate."""


def replace_once(text: str, old: str, new: str) -> str:
    """Fail closed if source metadata drifts or an anchor becomes ambiguous."""
    if text.count(old) != 1:
        raise ValueError(f"Expected exactly one build metadata anchor: {old!r}")
    return text.replace(old, new, 1)


def staged_config(path: Path, rewrites: list[tuple[str, str]]) -> str:
    """Apply anchored rewrites and prove the result is still valid metadata."""
    text = path.read_text("utf-8")
    for old, new in rewrites:
        text = replace_once(text, old, new)
    # Reject duplicate tables or malformed staged metadata before building.
    tomllib.loads(text)
    return text


# ------------------------------------------------------------- the scan gate


def community_needles() -> tuple[str, ...]:
    """Load the forbidden sentences from the Community notice module itself.

    Read out of the canonical source file rather than an installed package,
    so the gate cannot pass because the adapter happened not to be installed,
    and cannot drift when the approved wording is revised.
    """
    return adapter_source.notices_module().guidance_needles()


def _members(path: Path):
    """Yield (label, text) for every scannable member of an artifact."""
    if path.suffix in {".whl", ".mcpb", ".zip"}:
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if Path(name).suffix in _BINARY_SUFFIXES or name.endswith("/"):
                    continue
                yield f"{path.name}::{name}", archive.read(name).decode(
                    "utf-8", errors="replace")
    elif path.suffix not in _BINARY_SUFFIXES:
        yield path.name, path.read_text("utf-8", errors="replace")


#: Owner-test mode. An unfilled copy slot or a builder's home folder in
#: any shipped member fails the build. The installed-server smoke separately
#: checks strings assembled at runtime, where source comments are absent.
OWNER_TEST_MARKERS = ("COPY" "_SLOT:", "COPY PENDING", "FACTS TO CONVEY")
OWNER_TEST_SLOT_MODULES = (
    "kitchensink4ppt/core/messages.py",
    "kitchensink4ppt/surface/copy.py",
)


def require_archive_source() -> str:
    """Return the source commit injected by ``git archive``, or refuse.

    A working tree can have line-ending conversions, untracked files and
    uncommitted edits.  Owner-test artifacts therefore come only from an
    archive of one committed tree.  ``export-subst`` changes the marker to a
    40-hex commit only inside ``git archive``; its literal checkout form is
    deliberately invalid here.
    """
    try:
        revision = ARCHIVE_REVISION_FILE.read_text("ascii").strip()
    except OSError as exc:
        raise SystemExit(
            "owner-test release build refused: archive revision marker is missing"
        ) from exc
    if re.fullmatch(r"[0-9a-f]{40}", revision) is None:
        raise SystemExit(
            "owner-test release build refused: run this command from a clean "
            "`git archive HEAD` extraction, never from a checkout or worktree"
        )
    return revision


def _home_needles() -> tuple[str, ...]:
    home = Path.home()
    forms = {str(home), home.as_posix(), str(home).replace("\\", "\\\\")}
    return tuple(sorted(form for form in forms if len(form) > 3))


def _unfilled_slot_references(source: str) -> list[int]:
    """Return source lines that still assemble customer-copy slots.

    Slot markers are deliberately split (``"COPY" "_SLOT:"``), so a byte
    scan alone cannot prove that the writer filled them.  Check the syntax
    tree for the two supported slot shapes: a load of the ``SLOT`` sentinel
    and a call to the historical ``_slot`` helper.  A dormant helper
    definition is allowed; a call or sentinel load is not.
    """
    try:
        tree = ast.parse(source)
    except SyntaxError:
        return []
    lines: set[int] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load):
            if node.id == "SLOT":
                lines.add(node.lineno)
        if isinstance(node, ast.Call):
            function = node.func
            name = (
                function.id if isinstance(function, ast.Name)
                else function.attr if isinstance(function, ast.Attribute)
                else ""
            )
            if name == "_slot":
                lines.add(node.lineno)
    return sorted(lines)


def scan_owner_test(paths) -> dict:
    """Refuse unfilled copy slots and paths tied to the build machine."""
    hits: list[str] = []
    scanned = 0
    homes = _home_needles()
    for path in paths:
        for label, text in _members(Path(path)):
            scanned += 1
            for marker in OWNER_TEST_MARKERS:
                if marker in text:
                    hits.append(f"{label}: {marker}")
            # A writer may assemble the marker inside _slot(), so the literal
            # static scan above can be green while a runtime string is still
            # blocked. Any helper call left in either shipped message module
            # is an unfilled slot; the helper definition itself is permitted.
            if label.replace("\\", "/").endswith(OWNER_TEST_SLOT_MODULES):
                slot_lines = _unfilled_slot_references(text)
                if slot_lines:
                    hits.append(
                        f"{label}: {len(slot_lines)} unfilled copy slot(s) "
                        f"at lines {', '.join(map(str, slot_lines[:8]))}"
                    )
            lowered = text.lower()
            for home in homes:
                if home.lower() in lowered:
                    hits.append(f"{label}: home folder path")
    if hits:
        separator = chr(10) + "  "
        more = (
            f"{separator}... and {len(hits) - 40} more"
            if len(hits) > 40
            else ""
        )
        raise SystemExit(
            "owner-test ship gate refused the build:"
            + separator
            + separator.join(hits[:40])
            + more
        )
    return {"members_scanned": scanned, "markers": list(OWNER_TEST_MARKERS)}


def scan_business_artifacts(paths, needles=None, required=REQUIRED_NOTICES) -> dict:
    """Prove a Business artifact set carries no Community guidance and still
    carries its required notices. Raises rather than warning.

    Notice survival is reported as WHICH artifact carries each notice, not as
    a count. A count cannot tell you a disclaimer ended up in the wrong wheel,
    and it was a count that would have passed the defect where one
    product's legal notices sat in a wheel shared by four.
    """
    needles = tuple(needles if needles is not None else community_needles())
    if not needles:
        raise BuildScanError("The scan gate found no Community needles to "
                             "look for, which means it would pass anything.")
    leaks: list[str] = []
    carriers = {text: [] for text in required}
    orphan_wheels = []
    scanned = 0
    for path in paths:
        path = Path(path)
        carries_own = False
        for label, text in _members(path):
            scanned += 1
            for needle in needles:
                if needle in text:
                    leaks.append(f"{label}: {needle[:60]}...")
            for text_required in carriers:
                if (text_required in text
                        and path.name not in carriers[text_required]):
                    carriers[text_required].append(path.name)
            if REQUIRED_IN_EVERY_WHEEL in text:
                carries_own = True
        if path.suffix == ".whl" and not carries_own:
            orphan_wheels.append(path.name)
    missing = [t for t, seen in carriers.items() if not seen]
    if leaks:
        raise BuildScanError(
            "Community guidance reached a Business artifact:\n  "
            + "\n  ".join(leaks))
    if missing:
        raise BuildScanError(
            "Required notices are missing from the Business artifacts:\n  "
            + "\n  ".join(missing))
    if orphan_wheels:
        raise BuildScanError(
            f"These wheels do not carry {REQUIRED_IN_EVERY_WHEEL!r} "
            "themselves:\n  " + "\n  ".join(orphan_wheels))
    return {"members_scanned": scanned, "needles": len(needles),
            "required_notices": len(required), "carriers": carriers}


def wheel_license(path: Path) -> tuple[str, str]:
    """The license expression a wheel declares, read from its METADATA."""
    with zipfile.ZipFile(path) as archive:
        names = [n for n in archive.namelist() if n.endswith(".dist-info/METADATA")]
        if len(names) != 1:
            raise BuildScanError(f"{path.name} has {len(names)} METADATA files")
        declared = ""
        for line in archive.read(names[0]).decode("utf-8").splitlines():
            if line.startswith("License-Expression:"):
                declared = line.split(":", 1)[1].strip()
            elif line.startswith("License:") and not declared:
                declared = line.split(":", 1)[1].strip()
            elif not line.strip():
                break
    return path.name, declared


def scan_license_metadata(paths, edition: str) -> dict:
    """No distribution ships a wheel granted under the other edition's license.

    This is the check that keeps a Business bundle from carrying an
    AGPL-labelled core wheel beside a commercial launcher, which would put
    two conflicting grants in one customer's hands and make the shipped
    metadata, not the agreement, the thing a customer reads.
    """
    forbidden = FORBIDDEN_LICENSE[edition]
    declared, wrong = {}, []
    for path in paths:
        path = Path(path)
        if path.suffix != ".whl":
            continue
        name, expression = wheel_license(path)
        declared[name] = expression
        if not expression:
            wrong.append(f"{name}: declares no license expression")
        elif forbidden in expression:
            wrong.append(f"{name}: declares {expression}, forbidden in a "
                         f"{edition} distribution")
    if not declared:
        raise BuildScanError("The license gate found no wheels to check, "
                             "which means it would pass anything.")
    if wrong:
        raise BuildScanError(
            f"Conflicting license metadata in the {edition} distribution:\n  "
            + "\n  ".join(wrong))
    return declared


def scan_adapter_provenance(source: Path | None = None) -> str:
    """Prove the adapters being built are the canonical shared source.

    A forked adapter is the packaging defect the shared-source constraint
    exists to prevent, and it is invisible in a built wheel: same name, same
    version, different payload. The digest is the only mechanical evidence,
    so the build refuses without it.
    """
    source = adapter_source.ensure() if source is None else source
    digest = adapter_source.tree_digest(source)
    if digest != adapter_source.expected_digest():
        raise BuildScanError(
            "The license policy adapters are not the pinned canonical "
            f"revision: expected {adapter_source.expected_digest()}, found "
            f"{digest}.")
    vendored = [p for p in PACKAGES.glob("kitchensink4-license-*")]
    if vendored:
        raise BuildScanError(
            "The shared license policy adapters were copied into this "
            f"repository: {[p.name for p in vendored]}. They are one shared "
            "source for the whole suite and are consumed from the canonical "
            "revision, never forked per product.")
    return digest


def business_install_doc() -> str:
    """The Business install text, filtered out of the shared editions doc.

    Mechanical section removal, not a rewrite: the Community section is the
    one that carries guidance prose, so it is the one that does not ship.
    """
    lines = (ROOT / "docs/editions.md").read_text("utf-8").splitlines(keepends=True)
    out, skipping = [], False
    for line in lines:
        if line.startswith("## "):
            skipping = line.strip() == "## Community"
        if not skipping:
            out.append(line)
    return "".join(out)


# ------------------------------------------------------------ wheel builders


def _build_wheel(stage: Path, output: Path) -> Path:
    before = set(output.glob("*.whl"))
    subprocess.run([sys.executable, "-m", "build", "--wheel", "--no-isolation",
                    "--outdir", str(output), str(stage)], check=True,
                   stdin=subprocess.DEVNULL)
    made = set(output.glob("*.whl")) - before
    if len(made) != 1:
        raise ValueError(f"Expected exactly one new wheel, got {sorted(made)}")
    return made.pop()


def build_core(edition: str, output: Path, scratch: Path) -> tuple[Path, str]:
    """The editing core, from ONE source revision, labelled for its edition.

    Community gets X.Y.Z under AGPL, which is what goes to PyPI. Business gets
    X.Y.Z+commercial under the commercial grant, with the commercial notice as
    its LICENSE. Same tested source, two artifact identities, no wheel that
    says AGPL sitting inside a commercial bundle.

    NOTICE.md ships in BOTH. It carries the Alvut copyright, the required
    notice line and the Microsoft trademark disclaimer, and those obligations
    do not change with the grant. Relabelling a distribution never means
    dropping a notice.
    """
    stage = scratch / "core"
    stage.mkdir()
    shutil.copytree(ROOT / "src", stage / "src",
                    ignore=shutil.ignore_patterns("__pycache__", "*.egg-info"))
    for name in ("README.md", "LICENSE", "NOTICE.md"):
        shutil.copy2(ROOT / name, stage / name)
    version = BASE_VERSION
    rewrites = []
    if edition == "business":
        version = BASE_VERSION + COMMERCIAL_CORE_SUFFIX
        rewrites = [
            (f'version = "{BASE_VERSION}"', f'version = "{version}"'),
            (f'license = "{AGPL}"', f'license = "{COMMERCIAL}"'),
        ]
        notice = (ROOT / "business/COMMERCIAL_LICENSE_NOTICE.md").read_text("utf-8")
        (stage / "LICENSE").write_text(notice, encoding="utf-8")
    (stage / "pyproject.toml").write_text(
        staged_config(ROOT / "pyproject.toml", rewrites), encoding="utf-8")
    return _build_wheel(stage, output), version


def build_community_app(output: Path, scratch: Path, policy_version: str) -> tuple[Path, str]:
    """Keep the editing modules owned by the public app across pip upgrades."""
    if policy_version != "0.2.0":
        raise ValueError("Community requires kitchensink4-license-community==0.2.0")
    stage = scratch / "community-app"
    stage.mkdir()
    shutil.copytree(ROOT / "src", stage / "src",
                    ignore=shutil.ignore_patterns("__pycache__", "*.egg-info"))
    shutil.copytree(PACKAGES / "kitchensink4ppt-community/src/ks4ppt_community_launch",
                    stage / "src/ks4ppt_community_launch",
                    ignore=shutil.ignore_patterns("__pycache__", "*.egg-info"))
    for name in ("README.md", "LICENSE", "NOTICE.md"):
        shutil.copy2(ROOT / name, stage / name)
    config = staged_config(ROOT / "pyproject.toml", [
        ('name = "kitchensink4ppt-core"', 'name = "kitchensink4ppt"'),
        ('dependencies = [', 'dependencies = [\n    "kitchensink4-license-community==0.2.0",'),
    ])
    config += '\n[project.scripts]\nppt-mcp = "ks4ppt_community_launch:main"\nkitchensink4ppt = "ks4ppt_community_launch:main"\n'
    tomllib.loads(config)
    (stage / "pyproject.toml").write_text(config, encoding="utf-8", newline="\n")
    return _build_wheel(stage, output), BASE_VERSION


def build_policy_adapter(edition: str, output: Path, scratch: Path,
                         source: Path) -> tuple[Path, str]:
    """A license policy adapter wheel, built from the canonical shared source.

    Nothing product-specific is generated into either one, so the same
    adapter wheel serves every editor in the suite with identical contents,
    and this build copies the pinned revision rather than owning a fork of it.
    """
    name = COMMUNITY_POLICY_DIST if edition == "community" else BUSINESS_POLICY_DIST
    package = source / name
    stage = scratch / name
    shutil.copytree(package, stage, ignore=shutil.ignore_patterns(
        "__pycache__", "*.egg-info", "_edition.json"))
    version = tomllib.loads(
        (package / "pyproject.toml").read_text("utf-8"))["project"]["version"]
    return _build_wheel(stage, output), version


def build_launcher(edition: str, output: Path, scratch: Path, release_date: date,
                   keys: dict, core_version: str,
                   policy_version: str,
                   channel: str = DEFAULT_CHANNEL, roots: dict | None = None,
                   trust_list: dict | None = None,
                   require_chain: bool = False) -> tuple[Path, str]:
    """The edition distribution: the public console names, the pinned wheels,
    and, for Business, the product configuration and generated build metadata.

    The metadata lands HERE rather than in the shared adapter, because a
    generated _edition.json inside the adapter would make that wheel
    product-specific while keeping its name and version.
    """
    name = f"kitchensink4ppt-{edition}"
    source = PACKAGES / name
    stage = scratch / name
    shutil.copytree(source, stage, ignore=shutil.ignore_patterns(
        "__pycache__", "*.egg-info", "_edition.json"))
    policy_dist = (COMMUNITY_POLICY_DIST if edition == "community"
                   else BUSINESS_POLICY_DIST)
    rewrites = [
        (f'    "{CORE_DIST}",', f'    "{CORE_DIST}=={core_version}",'),
        (f'    "{policy_dist}",', f'    "{policy_dist}=={policy_version}",'),
    ]
    version = BASE_VERSION
    if edition == "business":
        version = f"{BASE_VERSION}+business"
        rewrites.append((f'version = "{BASE_VERSION}"', f'version = "{version}"'))
    (stage / "pyproject.toml").write_text(
        staged_config(source / "pyproject.toml", rewrites), encoding="utf-8")
    shutil.copy2(ROOT / "README.md", stage / "README.md")
    if edition == "business":
        notice = (ROOT / "business/COMMERCIAL_LICENSE_NOTICE.md").read_text("utf-8")
        (stage / "README.md").write_text(
            "# KitchenSink4PPT Business\n\nKitchenSink4AI Business, beta "
            "release. Built and verified like production, licensed for "
            "evaluation under the Pilot Evaluation License.\n\n" + notice,
            encoding="utf-8")
        (stage / "src/ks4ppt_business_launch/_edition.json").write_text(
            json.dumps({
                "edition": "business",
                # Must match the launcher's own PRODUCT_ID; the adapter
                # refuses the pair when it does not.
                "package": "kitchensink4ppt-business",
                "version": version,
                # The stage of the product, carried as data so the server
                # can report it and the adapter can refuse a value that is
                # not one of the two. It is deliberately NOT a PEP 440
                # pre-release suffix on the version: the wheel version stays
                # X.Y.Z+business so that pip resolution, the installer's
                # expectations and the published identity of the artifact do
                # not move with the label.
                "channel": channel,
                "release_date": release_date.isoformat(),
                "public_keys": keys,
                # THE ROOTS ARE THE PIN. A signing key can be rotated by
                # issuing it a certificate under an intermediate; a root
                # cannot, which is why only roots are frozen into an artifact
                # and why they are kept offline. public_keys above stays for
                # the receipts already issued under the flat model and comes
                # out a release after every one of them has been re-signed.
                "root_keys": dict(roots or {}),
                # The baseline signed list, carried so a build knows which
                # keys have been pulled without asking anything over a
                # network. A newer list rides in a receipt or an update.
                "trust_list": trust_list,
                # accept_flat false is the switch that ends the compatibility
                # window. It is set at build time and read at verify time; an
                # installation cannot change it, because an installation that
                # could relax it could accept a receipt this artifact was not
                # built to accept.
                "chain_policy": {"accept_flat": not require_chain},
            }, indent=2), encoding="utf-8")
    return _build_wheel(stage, output), version


BASE_VERSION = tomllib.loads(
    (ROOT / "pyproject.toml").read_text("utf-8"))["project"]["version"]


# ------------------------------------------------------------------- bundles


#: The key id a single unnamed --public-key gets, unchanged from before the
#: flags became repeatable.
DEFAULT_KEY_ID = "primary"


def _public_key_b64(path) -> str:
    """One Ed25519 public key file as raw base64, PEM or .pub.

    The issuer keys are handed over as PEM and the ROOT keys as the base64
    .pub files keygen.py writes, so the loader that reads a root has to accept
    both rather than make an operator convert a key by hand at the one moment
    where a mistake is expensive. The 32 raw bytes are the same either way and
    that is what a build records.
    """
    blob = Path(path).read_bytes()
    from cryptography.hazmat.primitives import serialization
    from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey

    if b"-----BEGIN" in blob:
        key = serialization.load_pem_public_key(blob)
        if not isinstance(key, Ed25519PublicKey):
            raise ValueError("An Ed25519 public key is required")
        return base64.b64encode(key.public_bytes(
            serialization.Encoding.Raw,
            serialization.PublicFormat.Raw)).decode("ascii")
    try:
        raw = base64.b64decode(b"".join(blob.split()), validate=True)
    except (ValueError, TypeError) as problem:
        raise ValueError(f"{path} is neither a PEM nor a base64 public key: "
                         f"{problem}") from problem
    if len(raw) != 32:
        raise ValueError(f"{path} holds {len(raw)} bytes, not the 32 bytes of an "
                         "Ed25519 public key")
    return base64.b64encode(raw).decode("ascii")


def root_keys(root_key, root_id) -> dict:
    """Map root key id to base64 public key, from paired repeatable arguments.

    Paired POSITIONALLY like --public-key and --key-id, and for the same
    reason. A build carries the global root AND the pre-trusted spare, so that
    losing the global root is a trust-list update rather than a recall: a
    product that already trusts the spare can be told to move to it by a list
    the spare signs. Both are named on the command line, because a root a
    build trusts and nobody wrote down is a root nobody can audit.
    """
    paths = ([] if root_key is None
             else [root_key] if isinstance(root_key, (str, Path))
             else list(root_key))
    ids = ([] if root_id is None
           else [root_id] if isinstance(root_id, str)
           else list(root_id))
    if not paths:
        return {}
    if len(ids) != len(paths):
        raise ValueError(
            f"{len(paths)} root key(s) and {len(ids)} root id(s). --root-key "
            "and --root-id are paired in order, so each root needs exactly "
            "one id.")
    keys: dict = {}
    for path, name in zip(paths, ids):
        if not isinstance(name, str) or not name.strip() or len(name) > 80:
            raise ValueError(f"Invalid root key id: {name!r}")
        if name in keys:
            raise ValueError(
                f"Duplicate root key id {name!r}. Each id names exactly one "
                "root, and a build that anchors a key the command line did "
                "not name is a build nobody can audit.")
        keys[name] = _public_key_b64(path)
    return keys


def load_trust_list(path, roots: dict) -> dict:
    """Read a signed trust list and prove it against the roots being built in.

    Proved HERE, at build time, rather than only in the adapter, because a
    build that ships a list its own roots do not vouch for is an artifact that
    fails at a customer's first status call. The adapter checks it again at
    verification time; this check is the one that stops the artifact from
    existing.
    """
    adapter_source.install_import_path()
    from ks4_license_business import chain

    data = json.loads(Path(path).read_text(encoding="utf-8"))
    try:
        chain.verify_trust_list(data, roots=roots)
    except chain.ChainError as problem:
        raise ValueError(f"{path} does not verify against the roots given: "
                         f"{problem}") from problem
    # The SIGNED object is what a build embeds, not the body it wraps: a
    # product has to be able to check that signature for itself.
    return data


def issuer_keys(public_key, key_id) -> dict:
    """Map key id to base64 raw public key, from paired repeatable arguments.

    --public-key and --key-id are paired POSITIONALLY: the nth id names the
    nth key. A build may need several, because the purchase flow has two
    issuers. The web tier signs self-serve receipts on the Worker with one
    key; the offline root signs manual and enterprise receipts with another.
    An installation cannot know which counter a customer bought at, so it
    carries both and the receipt's own key id selects between them.

    One pair is the ordinary case and behaves exactly as it did before this
    function existed, including the default id when no --key-id is given.
    Duplicate ids are refused rather than resolved: a map cannot hold one id
    twice, and silently keeping the last one would mean a build that trusts a
    different key than the command line named.
    """
    paths = ([] if public_key is None
             else [public_key] if isinstance(public_key, (str, Path))
             else list(public_key))
    ids = ([] if key_id is None
           else [key_id] if isinstance(key_id, str)
           else list(key_id))
    if not paths:
        return {}
    if not ids:
        ids = [DEFAULT_KEY_ID]
    if len(ids) != len(paths):
        raise ValueError(
            f"{len(paths)} public key(s) and {len(ids)} key id(s). "
            "--public-key and --key-id are paired in order, so each key needs "
            "exactly one id.")
    keys: dict[str, str] = {}
    for path, name in zip(paths, ids):
        if not isinstance(name, str) or not name.strip() or len(name) > 80:
            raise ValueError(f"Invalid issuer key id: {name!r}")
        if name in keys:
            raise ValueError(
                f"Duplicate issuer key id {name!r}. Each id names exactly one "
                "public key, and a build that trusts a key the command line "
                "did not name is a build nobody can audit.")
        # PEM or the one-line base64 .pub the key ceremony actually writes.
        # Reading only PEM here cost the first attempt at the beta build, and
        # the two encodings carry the same 32 bytes, so refusing one of them
        # bought nothing.
        keys[name] = _public_key_b64(path)
    return keys


def build(edition: str, output: Path, release_date: date, public_key=None,
          key_id=DEFAULT_KEY_ID, channel: str = DEFAULT_CHANNEL,
          root_key=None, root_id=None, trust_list=None,
          require_chain: bool = False, owner_test: bool = False) -> list[Path]:
    """Build behind every gate, then atomically publish one edition.

    The candidate lives in a same-filesystem staging directory that this call
    alone owns. No wheel or bundle appears at ``output`` until the Business,
    language, owner-test and license gates have all passed. A refusal removes
    only that staging directory, so an unrelated or concurrently-created
    destination is never treated as failed-build debris.
    """
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        raise FileExistsError(f"Build destination already exists: {output}")
    staging_root = Path(tempfile.mkdtemp(
        prefix=f".{output.name}.staging-", dir=output.parent))
    staging = staging_root / "candidate"
    try:
        artifacts = _build_staged(
            edition, staging, release_date, public_key, key_id, channel,
            root_key, root_id, trust_list, require_chain, owner_test)
        relative = [artifact.relative_to(staging) for artifact in artifacts]
        # staging and output share a parent filesystem, so the completed
        # artifact set becomes visible in one directory rename.
        staging.rename(output)
        return [output / path for path in relative]
    finally:
        shutil.rmtree(staging_root, ignore_errors=True)


def _build_staged(edition: str, output: Path, release_date: date,
                  public_key=None, key_id=DEFAULT_KEY_ID,
                  channel: str = DEFAULT_CHANNEL, root_key=None,
                  root_id=None, trust_list=None,
                  require_chain: bool = False,
                  owner_test: bool = False) -> list[Path]:
    if channel not in CHANNELS:
        raise ValueError(f"channel must be one of {CHANNELS}, got {channel!r}")
    business = edition == "business"
    keys = issuer_keys(public_key, key_id)
    roots = root_keys(root_key, root_id)
    if require_chain and not roots:
        raise ValueError(
            "--require-chain without --root-key would build an edition that "
            "refuses flat receipts and cannot check chained ones, which is an "
            "edition that verifies nothing.")
    trust = load_trust_list(trust_list, roots) if trust_list is not None else None
    # The adapters are proved canonical BEFORE anything is built, so a forked
    # adapter stops the build rather than shipping under a shared name.
    adapters = adapter_source.ensure()
    digest = scan_adapter_provenance(adapters)
    # Temp build roots are OS scratch, never inside the source checkout.
    output.mkdir(parents=True, exist_ok=False)
    with tempfile.TemporaryDirectory(prefix="ks4p-edition-") as scratch_dir:
        scratch = Path(scratch_dir)
        policy, policy_version = build_policy_adapter(edition, output, scratch,
                                                      adapters)
        if business:
            core, core_version = build_core(edition, output, scratch)
            launcher, version = build_launcher(
                edition, output, scratch, release_date, keys, core_version,
                policy_version, channel, roots, trust, require_chain)
            wheels = [core, policy, launcher]
        else:
            app, version = build_community_app(output, scratch, policy_version)
            wheels = [policy, app]
    package = "kitchensink4ppt-business" if business else "kitchensink4ppt"
    module = "ks4ppt_business_launch" if business else "ks4ppt_community_launch"

    manifest = json.loads((ROOT / "bundle/manifest.json").read_text("utf-8"))
    manifest.update(name=package, display_name=f"KitchenSink4PPT {edition.title()}",
                    version=version, license=("LicenseRef-KitchenSink-Commercial"
                    if business else "AGPL-3.0-only"))
    manifest["description"] = (
        f"KitchenSink4PPT {edition.title()}, beta release. Install the "
        "supplied wheels and register the launcher as described in INSTALL.md.")
    manifest["long_description"] = manifest["description"]
    manifest["author"] = {"name": "KitchenSink4AI"}
    manifest["server"]["mcp_config"]["command"] = "${user_config.python_path}"
    manifest["server"]["mcp_config"]["args"] = ["-m", module]
    # An edition bundle runs its own wheels, not the public Desktop bundle,
    # so it must not announce the bundle install route (core/update_flow.py).
    manifest["server"]["mcp_config"]["env"].pop("KS4P_INSTALL_ROUTE", None)
    manifest["user_config"]["python_path"] = {
        "type": "string", "title": f"{edition.title()} Python executable",
        "description": "Absolute path to Python in the dedicated environment containing this edition.",
        "required": True,
    }
    if business:
        manifest["user_config"]["license_file"] = {
            "type": "string", "title": "Business license receipt",
            "description": "Path to the offline signed JSON receipt, if supplied.",
            "required": False, "default": "",
        }
        manifest["server"]["mcp_config"]["env"]["KS4P_LICENSE_FILE"] = "${user_config.license_file}"
    archive = output / f"{package}-{version}.mcpb"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as bundle:
        bundle.writestr("manifest.json", json.dumps(manifest, indent=2))
        bundle.write(ROOT / "bundle/icon.png", "icon.png")
        bundle.write(ROOT / ("business/COMMERCIAL_LICENSE_NOTICE.md" if business else "LICENSE"),
                     "LICENSE")
        if business:
            # The shared editions doc carries Community guidance prose, so the
            # Business bundle ships the filtered install text instead.
            bundle.writestr("INSTALL.md", business_install_doc())
        else:
            bundle.write(ROOT / "docs/editions.md", "INSTALL.md")
        for wheel in wheels:
            bundle.write(wheel, wheel.name)
    artifacts = wheels + [archive]
    print(f"adapter gate: canonical shared source "
          f"{adapter_source.PINNED_COMMIT[:12]}, digest {digest[:16]}")
    if business:
        report = scan_business_artifacts(artifacts)
        print(f"scan gate: {report['members_scanned']} members, "
              f"{report['needles']} needles, required notices present in all "
              f"{len([a for a in artifacts if a.suffix == '.whl'])} wheels")
        for notice, carried_by in report["carriers"].items():
            print(f"  notice {notice[:40]!r}... carried by "
                  f"{len(carried_by)} artifact(s): {', '.join(carried_by)}")
        language = scan_prerelease_language(artifacts)
        print(f"language gate: {language['members_scanned']} customer-facing "
              f"members carry no pre-release wording; allowlisted verbatim "
              f"documents: {', '.join(language['allowlisted'])}")
    if owner_test:
        report = scan_owner_test(artifacts)
        print(
            f"owner-test gate: {report['members_scanned']} members carry no "
            "copy slot and no home folder"
        )
    # Both editions are checked. A commercial wheel loose in the public
    # distribution is the same defect wearing the other hat.
    declared = scan_license_metadata(artifacts, edition)
    print(f"license gate ({edition}): "
          + ", ".join(f"{n.split('-')[0]}={e}" for n, e in sorted(declared.items())))
    return artifacts


def build_set(editions: tuple[str, ...], output: Path, release_date: date,
              public_key=None, key_id=DEFAULT_KEY_ID,
              channel: str = DEFAULT_CHANNEL, root_key=None, root_id=None,
              trust_list=None, require_chain: bool = False,
              owner_test: bool = False) -> dict[str, list[Path]]:
    """Build an edition set and publish the whole set in one rename.

    A two-edition invocation is one release candidate.  If the second
    edition refuses, the first edition must not remain in the requested
    destination.  Per-edition staging still protects each inner build; this
    outer stage provides the all-or-nothing boundary for ``--edition both``.
    """
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        raise FileExistsError(f"Build destination already exists: {output}")
    staging_root = Path(tempfile.mkdtemp(
        prefix=f".{output.name}.set-staging-", dir=output.parent))
    candidate = staging_root / "candidate"
    staged: dict[str, list[Path]] = {}
    try:
        for edition in editions:
            staged[edition] = build(
                edition, candidate / edition, release_date, public_key,
                key_id, channel, root_key, root_id, trust_list,
                require_chain, owner_test,
            )
        relative = {
            edition: [artifact.relative_to(candidate) for artifact in artifacts]
            for edition, artifacts in staged.items()
        }
        candidate.rename(output)
        return {
            edition: [output / path for path in paths]
            for edition, paths in relative.items()
        }
    finally:
        shutil.rmtree(staging_root, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--edition", choices=["community", "business", "both"], default="both")
    parser.add_argument("--output", type=Path, required=True,
                        help="New destination: existing edition directories are never overwritten")
    parser.add_argument("--release-date", type=date.fromisoformat, required=True)
    parser.add_argument("--public-key", type=Path, action="append",
                        help="Issuer public PEM only; omitted builds cannot verify "
                             "receipts. Repeatable: pair each one with a --key-id "
                             "in the same order")
    parser.add_argument("--key-id", action="append",
                        help=f"Id for the public key in the same position "
                             f"(default: {DEFAULT_KEY_ID} for a single key)")
    parser.add_argument("--channel", choices=CHANNELS, default=DEFAULT_CHANNEL,
                        help="the release stage recorded in the Business build "
                             "metadata and reported by get_server_info "
                             f"(default: {DEFAULT_CHANNEL})")
    parser.add_argument("--root-key", type=Path, action="append",
                        help="Global root public key, PEM or base64 .pub. "
                             "Repeatable: pair each one with a --root-id in the "
                             "same order. A build trusts these and nothing else "
                             "at the top of a certificate chain")
    parser.add_argument("--root-id", action="append",
                        help="Id for the root key in the same position")
    parser.add_argument("--trust-list", type=Path,
                        help="Signed baseline trust list to embed; it must be "
                             "signed by one of the --root-key roots")
    parser.add_argument(
        "--owner-test",
        action="store_true",
        help=(
            "Owner-test mode: refuse any artifact carrying an "
            "unfilled copy slot or this machine's home folder"
        ),
    )
    parser.add_argument("--require-chain", action="store_true",
                        help="Refuse receipts that carry no certificate chain. "
                             "Off by default while receipts issued under the "
                             "flat model are still in customers' hands")
    args = parser.parse_args()
    archive_revision = require_archive_source() if args.owner_test else None
    for name in (args.key_id or []):
        if not name.strip() or len(name) > 80:
            parser.error("invalid key id")
    if args.key_id and not args.public_key:
        parser.error("--key-id names a --public-key; supply the key it names")
    for name in (args.root_id or []):
        if not name.strip() or len(name) > 80:
            parser.error("invalid root id")
    if args.root_id and not args.root_key:
        parser.error("--root-id names a --root-key; supply the key it names")
    if args.root_key and not args.root_id:
        parser.error("--root-key needs a --root-id: a root nobody named is a "
                     "root nobody can audit")
    if args.trust_list and not args.root_key:
        parser.error("--trust-list needs the --root-key roots it is signed by")
    try:
        # Pair them here, before a wheel is built, so a mismatched or
        # duplicated pair is a command-line error rather than a half-built
        # output directory. The trust list is proved against the roots in the
        # same pass, for the same reason.
        issuer_keys(args.public_key, args.key_id)
        roots = root_keys(args.root_key, args.root_id)
        if args.require_chain and not roots:
            parser.error("--require-chain needs --root-key")
        if args.trust_list is not None:
            load_trust_list(args.trust_list, roots)
    except ValueError as problem:
        parser.error(str(problem))
    editions = (("community", "business")
                if args.edition == "both" else (args.edition,))
    built = build_set(
        editions, args.output.resolve(), args.release_date, args.public_key,
        args.key_id, args.channel, args.root_key, args.root_id,
        args.trust_list, args.require_chain, args.owner_test,
    )
    if archive_revision:
        print(f"archive source: {archive_revision}")
    for edition in editions:
        for artifact in built[edition]:
            print(artifact)


if __name__ == "__main__":
    main()
