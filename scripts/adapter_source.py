"""Resolve the ONE canonical shared license policy adapter source.

ONE SOURCE, NEVER FORKED. The two license policy
adapters, `kitchensink4-license-community` and `kitchensink4-license-business`,
are ONE shared source serving every editor in the KitchenSink4AI suite. They
are never forked, never copied into a product repository, and never edited
here. Four products cannot each publish a wheel called
`kitchensink4-license-business` at one version with four different payloads
inside; pip, its cache and any co-installation would resolve one name to four
different things.

This repository therefore contains NO adapter sources. The build, the smoke
harness and the test suite all materialise them from the canonical revision
named below, and a recorded content digest proves that what was materialised
is that revision and not something that drifted.

WHERE THE CANONICAL SOURCE LIVES TODAY.
The adapters were authored in the Word editions repository and are pinned
here at commit PINNED_COMMIT. That is a
deliberate interim: the adapters have no repository of their own yet and no
published wheels yet. Relocating them (to their own repository, or to
published `kitchensink4-license-*` distributions on PyPI) is a separate
change, not one to make inside a product repository. When they move, exactly
two constants below change, plus ADAPTER_DIGEST if the contents change with
the move.

PIN NOTES (what each pinned revision brought).
  5bb0f20  Product-neutral Python, but the packaging notice files
           still named KitchenSink4Word, so one shared wheel would have
           shipped one product's legal notices to customers of all four. The
           defect was found by consuming the adapters from a product
           repository, and fixed in the canonical source rather than
           forked around.
  5d3842b  The adapters' NOTICE.md and LICENSE files carry
           family naming ("KitchenSink4AI licensing adapters") instead of one
           product's. Mechanical substitution in already-approved text: the
           AGPL body and the commercial notice's operative paragraphs were
           not edited, only their heading lines. The Word-specific trademark
           sentence left the shared notice; the general Microsoft
           non-affiliation disclaimer stayed, and this product's own
           trademark line stays in its root NOTICE.md and core wheel.
  3f9873a  The approved warning loses its em dash, because no em dash
           belongs in served or brand text.
           The clause break in "require a paid license ... purchase at"
           becomes a sentence break, so guidance_needles() yields eight
           needles where it yielded seven. Both new needles clear the
           40-character floor and neither collides with the commercial
           notice.
  59fdbd4  Two changes the self-fulfilling purchase
           flow needs. The Business build metadata may declare a release
           channel, and the adapter validates it against a closed set of two
           (beta, stable), so a mislabelled commercial artifact is a setup
           error instead of a string a customer reads. The receipt verifier
           selects the issuer public key BY THE ID THE RECEIPT NAMES out of
           the map the build carries, and refuses an unknown id with the
           existing signature_invalid category rather than letting a KeyError
           fall through to "unreadable". A build now carries several issuer
           keys, because the web tier signs self-serve receipts and the
           offline root signs manual and enterprise ones.

EDITIONS. KS4_ADAPTER_EDITION selects what is resolved. Unset (or "all")
is the private two-adapter path every Business build uses: both packages,
checked against ADAPTER_DIGEST, unchanged. "community" is the public path:
the public repositories carry and build the Community edition only and never
receive the Business adapter source, so only kitchensink4-license-community is
resolved, imported and digested, against COMMUNITY_DIGEST. Asking for the
Business adapter on that path is an error, not a fallback.

PUBLIC CI. The public repository reads the Community adapter from the
reviewed public release commit PUBLIC_ADAPTER_COMMIT of the license repository,
KitchenSink4AI/kitchensink4-license-community, the one repository that publishes
kitchensink4-license-community (`python scripts/adapter_source.py
--public-pin`). Until that commit exists the constant is empty and the pin
refuses, so no public run can use an unpinned or moving source.

PUBLISHED WHEEL. A release never uploads the adapter and never rebuilds it
for PyPI. It requires the exact wheel the publisher put on PyPI, pinned by
COMMUNITY_WHEEL_FILENAME and COMMUNITY_WHEEL_SHA256 (`python
scripts/adapter_source.py --fetch-wheel DIR` downloads and checks it). A
rebuild carries the build backend of its day, so its bytes could differ
from the published wheel's; the pin removes that.

Neutral code with product-specific notices is still a product-specific wheel,
so the neutrality assertions here cover FILES, not only Python. The
per-product distributions under packages/ keep their product-named notices,
and that direction is asserted too.
"""
from __future__ import annotations

import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

#: The canonical adapter revision: the product-neutral adapters with the
#: tightened contract validator, product-neutral packaging notices and the
#: em-dash-free approved warning. Anything earlier ships one
#: product's legal notices in a wheel shared by four, an em dash in served
#: text, or both.
#: 838fb02 supersedes a2d4af6. The update channel: the Business
#: adapter drops update_check_suppression and declares update_backend()
#: instead, returning a selector that names the licensed release channel
#: for the core to run. The Community adapter is unchanged except for the
#: comment recording that it implements neither hook.
#: df44aca supersedes 838fb02. The adapter distribution moves from 0.3.0 to
#: 0.4.0, and nothing else in the tree changes. The version is the point:
#: the previous revision removed update_check_suppression as well as adding
#: update_backend, so a core built against 0.3.0 calls a method that is gone
#: and falls back to the public PyPI channel on a Business install. Pinning
#: the version is what lets a build refuse that pairing.
#: 70e2a17 supersedes df44aca. The adapter source is unchanged; its own
#: LICENSE is not. The paragraph that called this a beta release
#: whose prices, eligibility, scope and support were not approved was replaced with the
#: paragraph that states the Business terms. That landed in the canonical
#: checkout and the pin stayed behind, so the sibling products materialised the
#: beta wording while the Word editor built the final one from its own tree, and
#: the release bundle build refused the bundle: one distribution, two payloads.
#: b2b404c: the Community adapter is the package of its own repository,
#: kitchensink4-license-community, which differs from the previous pin only in
#: comment wording and its PyPI author line. The Business adapter differs from
#: the previous pin only in comment wording.
#: 9c7dc45: the Community adapter is kitchensink4-license-community 0.2.0 (the plain AGPL
#: LICENSE text, trademark lines in NOTICE.md, tidier comments); the Business adapter is
#: unchanged.
PINNED_COMMIT = "9c7dc45404e294a63b685f4186db52769a7ed2e3"

#: Where the pinned revision is reachable from, relative to the product tree
#: that holds the editor repositories. Changing this is part of relocating
#: the adapters.
CANONICAL_REPO_RELATIVE = ("word-mcp",)

COMMUNITY_POLICY_DIST = "kitchensink4-license-community"
BUSINESS_POLICY_DIST = "kitchensink4-license-business"
ADAPTER_PACKAGES = (COMMUNITY_POLICY_DIST, BUSINESS_POLICY_DIST)

#: Content digest of the materialised canonical tree. Recorded rather than
#: computed, so a forked, edited or re-pointed adapter fails loudly here
#: instead of silently becoming this product's private copy.
ADAPTER_DIGEST = "70829e50d60509fc71daf0c49a3cdb0d4039e656ad6b8d4f1d46a009ad4fdac7"

#: Content digest of the Community adapter package ALONE at PINNED_COMMIT
#: (same algorithm, Community directory only). The Community-only path is
#: checked against this value and never against ADAPTER_DIGEST.
COMMUNITY_DIGEST = "ce7c493e264397b127fe07a0529469ea662f0b7f470af1f418afcd4d1ea2e1c5"

#: A pre-materialised canonical tree (a directory holding both package
#: directories). CI and offline builds set this.
SOURCE_ENV = "KS4PPT_ADAPTER_SOURCE"
#: A git checkout of the canonical repository, when it is not where the
#: default search looks.
REPO_ENV = "KS4PPT_ADAPTER_REPO"
#: "community" selects the Community-only path; unset or "all" keeps both.
ENV_EDITION = "KS4_ADAPTER_EDITION"
EDITIONS = ("all", "community")

#: Where public CI reads the Community adapter: the publisher's reviewed public
#: release commit carrying packages/kitchensink4-license-community.
#: Full 40-character SHA. Empty until that release exists; public_pin() refuses.
PUBLIC_ADAPTER_REPOSITORY = "KitchenSink4AI/kitchensink4-license-community"
PUBLIC_ADAPTER_COMMIT = "c33f1a4cda744bd4f95234993e37f03c27705278"

#: The exact kitchensink4-license-community wheel on PyPI that releases require,
#: as the publisher's post-verify job prints it. Empty until that release exists;
#: pinned_wheel() refuses.
COMMUNITY_WHEEL_FILENAME = "kitchensink4_license_community-0.2.0-py3-none-any.whl"
COMMUNITY_WHEEL_SHA256 = "4ad5bf6ae2a677a2b51b9c62b7f1a44464940350f8e8c9ce04635a91ea73ff81"
PYPI_JSON_URL = "https://pypi.org/pypi/{name}/{version}/json"
PYPI_FILES_HOST = "files.pythonhosted.org"


class AdapterSourceError(RuntimeError):
    """The canonical shared adapter source is unreachable or not canonical."""


def edition() -> str:
    """'all' (private, both adapters) or 'community' (public, Community only)."""
    value = (os.environ.get(ENV_EDITION) or "").strip().lower() or "all"
    if value not in EDITIONS:
        raise AdapterSourceError(
            f"{ENV_EDITION}={value!r} is not one of {', '.join(EDITIONS)}")
    return value


def required_dists() -> tuple[str, ...]:
    """The adapter packages the selected edition resolves."""
    if edition() == "community":
        return (COMMUNITY_POLICY_DIST,)
    return ADAPTER_PACKAGES


def expected_digest() -> str:
    return COMMUNITY_DIGEST if edition() == "community" else ADAPTER_DIGEST


def public_pin() -> dict:
    """Where public CI checks the Community adapter out from. Refuses unless pinned."""
    if not re.fullmatch(r"[0-9a-f]{40}", PUBLIC_ADAPTER_COMMIT):
        raise AdapterSourceError(
            "PUBLIC_ADAPTER_COMMIT is not a full 40-character commit SHA. Public "
            "builds read the Community adapter from the reviewed public "
            f"{PUBLIC_ADAPTER_REPOSITORY} release commit; set it (and "
            "COMMUNITY_DIGEST if the adapter content changed) in a reviewed edit.")
    return {"repository": PUBLIC_ADAPTER_REPOSITORY, "commit": PUBLIC_ADAPTER_COMMIT,
            "path": f"packages/{COMMUNITY_POLICY_DIST}", "digest": COMMUNITY_DIGEST}


def pinned_wheel() -> dict:
    """The exact published adapter wheel a release requires. Refuses unless pinned."""
    m = re.fullmatch(r"kitchensink4_license_community-([0-9]+(?:\.[0-9]+)*)-py3-none-any\.whl",
                     COMMUNITY_WHEEL_FILENAME)
    if not m or not re.fullmatch(r"[0-9a-f]{64}", COMMUNITY_WHEEL_SHA256):
        raise AdapterSourceError(
            "COMMUNITY_WHEEL_FILENAME and COMMUNITY_WHEEL_SHA256 are not set. Releases require "
            "the exact kitchensink4-license-community wheel its publisher put on PyPI; pin the "
            "values its post-verify job printed, in a reviewed edit.")
    return {"filename": COMMUNITY_WHEEL_FILENAME, "sha256": COMMUNITY_WHEEL_SHA256,
            "version": m.group(1)}


def _https_get(url: str) -> bytes:
    import urllib.request
    request = urllib.request.Request(url, headers={"User-Agent": "ks4-adapter-source/2"})
    with urllib.request.urlopen(request, timeout=60) as response:  # noqa: S310 (https only)
        return response.read()


def fetch_pinned_wheel(destination: Path, get=_https_get) -> Path:
    """Download the pinned wheel from PyPI into destination; refuse any other bytes."""
    import json
    import urllib.parse
    pin = pinned_wheel()
    meta = json.loads(get(PYPI_JSON_URL.format(name=COMMUNITY_POLICY_DIST, version=pin["version"])))
    entries = [f for f in meta.get("urls", []) if f.get("filename") == pin["filename"]]
    if len(entries) != 1:
        raise AdapterSourceError(
            f"{pin['filename']} is not a file of {COMMUNITY_POLICY_DIST} {pin['version']} on PyPI")
    url = entries[0].get("url", "")
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme != "https" or parsed.hostname != PYPI_FILES_HOST:
        raise AdapterSourceError(f"unexpected download URL {url!r}")
    if (entries[0].get("digests") or {}).get("sha256") != pin["sha256"]:
        raise AdapterSourceError("PyPI lists a different sha256 for the pinned wheel")
    data = get(url)
    if hashlib.sha256(data).hexdigest() != pin["sha256"]:
        raise AdapterSourceError("the downloaded wheel does not match COMMUNITY_WHEEL_SHA256")
    destination.mkdir(parents=True, exist_ok=True)
    target = destination / pin["filename"]
    target.write_bytes(data)
    return target


def package_dir(dist: str) -> Path:
    """One adapter package directory on the selected edition path."""
    if dist not in ADAPTER_PACKAGES:
        raise AdapterSourceError(f"{dist} is not a license policy adapter")
    if dist not in required_dists():
        raise AdapterSourceError(
            f"{dist} is not available on the {edition()} edition path "
            f"({ENV_EDITION}); the public repositories carry the Community adapter only")
    return ensure() / dist


def _run(args: list[str]) -> bytes:
    done = subprocess.run(args, capture_output=True, stdin=subprocess.DEVNULL)
    if done.returncode:
        raise AdapterSourceError(
            f"{' '.join(args[:3])} failed: "
            f"{done.stderr.decode('utf-8', 'replace').strip()}")
    return done.stdout


def locate_repo() -> Path:
    """Find the git checkout carrying the pinned canonical revision.

    The default search starts from the real repository rather than from
    ROOT, so it works identically from a git worktree in OS scratch and from
    the checkout itself.
    """
    override = os.environ.get(REPO_ENV)
    if override:
        return Path(override).resolve()
    common = _run(["git", "-C", str(ROOT), "rev-parse",
                   "--git-common-dir"]).decode().strip()
    real = (ROOT / common).resolve() if not Path(common).is_absolute() \
        else Path(common).resolve()
    # .../PowerPoint/editions/.git -> .../PowerPoint/editions -> the product tree
    product_tree = real.parent.parent
    return product_tree.joinpath(*CANONICAL_REPO_RELATIVE)


def _file_list(repo: Path) -> list[str]:
    listing = _run(["git", "-C", str(repo), "ls-tree", "-r", "--name-only",
                    PINNED_COMMIT, "--"]
                   + [f"packages/{name}" for name in required_dists()])
    paths = [line for line in listing.decode("utf-8").splitlines() if line]
    if not paths:
        raise AdapterSourceError(
            f"{repo} has no adapter packages at {PINNED_COMMIT[:12]}. The "
            "canonical shared adapter source is not reachable from here; set "
            f"{REPO_ENV} or {SOURCE_ENV}.")
    return paths


def materialize(destination: Path, repo: Path | None = None) -> Path:
    """Write the pinned adapter packages into destination.

    Read out of the commit object rather than a working tree, so an
    uncommitted edit in somebody's checkout cannot become the build input.
    """
    repo = repo or locate_repo()
    if not (repo / ".git").exists() and not repo.is_dir():
        raise AdapterSourceError(
            f"The canonical adapter repository is not at {repo}. Set "
            f"{REPO_ENV} to a checkout carrying {PINNED_COMMIT[:12]}, or "
            f"{SOURCE_ENV} to an already-materialised tree.")
    destination.mkdir(parents=True, exist_ok=True)
    for path in _file_list(repo):
        blob = _run(["git", "-C", str(repo), "show", f"{PINNED_COMMIT}:{path}"])
        # packages/<dist>/... -> <dist>/...
        target = destination / Path(path).relative_to("packages")
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(blob)
    return destination


def tree_digest(source: Path, packages: tuple[str, ...] | None = None) -> str:
    """Content digest over every adapter file, path included and sorted.
    Covers the selected edition's packages unless packages is given."""
    parts = []
    for package in (packages or required_dists()):
        root = source / package
        for path in sorted(p for p in root.rglob("*") if p.is_file()):
            if "__pycache__" in path.parts or path.suffix == ".pyc":
                continue
            parts.append((path.relative_to(source).as_posix(),
                          hashlib.sha256(path.read_bytes()).hexdigest()))
    if not parts:
        raise AdapterSourceError(f"No adapter files found under {source}")
    return hashlib.sha256(repr(sorted(parts)).encode()).hexdigest()


def _cache_root() -> Path:
    suffix = "-community" if edition() == "community" else ""
    return Path(tempfile.gettempdir()) / f"ks4-license-adapters-{PINNED_COMMIT[:12]}{suffix}"


def ensure(verify: bool = True) -> Path:
    """The canonical adapter tree, materialised once and reused.

    The cache lives in OS scratch, never inside a repository: a transient
    tree inside the vault goes through the backup mirror.
    """
    override = os.environ.get(SOURCE_ENV)
    if override:
        source = Path(override).resolve()
        missing = [n for n in required_dists() if not (source / n).is_dir()]
        if missing:
            raise AdapterSourceError(
                f"{SOURCE_ENV}={source} is missing {missing}.")
    else:
        source = _cache_root()
        if not all((source / name).is_dir() for name in required_dists()):
            staging = Path(tempfile.mkdtemp(prefix="ks4-adapters-"))
            try:
                materialize(staging)
                shutil.rmtree(source, ignore_errors=True)
                staging.replace(source)
            finally:
                shutil.rmtree(staging, ignore_errors=True)
    if verify:
        found = tree_digest(source)
        expected = expected_digest()
        if found != expected:
            raise AdapterSourceError(
                "The canonical shared adapter source is not the pinned "
                f"revision ({edition()} edition): expected digest {expected}, found {found}. "
                "Either the pin moved without this constant moving with it, "
                "or the adapters were edited outside the canonical source. "
                "Neither is something a product slice may resolve on its own.")
    return source


def source_paths(source: Path | None = None) -> list[str]:
    """The two import roots, for a sys.path a launcher would not need."""
    source = source or ensure()
    return [str(source / name / "src") for name in required_dists()]


def install_import_path(source: Path | None = None) -> Path:
    """Put the canonical adapters on sys.path for a developer process."""
    source = source or ensure()
    for entry in source_paths(source):
        if entry not in sys.path:
            sys.path.insert(0, entry)
    return source


def notices_module():
    """Load the Community notice module from the canonical source file.

    By file rather than by import, so the build gate cannot pass because the
    adapter happened not to be installed.
    """
    import importlib.util
    path = (ensure() / COMMUNITY_POLICY_DIST / "src" / "ks4_license_community"
            / "notices.py")
    spec = importlib.util.spec_from_file_location("_canonical_notices", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> None:
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--digest", action="store_true",
                        help="print the digest of the materialised tree")
    parser.add_argument("--public-pin", action="store_true",
                        help="print the pinned public Community adapter source")
    parser.add_argument("--fetch-wheel", metavar="DIR",
                        help="download the pinned published Community adapter wheel into DIR")
    parser.add_argument("--path", action="store_true",
                        help="print where the canonical tree was materialised")
    args = parser.parse_args()
    if args.public_pin:
        for key, value in public_pin().items():
            print(f"{key}={value}")
        return
    if args.fetch_wheel:
        print(fetch_pinned_wheel(Path(args.fetch_wheel)))
        return
    source = ensure(verify=not args.digest)
    if args.digest or not args.path:
        print(tree_digest(source))
    if args.path:
        print(source)


if __name__ == "__main__":
    main()
