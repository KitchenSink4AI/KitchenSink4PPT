"""Explicit policy injection for repository developer scripts.

A script that imports the editing core is doing what a launcher does, so it
has to choose an edition on purpose rather than have one default into place.
Measurement and counting scripts choose Community, because Community is the
edition whose published numbers they produce.

The Community adapter is not vendored here. It is materialised from the one
canonical shared source (see adapter_source.py), which is also what the build
and the test suite consume.

Import this before importing kitchensink4ppt.server.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

import adapter_source  # noqa: E402


def use_community() -> None:
    if str(ROOT / "src") not in sys.path:
        sys.path.insert(0, str(ROOT / "src"))
    adapter_source.install_import_path()
    from ks4_license_community.policy import CommunityPolicy, ProductConfig
    from kitchensink4ppt import __version__, license_policy
    if not license_policy.installed():
        license_policy.install(CommunityPolicy(
            ProductConfig(product_id="kitchensink4ppt",
                          display_name="KitchenSink4PPT"), __version__))


def main() -> None:
    """Run the source checkout under an explicit Community policy."""
    use_community()
    from kitchensink4ppt.server import main as server_main
    server_main()


if __name__ == "__main__":
    main()
