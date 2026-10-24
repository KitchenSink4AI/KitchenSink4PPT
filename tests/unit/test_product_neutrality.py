"""ONE shared adapter, naming no product.

The defect this file guards against is a packaging one rather than a
behavioural one. Four products cannot each publish a wheel called
kitchensink4-license-business at the same version with different payloads
inside; pip, its cache and any co-installation would then resolve one name to
four different things. So the adapters carry no product identity at all, the
launchers (whose distributions ARE product-specific) supply it, and this
repository holds no copy of the adapters to drift from.

Three properties follow and all three are tested here: one adapter wheel
serves every editor, a receipt still cannot wander between products, and
PowerPoint consumes the canonical shared source rather than a fork of it.
"""
import ast
import json
import tomllib
from pathlib import Path

import pytest

import adapter_source
from ks4_license_business import receipts
from ks4_license_business.policy import (BuildMetadataError, BusinessPolicy,
                                         ProductConfig, ProductConfigError)
from ks4_license_community import notices
from ks4_license_community.policy import CommunityPolicy
from ks4_license_community.policy import ProductConfig as CommunityConfig

ROOT = Path(__file__).resolve().parents[2]
PACKAGES = ROOT / "packages"
CANONICAL = adapter_source.ensure()
ADAPTERS = {
    "business": CANONICAL / adapter_source.BUSINESS_POLICY_DIST / "src/ks4_license_business",
    "community": CANONICAL / adapter_source.COMMUNITY_POLICY_DIST / "src/ks4_license_community",
}
FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "receipts"
MANIFEST = json.loads((FIXTURES / "MANIFEST.json").read_text(encoding="utf-8"))

#: Every marker that would make a shared wheel product-specific: brand names,
#: PyPI names, receipt env prefixes and core module names. Used by BOTH the
#: Python scan and the notice-file scan, because neutral code with
#: product-specific notices is still a product-specific wheel.
PRODUCT_MARKERS = (
    "kitchensink4word", "kitchensink4xl", "kitchensink4ppt", "kitchensink4web",
    "KitchenSink4Word", "KitchenSink4XL", "KitchenSink4PPT", "KitchenSink4Web",
    "KS4W_", "KS4XL_", "KS4P_", "KS4WEB_", "word_mcp", "xlsx_mcp",
)

PPT = ProductConfig(product_id="kitchensink4ppt-business",
                    display_name="KitchenSink4PPT Business",
                    license_file_env="KS4P_LICENSE_FILE")
WORD = ProductConfig(product_id="kitchensink4word-business",
                     display_name="KitchenSink4Word Business",
                     license_file_env="KS4W_LICENSE_FILE")


def _build(package: str, version: str = "1.2.2+business") -> dict:
    return {"edition": "business", "package": package, "version": version,
            "release_date": "2026-09-10",
            "public_keys": MANIFEST["build"]["public_keys"]}


# ------------------------------------- one shared source


def test_this_repository_contains_no_adapter_copy():
    """Vendoring the adapters here is the defect, not a convenience. The
    build refuses too; this is the cheap static half of the same check."""
    assert list(PACKAGES.glob("kitchensink4-license-*")) == []
    assert not list(ROOT.rglob("ks4_license_community"))
    assert not list(ROOT.rglob("ks4_license_business"))


def test_the_adapters_in_use_are_the_pinned_canonical_revision():
    """The only mechanical evidence that this build consumes the shared
    source rather than a fork: same name, same version, proven same bytes."""
    assert adapter_source.tree_digest(CANONICAL) == adapter_source.ADAPTER_DIGEST
    assert len(adapter_source.PINNED_COMMIT) == 40


def test_the_canonical_source_is_read_from_a_commit_not_a_working_tree():
    """An uncommitted edit in somebody's checkout must not be able to become
    a build input."""
    source = (Path(adapter_source.__file__)).read_text(encoding="utf-8")
    tree = ast.parse(source)
    materialize = next(n for n in tree.body if isinstance(n, ast.FunctionDef)
                       and n.name == "materialize")
    dumped = ast.dump(materialize)
    assert "PINNED_COMMIT" in dumped
    assert "'show'" in dumped or '"show"' in dumped


def test_a_drifted_adapter_source_is_refused(tmp_path, monkeypatch):
    """Edit the materialised adapter and everything that consumes it stops."""
    import shutil
    fork = tmp_path / "forked"
    shutil.copytree(CANONICAL, fork)
    target = fork / adapter_source.COMMUNITY_POLICY_DIST / "src/ks4_license_community/notices.py"
    target.write_text(target.read_text(encoding="utf-8") + "\n# local edit\n",
                      encoding="utf-8")
    monkeypatch.setenv(adapter_source.SOURCE_ENV, str(fork))
    with pytest.raises(adapter_source.AdapterSourceError, match="not the pinned"):
        adapter_source.ensure()


# ------------------------------------------- no product names in the wheels


@pytest.mark.parametrize("edition", ["business", "community"])
def test_no_product_name_appears_anywhere_in_a_shared_adapter(edition):
    """The literal check on the Python payload."""
    offenders = []
    for path in ADAPTERS[edition].rglob("*.py"):
        if "__pycache__" in path.parts:
            continue
        text = path.read_text(encoding="utf-8")
        for name in PRODUCT_MARKERS:
            if name in text:
                offenders.append(f"{path.name}: {name}")
    assert offenders == []


@pytest.mark.parametrize("edition", ["business", "community"])
def test_no_product_name_appears_in_a_shared_adapter_notice_or_license(edition):
    """Neutral Python with Word-named NOTICE.md and LICENSE files is still a
    product-specific wheel. One shared wheel would have shipped one product's
    legal notices to customers of all four. Every non-directory file in the
    package is read, not only the .py ones, because that is the hole a
    payload-only scan leaves open."""
    package = ADAPTERS[edition].parents[1]
    offenders = []
    for path in package.rglob("*"):
        if not path.is_file() or "__pycache__" in path.parts:
            continue
        text = path.read_text(encoding="utf-8", errors="replace")
        for name in PRODUCT_MARKERS:
            if name in text:
                offenders.append(f"{path.relative_to(package).as_posix()}: {name}")
    assert offenders == []


def test_the_per_product_distributions_keep_their_own_notices():
    """The opposite direction. Neutralising the SHARED adapters must not
    neutralise the product distributions, whose notices are supposed to name
    the product a customer actually installed."""
    relatives = ["kitchensink4ppt-community/NOTICE.md"]
    if (PACKAGES / "kitchensink4ppt-business").is_dir():
        # The Business distribution is not in the Community source tree.
        relatives.append("kitchensink4ppt-business/LICENSE")
    for relative in relatives:
        text = (PACKAGES / relative).read_text(encoding="utf-8")
        assert "KitchenSink4PPT" in text, relative
    # The Community launcher's LICENSE is the plain AGPL-3.0 text, byte for
    # byte the repository's own; the copyright notice lives in NOTICE.md.
    assert ((PACKAGES / "kitchensink4ppt-community/LICENSE").read_bytes()
            == (ROOT / "LICENSE").read_bytes())
    # And the product-specific trademark sentence that left the shared notice
    # is still here, where it belongs.
    root_notice = (ROOT / "NOTICE.md").read_text(encoding="utf-8")
    assert "Microsoft, PowerPoint, GitHub and Copilot are trademarks" in root_notice


def test_the_shared_notices_still_carry_their_obligations():
    """Neutral naming is a substitution, not a deletion. These obligations do
    not change with whose name sits on the heading line."""
    community = ADAPTERS["community"].parents[1]
    notice = (community / "NOTICE.md").read_text(encoding="utf-8")
    agpl = (community / "LICENSE").read_text(encoding="utf-8")
    commercial = (ADAPTERS["business"].parents[1] / "LICENSE").read_text(
        encoding="utf-8")
    # Notices are hard-wrapped prose, so obligations are checked against the
    # unwrapped text rather than against one particular line break.
    flat = " ".join(notice.split())
    assert "Copyright (c) 2026 Alvut Consulting, LLC" in flat
    assert "KitchenSink4AI is a product line of Alvut Consulting, LLC" in flat
    assert ("Not affiliated with, endorsed by, or sponsored by Microsoft "
            "Corporation.") in flat
    assert "Required Notice:" in flat
    assert "GNU AFFERO GENERAL PUBLIC LICENSE" in agpl
    assert agpl.encode("utf-8") == (ROOT / "LICENSE").read_bytes()
    assert "Third-party components retain their" in commercial


@pytest.mark.parametrize("edition", ["business", "community"])
def test_a_shared_adapter_ships_no_generated_metadata(edition):
    """Generated per-product files would make the wheel product-specific
    while leaving its name and version alone, which is the packaging defect
    described at the top of this file."""
    assert list(ADAPTERS[edition].rglob("_edition.json")) == []
    config = tomllib.loads(
        (ADAPTERS[edition].parents[1] / "pyproject.toml").read_text(encoding="utf-8"))
    package_data = config.get("tool", {}).get("setuptools", {}).get("package-data", {})
    assert package_data == {}


@pytest.mark.business_sources
def test_the_business_launcher_owns_every_product_specific_value():
    """All four product-specific values live in the launcher, which is the
    distribution that is allowed to be product-specific."""
    import ks4ppt_business_launch as launch
    assert launch.PRODUCT_ID == "kitchensink4ppt-business"
    assert launch.DISPLAY_NAME == "KitchenSink4PPT Business"
    # KS4P_, not KS4PPT_. Customer-visible and carried forward unchanged.
    assert launch.LICENSE_FILE_ENV == "KS4P_LICENSE_FILE"
    source = ast.parse((PACKAGES / "kitchensink4ppt-business/src"
                        / "ks4ppt_business_launch/__init__.py")
                       .read_text(encoding="utf-8"))
    assert "_edition.json" in ast.dump(source)


def test_the_community_launcher_owns_its_product_values():
    import ks4ppt_community_launch as launch
    assert launch.PRODUCT_ID == "kitchensink4ppt"
    assert launch.DISPLAY_NAME == "KitchenSink4PPT"


# ------------------------------------------------ two products, one wheel


def test_two_products_coexist_in_one_business_environment():
    """The coexistence case: one adapter import, two configured products,
    neither disturbing the other."""
    ppt = BusinessPolicy(PPT, _build("kitchensink4ppt-business"))
    word = BusinessPolicy(WORD, _build("kitchensink4word-business", "2.1.3+business"))
    assert ppt.product_id == "kitchensink4ppt-business"
    assert word.product_id == "kitchensink4word-business"
    assert ppt.distribution_version == "1.2.2+business"
    assert word.distribution_version == "2.1.3+business"
    assert "KitchenSink4PPT" in ppt.initialization_text()["introduction"]
    assert "KitchenSink4Word" in word.initialization_text()["introduction"]
    assert type(ppt).__module__ == type(word).__module__


def test_two_products_read_different_receipt_environment_variables(tmp_path,
                                                                   monkeypatch):
    receipt = (FIXTURES / "covered.json").read_bytes()
    path = tmp_path / "ppt-receipt.json"
    path.write_bytes(receipt)
    monkeypatch.setenv("KS4P_LICENSE_FILE", str(path))
    monkeypatch.delenv("KS4W_LICENSE_FILE", raising=False)
    ppt = BusinessPolicy(PPT, _build("kitchensink4ppt-business"))
    word = BusinessPolicy(WORD, _build("kitchensink4word-business"))
    assert ppt.describe_license()["status"] == "verified"
    # The Word policy does not see the PowerPoint receipt: different variable.
    assert word.describe_license()["reason"] == "not_configured"


def test_a_community_adapter_introduces_whatever_product_it_is_given():
    ppt = CommunityPolicy(CommunityConfig(product_id="kitchensink4ppt",
                                          display_name="KitchenSink4PPT"), "1.2.2")
    word = CommunityPolicy(CommunityConfig(product_id="kitchensink4word",
                                           display_name="KitchenSink4Word"), "2.1.3")
    assert ppt.initialization_text()["introduction"] == (
        "KitchenSink4PPT (kitchensink4ppt on PyPI), part of the "
        "KitchenSink4AI suite. ")
    assert word.initialization_text()["introduction"] == (
        "KitchenSink4Word (kitchensink4word on PyPI), part of the "
        "KitchenSink4AI suite. ")
    # The guidance itself is product-independent and identical.
    assert (ppt.initialization_text()["instructions"]
            == word.initialization_text()["instructions"]
            == notices.COMMUNITY_GUIDANCE)


# -------------------------------------------- receipts cannot cross products


def test_a_ppt_receipt_is_refused_by_a_word_installation():
    raw = (FIXTURES / "covered.json").read_bytes()   # product: ppt-business
    build = _build("kitchensink4word-business")
    with pytest.raises(receipts.ReceiptError) as problem:
        receipts.verify_receipt(raw, build, WORD.product_id)
    assert problem.value.category == "product_mismatch"


def test_a_word_receipt_is_refused_by_a_ppt_installation():
    raw = (FIXTURES / "wrong_product.json").read_bytes()  # product: word-business
    with pytest.raises(receipts.ReceiptError) as problem:
        receipts.verify_receipt(raw, MANIFEST["build"], PPT.product_id)
    assert problem.value.category == "product_mismatch"
    # ...and the same bytes verify against the installation they belong to.
    assert receipts.verify_receipt(raw, _build("kitchensink4word-business"),
                                   WORD.product_id)["status"] == "verified"


def test_a_receipt_cannot_nominate_the_product_it_is_checked_against(tmp_path,
                                                                     monkeypatch):
    """Self-selection is the failure mode parameterisation could have
    introduced. The expected product comes from configuration, and a receipt
    claiming a different one is refused rather than believed."""
    raw = (FIXTURES / "wrong_product.json").read_bytes()
    path = tmp_path / "receipt.json"
    path.write_bytes(raw)
    monkeypatch.setenv("KS4P_LICENSE_FILE", str(path))
    status = BusinessPolicy(PPT, _build("kitchensink4ppt-business")).describe_license()
    assert (status["status"], status["reason"]) == ("unverified", "product_mismatch")


@pytest.mark.parametrize("expected", [None, "", "   ", 17, b"ppt"])
def test_verification_refuses_to_run_without_a_configured_product(expected):
    """There is no permissive default to fall into."""
    raw = (FIXTURES / "covered.json").read_bytes()
    with pytest.raises(ValueError, match="expected product"):
        receipts.verify_receipt(raw, MANIFEST["build"], expected)


def test_build_metadata_must_agree_with_the_launcher_configuration():
    with pytest.raises(BuildMetadataError, match="different product"):
        BusinessPolicy(PPT, _build("kitchensink4word-business"))


@pytest.mark.parametrize("bad", [
    {"product_id": "", "display_name": "X", "license_file_env": "Y"},
    {"product_id": "x", "display_name": "  ", "license_file_env": "Y"},
    {"product_id": "x", "display_name": "X", "license_file_env": 42},
])
def test_a_malformed_product_config_is_refused_at_construction(bad):
    with pytest.raises(ProductConfigError):
        ProductConfig(**bad)


def test_a_business_policy_refuses_a_bare_dict_for_configuration():
    with pytest.raises(ProductConfigError, match="ProductConfig"):
        BusinessPolicy({"product_id": "kitchensink4ppt-business"},
                       _build("kitchensink4ppt-business"))
