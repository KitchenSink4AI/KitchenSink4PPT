"""The licensing seam: injection, setup errors, adapter isolation, composition.

Two properties are under test here and they are deliberately separate. A core
without a policy is a BROKEN INSTALLATION and must say so before a tool
exists to call. A Business receipt that is missing or unverifiable is a
REPORTED PROPERTY of a working installation and blocks nothing. Confusing
the two is the failure this architecture exists to prevent.

A third property is PowerPoint's alone: this core emits the licensing text
BEFORE its product prose, where its siblings emit it after. That divergence
is carried entirely by the core's assembly of two labelled components, with
no special case in the shared adapter, and it is asserted here.
"""
import ast
import subprocess
import sys
from pathlib import Path

import pytest

import adapter_source
from kitchensink4ppt import license_policy

ROOT = Path(__file__).resolve().parents[2]
ADAPTERS = adapter_source.ensure()
BUSINESS_SRC = ADAPTERS / adapter_source.BUSINESS_POLICY_DIST / "src"
COMMUNITY_SRC = ADAPTERS / adapter_source.COMMUNITY_POLICY_DIST / "src"

PPT_CONFIG = dict(product_id="kitchensink4ppt", display_name="KitchenSink4PPT")
PPT_BUSINESS = dict(product_id="kitchensink4ppt-business",
                    display_name="KitchenSink4PPT Business",
                    license_file_env="KS4P_LICENSE_FILE")
PPT_BUILD = {"edition": "business", "package": "kitchensink4ppt-business",
             "version": "1.2.2+business", "release_date": "2026-09-10",
             "public_keys": {}}


class _Good:
    policy_schema = license_policy.POLICY_SCHEMA
    edition_id = "community"
    product_id = "kitchensink4ppt"
    distribution_version = "0.0.0"

    def describe_license(self, redacted: bool = True):
        return {"edition": "community"}

    def initialization_text(self):
        return {"introduction": "I. ", "instructions": "L."}


def _variant(**members):
    return type("Variant", (_Good,), members)()


# ------------------------------------------------------------ setup errors


def test_a_core_without_a_policy_refuses_before_any_tool_registers(tmp_path):
    """Imported with no policy installed, the server module raises and no
    tool object is ever created. Run in a subprocess because the suite's own
    conftest injects a policy for every other test.

    This also covers a stale ppt-mcp.exe from a pre-split install, which
    pointed straight at kitchensink4ppt.server:main; that path now ends
    here, in a named setup error, instead of starting an unlabelled server
    in a mixed environment.
    """
    script = tmp_path / "probe.py"
    script.write_text(
        "import sys\n"
        f"sys.path[:0] = {[str(COMMUNITY_SRC), str(ROOT / 'src')]!r}\n"
        "from kitchensink4ppt import license_policy\n"
        "import kitchensink4ppt.server as server\n"
        "print('SERVER_IMPORTED', hasattr(server, 'mcp'))\n"
        "try:\n"
        "    server.create_server()\n"
        "except license_policy.PolicySetupError as problem:\n"
        "    print('SETUP_ERROR', problem)\n"
        "else:\n"
        "    print('NO_ERROR')\n", encoding="utf-8")
    done = subprocess.run([sys.executable, "-X", "utf8", str(script)],
                          capture_output=True, text=True, stdin=subprocess.DEVNULL)
    assert done.returncode == 0, done.stderr
    assert "SETUP_ERROR" in done.stdout
    assert "SERVER_IMPORTED False" in done.stdout
    # The message has to tell an operator what to do, not just what failed.
    assert "launcher" in done.stdout
    assert "kitchensink4ppt-business" in done.stdout


def test_setup_error_is_not_a_receipt_error():
    assert issubclass(license_policy.PolicySetupError, RuntimeError)
    from ks4_license_business.receipts import ReceiptError
    assert not issubclass(ReceiptError, license_policy.PolicySetupError)
    assert not issubclass(license_policy.PolicySetupError, ValueError)


@pytest.mark.parametrize("missing", ["policy_schema", "edition_id", "product_id",
                                     "distribution_version", "describe_license",
                                     "initialization_text"])
def test_an_incomplete_policy_is_a_setup_error(missing):
    with pytest.raises(license_policy.PolicyContractError, match=missing):
        license_policy._validate(_variant(**{missing: None}))


def test_a_mismatched_policy_schema_is_a_setup_error():
    with pytest.raises(license_policy.PolicyContractError, match="schema"):
        license_policy._validate(_variant(policy_schema=999))


def test_initialization_text_must_carry_both_labelled_components():
    broken = _variant(initialization_text=lambda self: {"introduction": "x"})
    with pytest.raises(license_policy.PolicyContractError, match="instructions"):
        license_policy._validate(broken)


# ------------------------------------- the policy contract checks exact types


def test_a_boolean_policy_schema_is_refused():
    """True == 1, so an isinstance(schema, int) check accepted policy_schema
    =True and then compared equal to schema 1. Types are checked exactly."""
    with pytest.raises(license_policy.PolicyContractError, match="non-integer"):
        license_policy._validate(_variant(policy_schema=True))
    assert license_policy.installed() is True  # nothing was swapped in


@pytest.mark.parametrize("member,value", [
    ("edition_id", 42), ("product_id", 17), ("distribution_version", 0),
])
def test_integer_identity_members_are_refused(member, value):
    with pytest.raises(license_policy.PolicyContractError,
                       match="not a non-empty string"):
        license_policy._validate(_variant(**{member: value}))


@pytest.mark.parametrize("text", [
    {"introduction": None, "instructions": []},
    {"introduction": "ok", "instructions": []},
    {"introduction": 3, "instructions": "ok"},
    {"introduction": "", "instructions": "ok"},
    {"introduction": "   ", "instructions": "ok"},
])
def test_non_string_initialization_components_are_refused(text):
    with pytest.raises(license_policy.PolicyContractError,
                       match="not a non-empty string"):
        license_policy._validate(_variant(initialization_text=lambda self: text))


def test_a_policy_with_every_member_mistyped_is_refused_by_install():
    """A policy whose members all have the wrong types is refused at
    install() rather than accepted."""
    class Probe:
        policy_schema = True
        edition_id = 42
        product_id = 17
        distribution_version = 0

        def describe_license(self, redacted: bool = True):
            return {}

        def initialization_text(self):
            return {"introduction": None, "instructions": []}

    previous = license_policy.active()
    with pytest.raises(license_policy.PolicyContractError):
        license_policy.install(Probe())
    assert license_policy.active() is previous  # the good policy still stands


def test_initialization_text_that_is_not_a_dict_is_refused():
    broken = _variant(initialization_text=lambda self: "just a string")
    with pytest.raises(license_policy.PolicyContractError, match="must return a dict"):
        license_policy._validate(broken)


def test_an_unknown_edition_id_is_refused():
    with pytest.raises(license_policy.PolicyContractError, match="unknown edition_id"):
        license_policy._validate(_variant(edition_id="enterprise"))


@pytest.mark.parametrize("hook", ["additional_tool_registration",
                                  "update_check_suppression",
                                  "update_backend"])
def test_a_non_callable_optional_hook_is_refused(hook):
    with pytest.raises(license_policy.PolicyContractError, match=hook):
        license_policy._validate(_variant(**{hook: "not callable"}))


def test_the_two_setup_error_classes_are_distinguishable():
    """Different operator actions: start through a launcher, versus install a
    matched core and adapter."""
    assert issubclass(license_policy.PolicyMissingError, license_policy.PolicySetupError)
    assert issubclass(license_policy.PolicyContractError, license_policy.PolicySetupError)
    assert not issubclass(license_policy.PolicyContractError,
                          license_policy.PolicyMissingError)


def test_core_never_imports_a_default_policy():
    """No module in the core may import either adapter. The core receives a
    prepared object or it fails; it does not find one."""
    offenders = []
    for path in (ROOT / "src").rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"), str(path))
        for node in ast.walk(tree):
            names = []
            if isinstance(node, ast.Import):
                names = [alias.name for alias in node.names]
            elif isinstance(node, ast.ImportFrom) and node.module:
                names = [node.module]
            for name in names:
                if name.split(".")[0] in {"ks4_license_community",
                                          "ks4_license_business"}:
                    offenders.append(f"{path.relative_to(ROOT)}: {name}")
    assert offenders == []


def test_no_licensing_prose_survives_in_the_core():
    """The guidance moved; it did not get copied."""
    from ks4_license_community import notices
    for needle in notices.guidance_needles():
        for path in (ROOT / "src").rglob("*.py"):
            assert needle not in path.read_text(encoding="utf-8"), path


def test_the_core_no_longer_carries_an_edition_module():
    """edition.py and launch.py were deleted, not emptied."""
    assert not (ROOT / "src/kitchensink4ppt/edition.py").exists()
    assert not (ROOT / "src/kitchensink4ppt/launch.py").exists()


# -------------------------------------------------------- adapter isolation


def _imports(root: Path) -> set[str]:
    found = set()
    for path in root.rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"), str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                found.update(alias.name.split(".")[0] for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                found.add(node.module.split(".")[0])
    return found


def test_business_package_imports_no_community_module():
    assert "ks4_license_community" not in _imports(BUSINESS_SRC)


def test_business_package_does_not_depend_on_the_editing_core():
    """Business must have no transitive dependency on Community, and the
    adapters reach the contract by shape rather than by import, so neither
    one can pull an editing surface in behind it."""
    assert "kitchensink4ppt" not in _imports(BUSINESS_SRC)


def test_community_package_cannot_reach_a_deck_api():
    reached = _imports(COMMUNITY_SRC)
    assert "kitchensink4ppt" not in reached
    assert "ks4_license_business" not in reached
    assert not (reached & {"pptx", "lxml", "svgelements", "win32com"})


def test_community_package_declares_no_dependencies():
    import tomllib
    config = tomllib.loads(
        (ADAPTERS / adapter_source.COMMUNITY_POLICY_DIST / "pyproject.toml")
        .read_text(encoding="utf-8"))
    assert config["project"]["dependencies"] == []


def test_business_package_imports_cleanly_without_community_installed(tmp_path):
    script = tmp_path / "probe.py"
    script.write_text(
        "import sys\n"
        f"sys.path[:0] = {[str(BUSINESS_SRC)]!r}\n"
        "import ks4_license_business\n"
        "print('COMMUNITY_LOADED', any(m.startswith('ks4_license_community') "
        "for m in sys.modules))\n"
        "print('CORE_LOADED', any(m.startswith('kitchensink4ppt') for m in sys.modules))\n",
        encoding="utf-8")
    done = subprocess.run([sys.executable, "-X", "utf8", str(script)],
                          capture_output=True, text=True, stdin=subprocess.DEVNULL)
    assert done.returncode == 0, done.stderr
    assert "COMMUNITY_LOADED False" in done.stdout
    assert "CORE_LOADED False" in done.stdout


# ------------------------------------------------------- text composition


def test_labelled_components_survive_either_assembly_order():
    """Word, Excel and Web put the licensing text last; PowerPoint puts it
    ahead of the product prose. One contract has to serve both without the
    adapter knowing which core it is talking to."""
    from ks4_license_community.policy import CommunityPolicy, ProductConfig
    text = CommunityPolicy(ProductConfig(**PPT_CONFIG), "1.2.2").initialization_text()
    prose = "PRODUCT PROSE."
    word_order = text["introduction"] + prose + " " + text["instructions"]
    ppt_order = text["introduction"] + text["instructions"] + " " + prose
    assert word_order.index(prose) < word_order.index(text["instructions"])
    assert ppt_order.index(text["instructions"]) < ppt_order.index(prose)
    assert text["introduction"] == word_order[:len(text["introduction"])]
    assert text["introduction"] == ppt_order[:len(text["introduction"])]


@pytest.mark.parametrize("edition", ["community", "business"])
def test_this_core_composes_the_powerpoint_order(edition, monkeypatch):
    """THE PowerPoint divergence, asserted against what the server actually
    emits rather than against an idea of it.

    The order is introduction, then the licensing instructions, then the
    product prose. It was PowerPoint's order before the split and it is
    PowerPoint's order after, and the shared adapter has no idea either way:
    it hands over two labelled components and this core places them.
    """
    from kitchensink4ppt import server
    if edition == "community":
        from ks4_license_community.policy import CommunityPolicy, ProductConfig
        policy = CommunityPolicy(ProductConfig(**PPT_CONFIG), "1.2.2")
    else:
        from ks4_license_business.policy import BusinessPolicy, ProductConfig
        policy = BusinessPolicy(ProductConfig(**PPT_BUSINESS), PPT_BUILD)
    text = policy.initialization_text()
    emitted = (text["introduction"] + text["instructions"] + " "
               + "PowerPoint (.pptx) editor:")
    # The shipped Community assembly is what the server module built at
    # import; both editions are checked against the same recipe.
    if edition == "community":
        assert server.create_server().instructions.startswith(emitted)
    prose_at = emitted.index("PowerPoint (.pptx) editor:")
    assert emitted.index(text["instructions"]) < prose_at
    assert emitted.startswith(text["introduction"])


def test_the_shipped_instructions_put_licensing_before_the_product_prose():
    from kitchensink4ppt import server
    instructions = server.create_server().instructions
    licensing = license_policy.active().initialization_text()["instructions"]
    assert licensing in instructions
    assert instructions.index(licensing) < instructions.index(
        "PowerPoint (.pptx) editor:")


def test_additional_tool_registration_is_optional_and_neither_policy_uses_it():
    from ks4_license_business.policy import BusinessPolicy, ProductConfig
    from ks4_license_community.policy import CommunityPolicy
    from ks4_license_community.policy import ProductConfig as CommunityConfig
    community = CommunityPolicy(CommunityConfig(**PPT_CONFIG), "1.2.2")
    business = BusinessPolicy(ProductConfig(**PPT_BUSINESS), PPT_BUILD)
    assert not hasattr(community, "additional_tool_registration")
    assert not hasattr(business, "additional_tool_registration")
    previous = license_policy.active()
    try:
        license_policy.install(community)
        license_policy.register_additional_tools(object())  # declines, no error
    finally:
        license_policy.install(previous)


def test_server_calls_the_optional_policy_registration_hook(monkeypatch):
    from kitchensink4ppt import server

    seen = []
    monkeypatch.setattr(
        license_policy, "register_additional_tools", lambda mcp: seen.append(mcp)
    )
    built = server.create_server()
    assert seen == [built]


def test_both_shipped_policies_satisfy_the_tightened_validator():
    from ks4_license_business.policy import BusinessPolicy, ProductConfig
    from ks4_license_community.policy import CommunityPolicy
    from ks4_license_community.policy import ProductConfig as CommunityConfig
    license_policy._validate(CommunityPolicy(CommunityConfig(**PPT_CONFIG), "1.2.2"))
    license_policy._validate(BusinessPolicy(ProductConfig(**PPT_BUSINESS), PPT_BUILD))


def test_only_the_business_policy_declares_an_update_backend():
    from ks4_license_business.policy import BusinessPolicy, ProductConfig
    from ks4_license_community.policy import CommunityPolicy
    from ks4_license_community.policy import ProductConfig as CommunityConfig
    community = CommunityPolicy(CommunityConfig(**PPT_CONFIG), "1.2.2")
    assert not hasattr(community, "update_check_suppression")
    assert not hasattr(community, "update_backend")
    business = BusinessPolicy(ProductConfig(**PPT_BUSINESS), PPT_BUILD)
    # The Business adapter dropped its suppression return when it gained a
    # channel of its own. Declaring both would report the check as off
    # while also knowing how to run it, which is a mismatched pair.
    assert not hasattr(business, "update_check_suppression")
    assert business.update_backend().backend_id == "business"
