"""Deleting an animated shape lays out the rest of the main sequence the way
PowerPoint does when it deletes the same shape.

Measured in PowerPoint 16 on made-up decks (Shape.Delete, then SaveCopyAs):

- An after-previous effect that followed the deleted shape's on-click
  effect at the start of the first click group starts with the slide: the
  group keeps delay="indefinite" and gains onBegin of the main sequence,
  and its slot starts at 0. The slide show then needs no click (it needed
  one before the fix).
- When the deleted effect led a later click group, what is left of that
  group joins the group before it, starting when that group's effects
  end.
- A slot that started at the end of the deleted effect (evt="onEnd" naming
  its time node) starts when the slots before it end; no condition is left
  naming a removed time node, and no hole is left where the effect was.
- Deleting an effect that nothing follows leaves the rest untouched.
"""

from __future__ import annotations

import pytest
from lxml import etree

from kitchensink4ppt.core.package import PptxPackage, qn
from kitchensink4ppt.ops import animations as an
from kitchensink4ppt.ops import shapes as shp
from kitchensink4ppt.ops.read import get_slide_info


def _deck(make_deck, seed):
    pkg = PptxPackage(make_deck(f"anim{seed}.pptx", seed=seed, extra_slides=0))
    info = get_slide_info(pkg, 0)
    a, b = [s["id"] for s in info["shapes"]][:2]
    c = shp.insert_shape(pkg, 0, "rect", 1, 5, 2, 1, text="Third")["shape_id"]
    return pkg, info["part"], a, b, c


def _effects(pkg, part):
    timing = pkg.root(part).find(qn("p:timing"))
    out = {}
    for ctn in timing.iter(qn("p:cTn")):
        if ctn.get("presetID") is not None:
            out[int(next(ctn.iter(qn("p:spTgt"))).get("spid"))] = ctn
    return out


def _chain_on_end(target_ctn, dependent_ctn, delay="0"):
    """Start the dependent effect's slot at the END of the target effect."""
    slot = dependent_ctn.getparent().getparent().getparent()
    st = slot.find(qn("p:stCondLst"))
    for c in list(st):
        st.remove(c)
    cond = etree.SubElement(st, qn("p:cond"), evt="onEnd", delay=delay)
    etree.SubElement(cond, qn("p:tn"), val=target_ctn.get("id"))


def _layout(pkg, part):
    """One string per click group: its start conditions, then per slot its
    start conditions and its effects (spid:nodeType). A condition naming a
    time node that is not in the tree shows as DANGLING."""
    timing = pkg.root(part).find(qn("p:timing"))
    roles = {}
    for ctn in timing.iter(qn("p:cTn")):
        if ctn.get("nodeType") == "mainSeq":
            roles[ctn.get("id")] = "mainSeq"
        elif ctn.get("presetID") is not None:
            roles[ctn.get("id")] = "eff" + next(ctn.iter(qn("p:spTgt"))).get("spid")

    def conds(ctn):
        out = []
        st = ctn.find(qn("p:stCondLst"))
        for c in st.findall(qn("p:cond")) if st is not None else []:
            tn = c.find(qn("p:tn"))
            s = c.get("evt", "")
            if tn is not None:
                s += "@" + roles.get(tn.get("val"), "DANGLING")
            out.append(f"{s}+{c.get('delay', '0')}")
        return "[" + ",".join(out) + "]"

    ms = next(c for c in timing.iter(qn("p:cTn")) if c.get("nodeType") == "mainSeq")
    lines = []
    for g in ms.find(qn("p:childTnLst")).findall(qn("p:par")):
        gctn = g.find(qn("p:cTn"))
        parts = ["G" + conds(gctn)]
        for s in gctn.find(qn("p:childTnLst")).findall(qn("p:par")):
            sctn = s.find(qn("p:cTn"))
            effs = [
                f"{next(e.iter(qn('p:spTgt'))).get('spid')}:{e.find(qn('p:cTn')).get('nodeType')}"
                for e in sctn.find(qn("p:childTnLst")).findall(qn("p:par"))
            ]
            parts.append("S" + conds(sctn) + "{" + " ".join(effs) + "}")
        lines.append(" ".join(parts))
    return lines


def _no_dangling(pkg, part):
    timing = pkg.root(part).find(qn("p:timing"))
    ids = {c.get("id") for c in timing.iter(qn("p:cTn"))}
    return [t.get("val") for t in timing.iter(qn("p:tn")) if t.get("val") not in ids]


def _delete_and_reload(pkg, part, spid):
    res = shp.delete_shape(pkg, 0, spid)
    path = pkg.save(do_backup=False)
    return res, PptxPackage(path)


@pytest.mark.parametrize("delay", ["0", "500"])
def test_after_previous_chained_to_the_deleted_leader_starts_with_the_slide(make_deck, delay):
    """The reported deck: B's slot starts at the end of A's effect; A is
    deleted. PowerPoint: B starts with the slide (onBegin), slot at 0."""
    pkg, part, a, b, _c = _deck(make_deck, 1)
    an.add_entrance_animation(pkg, 0, a, "fade")
    an.add_entrance_animation(pkg, 0, b, "appear", trigger="after_previous")
    e = _effects(pkg, part)
    _chain_on_end(e[a], e[b], delay)
    res, pkg2 = _delete_and_reload(pkg, part, a)
    assert _no_dangling(pkg2, part) == []
    assert _layout(pkg2, part) == [
        f"G[+indefinite,onBegin@mainSeq+0] S[+0]{{{b}:afterEffect}}"
    ]
    assert res["timing_report"]["retimed_shape_ids"] == [b]
    eff = an.list_animations(pkg2, 0)["effects"]
    assert [(x["shape_id"], x["group"], x["trigger"], x["delay_ms"]) for x in eff] == [
        (b, 0, "after_previous", 0)
    ]


def test_after_previous_left_leading_the_first_group_starts_with_the_slide(make_deck):
    """No event chain: B's slot started at 500 (A's end). PowerPoint: onBegin
    on the group and B's slot at 0."""
    pkg, part, a, b, _c = _deck(make_deck, 3)
    an.add_entrance_animation(pkg, 0, a, "fade")
    an.add_entrance_animation(pkg, 0, b, "appear", trigger="after_previous")
    _res, pkg2 = _delete_and_reload(pkg, part, a)
    assert _layout(pkg2, part) == [
        f"G[+indefinite,onBegin@mainSeq+0] S[+0]{{{b}:afterEffect}}"
    ]


@pytest.mark.parametrize("chained", [False, True])
def test_what_is_left_of_a_later_group_joins_the_group_before(make_deck, chained):
    """A on click, C on click, B after C; C is deleted. PowerPoint: one click
    group, B's slot at 500 after A, with or without an onEnd chain to C."""
    pkg, part, a, b, c = _deck(make_deck, 4)
    an.add_entrance_animation(pkg, 0, a, "fade")
    an.add_entrance_animation(pkg, 0, c, "fade")
    an.add_entrance_animation(pkg, 0, b, "appear", trigger="after_previous")
    if chained:
        e = _effects(pkg, part)
        _chain_on_end(e[c], e[b])
    res, pkg2 = _delete_and_reload(pkg, part, c)
    assert _no_dangling(pkg2, part) == []
    assert _layout(pkg2, part) == [
        f"G[+indefinite] S[+0]{{{a}:clickEffect}} S[+500]{{{b}:afterEffect}}"
    ]
    assert res["timing_report"]["retimed_shape_ids"] == [b]


def test_a_chain_through_the_deleted_effect_closes_up(make_deck):
    """A on click, B after A (onEnd), C after B (onEnd); B is deleted.
    PowerPoint: C's slot starts at 500, when A ends."""
    pkg, part, a, b, c = _deck(make_deck, 7)
    an.add_entrance_animation(pkg, 0, a, "fade")
    an.add_entrance_animation(pkg, 0, b, "fade", trigger="after_previous")
    an.add_entrance_animation(pkg, 0, c, "appear", trigger="after_previous")
    e = _effects(pkg, part)
    _chain_on_end(e[a], e[b])
    _chain_on_end(e[b], e[c])
    _res, pkg2 = _delete_and_reload(pkg, part, b)
    assert _no_dangling(pkg2, part) == []
    assert _layout(pkg2, part) == [
        f"G[+indefinite] S[+0]{{{a}:clickEffect}} S[+500]{{{c}:afterEffect}}"
    ]


def test_deleting_the_last_effect_leaves_the_rest_byte_identical(make_deck):
    pkg, part, a, b, _c = _deck(make_deck, 6)
    an.add_entrance_animation(pkg, 0, a, "fade")
    an.add_entrance_animation(pkg, 0, b, "appear", trigger="after_previous")
    e = _effects(pkg, part)
    _chain_on_end(e[a], e[b])

    def slot_and_group_start(ctn):
        slot = ctn.getparent().getparent().getparent().getparent()
        group_ctn = slot.getparent().getparent()
        return etree.tostring(slot), etree.tostring(group_ctn.find(qn("p:stCondLst")))

    before = slot_and_group_start(e[a])
    res, pkg2 = _delete_and_reload(pkg, part, b)
    assert res["timing_report"]["retimed_shape_ids"] == []
    assert slot_and_group_start(_effects(pkg2, part)[a]) == before
    assert _layout(pkg2, part) == [f"G[+indefinite] S[+0]{{{a}:clickEffect}}"]


def _retype_with_previous(pkg, part, host_spid, moved_spid):
    """Move moved_spid's effect into host_spid's slot as a with-previous effect."""
    e = _effects(pkg, part)
    host = e[host_spid].getparent().getparent()
    par = e[moved_spid].getparent()
    old_slot = par.getparent().getparent().getparent()
    host.append(par)
    e[moved_spid].set("nodeType", "withEffect")
    old_slot.getparent().remove(old_slot)


def test_with_previous_left_leading_the_first_group_starts_with_the_slide(make_deck):
    pkg, part, a, b, _c = _deck(make_deck, 9)
    an.add_entrance_animation(pkg, 0, a, "fade")
    an.add_entrance_animation(pkg, 0, b, "fade", trigger="after_previous")
    _retype_with_previous(pkg, part, a, b)
    _res, pkg2 = _delete_and_reload(pkg, part, a)
    assert _layout(pkg2, part) == [
        f"G[+indefinite,onBegin@mainSeq+0] S[+0]{{{b}:withEffect}}"
    ]


def test_with_previous_left_alone_in_a_slot_joins_the_slot_before(make_deck):
    """A on click; B after A with C alongside it; B is deleted. PowerPoint:
    C now starts with A, in A's slot."""
    pkg, part, a, b, c = _deck(make_deck, 10)
    an.add_entrance_animation(pkg, 0, a, "fade")
    an.add_entrance_animation(pkg, 0, b, "fade", trigger="after_previous")
    an.add_entrance_animation(pkg, 0, c, "appear", trigger="after_previous")
    _retype_with_previous(pkg, part, b, c)
    res, pkg2 = _delete_and_reload(pkg, part, b)
    assert _layout(pkg2, part) == [
        f"G[+indefinite] S[+0]{{{a}:clickEffect {c}:withEffect}}"
    ]
    assert res["timing_report"]["retimed_shape_ids"] == [c]


def test_an_effect_delay_counts_when_slots_close_up(make_deck):
    """PowerPoint keeps an after-previous delay on the effect node: B has
    300 ms, so C (after B) starts at 300 + 500 = 800 once A is gone."""
    pkg, part, a, b, c = _deck(make_deck, 11)
    an.add_entrance_animation(pkg, 0, a, "fade")
    an.add_entrance_animation(pkg, 0, b, "fade", trigger="after_previous")
    an.add_entrance_animation(pkg, 0, c, "appear", trigger="after_previous")
    e = _effects(pkg, part)
    e[b].find(qn("p:stCondLst")).find(qn("p:cond")).set("delay", "300")
    e[c].getparent().getparent().getparent().find(qn("p:stCondLst")).find(
        qn("p:cond")
    ).set("delay", "1300")
    _res, pkg2 = _delete_and_reload(pkg, part, a)
    assert _layout(pkg2, part) == [
        f"G[+indefinite,onBegin@mainSeq+0] S[+0]{{{b}:afterEffect}} S[+800]{{{c}:afterEffect}}"
    ]


def test_a_gap_written_by_delay_ms_is_kept(make_deck):
    """add_entrance_animation writes delay_ms into the slot start (750 =
    A's 500 + 250). The gap stays when A is deleted (PowerPoint's own
    delete drops it: it reads that gap as no delay)."""
    pkg, part, a, b, _c = _deck(make_deck, 8)
    an.add_entrance_animation(pkg, 0, a, "fade")
    an.add_entrance_animation(pkg, 0, b, "appear", trigger="after_previous", delay_ms=250)
    _res, pkg2 = _delete_and_reload(pkg, part, a)
    assert _layout(pkg2, part) == [
        f"G[+indefinite,onBegin@mainSeq+0] S[+250]{{{b}:afterEffect}}"
    ]
