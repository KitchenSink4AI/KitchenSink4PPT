"""The adjust handles of every DrawingML preset shape.

Each preset geometry (a:prstGeom prst="...") draws from a short list of
named guides, its adjust handles: the yellow handles PowerPoint shows on an
arc, a rounded rectangle or an arrow. A shape stores the ones that differ
from the preset's default as a:avLst/a:gd elements, fmla "val N".

GUIDES maps all 187 preset names of ECMA-376 to their handles, in the
preset's own order, with the default each one takes when the shape stores
nothing. The names and defaults match the presetShapeDefinitions of the
standard. A preset with an empty tuple has no handles.

Most handles count in 100000ths: of the shape's shorter side, of its width
or height, or of a ratio, depending on the preset. The handles in ANGLES
are angles instead, in 60000ths of a degree, measured clockwise from the
3 o'clock direction (0 points right, 90 degrees points down, 180 left and
270 up). An arc is drawn clockwise from its adj1 angle to its adj2 angle,
so the default arc (270 to 0) is the top right quarter of the ellipse.

What PowerPoint opens, measured on PowerPoint 16 with an invisible
PowerPoint and a full load of the deck: a shape's a:avLst is either empty or
the preset's WHOLE list of handles in the preset's own order. A list that
names only some of the handles (even the first one alone), lists them out
of order, names a guide the preset does not have, names one twice, or holds
a value past a signed 64-bit integer makes PowerPoint refuse the whole deck.
A whole list opens whatever its values are: an angle past a full turn, a
negative angle and a value past the preset's own limits all open, and
PowerPoint pins each value when it draws. A preset without handles ignores
its list. list_problem says which rule a stored list breaks.
"""

from __future__ import annotations

#: preset name -> ((handle name, default value), ...) in the preset's order.
GUIDES: dict[str, tuple[tuple[str, int], ...]] = {
    "accentBorderCallout1": (('adj1', 18750), ('adj2', -8333), ('adj3', 112500), ('adj4', -38333)),
    "accentBorderCallout2": (('adj1', 18750), ('adj2', -8333), ('adj3', 18750), ('adj4', -16667), ('adj5', 112500), ('adj6', -46667)),
    "accentBorderCallout3": (('adj1', 18750), ('adj2', -8333), ('adj3', 18750), ('adj4', -16667), ('adj5', 100000), ('adj6', -16667), ('adj7', 112963), ('adj8', -8333)),
    "accentCallout1": (('adj1', 18750), ('adj2', -8333), ('adj3', 112500), ('adj4', -38333)),
    "accentCallout2": (('adj1', 18750), ('adj2', -8333), ('adj3', 18750), ('adj4', -16667), ('adj5', 112500), ('adj6', -46667)),
    "accentCallout3": (('adj1', 18750), ('adj2', -8333), ('adj3', 18750), ('adj4', -16667), ('adj5', 100000), ('adj6', -16667), ('adj7', 112963), ('adj8', -8333)),
    "actionButtonBackPrevious": (),
    "actionButtonBeginning": (),
    "actionButtonBlank": (),
    "actionButtonDocument": (),
    "actionButtonEnd": (),
    "actionButtonForwardNext": (),
    "actionButtonHelp": (),
    "actionButtonHome": (),
    "actionButtonInformation": (),
    "actionButtonMovie": (),
    "actionButtonReturn": (),
    "actionButtonSound": (),
    "arc": (('adj1', 16200000), ('adj2', 0)),
    "bentArrow": (('adj1', 25000), ('adj2', 25000), ('adj3', 25000), ('adj4', 43750)),
    "bentConnector2": (),
    "bentConnector3": (('adj1', 50000),),
    "bentConnector4": (('adj1', 50000), ('adj2', 50000)),
    "bentConnector5": (('adj1', 50000), ('adj2', 50000), ('adj3', 50000)),
    "bentUpArrow": (('adj1', 25000), ('adj2', 25000), ('adj3', 25000)),
    "bevel": (('adj', 12500),),
    "blockArc": (('adj1', 10800000), ('adj2', 0), ('adj3', 25000)),
    "borderCallout1": (('adj1', 18750), ('adj2', -8333), ('adj3', 112500), ('adj4', -38333)),
    "borderCallout2": (('adj1', 18750), ('adj2', -8333), ('adj3', 18750), ('adj4', -16667), ('adj5', 112500), ('adj6', -46667)),
    "borderCallout3": (('adj1', 18750), ('adj2', -8333), ('adj3', 18750), ('adj4', -16667), ('adj5', 100000), ('adj6', -16667), ('adj7', 112963), ('adj8', -8333)),
    "bracePair": (('adj', 8333),),
    "bracketPair": (('adj', 16667),),
    "callout1": (('adj1', 18750), ('adj2', -8333), ('adj3', 112500), ('adj4', -38333)),
    "callout2": (('adj1', 18750), ('adj2', -8333), ('adj3', 18750), ('adj4', -16667), ('adj5', 112500), ('adj6', -46667)),
    "callout3": (('adj1', 18750), ('adj2', -8333), ('adj3', 18750), ('adj4', -16667), ('adj5', 100000), ('adj6', -16667), ('adj7', 112963), ('adj8', -8333)),
    "can": (('adj', 25000),),
    "chartPlus": (),
    "chartStar": (),
    "chartX": (),
    "chevron": (('adj', 50000),),
    "chord": (('adj1', 2700000), ('adj2', 16200000)),
    "circularArrow": (('adj1', 12500), ('adj2', 1142319), ('adj3', 20457681), ('adj4', 10800000), ('adj5', 12500)),
    "cloud": (),
    "cloudCallout": (('adj1', -20833), ('adj2', 62500)),
    "corner": (('adj1', 50000), ('adj2', 50000)),
    "cornerTabs": (),
    "cube": (('adj', 25000),),
    "curvedConnector2": (),
    "curvedConnector3": (('adj1', 50000),),
    "curvedConnector4": (('adj1', 50000), ('adj2', 50000)),
    "curvedConnector5": (('adj1', 50000), ('adj2', 50000), ('adj3', 50000)),
    "curvedDownArrow": (('adj1', 25000), ('adj2', 50000), ('adj3', 25000)),
    "curvedLeftArrow": (('adj1', 25000), ('adj2', 50000), ('adj3', 25000)),
    "curvedRightArrow": (('adj1', 25000), ('adj2', 50000), ('adj3', 25000)),
    "curvedUpArrow": (('adj1', 25000), ('adj2', 50000), ('adj3', 25000)),
    "decagon": (('vf', 105146),),
    "diagStripe": (('adj', 50000),),
    "diamond": (),
    "dodecagon": (),
    "donut": (('adj', 25000),),
    "doubleWave": (('adj1', 6250), ('adj2', 0)),
    "downArrow": (('adj1', 50000), ('adj2', 50000)),
    "downArrowCallout": (('adj1', 25000), ('adj2', 25000), ('adj3', 25000), ('adj4', 64977)),
    "ellipse": (),
    "ellipseRibbon": (('adj1', 25000), ('adj2', 50000), ('adj3', 12500)),
    "ellipseRibbon2": (('adj1', 25000), ('adj2', 50000), ('adj3', 12500)),
    "flowChartAlternateProcess": (),
    "flowChartCollate": (),
    "flowChartConnector": (),
    "flowChartDecision": (),
    "flowChartDelay": (),
    "flowChartDisplay": (),
    "flowChartDocument": (),
    "flowChartExtract": (),
    "flowChartInputOutput": (),
    "flowChartInternalStorage": (),
    "flowChartMagneticDisk": (),
    "flowChartMagneticDrum": (),
    "flowChartMagneticTape": (),
    "flowChartManualInput": (),
    "flowChartManualOperation": (),
    "flowChartMerge": (),
    "flowChartMultidocument": (),
    "flowChartOfflineStorage": (),
    "flowChartOffpageConnector": (),
    "flowChartOnlineStorage": (),
    "flowChartOr": (),
    "flowChartPredefinedProcess": (),
    "flowChartPreparation": (),
    "flowChartProcess": (),
    "flowChartPunchedCard": (),
    "flowChartPunchedTape": (),
    "flowChartSort": (),
    "flowChartSummingJunction": (),
    "flowChartTerminator": (),
    "foldedCorner": (('adj', 16667),),
    "frame": (('adj1', 12500),),
    "funnel": (),
    "gear6": (('adj1', 15000), ('adj2', 3526)),
    "gear9": (('adj1', 10000), ('adj2', 1763)),
    "halfFrame": (('adj1', 33333), ('adj2', 33333)),
    "heart": (),
    "heptagon": (('hf', 102572), ('vf', 105210)),
    "hexagon": (('adj', 25000), ('vf', 115470)),
    "homePlate": (('adj', 50000),),
    "horizontalScroll": (('adj', 12500),),
    "irregularSeal1": (),
    "irregularSeal2": (),
    "leftArrow": (('adj1', 50000), ('adj2', 50000)),
    "leftArrowCallout": (('adj1', 25000), ('adj2', 25000), ('adj3', 25000), ('adj4', 64977)),
    "leftBrace": (('adj1', 8333), ('adj2', 50000)),
    "leftBracket": (('adj', 8333),),
    "leftCircularArrow": (('adj1', 12500), ('adj2', -1142319), ('adj3', 1142319), ('adj4', 10800000), ('adj5', 12500)),
    "leftRightArrow": (('adj1', 50000), ('adj2', 50000)),
    "leftRightArrowCallout": (('adj1', 25000), ('adj2', 25000), ('adj3', 25000), ('adj4', 48123)),
    "leftRightCircularArrow": (('adj1', 12500), ('adj2', 1142319), ('adj3', 20457681), ('adj4', 11942319), ('adj5', 12500)),
    "leftRightRibbon": (('adj1', 50000), ('adj2', 50000), ('adj3', 16667)),
    "leftRightUpArrow": (('adj1', 25000), ('adj2', 25000), ('adj3', 25000)),
    "leftUpArrow": (('adj1', 25000), ('adj2', 25000), ('adj3', 25000)),
    "lightningBolt": (),
    "line": (),
    "lineInv": (),
    "mathDivide": (('adj1', 23520), ('adj2', 5880), ('adj3', 11760)),
    "mathEqual": (('adj1', 23520), ('adj2', 11760)),
    "mathMinus": (('adj1', 23520),),
    "mathMultiply": (('adj1', 23520),),
    "mathNotEqual": (('adj1', 23520), ('adj2', 6600000), ('adj3', 11760)),
    "mathPlus": (('adj1', 23520),),
    "moon": (('adj', 50000),),
    "noSmoking": (('adj', 18750),),
    "nonIsoscelesTrapezoid": (('adj1', 25000), ('adj2', 25000)),
    "notchedRightArrow": (('adj1', 50000), ('adj2', 50000)),
    "octagon": (('adj', 29289),),
    "parallelogram": (('adj', 25000),),
    "pentagon": (('hf', 105146), ('vf', 110557)),
    "pie": (('adj1', 0), ('adj2', 16200000)),
    "pieWedge": (),
    "plaque": (('adj', 16667),),
    "plaqueTabs": (),
    "plus": (('adj', 25000),),
    "quadArrow": (('adj1', 22500), ('adj2', 22500), ('adj3', 22500)),
    "quadArrowCallout": (('adj1', 18515), ('adj2', 18515), ('adj3', 18515), ('adj4', 48123)),
    "rect": (),
    "ribbon": (('adj1', 16667), ('adj2', 50000)),
    "ribbon2": (('adj1', 16667), ('adj2', 50000)),
    "rightArrow": (('adj1', 50000), ('adj2', 50000)),
    "rightArrowCallout": (('adj1', 25000), ('adj2', 25000), ('adj3', 25000), ('adj4', 64977)),
    "rightBrace": (('adj1', 8333), ('adj2', 50000)),
    "rightBracket": (('adj', 8333),),
    "round1Rect": (('adj', 16667),),
    "round2DiagRect": (('adj1', 16667), ('adj2', 0)),
    "round2SameRect": (('adj1', 16667), ('adj2', 0)),
    "roundRect": (('adj', 16667),),
    "rtTriangle": (),
    "smileyFace": (('adj', 4653),),
    "snip1Rect": (('adj', 16667),),
    "snip2DiagRect": (('adj1', 0), ('adj2', 16667)),
    "snip2SameRect": (('adj1', 16667), ('adj2', 0)),
    "snipRoundRect": (('adj1', 16667), ('adj2', 16667)),
    "squareTabs": (),
    "star10": (('adj', 42533), ('hf', 105146)),
    "star12": (('adj', 37500),),
    "star16": (('adj', 37500),),
    "star24": (('adj', 37500),),
    "star32": (('adj', 37500),),
    "star4": (('adj', 12500),),
    "star5": (('adj', 19098), ('hf', 105146), ('vf', 110557)),
    "star6": (('adj', 28868), ('hf', 115470)),
    "star7": (('adj', 34601), ('hf', 102572), ('vf', 105210)),
    "star8": (('adj', 37500),),
    "straightConnector1": (),
    "stripedRightArrow": (('adj1', 50000), ('adj2', 50000)),
    "sun": (('adj', 25000),),
    "swooshArrow": (('adj1', 25000), ('adj2', 16667)),
    "teardrop": (('adj', 100000),),
    "trapezoid": (('adj', 25000),),
    "triangle": (('adj', 50000),),
    "upArrow": (('adj1', 50000), ('adj2', 50000)),
    "upArrowCallout": (('adj1', 25000), ('adj2', 25000), ('adj3', 25000), ('adj4', 64977)),
    "upDownArrow": (('adj1', 50000), ('adj2', 50000)),
    "upDownArrowCallout": (('adj1', 25000), ('adj2', 25000), ('adj3', 25000), ('adj4', 48123)),
    "uturnArrow": (('adj1', 25000), ('adj2', 25000), ('adj3', 25000), ('adj4', 43750), ('adj5', 75000)),
    "verticalScroll": (('adj', 12500),),
    "wave": (('adj1', 12500), ('adj2', 0)),
    "wedgeEllipseCallout": (('adj1', -20833), ('adj2', 62500)),
    "wedgeRectCallout": (('adj1', -20833), ('adj2', 62500)),
    "wedgeRoundRectCallout": (('adj1', -20833), ('adj2', 62500), ('adj3', 16667)),
}

#: preset name -> the handles that are angles (60000ths of a degree,
#: clockwise from 3 o'clock).
ANGLES: dict[str, frozenset[str]] = {
    "arc": frozenset({"adj1", "adj2"}),
    "blockArc": frozenset({"adj1", "adj2"}),
    "chord": frozenset({"adj1", "adj2"}),
    "pie": frozenset({"adj1", "adj2"}),
    "circularArrow": frozenset({"adj2", "adj3", "adj4"}),
    "leftCircularArrow": frozenset({"adj2", "adj3", "adj4"}),
    "leftRightCircularArrow": frozenset({"adj2", "adj3", "adj4"}),
    "mathNotEqual": frozenset({"adj2"}),
}

#: 60000ths of a degree per degree, and in one full turn.
PER_DEGREE = 60000
FULL_TURN = 360 * PER_DEGREE

#: The values this server writes: a signed 32-bit whole number, far past
#: any value a preset uses (PowerPoint reads up to a signed 64-bit one).
VALUE_MIN, VALUE_MAX = -(2 ** 31), 2 ** 31 - 1
#: The values PowerPoint reads in a stored list: a signed 64-bit integer.
STORED_MIN, STORED_MAX = -(2 ** 63), 2 ** 63 - 1


def handles(preset: str) -> tuple[tuple[str, int], ...] | None:
    """The preset's handles with their defaults, or None for a name that is
    not a DrawingML preset."""
    return GUIDES.get(preset)


def is_angle(preset: str, name: str) -> bool:
    return name in ANGLES.get(preset, frozenset())


def degrees_of(value: int) -> float:
    """A stored angle as degrees, rounded to a thousandth."""
    return round(value / PER_DEGREE, 3)


def stored_value(fmla: str) -> int | None:
    """The whole number in a stored "val N" formula, or None for any other
    formula."""
    parts = (fmla or "").split()
    if len(parts) == 2 and parts[0] == "val":
        try:
            return int(parts[1])
        except ValueError:
            return None
    return None


def list_problem(preset: str, stored: list[tuple[str, str]]) -> str | None:
    """Why PowerPoint would refuse to open a deck holding this stored adjust
    list (name, formula pairs in document order), in words, or None when it
    would open it. An empty list, a preset without handles and a name that
    is not a DrawingML preset are never a problem here."""
    known = GUIDES.get(preset)
    if not known or not stored:
        return None
    names = [name for name, _default in known]
    got = [name for name, _fmla in stored]
    have = ", ".join(names)
    unknown = [name for name in got if name not in names]
    if unknown:
        return (f"it names {', '.join(repr(n) for n in dict.fromkeys(unknown))}, which "
                f"the {preset} preset does not have (its handles are {have})")
    twice = [name for name in dict.fromkeys(got) if got.count(name) > 1]
    if twice:
        return f"it names {', '.join(twice)} more than once"
    if got != names:
        missing = [name for name in names if name not in got]
        if missing:
            return (f"it stores only {', '.join(got)} of the {preset} preset's "
                    f"handles ({have})")
        return f"it lists the {preset} preset's handles out of order ({', '.join(got)})"
    for name, fmla in stored:
        value = stored_value(fmla)
        if value is not None and not STORED_MIN <= value <= STORED_MAX:
            return f"{name} holds {value}, more than PowerPoint can read"
    return None


__all__ = ["ANGLES", "FULL_TURN", "GUIDES", "PER_DEGREE", "STORED_MAX", "STORED_MIN",
           "VALUE_MAX", "VALUE_MIN", "degrees_of", "handles", "is_angle", "list_problem",
           "stored_value"]
