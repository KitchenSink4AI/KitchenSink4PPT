"""xsd:boolean attributes, read the way the schema spells them.

An OOXML boolean has four spellings: "1" and "true", "0" and "false", with
surrounding whitespace collapsed. PowerPoint writes "1" and "0", but other
producers write "true" and "false", and a reader that compared only against
"1" or "0" misread them: a slide hidden with show="false" read as shown, a
shape hidden with hidden="true" read as visible, a text box marked
txBox="true" read as an autoshape, and two operations disagreed about the
same shape. Every boolean read goes through here.
"""

from __future__ import annotations

_TRUE = frozenset(("1", "true"))
_FALSE = frozenset(("0", "false"))


def xml_bool(value: str | None, default: bool = False) -> bool:
    """The boolean an attribute value spells; `default` when the attribute is
    absent or is not one of the four schema spellings."""
    if value is None:
        return default
    text = value.strip()
    if text in _TRUE:
        return True
    if text in _FALSE:
        return False
    return default


def attr_bool(element, name: str, default: bool = False) -> bool:
    """xml_bool of element's attribute `name`; `default` when the element
    itself is missing (None)."""
    if element is None:
        return default
    return xml_bool(element.get(name), default)
