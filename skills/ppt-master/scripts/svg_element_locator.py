#!/usr/bin/env python3
"""
PPT Master - SVG Element Locator

Names an element a validation message reports, so a repair can patch that
element in place: its id when it has one, else its path and visible text.

Usage:
    Imported by the SVG quality checker and the semantic marker validator.

Examples:
    element_locator(root, elem)  # '<text> at /svg/g[2]/text[1] (text=\'Why…\')'

Dependencies:
    None (only uses standard library)
"""

from __future__ import annotations

from xml.etree import ElementTree as ET

# Matches the checker's own text excerpts in overflow messages.
_SNIPPET_CHARS = 20


def _local_name(elem: ET.Element) -> str:
    tag = elem.tag
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def element_path(root: ET.Element, elem: ET.Element) -> str:
    """XPath-style position, 1-based among same-tag siblings: /svg/g[2]/text[1]."""
    parents = {child: parent for parent in root.iter() for child in parent}
    steps: list[str] = []
    current = elem
    while current is not root and current in parents:
        parent = parents[current]
        tag = _local_name(current)
        siblings = [child for child in parent if _local_name(child) == tag]
        steps.append(f"{tag}[{siblings.index(current) + 1}]")
        current = parent
    steps.append(_local_name(root))
    return "/" + "/".join(reversed(steps))


def text_snippet(elem: ET.Element) -> str:
    """The element's visible text as ` (text='…')`, or empty when it has none."""
    text = " ".join("".join(elem.itertext()).split())
    if not text:
        return ""
    if len(text) > _SNIPPET_CHARS:
        text = f"{text[:_SNIPPET_CHARS]}…"
    return f" (text={text!r})"


def element_locator(root: ET.Element, elem: ET.Element) -> str:
    """`<tag id="x">` when the element has an id, else `<tag> at <path>` plus its text."""
    tag = _local_name(elem)
    elem_id = (elem.get("id") or "").strip()
    if elem_id:
        return f'<{tag} id="{elem_id}">'
    return f"<{tag}> at {element_path(root, elem)}{text_snippet(elem)}"
