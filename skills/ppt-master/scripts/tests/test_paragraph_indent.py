"""Indented lines stay one text frame and become paragraph left margins.

A source that indents a bullet level by moving x must survive as <a:pPr marL>
on one shape, not as several shapes. Fixtures are built inside the test;
nothing here depends on a sample deck.
"""

from __future__ import annotations

import unittest
from xml.etree import ElementTree as ET

from svg_finalize.flatten_tspan import (
    PARAGRAPH_INDENT_ATTR,
    PARAGRAPH_SOFT_BREAK_ATTR,
    classify_paragraph_block,
    flatten_text_with_tspans,
)
from svg_to_pptx.drawingml.elements import _paragraph_pr_xml
from svg_to_pptx.pptx_package.builder import (
    _set_no_bullet_paragraph_properties,
)

DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

SVG_NS = "http://www.w3.org/2000/svg"
QN_TEXT = f"{{{SVG_NS}}}text"
QN_TSPAN = f"{{{SVG_NS}}}tspan"


def _text(base_x: float, lines: list[tuple[float, float, str]]) -> ET.Element:
    """One <text> whose children are (x, dy, content) visual lines."""
    text = ET.Element(QN_TEXT, {"x": str(base_x), "y": "100", "font-size": "24"})
    for x, dy, content in lines:
        tspan = ET.SubElement(text, QN_TSPAN, {"x": str(x), "dy": str(dy)})
        tspan.text = content
    return text


def _classify(text: ET.Element, preserve: bool = True):
    return classify_paragraph_block(text, preserve_line_breaks=preserve)


def _flatten(text: ET.Element) -> ET.Element:
    """Run the real in-memory flattener over a one-text document."""
    root = ET.Element(f"{{{SVG_NS}}}svg")
    root.append(text)
    flatten_text_with_tspans(
        ET.ElementTree(root),
        merge_paragraphs=True,
        preserve_line_breaks=True,
    )
    return root.find(QN_TEXT)


class ClassifyIndentedBlock(unittest.TestCase):
    """The classifier must accept a block whose lines sit at different x."""

    def test_indented_levels_are_mergeable(self):
        block = _text(30, [(30, 0, "Level 0"), (54, 32, "Level 1")])
        self.assertIsNotNone(_classify(block))

    def test_three_levels_are_mergeable(self):
        block = _text(
            30,
            [(30, 0, "a"), (54, 32, "b"), (78, 32, "c")],
        )
        self.assertIsNotNone(_classify(block))

    def test_uniform_x_still_mergeable(self):
        block = _text(30, [(30, 0, "a"), (30, 32, "b")])
        self.assertIsNotNone(_classify(block))

    def test_missing_frame_origin_with_positioned_line_is_rejected(self):
        block = _text(30, [(30, 0, "a"), (54, 32, "b")])
        del block.attrib["x"]
        self.assertIsNone(_classify(block))

    def test_non_numeric_frame_origin_is_rejected(self):
        block = _text(30, [(30, 0, "a"), (54, 32, "b")])
        block.set("x", "nonsense")
        self.assertIsNone(_classify(block))


class IndentForcesParagraphBreak(unittest.TestCase):
    """marL cannot change inside an <a:p>, so an x change must break one."""

    def test_indent_change_breaks_even_at_base_line_height(self):
        # Same dy on both lines would otherwise reflow into one paragraph.
        block = _text(30, [(30, 0, "a"), (54, 32, "b"), (54, 32, "c")])
        _base, _extras, breaks, _lines, _synthetic = _classify(block)
        self.assertEqual(breaks[1], "paragraph")
        # Line 2 repeats line 1's x, so it is free to stay a visual break.
        self.assertNotEqual(breaks[2], "paragraph")

    def test_indent_change_outranks_explicit_soft_break(self):
        block = _text(30, [(30, 0, "a"), (54, 32, "b")])
        list(block)[1].set(PARAGRAPH_SOFT_BREAK_ATTR, "1")
        _base, _extras, breaks, _lines, _synthetic = _classify(block)
        self.assertEqual(breaks[1], "paragraph")

    def test_uniform_x_honours_explicit_soft_break(self):
        block = _text(30, [(30, 0, "a"), (30, 32, "b")])
        list(block)[1].set(PARAGRAPH_SOFT_BREAK_ATTR, "1")
        _base, _extras, breaks, _lines, _synthetic = _classify(block)
        self.assertEqual(breaks[1], "soft")


class EmittedIndentAttribute(unittest.TestCase):
    """The flattener records each line's offset from its frame origin."""

    def test_offsets_are_recorded_per_line(self):
        text = _flatten(
            _text(30, [(30, 0, "a"), (54, 32, "b"), (78, 32, "c")])
        )
        got = [child.get(PARAGRAPH_INDENT_ATTR) for child in text]
        self.assertIsNone(got[0])
        self.assertAlmostEqual(float(got[1]), 24.0)
        self.assertAlmostEqual(float(got[2]), 48.0)

    def test_uniform_block_records_no_indent(self):
        text = _flatten(_text(30, [(30, 0, "a"), (30, 32, "b")]))
        for child in text:
            self.assertIsNone(child.get(PARAGRAPH_INDENT_ATTR))

    def test_negative_offset_clamps_to_none(self):
        # ST_TextMargin is non-negative; a line left of its frame is not indented.
        text = _flatten(_text(30, [(30, 0, "a"), (10, 32, "b")]))
        for child in text:
            self.assertIsNone(child.get(PARAGRAPH_INDENT_ATTR))

    def test_indent_survives_a_multi_run_line(self):
        # A line with inline runs is wrapped in a container that has no x, so
        # the offset has to be read before normalization.
        block = _text(30, [(30, 0, "a"), (54, 32, "bullet ")])
        inline = ET.SubElement(list(block)[1], QN_TSPAN, {"fill": "#FF0000"})
        inline.text = "text"
        text = _flatten(block)
        self.assertAlmostEqual(
            float(list(text)[1].get(PARAGRAPH_INDENT_ATTR)), 24.0
        )


class ParagraphPropertiesXml(unittest.TestCase):
    """marL carries the indent; a bullet adds its hanging offset on top."""

    BULLET = {"char": "•", "margin_px": 12.0}

    def _attr(self, xml: str, name: str) -> int | None:
        # The builder emits a bare <a:pPr> fragment; bind the prefix to parse it.
        wrapped = (
            '<root xmlns:a="http://schemas.openxmlformats.org/'
            f'drawingml/2006/main">{xml}</root>'
        )
        element = list(ET.fromstring(wrapped))[0]
        raw = element.get(name)
        return None if raw is None else int(raw)

    def test_indent_without_bullet_sets_only_mar_l(self):
        xml = _paragraph_pr_xml(algn="l", font_size=16.0, left_indent_px=24.0)
        self.assertEqual(self._attr(xml, "marL"), 24 * 9525)
        self.assertIsNone(self._attr(xml, "indent"))

    def test_bullet_indent_adds_to_the_hanging_margin(self):
        xml = _paragraph_pr_xml(
            algn="l", font_size=16.0, bullet=self.BULLET, left_indent_px=24.0
        )
        self.assertEqual(self._attr(xml, "marL"), (24 + 12) * 9525)
        self.assertEqual(self._attr(xml, "indent"), -12 * 9525)

    def test_bullet_without_indent_is_unchanged(self):
        xml = _paragraph_pr_xml(algn="l", font_size=16.0, bullet=self.BULLET)
        self.assertEqual(self._attr(xml, "marL"), 12 * 9525)
        self.assertEqual(self._attr(xml, "indent"), -12 * 9525)

    def test_no_bullet_no_indent_emits_neither(self):
        xml = _paragraph_pr_xml(algn="l", font_size=16.0)
        self.assertIsNone(self._attr(xml, "marL"))
        self.assertIsNone(self._attr(xml, "indent"))

    def test_negative_indent_never_reaches_mar_l(self):
        xml = _paragraph_pr_xml(algn="l", font_size=16.0, left_indent_px=-30.0)
        self.assertIsNone(self._attr(xml, "marL"))

    def test_deeper_levels_order_strictly_increasing(self):
        margins = [
            self._attr(
                _paragraph_pr_xml(
                    algn="l",
                    font_size=16.0,
                    bullet=self.BULLET,
                    left_indent_px=depth,
                ),
                "marL",
            )
            for depth in (0.0, 24.0, 48.0)
        ]
        self.assertEqual(margins, sorted(margins))
        self.assertEqual(len(set(margins)), len(margins))


class BulletSuppressionKeepsIndentLevel(unittest.TestCase):
    """A body placeholder drops inherited bullets but keeps its indent depth."""

    def _p_pr(self, **attrs: str) -> ET.Element:
        return ET.Element(f"{{{DML_NS}}}pPr", attrs)

    def test_hanging_indent_is_removed_and_level_kept(self):
        # marL = level(24px) + bullet margin(12px), indent = -12px.
        props = self._p_pr(marL=str(36 * 9525), indent=str(-12 * 9525))
        _set_no_bullet_paragraph_properties(props, replace_existing=True)
        self.assertEqual(props.get("marL"), str(24 * 9525))
        self.assertEqual(props.get("indent"), "0")

    def test_plain_indent_survives_untouched(self):
        props = self._p_pr(marL=str(24 * 9525))
        _set_no_bullet_paragraph_properties(props, replace_existing=True)
        self.assertEqual(props.get("marL"), str(24 * 9525))

    def test_bullet_only_paragraph_still_collapses_to_zero(self):
        # The pre-existing case: no indent level, only the bullet's own margin.
        props = self._p_pr(marL=str(12 * 9525), indent=str(-12 * 9525))
        _set_no_bullet_paragraph_properties(props, replace_existing=True)
        self.assertEqual(props.get("marL"), "0")
        self.assertEqual(props.get("indent"), "0")

    def test_bare_paragraph_still_collapses_to_zero(self):
        props = self._p_pr()
        _set_no_bullet_paragraph_properties(props, replace_existing=True)
        self.assertEqual(props.get("marL"), "0")
        self.assertEqual(props.get("indent"), "0")

    def test_bullet_is_suppressed(self):
        props = self._p_pr(marL=str(36 * 9525), indent=str(-12 * 9525))
        ET.SubElement(props, f"{{{DML_NS}}}buChar", {"char": "•"})
        _set_no_bullet_paragraph_properties(props, replace_existing=True)
        self.assertIsNone(props.find(f"{{{DML_NS}}}buChar"))
        self.assertIsNotNone(props.find(f"{{{DML_NS}}}buNone"))

    def test_non_numeric_values_fall_back_to_zero(self):
        props = self._p_pr(marL="wide", indent="-12")
        _set_no_bullet_paragraph_properties(props, replace_existing=True)
        self.assertEqual(props.get("marL"), "0")


if __name__ == "__main__":
    unittest.main()
