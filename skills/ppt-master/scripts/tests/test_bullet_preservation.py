"""A source template's bullet survives the round trip as a declaration.

Import used to flatten <a:buChar>/<a:buAutoNum> into a literal text prefix, so
export had to guess a bullet back from the glyph and only recognised nine
characters. The declaration now travels with the run that draws it. Fixtures are
built inside the test; nothing here depends on a sample deck.
"""

from __future__ import annotations

import json
import unittest
from xml.etree import ElementTree as ET

from pptx_to_svg.txbody_to_svg import TextRun, _bullet_spec_attr, _resolve_bullet
from svg_to_pptx.drawingml.elements import (
    _BULLET_SPEC_ATTR,
    _BULLET_SPEC_KEY,
    _build_bullet_xml,
    _declared_bullet_spec,
    _extract_text_bullet,
)
from svg_to_pptx.pptx_package.builder import (
    _set_no_bullet_paragraph_properties,
    _set_placeholder_no_inherited_bullets,
)
from svg_to_pptx.pptx_package.template_structure import TemplateElementSpec

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _p_pr(xml: str) -> ET.Element:
    return ET.fromstring(f'<a:pPr xmlns:a="{A_NS}">{xml}</a:pPr>')


def _parse_bullet_xml(xml: str) -> ET.Element:
    return ET.fromstring(f'<root xmlns:a="{A_NS}">{xml}</root>')


class ResolveBulletDeclaration(unittest.TestCase):
    """Import keeps what the template declared beside what it draws."""

    def test_bu_char_keeps_glyph_and_typeface(self):
        prefix, spec = _resolve_bullet(
            (_p_pr('<a:buFont typeface="Arial"/><a:buChar char="•"/>'),), 0, {}
        )
        self.assertEqual(prefix, "• ")
        self.assertEqual(spec, {"kind": "character", "char": "•", "font": "Arial"})

    def test_wingdings_disc_substitutes_only_the_drawn_glyph(self):
        # The SVG must draw a disc, but the declaration keeps the real pairing
        # so PowerPoint renders it from the font rather than from a lookalike.
        prefix, spec = _resolve_bullet(
            (_p_pr('<a:buFont typeface="Wingdings"/><a:buChar char="l"/>'),), 0, {}
        )
        self.assertEqual(prefix, "● ")
        self.assertEqual(spec["char"], "l")
        self.assertEqual(spec["font"], "Wingdings")

    def test_unrecognised_glyph_is_still_declared(self):
        # "§" is not in the export-side glyph table; only a declaration saves it.
        _prefix, spec = _resolve_bullet(
            (_p_pr('<a:buFont typeface="Wingdings"/><a:buChar char="§"/>'),), 0, {}
        )
        self.assertEqual(spec["kind"], "character")
        self.assertEqual(spec["char"], "§")

    def test_auto_number_keeps_its_scheme(self):
        prefix, spec = _resolve_bullet(
            (_p_pr('<a:buAutoNum type="arabicPeriod"/>'),), 0, {}
        )
        self.assertEqual(prefix, "1. ")
        self.assertEqual(spec["kind"], "auto_number")
        self.assertEqual(spec["scheme"], "arabicPeriod")

    def test_auto_number_start_at_is_kept(self):
        _prefix, spec = _resolve_bullet(
            (_p_pr('<a:buAutoNum type="alphaLcParenR" startAt="3"/>'),), 0, {}
        )
        self.assertEqual(spec["start_at"], 3)
        self.assertEqual(spec["scheme"], "alphaLcParenR")

    def test_bu_none_is_declared_as_none(self):
        prefix, spec = _resolve_bullet((_p_pr("<a:buNone/>"),), 0, {})
        self.assertEqual(prefix, "")
        self.assertEqual(spec, {"kind": "none"})

    def test_silence_declares_nothing(self):
        prefix, spec = _resolve_bullet((_p_pr(""),), 0, {})
        self.assertEqual(prefix, "")
        self.assertIsNone(spec)


class SerializedOntoTheRun(unittest.TestCase):
    """The declaration reaches the SVG on the run that renders it."""

    def test_attribute_round_trips_through_json(self):
        run = TextRun(
            text="• ", font_size_px=16.0, font_family="Arial", fill="#000000",
            bullet_spec={"kind": "character", "char": "•", "font": "Arial"},
        )
        attr = _bullet_spec_attr(run)
        self.assertIn(_BULLET_SPEC_ATTR, attr)
        element = ET.fromstring(f"<tspan{attr}/>")
        self.assertEqual(
            _declared_bullet_spec(element),
            {"kind": "character", "char": "•", "font": "Arial"},
        )

    def test_a_run_without_a_bullet_emits_nothing(self):
        run = TextRun(
            text="text", font_size_px=16.0, font_family="Arial", fill="#000000"
        )
        self.assertEqual(_bullet_spec_attr(run), "")

    def test_quotes_in_a_glyph_do_not_break_the_attribute(self):
        run = TextRun(
            text='" ', font_size_px=16.0, font_family="Arial", fill="#000000",
            bullet_spec={"kind": "character", "char": '"'},
        )
        element = ET.fromstring(f"<tspan{_bullet_spec_attr(run)}/>")
        self.assertEqual(_declared_bullet_spec(element)["char"], '"')

    def test_malformed_payload_is_ignored(self):
        element = ET.fromstring(f'<tspan {_BULLET_SPEC_ATTR}="not json"/>')
        self.assertIsNone(_declared_bullet_spec(element))

    def test_payload_without_a_kind_is_ignored(self):
        element = ET.Element("tspan", {_BULLET_SPEC_ATTR: json.dumps({})})
        self.assertIsNone(_declared_bullet_spec(element))


class ExportPrefersTheDeclaration(unittest.TestCase):
    """A declared bullet outranks the nine-glyph heuristic."""

    def _runs(self, spec, marker="§ ", body="Level 2"):
        return [
            {"text": marker, "fill": "#000000", _BULLET_SPEC_KEY: spec},
            {"text": body, "fill": "#000000"},
        ]

    def test_unrecognised_glyph_still_becomes_a_bullet(self):
        runs, bullet = _extract_text_bullet(
            self._runs({"kind": "character", "char": "§", "font": "Wingdings"})
        )
        self.assertEqual(bullet["char"], "§")
        self.assertEqual(bullet["font"], "Wingdings")
        self.assertEqual("".join(r["text"] for r in runs), "Level 2")

    def test_the_drawn_prefix_is_dropped(self):
        runs, _bullet = _extract_text_bullet(
            self._runs({"kind": "character", "char": "•"}, marker="• ")
        )
        self.assertNotIn("•", "".join(r["text"] for r in runs))

    def test_auto_number_prefix_is_dropped_and_scheme_kept(self):
        runs, bullet = _extract_text_bullet(
            self._runs({"kind": "auto_number", "scheme": "arabicPeriod"}, marker="1. ")
        )
        self.assertEqual(bullet["scheme"], "arabicPeriod")
        self.assertNotIn("1.", "".join(r["text"] for r in runs))

    def test_declared_none_removes_the_prefix_without_a_bullet(self):
        runs, bullet = _extract_text_bullet(
            self._runs({"kind": "none"}, marker="")
        )
        self.assertIsNone(bullet)
        self.assertEqual("".join(r["text"] for r in runs), "Level 2")

    def test_undeclared_runs_still_use_the_glyph_heuristic(self):
        runs, bullet = _extract_text_bullet(
            [{"text": "• ", "fill": "#000"}, {"text": "Point", "fill": "#000"}]
        )
        self.assertIsNotNone(bullet)
        self.assertEqual(bullet["char"], "•")
        self.assertNotIn("declared", bullet)


class BulletXmlRestatesTheSource(unittest.TestCase):
    """The emitted <a:pPr> children match what the template declared."""

    def test_typeface_is_restated(self):
        root = _parse_bullet_xml(
            _build_bullet_xml({"char": "§", "font": "Wingdings"}, None)
        )
        self.assertEqual(root.find(f"{{{A_NS}}}buFont").get("typeface"), "Wingdings")
        self.assertEqual(root.find(f"{{{A_NS}}}buChar").get("char"), "§")

    def test_missing_typeface_falls_back_to_inherited(self):
        root = _parse_bullet_xml(_build_bullet_xml({"char": "•"}, None))
        self.assertIsNotNone(root.find(f"{{{A_NS}}}buFontTx"))
        self.assertIsNone(root.find(f"{{{A_NS}}}buFont"))

    def test_auto_number_emits_bu_auto_num_not_bu_char(self):
        root = _parse_bullet_xml(
            _build_bullet_xml({"scheme": "romanUcPeriod", "start_at": 4}, None)
        )
        node = root.find(f"{{{A_NS}}}buAutoNum")
        self.assertEqual(node.get("type"), "romanUcPeriod")
        self.assertEqual(node.get("startAt"), "4")
        self.assertIsNone(root.find(f"{{{A_NS}}}buChar"))

    def test_auto_number_without_start_at_omits_the_attribute(self):
        root = _parse_bullet_xml(_build_bullet_xml({"scheme": "arabicPeriod"}, None))
        self.assertIsNone(root.find(f"{{{A_NS}}}buAutoNum").get("startAt"))

    def test_no_bullet_emits_nothing(self):
        self.assertEqual(_build_bullet_xml(None, None), "")


class SuppressionKeepsDeclaredBullets(unittest.TestCase):
    """A prose placeholder loses inherited bullets, never its declared one."""

    def test_declared_bullet_survives(self):
        props = _p_pr('<a:buFont typeface="Arial"/><a:buChar char="•"/>')
        props.set("marL", "228600")
        _set_no_bullet_paragraph_properties(props)
        self.assertIsNotNone(props.find(f"{{{A_NS}}}buChar"))
        self.assertIsNone(props.find(f"{{{A_NS}}}buNone"))
        self.assertEqual(props.get("marL"), "228600")

    def test_auto_number_survives(self):
        props = _p_pr('<a:buAutoNum type="arabicPeriod"/>')
        _set_no_bullet_paragraph_properties(props)
        self.assertIsNotNone(props.find(f"{{{A_NS}}}buAutoNum"))

    def test_a_paragraph_with_no_bullet_is_still_suppressed(self):
        props = _p_pr("")
        _set_no_bullet_paragraph_properties(props)
        self.assertIsNotNone(props.find(f"{{{A_NS}}}buNone"))
        self.assertEqual(props.get("marL"), "0")


class PlaceholderPolicyKeepsDeclaredBullets(unittest.TestCase):
    """The body-placeholder policy itself, not just the helper beneath it."""

    def _shape(self, p_pr_xml: str) -> ET.Element:
        return ET.fromstring(
            f'<p:sp xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:txBody>'
            f"<a:p><a:pPr>{p_pr_xml}</a:pPr></a:p>"
            "</p:txBody></p:sp>"
        )

    def _spec(self, placeholder: str) -> TemplateElementSpec:
        return TemplateElementSpec(
            element_id="slot-1", order=0, tag="g", placeholder=placeholder
        )

    def test_body_placeholder_keeps_a_declared_bullet(self):
        shape = self._shape('<a:buFont typeface="Arial"/><a:buChar char="•"/>')
        _set_placeholder_no_inherited_bullets(shape, self._spec("body"))
        self.assertIsNotNone(shape.find(f".//{{{A_NS}}}buChar"))
        self.assertIsNone(shape.find(f".//{{{A_NS}}}buNone"))

    def test_body_placeholder_keeps_declared_auto_numbering(self):
        shape = self._shape('<a:buAutoNum type="arabicPeriod"/>')
        _set_placeholder_no_inherited_bullets(shape, self._spec("body"))
        self.assertIsNotNone(shape.find(f".//{{{A_NS}}}buAutoNum"))

    def test_body_placeholder_without_a_bullet_is_suppressed(self):
        shape = self._shape("")
        _set_placeholder_no_inherited_bullets(shape, self._spec("body"))
        self.assertIsNotNone(shape.find(f".//{{{A_NS}}}buNone"))

    def test_a_non_prose_placeholder_is_left_alone(self):
        shape = self._shape("")
        _set_placeholder_no_inherited_bullets(shape, self._spec("title"))
        self.assertIsNone(shape.find(f".//{{{A_NS}}}buNone"))


if __name__ == "__main__":
    unittest.main()
