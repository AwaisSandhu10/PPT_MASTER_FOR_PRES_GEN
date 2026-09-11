"""PowerPoint text-fitting behaviour across the template round trip.

An imported placeholder must keep the wrapping, insets, anchoring and autofit
its source template gave it, and must keep them as an inheritance chain rather
than as one resolved value stamped onto every slide. Fixtures are built inside
the test; nothing here depends on a sample deck.
"""

from __future__ import annotations

import base64
import hashlib
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from xml.etree import ElementTree as ET

import mirror_template_materialize as M
from pptx_to_svg.emu_units import Xfrm
from pptx_to_svg.shape_walker import (
    SHAPE,
    ScopedBodyPr,
    ShapeNode,
    walk_sp_tree,
)
from pptx_to_svg.slide_to_svg import (
    AssemblyContext,
    _placeholder_body_pr_metadata,
)
from svg_to_pptx.pptx_package.builder import (
    _apply_imported_body_properties,
    _clear_autofit,
    _normalize_placeholder_body_properties,
)
from svg_to_pptx.pptx_package.template_structure import (
    TemplateStructureError,
    _parse_placeholder_body_properties,
    parse_template_slides,
)

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SVG = "http://www.w3.org/2000/svg"

BOUNDS = (0, 0, 914400, 914400)


def _part(*shapes: str) -> ET.Element:
    """Wrap shape XML in the p:cSld/p:spTree a slide, layout or master needs."""
    return ET.fromstring(
        f'<root xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">'
        f'<p:cSld><p:spTree>{"".join(shapes)}</p:spTree></p:cSld></root>'
    )


def _sp(body_pr: str, *, ph: str = '<p:ph type="body" idx="1"/>') -> str:
    return (
        "<p:sp><p:nvSpPr><p:cNvPr id='2' name='x'/><p:cNvSpPr/>"
        f"<p:nvPr>{ph}</p:nvPr></p:nvSpPr>"
        "<p:spPr><a:xfrm><a:off x='0' y='0'/><a:ext cx='100' cy='100'/>"
        "</a:xfrm></p:spPr>"
        f"<p:txBody>{body_pr}<a:p><a:r><a:t>t</a:t></a:r></a:p></p:txBody>"
        "</p:sp>"
    )


def _ctx(scope: str = "slide") -> AssemblyContext:
    return AssemblyContext(
        palette=None, pkg=None, slide_part=None, placeholder_scope=scope,
    )


def _node(body_pr: str, inherited: tuple[ScopedBodyPr, ...] = ()) -> ShapeNode:
    sp = _part(_sp(body_pr))[0][0][0]
    return ShapeNode(
        kind=SHAPE,
        xml=sp,
        xfrm=Xfrm(),
        placeholder=walk_sp_tree(_part(_sp(body_pr)))[0].placeholder,
        inherited_body_properties=inherited,
    )


def _scoped(scope: str, body_pr: str) -> ScopedBodyPr:
    return ScopedBodyPr(scope, ET.fromstring(f'<a:bodyPr xmlns:a="{A}" '
                                             + body_pr[len("<a:bodyPr "):]))


def _metadata_entries(markup: str) -> dict[str, ET.Element]:
    """Decode the emitted metadata parts into scope -> bodyPr element."""
    entries: dict[str, ET.Element] = {}
    for chunk in markup.split("\n"):
        if not chunk.strip():
            continue
        element = ET.fromstring(chunk)
        payload = base64.b64decode(element.text)
        entries[element.get("data-pptx-scope")] = ET.fromstring(payload)
    return entries


def _metadata_markup(scope: str, body_pr_xml: str, *, digest=True) -> str:
    raw = body_pr_xml.encode("utf-8")
    sha = (
        f'data-pptx-ooxml-sha256="{hashlib.sha256(raw).hexdigest()}" '
        if digest else ""
    )
    return (
        f'<metadata xmlns="{SVG}" data-pptx-part="placeholder-bodypr" '
        f'data-pptx-scope="{scope}" data-pptx-encoding="base64" {sha}>'
        f'{base64.b64encode(raw).decode("ascii")}</metadata>'
    )


def _slot(*metadata: str) -> ET.Element:
    return ET.fromstring(
        f'<g xmlns="{SVG}" id="slot-1" data-pptx-placeholder="body" '
        'data-pptx-bounds="10 10 100 50" data-pptx-idx="1">'
        f'{"".join(metadata)}<text data-pptx-carrier="true"></text></g>'
    )


def _document(*slots: ET.Element) -> ET.Element:
    root = ET.Element(f"{{{SVG}}}svg", {
        "version": "1.1", "width": "960", "height": "540",
        "viewBox": "0 0 960 540",
        "data-pptx-master": "master_1", "data-pptx-master-name": "M",
        "data-pptx-layout": "layout_1", "data-pptx-layout-name": "L",
        "data-pptx-show-master-shapes": "true",
        "data-pptx-show-inherited-shapes": "true",
    })
    for slot in slots:
        root.append(slot)
    return root


def _body_pr(**attrs) -> ET.Element:
    return ET.Element(f"{{{A}}}bodyPr", {k: str(v) for k, v in attrs.items()})


def _autofit_tag(body_pr: ET.Element) -> str | None:
    for child in body_pr:
        name = child.tag.rsplit("}", 1)[-1]
        if name in {"noAutofit", "normAutofit", "spAutoFit"}:
            return name
    return None


class ImportScopeTests(unittest.TestCase):
    """The chain must be recorded per level, not merged into one value."""

    def test_walk_labels_layout_and_master_levels(self):
        nodes = walk_sp_tree(
            _part(_sp("<a:bodyPr/>")),
            layout_xml=_part(_sp('<a:bodyPr anchor="ctr"/>')),
            master_xml=_part(_sp('<a:bodyPr wrap="none"/>')),
        )
        self.assertEqual(
            [scoped.scope for scoped in nodes[0].inherited_body_properties],
            ["layout", "master"],
        )

    def test_each_level_keeps_only_its_own_declarations(self):
        """The whole point: nothing is resolved down onto the slide."""
        node = _node(
            '<a:bodyPr wrap="none"/>',
            (_scoped("layout", '<a:bodyPr anchor="ctr"><a:normAutofit '
                               'fontScale="62500"/></a:bodyPr>'),),
        )
        entries = _metadata_entries(_placeholder_body_pr_metadata(node, _ctx()))

        self.assertEqual(entries["slide"].get("wrap"), "none")
        self.assertIsNone(entries["slide"].get("anchor"))
        self.assertIsNone(_autofit_tag(entries["slide"]))
        self.assertEqual(entries["layout"].get("anchor"), "ctr")
        self.assertEqual(_autofit_tag(entries["layout"]), "normAutofit")

    def test_layout_part_labels_its_own_body_pr_as_layout(self):
        node = _node('<a:bodyPr wrap="none"/>')
        entries = _metadata_entries(
            _placeholder_body_pr_metadata(node, _ctx("layout"))
        )
        self.assertEqual(list(entries), ["layout"])

    def test_non_placeholder_shape_emits_nothing(self):
        sp = _part(_sp("<a:bodyPr/>", ph=""))[0][0][0]
        node = ShapeNode(kind=SHAPE, xml=sp, xfrm=Xfrm(), placeholder=None)
        self.assertEqual(_placeholder_body_pr_metadata(node, _ctx()), "")

    def test_relationship_bearing_body_pr_is_skipped(self):
        node = _node('<a:bodyPr><a:blipFill><a:blip r:embed="rId9"/>'
                     "</a:blipFill></a:bodyPr>")
        self.assertEqual(_placeholder_body_pr_metadata(node, _ctx()), "")


class ContractParseTests(unittest.TestCase):
    def test_returns_levels_nearest_first(self):
        slot = _slot(
            _metadata_markup("master", '<a:bodyPr xmlns:a="%s" wrap="none"/>' % A),
            _metadata_markup("slide", '<a:bodyPr xmlns:a="%s"/>' % A),
            _metadata_markup("layout", '<a:bodyPr xmlns:a="%s"/>' % A),
        )
        parsed = _parse_placeholder_body_properties(
            slot, svg_path=Path("a.svg"), element_id="slot-1",
        )
        self.assertEqual([scope for scope, _ in parsed],
                         ["slide", "layout", "master"])

    def test_other_metadata_parts_are_ignored(self):
        slot = _slot(f'<metadata xmlns="{SVG}" data-pptx-part="txbody">e30=</metadata>')
        self.assertEqual(
            _parse_placeholder_body_properties(
                slot, svg_path=Path("a.svg"), element_id="slot-1"),
            (),
        )

    def _assert_rejected(self, slot, message):
        with self.assertRaises(TemplateStructureError) as caught:
            _parse_placeholder_body_properties(
                slot, svg_path=Path("a.svg"), element_id="slot-1")
        self.assertIn(message, str(caught.exception))

    def test_rejects_unknown_scope(self):
        self._assert_rejected(
            _slot(_metadata_markup("theme", f'<a:bodyPr xmlns:a="{A}"/>')),
            "unsupported data-pptx-scope",
        )

    def test_rejects_non_base64_encoding(self):
        self._assert_rejected(
            _slot(f'<metadata xmlns="{SVG}" data-pptx-part="placeholder-bodypr" '
                  'data-pptx-scope="slide">x</metadata>'),
            "data-pptx-encoding='base64'",
        )

    def test_rejects_hash_mismatch(self):
        markup = _metadata_markup("slide", f'<a:bodyPr xmlns:a="{A}"/>')
        tampered = markup.replace(
            'data-pptx-ooxml-sha256="', 'data-pptx-ooxml-sha256="00',
        )
        self._assert_rejected(_slot(tampered), "does not match")

    def test_rejects_duplicate_scope(self):
        entry = _metadata_markup("slide", f'<a:bodyPr xmlns:a="{A}"/>')
        self._assert_rejected(_slot(entry, entry), "more than one")

    def test_round_trips_through_parse_template_slides(self):
        markup = _metadata_markup(
            "layout",
            f'<a:bodyPr xmlns:a="{A}" wrap="square"><a:normAutofit '
            'fontScale="70000"/></a:bodyPr>',
        )
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "001_content.svg"
            path.write_bytes(ET.tostring(_document(_slot(markup))))
            specs = parse_template_slides([path])
        carried = specs[0].placeholders[0].placeholder_body_properties
        self.assertEqual([scope for scope, _ in carried], ["layout"])
        self.assertIn("fontScale=\"70000\"", carried[0][1])


class ExportApplyTests(unittest.TestCase):
    LAYOUT = (
        "layout",
        f'<a:bodyPr xmlns:a="{A}" wrap="none" anchor="b" lIns="91440">'
        '<a:normAutofit fontScale="62500" lnSpcReduction="10000"/></a:bodyPr>',
    )

    def test_absent_scope_leaves_the_element_untouched(self):
        body_pr = _body_pr(wrap="square")
        self.assertFalse(
            _apply_imported_body_properties(body_pr, (self.LAYOUT,), "slide")
        )
        self.assertEqual(body_pr.get("wrap"), "square")

    def test_source_attributes_and_autofit_are_restored(self):
        body_pr = _body_pr(wrap="square", anchor="t")
        body_pr.append(ET.Element(f"{{{A}}}noAutofit"))

        self.assertTrue(
            _apply_imported_body_properties(body_pr, (self.LAYOUT,), "layout")
        )
        self.assertEqual(body_pr.get("wrap"), "none")
        self.assertEqual(body_pr.get("anchor"), "b")
        self.assertEqual(body_pr.get("lIns"), "91440")
        self.assertEqual(_autofit_tag(body_pr), "normAutofit")

    def test_norm_autofit_scale_survives_verbatim(self):
        body_pr = _body_pr()
        _apply_imported_body_properties(body_pr, (self.LAYOUT,), "layout")
        autofit = body_pr[0]
        self.assertEqual(autofit.get("fontScale"), "62500")
        self.assertEqual(autofit.get("lnSpcReduction"), "10000")

    def test_a_silent_level_does_not_erase_the_one_beneath_it(self):
        """Folding master then layout must not lose the master's autofit."""
        body_pr = _body_pr()
        chain = (
            ("master", f'<a:bodyPr xmlns:a="{A}" lIns="0" anchor="t">'
                       "<a:noAutofit/></a:bodyPr>"),
            ("layout", f'<a:bodyPr xmlns:a="{A}" anchor="ctr"/>'),
        )
        for scope in ("master", "layout"):
            _apply_imported_body_properties(body_pr, chain, scope)

        self.assertEqual(_autofit_tag(body_pr), "noAutofit")   # from master
        self.assertEqual(body_pr.get("lIns"), "0")             # from master
        self.assertEqual(body_pr.get("anchor"), "ctr")         # layout wins

    def test_folding_does_not_duplicate_a_child(self):
        body_pr = _body_pr()
        chain = (
            ("master", f'<a:bodyPr xmlns:a="{A}"><a:noAutofit/></a:bodyPr>'),
            ("layout", f'<a:bodyPr xmlns:a="{A}"><a:spAutoFit/></a:bodyPr>'),
        )
        for scope in ("master", "layout"):
            _apply_imported_body_properties(body_pr, chain, scope)
        self.assertEqual(len(body_pr), 1)
        self.assertEqual(_autofit_tag(body_pr), "spAutoFit")

    def test_clear_autofit_drops_a_generated_element(self):
        body_pr = _body_pr()
        body_pr.append(ET.Element(f"{{{A}}}noAutofit"))
        _clear_autofit(body_pr)
        self.assertIsNone(_autofit_tag(body_pr))


class ExportNormalizeTests(unittest.TestCase):
    def test_declared_values_are_never_overwritten(self):
        body_pr = _body_pr(wrap="none", anchor="b")
        _normalize_placeholder_body_properties(
            body_pr, BOUNDS, BOUNDS, default_autofit=True,
        )
        self.assertEqual(body_pr.get("wrap"), "none")
        self.assertEqual(body_pr.get("anchor"), "b")

    def test_absent_values_are_supplied(self):
        body_pr = _body_pr()
        _normalize_placeholder_body_properties(
            body_pr, BOUNDS, BOUNDS, default_autofit=True,
        )
        self.assertEqual(body_pr.get("wrap"), "square")
        self.assertEqual(body_pr.get("anchor"), "ctr")
        self.assertEqual(_autofit_tag(body_pr), "noAutofit")

    def test_slide_level_never_stamps_an_autofit(self):
        """A stamped Slide autofit would shadow the Layout's."""
        body_pr = _body_pr()
        _normalize_placeholder_body_properties(
            body_pr, BOUNDS, BOUNDS, default_autofit=False,
        )
        self.assertIsNone(_autofit_tag(body_pr))

    def test_declared_autofit_is_kept_over_the_default(self):
        body_pr = _body_pr()
        body_pr.append(ET.Element(f"{{{A}}}spAutoFit"))
        _normalize_placeholder_body_properties(
            body_pr, BOUNDS, BOUNDS, default_autofit=True,
        )
        self.assertEqual(_autofit_tag(body_pr), "spAutoFit")


class MaterializerCarryTests(unittest.TestCase):
    def _source(self, *scopes: str) -> ET.Element:
        return ET.fromstring(
            f'<g xmlns="{SVG}">'
            + "".join(
                _metadata_markup(scope, f'<a:bodyPr xmlns:a="{A}" '
                                        f'anchor="{scope[0]}"/>')
                for scope in scopes
            )
            + "</g>"
        )

    def _scopes(self, wrapper: ET.Element) -> list[str]:
        return [
            child.get("data-pptx-scope")
            for child in wrapper
            if child.get("data-pptx-part") == "placeholder-bodypr"
        ]

    def test_slot_carries_the_source_chain(self):
        wrapper = ET.Element(f"{{{SVG}}}g")
        M._preserve_body_properties(wrapper, self._source("slide", "layout"))
        self.assertEqual(self._scopes(wrapper), ["slide", "layout"])

    def test_empty_slot_still_carries_the_layout_level(self):
        """An unfilled slot has no artwork but must keep its text behaviour."""
        wrapper = ET.Element(f"{{{SVG}}}g")
        M._preserve_body_properties(wrapper, None, self._source("layout"))
        self.assertEqual(self._scopes(wrapper), ["layout"])

    def test_nearest_source_wins_for_a_repeated_scope(self):
        wrapper = ET.Element(f"{{{SVG}}}g")
        M._preserve_body_properties(
            wrapper, self._source("layout"), self._source("layout"),
        )
        self.assertEqual(self._scopes(wrapper), ["layout"])


if __name__ == "__main__":
    unittest.main()
