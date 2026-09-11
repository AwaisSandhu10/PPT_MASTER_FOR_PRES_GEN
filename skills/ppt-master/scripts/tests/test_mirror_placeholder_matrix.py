"""Placeholder/content matrix for PPTX template import and mirror materialization.

Covers every supported placeholder role in both content states, every OOXML
element that may host a <p:ph>, and the negative cases that must still fail.
Fixtures are built inside the test; nothing here depends on a sample deck.
"""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from xml.etree import ElementTree as ET

import mirror_template_materialize as M
from svg_to_pptx.pptx_package.template_structure import parse_template_slides
from template_import.manifest import (
    UnsupportedPlaceholderHostError,
    extract_placeholders,
    placeholder_semantic_role,
)

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
SVG = "http://www.w3.org/2000/svg"

TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table"
CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DIAGRAM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
OLE_URI = "http://schemas.openxmlformats.org/presentationml/2006/ole"


# ---------------------------------------------------------------------------
# OOXML part fixtures
# ---------------------------------------------------------------------------

def _part(*shapes: str, tag: str = "sld") -> ET.Element:
    body = "".join(shapes)
    return ET.fromstring(
        f'<p:{tag} xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">'
        f'<p:cSld><p:spTree>'
        f"<p:nvGrpSpPr/><p:grpSpPr/>{body}"
        f"</p:spTree></p:cSld></p:{tag}>"
    )


def _ph(ph_type: str | None, idx: str | None) -> str:
    attrs = "".join(
        f' {name}="{value}"'
        for name, value in (("type", ph_type), ("idx", idx))
        if value is not None
    )
    return f"<p:ph{attrs}/>"


def _xfrm(x=100, y=100, cx=4000000, cy=3000000, tag="a:xfrm") -> str:
    return (
        f'<{tag}><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></{tag}>'
    )


def _sp(shape_id: str, ph_type: str | None, idx: str | None,
        text: str = "", blip: bool = False, geometry: bool = True) -> str:
    body = ""
    if text:
        body = f"<a:p><a:r><a:t>{text}</a:t></a:r></a:p>"
    fill = '<a:blipFill><a:blip r:embed="rId1"/></a:blipFill>' if blip else ""
    return (
        f"<p:sp><p:nvSpPr>"
        f'<p:cNvPr id="{shape_id}" name="Shape {shape_id}"/>'
        f"<p:cNvSpPr/><p:nvPr>{_ph(ph_type, idx)}</p:nvPr></p:nvSpPr>"
        f"<p:spPr>{_xfrm() if geometry else ''}</p:spPr>{fill}"
        f"<p:txBody><a:bodyPr/>{body or '<a:p/>'}</p:txBody></p:sp>"
    )


def _pic(shape_id: str, idx: str, ph_type: str | None = "pic",
         blip: bool = True) -> str:
    fill = '<a:blip r:embed="rId1"/>' if blip else ""
    return (
        f"<p:pic><p:nvPicPr>"
        f'<p:cNvPr id="{shape_id}" name="Picture {shape_id}"/>'
        f"<p:cNvPicPr/><p:nvPr>{_ph(ph_type, idx)}</p:nvPr></p:nvPicPr>"
        f"<p:blipFill>{fill}</p:blipFill>"
        f"<p:spPr>{_xfrm()}</p:spPr></p:pic>"
    )


def _graphic_frame(shape_id: str, idx: str, uri: str,
                   ph_type: str | None = None, payload: str = "") -> str:
    return (
        f"<p:graphicFrame><p:nvGraphicFramePr>"
        f'<p:cNvPr id="{shape_id}" name="Frame {shape_id}"/>'
        f"<p:cNvGraphicFramePr/><p:nvPr>{_ph(ph_type, idx)}</p:nvPr>"
        f"</p:nvGraphicFramePr>{_xfrm(tag='p:xfrm')}"
        f'<a:graphic><a:graphicData uri="{uri}">{payload}</a:graphicData>'
        f"</a:graphic></p:graphicFrame>"
    )


def _by_idx(records, idx):
    return next(item for item in records if item["idx"] == idx)


# ---------------------------------------------------------------------------
# Import: every host element, both content states
# ---------------------------------------------------------------------------

class ImportPlaceholderHostTests(unittest.TestCase):
    """<p:ph> is legal on p:sp, p:pic and p:graphicFrame, not only p:sp."""

    def test_picture_placeholder_on_pic_is_found(self):
        records = extract_placeholders(_part(_pic("5", "13")))
        record = _by_idx(records, "13")
        self.assertEqual(record["host"], "pic")
        self.assertEqual(record["semanticRole"], "picture")
        self.assertEqual(record["contentState"], "populated")

    def test_table_placeholder_on_graphic_frame_is_found(self):
        part = _part(_graphic_frame("7", "13", TABLE_URI, "tbl",
                                    payload="<a:tbl/>"))
        record = _by_idx(extract_placeholders(part), "13")
        self.assertEqual(record["host"], "graphicFrame")
        self.assertEqual(record["semanticRole"], "table")
        self.assertEqual(record["contentState"], "populated")

    def test_chart_placeholder_on_graphic_frame_is_found(self):
        part = _part(_graphic_frame("8", "14", CHART_URI, "chart"))
        record = _by_idx(extract_placeholders(part), "14")
        self.assertEqual(record["semanticRole"], "chart")
        self.assertEqual(record["contentState"], "populated")

    def test_graphic_frame_geometry_uses_p_xfrm(self):
        """A graphicFrame keeps geometry at p:xfrm, not p:spPr/a:xfrm."""
        part = _part(_graphic_frame("7", "13", TABLE_URI, "tbl"))
        self.assertIsNotNone(_by_idx(extract_placeholders(part), "13")["geometry"])

    def test_empty_and_populated_text_placeholders_differ(self):
        part = _part(_sp("2", "body", "10", text="Hello"),
                     _sp("3", "body", "11"))
        records = extract_placeholders(part)
        self.assertEqual(_by_idx(records, "10")["contentState"], "populated")
        self.assertEqual(_by_idx(records, "11")["contentState"], "empty")

    def test_empty_picture_placeholder_hosted_by_sp(self):
        """An unfilled picture placeholder is an empty p:sp, not a p:pic."""
        record = _by_idx(extract_placeholders(_part(_sp("4", "pic", "16"))), "16")
        self.assertEqual(record["host"], "sp")
        self.assertEqual(record["semanticRole"], "picture")
        self.assertEqual(record["contentState"], "empty")

    def test_media_placeholder_keeps_its_role(self):
        record = _by_idx(
            extract_placeholders(_part(_pic("9", "12", ph_type="media"))), "12",
        )
        self.assertEqual(record["semanticRole"], "media")

    def test_group_hosted_placeholder_is_named_not_misbound(self):
        grp = (
            f"<p:grpSp><p:nvGrpSpPr>"
            f'<p:cNvPr id="20" name="Group 20"/><p:cNvGrpSpPr/>'
            f"<p:nvPr>{_ph('obj', '9')}</p:nvPr></p:nvGrpSpPr>"
            f"<p:grpSpPr>{_xfrm()}</p:grpSpPr></p:grpSp>"
        )
        with self.assertRaises(UnsupportedPlaceholderHostError) as ctx:
            extract_placeholders(_part(grp))
        self.assertIn("grpSp", str(ctx.exception))

    def test_placeholder_inside_a_group_is_still_found(self):
        grp = (
            f"<p:grpSp><p:nvGrpSpPr>"
            f'<p:cNvPr id="20" name="Group 20"/><p:cNvGrpSpPr/><p:nvPr/>'
            f"</p:nvGrpSpPr><p:grpSpPr>{_xfrm()}</p:grpSpPr>"
            f'{_sp("21", "body", "7", text="nested")}</p:grpSp>'
        )
        self.assertEqual(len(extract_placeholders(_part(grp))), 1)


class ImportTypeResolutionTests(unittest.TestCase):
    def test_typeless_slide_placeholder_resolves_through_the_layout(self):
        layout = _part(_sp("9", "pic", "13"), tag="sldLayout")
        slide = _part(_sp("4", None, "13"))
        record = _by_idx(extract_placeholders(slide, layout_root=layout), "13")
        self.assertIsNone(record["type"])
        self.assertEqual(record["resolvedType"], "pic")
        self.assertEqual(record["semanticRole"], "picture")

    def test_typeless_placeholder_without_a_layout_stays_object(self):
        record = _by_idx(extract_placeholders(_part(_sp("4", None, "13"))), "13")
        self.assertEqual(record["semanticRole"], "object")

    def test_remaining_ooxml_types_never_fall_through_to_other(self):
        for ph_type in ("clipArt", "dgm", "sldImg", "hdr", "ctrTitle", "media"):
            with self.subTest(ph_type=ph_type):
                self.assertNotEqual(placeholder_semantic_role(ph_type), "other")

    def test_content_state_absent_is_unrepresentable(self):
        """A record cannot say "absent"; absence is a missing record."""
        part = _part(_sp("2", "body", "10"), _pic("5", "13"))
        for record in extract_placeholders(part):
            self.assertIn(record["contentState"], {"empty", "populated"})


# ---------------------------------------------------------------------------
# Materialization: role x content state
# ---------------------------------------------------------------------------

def _plan(role, slot_id="slot-l1-1", idx=1):
    return M.SlotPlan(
        slot_id=slot_id, semantic_role=role, placeholder_type=None,
        idx=idx, shape_id="1", bounds=(10.0, 20.0, 300.0, 200.0),
    )


def _svg(markup: str) -> ET.Element:
    return ET.fromstring(f'<g xmlns="{SVG}">{markup}</g>')


def _wrap(plan, source=None, *, state, layout_guide=None, master_guide=None):
    return M._slot_wrapper(
        plan, source, layout_guide=layout_guide,
        master_guide=master_guide, content_state=state,
    )[0]


def _tags(wrapper):
    return [child.tag.rsplit("}", 1)[-1] for child in wrapper]


def _carrier(wrapper):
    return next(
        (c for c in wrapper if c.get("data-pptx-carrier") == "true"), None,
    )


class UnboundSlotTests(unittest.TestCase):
    """Every role must materialize as a valid slot with no bound content."""

    ROLES = ("title", "subtitle", "body", "date", "footer", "slide-number",
             "picture", "media", "object", "chart", "table")

    def test_every_role_materializes_when_absent(self):
        for role in self.ROLES:
            for state in ("absent", "empty"):
                with self.subTest(role=role, state=state):
                    wrapper = _wrap(_plan(role), None, state=state)
                    self.assertEqual(
                        wrapper.get("data-pptx-placeholder"), role,
                    )
                    self.assertEqual(wrapper.get("data-pptx-idx"), "1")
                    self.assertEqual(
                        wrapper.get("data-pptx-bounds"), "10 20 300 200",
                    )
                    self.assertIsNotNone(_carrier(wrapper))

    def test_unbound_picture_keeps_an_image_carrier(self):
        """The blank tag matches the filled tag so one Layout stays consistent."""
        wrapper = _wrap(_plan("picture"), None, state="empty")
        self.assertEqual(_tags(wrapper), ["image"])

    def test_unbound_object_uses_a_shape_carrier_not_a_proxy(self):
        wrapper = _wrap(_plan("object"), None, state="empty")
        self.assertEqual(_tags(wrapper), ["rect"])
        self.assertIsNone(wrapper.get("data-pptx-binding"))

    def test_unbound_table_declares_empty_binding_without_a_marker(self):
        wrapper = _wrap(_plan("table"), None, state="empty")
        self.assertEqual(wrapper.get("data-pptx-binding"), "empty")
        carrier = _carrier(wrapper)
        self.assertEqual(carrier.tag.rsplit("}", 1)[-1], "text")
        self.assertIsNone(carrier.get("data-pptx-replace-with"))

    def test_unbound_chart_declares_empty_binding(self):
        wrapper = _wrap(_plan("chart"), None, state="empty")
        self.assertEqual(wrapper.get("data-pptx-binding"), "empty")

    def test_a_layout_prompt_guide_is_never_treated_as_content(self):
        """The guide supplies bounds, never artwork."""
        guide = _svg(
            '<path d="M 0 0 L 10 10" fill="none" data-pptx-part="geometry"/>'
            "<text>Click icon to add table</text>"
        )
        wrapper = _wrap(_plan("table"), None, state="absent",
                        layout_guide=guide)
        self.assertEqual(wrapper.get("data-pptx-binding"), "empty")
        self.assertNotIn("Click icon to add table",
                         "".join(wrapper.itertext()))


class BoundSlotTests(unittest.TestCase):
    def test_picture_binds_a_single_image(self):
        source = _svg('<image href="a.png" width="10" height="10"/>')
        self.assertEqual(_tags(_wrap(_plan("picture"), source,
                                     state="populated")), ["image"])

    def test_cropped_picture_binds_the_outer_svg_only(self):
        """<g><svg><image/></svg></g> is one carrier, not two."""
        source = _svg(
            '<svg viewBox="0 0 1 1" width="10" height="10">'
            '<image href="a.png" width="10" height="10"/></svg>'
        )
        self.assertEqual(_tags(_wrap(_plan("picture"), source,
                                     state="populated")), ["svg"])

    def test_object_with_one_atom_binds_a_carrier(self):
        source = _svg("<text>Text or image</text>")
        wrapper = _wrap(_plan("object"), source, state="populated")
        self.assertIsNone(wrapper.get("data-pptx-binding"))
        self.assertIsNotNone(_carrier(wrapper))

    def test_object_with_composite_content_uses_a_proxy(self):
        source = _svg('<text>a</text><rect width="4" height="4" fill="#000"/>')
        wrapper = _wrap(_plan("object"), source, state="populated")
        self.assertEqual(wrapper.get("data-pptx-binding"), "proxy")
        self.assertIsNone(_carrier(wrapper))

    def test_native_table_marker_binds_as_a_carrier(self):
        source = ET.fromstring(
            f'<g xmlns="{SVG}" data-pptx-replace-with="table">'
            '<metadata type="application/json">{}</metadata>'
            '<rect width="4" height="4" fill="#000"/></g>'
        )
        wrapper = _wrap(_plan("table"), source, state="populated")
        carrier = _carrier(wrapper)
        self.assertEqual(carrier.get("data-pptx-replace-with"), "table")

    def test_fallback_only_table_keeps_its_slot_as_a_proxy(self):
        """Native conversion is best-effort; fallback artwork is still content."""
        source = ET.fromstring(
            f'<g xmlns="{SVG}" data-pptx-import-source="pptx" '
            'data-pptx-replacement-status="unsupported-table-style">'
            '<text>cell</text><rect width="4" height="4" fill="#000"/></g>'
        )
        wrapper = _wrap(_plan("table"), source, state="populated")
        self.assertEqual(wrapper.get("data-pptx-binding"), "proxy")


class PlaceholderIdentityTests(unittest.TestCase):
    """A bound carrier keeps the canonical <p:ph> identity of its placeholder."""

    SOURCE = (
        f'<g xmlns="{SVG}" data-pptx-object="picture" data-ph-type="pic" '
        'data-pptx-placeholder-index="13" data-pptx-placeholder-size="quarter">'
        '<svg data-pptx-object="picture" data-ph-type="pic" '
        'data-pptx-placeholder-index="13" data-pptx-placeholder-size="quarter" '
        'viewBox="0 0 1 1" width="10" height="10">'
        '<image href="a.png" width="10" height="10"/></svg></g>'
    )

    def test_cropped_carrier_keeps_type_index_and_size(self):
        source = ET.fromstring(self.SOURCE)
        carrier = _carrier(_wrap(_plan("picture"), source, state="populated"))
        self.assertEqual(carrier.tag.rsplit("}", 1)[-1], "svg")
        self.assertEqual(carrier.get("data-ph-type"), "pic")
        self.assertEqual(carrier.get("data-pptx-placeholder-index"), "13")
        self.assertEqual(carrier.get("data-pptx-placeholder-size"), "quarter")

    def test_identity_is_inherited_when_the_carrier_lacks_it(self):
        source = ET.fromstring(
            f'<g xmlns="{SVG}" data-ph-type="pic" '
            'data-pptx-placeholder-index="7" data-pptx-placeholder-size="half">'
            '<image href="a.png" width="10" height="10"/></g>'
        )
        carrier = _carrier(_wrap(_plan("picture"), source, state="populated"))
        self.assertEqual(carrier.get("data-ph-type"), "pic")
        self.assertEqual(carrier.get("data-pptx-placeholder-index"), "7")
        self.assertEqual(carrier.get("data-pptx-placeholder-size"), "half")

    def test_ph_type_survives_ir_stripping_only_on_a_carrier(self):
        root = ET.fromstring(
            f'<svg xmlns="{SVG}">'
            '<image data-pptx-carrier="true" data-ph-type="pic"/>'
            '<rect data-ph-type="pic"/></svg>'
        )
        M._strip_source_refs(root)
        self.assertEqual(root[0].get("data-ph-type"), "pic")
        self.assertIsNone(root[1].get("data-ph-type"))

    def test_carrier_identity_rebuilds_a_native_placeholder_marker(self):
        from svg_to_pptx.drawingml.elements import _imported_placeholder_xml
        source = ET.fromstring(self.SOURCE)
        carrier = _carrier(_wrap(_plan("picture"), source, state="populated"))
        self.assertEqual(
            _imported_placeholder_xml(carrier),
            '<p:ph type="pic" idx="13" sz="quarter"/>',
        )


class SlotValidationTests(unittest.TestCase):
    """Validation must stay strict for genuinely malformed content."""

    def test_two_picture_carriers_still_fail(self):
        source = _svg(
            '<image href="a.png" width="10" height="10"/>'
            '<image href="b.png" width="10" height="10"/>'
        )
        with self.assertRaises(M.MirrorMaterializationError):
            _wrap(_plan("picture"), source, state="populated")

    def test_table_without_marker_or_fallback_still_fails(self):
        source = _svg('<rect width="4" height="4" fill="#000"/>')
        with self.assertRaises(M.MirrorMaterializationError) as ctx:
            _wrap(_plan("table"), source, state="populated")
        self.assertIn("native marker", str(ctx.exception))

    def test_populated_picture_with_no_carrier_still_fails(self):
        source = _svg('<rect width="4" height="4" fill="#000"/>')
        with self.assertRaises(M.MirrorMaterializationError):
            _wrap(_plan("picture"), source, state="populated")


# ---------------------------------------------------------------------------
# Cross-prototype consistency and the downstream contract
# ---------------------------------------------------------------------------

def _document(*slots: ET.Element, layout="layout_1", master="master_1"):
    root = ET.Element(f"{{{SVG}}}svg", {
        "version": "1.1", "width": "960", "height": "540",
        "viewBox": "0 0 960 540",
        "data-pptx-master": master, "data-pptx-master-name": "M",
        "data-pptx-layout": layout, "data-pptx-layout-name": "L",
        "data-pptx-show-master-shapes": "true",
        "data-pptx-show-inherited-shapes": "true",
    })
    for slot in slots:
        root.append(slot)
    return root


class ProxyPromotionTests(unittest.TestCase):
    def test_one_slot_filled_on_one_page_and_blank_on_another_agrees(self):
        """A Layout declares one binding, so proxy must win layout-wide."""
        filled = _wrap(_plan("object"), _svg('<text>a</text><rect width="4" '
                                             'height="4" fill="#000"/>'),
                       state="populated")
        blank = _wrap(_plan("object"), None, state="empty")
        roots = [(Path("a.svg"), _document(filled)),
                 (Path("b.svg"), _document(blank))]
        self.assertEqual(M._promote_proxy_bindings(roots), 1)
        promoted = list(roots[1][1])[0]
        self.assertEqual(promoted.get("data-pptx-binding"), "proxy")
        self.assertIsNone(_carrier(promoted))

    def test_unrelated_layouts_are_not_promoted(self):
        filled = _wrap(_plan("object"), _svg('<text>a</text><rect width="4" '
                                             'height="4" fill="#000"/>'),
                       state="populated")
        blank = _wrap(_plan("object"), None, state="empty")
        roots = [(Path("a.svg"), _document(filled, layout="layout_1")),
                 (Path("b.svg"), _document(blank, layout="layout_2"))]
        self.assertEqual(M._promote_proxy_bindings(roots), 0)


class DownstreamContractTests(unittest.TestCase):
    """Every slot this module emits must satisfy the export contract."""

    CASES = (
        ("title", None, "absent"),
        ("body", None, "empty"),
        ("picture", None, "empty"),
        ("media", None, "empty"),
        ("object", None, "empty"),
        ("table", None, "empty"),
        ("chart", None, "absent"),
    )

    def test_unbound_slots_validate(self):
        with TemporaryDirectory() as tmp:
            slots = [
                _wrap(_plan(role, slot_id=f"slot-l1-{n}", idx=n), source,
                      state=state)
                for n, (role, source, state) in enumerate(self.CASES, start=1)
            ]
            path = Path(tmp) / "001_content.svg"
            path.write_bytes(ET.tostring(_document(*slots)))
            specs = parse_template_slides([path])
            self.assertEqual(len(specs[0].placeholders), len(self.CASES))

    def test_bound_slots_validate(self):
        with TemporaryDirectory() as tmp:
            marker = ET.fromstring(
                f'<g xmlns="{SVG}" data-pptx-replace-with="table">'
                '<metadata type="application/json">{}</metadata>'
                '<rect width="4" height="4" fill="#000"/></g>'
            )
            slots = [
                _wrap(_plan("title", slot_id="slot-l1-1", idx=1),
                      _svg("<text>Heading</text>"), state="populated"),
                _wrap(_plan("picture", slot_id="slot-l1-2", idx=2),
                      _svg('<image href="a.png" width="10" height="10"/>'),
                      state="populated"),
                _wrap(_plan("object", slot_id="slot-l1-3", idx=3),
                      _svg('<text>a</text><rect width="4" height="4" '
                           'fill="#000"/>'), state="populated"),
                _wrap(_plan("table", slot_id="slot-l1-4", idx=4), marker,
                      state="populated"),
            ]
            path = Path(tmp) / "001_content.svg"
            path.write_bytes(ET.tostring(_document(*slots)))
            specs = parse_template_slides([path])
            bindings = {
                item.placeholder: item.placeholder_binding
                for item in specs[0].placeholders
            }
            self.assertEqual(bindings["object"], "proxy")
            self.assertEqual(bindings["table"], "carrier")


# ---------------------------------------------------------------------------
# End-to-end structure: prototype counts are derived, never constants
# ---------------------------------------------------------------------------

import hashlib
import json
import zipfile

from svg_authoring_view import semantic_subtree_sha256

DML = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _shape(scope: str, shape_id: int, label: str) -> str:
    return (
        f'<g id="{scope}-shape-{shape_id}" data-pptx-object="shape" '
        f'data-pptx-shape-id="{shape_id}" data-pptx-shape-scope="{scope}" '
        f'data-pptx-frame="10 20 300 200" data-ph-type="body" '
        f'data-pptx-placeholder-index="{shape_id}">'
        f'<text x="10" y="40">{label}</text></g>'
    )


def _page(body: str) -> str:
    return (
        f'<svg xmlns="{SVG}" version="1.1" width="960" height="540" '
        f'viewBox="0 0 960 540">{body}</svg>'
    )


def _authored(scope: str, shape_id: int, label: str) -> str:
    return (
        f'<g id="{scope}-shape-{shape_id}" data-pptx-object="shape" '
        f'data-pptx-frame="10 20 300 200" data-ph-type="body" '
        f'data-pptx-placeholder-index="{shape_id}" '
        f'data-pptx-source-ref="{scope}:{shape_id}">'
        f'<text x="10" y="40">{label}</text></g>'
    )


def _minimal_pptx(path: Path, theme_parts: list[str]) -> None:
    theme = (
        f'<a:theme xmlns:a="{DML}" name="t"><a:themeElements/></a:theme>'
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        for part in theme_parts:
            archive.writestr(part, theme)


def _build_workspace(root: Path, *, masters: int, layouts: int, slides: int):
    """Write one import workspace with the given structural dimensions."""
    for sub in ("analysis", "sources", "svg", "authoring-svg"):
        (root / sub).mkdir(parents=True, exist_ok=True)

    theme_parts = [f"ppt/theme/theme{n}.xml" for n in range(1, masters + 1)]
    pptx = root / "sources" / "source.pptx"
    _minimal_pptx(pptx, theme_parts)

    master_keys = [f"master_{n:02d}" for n in range(1, masters + 1)]
    layout_keys = [f"layout_{n:02d}" for n in range(1, layouts + 1)]

    documents, native_masters, native_layouts, native_slides = [], [], [], []
    inh = {"masters": [], "layouts": [], "slides": []}

    def emit(name: str, scope: str, shape_id: int, label: str):
        lossless = _page(_shape(scope, shape_id, label)).encode()
        authoring = _page(_authored(scope, shape_id, label)).encode()
        (root / "svg" / name).write_bytes(lossless)
        (root / "authoring-svg" / name).write_bytes(authoring)
        subtree = ET.fromstring(authoring)[0]
        documents.append({
            "source": name,
            "authoring": name,
            "source_sha256": _sha(lossless),
            "initial_authoring_sha256": _sha(authoring),
            "source_refs": {
                f"{scope}:{shape_id}": {
                    "source_path": [0],
                    "initial_authoring_subtree_sha256": semantic_subtree_sha256(
                        subtree,
                    ),
                    "representation": "inline",
                },
            },
        })

    for index, key in enumerate(master_keys, start=1):
        name = f"{key}_slideMaster{index}.svg"
        part = f"ppt/slideMasters/slideMaster{index}.xml"
        emit(name, "master", 100 + index, f"master {index}")
        native_masters.append({
            "key": key, "name": f"Master {index}", "packagePart": part,
            "themePart": theme_parts[index - 1], "theme": {},
            "backgroundAsset": None, "imageAssets": [], "shapeImageAssets": [],
            "drawableShapeCount": 0, "layoutKeys": [],
            "placeholders": [{
                "type": "body", "resolvedType": "body", "idx": str(100 + index),
                "size": None, "orient": None, "semanticRole": "body",
                "host": "sp", "contentState": "populated",
                "shapeId": str(100 + index), "shapeName": "Body",
                "geometry": {"x": 10, "y": 20, "width": 300, "height": 200},
                "textSamples": [],
            }],
        })
        inh["masters"].append({
            "file": name, "partPath": part,
            "themePath": theme_parts[index - 1],
        })

    for index, key in enumerate(layout_keys, start=1):
        name = f"{key}_slideLayout{index}.svg"
        part = f"ppt/slideLayouts/slideLayout{index}.xml"
        master_index = (index - 1) % masters
        emit(name, "layout", 200 + index, f"layout {index}")
        native_layouts.append({
            "key": key, "name": f"Layout {index}", "type": None,
            "packagePart": part, "masterKey": master_keys[master_index],
            "showMasterShapes": True, "backgroundAsset": None,
            "imageAssets": [], "shapeImageAssets": [], "drawableShapeCount": 0,
            "usedBySlides": [], "svgFile": name,
            "placeholders": [{
                "type": "body", "resolvedType": "body", "idx": str(200 + index),
                "size": None, "orient": None, "semanticRole": "body",
                "host": "sp", "contentState": "populated",
                "shapeId": str(200 + index), "shapeName": "Body",
                "geometry": {"x": 10, "y": 20, "width": 300, "height": 200},
                "textSamples": [],
            }],
        })
        inh["layouts"].append({
            "file": name, "partPath": part,
            "master": f"{master_keys[master_index]}_slideMaster"
                      f"{master_index + 1}.svg",
            "parentPartPath": f"ppt/slideMasters/slideMaster"
                              f"{master_index + 1}.xml",
            "themePath": theme_parts[master_index], "showMasterShapes": True,
        })

    # Slides deliberately reuse a subset of the Layouts, so some Layouts are
    # demonstrated (one or more times) and the rest are not demonstrated at all.
    for index in range(1, slides + 1):
        name = f"slide_{index:02d}.svg"
        layout_index = (index - 1) % max(1, min(layouts, slides))
        layout = native_layouts[layout_index]
        emit(name, "slide", 300 + index, f"slide {index}")
        native_slides.append({
            "index": index, "packagePart": f"ppt/slides/slide{index}.xml",
            "pageType": "content_candidate", "layoutKey": layout["key"],
            "masterKey": layout["masterKey"], "showInheritedShapes": True,
            "layeredSvgFile": name, "flatSvgFile": name,
            "placeholders": [{
                "type": "body", "resolvedType": "body", "idx": layout[
                    "placeholders"][0]["idx"],
                "size": None, "orient": None, "semanticRole": "body",
                "host": "sp", "contentState": "populated",
                "shapeId": str(300 + index), "shapeName": "Body",
                "geometry": {"x": 10, "y": 20, "width": 300, "height": 200},
                "textSamples": [],
            }],
        })
        inh["slides"].append({
            "file": name, "index": index, "layout": layout["svgFile"],
            "master": f"{layout['masterKey']}_slideMaster"
                      f"{master_keys.index(layout['masterKey']) + 1}.svg",
            "showInheritedShapes": True,
        })

    (root / "analysis" / "native_structure.json").write_text(json.dumps({
        "schema": "ppt-master.native-structure.v1",
        "source": {"name": "synthetic.pptx",
                   "templateFile": "sources/source.pptx",
                   "sha256": _sha(pptx.read_bytes())},
        "slideSize": {"width_px": 960, "height_px": 540},
        "strategy": {},
        "masters": native_masters,
        "layouts": native_layouts,
        "slides": native_slides,
    }))
    (root / "svg" / "inheritance.json").write_text(json.dumps(inh))
    (root / "authoring-svg" / "authoring_manifest.json").write_text(json.dumps({
        "schema": "ppt-master.svg-authoring-ir.v1",
        "projection_kind": "layered",
        "source_root": "../svg",
        "authoring_root": ".",
        "source_ref_attribute": "data-pptx-source-ref",
        "file_count": len(documents),
        "source_ref_count": len(documents),
        "documents": documents,
    }))
    return {
        "masters": set(master_keys),
        "layouts": set(layout_keys),
        "slide_layouts": {item["layoutKey"] for item in native_slides},
        "slides": slides,
    }


class PrototypeCountInvariantTests(unittest.TestCase):
    """One prototype per source Slide, for any deck; never a fixed number."""

    DIMENSIONS = ((1, 3, 2), (3, 7, 4))

    def test_counts_follow_the_source_roster(self):
        for masters, layouts, slides in self.DIMENSIONS:
            with self.subTest(masters=masters, layouts=layouts, slides=slides):
                with TemporaryDirectory() as tmp:
                    source = Path(tmp) / "import"
                    facts = _build_workspace(
                        source, masters=masters, layouts=layouts, slides=slides,
                    )
                    report = M.materialize_mirror_template(
                        source, Path(tmp) / "template",
                    )
                    expected = facts["slides"]
                    self.assertEqual(report["template_svg_count"], expected)

                    emitted = sorted(
                        (Path(tmp) / "template" / "templates").glob("*.svg"),
                    )
                    self.assertEqual(len(emitted), expected)
                    seen_layouts, seen_masters = set(), set()
                    for path in emitted:
                        root = ET.parse(path).getroot()
                        seen_layouts.add(root.get("data-pptx-layout"))
                        seen_masters.add(root.get("data-pptx-master"))
                    # Retained structure is exactly what the Slides reach.
                    self.assertEqual(seen_layouts, facts["slide_layouts"])
                    self.assertTrue(seen_masters <= facts["masters"])

    def test_layouts_no_slide_uses_are_omitted(self):
        """Unused Layouts are reported as omitted, not silently dropped."""
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "import"
            facts = _build_workspace(source, masters=2, layouts=6, slides=2)
            report = M.materialize_mirror_template(
                source, Path(tmp) / "template",
            )
            undemonstrated = facts["layouts"] - facts["slide_layouts"]
            self.assertTrue(undemonstrated, "fixture must leave Layouts unused")
            self.assertEqual(
                set(report["omitted_structure"]["layout_keys"]), undemonstrated,
            )
            emitted = {
                ET.parse(path).getroot().get("data-pptx-layout")
                for path in (Path(tmp) / "template" / "templates").glob("*.svg")
            }
            self.assertFalse(undemonstrated & emitted)


class ParagraphBlockEvidenceTests(unittest.TestCase):
    """An explicit paragraph marker outranks the dy-gap heuristic.

    A placeholder merged from several source text blocks is one PowerPoint text
    frame even when the blocks sit far apart. Losing that to a distance guess
    lowers the carrier into several shapes, which no <p:ph> can bind.
    """

    @staticmethod
    def _block(gap: float, marker: bool):
        marked = ' data-paragraph-soft-break="0"' if marker else ""
        return ET.fromstring(
            f'<text xmlns="{SVG}" x="10" y="20" font-size="18.67">'
            '<tspan x="10" dy="0">first</tspan>'
            '<tspan x="10" dy="18.67">second</tspan>'
            f'<tspan x="10" dy="{gap}"{marked}>far below</tspan></text>'
        )

    def test_wide_gap_without_a_marker_is_still_a_section_break(self):
        from svg_finalize.flatten_tspan import classify_paragraph_block
        self.assertIsNone(
            classify_paragraph_block(self._block(89.33, marker=False),
                                     preserve_line_breaks=True),
        )

    def test_wide_gap_with_an_explicit_marker_stays_one_block(self):
        from svg_finalize.flatten_tspan import classify_paragraph_block
        result = classify_paragraph_block(self._block(89.33, marker=True),
                                          preserve_line_breaks=True)
        self.assertIsNotNone(result)
        self.assertEqual(result[2][-1], "paragraph")

    def test_overlapping_lines_are_still_rejected(self):
        from svg_finalize.flatten_tspan import classify_paragraph_block
        self.assertIsNone(
            classify_paragraph_block(self._block(1.0, marker=True),
                                     preserve_line_breaks=True),
        )
