"""Generic template-context extraction from arbitrary PPTX templates.

Every fixture is a synthetic package built inside the test, with deliberately
unremarkable dimensions and counts. No assertion may encode a number taken from
any particular source deck: the shapes under test are the relationships, not the
totals.
"""

from __future__ import annotations

import json
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from template_import.template_context import (
    SCHEMA,
    build_template_context,
    write_template_context,
)

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EMU_PER_IN = 914400


def _sp(
    *,
    ph: str | None,
    name: str,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    body_pr: str = "<a:bodyPr/>",
    lst_style: str = "",
    text: str = "",
    paragraphs: str | None = None,
) -> str:
    nv_pr = f"<p:nvPr>{ph or ''}</p:nvPr>"
    body = paragraphs if paragraphs is not None else (
        f"<a:p><a:r><a:t>{text}</a:t></a:r></a:p>" if text else "<a:p/>"
    )
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='2' name='{name}'/><p:cNvSpPr/>{nv_pr}"
        "</p:nvSpPr><p:spPr><a:xfrm>"
        f"<a:off x='{int(x_in * EMU_PER_IN)}' y='{int(y_in * EMU_PER_IN)}'/>"
        f"<a:ext cx='{int(w_in * EMU_PER_IN)}' cy='{int(h_in * EMU_PER_IN)}'/>"
        "</a:xfrm></p:spPr>"
        f"<p:txBody>{body_pr}{lst_style}{body}</p:txBody></p:sp>"
    )


def _pic(*, ph: str, name: str, x_in: float, y_in: float,
         w_in: float, h_in: float) -> str:
    return (
        f"<p:pic><p:nvPicPr><p:cNvPr id='9' name='{name}'/><p:cNvPicPr/>"
        f"<p:nvPr>{ph}</p:nvPr></p:nvPicPr>"
        "<p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill>"
        "<p:spPr><a:xfrm>"
        f"<a:off x='{int(x_in * EMU_PER_IN)}' y='{int(y_in * EMU_PER_IN)}'/>"
        f"<a:ext cx='{int(w_in * EMU_PER_IN)}' cy='{int(h_in * EMU_PER_IN)}'/>"
        "</a:xfrm></p:spPr></p:pic>"
    )


def _part(root_tag: str, shapes: str, *, name: str = "L", extra: str = "") -> str:
    return (
        f'<?xml version="1.0"?>'
        f'<p:{root_tag} xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">'
        f'<p:cSld name="{name}"><p:spTree>'
        "<p:nvGrpSpPr><p:cNvPr id='1' name=''/><p:cNvGrpSpPr/><p:nvPr/>"
        "</p:nvGrpSpPr><p:grpSpPr/>"
        f"{shapes}</p:spTree></p:cSld>{extra}</p:{root_tag}>"
    )


def _rels(entries: list[tuple[str, str, str]]) -> str:
    items = "".join(
        f'<Relationship Id="{rid}" Type="{RT}/{kind}" Target="{target}"/>'
        for rid, kind, target in entries
    )
    return (
        f'<?xml version="1.0"?>'
        f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{items}</Relationships>"
    )


class Deck:
    """A minimal but valid PPTX built entirely in-test."""

    def __init__(self, *, width_in: float = 13.333, height_in: float = 7.5):
        self.width_in = width_in
        self.height_in = height_in
        self.masters: list[tuple[str, str]] = []   # (master_xml, theme_name)
        self.layouts: list[tuple[int, str]] = []   # (master_index, layout_xml)
        self.slides: list[tuple[int, str]] = []    # (layout_index, slide_xml)

    def add_master(self, xml: str) -> int:
        self.masters.append((xml, "Test Theme"))
        return len(self.masters) - 1

    def add_layout(self, master_index: int, xml: str) -> int:
        self.layouts.append((master_index, xml))
        return len(self.layouts) - 1

    def add_slide(self, layout_index: int, xml: str) -> int:
        self.slides.append((layout_index, xml))
        return len(self.slides) - 1

    def _theme(self) -> str:
        return (
            f'<?xml version="1.0"?><a:theme xmlns:a="{A}" name="T">'
            "<a:themeElements><a:clrScheme name='c'>"
            + "".join(
                f"<a:{n}><a:srgbClr val='000000'/></a:{n}>"
                for n in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2",
                          "accent3", "accent4", "accent5", "accent6",
                          "hlink", "folHlink")
            )
            + "</a:clrScheme>"
            "<a:fontScheme name='f'>"
            "<a:majorFont><a:latin typeface='Major Test Face'/><a:ea typeface=''/>"
            "<a:cs typeface=''/></a:majorFont>"
            "<a:minorFont><a:latin typeface='Minor Test Face'/><a:ea typeface=''/>"
            "<a:cs typeface=''/></a:minorFont></a:fontScheme>"
            "<a:fmtScheme name='s'><a:fillStyleLst/><a:lnStyleLst/>"
            "<a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>"
            "</a:themeElements></a:theme>"
        )

    def write(self, path: Path) -> Path:
        layouts_by_master: dict[int, list[int]] = {}
        for index, (master_index, _) in enumerate(self.layouts):
            layouts_by_master.setdefault(master_index, []).append(index)

        with zipfile.ZipFile(path, "w") as z:
            overrides = []
            z.writestr("_rels/.rels", _rels([
                ("rId1", "officeDocument", "ppt/presentation.xml"),
            ]))

            master_ids = "".join(
                f'<p:sldMasterId id="{2147483648 + i}" r:id="rIdM{i}"/>'
                for i in range(len(self.masters))
            )
            slide_ids = "".join(
                f'<p:sldId id="{256 + i}" r:id="rIdS{i}"/>'
                for i in range(len(self.slides))
            )
            z.writestr("ppt/presentation.xml",
                f'<?xml version="1.0"?><p:presentation xmlns:p="{P}" '
                f'xmlns:r="{R}" xmlns:a="{A}">'
                f"<p:sldMasterIdLst>{master_ids}</p:sldMasterIdLst>"
                f"<p:sldIdLst>{slide_ids}</p:sldIdLst>"
                f'<p:sldSz cx="{int(self.width_in * EMU_PER_IN)}" '
                f'cy="{int(self.height_in * EMU_PER_IN)}"/>'
                "</p:presentation>")
            z.writestr("ppt/_rels/presentation.xml.rels", _rels(
                [(f"rIdM{i}", "slideMaster", f"slideMasters/slideMaster{i + 1}.xml")
                 for i in range(len(self.masters))]
                + [(f"rIdS{i}", "slide", f"slides/slide{i + 1}.xml")
                   for i in range(len(self.slides))]
            ))

            for i, (xml, _) in enumerate(self.masters):
                own = layouts_by_master.get(i, [])
                ids = "".join(
                    f'<p:sldLayoutId id="{2147484000 + j}" r:id="rIdL{j}"/>'
                    for j in own
                )
                xml = xml.replace(
                    "</p:cSld>",
                    f"</p:cSld><p:sldLayoutIdLst>{ids}</p:sldLayoutIdLst>",
                )
                z.writestr(f"ppt/slideMasters/slideMaster{i + 1}.xml", xml)
                z.writestr(f"ppt/slideMasters/_rels/slideMaster{i + 1}.xml.rels",
                    _rels(
                        [(f"rIdL{j}", "slideLayout",
                          f"../slideLayouts/slideLayout{j + 1}.xml") for j in own]
                        + [("rIdT", "theme", f"../theme/theme{i + 1}.xml")]
                    ))
                z.writestr(f"ppt/theme/theme{i + 1}.xml", self._theme())
                overrides.append((f"/ppt/slideMasters/slideMaster{i + 1}.xml",
                                  "slideMaster"))

            for j, (master_index, xml) in enumerate(self.layouts):
                z.writestr(f"ppt/slideLayouts/slideLayout{j + 1}.xml", xml)
                z.writestr(f"ppt/slideLayouts/_rels/slideLayout{j + 1}.xml.rels",
                    _rels([("rIdM", "slideMaster",
                            f"../slideMasters/slideMaster{master_index + 1}.xml")]))
                overrides.append((f"/ppt/slideLayouts/slideLayout{j + 1}.xml",
                                  "slideLayout"))

            for k, (layout_index, xml) in enumerate(self.slides):
                z.writestr(f"ppt/slides/slide{k + 1}.xml", xml)
                z.writestr(f"ppt/slides/_rels/slide{k + 1}.xml.rels",
                    _rels([("rIdL", "slideLayout",
                            f"../slideLayouts/slideLayout{layout_index + 1}.xml")]))
                overrides.append((f"/ppt/slides/slide{k + 1}.xml", "slide"))

            types = {
                "slideMaster": "application/vnd.openxmlformats-officedocument."
                               "presentationml.slideMaster+xml",
                "slideLayout": "application/vnd.openxmlformats-officedocument."
                               "presentationml.slideLayout+xml",
                "slide": "application/vnd.openxmlformats-officedocument."
                         "presentationml.slide+xml",
            }
            z.writestr("[Content_Types].xml",
                f'<?xml version="1.0"?><Types xmlns="{CT}">'
                '<Default Extension="rels" ContentType="application/vnd.'
                'openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/ppt/presentation.xml" ContentType='
                '"application/vnd.openxmlformats-officedocument.presentationml.'
                'presentation.main+xml"/>'
                + "".join(
                    f'<Override PartName="{part}" ContentType="{types[kind]}"/>'
                    for part, kind in overrides
                )
                + "</Types>")
        return path


def _native_structure(deck: Deck) -> dict:
    """The minimal native_structure fields the context builder joins against."""
    return {
        "source": {"name": "test", "templateFile": "t.pptx", "sha256": "0" * 64},
        "masters": [
            {"key": f"master_{i + 1:02d}",
             "packagePart": f"ppt/slideMasters/slideMaster{i + 1}.xml"}
            for i in range(len(deck.masters))
        ],
        "layouts": [
            {"key": f"layout_{j + 1:02d}",
             "packagePart": f"ppt/slideLayouts/slideLayout{j + 1}.xml",
             "usedBySlides": [
                 k + 1 for k, (li, _) in enumerate(deck.slides) if li == j
             ]}
            for j in range(len(deck.layouts))
        ],
    }


def _build(deck: Deck, tmp: str) -> dict:
    path = deck.write(Path(tmp) / "deck.pptx")
    return build_template_context(path, _native_structure(deck))


def _by_role(layout: dict, role: str) -> dict:
    return next(p for p in layout["placeholders"] if p["semantic_role"] == role)


BODY_LST = (
    "<a:lstStyle>"
    "<a:lvl1pPr marL='0' indent='0'><a:buNone/>"
    "<a:defRPr sz='2000'/></a:lvl1pPr>"
    "<a:lvl2pPr marL='342900' indent='-342900'>"
    "<a:buFont typeface='Arial'/><a:buChar char='•'/>"
    "<a:defRPr sz='1600'/></a:lvl2pPr>"
    "<a:lvl3pPr marL='685800' indent='-342900'>"
    "<a:buAutoNum type='arabicPeriod' startAt='3'/>"
    "<a:defRPr sz='1400'/></a:lvl3pPr>"
    "</a:lstStyle>"
)


def _simple_deck() -> Deck:
    deck = Deck()
    master = deck.add_master(_part("sldMaster", _sp(
        ph="<p:ph type='body' idx='1'/>", name="Master Body",
        x_in=1, y_in=1, w_in=4, h_in=3,
        body_pr="<a:bodyPr anchor='ctr' lIns='45720' tIns='0' "
                "rIns='45720' bIns='0'/>",
    ), name="M"))
    layout = deck.add_layout(master, _part("sldLayout", "".join([
        _sp(ph="<p:ph type='title'/>", name="Title 1",
            x_in=0.5, y_in=0.4, w_in=9.0, h_in=1.0,
            lst_style="<a:lstStyle><a:lvl1pPr><a:defRPr sz='3200'/>"
                      "</a:lvl1pPr></a:lstStyle>"),
        _sp(ph="<p:ph type='body' idx='1'/>", name="Body 1",
            x_in=0.5, y_in=1.8, w_in=5.0, h_in=4.0, lst_style=BODY_LST),
        _pic(ph="<p:ph type='pic' idx='2'/>", name="Picture 1",
             x_in=6.0, y_in=1.8, w_in=4.0, h_in=4.0),
        _sp(ph=None, name="Decoration", x_in=0, y_in=7.0,
            w_in=13.3, h_in=0.4, text="static footer chrome"),
    ]), name="Test Layout"))
    return deck, layout


class StructureTests(unittest.TestCase):
    def test_schema_and_layout_coverage_are_derived(self):
        deck, _ = _simple_deck()
        extra_master = deck.add_master(_part("sldMaster", "", name="M2"))
        deck.add_layout(extra_master, _part("sldLayout", "", name="Empty"))
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        self.assertEqual(ctx["schema"], SCHEMA)
        self.assertEqual(len(ctx["layouts"]), len(deck.layouts))
        self.assertEqual(len(ctx["masters"]), len(deck.masters))

    def test_placeholder_roles_and_geometry(self):
        deck, _ = _simple_deck()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        layout = ctx["layouts"][0]
        roles = {p["semantic_role"] for p in layout["placeholders"]}
        self.assertEqual(roles, {"title", "body", "picture"})

        title = _by_role(layout, "title")
        self.assertEqual(title["geometry_in"]["width"], 9.0)
        self.assertEqual(title["geometry_in"]["height"], 1.0)
        self.assertEqual(title["levels"][0]["font"]["size_pt"], 32.0)

    def test_picture_placeholder_is_a_pic_host(self):
        deck, _ = _simple_deck()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        picture = _by_role(ctx["layouts"][0], "picture")
        self.assertEqual(picture["host"], "pic")
        self.assertTrue(picture["is_placeholder"])

    def test_ordinary_text_box_is_not_a_placeholder(self):
        deck, _ = _simple_deck()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        layout = ctx["layouts"][0]
        self.assertTrue(
            all(p["is_placeholder"] for p in layout["placeholders"])
        )
        static = layout["static_shapes"]
        self.assertEqual([s["name"] for s in static], ["Decoration"])
        self.assertFalse(static[0]["is_placeholder"])
        self.assertTrue(static[0]["has_text"])

    def test_layout_with_no_slides_still_emits_structure(self):
        deck, _ = _simple_deck()
        lonely = deck.add_layout(0, _part("sldLayout", _sp(
            ph="<p:ph type='title'/>", name="T", x_in=1, y_in=1, w_in=3, h_in=1,
        ), name="Unused"))
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        entry = ctx["layouts"][lonely]
        self.assertEqual(entry["exemplars"], [])
        self.assertEqual(entry["used_by_slides"], [])
        self.assertTrue(entry["placeholders"])


class ListStructureTests(unittest.TestCase):
    def _levels(self):
        deck, _ = _simple_deck()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        return _by_role(ctx["layouts"][0], "body")["levels"]

    def test_bullet_kinds_are_distinguished_not_flattened(self):
        levels = self._levels()
        kinds = [level["bullet"]["kind"] for level in levels[:3]]
        self.assertEqual(kinds, ["none", "character", "auto_number"])

    def test_character_bullet_keeps_glyph_and_font(self):
        bullet = self._levels()[1]["bullet"]
        self.assertEqual(bullet["char"], "•")
        self.assertEqual(bullet["font"], "Arial")

    def test_auto_number_keeps_scheme_and_start(self):
        bullet = self._levels()[2]["bullet"]
        self.assertEqual(bullet["scheme"], "arabicPeriod")
        self.assertEqual(bullet["start_at"], 3)

    def test_indentation_is_reported_per_level(self):
        levels = self._levels()
        self.assertEqual(levels[0]["left_margin_in"], 0.0)
        self.assertGreater(levels[1]["left_margin_in"], 0.0)
        self.assertLess(levels[1]["indent_in"], 0.0)

    def test_levels_defined_counts_declared_levels(self):
        deck, _ = _simple_deck()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        body = _by_role(ctx["layouts"][0], "body")
        self.assertEqual(body["levels_defined"], len(body["levels"]))
        self.assertGreaterEqual(body["levels_defined"], 3)


class InheritanceTests(unittest.TestCase):
    def test_master_body_properties_reach_the_layout(self):
        deck, _ = _simple_deck()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        body = _by_role(ctx["layouts"][0], "body")["text_behavior"]
        self.assertEqual(body["anchor"], "ctr")
        self.assertEqual(body["provenance"]["anchor"], "master")
        self.assertEqual(body["insets_in"]["top"], 0.0)
        self.assertEqual(body["provenance"]["insets"], "master")

    def test_unstated_values_are_marked_default(self):
        deck, _ = _simple_deck()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        body = _by_role(ctx["layouts"][0], "body")["text_behavior"]
        self.assertEqual(body["wrap"], "square")
        self.assertEqual(body["provenance"]["wrap"], "default")

    def test_theme_font_tokens_are_resolved(self):
        deck = Deck()
        master = deck.add_master(_part("sldMaster", "", name="M"))
        deck.add_layout(master, _part("sldLayout", _sp(
            ph="<p:ph type='title'/>", name="T", x_in=1, y_in=1, w_in=5, h_in=1,
            lst_style="<a:lstStyle><a:lvl1pPr><a:defRPr sz='2400'>"
                      "<a:latin typeface='+mj-lt'/></a:defRPr>"
                      "</a:lvl1pPr></a:lstStyle>",
        ), name="L"))
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        family = ctx["layouts"][0]["placeholders"][0]["levels"][0]["font"]["family"]
        self.assertEqual(family, "Major Test Face")
        self.assertNotIn("+", family)

    def test_a_level_without_a_typeface_falls_back_to_the_theme(self):
        """Only the size is stated here; the face must still resolve."""
        deck = Deck()
        master = deck.add_master(_part("sldMaster", "", name="M"))
        deck.add_layout(master, _part("sldLayout", _sp(
            ph="<p:ph type='body' idx='1'/>", name="B",
            x_in=1, y_in=1, w_in=4, h_in=3,
            lst_style="<a:lstStyle><a:lvl1pPr><a:defRPr sz='1800'/>"
                      "</a:lvl1pPr></a:lstStyle>",
        ), name="L"))
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        level = ctx["layouts"][0]["placeholders"][0]["levels"][0]
        self.assertEqual(level["font"]["family"], "Minor Test Face")
        self.assertEqual(level["font"]["size_pt"], 18.0)

    def test_a_typeface_set_above_the_level_is_inherited(self):
        """Size on lvl2, face only on lvl1: the face must still be found."""
        deck = Deck()
        master = deck.add_master(_part("sldMaster", _sp(
            ph="<p:ph type='body' idx='1'/>", name="MB",
            x_in=1, y_in=1, w_in=4, h_in=3,
            lst_style="<a:lstStyle><a:lvl2pPr><a:defRPr>"
                      "<a:latin typeface='Inherited Face'/></a:defRPr>"
                      "</a:lvl2pPr></a:lstStyle>",
        ), name="M"))
        deck.add_layout(master, _part("sldLayout", _sp(
            ph="<p:ph type='body' idx='1'/>", name="B",
            x_in=1, y_in=1, w_in=4, h_in=3,
            lst_style="<a:lstStyle><a:lvl2pPr><a:defRPr sz='1100'/>"
                      "</a:lvl2pPr></a:lstStyle>",
        ), name="L"))
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        level = next(
            v for v in ctx["layouts"][0]["placeholders"][0]["levels"]
            if v["level"] == 1
        )
        self.assertEqual(level["font"]["family"], "Inherited Face")
        self.assertEqual(level["font"]["size_pt"], 11.0)


class TextBehaviorTests(unittest.TestCase):
    def _behavior(self, body_pr: str) -> dict:
        deck = Deck()
        master = deck.add_master(_part("sldMaster", "", name="M"))
        deck.add_layout(master, _part("sldLayout", _sp(
            ph="<p:ph type='body' idx='1'/>", name="B",
            x_in=1, y_in=1, w_in=4, h_in=3, body_pr=body_pr,
            lst_style="<a:lstStyle><a:lvl1pPr><a:defRPr sz='1800'/>"
                      "</a:lvl1pPr></a:lstStyle>",
        ), name="L"))
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        return ctx["layouts"][0]["placeholders"][0]

    def test_insets_are_reported_in_inches(self):
        record = self._behavior(
            "<a:bodyPr lIns='914400' tIns='457200' rIns='0' bIns='0'/>"
        )
        insets = record["text_behavior"]["insets_in"]
        self.assertEqual(insets["left"], 1.0)
        self.assertEqual(insets["top"], 0.5)
        self.assertEqual(insets["right"], 0.0)

    def test_wrap_none_makes_chars_per_line_meaningless(self):
        record = self._behavior("<a:bodyPr wrap='none'/>")
        self.assertEqual(record["text_behavior"]["wrap"], "none")
        self.assertIsNone(record["levels"][0]["capacity"]["chars_per_line"])
        self.assertIsNotNone(record["levels"][0]["capacity"]["lines"])

    def test_norm_autofit_scale_and_reduction_are_reported(self):
        record = self._behavior(
            "<a:bodyPr><a:normAutofit fontScale='62500' "
            "lnSpcReduction='10000'/></a:bodyPr>"
        )
        autofit = record["text_behavior"]["autofit"]
        self.assertEqual(autofit["mode"], "normAutofit")
        self.assertEqual(autofit["font_scale_pct"], 62.5)
        self.assertEqual(autofit["line_space_reduction_pct"], 10.0)

    def test_shape_autofit_is_reported(self):
        record = self._behavior("<a:bodyPr><a:spAutoFit/></a:bodyPr>")
        self.assertEqual(record["text_behavior"]["autofit"]["mode"], "spAutoFit")

    def test_absent_autofit_is_marked_as_default(self):
        record = self._behavior("<a:bodyPr/>")
        autofit = record["text_behavior"]["autofit"]
        self.assertEqual(autofit["mode"], "noAutofit")
        self.assertEqual(autofit["source"], "default")

    def test_vertical_text_is_reported(self):
        record = self._behavior("<a:bodyPr vert='eaVert'/>")
        self.assertEqual(record["text_behavior"]["vertical"], "eaVert")


class CapacityTests(unittest.TestCase):
    def _capacity(self, *, w_in: float, h_in: float, sz: int) -> dict:
        deck = Deck()
        master = deck.add_master(_part("sldMaster", "", name="M"))
        deck.add_layout(master, _part("sldLayout", _sp(
            ph="<p:ph type='body' idx='1'/>", name="B",
            x_in=0.5, y_in=0.5, w_in=w_in, h_in=h_in,
            body_pr="<a:bodyPr lIns='0' tIns='0' rIns='0' bIns='0'/>",
            lst_style=f"<a:lstStyle><a:lvl1pPr><a:defRPr sz='{sz}'/>"
                      "</a:lvl1pPr></a:lstStyle>",
        ), name="L"))
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        return ctx["layouts"][0]["placeholders"][0]["levels"][0]["capacity"]

    def test_capacity_is_never_claimed_exact(self):
        self.assertFalse(self._capacity(w_in=4, h_in=3, sz=1800)["exact"])

    def test_a_taller_box_holds_more_lines(self):
        short = self._capacity(w_in=4, h_in=1, sz=1800)["lines"]
        tall = self._capacity(w_in=4, h_in=4, sz=1800)["lines"]
        self.assertGreater(tall, short)

    def test_a_wider_box_holds_more_characters(self):
        narrow = self._capacity(w_in=2, h_in=3, sz=1800)["chars_per_line"]
        wide = self._capacity(w_in=8, h_in=3, sz=1800)["chars_per_line"]
        self.assertGreater(wide, narrow)

    def test_smaller_type_holds_more_of_both(self):
        large = self._capacity(w_in=4, h_in=3, sz=3600)
        small = self._capacity(w_in=4, h_in=3, sz=1200)
        self.assertGreater(small["chars_per_line"], large["chars_per_line"])
        self.assertGreater(small["lines"], large["lines"])

    def test_capacity_basis_is_declared_and_not_exact(self):
        deck, _ = _simple_deck()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        self.assertFalse(ctx["capacity_basis"]["exact"])
        self.assertIn("estimate_text_width", ctx["capacity_basis"]["estimator"])


class ExemplarTests(unittest.TestCase):
    def _deck_with_slide(self) -> Deck:
        deck, layout = _simple_deck()
        deck.add_slide(layout, _part("sld", "".join([
            _sp(ph="<p:ph type='title'/>", name="T",
                x_in=0.5, y_in=0.4, w_in=9.0, h_in=1.0, text="Real Headline"),
            _sp(ph="<p:ph type='body' idx='1'/>", name="B",
                x_in=0.5, y_in=1.8, w_in=5.0, h_in=4.0,
                paragraphs="<a:p><a:r><a:t>Lead line</a:t></a:r></a:p>"
                           "<a:p><a:pPr lvl='1'/><a:r><a:t>Nested point</a:t>"
                           "</a:r></a:p>"),
        ]), name="S"))
        return deck

    def test_exemplar_text_comes_from_the_slide(self):
        deck = self._deck_with_slide()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        exemplars = ctx["layouts"][0]["exemplars"]
        self.assertEqual(len(exemplars), 1)
        by_idx = exemplars[0]["by_idx"]
        self.assertEqual(by_idx["title"]["paragraphs"][0]["text"], "Real Headline")

    def test_exemplar_preserves_paragraph_levels(self):
        deck = self._deck_with_slide()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        paragraphs = ctx["layouts"][0]["exemplars"][0]["by_idx"]["1"]["paragraphs"]
        self.assertEqual([p["level"] for p in paragraphs], [0, 1])
        self.assertEqual(paragraphs[1]["text"], "Nested point")

    def test_demonstrated_levels_come_from_exemplars_only(self):
        deck = self._deck_with_slide()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        body = _by_role(ctx["layouts"][0], "body")
        self.assertEqual(body["levels_demonstrated"], [0, 1])
        # The layout declares more levels than any slide demonstrates.
        self.assertGreater(body["levels_defined"], len(body["levels_demonstrated"]))

    def test_used_by_slide_count_is_recorded(self):
        deck = self._deck_with_slide()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        self.assertEqual(ctx["layouts"][0]["used_by_slide_count"], 1)


class ProvenanceTests(unittest.TestCase):
    def test_no_semantic_inference_is_emitted(self):
        deck, _ = _simple_deck()
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        self.assertFalse(ctx["semantic_inference"]["present"])
        blob = json.dumps(ctx)
        for invented in ("caption", "quote", "eyebrow", "primary_content"):
            self.assertNotIn(f'"{invented}"', blob)

    def test_written_artifact_is_valid_json_at_the_expected_path(self):
        deck, _ = _simple_deck()
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            pptx = deck.write(root / "deck.pptx")
            write_template_context(pptx, root, _native_structure(deck))
            written = root / "analysis" / "template_context.json"
            self.assertTrue(written.is_file())
            self.assertEqual(json.loads(written.read_text())["schema"], SCHEMA)


class ArbitraryShapeTests(unittest.TestCase):
    """Counts must follow the input, never a constant."""

    def _deck(self, masters: int, layouts_each: int, slides: int) -> Deck:
        deck = Deck(width_in=9.0, height_in=6.75)
        for m in range(masters):
            master = deck.add_master(_part("sldMaster", "", name=f"M{m}"))
            for n in range(layouts_each):
                deck.add_layout(master, _part("sldLayout", _sp(
                    ph="<p:ph type='body' idx='1'/>", name=f"B{m}{n}",
                    x_in=1, y_in=1, w_in=3, h_in=2,
                    lst_style="<a:lstStyle><a:lvl1pPr><a:defRPr sz='1400'/>"
                              "</a:lvl1pPr></a:lstStyle>",
                ), name=f"L{m}{n}"))
        for s in range(slides):
            deck.add_slide(s % len(deck.layouts), _part("sld", _sp(
                ph="<p:ph type='body' idx='1'/>", name="B",
                x_in=1, y_in=1, w_in=3, h_in=2, text=f"slide {s}",
            ), name=f"S{s}"))
        return deck

    def test_counts_follow_the_input_deck(self):
        for masters, layouts_each, slides in ((1, 3, 2), (3, 2, 5), (2, 1, 0)):
            with self.subTest(masters=masters, layouts=layouts_each):
                deck = self._deck(masters, layouts_each, slides)
                with TemporaryDirectory() as tmp:
                    ctx = _build(deck, tmp)
                self.assertEqual(len(ctx["masters"]), masters)
                self.assertEqual(len(ctx["layouts"]), masters * layouts_each)
                self.assertEqual(
                    sum(len(item["placeholders"]) for item in ctx["layouts"]),
                    masters * layouts_each,
                )

    def test_every_layout_key_joins_to_native_structure(self):
        deck = self._deck(2, 2, 3)
        native = _native_structure(deck)
        with TemporaryDirectory() as tmp:
            path = deck.write(Path(tmp) / "deck.pptx")
            ctx = build_template_context(path, native)
        known = {entry["key"] for entry in native["layouts"]}
        self.assertEqual({item["key"] for item in ctx["layouts"]}, known)
        masters = {entry["key"] for entry in native["masters"]}
        self.assertTrue(
            {item["master_key"] for item in ctx["layouts"]}.issubset(masters)
        )

    def test_slide_size_follows_the_deck(self):
        deck = self._deck(1, 1, 0)
        with TemporaryDirectory() as tmp:
            ctx = _build(deck, tmp)
        self.assertEqual(ctx["slide_size_in"]["width"], 9.0)
        self.assertEqual(ctx["slide_size_in"]["height"], 6.75)


if __name__ == "__main__":
    unittest.main()
