"""Generator-facing template context extracted from a source PPTX.

This module is a **serializer**, not a parser. Every structural and typographic
fact is resolved by the existing import pipeline — ``walk_sp_tree`` for shapes and
master/layout inheritance, ``txbody_to_svg`` for list styles and text metrics — and
this module only shapes those results into JSON that an external generator can read.

The artifact is derived context. The source PPTX remains the source of truth.

Facts are labelled by how they were obtained:

- **extracted** — read from the PPTX with inheritance applied
- **derived** — computed from extracted values (capacity, unit conversions)
- **exemplar** — observed in a real slide that uses the layout

Semantic purpose ("caption", "quote", "primary content") is deliberately absent.
PowerPoint does not encode it, so inventing it here would present a guess as a fact.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from pptx_to_svg import txbody_to_svg as T
from pptx_to_svg.ooxml_loader import OoxmlPackage
from pptx_to_svg.shape_walker import walk_sp_tree
from svg_to_pptx.drawingml.utils import estimate_text_width
from template_import.manifest import placeholder_semantic_role

SCHEMA = "ppt-master.template-context.v1"
CONTEXT_NAME = "analysis/template_context.json"

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A_NS, "p": P_NS}

PX_PER_INCH = 96.0
PT_PER_PX = 0.75
EMU_PER_PX = 9525.0

# PowerPoint's own defaults when no part in the chain declares an inset.
DEFAULT_INSETS_EMU = {"left": 91440, "top": 45720, "right": 91440, "bottom": 45720}
_INSET_ATTR = {"left": "lIns", "top": "tIns", "right": "rIns", "bottom": "bIns"}

_AUTOFIT_TAGS = ("noAutofit", "normAutofit", "spAutoFit")
_MAX_LIST_LEVELS = 9

# Mixed-case English used to derive an average advance width. Any fixed sample
# works; this one is declared once so the estimate is reproducible.
REFERENCE_SAMPLE = (
    "The quick brown fox jumps over the lazy dog and then continues "
    "running past the riverbank in the morning light"
)

# Exemplars are evidence, not a catalogue. A handful shows how a layout is used.
MAX_EXEMPLARS_PER_LAYOUT = 3


def _in(px: float) -> float:
    return round(px / PX_PER_INCH, 3)


def _pt(px: float) -> float:
    return round(px * PT_PER_PX, 1)


def _local(tag: object) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _theme_fonts(package: OoxmlPackage) -> dict[str, str]:
    """Return the deck's major/minor Latin faces for +mj-lt / +mn-lt tokens."""
    for master in package.iter_all_masters():
        theme = package.resolve_theme(master)
        if theme is None:
            continue
        scheme = theme.xml.find(".//a:fontScheme", NS)
        if scheme is None:
            continue
        fonts = {}
        for tag, key in (("majorFont", "majorLatin"), ("minorFont", "minorLatin")):
            latin = scheme.find(f"a:{tag}/a:latin", NS)
            if latin is not None and latin.get("typeface"):
                fonts[key] = latin.get("typeface")
        if fonts:
            return fonts
    return {}


def _body_property_chain(node) -> list[tuple[str, ET.Element]]:
    """Return this placeholder's bodyPr levels, nearest first.

    The node's own bodyPr belongs to the part being walked; ``walk_sp_tree``
    already labels every inherited level with the part that declared it.
    """
    chain: list[tuple[str, ET.Element]] = []
    own = node.xml.find("p:txBody/a:bodyPr", NS)
    if own is not None:
        chain.append(("layout", own))
    for scoped in node.inherited_body_properties:
        chain.append((scoped.scope, scoped.body_pr))
    return chain


def _resolved_attr(
    chain: list[tuple[str, ET.Element]],
    name: str,
) -> tuple[str | None, str]:
    """Return (value, the scope that declared it) or (None, "default")."""
    for scope, body_pr in chain:
        if name in body_pr.attrib:
            return body_pr.attrib[name], scope
    return None, "default"


def _resolved_autofit(chain: list[tuple[str, ET.Element]]) -> dict[str, Any]:
    for scope, body_pr in chain:
        for child in body_pr:
            if _local(child.tag) not in _AUTOFIT_TAGS:
                continue
            scale = child.get("fontScale")
            reduction = child.get("lnSpcReduction")
            return {
                "mode": _local(child.tag),
                # OOXML stores these as thousandths of a percent.
                "font_scale_pct": round(int(scale) / 1000, 2) if scale else None,
                "line_space_reduction_pct": (
                    round(int(reduction) / 1000, 2) if reduction else None
                ),
                "source": scope,
            }
    return {
        "mode": "noAutofit",
        "font_scale_pct": None,
        "line_space_reduction_pct": None,
        "source": "default",
    }


def _text_behavior(node) -> dict[str, Any]:
    chain = _body_property_chain(node)
    wrap, wrap_src = _resolved_attr(chain, "wrap")
    anchor, anchor_src = _resolved_attr(chain, "anchor")
    vertical, vertical_src = _resolved_attr(chain, "vert")
    anchor_ctr, _ = _resolved_attr(chain, "anchorCtr")
    columns, _ = _resolved_attr(chain, "numCol")
    column_spacing, _ = _resolved_attr(chain, "spcCol")

    insets: dict[str, float] = {}
    inset_source = "default"
    for name, attr in _INSET_ATTR.items():
        raw, source = _resolved_attr(chain, attr)
        emu = float(raw) if raw is not None else DEFAULT_INSETS_EMU[name]
        insets[name] = emu / EMU_PER_PX
        if source != "default":
            inset_source = source

    return {
        "wrap": wrap or "square",
        "anchor": anchor or "t",
        "anchor_center": anchor_ctr == "1",
        "vertical": vertical or "horz",
        "columns": {
            "count": int(columns) if columns else 1,
            "spacing_in": _in(float(column_spacing) / EMU_PER_PX)
            if column_spacing
            else 0.0,
        },
        "autofit": _resolved_autofit(chain),
        "insets_in": {name: _in(px) for name, px in insets.items()},
        "provenance": {
            "wrap": wrap_src,
            "anchor": anchor_src,
            "vertical": vertical_src,
            "insets": inset_source,
        },
    }, insets


def _bullet(level_prs: tuple[ET.Element | None, ...], level: int) -> dict[str, Any]:
    """Return the level's list-marker structure, not just whether it has one."""
    if T._child_chain(level_prs, "a:buNone") is not None:
        return {"kind": "none"}
    bu_char = T._child_chain(level_prs, "a:buChar")
    if bu_char is not None:
        bu_font = T._child_chain(level_prs, "a:buFont")
        return {
            "kind": "character",
            "char": bu_char.get("char"),
            "font": bu_font.get("typeface") if bu_font is not None else None,
        }
    bu_auto = T._child_chain(level_prs, "a:buAutoNum")
    if bu_auto is not None:
        try:
            start_at = int(bu_auto.get("startAt", "1"))
        except ValueError:
            start_at = 1
        return {
            "kind": "auto_number",
            "scheme": bu_auto.get("type", "arabicPeriod"),
            "start_at": start_at,
        }
    return {"kind": "none"}


def _line_spacing(level_prs: tuple[ET.Element | None, ...]) -> dict[str, Any]:
    """Return proportional spacing, plus an exact point value when one is set."""
    exact_pt = None
    ln_spc = T._child_chain(level_prs, "a:lnSpc")
    if ln_spc is not None:
        spc_pts = ln_spc.find("a:spcPts", NS)
        if spc_pts is not None:
            try:
                exact_pt = round(int(spc_pts.get("val", "0")) / 100, 1)
            except ValueError:
                exact_pt = None
    return {"ratio": round(T._line_height_ratio(level_prs), 3), "exact_pt": exact_pt}


def _default_run_properties(
    level_prs: tuple[ET.Element | None, ...],
) -> tuple[ET.Element, ...]:
    """Return every defRPr in the level chain, nearest first.

    A level may set its size but leave the typeface to the level above it, so
    resolution has to see the whole chain rather than only the nearest entry.
    """
    found = []
    for level_pr in level_prs:
        if level_pr is None:
            continue
        def_rpr = level_pr.find("a:defRPr", NS)
        if def_rpr is not None:
            found.append(def_rpr)
    return tuple(found)


def _capacity(
    *,
    font_px: float,
    line_ratio: float,
    width_px: float,
    height_px: float,
    insets: dict[str, float],
    margin_left_px: float,
    indent_px: float,
    bullet_prefix: str,
    wrap: str,
) -> dict[str, Any]:
    """Approximate how much text fits, using the importer's own estimator.

    Deliberately not an exact character limit: advance widths are estimated
    rather than measured from font files, so this is a physical constraint to
    reason with, not a rule to enforce.
    """
    usable_w = (
        width_px
        - insets["left"]
        - insets["right"]
        - margin_left_px
        - max(-indent_px, 0.0)
        - (estimate_text_width(bullet_prefix, font_px) if bullet_prefix else 0.0)
    )
    usable_h = height_px - insets["top"] - insets["bottom"]
    line_h = font_px * line_ratio

    chars_per_line = None
    if wrap != "none":
        advance = estimate_text_width(REFERENCE_SAMPLE, font_px) / len(REFERENCE_SAMPLE)
        chars_per_line = max(int(usable_w / advance), 0) if advance > 0 else None

    return {
        "chars_per_line": chars_per_line,
        "lines": max(int(usable_h / line_h), 0) if line_h > 0 else None,
        "exact": False,
    }


def _levels(
    node,
    *,
    theme_fonts: dict[str, str],
    insets: dict[str, float],
    wrap: str,
) -> list[dict[str, Any]]:
    own_lst = node.xml.find("p:txBody/a:lstStyle", NS)
    chain = ((own_lst,) if own_lst is not None else ()) + node.inherited_lst_styles

    levels: list[dict[str, Any]] = []
    for level in range(_MAX_LIST_LEVELS):
        level_prs = T._lst_style_level_prs(chain, level)
        if not level_prs:
            continue
        def_rprs = _default_run_properties(level_prs)
        def_rpr = def_rprs[0] if def_rprs else None
        font_px = T._font_size_px(def_rprs, T.DEFAULT_FONT_SIZE_PX)
        family = T._resolve_theme_typeface(
            T._typeface_chain(def_rprs, "latin"), theme_fonts
        )
        if family is None:
            # PowerPoint falls back to the theme's minor Latin face for body
            # text when nothing in the chain names one.
            family = theme_fonts.get("minorLatin")
        spacing = _line_spacing(level_prs)
        margin_left_px = T._emu_px_attr_chain(level_prs, "marL", 0.0)
        indent_px = T._emu_px_attr_chain(level_prs, "indent", 0.0)
        bullet = _bullet(level_prs, level)
        # autonum_state is per-call: this renders one representative marker,
        # never a running count across the document.
        prefix = T._resolve_bullet_prefix(level_prs, level, {})

        levels.append({
            "level": level,
            "font": {
                "family": family,
                "size_pt": _pt(font_px),
                "bold": (def_rpr.get("b") == "1") if def_rpr is not None else False,
                "italic": (def_rpr.get("i") == "1") if def_rpr is not None else False,
            },
            "line_spacing": spacing,
            "space_before_pt": _pt(
                T._spacing_points_px(level_prs, "a:spcBef/a:spcPts")
            ),
            "space_after_pt": _pt(T._spacing_points_px(level_prs, "a:spcAft/a:spcPts")),
            "alignment": T._attr_chain(level_prs, "algn") or "l",
            "left_margin_in": _in(margin_left_px),
            "indent_in": _in(indent_px),
            "bullet": bullet,
            "rendered_prefix": prefix or None,
            "capacity": _capacity(
                font_px=font_px,
                line_ratio=spacing["ratio"],
                width_px=node.xfrm.w,
                height_px=node.xfrm.h,
                insets=insets,
                margin_left_px=margin_left_px,
                indent_px=indent_px,
                bullet_prefix=prefix,
                wrap=wrap,
            ),
        })
    return levels


def _geometry(node) -> dict[str, Any]:
    return {
        "x": _in(node.xfrm.x),
        "y": _in(node.xfrm.y),
        "width": _in(node.xfrm.w),
        "height": _in(node.xfrm.h),
        "rotation_deg": round(node.xfrm.rot, 2),
        "flip_horizontal": bool(node.xfrm.flip_h),
        "flip_vertical": bool(node.xfrm.flip_v),
    }


def _shape_text(element: ET.Element) -> str:
    return " ".join(
        node.text or "" for node in element.iter(f"{{{A_NS}}}t")
    ).strip()


def _paragraphs(element: ET.Element) -> list[dict[str, Any]]:
    """Return a placeholder's paragraphs with their outline levels."""
    paragraphs: list[dict[str, Any]] = []
    for para in element.findall("p:txBody/a:p", NS):
        text = "".join(
            node.text or "" for node in para.iter(f"{{{A_NS}}}t")
        ).strip()
        if not text:
            continue
        p_pr = para.find("a:pPr", NS)
        try:
            level = int(p_pr.get("lvl", "0")) if p_pr is not None else 0
        except ValueError:
            level = 0
        paragraphs.append({"level": level, "text": text})
    return paragraphs


def _placeholder_key(placeholder) -> str:
    return str(placeholder.idx) if placeholder.idx is not None else placeholder.type


def _exemplars(
    package: OoxmlPackage,
    layout_part: str,
    layout_xml: ET.Element,
    master_xml: ET.Element | None,
) -> tuple[list[dict[str, Any]], int, dict[str, set[int]]]:
    """Collect real content from slides using this layout.

    Evidence of how the template is actually filled. PowerPoint records no
    semantic purpose, so this is the only signal distinguishing, say, a one-line
    label from a full content column beyond geometry.
    """
    exemplars: list[dict[str, Any]] = []
    demonstrated: dict[str, set[int]] = {}
    total = 0
    for slide in package.iter_slides():
        if slide.layout is None or slide.layout.path != layout_part:
            continue
        total += 1
        by_idx: dict[str, Any] = {}
        for node in walk_sp_tree(
            slide.part.xml, layout_xml=layout_xml, master_xml=master_xml
        ):
            if node.placeholder is None:
                continue
            paragraphs = _paragraphs(node.xml)
            if not paragraphs:
                continue
            key = _placeholder_key(node.placeholder)
            by_idx[key] = {
                "paragraphs": paragraphs,
                "char_count": sum(len(p["text"]) for p in paragraphs),
            }
            demonstrated.setdefault(key, set()).update(
                p["level"] for p in paragraphs
            )
        if by_idx and len(exemplars) < MAX_EXEMPLARS_PER_LAYOUT:
            exemplars.append({"slide": slide.index, "by_idx": by_idx})
    return exemplars, total, demonstrated


def build_template_context(
    source_pptx: Path,
    native_structure: dict[str, Any],
) -> dict[str, Any]:
    """Build the generator-facing context for every layout in the package.

    Layout and master keys are taken from *native_structure* so the two
    artifacts join cleanly rather than deriving keys twice.
    """
    layout_key_by_part = {
        entry["packagePart"]: entry["key"]
        for entry in native_structure.get("layouts", [])
    }
    master_key_by_part = {
        entry["packagePart"]: entry["key"]
        for entry in native_structure.get("masters", [])
    }
    used_by_slides = {
        entry["packagePart"]: entry.get("usedBySlides", [])
        for entry in native_structure.get("layouts", [])
    }

    package = OoxmlPackage(source_pptx)
    package.open()
    try:
        theme_fonts = _theme_fonts(package)
        width_px, height_px = package.slide_size_px

        masters = [
            {
                "key": master_key_by_part.get(master.path, master.path),
                "name": (
                    master.xml.find("p:cSld", NS).get("name")
                    if master.xml.find("p:cSld", NS) is not None
                    else None
                ),
                "package_part": master.path,
            }
            for master in package.iter_all_masters()
        ]

        layouts = []
        for layout, master in package.iter_all_layouts_with_parent():
            c_sld = layout.xml.find("p:cSld", NS)
            exemplars, slide_count, demonstrated = _exemplars(
                package, layout.path, layout.xml, master.xml if master else None
            )

            placeholders = []
            static_shapes = []
            for node in walk_sp_tree(
                layout.xml, master_xml=master.xml if master else None
            ):
                if node.placeholder is None:
                    text = _shape_text(node.xml)
                    static_shapes.append({
                        "kind": node.kind,
                        "name": node.name or None,
                        "is_placeholder": False,
                        "geometry_in": _geometry(node),
                        "has_text": bool(text),
                        "text_sample": text[:200] or None,
                    })
                    continue

                behavior, insets = _text_behavior(node)
                levels = _levels(
                    node,
                    theme_fonts=theme_fonts,
                    insets=insets,
                    wrap=behavior["wrap"],
                )
                key = _placeholder_key(node.placeholder)
                placeholders.append({
                    "idx": node.placeholder.idx,
                    "placeholder_type": node.placeholder.type,
                    "semantic_role": placeholder_semantic_role(
                        node.placeholder.type or "obj"
                    ),
                    "size_hint": node.placeholder.sz,
                    "orientation": node.placeholder.orient,
                    "host": node.kind,
                    "is_placeholder": True,
                    "shape_name": node.name or None,
                    "geometry_in": _geometry(node),
                    "text_behavior": behavior,
                    "levels_defined": len(levels),
                    "levels_demonstrated": sorted(demonstrated.get(key, ())),
                    "levels": levels,
                })

            layouts.append({
                "key": layout_key_by_part.get(layout.path, layout.path),
                "name": c_sld.get("name") if c_sld is not None else None,
                "package_part": layout.path,
                "master_key": (
                    master_key_by_part.get(master.path, master.path)
                    if master
                    else None
                ),
                "used_by_slides": used_by_slides.get(layout.path, []),
                "used_by_slide_count": slide_count,
                "placeholders": placeholders,
                "static_shapes": static_shapes,
                "exemplars": exemplars,
            })
    finally:
        package.close()

    return {
        "schema": SCHEMA,
        "source": native_structure.get("source", {}),
        "slide_size_in": {"width": _in(width_px), "height": _in(height_px)},
        "theme": {
            "major_latin": theme_fonts.get("majorLatin"),
            "minor_latin": theme_fonts.get("minorLatin"),
        },
        "capacity_basis": {
            "method": "reference-sample",
            "estimator": "svg_to_pptx.drawingml.utils.estimate_text_width",
            "exact": False,
            "note": (
                "Approximate. Advance widths are estimated by the same routine "
                "the importer wraps text with, not measured from font files. "
                "Use as a physical constraint, not a hard character limit."
            ),
        },
        "semantic_inference": {
            "present": False,
            "note": (
                "PowerPoint encodes placeholder type, geometry, typography and "
                "list structure, but not purpose. Labels such as caption, quote "
                "or primary content are not emitted; derive them from geometry, "
                "typography and exemplars."
            ),
        },
        "masters": masters,
        "layouts": layouts,
    }


def write_template_context(
    source_pptx: Path,
    output_dir: Path,
    native_structure: dict[str, Any],
) -> dict[str, Any]:
    """Write analysis/template_context.json and return the context."""
    context = build_template_context(source_pptx, native_structure)
    path = output_dir / CONTEXT_NAME
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(context, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return context
