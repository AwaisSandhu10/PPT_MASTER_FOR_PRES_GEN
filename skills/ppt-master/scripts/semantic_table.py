#!/usr/bin/env python3
"""Compact, expand, and project the canonical semantic-table.v2 payload."""

from __future__ import annotations

import copy
import json
import re
import sys
from collections.abc import Iterable
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


SEMANTIC_TABLE_SCHEMA = "ppt-master.semantic-table.v2"

_TOP_LEVEL_FIELDS = {
    "schema",
    "name",
    "x",
    "y",
    "width",
    "height",
    "strict_grid",
    "header_rows",
    "column_widths",
    "row_heights",
    "style",
    "defaults",
    "cell_styles",
    "columns",
    "rows",
}
_TABLE_STYLE_FIELDS = {
    "band_row",
    "font_family",
    "font_size",
    "header_font_size",
    "header_fill",
    "header_text",
    "body_fill",
    "body_text",
    "band_fill",
    "border_color",
    "border_width",
    "padding",
    "valign",
    "lang",
    "table_style_id",
}
_CELL_FORMAT_FIELDS = (
    "fill",
    "fill_opacity",
    "color",
    "font_size",
    "bold",
    "align",
    "valign",
    "borders",
    "padding",
    "padding_left",
    "padding_right",
    "padding_top",
    "padding_bottom",
    "border_color",
    "border_width",
    "lang",
    "anchor_center",
    "horizontal_overflow",
)
_CELL_FIELDS = set(_CELL_FORMAT_FIELDS) | {
    "text",
    "paragraphs",
    "row_span",
    "col_span",
    "merge_continuation",
    "cell_style",
}
_PARAGRAPH_DEFAULT_FIELDS = (
    "align",
    "line_spacing_percent",
)
_PARAGRAPH_FIELDS = set(_PARAGRAPH_DEFAULT_FIELDS) | {"text", "runs"}
_PLAIN_TEXT_RUN_FIELDS = ("bold", "color", "font_size")
_RUN_DEFAULT_FIELDS = (
    "bold",
    "italic",
    "underline",
    "strike",
    "color",
    "font_size",
    "font_family",
    "lang",
    "alt_lang",
    "baseline_percent",
    "outline",
)
_RUN_FIELDS = set(_RUN_DEFAULT_FIELDS) | {"text"}
_DEFAULT_SECTIONS = {
    "cell": set(_CELL_FORMAT_FIELDS),
    "paragraph": set(_PARAGRAPH_DEFAULT_FIELDS),
    "run": set(_RUN_DEFAULT_FIELDS),
}
_MERGED_OBJECT_FIELDS = {"padding"}
_STYLE_NAME_RE = re.compile(r"^[a-z][a-z0-9]*(?:-[a-z0-9]+)*$")


def _canonical_json(value: Any) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def _payload_size(payload: dict[str, Any], defaults: dict[str, Any]) -> int:
    return len(_canonical_json({"payload": payload, "defaults": defaults}))


def _require_object(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise RuntimeError(f"Native PPTX table {label} must be an object")
    return value


def _reject_unknown_fields(
    value: dict[str, Any],
    allowed: set[str],
    label: str,
) -> None:
    unknown = sorted(set(value) - allowed)
    if unknown:
        raise RuntimeError(
            f"Native PPTX table {label} contains unsupported field(s): "
            + ", ".join(unknown)
        )


def _format_layers(*layers: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for layer in layers:
        for key, value in layer.items():
            if (
                key in _MERGED_OBJECT_FIELDS
                and isinstance(value, dict)
                and isinstance(result.get(key), dict)
            ):
                merged = copy.deepcopy(result[key])
                merged.update(copy.deepcopy(value))
                result[key] = merged
            else:
                result[key] = copy.deepcopy(value)
    return result


def _validated_defaults(payload: dict[str, Any]) -> dict[str, dict[str, Any]]:
    raw_defaults = payload.get("defaults", {})
    defaults = _require_object(raw_defaults, "defaults")
    _reject_unknown_fields(defaults, set(_DEFAULT_SECTIONS), "defaults")
    result: dict[str, dict[str, Any]] = {}
    for section, allowed in _DEFAULT_SECTIONS.items():
        raw_section = defaults.get(section, {})
        section_data = _require_object(raw_section, f"defaults.{section}")
        _reject_unknown_fields(
            section_data,
            allowed,
            f"defaults.{section}",
        )
        result[section] = copy.deepcopy(section_data)
    return result


def _validated_cell_styles(payload: dict[str, Any]) -> dict[str, dict[str, Any]]:
    raw_styles = payload.get("cell_styles", {})
    styles = _require_object(raw_styles, "cell_styles")
    result: dict[str, dict[str, Any]] = {}
    for name, raw_style in styles.items():
        if not isinstance(name, str) or not _STYLE_NAME_RE.fullmatch(name):
            raise RuntimeError(
                "Native PPTX table cell style names must use lower-case kebab-case"
            )
        style = _require_object(raw_style, f"cell_styles.{name}")
        _reject_unknown_fields(
            style,
            set(_CELL_FORMAT_FIELDS),
            f"cell_styles.{name}",
        )
        result[name] = copy.deepcopy(style)
    return result


def _expand_run(value: Any, run_defaults: dict[str, Any]) -> dict[str, Any]:
    run = _require_object(value, "run")
    _reject_unknown_fields(run, _RUN_FIELDS, "run")
    return _format_layers(run_defaults, run)


def _expand_paragraph(
    value: Any,
    paragraph_defaults: dict[str, Any],
    run_defaults: dict[str, Any],
) -> dict[str, Any]:
    if isinstance(value, str):
        paragraph = {"text": value}
    else:
        paragraph = _require_object(value, "paragraph")
        _reject_unknown_fields(paragraph, _PARAGRAPH_FIELDS, "paragraph")
    expanded = _format_layers(paragraph_defaults, paragraph)
    if "runs" in expanded:
        runs = expanded["runs"]
        if not isinstance(runs, list):
            raise RuntimeError("Native PPTX table paragraph runs must be a list")
        expanded["runs"] = [_expand_run(run, run_defaults) for run in runs]
    return expanded


def _expand_cell(
    value: Any,
    defaults: dict[str, dict[str, Any]],
    cell_styles: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    if isinstance(value, dict):
        cell = copy.deepcopy(value)
        _reject_unknown_fields(cell, _CELL_FIELDS, "cell")
    else:
        cell = {"text": "" if value is None else str(value)}

    style_name = cell.pop("cell_style", None)
    if style_name is None:
        style = {}
    elif not isinstance(style_name, str) or style_name not in cell_styles:
        raise RuntimeError(
            f"Native PPTX table cell references unknown cell_style: {style_name!r}"
        )
    else:
        style = cell_styles[style_name]

    expanded = _format_layers(defaults["cell"], style, cell)
    # A plain-text cell is one run, and a text-only paragraph is one run in
    # the cell's colour: run defaults reach both through the cell unless the
    # cell (or its cell_style / cell defaults) already set the field. Runs
    # spelled out under paragraphs[].runs take the same defaults directly.
    for field in _PLAIN_TEXT_RUN_FIELDS:
        if field not in expanded and field in defaults["run"]:
            expanded[field] = copy.deepcopy(defaults["run"][field])
    if "paragraphs" in expanded:
        paragraphs = expanded["paragraphs"]
        if not isinstance(paragraphs, list):
            raise RuntimeError("Native PPTX table cell paragraphs must be a list")
        expanded["paragraphs"] = [
            _expand_paragraph(
                paragraph,
                defaults["paragraph"],
                defaults["run"],
            )
            for paragraph in paragraphs
        ]
    return expanded


def expand_semantic_table_payload(payload: dict[str, Any]) -> dict[str, Any]:
    """Expand semantic-table.v2 defaults and named styles into canonical cells."""
    source = _require_object(payload, "payload")
    if source.get("schema") != SEMANTIC_TABLE_SCHEMA:
        raise RuntimeError(
            "Native PPTX table metadata requires schema "
            f"{SEMANTIC_TABLE_SCHEMA!r}"
        )
    _reject_unknown_fields(source, _TOP_LEVEL_FIELDS, "payload")
    style = source.get("style", {})
    if not isinstance(style, dict):
        raise RuntimeError("Native PPTX table style must be an object")
    _reject_unknown_fields(style, _TABLE_STYLE_FIELDS, "style")

    defaults = _validated_defaults(source)
    cell_styles = _validated_cell_styles(source)
    expanded = {
        key: copy.deepcopy(value)
        for key, value in source.items()
        if key not in {"schema", "defaults", "cell_styles", "columns", "rows"}
    }
    if "columns" in source:
        columns = source["columns"]
        if not isinstance(columns, list):
            raise RuntimeError("Native PPTX table columns must be a list")
        expanded["columns"] = [
            _expand_cell(cell, defaults, cell_styles) for cell in columns
        ]
    if "rows" in source:
        rows = source["rows"]
        if not isinstance(rows, list):
            raise RuntimeError("Native PPTX table rows must be a list")
        expanded_rows: list[list[dict[str, Any]]] = []
        for row_index, row in enumerate(rows, start=1):
            if not isinstance(row, list):
                raise RuntimeError(
                    f"Native PPTX table row {row_index} must be a list"
                )
            expanded_rows.append(
                [_expand_cell(cell, defaults, cell_styles) for cell in row]
            )
        expanded["rows"] = expanded_rows
    return expanded


def _iter_cells(payload: dict[str, Any]) -> Iterable[dict[str, Any]]:
    columns = payload.get("columns")
    if isinstance(columns, list):
        for cell in columns:
            if isinstance(cell, dict):
                yield cell
    rows = payload.get("rows")
    if isinstance(rows, list):
        for row in rows:
            if not isinstance(row, list):
                continue
            for cell in row:
                if isinstance(cell, dict):
                    yield cell


def _iter_paragraphs(cells: Iterable[dict[str, Any]]) -> Iterable[dict[str, Any]]:
    for cell in cells:
        paragraphs = cell.get("paragraphs")
        if not isinstance(paragraphs, list):
            continue
        for paragraph in paragraphs:
            if isinstance(paragraph, dict):
                yield paragraph


def _iter_runs(paragraphs: Iterable[dict[str, Any]]) -> Iterable[dict[str, Any]]:
    for paragraph in paragraphs:
        runs = paragraph.get("runs")
        if not isinstance(runs, list):
            continue
        for run in runs:
            if isinstance(run, dict):
                yield run


def _most_common_value(items: list[dict[str, Any]], field: str) -> Any:
    counts: dict[str, int] = {}
    values: dict[str, Any] = {}
    order: list[str] = []
    for item in items:
        signature = _canonical_json(item[field])
        if signature not in counts:
            counts[signature] = 0
            values[signature] = item[field]
            order.append(signature)
        counts[signature] += 1
    winner = max(order, key=lambda signature: counts[signature])
    return copy.deepcopy(values[winner])


def _promote_defaults(
    payload: dict[str, Any],
    defaults: dict[str, dict[str, Any]],
    section: str,
    items: list[dict[str, Any]],
    fields: tuple[str, ...],
) -> None:
    if not items:
        return
    for field in fields:
        if any(field not in item for item in items):
            continue
        value = _most_common_value(items, field)
        before = _payload_size(payload, defaults)
        matching = [item for item in items if item[field] == value]
        defaults[section][field] = value
        for item in matching:
            del item[field]
        after = _payload_size(payload, defaults)
        if after < before:
            continue
        del defaults[section][field]
        for item in matching:
            item[field] = copy.deepcopy(value)


def _compact_plain_paragraphs(payload: dict[str, Any]) -> None:
    for cell in _iter_cells(payload):
        paragraphs = cell.get("paragraphs")
        if not isinstance(paragraphs, list):
            continue
        for index, paragraph in enumerate(paragraphs):
            if (
                isinstance(paragraph, dict)
                and set(paragraph) == {"text"}
                and isinstance(paragraph["text"], str)
            ):
                paragraphs[index] = paragraph["text"]


def _cell_style_signature(cell: dict[str, Any]) -> dict[str, Any]:
    return {
        field: copy.deepcopy(cell[field])
        for field in _CELL_FORMAT_FIELDS
        if field in cell
    }


def _factor_cell_styles(
    payload: dict[str, Any],
    defaults: dict[str, dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    cells = list(_iter_cells(payload))
    signatures: dict[str, dict[str, Any]] = {}
    matching_cells: dict[str, list[dict[str, Any]]] = {}
    order: list[str] = []
    for cell in cells:
        style = _cell_style_signature(cell)
        if not style:
            continue
        signature = _canonical_json(style)
        if signature not in signatures:
            signatures[signature] = style
            matching_cells[signature] = []
            order.append(signature)
        matching_cells[signature].append(cell)

    cell_styles: dict[str, dict[str, Any]] = {}
    for signature in order:
        matched = matching_cells[signature]
        if len(matched) < 2:
            continue
        name = f"cell-{len(cell_styles) + 1}"
        before = len(
            _canonical_json(
                {"payload": payload, "defaults": defaults, "cell_styles": cell_styles}
            )
        )
        style = signatures[signature]
        cell_styles[name] = copy.deepcopy(style)
        for cell in matched:
            for field in style:
                del cell[field]
            cell["cell_style"] = name
        after = len(
            _canonical_json(
                {"payload": payload, "defaults": defaults, "cell_styles": cell_styles}
            )
        )
        if after < before:
            continue
        del cell_styles[name]
        for cell in matched:
            del cell["cell_style"]
            cell.update(copy.deepcopy(style))
    return cell_styles


def compact_semantic_table_payload(payload: dict[str, Any]) -> dict[str, Any]:
    """Return a deterministic, lossless semantic-table.v2 representation."""
    source = _require_object(payload, "payload")
    if "schema" in source:
        working = expand_semantic_table_payload(source)
    else:
        working = copy.deepcopy(source)

    defaults: dict[str, dict[str, Any]] = {
        "cell": {},
        "paragraph": {},
        "run": {},
    }
    cells = list(_iter_cells(working))
    total_cells = len(working.get("columns", [])) + sum(
        len(row) for row in working.get("rows", []) if isinstance(row, list)
    )
    if cells and len(cells) == total_cells:
        _promote_defaults(
            working,
            defaults,
            "cell",
            cells,
            _CELL_FORMAT_FIELDS,
        )

    paragraphs = list(_iter_paragraphs(cells))
    paragraph_count = sum(
        len(cell["paragraphs"])
        for cell in cells
        if isinstance(cell.get("paragraphs"), list)
    )
    if paragraphs and len(paragraphs) == paragraph_count:
        _promote_defaults(
            working,
            defaults,
            "paragraph",
            paragraphs,
            _PARAGRAPH_DEFAULT_FIELDS,
        )

    runs = list(_iter_runs(paragraphs))
    run_count = sum(
        len(paragraph["runs"])
        for paragraph in paragraphs
        if isinstance(paragraph.get("runs"), list)
    )
    if runs and len(runs) == run_count:
        _promote_defaults(
            working,
            defaults,
            "run",
            runs,
            _RUN_DEFAULT_FIELDS,
        )

    _compact_plain_paragraphs(working)
    cell_styles = _factor_cell_styles(working, defaults)
    compact_defaults = {
        section: values for section, values in defaults.items() if values
    }

    result: dict[str, Any] = {"schema": SEMANTIC_TABLE_SCHEMA}
    for key, value in working.items():
        if key not in {"columns", "rows"}:
            result[key] = value
    if compact_defaults:
        result["defaults"] = compact_defaults
    if cell_styles:
        result["cell_styles"] = cell_styles
    if "columns" in working:
        result["columns"] = working["columns"]
    if "rows" in working:
        result["rows"] = working["rows"]

    expand_semantic_table_payload(result)
    return result


# ---------------------------------------------------------------------------
# Fallback projection (SVG-first): derive the payload from the drawn table.
# ---------------------------------------------------------------------------

SVG_NS = "http://www.w3.org/2000/svg"
_EDGE_TOLERANCE = 2.0
# Style fields a drawing cannot express; carried over from an existing payload.
_NON_DERIVABLE_STYLE = ("padding", "valign", "lang", "table_style_id")
_BOLD_WEIGHTS = {"bold", "600", "700", "800", "900"}


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _num(value: Any) -> float | None:
    try:
        number = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    return number


def _tidy(value: float) -> Any:
    return int(value) if float(value).is_integer() else round(value, 2)


def _inherited(elem: ET.Element, parents: dict, name: str) -> str | None:
    current: ET.Element | None = elem
    while current is not None:
        value = current.get(name)
        if value is not None:
            return value
        current = parents.get(id(current))
    return None


def _dedupe_edges(values: list[float]) -> list[float]:
    edges: list[float] = []
    for value in sorted(values):
        if edges and abs(edges[-1] - value) <= _EDGE_TOLERANCE:
            edges[-1] = (edges[-1] + value) / 2
        else:
            edges.append(value)
    return edges


def _frame(group: ET.Element) -> tuple[float, float, float, float]:
    raw = group.get("data-pptx-bounds")
    parts = [_num(value) for value in (raw or "").replace(",", " ").split()]
    if len(parts) != 4 or any(v is None for v in parts) or parts[2] <= 0 or parts[3] <= 0:
        raise RuntimeError("table group needs a positive data-pptx-bounds frame")
    return parts[0], parts[1], parts[2], parts[3]


def _collect(group: ET.Element) -> tuple[dict, list, list, list]:
    """Return the parent map plus the drawn rects, lines, and texts."""
    parents: dict[int, ET.Element] = {}
    rects: list[ET.Element] = []
    lines: list[ET.Element] = []
    texts: list[ET.Element] = []
    stack = [group]
    while stack:
        node = stack.pop()
        for child in list(node):
            if _local(child.tag) == "metadata":
                continue
            parents[id(child)] = node
            stack.append(child)
            kind = _local(child.tag)
            if kind == "rect":
                rects.append(child)
            elif kind == "line":
                lines.append(child)
            elif kind == "text":
                texts.append(child)
    return parents, rects, lines, texts


def _grid_edges(lines: list, frame: tuple) -> tuple[list[float], list[float]]:
    x, y, width, height = frame
    left, right, top, bottom = x, x + width, y, y + height
    cols, rows = [left, right], [top, bottom]
    for line in lines:
        x1, y1 = _num(line.get("x1")), _num(line.get("y1"))
        x2, y2 = _num(line.get("x2")), _num(line.get("y2"))
        if None in (x1, y1, x2, y2):
            continue
        vertical = abs(x1 - x2) <= _EDGE_TOLERANCE
        horizontal = abs(y1 - y2) <= _EDGE_TOLERANCE
        if vertical and left - _EDGE_TOLERANCE <= x1 <= right + _EDGE_TOLERANCE:
            cols.append(x1)
        elif horizontal and top - _EDGE_TOLERANCE <= y1 <= bottom + _EDGE_TOLERANCE:
            rows.append(y1)
    return _dedupe_edges(cols), _dedupe_edges(rows)


def _slot(value: float, edges: list[float]) -> int | None:
    for index in range(len(edges) - 1):
        low, high = edges[index], edges[index + 1]
        if low - _EDGE_TOLERANCE <= value <= high + _EDGE_TOLERANCE:
            return index
    return None


def _text_cell(elem: ET.Element, parents: dict, cols: list, rows: list) -> tuple | None:
    """Place one <text> in the grid and read the style the drawing gives it."""
    x, y = _num(elem.get("x")), _num(elem.get("y"))
    if x is None or y is None:
        return None
    anchor = (_inherited(elem, parents, "text-anchor") or "start").strip().lower()
    # Map the anchor back to the glyph box's horizontal centre of gravity so an
    # end-anchored numeric column lands in its own cell, not the neighbour's.
    probe = x - 1.0 if anchor in {"end", "right"} else x + 1.0
    column = _slot(probe, cols)
    row = _slot(y, rows)
    if column is None or row is None:
        return None
    content = "".join(elem.itertext()).strip()
    if not content:
        return None
    cell: dict[str, Any] = {"text": content}
    align = {"middle": "ctr", "end": "r", "right": "r"}.get(anchor, "l")
    cell["align"] = align
    colour = _inherited(elem, parents, "fill")
    if colour:
        cell["color"] = colour.strip().upper()
    weight = (_inherited(elem, parents, "font-weight") or "").strip().lower()
    if weight in _BOLD_WEIGHTS:
        cell["bold"] = True
    size = _num(_inherited(elem, parents, "font-size"))
    if size is not None:
        cell["font_size"] = _tidy(size)
    return row, column, cell


def _row_fill(rects: list, frame: tuple, rows: list, index: int) -> str | None:
    """Return a fill that covers the whole of one grid row."""
    x, _y, width, _height = frame
    top, bottom = rows[index], rows[index + 1]
    for rect in rects:
        fill = (rect.get("fill") or "").strip()
        if not fill or fill.lower() == "none":
            continue
        rx, ry = _num(rect.get("x")) or 0.0, _num(rect.get("y")) or 0.0
        rw, rh = _num(rect.get("width")), _num(rect.get("height"))
        if rw is None or rh is None:
            continue
        spans_width = rx <= x + _EDGE_TOLERANCE and rx + rw >= x + width - _EDGE_TOLERANCE
        spans_row = ry <= top + _EDGE_TOLERANCE and ry + rh >= bottom - _EDGE_TOLERANCE
        if spans_width and spans_row and rh <= (bottom - top) + _EDGE_TOLERANCE * 2:
            return fill.upper()
    return None


def _drawn_segments(lines: list) -> tuple[list, list]:
    """Split drawn rules into vertical and horizontal spans."""
    vertical, horizontal = [], []
    for line in lines:
        x1, y1 = _num(line.get("x1")), _num(line.get("y1"))
        x2, y2 = _num(line.get("x2")), _num(line.get("y2"))
        if None in (x1, y1, x2, y2):
            continue
        if abs(x1 - x2) <= _EDGE_TOLERANCE:
            vertical.append((x1, min(y1, y2), max(y1, y2)))
        elif abs(y1 - y2) <= _EDGE_TOLERANCE:
            horizontal.append((y1, min(x1, x2), max(x1, x2)))
    return vertical, horizontal


def _covered(position: float, low: float, high: float, segments: list) -> bool:
    for at, start, end in segments:
        if (
            abs(at - position) <= _EDGE_TOLERANCE
            and start <= low + _EDGE_TOLERANCE
            and end >= high - _EDGE_TOLERANCE
        ):
            return True
    return False


def project_table_payload(
    group: ET.Element,
    *,
    header_rows: int = 1,
    existing: dict[str, Any] | None = None,
    root: ET.Element | None = None,
) -> dict[str, Any]:
    """Derive a semantic-table.v2 payload from one drawn fallback group."""
    frame = _frame(group)
    parents, rects, lines, texts = _collect(group)
    if root is not None:
        ancestors = {id(c): n for n in root.iter() for c in list(n)}
        parents.update({k: v for k, v in ancestors.items() if k not in parents})
    cols, rows = _grid_edges(lines, frame)
    if len(cols) < 2 or len(rows) < 2:
        raise RuntimeError("fallback has no resolvable cell grid")
    col_count, row_count = len(cols) - 1, len(rows) - 1
    if not 0 <= header_rows < row_count:
        raise RuntimeError(f"header_rows must be between 0 and {row_count - 1}")

    grid: list[list[dict]] = [[{} for _ in range(col_count)] for _ in range(row_count)]
    sizes: list[float] = []
    for elem in texts:
        placed = _text_cell(elem, parents, cols, rows)
        if placed is None:
            continue
        row, column, cell = placed
        size = cell.get("font_size")
        if size is not None:
            sizes.append(float(size))
        if grid[row][column]:
            grid[row][column]["text"] += " " + cell["text"]
        else:
            grid[row][column] = cell

    vertical, horizontal = _drawn_segments(lines)
    full_edges = covered_edges = 0
    for row in range(row_count):
        top, bottom = rows[row], rows[row + 1]
        fill = _row_fill(rects, frame, rows, row)
        for column in range(col_count):
            left, right = cols[column], cols[column + 1]
            cell = grid[row][column]
            if fill and "fill" not in cell:
                cell["fill"] = fill
            sides = {
                "left": _covered(left, top, bottom, vertical),
                "right": _covered(right, top, bottom, vertical),
                "top": _covered(top, left, right, horizontal),
                "bottom": _covered(bottom, left, right, horizontal),
            }
            full_edges += 4
            covered_edges += sum(sides.values())
            cell["_sides"] = sides

    body_size = max(set(sizes), key=sizes.count) if sizes else None
    header_size = None
    if header_rows:
        header_sizes = [
            float(cell["font_size"])
            for cell in grid[0] if cell.get("font_size") is not None
        ]
        if header_sizes:
            header_size = max(set(header_sizes), key=header_sizes.count)
            if header_size == body_size:
                header_size = None

    outline = next(
        (
            rect for rect in rects
            if (rect.get("stroke") or "").strip().lower() not in {"", "none"}
        ),
        None,
    )
    style: dict[str, Any] = {}
    family = _inherited(group, parents, "font-family") or group.get("font-family")
    if family:
        style["font_family"] = family.split(",")[0].strip().strip("'\"")
    if body_size is not None:
        style["font_size"] = _tidy(body_size)
    if header_size is not None:
        style["header_font_size"] = _tidy(header_size)
    if outline is not None:
        style["border_color"] = (outline.get("stroke") or "").strip().upper()
    # Cell rules, not the outer frame, set the per-cell border width.
    rule_widths = [
        width for line in lines
        if (width := _num(_inherited(line, parents, "stroke-width"))) is not None
    ]
    if rule_widths:
        style["border_width"] = _tidy(max(set(rule_widths), key=rule_widths.count))
    elif outline is not None and (width := _num(outline.get("stroke-width"))):
        style["border_width"] = _tidy(width)
    if existing and isinstance(existing.get("style"), dict):
        for key in _NON_DERIVABLE_STYLE:
            if key in existing["style"]:
                style[key] = copy.deepcopy(existing["style"][key])

    # Sparse rules must be projected per side; a complete grid stays uniform.
    sparse = covered_edges < full_edges
    for row in range(row_count):
        for column in range(col_count):
            cell = grid[row][column]
            sides = cell.pop("_sides")
            if not sparse:
                continue
            border_colour = style.get("border_color")
            cell["borders"] = {
                side: (
                    {"style": "solid", "color": border_colour}
                    if drawn and border_colour else
                    {"style": "solid"} if drawn else {"style": "none"}
                )
                for side, drawn in sides.items()
            }

    payload: dict[str, Any] = {
        "schema": SEMANTIC_TABLE_SCHEMA,
        "x": _tidy(frame[0]),
        "y": _tidy(frame[1]),
        "width": _tidy(frame[2]),
        "height": _tidy(frame[3]),
        "column_widths": [_tidy(cols[i + 1] - cols[i]) for i in range(col_count)],
        "row_heights": [_tidy(rows[i + 1] - rows[i]) for i in range(row_count)],
    }
    name = group.get("id") or (existing or {}).get("name")
    if name:
        payload["name"] = name
    if existing and "strict_grid" in existing:
        payload["strict_grid"] = existing["strict_grid"]
    if style:
        payload["style"] = style
    if header_rows:
        payload["header_rows"] = header_rows
        payload["columns"] = [dict(cell) or {"text": ""} for cell in grid[0]]
        payload["rows"] = [[dict(cell) or {"text": ""} for cell in row] for row in grid[1:]]
    else:
        payload["rows"] = [[dict(cell) or {"text": ""} for cell in row] for row in grid]

    # Round-trip through the validator so a bad projection fails here, loudly,
    # rather than at export.
    expand_semantic_table_payload(payload)
    return compact_semantic_table_payload(payload)


def find_table_groups(root: ET.Element) -> list[ET.Element]:
    """Return every active table replacement group, outermost first."""
    return [
        elem for elem in root.iter()
        if _local(elem.tag) == "g"
        and (elem.get("data-pptx-replace-with") or "").strip() == "table"
    ]


def _select_group(root: ET.Element, group_id: str | None) -> ET.Element:
    groups = find_table_groups(root)
    if group_id:
        for group in groups:
            if (group.get("id") or "") == group_id:
                return group
        raise RuntimeError(f"no table group with id {group_id!r}")
    if not groups:
        raise RuntimeError('no <g data-pptx-replace-with="table"> in this SVG')
    if len(groups) > 1:
        names = ", ".join(group.get("id") or "<no id>" for group in groups)
        raise RuntimeError(f"several table groups; choose one with --group: {names}")
    return groups[0]


def _existing_payload(group: ET.Element) -> dict[str, Any] | None:
    for child in group:
        if _local(child.tag) == "metadata":
            try:
                return json.loads("".join(child.itertext()))
            except ValueError:
                return None
    return None


def _load_payload(path: str) -> dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def main(argv: list[str] | None = None) -> int:
    """Run the CLI entry point."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Project, compact, expand, or lint a semantic-table.v2 payload.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    project = sub.add_parser(
        "project",
        help="Derive the payload from a drawn table fallback (SVG-first).",
    )
    project.add_argument("svg", help="SVG file containing the table group")
    project.add_argument("--group", help="id of the table group, when the page has several")
    project.add_argument("--header-rows", type=int, default=1)

    for name, helptext in (
        ("compact", "Factor defaults and repeated cell styles."),
        ("expand", "Expand defaults and named styles into canonical cells."),
        ("lint", "Validate a payload; print nothing and exit 0 when it is legal."),
    ):
        command = sub.add_parser(name, help=helptext)
        command.add_argument("payload", help="JSON payload file")

    args = parser.parse_args(argv)
    try:
        if args.command == "project":
            root = ET.parse(args.svg).getroot()
            group = _select_group(root, args.group)
            result = project_table_payload(
                group,
                header_rows=args.header_rows,
                existing=_existing_payload(group),
                root=root,
            )
        elif args.command == "compact":
            result = compact_semantic_table_payload(_load_payload(args.payload))
        elif args.command == "expand":
            result = expand_semantic_table_payload(_load_payload(args.payload))
        else:
            expand_semantic_table_payload(_load_payload(args.payload))
            return 0
    except (ET.ParseError, OSError, RuntimeError, ValueError) as exc:
        print(f"{args.command} failed: {exc}", file=sys.stderr)
        return 2

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
