#!/usr/bin/env python3
"""PPT Master - SVG Text Inspection

List every <text> node with its resolved geometry, style, measured width, and
owning container, so a repair pass never reads a whole decorated SVG.

Containers follow the checker: an authored page's text is bounded by its root
<g> data-pptx-bounds module and by the root viewBox, while data-pptx-frame is
the round-trip route's container and is reported separately.

Usage:
    python3 scripts/svg_text_inspect.py <svg-file-or-directory> [options]
Examples:
    python3 scripts/svg_text_inspect.py projects/demo/svg_output --overflowing-only
    python3 scripts/svg_text_inspect.py svg_output/01_cover.svg --scope title-slot --json
Dependencies:
    Standard library and PPT Master sibling modules
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from xml.etree import ElementTree as ET


_SCRIPTS_DIR = Path(__file__).resolve().parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from console_encoding import configure_utf8_stdio  # noqa: E402
from slide_roster import discover_slide_svgs  # noqa: E402
from svg_quality.checker import (  # noqa: E402
    SVG_NS,
    SVGQualityChecker,
    _effective_presentation_value,
    _estimate_single_line_text_frame_width,
    _local_name,
    _parse_viewbox_values,
    _resolve_project_font_sizes,
    _resolve_project_letter_spacings,
)
from template_text_slots import _text_selector  # noqa: E402


def _round(value: float | None) -> float | None:
    return None if value is None else round(float(value), 2)


def _rect(bounds: tuple[float, float, float, float] | None) -> dict | None:
    if bounds is None:
        return None
    left, top, right, bottom = bounds
    return {
        'x': _round(left),
        'y': _round(top),
        'width': _round(right - left),
        'height': _round(bottom - top),
    }


def _parent_maps(root: ET.Element) -> tuple[dict, dict]:
    """Build the id-keyed map the checker uses and the element-keyed one slots use."""
    by_id: dict[int, ET.Element] = {}
    by_child: dict[ET.Element, ET.Element] = {}
    for parent in root.iter():
        for child in list(parent):
            by_id[id(child)] = parent
            by_child[child] = parent
    return by_id, by_child


def _module_map(root: ET.Element) -> dict[int, dict]:
    """Map each text node to its root <g> module, as the checker's walk does."""
    modules: dict[int, dict] = {}
    for module in list(root):
        if _local_name(module) != 'g':
            continue
        resolved = SVGQualityChecker._resolved_root_module_bounds(module)
        if resolved is None:
            continue
        attribute, bounds = resolved
        record = {
            'label': f"{(module.get('id') or '<g>')} {attribute}",
            'bounds': bounds,
        }
        for text_element in module.iter(f'{{{SVG_NS}}}text'):
            modules[id(text_element)] = record
    return modules


def _frame_record(text_element: ET.Element, parent_by_id: dict) -> dict | None:
    """Resolve the round-trip data-pptx-frame container, when one applies."""
    label, bounds, error, inferred = SVGQualityChecker._roundtrip_text_frame(
        text_element,
        parent_by_id,
    )
    if label is None and bounds is None and error is None:
        return None
    record = {'label': label, 'inferred': inferred, 'error': error}
    record.update(_rect(bounds) or {})
    return record


def _line_records(lines: list, include_headroom: bool) -> list[dict]:
    records = []
    for _line_element, x, y, runs, font_size in lines:
        try:
            width = float(_estimate_single_line_text_frame_width(
                runs,
                include_headroom=include_headroom,
            ))
        except (KeyError, TypeError, ValueError):
            width = None
        records.append({
            'text': ''.join(str(run.get('text', '')) for run in runs),
            'x': _round(x),
            'y': _round(y),
            'font_size': _round(font_size),
            'width': _round(width),
            'runs': len(runs),
        })
    return records


def _overflow(
    bounds: tuple[float, float, float, float] | None,
    container: tuple[float, float, float, float] | None,
) -> dict | None:
    """Compare bounds to a container with the checker's own tolerance."""
    if bounds is None or container is None:
        return None
    # The checker returns None for "within tolerance"; that is the fits case.
    metrics = SVGQualityChecker._bounds_overflow_metrics(bounds, container)
    if metrics is None:
        return {'axes': None, 'horizontal_ratio': 0.0, 'vertical_ratio': 0.0, 'fits': True}
    axes, horizontal, vertical = metrics
    return {
        'axes': axes,
        'horizontal_ratio': _round(horizontal),
        'vertical_ratio': _round(vertical),
        'fits': False,
    }


def _scoped_text_ids(root: ET.Element, scope_id: str) -> set[int]:
    """Collect text nodes under the element carrying ``scope_id``."""
    scoped: set[int] = set()
    for element in root.iter():
        if (element.get('id') or '').strip() != scope_id:
            continue
        for text_element in element.iter(f'{{{SVG_NS}}}text'):
            scoped.add(id(text_element))
    return scoped


def inspect_svg(path: Path, *, scope_id: str | None = None) -> list[dict]:
    """Project one SVG's text nodes into compact records."""
    root = ET.parse(path).getroot()
    font_sizes = _resolve_project_font_sizes(root)
    letter_spacings = _resolve_project_letter_spacings(root, font_sizes)
    parent_by_id, parent_by_child = _parent_maps(root)
    modules = _module_map(root)
    scoped = _scoped_text_ids(root, scope_id) if scope_id else None

    viewbox = _parse_viewbox_values(root.get('viewBox') or '')
    canvas = None
    if viewbox is not None:
        x, y, width, height = viewbox
        canvas = (x, y, x + width, y + height)

    records: list[dict] = []
    for text_element in root.iter(f'{{{SVG_NS}}}text'):
        if scoped is not None and id(text_element) not in scoped:
            continue
        lines = SVGQualityChecker._resolved_text_lines(
            text_element,
            parent_by_id,
            font_sizes,
            letter_spacings,
        )
        # The checker measures page fit without headroom and module fit with
        # it; reporting one number for both would disagree with the gate.
        module_bounds = SVGQualityChecker._estimated_text_bounds(
            text_element, parent_by_id, font_sizes, letter_spacings,
            include_headroom=True,
        )
        canvas_bounds = SVGQualityChecker._estimated_text_bounds(
            text_element, parent_by_id, font_sizes, letter_spacings,
            include_headroom=False,
        )
        module = modules.get(id(text_element))
        line_records = _line_records(lines, True) if lines else []
        records.append({
            'file': path.name,
            'selector': _text_selector(text_element, parent_by_child, compact=True),
            'text': ''.join(text_element.itertext()).strip(),
            'hidden': SVGQualityChecker._is_hidden_element(text_element, parent_by_id),
            'font_family': _effective_presentation_value(
                text_element, 'font-family', parent_by_id,
            ),
            'font_weight': _effective_presentation_value(
                text_element, 'font-weight', parent_by_id,
            ),
            'fill': _effective_presentation_value(text_element, 'fill', parent_by_id),
            'text_anchor': _effective_presentation_value(
                text_element, 'text-anchor', parent_by_id,
            ) or 'start',
            'lines': line_records,
            'width': max(
                (line['width'] for line in line_records if line['width'] is not None),
                default=None,
            ),
            'bounds': _rect(module_bounds),
            'module': None if module is None else {
                'label': module['label'], **(_rect(module['bounds']) or {}),
            },
            'module_overflow': _overflow(
                module_bounds, None if module is None else module['bounds'],
            ),
            'canvas_overflow': _overflow(canvas_bounds, canvas),
            'frame': _frame_record(text_element, parent_by_id),
        })
    return records


def _verdict(record: dict) -> str:
    """Name the first failing container, matching checker severity order."""
    canvas = record['canvas_overflow']
    module = record['module_overflow']
    if canvas is not None and not canvas['fits']:
        return f"OFF-CANVAS {canvas['axes']}"
    if module is None:
        return 'no module bounds'
    if not module['fits']:
        severity = 'error' if module['horizontal_ratio'] > 0.05 or module['vertical_ratio'] > 0.05 else 'warn'
        return f"OVERFLOW {module['axes']} ({severity})"
    return 'fits'


def _fits(record: dict) -> bool:
    canvas = record['canvas_overflow']
    module = record['module_overflow']
    return (
        (canvas is None or canvas['fits'])
        and (module is None or module['fits'])
    )


def _format_record(record: dict) -> str:
    """Render one record as a single scannable line plus its text."""
    module = record['module'] or {}
    size = record['lines'][0]['font_size'] if record['lines'] else None
    module_width = module.get('width')
    slack = (
        None
        if module_width is None or record['width'] is None
        else round(module_width - record['width'], 2)
    )
    head = (
        f"  {record['selector']}  {size}px/{record['font_family']} "
        f"anchor={record['text_anchor']} w={record['width']} "
        f"module={module_width} slack={slack}  [{_verdict(record)}]"
    )
    if record['hidden']:
        head += ' (hidden)'
    return f"{head}\n      {record['text'][:96]!r}"


def _targets(target: Path) -> list[Path]:
    if target.is_file():
        return [target]
    svg_root = target / 'svg_output' if (target / 'svg_output').is_dir() else target
    return discover_slide_svgs(svg_root) if svg_root.is_dir() else []


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description='List <text> nodes with geometry, style, width, and owning container.',
    )
    parser.add_argument('target', type=Path, help='SVG file, svg_output/, or project')
    parser.add_argument('--scope', help='Only text under the element with this id')
    parser.add_argument('--json', action='store_true')
    parser.add_argument(
        '--overflowing-only',
        action='store_true',
        help='Skip visible text that fits every container the gate checks.',
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    configure_utf8_stdio()
    args = build_parser().parse_args(argv)
    paths = _targets(args.target)
    if not paths:
        print(f'No SVG files found in: {args.target}', file=sys.stderr)
        return 2

    records: list[dict] = []
    for path in paths:
        try:
            records.extend(inspect_svg(path, scope_id=args.scope))
        except (ET.ParseError, OSError, ValueError) as exc:
            print(f'{path.name}: {exc}', file=sys.stderr)
            return 1

    if args.overflowing_only:
        records = [
            record for record in records
            if not record['hidden'] and not _fits(record)
        ]

    if args.json:
        print(json.dumps(records, ensure_ascii=False, indent=2))
        return 0

    if not records:
        print('[OK] No matching text nodes.')
        return 0
    current_file = None
    for record in records:
        if record['file'] != current_file:
            current_file = record['file']
            print(f'\n{current_file}')
        print(_format_record(record))
    files = len({record['file'] for record in records})
    print(f'\n{len(records)} text node(s) in {files} file(s)')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
