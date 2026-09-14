#!/usr/bin/env python3
"""Focused tests for projecting a semantic-table.v2 payload from its fallback."""

from __future__ import annotations

import json
import subprocess
import sys
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET


SCRIPTS_DIR = Path(__file__).resolve().parents[1]
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from semantic_table import (  # noqa: E402
    expand_semantic_table_payload,
    project_table_payload,
)


SCRIPT = SCRIPTS_DIR / 'semantic_table.py'
CHECKER = SCRIPTS_DIR / 'svg_quality_checker.py'
RECORD_TABLE = SCRIPTS_DIR.parent / 'templates' / 'tables' / 'record_table.svg'

# Two columns, one header row and two body rows, with a banded middle row and
# a deliberately sparse rule set (no verticals inside the body).
FALLBACK = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 200"
     font-family="Georgia, serif" font-size="12">
  <g id="t" data-pptx-bounds="0 0 200 90" data-pptx-replace-with="table">
    <rect x="0" y="0" width="200" height="30" fill="#EEEEEE"/>
    <rect x="0" y="60" width="200" height="30" fill="#FAFAFA"/>
    <rect x="0" y="0" width="200" height="90" fill="none" stroke="#CCCCCC" stroke-width="2"/>
    <g stroke="#CCCCCC" stroke-width="1">
      <line x1="100" y1="0" x2="100" y2="30"/>
      <line x1="0" y1="30" x2="200" y2="30"/>
      <line x1="0" y1="60" x2="200" y2="60"/>
    </g>
    <g font-size="11" font-weight="700" fill="#333333">
      <text x="10" y="20">NAME</text>
      <text x="190" y="20" text-anchor="end">QTY</text>
    </g>
    <text x="10" y="50" fill="#111111">Widget</text>
    <text x="190" y="50" text-anchor="end">12</text>
    <text x="10" y="80" fill="#111111">Gadget</text>
    <text x="150" y="80" text-anchor="middle">7</text>
  </g>
</svg>
"""


def _group(svg: str = FALLBACK, group_id: str = 't') -> tuple[ET.Element, ET.Element]:
    root = ET.fromstring(svg)
    group = next(
        elem for elem in root.iter()
        if elem.tag.endswith('}g') and elem.get('id') == group_id
    )
    return root, group


class SemanticTableProjectionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.root, self.group = _group()
        self.payload = project_table_payload(self.group, root=self.root)

    def test_grid_is_derived_from_drawn_rules(self) -> None:
        self.assertEqual(self.payload['column_widths'], [100, 100])
        self.assertEqual(self.payload['row_heights'], [30, 30, 30])
        self.assertEqual(self.payload['header_rows'], 1)
        self.assertEqual(len(self.payload['columns']), 2)
        self.assertEqual(len(self.payload['rows']), 2)

    def test_alignment_comes_from_text_anchor(self) -> None:
        expanded = expand_semantic_table_payload(self.payload)
        self.assertEqual(expanded['columns'][0]['align'], 'l')
        self.assertEqual(expanded['columns'][1]['align'], 'r')
        self.assertEqual(expanded['rows'][0][1]['align'], 'r')
        self.assertEqual(expanded['rows'][1][1]['align'], 'ctr')

    def test_text_style_and_row_fill_reach_the_cells(self) -> None:
        expanded = expand_semantic_table_payload(self.payload)
        header = expanded['columns'][0]
        self.assertEqual(header['text'], 'NAME')
        self.assertTrue(header['bold'])
        self.assertEqual(header['fill'], '#EEEEEE')
        self.assertEqual(expanded['rows'][0][0]['color'], '#111111')
        # The banded body row's whole-row fill must reach every cell in it.
        self.assertEqual(expanded['rows'][1][0]['fill'], '#FAFAFA')
        self.assertEqual(expanded['rows'][1][1]['fill'], '#FAFAFA')

    def test_sparse_rules_become_per_side_borders(self) -> None:
        expanded = expand_semantic_table_payload(self.payload)
        borders = expanded['columns'][0]['borders']
        self.assertEqual(borders['right']['style'], 'solid')
        self.assertEqual(borders['bottom']['style'], 'solid')
        # No vertical rule runs through the body, so that edge is not drawn.
        self.assertEqual(expanded['rows'][0][0]['borders']['right']['style'], 'none')

    def test_style_derives_family_size_and_border_width(self) -> None:
        style = self.payload['style']
        self.assertEqual(style['font_family'], 'Georgia')
        self.assertEqual(style['font_size'], 12)
        self.assertEqual(style['header_font_size'], 11)
        self.assertEqual(style['border_color'], '#CCCCCC')
        # Cell rules are 1px; the 2px outer frame must not set the cell width.
        self.assertEqual(style['border_width'], 1)

    def test_non_derivable_style_is_carried_from_the_existing_payload(self) -> None:
        existing = {
            'style': {
                'padding': {'left': 9, 'right': 9, 'top': 4, 'bottom': 4},
                'valign': 'middle',
                'table_style_id': 'keep-me',
            },
        }
        payload = project_table_payload(self.group, root=self.root, existing=existing)
        self.assertEqual(payload['style']['padding']['left'], 9)
        self.assertEqual(payload['style']['valign'], 'middle')
        self.assertEqual(payload['style']['table_style_id'], 'keep-me')

    def test_projection_of_the_canonical_table_clears_parity_warnings(self) -> None:
        """The repo's own hand-written payload trips five parity findings."""
        def parity(path: Path) -> int:
            output = subprocess.run(
                [sys.executable, str(CHECKER), str(path)],
                capture_output=True, text=True, check=False,
            ).stdout
            return output.count('blocks --native-charts-and-tables export')

        self.assertGreater(parity(RECORD_TABLE), 0)

        projected = subprocess.run(
            [sys.executable, str(SCRIPT), 'project', str(RECORD_TABLE)],
            capture_output=True, text=True, check=True,
        ).stdout
        body = json.dumps(json.loads(projected), ensure_ascii=False, indent=2)

        import re
        import tempfile
        source = RECORD_TABLE.read_text(encoding='utf-8')
        patched = re.sub(
            r'(<metadata[^>]*>).*?(</metadata>)',
            lambda m: f'{m.group(1)}\n{body}\n{m.group(2)}',
            source, count=1, flags=re.S,
        )
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / RECORD_TABLE.name
            target.write_text(patched, encoding='utf-8')
            self.assertEqual(parity(target), 0)

    def test_cli_requires_a_group_when_the_page_has_several(self) -> None:
        import tempfile
        two = FALLBACK.replace('</svg>', FALLBACK.split('\n', 2)[2].replace('id="t"', 'id="u"'))
        with tempfile.TemporaryDirectory() as tmp:
            page = Path(tmp) / 'two.svg'
            page.write_text(two, encoding='utf-8')
            result = subprocess.run(
                [sys.executable, str(SCRIPT), 'project', str(page)],
                capture_output=True, text=True, check=False,
            )
            self.assertEqual(result.returncode, 2)
            self.assertIn('--group', result.stderr)


if __name__ == '__main__':
    unittest.main()
