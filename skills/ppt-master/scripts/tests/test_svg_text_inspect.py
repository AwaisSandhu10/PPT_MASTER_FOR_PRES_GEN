#!/usr/bin/env python3
"""Focused tests for the authoring-time SVG text inspector."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPTS_DIR = Path(__file__).resolve().parents[1]
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from svg_text_inspect import inspect_svg  # noqa: E402


SCRIPT = SCRIPTS_DIR / 'svg_text_inspect.py'

# Authored pages are bounded by the root <g> data-pptx-bounds module, so the
# fixture carries one module the text fits, one it cannot, one whose text
# leaves the canvas, and one group with no module bounds at all.
PAGE_ONE = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720"
     width="1280" height="720" font-family="Arial" font-size="16">
  <rect width="1280" height="720" fill="#FFFFFF"/>
  <g id="title-slot" data-pptx-bounds="80 60 600 80">
    <text id="title" x="80" y="120" font-size="40" fill="#0F172A">Short title</text>
  </g>
  <g id="narrow-slot" data-pptx-bounds="80 300 120 60">
    <text id="overflowing" x="80" y="340" font-size="32" fill="#0F172A">This line is far too wide for its module</text>
  </g>
  <g id="edge-slot" data-pptx-bounds="1100 500 160 60">
    <text id="offcanvas" x="1100" y="540" font-size="28" fill="#0F172A">Runs past the right edge of the canvas</text>
  </g>
  <g id="loose">
    <rect x="80" y="620" width="400" height="60" fill="none"/>
    <text id="unbounded" x="80" y="660" font-size="20" fill="#0F172A">No module bounds</text>
  </g>
</svg>
"""

PAGE_TWO = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720"
     width="1280" height="720" font-family="Arial" font-size="16">
  <rect width="1280" height="720" fill="#FFFFFF"/>
  <g id="body-slot" data-pptx-bounds="80 60 900 120">
    <text id="body" x="80" y="120" font-size="24" fill="#334155">Second page body</text>
  </g>
</svg>
"""


def _run_cli(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SCRIPT), *args],
        capture_output=True,
        text=True,
        check=False,
    )


def _by_selector(records: list[dict]) -> dict[str, dict]:
    return {record['selector']: record for record in records}


class SvgTextInspectTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.page_one = self.root / '01_cover.svg'
        self.page_two = self.root / '02_body.svg'
        self.page_one.write_text(PAGE_ONE, encoding='utf-8')
        self.page_two.write_text(PAGE_TWO, encoding='utf-8')
        self.addCleanup(self._tmp.cleanup)

    def _records(self) -> dict[str, dict]:
        return _by_selector(inspect_svg(self.page_one))

    def test_text_is_associated_with_its_owning_module(self) -> None:
        records = self._records()
        module = records['#title']['module']
        self.assertEqual(module['label'], 'title-slot data-pptx-bounds')
        self.assertEqual(
            (module['x'], module['y'], module['width'], module['height']),
            (80.0, 60.0, 600.0, 80.0),
        )
        self.assertEqual(records['#title']['text'], 'Short title')
        self.assertEqual(records['#title']['text_anchor'], 'start')
        self.assertEqual(records['#title']['fill'], '#0F172A')
        self.assertFalse(records['#title']['hidden'])
        # Each text takes its own module, never a neighbour's.
        self.assertEqual(
            records['#overflowing']['module']['label'],
            'narrow-slot data-pptx-bounds',
        )

    def test_overflow_is_detected_per_container(self) -> None:
        records = self._records()

        self.assertTrue(records['#title']['module_overflow']['fits'])
        self.assertTrue(records['#title']['canvas_overflow']['fits'])

        overflowing = records['#overflowing']
        self.assertFalse(overflowing['module_overflow']['fits'])
        self.assertEqual(overflowing['module_overflow']['axes'], 'horizontal')
        self.assertTrue(overflowing['canvas_overflow']['fits'])
        self.assertGreater(overflowing['width'], overflowing['module']['width'])

        # Leaving the canvas is a separate, always-blocking finding.
        self.assertFalse(records['#offcanvas']['canvas_overflow']['fits'])

    def test_overflowing_only_filters_out_fitting_text(self) -> None:
        result = _run_cli(str(self.page_one), '--overflowing-only', '--json')
        self.assertEqual(result.returncode, 0, result.stderr)
        selectors = {record['selector'] for record in json.loads(result.stdout)}
        self.assertEqual(selectors, {'#overflowing', '#offcanvas'})

    def test_scope_option_limits_output_to_one_element(self) -> None:
        result = _run_cli(str(self.page_one), '--scope', 'title-slot', '--json')
        self.assertEqual(result.returncode, 0, result.stderr)
        records = json.loads(result.stdout)
        self.assertEqual([record['selector'] for record in records], ['#title'])

        empty = _run_cli(str(self.page_one), '--scope', 'absent-slot', '--json')
        self.assertEqual(empty.returncode, 0, empty.stderr)
        self.assertEqual(json.loads(empty.stdout), [])

    def test_json_output_carries_the_documented_fields(self) -> None:
        result = _run_cli(str(self.page_one), '--json')
        self.assertEqual(result.returncode, 0, result.stderr)
        records = json.loads(result.stdout)
        self.assertEqual(len(records), 4)
        for record in records:
            self.assertEqual(
                {
                    'file', 'selector', 'text', 'hidden', 'font_family',
                    'font_weight', 'fill', 'text_anchor', 'lines', 'width',
                    'bounds', 'module', 'module_overflow', 'canvas_overflow',
                    'frame',
                },
                set(record),
            )
        title = _by_selector(records)['#title']
        self.assertEqual(title['file'], '01_cover.svg')
        self.assertEqual(title['lines'][0]['font_size'], 40.0)
        self.assertEqual(title['lines'][0]['x'], 80.0)

    def test_missing_module_bounds_and_inferred_frame_are_reported(self) -> None:
        record = self._records()['#unbounded']
        # No root <g> data-pptx-bounds, so the gate has no module to compare.
        self.assertIsNone(record['module'])
        self.assertIsNone(record['module_overflow'])
        self.assertTrue(record['canvas_overflow']['fits'])
        # data-pptx-frame is absent too, so the round-trip container is only
        # inferable from the sibling rect and must say so.
        self.assertTrue(record['frame']['inferred'])
        self.assertEqual(
            (record['frame']['x'], record['frame']['width']), (80.0, 400.0),
        )

    def test_directory_target_scans_every_slide(self) -> None:
        result = _run_cli(str(self.root), '--json')
        self.assertEqual(result.returncode, 0, result.stderr)
        records = json.loads(result.stdout)
        self.assertEqual(
            [record['file'] for record in records],
            ['01_cover.svg'] * 4 + ['02_body.svg'],
        )
        self.assertEqual(_by_selector(records)['#body']['text'], 'Second page body')


if __name__ == '__main__':
    unittest.main()
