"""Exercise source-colored PDF path gates with the existing PyMuPDF validator."""
import argparse
import contextlib
import io
import json
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

import pymupdf
import SummarizeDocxInlineFrameGeometry as geometry


class ColoredPathChecks(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = output_directory
        cls.root.mkdir(parents=True, exist_ok=False)
        cls.candidate_directory = cls.root / 'candidate'
        target = cls.candidate_directory / 'colored-path-probe-final-preserve/candidate.pdf'
        target.parent.mkdir(parents=True)
        with pymupdf.open() as pdf:
            page = pdf.new_page(width=100, height=100)
            page.draw_rect(pymupdf.Rect(10, 20, 40, 60), color=None, fill=(0, 170/255, 102/255))
            page.draw_rect(pymupdf.Rect(50, 20, 70, 40), color=None, fill=(34/255, 85/255, 204/255))
            page.draw_rect(pymupdf.Rect(-5, 80, 5, 90), color=None, fill=(0, 170/255, 102/255))
            pdf.save(target)
        cls.pdf = target
        cls.identity = {'Id': 'colored-path-probe', 'Mode': 'final-preserve',
                        'InputSha256': 'synthetic-source', 'PdfSha256': geometry.sha(target)}
        (cls.candidate_directory / 'outcomes.json').write_text(json.dumps({'Rows': [cls.identity]}))
        cls.number = 0

    def gate(self, expected):
        type(self).number += 1
        golden = self.root / f'golden-{self.number}.json'
        result = self.root / f'report-{self.number}.json'
        row = {**self.identity, 'WordPdfSha256': 'synthetic-reference',
               'WordMathAnchors': [], 'WordFrames': [], 'WordColoredPaths': expected}
        golden.write_text(json.dumps({'Rows': [row]}))
        arguments = ['geometry', '--golden', str(golden), '--candidate', str(self.candidate_directory), '--output', str(result)]
        with patch.object(sys, 'argv', arguments), contextlib.redirect_stdout(io.StringIO()):
            code = geometry.main()
        return code, json.loads(result.read_text())['Rows'][0]

    def test_source_color_and_bounds(self):
        paths = geometry.measure_colored_paths(self.pdf, ['00AA66'])
        self.assertEqual(2, len(paths))
        self.assertEqual([10, 20, 40, 60], paths[0]['Rect'])
        self.assertTrue(paths[0]['InsidePage'])
        self.assertFalse(paths[1]['InsidePage'])

    def test_missing_witness_fails(self):
        code, row = self.gate([{'Page': 1, 'FillHex': 'FF0000', 'Rect': [50, 20, 70, 40]}])
        self.assertEqual(1, code)
        self.assertIn('ColoredPathCountOrSourceOrder', row['Failures'])

    def test_original_one_point_boundary(self):
        code, _ = self.gate([{'Page': 1, 'FillHex': '2255CC', 'Rect': [50.5, 20, 70.5, 40]}])
        self.assertEqual(0, code)
        code, row = self.gate([{'Page': 1, 'FillHex': '2255CC', 'Rect': [51, 20, 71, 40]}])
        self.assertEqual(1, code)
        self.assertIn('ColoredPathGeometry', row['Failures'])

    def test_offpage_path_fails(self):
        code, row = self.gate([{'Page': 1, 'FillHex': '00AA66', 'Rect': [10, 20, 40, 60]},
                               {'Page': 1, 'FillHex': '00AA66', 'Rect': [-5, 80, 5, 90]}])
        self.assertEqual(1, code)
        self.assertIn('OffPageColoredPath', row['Failures'])

    def test_invalid_color_fails_before_reading(self):
        with self.assertRaises(ValueError):
            geometry.measure_colored_paths(self.root / 'absent.pdf', ['GGAA66'])


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-directory', required=True, type=Path)
    args = parser.parse_args()
    output_directory = args.output_directory.resolve()
    unittest.main(argv=[sys.argv[0]])
