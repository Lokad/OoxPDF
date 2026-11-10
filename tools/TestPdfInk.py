"""Focused validation-tool tests; existing pinned fonts, no library build or Office."""

import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

import pymupdf as fitz

from ScreenPdfInk import analyze_pdf


ROOT = Path(__file__).resolve().parent.parent
FONT_DIRECTORY = ROOT / "tests/Lokad.OoxPdf.Tests/Cases/PortableFonts"


def write_pdf(path, prefix=b""):
    document = fitz.open()
    page = document.new_page(width=220, height=160)
    for index, name in enumerate(("DejaVuSans.ttf", "DejaVuMathTeXGyre.ttf")):
        resource = "F" + str(index + 1)
        page.insert_font(fontname=resource, fontbuffer=(FONT_DIRECTORY / name).read_bytes())
        page.insert_text((20 + index * 60, 50), "A", fontname=resource, fontsize=24)
    # Two actual programs deliberately share the names used by native tracing.
    aliases = {}
    for xref in range(1, document.xref_length()):
        for key in ("BaseFont", "FontName"):
            kind, name = document.xref_get_key(xref, key)
            if kind == "name":
                alias = aliases.setdefault(name, "/" + chr(65 + len(aliases)) * 6 + "+SharedFace")
                document.xref_set_key(xref, key, alias)
    if prefix:
        first = page.get_contents()[0]
        document.update_stream(first, prefix + document.xref_stream(first) + b"\nQ\n")
    document.save(path, deflate=True)
    document.close()


class PdfInkTests(unittest.TestCase):
    def setUp(self):
        (ROOT / "artifacts").mkdir(exist_ok=True)
        self.temporary = tempfile.TemporaryDirectory(prefix="pdf-ink-tests-", dir=ROOT / "artifacts")
        self.root = Path(self.temporary.name)
        self.pdf = self.root / "font alias.pdf"

    def tearDown(self):
        self.temporary.cleanup()

    def test_display_name_collision_binds_both_actual_programs(self):
        write_pdf(self.pdf)
        result = analyze_pdf(self.pdf, hashlib.sha256(self.pdf.read_bytes()).hexdigest())
        self.assertEqual(2, result["GlyphsMeasured"])
        self.assertEqual([], result["ClippedTextTraces"])
        self.assertEqual([], result["UnmeasuredTextTraces"])
        self.assertEqual(2, len(result["Pages"][0]["FontPrograms"]))
        self.assertEqual(2, len(result["Pages"][0]["DisplayNameCollisions"]["SharedFace"]))

    def test_clipped_glyph_retains_its_exact_program_identity(self):
        write_pdf(self.pdf, b"q\n0 0 5 160 re W n\n")
        result = analyze_pdf(self.pdf)
        self.assertEqual([], result["UnmeasuredTextTraces"])
        self.assertEqual(1, len(result["ClippedTextTraces"]))
        clipped = result["ClippedTextTraces"][0]
        self.assertIn(clipped["FontProgramSha256"], result["Pages"][0]["FontPrograms"])
        self.assertEqual(1, len(clipped["CutGlyphs"]))
        self.assertTrue(clipped["CutGlyphs"][0]["FullyHidden"])

    def test_quarter_turn_clip_is_qualified(self):
        write_pdf(self.pdf, b"q\n0 1 -1 0 150 0 cm\n0 0 160 160 re W n\n")
        result = analyze_pdf(self.pdf)
        self.assertEqual([], result["UnmeasuredTextTraces"])

    def test_oblique_clip_remains_unknown(self):
        write_pdf(self.pdf, b"q\n0.707107 0.707107 -0.707107 0.707107 100 0 cm\n0 0 160 160 re W n\n")
        result = analyze_pdf(self.pdf)
        self.assertEqual(1, len(result["UnmeasuredTextTraces"]))
        self.assertEqual("UnknownClip", result["UnmeasuredTextTraces"][0]["Reason"])

    def test_pdf_identity_mismatch_refuses_analysis(self):
        write_pdf(self.pdf)
        with self.assertRaisesRegex(ValueError, "PDF hash differs"):
            analyze_pdf(self.pdf, "00" * 32)

    def test_manifest_selection_and_library_identity(self):
        target = self.root / "probe-final-preserve"
        target.mkdir()
        self.pdf = target / "candidate.pdf"
        write_pdf(self.pdf)
        outcomes = self.root / "outcomes.json"
        outcomes.write_text(json.dumps({
            "LibrarySha256": "AA" * 32, "Rows": [{
                "Id": "probe", "Mode": "final-preserve",
                "PdfSha256": hashlib.sha256(self.pdf.read_bytes()).hexdigest().upper()
            }]
        }), encoding="utf-8")
        manifest = self.root / "manifest.json"
        manifest.write_text('[{"Id":"probe"}]', encoding="utf-8")
        output = self.root / "report.json"
        command = [sys.executable, str(ROOT / "tools/ScreenPdfInk.py"), "--outcomes", str(outcomes),
                   "--manifest", str(manifest), "--expected-library-sha256", "AA" * 32,
                   "--output", str(output)]
        successful = subprocess.run(command, capture_output=True, text=True)
        self.assertEqual(0, successful.returncode, successful.stderr)
        self.assertEqual(1, len(json.loads(output.read_text())["Rows"]))
        duplicate = subprocess.run(command, capture_output=True, text=True)
        self.assertNotEqual(0, duplicate.returncode)
        wrong = self.root / "wrong.json"
        command[command.index("AA" * 32)] = "BB" * 32
        command[-1] = str(wrong)
        refused = subprocess.run(command, capture_output=True, text=True)
        self.assertNotEqual(0, refused.returncode)
        self.assertFalse(wrong.exists())


if __name__ == "__main__":
    unittest.main()
