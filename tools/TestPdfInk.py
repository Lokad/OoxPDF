"""Focused validation-tool tests; existing pinned fonts, no library build or Office."""

import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

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
        self.assertNotIn("TraceInkBounds", result)
        traced = analyze_pdf(self.pdf, include_trace_bounds=True)
        self.assertEqual(2, len(traced["TraceInkBounds"]))
        self.assertEqual(set(result["Pages"][0]["FontPrograms"]),
                         {frame["FontProgramSha256"] for frame in traced["TraceInkBounds"]})
        self.assertTrue(all(frame["Ink"][2] > frame["Ink"][0] and frame["Ink"][3] > frame["Ink"][1]
                            for frame in traced["TraceInkBounds"]))

    def test_clipped_glyph_retains_its_exact_program_identity(self):
        write_pdf(self.pdf, b"q\n0 0 5 160 re W n\n")
        result = analyze_pdf(self.pdf)
        self.assertEqual([], result["UnmeasuredTextTraces"])
        self.assertEqual(1, len(result["ClippedTextTraces"]))
        clipped = result["ClippedTextTraces"][0]
        self.assertIn(clipped["FontProgramSha256"], result["Pages"][0]["FontPrograms"])
        self.assertEqual(1, len(clipped["CutGlyphs"]))
        self.assertTrue(clipped["CutGlyphs"][0]["FullyHidden"])

    def test_shared_font_object_keeps_page_local_resource_names(self):
        font = (FONT_DIRECTORY / "DejaVuSans.ttf").read_bytes()
        with fitz.open() as document:
            xrefs = []
            for resource in ("First", "Second"):
                page = document.new_page(width=220, height=160)
                xrefs.append(page.insert_font(fontname=resource, fontbuffer=font))
                page.insert_text((20, 50), "A", fontname=resource, fontsize=24)
            self.assertEqual(xrefs[0], xrefs[1])
            document.save(self.pdf, deflate=True)

        with patch.object(fitz.Document, "extract_font", autospec=True,
                          side_effect=fitz.Document.extract_font) as extract_font:
            result = analyze_pdf(self.pdf, text_witnesses=[(1, "A"), (2, "A")])
        self.assertEqual([xrefs[0]], [call.args[1] for call in extract_font.call_args_list])
        self.assertEqual(2, result["GlyphsMeasured"])
        digest = hashlib.sha256(font).hexdigest().upper()
        self.assertEqual([{digest: ["First"]}, {digest: ["Second"]}],
                         [page["FontResourceBindings"] for page in result["Pages"]])
        self.assertEqual([("UnclippedInk", digest)] * 2,
                         [(row["Status"], row["FontProgramSha256"]) for row in result["TextWitnesses"]])

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

    def test_unembedded_font_remains_unknown(self):
        document = fitz.open()
        page = document.new_page(width=220, height=160)
        page.insert_text((20, 50), "A", fontname="helv", fontsize=24)
        document.save(self.pdf)
        document.close()
        result = analyze_pdf(self.pdf, text_witnesses=[(1, "A")])
        self.assertEqual(0, result["GlyphsMeasured"])
        self.assertEqual("UnavailableEmbeddedTrueTypeProgram",
                         result["UnmeasuredTextTraces"][0]["Reason"])
        self.assertEqual("Unverified", result["TextWitnesses"][0]["Status"])
        self.assertEqual("UnavailableEmbeddedTrueTypeProgram", result["TextWitnesses"][0]["Reason"])

    def test_named_text_witness_requires_ink_for_the_complete_phrase(self):
        for name, render_mode, clip, status in (
                ("visible", 0, None, "UnclippedInk"),
                ("absent", None, None, "NoMatchingPaintSpan"),
                ("invisible", 3, None, "NoMatchingPaintSpan"),
                ("clipped", 0, b"0 0 80 160 re W n", "Clipped")):
            with self.subTest(name=name), fitz.open() as document:
                page = document.new_page(width=220, height=160)
                if render_mode is not None:
                    page.insert_font(fontname="F1", fontfile=FONT_DIRECTORY / "DejaVuSans.ttf")
                    page.insert_text((20, 50), "Source text", fontname="F1", fontsize=20,
                                     render_mode=render_mode)
                    if clip:
                        contents = page.get_contents()[0]
                        document.update_stream(contents, b"q\n" + clip + b"\n" +
                                               document.xref_stream(contents) + b"\nQ\n")
                pdf = self.root / (name + ".pdf")
                document.save(pdf, deflate=True)
                output = self.root / (name + ".json")
                command = [sys.executable, str(ROOT / "tools/ScreenPdfInk.py"), "--pdf", str(pdf),
                           "--expected-pdf-sha256", hashlib.sha256(pdf.read_bytes()).hexdigest(),
                           "--text-witness", "1", "Source text", "--output", str(output)]
                run = subprocess.run(command, capture_output=True, text=True)
                self.assertEqual(0 if name == "visible" else 1, run.returncode, run.stderr)
                matched = json.loads(output.read_text())["Rows"][0]["TextWitnesses"]
                self.assertEqual([status], [match["Status"] for match in matched])
                if name == "visible":
                    result = analyze_pdf(pdf, text_witnesses=[(1, "Source tex"), (2, "Source text")])
                    self.assertEqual(["NoMatchingPaintSpan", "NoMatchingPaintSpan"],
                                     [match["Status"] for match in result["TextWitnesses"]])

    def test_named_witness_in_zero_alpha_group_remains_unverified(self):
        with fitz.open() as source, fitz.open() as document:
            source_page = source.new_page(width=220, height=160)
            source_page.insert_font(fontname="F1", fontfile=FONT_DIRECTORY / "DejaVuSans.ttf")
            source_page.insert_text((20, 50), "Source text", fontname="F1", fontsize=20)
            page = document.new_page(width=220, height=160)
            form = page.show_pdf_page(page.rect, source, 0)
            document.xref_set_key(form, "Group", "<< /S /Transparency /I true /K false >>")
            resources = int(document.xref_get_key(page.xref, "Resources")[1].split()[0])
            document.xref_set_key(resources, "ExtGState", "<< /Hidden << /ca 0 /CA 0 >> >>")
            contents, = page.get_contents()
            document.update_stream(contents, b"q /Hidden gs\n" + document.xref_stream(contents) + b"\nQ\n")
            document.save(self.pdf)
        witness, = analyze_pdf(self.pdf, text_witnesses=[(1, "Source text")])["TextWitnesses"]
        self.assertEqual("Unverified", witness["Status"])
        self.assertEqual("TransparencyGroup", witness["Reason"])

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
