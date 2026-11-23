"""Source audit controls, with optional native picture checks on small synthetic PDFs."""
import sys
import json
import subprocess
from collections import Counter
from zipfile import ZipFile
sys.dont_write_bytecode = True
from xml.etree import ElementTree as E
import unittest
from unittest.mock import patch
from pathlib import Path
from tempfile import TemporaryDirectory, gettempdir
from AuditCompositionContent import (inspect_reference, operation_text, reference_simple_math, resolve_reference,
                                     screen_picture_ink, sha, unresolved_fields, validate_outcome_modes, words)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

class OperationTextTests(unittest.TestCase):
    def operation(self, text='r', x=108.226, y=73.499, advance=2.67732, **changes):
        item = dict(PageNumber=1, ObjectNumber=4, Generation=0, Font='Eq2', FillColor='g:0',
                    FontSize=8.04, CharacterSpacing=0, EffectiveA=1, EffectiveB=0, EffectiveC=0, EffectiveD=1,
                    EffectiveX=x, EffectiveY=y, EmittedAdvancePoints=advance, DecodedText=text)
        return item | changes

    def test_individually_positioned_glyphs_and_split_runs_are_words(self):
        items = [self.operation(), self.operation('a', 110.904, advance=4.47024),
                 self.operation('t', 115.375, advance=2.23512), self.operation('e', 117.609, advance=4.47024)]
        self.assertEqual({'rate': 1}, words(operation_text(items)))
        self.assertEqual('rate', operation_text([self.operation('ra', advance=7.148), self.operation('te', 115.374)]))

    def test_whole_words_explicit_spaces_and_punctuation_survive(self):
        items = [self.operation('rate ', advance=20), self.operation('time', 128.226)]
        self.assertEqual({'rate': 1, 'time': 1}, words(operation_text(items)))
        self.assertEqual({'rate': 1, 'time': 1}, words(operation_text([
            self.operation('rate,', advance=20), self.operation('time', 128.226)])))
        self.assertEqual("rate's", operation_text([
            self.operation('rate', advance=20), self.operation("'s", 128.226)]))

    def test_word_gaps_columns_baselines_overlap_and_paint_order_are_boundaries(self):
        left = self.operation('rate', advance=20)
        for changes in [dict(x=130), dict(x=300), dict(y=74), dict(x=120), dict(x=90)]:
            self.assertEqual({'rate': 1, 'time': 1}, words(operation_text([left, self.operation('time', **changes)])))
        self.assertEqual('time rate', operation_text([self.operation('time', x=130), left]))

    def test_resource_text_state_and_page_changes_are_boundaries(self):
        left = self.operation()
        for key, value in dict(PageNumber=2, ObjectNumber=5, Generation=1, Font='Other', FillColor='g:1',
                               FontSize=9, CharacterSpacing=1, EffectiveA=2, EffectiveB=1,
                               EffectiveC=1, EffectiveD=2).items():
            with self.subTest(key=key):
                self.assertEqual('r ate', operation_text([left, self.operation('ate', x=110.90332, **{key:value})]))

    def test_rotated_scaled_and_reflected_effective_frames(self):
        for a,b,c,d in [(0,1,-1,0), (2,0,0,2), (-1,0,0,1)]:
            frame = dict(EffectiveA=a, EffectiveB=b, EffectiveC=c, EffectiveD=d)
            left = self.operation(**frame)
            right = self.operation('ate', x=108.226+2.67732*a, y=73.499+2.67732*b, **frame)
            self.assertEqual('rate', operation_text([left, right]))

    def test_missing_invalid_zero_and_nonfinite_metrics_stay_separate(self):
        for key in ['FontSize','CharacterSpacing','EffectiveA','EffectiveB','EffectiveC','EffectiveD',
                    'EffectiveX','EffectiveY','EmittedAdvancePoints']:
            for value in [None, '1', True, float('nan'), float('inf'), -float('inf')]:
                for side in [0,1]:
                    items=[self.operation(), self.operation('ate', x=110.90332)]
                    items[side][key]=value
                    self.assertEqual('r ate', operation_text(items), (key,value,side))
        for key in ['PageNumber','ObjectNumber','Generation','Font','FillColor']:
            items=[self.operation(),self.operation('ate', x=110.90332)]
            for item in items: item.pop(key)
            self.assertEqual('r ate', operation_text(items))
        for changes in [dict(EmittedAdvancePoints=0),dict(EmittedAdvancePoints=-1),dict(FontSize=0),
                        dict(EffectiveA=0,EffectiveD=0)]:
            self.assertEqual('r ate', operation_text([self.operation(**changes),self.operation('ate',x=110.90332)]))

    def test_replacement_text_and_empty_assembly_continuations(self):
        self.assertEqual('root', operation_text([self.operation('root')]))
        self.assertEqual('r  ate', operation_text([
            self.operation(), self.operation('', x=110.90332), self.operation('ate', x=110.90332)]))
        self.assertEqual('r ate', operation_text([self.operation(),self.operation('ate',x=110.90332,EmittedAdvancePoints=None)]))

    def test_rounding_tolerance_is_small_and_geometry_only(self):
        left=self.operation()
        self.assertEqual('rate', operation_text([left,self.operation('ate',x=110.90532)]))
        self.assertEqual('r ate', operation_text([left,self.operation('ate',x=110.90732)]))
        self.assertEqual('r ate', operation_text([left,self.operation('ate',x=110.90332,y=73.503)]))

    def test_reference_math_confirms_a_source_subset_with_separate_operators(self):
        spans = [self.operation('𝑎', x=110), self.operation('+', x=115), self.operation('𝑏', x=120)]
        expected = Counter({'a+b': 3, 'x+y': 1})
        second_page = [span | {'PageNumber': 2} for span in spans]
        self.assertEqual({'a+b': 2}, reference_simple_math(spans + second_page, expected))
        self.assertEqual({'a+b': 1}, reference_simple_math(spans + second_page, Counter({'a+b': 1})))
        self.assertEqual({}, reference_simple_math(spans, Counter({'a+c': 1})))

    def test_reference_math_does_not_join_unrelated_results(self):
        spans = [self.operation('𝑎', x=110), self.operation('+', x=115), self.operation('𝑏', x=120)]
        expected = Counter({'a+b': 1})
        for change in ({'EffectiveX': 150}, {'EffectiveX': 114}, {'EffectiveY': 78}, {'Font': 'F2'},
                       {'PageNumber': 2}, {'ObjectNumber': 5}, {'EffectiveB': 1}, {'DecodedText': ''}):
            with self.subTest(change=change):
                self.assertEqual({}, reference_simple_math(spans[:2] + [spans[2] | change], expected))
        self.assertEqual({}, reference_simple_math([self.operation('a+b')], expected))


class PictureInkTests(unittest.TestCase):
    def setUp(self):
        try:
            import pymupdf
        except ModuleNotFoundError:
            self.skipTest('PyMuPDF is needed for picture witnesses')
        self.fitz = pymupdf
        temporary = TemporaryDirectory(prefix='ooxpdf-picture-audit-')
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.source = self.root / 'source.docx'
        self.png = pymupdf.Pixmap(pymupdf.csRGB, 2, 1, bytes([47, 128, 237, 39, 174, 96]), False).tobytes('png')

    def write_source(self, svg=False):
        with ZipFile(self.source, 'w') as archive:
            archive.writestr('word/media/picture.png', self.png)
            if svg:
                archive.writestr('word/media/drawing.svg', '<svg/>')
                archive.writestr('word/media/drawing.png', self.png)

    def write_pdf(self, name, pictures):
        path = self.root / (name + '.pdf')
        with self.fitz.open() as doc:
            for kind in pictures:
                page = doc.new_page(width=100, height=100)
                rect = self.fitz.Rect(10, 10, 70, 40)
                if kind in ('png', 'clipped', 'covered', 'corner-overlap', 'rotated'):
                    page.insert_image(rect, stream=self.png)
                elif kind == 'svg':
                    page.draw_rect(rect, color=None, fill=(0., 170 / 255, 102 / 255))
                if kind == 'clipped':
                    xref, = page.get_contents()
                    doc.update_stream(xref, b'q 80 0 10 10 re W n\n' + doc.xref_stream(xref) + b'\nQ')
                if kind in ('covered', 'cover-alone'):
                    page.draw_rect(rect, color=None, fill=(1, 1, 1))
                    for top, fill in [(12, (47, 128, 237)), (25, (39, 174, 96))]:
                        page.draw_rect(self.fitz.Rect(12, top, 18, top + 6), color=None,
                                       fill=tuple(channel / 255 for channel in fill))
                if kind == 'corner-overlap':
                    page.draw_rect(self.fitz.Rect(10, 10, 20, 20), color=None, fill=(1, 1, 1))
                if kind == 'rotated':
                    page.set_rotation(90)
            doc.save(path)
        return path

    def test_removed_and_fully_clipped_pngs_lose_their_visible_witness(self):
        self.write_source()
        expected = {'ExpectedInlineImageCount': 1}
        reference = self.write_pdf('office', ['png'])
        for kind, status in [('png', 'Picture witnesses present'), ('none', 'Missing picture ink'),
                             ('clipped', 'Missing picture ink')]:
            with self.subTest(kind=kind):
                candidate = self.write_pdf(kind, [kind])
                if kind == 'clipped':
                    with self.fitz.open(candidate) as pdf:
                        self.assertEqual(1, len(pdf[0].get_images()))
                result = screen_picture_ink(self.source, candidate, reference, expected, self.fitz)
                self.assertEqual(status, result['Status'])

    def test_foreground_paint_cannot_supply_png_witness_colors(self):
        self.write_source()
        expected = {'ExpectedInlineImageCount': 1}
        reference = self.write_pdf('office', ['png'])
        covered = self.write_pdf('covered', ['covered'])
        cover_alone = self.write_pdf('cover-alone', ['cover-alone'])
        with self.fitz.open(covered) as candidate, self.fitz.open(cover_alone) as without_picture:
            self.assertEqual(1, len(candidate[0].get_image_info()))
            self.assertEqual(candidate[0].get_pixmap().samples, without_picture[0].get_pixmap().samples)
        result = screen_picture_ink(self.source, covered, reference, expected, self.fitz)
        self.assertEqual('Unverified candidate picture witnesses', result['Status'])
        self.assertEqual({}, result['CandidateByPage'])
        self.assertEqual({1: 1}, result['UnverifiedCandidatePngsByPage'])
        self.assertEqual('Missing picture ink',
                         screen_picture_ink(self.source, cover_alone, reference, expected, self.fitz)['Status'])
        self.assertEqual('Unverified Office picture witnesses',
                         screen_picture_ink(self.source, reference, covered, expected, self.fitz)['Status'])
        for kind in ('corner-overlap', 'rotated'):
            with self.subTest(kind=kind):
                candidate = self.write_pdf(kind, [kind])
                self.assertEqual('Picture witnesses present',
                                 screen_picture_ink(self.source, candidate, reference, expected, self.fitz)['Status'])

    def test_repeated_static_svg_witnesses_follow_their_reference_pages(self):
        self.write_source(svg=True)
        expected = {'ExpectedInlineImageCount': 2, 'SvgWitnessColor': [0, 170, 102]}
        reference = self.write_pdf('office', ['png', 'svg', 'svg', 'none'])
        for name, pictures, status in [
            ('same', ['png', 'svg', 'svg'], 'Picture witnesses present'),
            ('omit-copy', ['png', 'svg', 'none'], 'Missing picture ink'),
            ('moved-copy', ['png', 'svg', 'none', 'svg'], 'Picture page counts differ'),
        ]:
            with self.subTest(name=name):
                candidate = self.write_pdf(name, pictures)
                result = screen_picture_ink(self.source, candidate, reference, expected, self.fitz)
                self.assertEqual(status, result['Status'])


def document(content):
    return E.fromstring(f'<w:document xmlns:w="{W}"><w:body>{content}</w:body></w:document>')

def complex_field(instruction=' REF Target ', result='North', packed=False):
    children = ['<w:fldChar w:fldCharType="begin"/>', f'<w:instrText>{instruction}</w:instrText>',
                '<w:fldChar w:fldCharType="separate"/>', f'<w:t>{result}</w:t>', '<w:fldChar w:fldCharType="end"/>']
    return '<w:r>'+''.join(children)+'</w:r>' if packed else ''.join('<w:r>'+child+'</w:r>' for child in children)

class ReferenceResolutionTests(unittest.TestCase):
    def setUp(self):
        self.temporary = TemporaryDirectory(prefix='ooxpdf-audit-reference-')
        self.root = Path(self.temporary.name).resolve()
        self.assertEqual(Path(gettempdir()).resolve(), self.root.parent)
        self.assertTrue(self.root.name.startswith('ooxpdf-audit-reference-'))
        self.addCleanup(self.temporary.cleanup)
        self.first, self.second = self.root / 'first', self.root / 'second'
        self.first.mkdir()
        self.second.mkdir()

    def pdf(self, root, identifier, content=b'Pinned reference'):
        folder = root / (identifier + '-final-preserve')
        folder.mkdir(exist_ok=True)
        pdf = folder / 'reference.pdf'
        pdf.write_bytes(content)
        return pdf

    def test_single_multiple_and_identical_cache_entries(self):
        first = self.pdf(self.first, 'report')
        second = self.pdf(self.second, 'specification', b'Other pinned reference')
        self.assertEqual(first, resolve_reference([self.first], 'report', 'final-preserve', sha(first)))
        roots = [self.first, self.second]
        self.assertEqual(first, resolve_reference(roots, 'report', 'final-preserve', sha(first)))
        self.assertEqual(second, resolve_reference(roots, 'specification', 'final-preserve', sha(second)))
        duplicate = self.pdf(self.second, 'report')
        self.assertEqual(first, resolve_reference(roots, 'report', 'final-preserve', sha(first)))
        self.assertEqual(duplicate, resolve_reference(list(reversed(roots)), 'report', 'final-preserve', sha(first)))

    def test_missing_and_conflicting_entries_cannot_hide_beside_valid_cache(self):
        first = self.pdf(self.first, 'report')
        roots = [self.first, self.second]
        with self.assertRaisesRegex(ValueError, 'Reference PDF missing'):
            resolve_reference(roots, 'missing', 'final-preserve', sha(first))
        (self.second / 'report-final-preserve').mkdir()
        with self.assertRaisesRegex(ValueError, 'Reference PDF missing'):
            resolve_reference(roots, 'report', 'final-preserve', sha(first))
        self.pdf(self.second, 'report', b'Stale reference')
        with self.assertRaisesRegex(ValueError, 'Reference PDF identity differs'):
            resolve_reference(roots, 'report', 'final-preserve', sha(first))
        with self.assertRaisesRegex(ValueError, 'Reference PDF identity differs'):
            resolve_reference(list(reversed(roots)), 'report', 'final-preserve', sha(first))


class CachedFieldAuditTests(unittest.TestCase):
    def test_selected_presets_preserve_presentation_and_default_both(self):
        cases = {'doc': {'Id': 'doc', 'Input': 'report.docx'},
                 'deck': {'Id': 'deck', 'Input': 'briefing.pptx'}}
        for preset, modes in [('Both', ['final-preserve', 'all-word-compatible']),
                              ('Final', ['final-preserve']), ('Review', ['all-word-compatible'])]:
            rows = [{'Id': 'doc', 'Mode': mode} for mode in modes] + [{'Id': 'deck', 'Mode': 'presentation'}]
            validate_outcome_modes(cases, {'DocxPreset': preset, 'Rows': rows}, {'Rows': rows})
            if preset == 'Both':
                validate_outcome_modes(cases, {'Rows': rows}, {'Rows': rows})

    def test_preset_selection_does_not_accept_missing_extra_or_duplicate_rows(self):
        cases = {'doc': {'Id': 'doc', 'Input': 'report.docx'}}
        review = {'Id': 'doc', 'Mode': 'all-word-compatible'}
        final = {'Id': 'doc', 'Mode': 'final-preserve'}
        for actual, compared in [([], [review]), ([review], []), ([review, review], [review]),
                                 ([review], [review, review]), ([review, final], [review]),
                                 ([review], [review, final])]:
            with self.assertRaises(ValueError):
                validate_outcome_modes(cases, {'DocxPreset': 'Review', 'Rows': actual}, {'Rows': compared})
        with self.assertRaises(ValueError):
            validate_outcome_modes(cases, {'DocxPreset': 'Unknown', 'Rows': []}, {'Rows': []})

    def test_nested_paragraphs_without_fields_are_unchanged(self):
        self.assertEqual([], unresolved_fields(document('<w:p><w:r><w:t>Outer</w:t><w:txbxContent><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:txbxContent></w:r></w:p>')))

    def test_literal_page_text_is_ordinary_prose(self):
        self.assertEqual([], unresolved_fields(document('<w:p><w:r><w:t>PAGE NUMPAGES REF Target</w:t></w:r></w:p>')))

    def test_simple_cached_result(self):
        for opcode in ["REF Target", "PAGE2", "NUMPAGES_ARCHIVE"]:
            self.assertEqual([], unresolved_fields(document(
                f'<w:p><w:fldSimple w:instr=" {opcode} "><w:r><w:t>North</w:t></w:r></w:fldSimple></w:p>')))

    def test_packed_and_split_complex_results(self):
        for packed in [False, True]:
            self.assertEqual([], unresolved_fields(document('<w:p>'+complex_field(packed=packed)+'</w:p>')))
        for opcode in ["PAGE2", "NUMPAGES_ARCHIVE"]:
            self.assertEqual([], unresolved_fields(document('<w:p>'+complex_field(instruction=opcode)+'</w:p>')))

    def test_nested_cached_results(self):
        inner = complex_field(result='North')
        outer = complex_field(result='PREFIX').replace('<w:t>PREFIX</w:t>', '<w:t>South</w:t>')
        outer = outer.replace('<w:r><w:fldChar w:fldCharType="end"/></w:r>', inner+'<w:r><w:fldChar w:fldCharType="end"/></w:r>')
        self.assertEqual([], unresolved_fields(document('<w:p>'+outer+'</w:p>')))

    def test_dynamic_fields_stay_unresolved(self):
        for opcode in ['PAGE', 'NUMPAGES', r'page \* MERGEFORMAT', r'page\* MERGEFORMAT']:
            self.assertTrue(unresolved_fields(document('<w:p>'+complex_field(instruction=opcode)+'</w:p>')))
            self.assertTrue(unresolved_fields(document(f'<w:p><w:fldSimple w:instr="{opcode}"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>')))

    def test_missing_cached_result(self):
        self.assertTrue(unresolved_fields(document('<w:p>'+complex_field(result='')+'</w:p>')))
        self.assertTrue(unresolved_fields(document('<w:p><w:fldSimple w:instr=" REF Target "/></w:p>')))

    def test_missing_end(self):
        content = complex_field().replace('<w:fldChar w:fldCharType="end"/>','')
        self.assertTrue(unresolved_fields(document('<w:p>'+content+'</w:p>')))

    def test_orphan_instruction_and_control(self):
        for content in ['<w:r><w:instrText>REF Target</w:instrText></w:r>', '<w:r><w:fldChar w:fldCharType="end"/></w:r>']:
            self.assertTrue(unresolved_fields(document('<w:p>'+content+'</w:p>')))

    def test_instruction_text_cannot_be_a_cached_result(self):
        content = complex_field().replace('<w:fldChar w:fldCharType="separate"/>','')
        self.assertTrue(unresolved_fields(document('<w:p>'+content+'</w:p>')))

    def test_cross_paragraph_field_stays_unresolved(self):
        content = complex_field().replace('<w:r><w:fldChar w:fldCharType="separate"/></w:r>', '</w:p><w:p><w:r><w:fldChar w:fldCharType="separate"/></w:r>')
        self.assertTrue(unresolved_fields(document('<w:p>'+content+'</w:p>')))

    def test_nested_dynamic_result_stays_unresolved(self):
        content = '<w:fldSimple w:instr=" REF Target ">'+complex_field(instruction='PAGE')+'</w:fldSimple>'
        self.assertTrue(unresolved_fields(document('<w:p>'+content+'</w:p>')))

class ReferenceInspectionTests(unittest.TestCase):
    def test_empty_reference_reports_keep_actual_blank_pages(self):
        for page_count in [1, 2]:
            with self.subTest(page_count=page_count), TemporaryDirectory() as temporary:
                root = Path(temporary)
                pdf, inspector = root / "reference.pdf", root / "inspector.dll"
                pdf.write_bytes(b"pinned reference")
                inspector.write_bytes(b"pinned inspector")
                cache = root / "cache"
                key = sha(pdf) + "-" + sha(inspector)[:12]
                legacy = cache / key
                legacy.mkdir(parents=True)
                legacy_operations = legacy / "text-operations.json"
                legacy_operations.write_text("[]", encoding="utf-8")
                (legacy / "identity.json").write_text(json.dumps({"PdfSha256": sha(pdf),
                    "InspectorSha256": sha(inspector), "OperationsSha256": sha(legacy_operations)}), encoding="utf-8")
                legacy_log = legacy / "inspect.log"
                legacy_log.write_text("9 0 obj: page\n", encoding="utf-8")

                def inspect(argv, *, stdout, stderr, check):
                    self.assertEqual(["dotnet", str(inspector), str(pdf)], argv[:3])
                    self.assertEqual("--text-only", argv[4])
                    self.assertEqual(subprocess.STDOUT, stderr)
                    self.assertTrue(check)
                    (Path(argv[3]) / "text-operations.json").write_text("[]", encoding="utf-8")
                    stdout.write("2 0 obj: pages\n" + "".join(f"{3 + index} 0 obj: page\n" for index in range(page_count)))

                with patch("AuditCompositionContent.subprocess.run", side_effect=inspect) as process:
                    result = inspect_reference(pdf, inspector, cache)
                    self.assertEqual(result, inspect_reference(pdf, inspector, cache))
                    process.assert_called_once()

                folder = cache / ("v2-" + key)
                operations = folder / "text-operations.json"
                log = folder / "inspect.log"
                self.assertEqual(([], sha(operations), page_count, sha(log)), result)
                self.assertEqual("9 0 obj: page\n", legacy_log.read_text(encoding="utf-8"))
                original_log = log.read_bytes()
                log.write_text("9 0 obj: page\n" * (3 - page_count), encoding="utf-8")
                with self.assertRaisesRegex(ValueError, "Reference text cache identity differs"):
                    inspect_reference(pdf, inspector, cache)
                log.write_bytes(original_log)
                operations.write_text('[{"DecodedText":"changed"}]', encoding="utf-8")
                with self.assertRaisesRegex(ValueError, "Reference text cache identity differs"):
                    inspect_reference(pdf, inspector, cache)

    def test_cold_inspection_without_page_evidence_leaves_no_entry(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            pdf, inspector = root / "reference.pdf", root / "inspector.dll"
            pdf.write_bytes(b"pinned reference")
            inspector.write_bytes(b"pinned inspector")
            cache = root / "cache"

            def inspect(argv, *, stdout, **_):
                (Path(argv[3]) / "text-operations.json").write_text("[]", encoding="utf-8")
                stdout.write("2 0 obj: pages\n")

            with patch("AuditCompositionContent.subprocess.run", side_effect=inspect):
                with self.assertRaisesRegex(ValueError, "No inspected reference pages"):
                    inspect_reference(pdf, inspector, cache)
            self.assertEqual([], list(cache.iterdir()))


class AuditFrameworkBindingTests(unittest.TestCase):
    def test_audit_uses_pinned_framework_and_preserves_legacy_identity(self):
        for framework in [None, 'net10.0', 'net11.0']:
            with self.subTest(framework=framework), TemporaryDirectory() as temporary:
                root = Path(temporary)
                artifact, candidate, references = (root / name for name in ['artifact', 'candidate', 'references'])
                for path in [artifact, candidate, references]:
                    path.mkdir()
                inspector = root / 'inspector.dll'
                inspector.write_bytes(b'inspector-control')
                selected = framework or 'net10.0'
                for tfm in ['net10.0', 'net11.0']:
                    library = artifact / 'src/Lokad.OoxPdf/bin/Release' / tfm / 'Lokad.OoxPdf.dll'
                    library.parent.mkdir(parents=True)
                    library.write_bytes(tfm.encode())
                library = artifact / 'src/Lokad.OoxPdf/bin/Release' / selected / 'Lokad.OoxPdf.dll'
                source = artifact / 'tests/Lokad.OoxPdf.Tests/Cases/control.pptx'
                source.parent.mkdir(parents=True)
                with ZipFile(source, 'w') as package:
                    package.writestr('ppt/slides/slide1.xml', '<slide/>')
                manifest_name = 'tests/Lokad.OoxPdf.Tests/Cases/portfolio.json'
                manifest = artifact / manifest_name
                manifest.write_text(json.dumps([dict(Id='control', Input=source.relative_to(artifact).as_posix(),
                                                     Sha256=sha(source), HeldOut=False)]), encoding='utf-8')
                identity = dict(Revision='control', LibrarySha256=sha(library),
                                Inputs={manifest_name: sha(manifest), source.relative_to(artifact).as_posix(): sha(source)})
                if framework is not None:
                    identity['Framework'] = framework
                (artifact / 'identity.json').write_text(json.dumps(identity), encoding='utf-8')
                run = candidate / 'control-presentation'
                (run / 'inspect').mkdir(parents=True)
                pdf = run / 'candidate.pdf'
                pdf.write_bytes(b'%PDF-cached-control')
                operations = run / 'inspect/text-operations.json'
                operations.write_text('[]', encoding='utf-8')
                reference = references / 'control-presentation/reference.pdf'
                reference.parent.mkdir()
                reference.write_bytes(pdf.read_bytes())
                cache = root / 'cache' / ('v2-' + sha(reference) + '-' + sha(inspector)[:12])
                cache.mkdir(parents=True)
                ops = cache / 'text-operations.json'
                ops.write_text('[]', encoding='utf-8')
                log = cache / 'inspect.log'
                log.write_text('2 0 obj: pages\n3 0 obj: page\n', encoding='utf-8')
                (cache / 'identity.json').write_text(json.dumps(dict(PdfSha256=sha(reference), InspectorSha256=sha(inspector),
                                                                    OperationsSha256=sha(ops),
                                                                    InspectionLogSha256=sha(log))), encoding='utf-8')
                row = dict(Id='control', Mode='presentation', InputSha256=sha(source), PdfSha256=sha(pdf),
                           TextOperationsSha256=sha(operations),
                           InspectedPages=1, IgnoredEquationDiagnostics=0)
                (candidate / 'outcomes.json').write_text(json.dumps(dict(LibrarySha256=sha(library), ManifestSha256=sha(manifest),
                                                                       InspectorSha256=sha(inspector), Rows=[row])), encoding='utf-8')
                comparison = dict(Id='control', Mode='presentation', CandidatePdfSha256=sha(pdf), ReferencePdfSha256=sha(reference), Failures=[])
                (candidate / 'office-comparisons.json').write_text(json.dumps(dict(CandidateLibrarySha256=sha(library), Rows=[comparison])), encoding='utf-8')
                output = root / 'result.json'
                arguments = [sys.executable, str(Path(__file__).with_name('AuditCompositionContent.py')),
                             '--artifact', str(artifact), '--candidate', str(candidate), '--inspector', str(inspector),
                             '--references', str(references), '--cache', str(root / 'cache'), '--output', str(output),
                             '--manifest', manifest_name]
                result = subprocess.run(arguments, capture_output=True, text=True)
                self.assertEqual(0, result.returncode, result.stderr)
                report = json.loads(output.read_text(encoding='utf-8'))
                self.assertEqual(selected, report.get('Framework', 'net10.0'))
                self.assertEqual(0, report['LibraryBuilds'])
                if framework == 'net11.0':
                    # A report from another PDF, even with valid operations,
                    # cannot replace a pinned, legitimately empty inspection.
                    arguments[-3] = str(root / 'refused.json')
                    operations.write_text('[{"DecodedText":"Another PDF"}]', encoding='utf-8')
                    result = subprocess.run(arguments, capture_output=True, text=True)
                    self.assertEqual(1, result.returncode)
                    self.assertIn('Candidate text inspection differs', result.stderr)
                    self.assertFalse((root / 'refused.json').exists())
                    outcomes = candidate / 'outcomes.json'
                    legacy = json.loads(outcomes.read_text(encoding='utf-8'))
                    del legacy['Rows'][0]['TextOperationsSha256']
                    outcomes.write_text(json.dumps(legacy), encoding='utf-8')
                    result = subprocess.run(arguments, capture_output=True, text=True)
                    self.assertEqual(1, result.returncode)
                    self.assertIn('reinspect the candidate PDF', result.stderr)
                    self.assertFalse((root / 'refused.json').exists())
                # Never borrow another installed target when the selected asset changes.
                library.write_bytes(b'wrong-selected-asset')
                result = subprocess.run(arguments, capture_output=True, text=True)
                self.assertNotEqual(0, result.returncode)
                self.assertIn('Candidate/frozen library identity differs', result.stderr)

if __name__ == '__main__':
    unittest.main()
