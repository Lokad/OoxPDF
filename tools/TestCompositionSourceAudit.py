"""Adversarial controls for cached-result source auditing; no rendering/builds."""
import sys
import json
import subprocess
from zipfile import ZipFile
sys.dont_write_bytecode = True
from xml.etree import ElementTree as E
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory, gettempdir
from AuditCompositionContent import inspect_reference, unresolved_fields, validate_outcome_modes, resolve_reference, sha, operation_text, words

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
        self.assertEqual([], unresolved_fields(document('<w:p><w:fldSimple w:instr=" REF Target "><w:r><w:t>North</w:t></w:r></w:fldSimple></w:p>')))

    def test_packed_and_split_complex_results(self):
        for packed in [False, True]:
            self.assertEqual([], unresolved_fields(document('<w:p>'+complex_field(packed=packed)+'</w:p>')))

    def test_nested_cached_results(self):
        inner = complex_field(result='North')
        outer = complex_field(result='PREFIX').replace('<w:t>PREFIX</w:t>', '<w:t>South</w:t>')
        outer = outer.replace('<w:r><w:fldChar w:fldCharType="end"/></w:r>', inner+'<w:r><w:fldChar w:fldCharType="end"/></w:r>')
        self.assertEqual([], unresolved_fields(document('<w:p>'+outer+'</w:p>')))

    def test_dynamic_fields_stay_unresolved(self):
        for opcode in ['PAGE', 'NUMPAGES', 'page \\* MERGEFORMAT']:
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
            with TemporaryDirectory() as temporary:
                root = Path(temporary)
                pdf, inspector = root / "reference.pdf", root / "inspector.dll"
                pdf.write_bytes(b"pinned reference")
                inspector.write_bytes(b"pinned inspector")
                cache = root / "cache"
                folder = cache / (sha(pdf) + "-" + sha(inspector)[:12])
                folder.mkdir(parents=True)
                operations = folder / "text-operations.json"
                operations.write_text("[]", encoding="utf-8")
                (folder / "identity.json").write_text(json.dumps({"PdfSha256": sha(pdf),
                    "InspectorSha256": sha(inspector), "OperationsSha256": sha(operations)}), encoding="utf-8")
                log = folder / "inspect.log"
                log.write_text("2 0 obj: pages\n" + "".join(f"{3 + index} 0 obj: page\n" for index in range(page_count)), encoding="utf-8")
                observed, operations_hash, pages, log_hash = inspect_reference(pdf, inspector, cache)
                self.assertEqual([], observed)
                self.assertEqual(sha(operations), operations_hash)
                self.assertEqual(page_count, pages)
                self.assertEqual(sha(log), log_hash)
                log.write_text("2 0 obj: pages\n", encoding="utf-8")
                with self.assertRaises(ValueError):
                    inspect_reference(pdf, inspector, cache)

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
                (run / 'inspect/text-operations.json').write_text('[]', encoding='utf-8')
                reference = references / 'control-presentation/reference.pdf'
                reference.parent.mkdir()
                reference.write_bytes(pdf.read_bytes())
                cache = root / 'cache' / (sha(reference) + '-' + sha(inspector)[:12])
                cache.mkdir(parents=True)
                ops = cache / 'text-operations.json'
                ops.write_text('[]', encoding='utf-8')
                (cache / 'identity.json').write_text(json.dumps(dict(PdfSha256=sha(reference), InspectorSha256=sha(inspector),
                                                                    OperationsSha256=sha(ops))), encoding='utf-8')
                (cache / 'inspect.log').write_text('2 0 obj: pages\n3 0 obj: page\n', encoding='utf-8')
                row = dict(Id='control', Mode='presentation', InputSha256=sha(source), PdfSha256=sha(pdf),
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
                # Never borrow another installed target when the selected asset changes.
                library.write_bytes(b'wrong-selected-asset')
                result = subprocess.run(arguments, capture_output=True, text=True)
                self.assertNotEqual(0, result.returncode)
                self.assertIn('Candidate/frozen library identity differs', result.stderr)

if __name__ == '__main__':
    unittest.main()
