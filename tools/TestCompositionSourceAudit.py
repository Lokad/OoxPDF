"""Adversarial controls for cached-result source auditing; no rendering/builds."""
import sys
sys.dont_write_bytecode = True
from xml.etree import ElementTree as E
import unittest
from AuditCompositionContent import unresolved_fields, validate_outcome_modes

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def document(content):
    return E.fromstring(f'<w:document xmlns:w="{W}"><w:body>{content}</w:body></w:document>')

def complex_field(instruction=' REF Target ', result='North', packed=False):
    children = ['<w:fldChar w:fldCharType="begin"/>', f'<w:instrText>{instruction}</w:instrText>',
                '<w:fldChar w:fldCharType="separate"/>', f'<w:t>{result}</w:t>', '<w:fldChar w:fldCharType="end"/>']
    return '<w:r>'+''.join(children)+'</w:r>' if packed else ''.join('<w:r>'+child+'</w:r>' for child in children)

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

if __name__ == '__main__':
    unittest.main()
