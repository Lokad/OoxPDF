"""Mechanically replace the held-out composed-field report's footnote with an endnote."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
from copy import deepcopy
import argparse, hashlib, json
from DocxFixtureNotes import replace_footnotes_with_endnotes, w

root = Path(__file__).resolve().parent.parent
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--include-math', action='store_true', help='Add two supported fractions inside the endnote')
args = parser.parse_args()
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/Composition'
baseline = json.loads((cases / 'composed-fields-report.json').read_text(encoding='utf-8'))[0]
source = root / baseline['Input']
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == baseline['Sha256']
with ZipFile(source) as archive:
    original = {name:archive.read(name) for name in archive.namelist()}
entries = replace_footnotes_with_endnotes(original)
before = E.fromstring(original['word/document.xml'])
after = E.fromstring(entries['word/document.xml'])
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
assert ''.join(node.text or '' for node in before.iter(w + 't')) == ''.join(node.text or '' for node in after.iter(w + 't'))
assert [E.tostring(node) for node in before.iter(m + 'oMath')] == [E.tostring(node) for node in after.iter(m + 'oMath')]
assert sum(len(list(E.fromstring(data).iter(m + 'oMath'))) for name,data in entries.items()
           if name.startswith('word/') and name.endswith('.xml')) == baseline['ExpectedEquationCount']
changed = {'word/document.xml','word/footnotes.xml','word/endnotes.xml','word/_rels/document.xml.rels','[Content_Types].xml'}
assert all(entries[name] == data for name,data in original.items() if name not in changed)
assert all(entries[name] == data for name,data in original.items() if name.startswith('word/media/'))
if args.include_math:
    notes = E.fromstring(entries['word/endnotes.xml'])
    note = next(node for node in notes if int(node.get(w + 'id', '0')) > 0)
    paragraph = note.find(w + 'p')
    def text(value):
        run = E.SubElement(paragraph, w + 'r')
        E.SubElement(run, w + 't', {'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text = value
    text(' Calculation: ')
    E.SubElement(paragraph, w + 'bookmarkStart', {w + 'id':'55003', w + 'name':'NoteCalculation'})
    paragraph.append(deepcopy(next(after.iter(m + 'oMath'))))
    E.SubElement(paragraph, w + 'bookmarkEnd', {w + 'id':'55003'})
    text(' and ')
    paragraph.append(deepcopy(next(after.iter(m + 'oMath'))))
    entries['word/endnotes.xml'] = E.tostring(notes, encoding='utf-8', xml_declaration=True)
output = cases / ('composed-note-math-report.docx' if args.include_math else 'composed-endnotes-report.docx')
with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
    for name,data in sorted(entries.items()):
        info = ZipInfo(name, (2026,10,16,0,0,0))
        info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
manifest = [baseline | dict(Id=output.stem, Input=output.relative_to(root).as_posix(),
    Sha256=hashlib.sha256(output.read_bytes()).hexdigest().upper(),
    ExpectedEquationCount=baseline['ExpectedEquationCount'] + (2 if args.include_math else 0),
    ExpectedText=baseline['ExpectedText'] + (['Calculation:'] if args.include_math else []),
    Features=[item.replace('footnote', 'endnote') for item in baseline['Features']] +
        (['two supported fractions and a bookmark inside the endnote'] if args.include_math else []))]
(cases / (output.stem + '.json')).write_text(json.dumps(manifest, indent=2)+'\n', encoding='utf-8')
print('Pinned one held-out endnote report; existing prose, six equations, media and unrelated package parts preserved' +
      ('; two note fractions added.' if args.include_math else '.'))
