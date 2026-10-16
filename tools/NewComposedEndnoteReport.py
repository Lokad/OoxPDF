"""Mechanically replace the held-out composed-field report's footnote with an endnote."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json
from DocxFixtureNotes import replace_footnotes_with_endnotes, w

root = Path(__file__).resolve().parent.parent
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
output = cases / 'composed-endnotes-report.docx'
with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
    for name,data in sorted(entries.items()):
        info = ZipInfo(name, (2026,10,16,0,0,0))
        info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
manifest = [baseline | dict(Id=output.stem, Input=output.relative_to(root).as_posix(),
    Sha256=hashlib.sha256(output.read_bytes()).hexdigest().upper(),
    Features=[item.replace('footnote', 'endnote') for item in baseline['Features']])]
(cases / 'composed-endnotes-report.json').write_text(json.dumps(manifest, indent=2)+'\n', encoding='utf-8')
print('Pinned one held-out endnote report; prose, six equations, media and unrelated package parts preserved.')
