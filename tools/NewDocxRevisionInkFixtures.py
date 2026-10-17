"""Reuse the ordering controls and add a deletion beside an inserted picture."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import json, hashlib
root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/ReviewAuthors'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
E.register_namespace('w', w[1:-1])
selected = ['mark-with-text', 'image-before-text', 'image-after-text', 'format-before-text', 'format-after-text']
rows = [row for row in json.loads((cases / 'nontext-authors.json').read_text()) if row['Kind'] in selected]
source = cases / 'nontext-author-image-before-text.docx'
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == next(row['Sha256'] for row in rows if row['Kind'] == 'image-before-text')
with ZipFile(source) as archive:
    entries = {part: archive.read(part) for part in archive.namelist()}
document = E.fromstring(entries['word/document.xml'])
revision = next(node for node in document.iter(w + 'ins') if node.get(w + 'author') == 'Other Reviewer')
revision.tag = w + 'del'
for text in revision.iter(w + 't'):
    text.tag = w + 'delText'; text.text = 'Other deleted phrase.'
entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
target = cases / 'revision-ink-deletion-after-image.docx'
with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
    for part, data in sorted(entries.items()):
        info = ZipInfo(part, (2026, 10, 17, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
rows.append(dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=0,ExpectedText=['Public owner paragraph.', 'Public comment anchor.', 'Other deleted phrase.']))
(cases / 'revision-text-ink.json').write_text(json.dumps(rows,indent=2)+'\n')
print('Pinned six revision-ink controls; five existing inputs reused')
