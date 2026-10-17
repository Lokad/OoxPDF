"""Independent numbering color, underline and revision-owner controls."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
source = cases / 'list-family-run-arial.docx'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = []

def decorate(properties, color, underline):
    E.SubElement(properties, w + 'color', {w + 'val': color})
    E.SubElement(properties, w + 'u', {w + 'val': underline})

revision = {w + 'id': '310', w + 'author': 'Public Reviewer', w + 'date': '2026-10-17T00:00:00Z'}
for name in ['run-blue-double', 'mark-red-single', 'level-green-none', 'normal-purple-single',
             'run-blue-double-inserted', 'mark-inserted', 'level-green-double']:
    with ZipFile(source) as z:
        entries = {part: z.read(part) for part in z.namelist()}
    footer = E.fromstring(entries['word/footer1.xml'])
    paragraph = footer.find(w + 'p')
    first = paragraph.find(w + 'r')
    decorate(first.find(w + 'rPr'), '1964AE', 'double')
    if name in ['mark-red-single', 'level-green-none', 'level-green-double']:
        decorate(E.SubElement(paragraph.find(w + 'pPr'), w + 'rPr'), 'AA2200', 'single')
    if name.startswith('level-green-'):
        numbering = E.fromstring(entries['word/numbering.xml'])
        decorate(E.SubElement(numbering.find(w + 'abstractNum/' + w + 'lvl'), w + 'rPr'), '008844', 'double' if name.endswith('double') else 'none')
        entries['word/numbering.xml'] = E.tostring(numbering, encoding='utf-8', xml_declaration=True)
    if name == 'normal-purple-single':
        styles = E.fromstring(entries['word/styles.xml'])
        normal = E.SubElement(styles, w + 'style', {w + 'type': 'paragraph', w + 'styleId': 'Normal', w + 'default': '1'})
        decorate(E.SubElement(normal, w + 'rPr'), '8822AA', 'single')
        entries['word/styles.xml'] = E.tostring(styles, encoding='utf-8', xml_declaration=True)
    if name == 'run-blue-double-inserted':
        index = list(paragraph).index(first)
        paragraph.remove(first)
        insertion = E.Element(w + 'ins', revision)
        insertion.append(first)
        paragraph.insert(index, insertion)
    if name == 'mark-inserted':
        E.SubElement(E.SubElement(paragraph.find(w + 'pPr'), w + 'rPr'), w + 'ins', revision)
    entries['word/footer1.xml'] = E.tostring(footer, encoding='utf-8', xml_declaration=True)
    target = cases / ('list-decoration-' + name + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as z:
        for part, data in entries.items():
            info = ZipInfo(part, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=1, HeldOut=False, ExpectedText=['share the same input series.']))
(cases / 'list-decorations.json').write_text(json.dumps(rows, indent=2)+'\n', encoding='utf-8')
print('Prepared',len(rows),'public decoration owner controls')
