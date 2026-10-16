"""Public mixed-story controls with no declared document typeface or theme."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
source = cases / 'list-family-no-family-default.docx'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
rows = []
for points in [11, 14]:
    with ZipFile(source) as z:
        entries = {part: z.read(part) for part in z.namelist()}
    document = E.fromstring(entries['word/document.xml'])
    body = document.find(w + 'body')
    section = body.find(w + 'sectPr')
    table = E.Element(w + 'tbl')
    properties = E.SubElement(table, w + 'tblPr')
    E.SubElement(properties, w + 'tblW', {w + 'w': '9360', w + 'type': 'dxa'})
    E.SubElement(properties, w + 'tblLayout', {w + 'type': 'fixed'})
    E.SubElement(E.SubElement(table, w + 'tblGrid'), w + 'gridCol', {w + 'w': '9360'})
    cell = E.SubElement(E.SubElement(table, w + 'tr'), w + 'tc')
    E.SubElement(E.SubElement(cell, w + 'tcPr'), w + 'tcW', {w + 'w': '9360', w + 'type': 'dxa'})
    E.SubElement(E.SubElement(E.SubElement(cell, w + 'p'), w + 'r'), w + 't').text = 'Table fallback sample.'
    body.insert(list(body).index(section), table)
    E.SubElement(section, w + 'headerReference', {w + 'type': 'default', r + 'id': 'header'})
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    header = E.Element(w + 'hdr')
    E.SubElement(E.SubElement(E.SubElement(header, w + 'p'), w + 'r'), w + 't').text = 'Header fallback sample.'
    entries['word/header1.xml'] = E.tostring(header, encoding='utf-8', xml_declaration=True)
    styles = E.fromstring(entries['word/styles.xml'])
    styles.find(w + 'docDefaults/' + w + 'rPrDefault/' + w + 'rPr/' + w + 'sz').set(w + 'val', str(points * 2))
    entries['word/styles.xml'] = E.tostring(styles, encoding='utf-8', xml_declaration=True)
    target = cases / ('default-typeface-mixed-' + str(points) + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as z:
        for part, data in entries.items():
            info = ZipInfo(part, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=0, HeldOut=False,
        ExpectedText=['Composed story.', 'Table fallback sample.', 'Header fallback sample.', 'share the same input series.']))
(cases / 'default-typefaces.json').write_text(json.dumps(rows, indent=2)+'\n', encoding='utf-8')
print('Pinned',len(rows),'mixed-story implicit-typeface controls')
