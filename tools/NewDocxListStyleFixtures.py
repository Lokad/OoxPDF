"""Independent numbering-level/paragraph-mark font-style controls."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
source = cases / 'static-frame-footer-plain-first.docx'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
rows = []
for name in ['mark-bold', 'level-bold', 'mark-bold-level-off', 'normal-bold', 'mark-italic', 'level-italic',
             'run-bold-untracked', 'run-bold-inserted']:
    with ZipFile(source) as z:
        entries = {part: z.read(part) for part in z.namelist()}
    if name.startswith('mark-'):
        xml = E.fromstring(entries['word/footer1.xml'])
        properties = xml.find(w + 'p/' + w + 'pPr')
        mark = E.SubElement(properties, w + 'rPr')
        E.SubElement(mark, w + ('i' if name == 'mark-italic' else 'b'))
        entries['word/footer1.xml'] = E.tostring(xml, encoding='utf-8', xml_declaration=True)
    if name in ['level-bold', 'mark-bold-level-off', 'level-italic']:
        xml = E.fromstring(entries['word/numbering.xml'])
        level = xml.find(w + 'abstractNum/' + w + 'lvl')
        properties = level.find(w + 'rPr')
        if properties is None:
            properties = E.SubElement(level, w + 'rPr')
        bold = E.SubElement(properties, w + ('i' if name == 'level-italic' else 'b'))
        if name.endswith('off'):
            bold.set(w + 'val', '0')
        entries['word/numbering.xml'] = E.tostring(xml, encoding='utf-8', xml_declaration=True)
    if name == 'normal-bold':
        xml = E.fromstring(entries['word/styles.xml'])
        normal = E.SubElement(xml, w + 'style', {w + 'type': 'paragraph', w + 'styleId': 'Normal', w + 'default': '1'})
        E.SubElement(E.SubElement(normal, w + 'rPr'), w + 'b')
        entries['word/styles.xml'] = E.tostring(xml, encoding='utf-8', xml_declaration=True)
    if name.startswith('run-bold-'):
        xml = E.fromstring(entries['word/footer1.xml'])
        for properties in xml.iter(w + 'rPr'):
            for change in list(properties.findall(w + 'rPrChange')):
                properties.remove(change)
        if name.endswith('inserted'):
            paragraph = xml.find(w + 'p')
            run = paragraph.find(w + 'r')
            index = list(paragraph).index(run)
            paragraph.remove(run)
            insertion = E.Element(w + 'ins', {w + 'id': '201', w + 'author': 'Public Reviewer', w + 'date': '2026-10-17T00:00:00Z'})
            insertion.append(run)
            paragraph.insert(index, insertion)
        entries['word/footer1.xml'] = E.tostring(xml, encoding='utf-8', xml_declaration=True)
    target = cases / ('list-style-' + name + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as z:
        for part, data in entries.items():
            info = ZipInfo(part, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=sum(len(list(E.fromstring(data).iter(m + 'oMath')))
            for part, data in entries.items() if part.endswith('.xml')),
        HeldOut=False, ExpectedText=['share the same input series.']))
(cases / 'list-styles.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'list font-style owner controls')
