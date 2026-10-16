"""Preserve list-label ownership when supported formulas use explicit math zones."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
from copy import deepcopy
import json
import hashlib
root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
rows = []
def save(target, entries, source, count):
    with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=count, HeldOut=False, ExpectedText=['share the same input series.']))
for story, suffix, part in [('body', '', 'word/document.xml'), ('cell', '-cell', 'word/document.xml'),
                          ('header', '-header', 'word/header1.xml'), ('footer', '-footer', 'word/footer1.xml')]:
    source = cases / ('math-composition-numbered' + suffix + '.docx')
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries[part])
    count = 0
    for paragraph in document.iter(w + 'p'):
        if paragraph.find(w + 'pPr/' + w + 'numPr') is None:
            continue
        for equation in list(paragraph.findall(m + 'oMath')):
            index = list(paragraph).index(equation)
            paragraph.remove(equation)
            display = E.Element(m + 'oMathPara')
            display.append(equation)
            paragraph.insert(index, display)
            count += 1
    entries[part] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    save(cases / ('math-numbered-display-' + story + '.docx'), entries, source, count)
for kind in ['math-only', 'math-break']:
    source = cases / 'math-numbered-display-body.docx'
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries['word/document.xml'])
    paragraph = next(p for p in document.iter(w + 'p') if p.find(w + 'pPr/' + w + 'numPr') is not None)
    properties = deepcopy(paragraph.find(w + 'pPr'))
    equation = deepcopy(paragraph.find(m + 'oMathPara'))
    paragraph.clear()
    paragraph.append(properties)
    paragraph.append(equation)
    if kind == 'math-break':
        E.SubElement(E.SubElement(paragraph, w + 'r'), w + 'br')
        paragraph.append(deepcopy(equation))
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    save(cases / ('math-numbered-display-' + kind + '.docx'), entries, source, 5 if kind == 'math-break' else 4)
(cases / 'math-numbered-display.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'numbered math-zone controls')
