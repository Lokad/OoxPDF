"""Unchanged equations beside a run-formatting revision, across four containers."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
rows = []
for story, kind in [('body', 'bold'), ('cell', 'bold'), ('header', 'bold'),
                    ('footer', 'bold'), ('body', 'void')]:
    source = cases / ('math-numbered-display-' + story + '.docx')
    part = 'word/' + (story + '1.xml' if story in ('header', 'footer') else 'document.xml')
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries[part])
    paragraph = next(p for p in document.iter(w + 'p') if list(p.iter(m + 'oMath')))
    run = next(r for r in paragraph.findall(w + 'r') if r.find(w + 't') is not None)
    properties = run.find(w + 'rPr')
    if properties is None:
        properties = E.Element(w + 'rPr')
        run.insert(0, properties)
    if kind == 'bold':
        E.SubElement(properties, w + 'b')
    change = E.SubElement(properties, w + 'rPrChange', {
        w + 'id': '96', w + 'author': 'Public Reviewer', w + 'date': '2026-10-16T00:00:00Z'})
    previous = E.SubElement(change, w + 'rPr')
    if kind == 'bold':
        E.SubElement(previous, w + 'b', {w + 'val': '0'})
    entries[part] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    target = cases / ('math-run-format-' + story + '-' + kind + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            info = ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=4, HeldOut=False, ExpectedText=['share the same input series.']))
(cases / 'math-run-format.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'math/run-formatting controls')
