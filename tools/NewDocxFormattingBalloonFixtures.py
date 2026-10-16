"""Public composition controls for run-formatting descriptions in Word mixed view."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
source = cases / 'math-run-format-body-bold.docx'
rows = []
for name, values in [('not-bold', [('b', '0', '1')]),
                     ('italic', [('i', '1', '0')]),
                     ('bold-italic', [('b', '1', '0'), ('i', '1', '0')])]:
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries['word/document.xml'])
    change = next(document.iter(w + 'rPrChange'))
    properties = next(p for p in document.iter(w + 'rPr') if change in p)
    previous = change.find(w + 'rPr')
    for item in list(previous):
        previous.remove(item)
    for item in list(properties):
        if item.tag in (w + 'b', w + 'i'):
            properties.remove(item)
    for prop, after, before in values:
        E.SubElement(properties, w + prop, {w + 'val': after})
        E.SubElement(previous, w + prop, {w + 'val': before})
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    target = cases / ('formatting-balloon-body-' + name + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
        for part, data in entries.items():
            info = ZipInfo(part, (2026, 10, 16, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=4, HeldOut=False, ExpectedText=['share the same input series.']))
(cases / 'formatting-balloons.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'formatting description controls')
