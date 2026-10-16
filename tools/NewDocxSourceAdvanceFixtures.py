"""Source kerning thresholds and nominal glyph advances, without layout guards."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
source = cases / 'overwide-word-default.docx'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = [dict(Id=source.stem, Input=source.relative_to(root).as_posix(),
             Sha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
             ExpectedEquationCount=0, ExpectedText=[], HeldOut=False)]
controls = [('word-kern-' + str(value), None, None, value) for value in (0, 16, 28)]
controls += [('word-kern-calibri-11-met', 'Calibri', 11, 22),
             ('word-kern-arial-11-met', 'Arial', 11, 22),
             ('word-kern-times-11-met', 'Times New Roman', 11, 22),
             ('word-kern-calibri-14-met', 'Calibri', 14, 28)]
for identity, family, size, threshold in controls:
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries['word/document.xml'])
    for run in document.iter(w + 'r'):
        properties = run.find(w + 'rPr')
        if properties is None:
            properties = E.Element(w + 'rPr')
            run.insert(0, properties)
        if family:
            E.SubElement(properties, w + 'rFonts', {w + name: family for name in ('ascii', 'hAnsi', 'eastAsia', 'cs')})
            E.SubElement(properties, w + 'sz', {w + 'val': str(size * 2)})
        E.SubElement(properties, w + 'kern', {w + 'val': str(threshold)})
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    target = cases / (identity + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=identity, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=rows[0]['Sha256'], ExpectedEquationCount=0, ExpectedText=[], HeldOut=False))
(cases / 'source-text-advances.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'source-advance controls')
