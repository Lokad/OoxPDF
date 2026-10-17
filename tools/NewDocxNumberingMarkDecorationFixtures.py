"""Qualify inserted-mark decoration against paragraph and level overrides."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import json
import hashlib

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
source = cases / 'list-decoration-mark-inserted.docx'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = json.loads((cases / 'numbering-marks.json').read_text(encoding='utf-8'))
for name in ['mark-red-double', 'level-green-none', 'level-green-double']:
    with ZipFile(source) as z:
        entries = {part: z.read(part) for part in z.namelist()}
    footer = E.fromstring(entries['word/footer1.xml'])
    mark = footer.find(w + 'p/' + w + 'pPr/' + w + 'rPr')
    E.SubElement(mark, w + 'color', {w + 'val': 'AA2200'})
    E.SubElement(mark, w + 'u', {w + 'val': 'double'})
    entries['word/footer1.xml'] = E.tostring(footer, encoding='utf-8', xml_declaration=True)
    if name.startswith('level-'):
        numbering = E.fromstring(entries['word/numbering.xml'])
        level = E.SubElement(numbering.find(w + 'abstractNum/' + w + 'lvl'), w + 'rPr')
        E.SubElement(level, w + 'color', {w + 'val': '008844'})
        E.SubElement(level, w + 'u', {w + 'val': 'double' if name.endswith('double') else 'none'})
        entries['word/numbering.xml'] = E.tostring(numbering, encoding='utf-8', xml_declaration=True)
    target = cases / ('numbering-mark-decoration-' + name + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as z:
        for part, data in entries.items():
            info = ZipInfo(part, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        Kind='plain', Story='footer', ExpectedEquationCount=1,
        SelectedSourceEquations=0, HeldOut=False, ExpectedText=['share the same input series.']))
(cases / 'numbering-mark-decorations.json').write_text(json.dumps(rows, indent=2)+'\n', encoding='utf-8')
print('Pinned', len(rows), 'decoration sources; eight existing inputs reused')
