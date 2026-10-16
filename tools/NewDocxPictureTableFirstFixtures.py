"""Pin first-table picture/math rows with and without a neighboring prose cell."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
source = cases / 'math-display-image-cell-before.docx'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = []
for shape in ['two-cells', 'single-cell']:
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries['word/document.xml'])
    body = document.find(w + 'body')
    table = body.find(w + 'tbl')
    for child in list(body):
        if child is table:
            break
        body.remove(child)
    cells = list(table.iter(w + 'tc'))
    paragraph = cells[-1].find(w + 'p')
    properties = paragraph.find(w + 'pPr')
    if properties is None:
        properties = E.Element(w + 'pPr')
        paragraph.insert(0, properties)
    mark = E.SubElement(properties, w + 'rPr')
    E.SubElement(mark, w + 'rFonts', {w + 'ascii': 'Arial', w + 'hAnsi': 'Arial'})
    E.SubElement(mark, w + 'sz', {w + 'val': '24'})
    if shape == 'single-cell':
        row = next(table.iter(w + 'tr'))
        row.remove(cells[0])
        grid = table.find(w + 'tblGrid')
        grid.clear()
        E.SubElement(grid, w + 'gridCol', {w + 'w': '9360'})
        cells[-1].find(w + 'tcPr').find(w + 'tcW').set(w + 'w', '9360')
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    target = cases / f'math-picture-first-table-{shape}.docx'
    with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=2, ExpectedInlineImageCount=1, HeldOut=False,
        ExpectedText=['Following prose.']))
(cases / 'math-picture-first-tables.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'first-table picture controls')
