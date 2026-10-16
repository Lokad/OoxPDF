"""Pin image-only math rows and adjacent/combined fraction ownership controls."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
target_root = cases / 'CompositionProbes'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
pr = '{http://schemas.openxmlformats.org/package/2006/relationships}'
ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
report_source = cases / 'Composition/report.docx'
with ZipFile(report_source) as archive:
    report = E.fromstring(archive.read('word/document.xml'))
    drawing = next(run for run in report.iter(w + 'r') if run.find(w + 'drawing') is not None)
    image = archive.read('word/media/trend.png')
rows = []
for container in ['body', 'cell']:
    source = target_root / f'math-fields-simple-{container}.docx'
    for position in ['before', 'between', 'combined']:
        with ZipFile(source) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        document = E.fromstring(entries['word/document.xml'])
        paragraph = next(p for p in document.iter(w + 'p') if list(p.iter(m + 'oMath')))
        equations = list(paragraph.iter(m + 'oMath'))
        assert len(equations) == 2
        for child in list(paragraph):
            if child.tag != w + 'pPr':
                paragraph.remove(child)
        if position == 'combined':
            for child in list(equations[1]):
                equations[0].append(deepcopy(child))
            equations = equations[:1]
        displays = []
        for equation in equations:
            group = E.Element(m + 'oMathPara')
            group.append(equation)
            displays.append(group)
        picture = deepcopy(drawing)
        for node in picture.iter():
            if r + 'embed' in node.attrib:
                node.set(r + 'embed', 'rIdRowImage')
        for item in [displays[0], picture, displays[1]] if position == 'between' else [picture, *displays]:
            paragraph.append(item)
        entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
        relationships = E.fromstring(entries['word/_rels/document.xml.rels'])
        E.SubElement(relationships, pr + 'Relationship', {'Id': 'rIdRowImage',
            'Type': r[1:-1] + '/image', 'Target': 'media/math-row-trend.png'})
        entries['word/_rels/document.xml.rels'] = E.tostring(relationships, encoding='utf-8', xml_declaration=True)
        types = E.fromstring(entries['[Content_Types].xml'])
        if not any(node.get('Extension') == 'png' for node in types):
            E.SubElement(types, ct + 'Default', {'Extension': 'png', 'ContentType': 'image/png'})
        entries['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
        entries['word/media/math-row-trend.png'] = image
        target = target_root / f'math-display-image-{container}-{position}.docx'
        with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
            for name, data in entries.items():
                info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            ImageSourceSha256=hashlib.sha256(report_source.read_bytes()).hexdigest().upper(),
            ExpectedEquationCount=len(equations), ExpectedInlineImageCount=1,
            HeldOut=False, ExpectedText=['Following prose.']))
(target_root / 'math-display-images.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'display/image fraction controls')
