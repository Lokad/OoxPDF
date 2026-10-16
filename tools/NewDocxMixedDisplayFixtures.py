"""Pin explicit-display source alongside ordinary prose without changing formulas."""
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
for kind in ['literal', 'fraction']:
    source = cases / f'math-printed-frame-{kind}-anchor12-multiple.docx'
    for grouping in ['first', 'all']:
        with ZipFile(source) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        document = E.fromstring(entries['word/document.xml'])
        equations = list(document.iter(m + 'oMath'))
        assert len(equations) == 2
        for equation in equations[:1] if grouping == 'first' else equations:
            parent = next(p for p in document.iter() if equation in list(p))
            assert parent.tag == w + 'p'
            index = list(parent).index(equation)
            parent.remove(equation)
            display = E.Element(m + 'oMathPara')
            display.append(equation)
            parent.insert(index, display)
        entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
        target = cases / f'math-mixed-display-{kind}-{grouping}.docx'
        with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
            for name, data in entries.items():
                info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            HeldOut=False, ExpectedEquationCount=2,
            ExpectedText=['Before formula:', 'and', 'Following prose.']))
for container, source_name, part in [
    ('cell', 'math-explicit-display-left-cell', 'word/document.xml'),
    ('footer', 'math-explicit-display-right-footer', 'word/math-tabs-footer.xml')]:
    source = cases / (source_name + '.docx')
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries[part])
    paragraph = next(p for p in document.iter(w + 'p') if list(p.iter(m + 'oMath')))
    def prose(text):
        run = E.Element(w + 'r')
        props = E.SubElement(run, w + 'rPr')
        E.SubElement(props, w + 'rFonts', {w + 'ascii': 'Arial', w + 'hAnsi': 'Arial'})
        E.SubElement(props, w + 'sz', {w + 'val': '24'})
        E.SubElement(run, w + 't', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = text
        return run
    paragraph.insert(1 if paragraph.find(w + 'pPr') is not None else 0, prose('Before formula: '))
    paragraph.append(prose(' End.'))
    entries[part] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    target = cases / f'math-mixed-display-{container}.docx'
    with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=2, ExpectedText=['Before formula:', 'End.']))
(cases / 'math-mixed-display.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'mixed display/prose controls')
