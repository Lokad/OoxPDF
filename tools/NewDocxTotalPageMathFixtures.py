"""Pin actual total=2, current page=1 and stale cache=777 composition controls."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = []
for form in ['simple', 'packed']:
    for container in ['body', 'cell']:
        source = cases / f'math-fields-{form}-{container}.docx'
        source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
        with ZipFile(source) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        document = E.fromstring(entries['word/document.xml'])
        for field in document.iter(w + 'fldSimple'):
            field.set(w + 'instr', ' NUMPAGES ')
            for text in field.iter(w + 't'):
                text.text = '777'
        for field in document.iter(w + 'instrText'):
            field.text = ' NUMPAGES '
        for run in document.iter(w + 'r'):
            if run.find(w + 'instrText') is not None:
                for text in run.iter(w + 't'):
                    text.text = '777'
        following = next(p for p in document.iter(w + 'p')
            if any(t.text == 'Following prose.' for t in p.iter(w + 't')))
        props = following.find(w + 'pPr')
        if props is None:
            props = E.Element(w + 'pPr')
            following.insert(0, props)
        E.SubElement(props, w + 'pageBreakBefore')
        entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
        target = cases / f'math-total-pages-{form}-{container}.docx'
        with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
            for name, data in entries.items():
                info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=source_sha, HeldOut=False, ExpectedEquationCount=2,
            ExpectedText=['Following prose.', 'Start', 'middle', 'End.']))
(cases / 'math-total-pages.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'total-page/math controls')
