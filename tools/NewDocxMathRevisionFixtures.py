"""Public inline text revision and unchanged equation composition controls."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import json
import hashlib
root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
rows = []
for story, part in [('body', 'word/document.xml'), ('cell', 'word/document.xml'),
                    ('header', 'word/header1.xml'), ('footer', 'word/footer1.xml')]:
    source = cases / ('math-numbered-display-' + story + '.docx')
    for kind in ['ins', 'del']:
        with ZipFile(source) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        document = E.fromstring(entries[part])
        paragraph = next(p for p in document.iter(w + 'p') if list(p.iter(m + 'oMath')))
        run = next(r for r in paragraph.findall(w + 'r') if r.find(w + 't') is not None)
        index = list(paragraph).index(run)
        paragraph.remove(run)
        revision = E.Element(w + kind, {w + 'id': '71', w + 'author': 'Public Reviewer', w + 'date': '2026-10-16T00:00:00Z'})
        revision.append(run)
        paragraph.insert(index, revision)
        if kind == 'del':
            for text in run.iter(w + 't'):
                text.tag = w + 'delText'
        entries[part] = E.tostring(document, encoding='utf-8', xml_declaration=True)
        target = cases / ('math-revision-' + story + '-' + kind + '.docx')
        with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
            for name, data in entries.items():
                info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            ExpectedEquationCount=4, HeldOut=False, ExpectedText=['share the same input series.']))
(cases / 'math-revisions.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'math revision controls')
