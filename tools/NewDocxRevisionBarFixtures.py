"""Deterministic public controls for revision bars in authored/printed page frames."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
output = cases / 'RevisionBars'
output.mkdir(exist_ok=True)
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = []
for name, left, right, width, height, report in [
    ('left54', 1080, 1440, 12240, 15840, False),
    ('left108', 2160, 1440, 12240, 15840, False),
    ('left144', 2880, 1440, 12240, 15840, False),
    ('right144', 1440, 2880, 12240, 15840, False),
    ('landscape', 1440, 1440, 15840, 12240, False),
    ('comment-scaled', 1440, 1440, 12240, 15840, True),
    ('left36', 720, 1440, 12240, 15840, False),
    ('left90', 1800, 1440, 12240, 15840, False),
]:
    source = cases / ('Composition/report.docx' if report else 'CompositionProbes/math-revision-body-del.docx')
    target = output / ('revision-bar-' + name + '.docx')
    with ZipFile(source) as archive, ZipFile(target, 'w', ZIP_DEFLATED) as result:
        for entry in archive.infolist():
            data = archive.read(entry.filename)
            if entry.filename == 'word/document.xml':
                document = E.fromstring(data)
                margins = next(document.iter(w + 'pgMar'))
                margins.set(w + 'left', str(left))
                margins.set(w + 'right', str(right))
                size = next(document.iter(w + 'pgSz'))
                size.set(w + 'w', str(width))
                size.set(w + 'h', str(height))
                if report:
                    paragraph = next(document.iter(w + 'p'))
                    run = paragraph.find(w + 'r')
                    index = list(paragraph).index(run)
                    paragraph.remove(run)
                    revision = E.Element(w + 'ins', {w + 'id': '93', w + 'author': 'Public Reviewer', w + 'date': '2026-10-16T00:00:00Z'})
                    revision.append(run)
                    paragraph.insert(index, revision)
                data = E.tostring(document, encoding='utf-8', xml_declaration=True)
            info = ZipInfo(entry.filename, (1980, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            result.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), HeldOut=False,
        ExpectedEquationCount=6 if report else 4, ExpectedText=[]))
(output / 'revision-bars.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'public revision-bar controls')
