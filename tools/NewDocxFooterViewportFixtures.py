"""Verify independent footer viewport controls; --write publishes atomically."""
import argparse
import copy
import hashlib
import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--write', action='store_true')
args = parser.parse_args()
parent = next(row for row in json.loads((cases / 'static-table-origins.json').read_text())
              if row['Id'] == 'static-table-origin-footer-wide')
original = (root / parent['Input']).read_bytes()
assert sha(original) == parent['Sha256']
with ZipFile(BytesIO(original)) as archive:
    infos = archive.infolist()
    parts = {entry.filename: archive.read(entry.filename) for entry in infos}
rows, jobs = [], []
for variant, attribute, before, after in [('bottom-margin', 'bottom', '1440', '1800'),
                                         ('footer-distance', 'footer', '720', '1080')]:
    document = E.fromstring(parts['word/document.xml'])
    margin = document.find(w + 'body/' + w + 'sectPr/' + w + 'pgMar')
    assert margin is not None and margin.get(w + attribute) == before
    source_xml = E.tostring(document)
    margin.set(w + attribute, after)
    changed_xml = E.tostring(document, encoding='utf-8', xml_declaration=True)
    margin.set(w + attribute, before)
    assert E.tostring(document) == source_xml
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as archive:
        for entry in infos:
            archive.writestr(copy.copy(entry), changed_xml if entry.filename == 'word/document.xml' else parts[entry.filename])
    data = buffer.getvalue()
    identifier = 'footer-table-viewport-' + variant
    row = copy.deepcopy(parent)
    row.update(Id=identifier, Input='tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath/' + identifier + '.docx',
               Sha256=sha(data), SourceInput=parent['Input'], SourceSha256=parent['Sha256'], Variant=variant,
               MarginAttribute=attribute, SourceTwips=int(before), ChangedTwips=int(after),
               Change='Change only one section page-margin attribute; table, paragraph, drawing, math and all other package parts remain exact.')
    rows.append(row)
    jobs.append((cases / (identifier + '.docx'), data))
jobs.append((cases / 'footer-table-viewports.json',
             (json.dumps(rows, indent=2) + '\n').replace('\n', '\r\n').encode()))
for target, data in jobs:
    if args.write:
        stage = target.with_suffix(target.suffix + '.staging')
        with stage.open('xb') as output:
            output.write(data)
        stage.replace(target)
    else:
        assert target.read_bytes() == data, target.name
print('Two independent footer viewport controls verified; only one section attribute changes.')
