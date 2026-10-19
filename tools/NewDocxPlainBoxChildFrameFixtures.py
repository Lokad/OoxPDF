"""Verify source-exact plain-before/after-math textbox child controls."""
import argparse
import copy
import hashlib
import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parent.parent
directory = root / 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()
review = json.loads((directory / 'inline-box-review-frames.json').read_text())
source = next(case for case in review if case['Id'] == 'inline-box-review-body')
original = (root / source['Input']).read_bytes()
assert sha(original) == source['Sha256']
with ZipFile(BytesIO(original)) as archive:
    infos = archive.infolist()
    source_parts = {info.filename: archive.read(info.filename) for info in infos}
cases = json.loads((directory / 'inline-frame-ends.json').read_text())
cases += json.loads((directory / 'inline-box-owned-paragraphs.json').read_text())
cases += [case for case in json.loads((directory / 'inline-box-math.json').read_text()) if case['Id'] == 'inline-box-math-cell']
assert len(cases) == len({case['Id'] for case in cases}) == 22
prepared = []
for variant in ['before', 'after']:
    parts = dict(source_parts)
    document = E.fromstring(parts['word/document.xml'])
    box = document.find('.//' + w + 'txbxContent')
    assert box is not None and len(box.findall(w + 'p')) == 1
    paragraph = E.Element(w + 'p')
    E.SubElement(E.SubElement(paragraph, w + 'r'), w + 't').text = 'Public context.'
    box.insert(0 if variant == 'before' else len(box), paragraph)
    parts['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    output = BytesIO()
    with ZipFile(output, 'w') as archive:
        for info in infos:
            archive.writestr(copy.copy(info), parts[info.filename])
    data = output.getvalue()
    id = 'inline-box-child-plain-' + variant
    case = {'Id': id, 'Input': 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath/' + id + '.docx',
        'Sha256': sha(data), 'SourceSha256': source['Sha256'], 'HeldOut': True, 'Variant': variant,
        'MathPart': 'word/document.xml', 'ExpectedEquationCount': 1, 'ExpectedRenderedEquationCount': 1,
        'ExpectedImageCount': 0, 'ExpectedText': source['ExpectedText'] + ['Public context.'],
        'Change': 'Add an ordinary child paragraph before/after existing math; retain frame extents, source math and real body comment.'}
    cases.append(case)
    prepared.append((root / case['Input'], data))
assert len(cases) == 24
manifest = directory / 'inline-box-plain-child-frames.json'
manifest_bytes = (json.dumps(cases, indent=2) + '\n').replace('\n', '\r\n').encode()
if manifest.exists():
    assert json.loads(manifest.read_text()) == cases, 'Derived manifest differs'
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--write', action='store_true')
if parser.parse_args().write:
    for target, data in prepared + [(manifest, manifest_bytes)]:
        staging = target.with_suffix(target.suffix + '.staging')
        with staging.open('xb') as stream:
            stream.write(data)
        staging.replace(target)
else:
    for target, data in prepared:
        assert target.read_bytes() == data, target.name
print('Twenty-two existing and two mixed plain/math child controls verified.')
