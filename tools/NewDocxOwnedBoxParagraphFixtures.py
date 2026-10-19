"""Verify ordinary-text box controls derived without changing child frames."""
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
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()
siblings = json.loads((directory / 'inline-sibling-frames.json').read_text())
ordinary = json.loads((directory / 'inline-box-math.json').read_text())
cases = json.loads((directory / 'inline-box-review-frames.json').read_text())
cases += [case for case in ordinary if case['Id'] in
    ['inline-box-boxed-equation-valid', 'inline-box-boxed-text-valid', 'inline-box-mixed-source-report']]
prepared = []
for variant in ['first', 'wrap']:
    source = next(case for case in siblings if case['Id'] == 'inline-sibling-' + variant)
    data = (root / source['Input']).read_bytes()
    assert sha(data) == source['Sha256']
    with ZipFile(BytesIO(data)) as archive:
        infos = archive.infolist()
        parts = {info.filename: archive.read(info.filename) for info in infos}
    document = E.fromstring(parts['word/document.xml'])
    removed = 0
    for paragraph in document.iter(w + 'p'):
        for child in list(paragraph):
            if child.tag in [m + 'oMath', m + 'oMathPara']:
                paragraph.remove(child)
                removed += 1
    assert removed == 1 and not any(node.tag in [m + 'oMath', m + 'oMathPara'] for node in document.iter())
    parts['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    output = BytesIO()
    with ZipFile(output, 'w') as archive:
        for info in infos:
            archive.writestr(copy.copy(info), parts[info.filename])
    data = output.getvalue()
    id = 'inline-box-owned-text-' + variant
    case = {'Id': id, 'Input': 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath/' + id + '.docx',
        'Sha256': sha(data), 'SourceSha256': source['Sha256'], 'HeldOut': True, 'Variant': variant,
        'MathPart': 'word/document.xml', 'ExpectedEquationCount': 0, 'ExpectedRenderedEquationCount': 0,
        'ExpectedImageCount': 0, 'ExpectedText': source['ExpectedText'],
        'Change': 'Remove only parent mathematics; preserve source-ordered box, prose, real body comment and child frame.'}
    cases.append(case)
    prepared.append((root / case['Input'], data))
assert len(cases) == 10
manifest = directory / 'inline-box-owned-paragraphs.json'
manifest_bytes = (json.dumps(cases, indent=2) + '\n').replace('\n', '\r\n').encode()
if manifest.exists():
    assert json.loads(manifest.read_text()) == cases, 'Derived fixture manifest differs'
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
print('Eight existing controls and two ordinary-text source-owned box controls verified.')
