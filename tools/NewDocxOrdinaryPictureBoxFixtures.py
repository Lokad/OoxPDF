"""Verify ordinary picture/box composition from pinned body and cell sources."""
import argparse
import copy
import hashlib
import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parent.parent
directory = root / 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()
sources = json.loads((directory / 'inline-sibling-frames.json').read_text())
mixed = next(case for case in sources if case['Id'] == 'inline-sibling-mixed')
cell = next(case for case in sources if case['Id'] == 'inline-sibling-cell')

def unpack(case):
    data = (root / case['Input']).read_bytes()
    assert sha(data) == case['Sha256']
    with ZipFile(BytesIO(data)) as archive:
        infos = archive.infolist()
        return infos, {info.filename: archive.read(info.filename) for info in infos}

mixed_infos, mixed_parts = unpack(mixed)
document = E.fromstring(mixed_parts['word/document.xml'])
paragraph = next(p for p in document.iter(w + 'p') if p.find(w + 'r/' + w + 'drawing') is not None)
equations = [node for node in paragraph if node.tag in [m + 'oMath', m + 'oMathPara']]
assert len(equations) == 2
for equation in equations:
    paragraph.remove(equation)
assert not any(node.tag in [m + 'oMath', m + 'oMathPara'] for node in document.iter())

prepared = []
cases = []
for variant in ['body', 'cell']:
    if variant == 'body':
        infos, parts = mixed_infos, dict(mixed_parts)
        tree = document
    else:
        infos, parts = unpack(cell)
        tree = E.fromstring(parts['word/document.xml'])
        parent = tree.find('.//' + w + 'tc')
        old = next(p for p in parent.findall(w + 'p') if p.find('.//' + w + 'txbxContent') is not None)
        index = list(parent).index(old)
        parent.remove(old)
        parent.insert(index, copy.deepcopy(paragraph))
        links = E.fromstring(parts['word/_rels/document.xml.rels'])
        assert not any(link.get('Id') == 'rSiblingPicture' for link in links)
        E.SubElement(links, rel + 'Relationship', {'Id': 'rSiblingPicture',
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'Target': 'media/sibling-frame.png'})
        parts['word/_rels/document.xml.rels'] = E.tostring(links, encoding='utf-8', xml_declaration=True)
        parts['word/media/sibling-frame.png'] = mixed_parts['word/media/sibling-frame.png']
        types = E.fromstring(parts['[Content_Types].xml'])
        if not any(node.tag == ct + 'Default' and node.get('Extension') == 'png' for node in types):
            E.SubElement(types, ct + 'Default', {'Extension': 'png', 'ContentType': 'image/png'})
        parts['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
    parts['word/document.xml'] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
    output = BytesIO()
    with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
        for info in infos:
            archive.writestr(copy.copy(info), parts[info.filename])
        for name in sorted(set(parts) - {info.filename for info in infos}):
            info = ZipInfo(name, (2026, 10, 19, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, parts[name])
    data = output.getvalue()
    id = 'inline-box-picture-' + variant
    case = {'Id': id, 'Input': 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath/' + id + '.docx',
        'Sha256': sha(data), 'SourceSha256': (mixed if variant == 'body' else cell)['Sha256'],
        'SiblingSourceSha256': mixed['Sha256'], 'HeldOut': True, 'Variant': variant, 'MathPart': 'word/document.xml',
        'ExpectedEquationCount': 0, 'ExpectedRenderedEquationCount': 0, 'ExpectedImageCount': 1,
        'ExpectedText': mixed['ExpectedText'], 'Change': 'Remove parent mathematics while retaining ordinary text, source picture, textbox, real comment and container constraints.'}
    cases.append(case)
    prepared.append((root / case['Input'], data))
manifest = directory / 'inline-box-ordinary-pictures.json'
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
print('Two ordinary picture/textbox body/cell controls verified.')
