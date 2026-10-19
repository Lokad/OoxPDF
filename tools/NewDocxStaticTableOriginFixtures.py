"""Verify bounded body/header/footer table-origin controls from one pinned source."""
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
rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()
source = next(case for case in json.loads((directory / 'inline-box-review-frames.json').read_text())
    if case['Id'] == 'inline-box-review-static-cell')
original = (root / source['Input']).read_bytes()
assert sha(original) == source['Sha256']
with ZipFile(BytesIO(original)) as archive:
    infos = archive.infolist()
    original_parts = {info.filename: archive.read(info.filename) for info in infos}
part = 'word/inline-box-review-header.xml'
prepared, cases = [], [source]
for variant in ['header-zero', 'header-wide', 'header-prose-wide', 'footer-wide', 'body-wide']:
    parts = original_parts.copy()
    story = E.fromstring(parts[part])
    table = story.find(w + 'tbl')
    assert table is not None
    margins = table.find(w + 'tblPr/' + w + 'tblCellMar')
    for side in ['left', 'right']:
        margins.find(w + side).set(w + 'w', '0' if variant == 'header-zero' else '200')
    expected = source['ExpectedText']
    equations = source['ExpectedEquationCount']
    if variant == 'header-prose-wide':
        paragraph = table.find(w + 'tr/' + w + 'tc/' + w + 'p')
        paragraph.clear()
        E.SubElement(E.SubElement(paragraph, w + 'r'), w + 't').text = 'Public table cell text.'
        expected = ['Before public math table.', 'After public math table.', 'Public table cell text.']
        equations = 0
    document = E.fromstring(parts['word/document.xml'])
    section = document.find(w + 'body/' + w + 'sectPr')
    reference = section.find(w + 'headerReference')
    assert reference is not None
    math_part = part
    if variant == 'body-wide':
        story.remove(table)
        document.find(w + 'body').insert(1, table)
        section.remove(reference)
        math_part = 'word/document.xml'
    elif variant == 'footer-wide':
        story.tag = w + 'ftr'
        reference.tag = w + 'footerReference'
        relationships = E.fromstring(parts['word/_rels/document.xml.rels'])
        reference_id = reference.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        relationship = next(node for node in relationships if node.get('Id') == reference_id)
        relationship.set('Type', relationship.get('Type').replace('/header', '/footer'))
        parts['word/_rels/document.xml.rels'] = E.tostring(relationships, encoding='utf-8', xml_declaration=True)
        # The public part name stays stable; only its content type changes.
        ct = E.fromstring(parts['[Content_Types].xml'])
        next(node for node in ct if node.get('PartName') == '/' + part).set('ContentType',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
        parts['[Content_Types].xml'] = E.tostring(ct, encoding='utf-8', xml_declaration=True)
    parts[part] = E.tostring(story, encoding='utf-8', xml_declaration=True)
    parts['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    output = BytesIO()
    with ZipFile(output, 'w') as archive:
        for info in infos:
            archive.writestr(copy.copy(info), parts[info.filename])
    data = output.getvalue()
    id = 'static-table-origin-' + variant
    case = {'Id': id, 'Input': 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath/' + id + '.docx',
        'Sha256': sha(data), 'SourceSha256': source['Sha256'], 'HeldOut': True, 'Variant': variant,
        'MathPart': math_part, 'ExpectedEquationCount': equations, 'ExpectedRenderedEquationCount': equations,
        'ExpectedImageCount': 0, 'ExpectedText': expected,
        'Change': 'Vary authored cell margin or container/content while retaining the source table, frame and real body comment.'}
    cases.append(case)
    prepared.append((root / case['Input'], data))
manifest = directory / 'static-table-origins.json'
manifest_bytes = (json.dumps(cases, indent=2) + '\n').replace('\n', '\r\n').encode()
if manifest.exists():
    assert json.loads(manifest.read_text()) == cases
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
print('One retained and five derived static/body table-origin controls verified.')
