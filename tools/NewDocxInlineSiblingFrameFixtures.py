"""Verify source-owned sibling-frame controls; --write publishes validated bytes."""
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
a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
wp = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()
ordinary = json.loads((directory / 'inline-box-math.json').read_text())
review = json.loads((directory / 'inline-box-review-frames.json').read_text())
by_variant = {case['Variant']: case for case in ordinary}
review_by_variant = {case['Variant']: case for case in review}

def unpack(case):
    data = (root / case['Input']).read_bytes()
    assert sha(data) == case['Sha256']
    with ZipFile(BytesIO(data)) as archive:
        return archive.infolist(), {info.filename: archive.read(info.filename) for info in archive.infolist()}

def pack(infos, parts):
    buffer = BytesIO()
    with ZipFile(buffer, 'w', ZIP_DEFLATED) as archive:
        for info in infos:
            archive.writestr(copy.copy(info), parts[info.filename])
        for name in sorted(set(parts) - {info.filename for info in infos}):
            info = ZipInfo(name, (2026, 10, 19, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, parts[name])
    return buffer.getvalue()

seed = by_variant['body-equation-inline-box']
_, seed_parts = unpack(seed)
sibling = E.fromstring(seed_parts['word/document.xml']).find(w + 'body').findall(w + 'p')[1]
source_image = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/math-composition-unnumbered.docx'
image_bytes = source_image.read_bytes()
assert sha(image_bytes) == by_variant['mixed-source-report']['SourceSha256']
with ZipFile(BytesIO(image_bytes)) as archive:
    image_tree = E.fromstring(archive.read('word/document.xml'))
    image_run = next(run for run in image_tree.iter(w + 'r') if run.find('.//' + a + 'blip') is not None)
    image_id = image_run.find('.//' + a + 'blip').get(r + 'embed')
    image_link = next(link for link in E.fromstring(archive.read('word/_rels/document.xml.rels')) if link.get('Id') == image_id)
    picture_bytes = archive.read('word/' + image_link.get('Target'))

prepared = []
cases = [seed]
for variant in ['first', 'mixed', 'wrap', 'cell', 'header', 'footer']:
    template = review_by_variant[variant if variant in ['cell', 'header', 'footer'] else 'body']
    infos, parts = unpack(template)
    paragraph = copy.deepcopy(sibling)
    drawing_run = next(run for run in paragraph.findall(w + 'r') if run.find(w + 'drawing') is not None)
    if variant == 'first':
        paragraph.remove(drawing_run)
        paragraph.insert(1 if paragraph.find(w + 'pPr') is not None else 0, drawing_run)
    elif variant == 'mixed':
        equation = copy.deepcopy(paragraph.find(m + 'oMath'))
        assert equation is not None
        paragraph.insert(list(paragraph).index(drawing_run), equation)
        picture = copy.deepcopy(image_run)
        picture.find('.//' + a + 'blip').set(r + 'embed', 'rSiblingPicture')
        picture.find('.//' + wp + 'docPr').set('id', '222')
        paragraph.insert(list(paragraph).index(drawing_run), picture)
        links = E.fromstring(parts['word/_rels/document.xml.rels'])
        E.SubElement(links, rel + 'Relationship', {'Id': 'rSiblingPicture',
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'Target': 'media/sibling-frame.png'})
        parts['word/_rels/document.xml.rels'] = E.tostring(links, encoding='utf-8', xml_declaration=True)
        parts['word/media/sibling-frame.png'] = picture_bytes
        types = E.fromstring(parts['[Content_Types].xml'])
        if not any(node.tag == ct + 'Default' and node.get('Extension') == 'png' for node in types):
            E.SubElement(types, ct + 'Default', {'Extension': 'png', 'ContentType': 'image/png'})
        parts['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
    elif variant == 'wrap':
        run = E.Element(w + 'r')
        E.SubElement(run, w + 't', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = ' Additional public explanation keeps the complete calculation frame together. '
        paragraph.insert(list(paragraph).index(drawing_run), run)
    part = 'word/inline-box-' + variant + '.xml' if variant in ['header', 'footer'] else 'word/document.xml'
    tree = E.fromstring(parts[part])
    if variant in ['header', 'footer']:
        parent = tree
    elif variant == 'cell':
        parent = tree.find(w + 'body').find(w + 'tbl').find('.//' + w + 'tc')
    else:
        parent = tree.find(w + 'body')
    old = next(node for node in parent.findall(w + 'p') if node.find('.//' + w + 'txbxContent') is not None)
    index = list(parent).index(old)
    parent.remove(old)
    parent.insert(index, paragraph)
    parts[part] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
    data = pack(infos, parts)
    id = 'inline-sibling-' + variant
    case = {'Id': id, 'Input': 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath/' + id + '.docx',
        'Sha256': sha(data), 'SourceSha256': template['Sha256'], 'SiblingSourceSha256': seed['Sha256'],
        'HeldOut': False, 'Variant': variant, 'MathPart': part,
        'ExpectedEquationCount': 2 if variant == 'mixed' else 1,
        'ExpectedRenderedEquationCount': 2 if variant == 'mixed' else 1,
        'ExpectedImageCount': 1 if variant == 'mixed' else 0,
        'ExpectedText': ['Before public math table.', 'Before inline', 'after equation.', 'Inline box content.', 'After public math table.']}
    cases.append(case)
    prepared.append((root / case['Input'], data))

manifest = directory / 'inline-sibling-frames.json'
manifest_bytes = (json.dumps(cases, indent=2) + '\n').replace('\n', '\r\n').encode()
if manifest.exists():
    assert json.loads(manifest.read_text()) == cases, 'Derived fixture identity differs'
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--write', action='store_true')
if parser.parse_args().write:
    for target, data in prepared + [(manifest, manifest_bytes)]:
        temporary = target.with_suffix(target.suffix + '.staging')
        with temporary.open('xb') as output:
            output.write(data)
        temporary.replace(target)
else:
    for target, data in prepared:
        assert target.read_bytes() == data, target.name
print('Canonical sibling source and six derived review/container fixtures verified.')
