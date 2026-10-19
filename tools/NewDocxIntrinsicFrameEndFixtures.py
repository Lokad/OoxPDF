"""Verify four source-derived intrinsic-frame controls; --write publishes pinned bytes."""
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
a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
wp = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()
sources = json.loads((directory / 'inline-sibling-frames.json').read_text())
by_id = {case['Id']: case for case in sources}

def unpack(case):
    data = (root / case['Input']).read_bytes()
    assert sha(data) == case['Sha256']
    with ZipFile(BytesIO(data)) as archive:
        return archive.infolist(), {info.filename: archive.read(info.filename) for info in archive.infolist()}

def pack(infos, parts):
    output = BytesIO()
    with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
        for info in infos:
            archive.writestr(copy.copy(info), parts[info.filename])
        for name in sorted(set(parts) - {info.filename for info in infos}):
            info = ZipInfo(name, (2026, 10, 19, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, parts[name])
    return output.getvalue()

def xml_bytes(tree):
    return E.tostring(tree, encoding='utf-8', xml_declaration=True)

prepared = []
cases = list(sources)
for variant in ['cell-picture', 'cell-space', 'cell-size12', 'wrap-size12']:
    source = by_id['inline-sibling-wrap' if variant.startswith('wrap') else 'inline-sibling-cell']
    infos, parts = unpack(source)
    if variant.endswith('size12'):
        styles = E.fromstring(parts['word/styles.xml'])
        for node in styles.iter():
            if node.tag in [w + 'sz', w + 'szCs'] and node.get(w + 'val') == '22':
                node.set(w + 'val', '24')
        parts['word/styles.xml'] = xml_bytes(styles)
    else:
        document = E.fromstring(parts['word/document.xml'])
        paragraph = next(p for p in document.iter(w + 'p') if p.find('.//' + w + 'txbxContent') is not None)
        box_run = next(run for run in paragraph.findall(w + 'r') if run.find('.//' + w + 'txbxContent') is not None)
        if variant == 'cell-space':
            run = E.SubElement(paragraph, w + 'r')
            E.SubElement(run, w + 't', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = ' '
        else:
            _, mixed_parts = unpack(by_id['inline-sibling-mixed'])
            mixed = E.fromstring(mixed_parts['word/document.xml'])
            picture = copy.deepcopy(next(run for run in mixed.iter(w + 'r') if run.find('.//' + a + 'blip') is not None))
            picture.find('.//' + a + 'blip').set(r + 'embed', 'rEndFramePicture')
            for extent in [picture.find('.//' + wp + 'extent'), picture.find('.//' + a + 'xfrm/' + a + 'ext')]:
                extent.set('cx', '2743200')
                extent.set('cy', '762000')
            paragraph.insert(list(paragraph).index(box_run), picture)
            paragraph.remove(box_run)
            links = E.fromstring(parts['word/_rels/document.xml.rels'])
            E.SubElement(links, rel + 'Relationship', {'Id': 'rEndFramePicture',
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'Target': 'media/end-frame.png'})
            parts['word/_rels/document.xml.rels'] = xml_bytes(links)
            parts['word/media/end-frame.png'] = mixed_parts['word/media/sibling-frame.png']
            types = E.fromstring(parts['[Content_Types].xml'])
            if not any(node.tag == ct + 'Default' and node.get('Extension') == 'png' for node in types):
                E.SubElement(types, ct + 'Default', {'Extension': 'png', 'ContentType': 'image/png'})
            parts['[Content_Types].xml'] = xml_bytes(types)
        parts['word/document.xml'] = xml_bytes(document)
    data = pack(infos, parts)
    id = 'inline-end-' + variant
    expected = list(source['ExpectedText'])
    if variant == 'cell-picture':
        expected.remove('Inline box content.')
    case = {'Id': id, 'Input': 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath/' + id + '.docx',
        'Sha256': sha(data), 'SourceSha256': source['Sha256'], 'HeldOut': True, 'Variant': variant,
        'MathPart': 'word/document.xml', 'ExpectedEquationCount': 1, 'ExpectedRenderedEquationCount': 1,
        'ExpectedImageCount': 1 if variant == 'cell-picture' else 0, 'ExpectedText': expected}
    cases.append(case)
    prepared.append((root / case['Input'], data))

manifest = directory / 'inline-frame-ends.json'
manifest_bytes = (json.dumps(cases, indent=2) + '\n').replace('\n', '\r\n').encode()
if manifest.exists():
    assert json.loads(manifest.read_text()) == cases, 'Derived manifest differs'
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
print('Seven existing and four source-derived intrinsic-frame controls verified.')
