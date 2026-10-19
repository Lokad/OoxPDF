"""Verify ordinary text/box/picture paragraphs moved into selected static stories."""
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
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()
manifests = ['inline-box-owned-paragraphs.json', 'inline-box-ordinary-pictures.json']
sources = {case['Id']: case for name in manifests for case in json.loads((directory / name).read_text())}
prepared, cases = [], []
for content, source_id in [('box', 'inline-box-owned-text-first'), ('picture', 'inline-box-picture-body')]:
    source = sources[source_id]
    original = (root / source['Input']).read_bytes()
    assert sha(original) == source['Sha256']
    with ZipFile(BytesIO(original)) as archive:
        infos = archive.infolist()
        original_parts = {info.filename: archive.read(info.filename) for info in infos}
    for kind in ['header', 'footer']:
        parts = original_parts.copy()
        document = E.fromstring(parts['word/document.xml'])
        body = document.find(w + 'body')
        paragraph = next(node for node in body if node.tag == w + 'p' and node.find('.//' + w + 'txbxContent') is not None)
        body.remove(paragraph)
        section = body.find(w + 'sectPr')
        for child in list(section):
            if child.tag in [w + 'headerReference', w + 'footerReference']:
                section.remove(child)
        id = 'static-inline-ordinary-' + kind + '-' + content
        part = 'word/' + id + '.xml'
        story = E.Element(w + ('hdr' if kind == 'header' else 'ftr'))
        story.append(paragraph)
        parts[part] = E.tostring(story, encoding='utf-8', xml_declaration=True)
        relationships = E.fromstring(parts['word/_rels/document.xml.rels'])
        source_ids = {node.get(r + 'embed') for node in paragraph.iter() if node.get(r + 'embed') is not None}
        story_relationships = E.Element(rel + 'Relationships')
        for node in relationships:
            if node.get('Id') in source_ids:
                story_relationships.append(copy.deepcopy(node))
        assert len(story_relationships) == (1 if content == 'picture' else 0)
        parts['word/_rels/' + id + '.xml.rels'] = E.tostring(story_relationships, encoding='utf-8', xml_declaration=True)
        relationship_id = 'rOrdinaryStaticInline'
        assert all(node.get('Id') != relationship_id for node in relationships)
        E.SubElement(relationships, rel + 'Relationship', Id=relationship_id, Target=id + '.xml',
            Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + kind)
        section.insert(0, E.Element(w + kind + 'Reference', {w + 'type': 'default', r + 'id': relationship_id}))
        parts['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
        parts['word/_rels/document.xml.rels'] = E.tostring(relationships, encoding='utf-8', xml_declaration=True)
        types = E.fromstring(parts['[Content_Types].xml'])
        E.SubElement(types, ct + 'Override', PartName='/' + part,
            ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + '+xml')
        parts['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
        output = BytesIO()
        with ZipFile(output, 'w') as archive:
            for info in infos:
                archive.writestr(copy.copy(info), parts[info.filename])
            for name in sorted(set(parts) - {info.filename for info in infos}):
                info = ZipInfo(name, (2026, 10, 19, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                archive.writestr(info, parts[name])
        data = output.getvalue()
        cases.append({'Id': id, 'Input': 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath/' + id + '.docx',
            'Sha256': sha(data), 'SourceSha256': source['Sha256'], 'HeldOut': True,
            'Variant': kind + '-' + content, 'MathPart': part, 'ExpectedEquationCount': 0,
            'ExpectedRenderedEquationCount': 0, 'ExpectedImageCount': 1 if content == 'picture' else 0,
            'ExpectedText': source['ExpectedText'],
            'Change': 'Move one source-exact ordinary paragraph into a selected header/footer; retain sibling order, complete box, picture resources and real body comment.'})
        prepared.append((root / cases[-1]['Input'], data))
manifest = directory / 'static-inline-ordinary.json'
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
print('Four source-exact ordinary mixed static-story controls verified.')
