"""Compose already-qualified inline source fragments into the public report."""
import copy
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E

w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
wp = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()


def add_source_frames(entries, root):
    directory = root / 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath'
    sources = {case['Id']: case for name in ['inline-sibling-frames.json', 'static-inline-ordinary.json']
        for case in json.loads((directory / name).read_text())}
    selected = [('word/document.xml', 'inline-sibling-mixed', 'word/document.xml'),
        ('word/header1.xml', 'static-inline-ordinary-header-box', 'word/static-inline-ordinary-header-box.xml'),
        ('word/footer1.xml', 'static-inline-ordinary-footer-picture', 'word/static-inline-ordinary-footer-picture.xml')]
    next_id = max(int(node.get('id')) for name, data in entries.items()
        if name.startswith('word/') and name.endswith('.xml')
        for node in E.fromstring(data).iter(wp + 'docPr')) + 1
    pins = []
    for part, id, source_part in selected:
        source = sources[id]
        source_bytes = (root / source['Input']).read_bytes()
        assert sha(source_bytes) == source['Sha256']
        with ZipFile(root / source['Input']) as archive:
            source_tree = E.fromstring(archive.read(source_part))
            paragraph = copy.deepcopy(next(node for node in source_tree.iter(w + 'p')
                if node.find('.//' + w + 'txbxContent') is not None))
            assert not any(node.tag in [w + 'commentReference', w + 'commentRangeStart', w + 'commentRangeEnd',
                w + 'footnoteReference', w + 'endnoteReference', w + 'hyperlink'] for node in paragraph.iter())
            for node in paragraph.iter(wp + 'docPr'):
                node.set('id', str(next_id)); node.set('name', 'Public report inline frame ' + str(next_id))
                next_id += 1
            relationship_part = str(Path(source_part).parent / '_rels' / (Path(source_part).name + '.rels')).replace('\\', '/')
            old_relationships = E.fromstring(archive.read(relationship_part))
            target_relationship_part = str(Path(part).parent / '_rels' / (Path(part).name + '.rels')).replace('\\', '/')
            relationships = E.fromstring(entries[target_relationship_part]) if target_relationship_part in entries else E.Element(rel + 'Relationships')
            for node in paragraph.iter():
                if node.get(r + 'embed') is None:
                    continue
                image = next(item for item in old_relationships if item.get('Id') == node.get(r + 'embed'))
                image_bytes = archive.read('word/' + image.get('Target'))
                media = 'word/media/composed-inline-frame.png'
                assert media not in entries or entries[media] == image_bytes
                entries[media] = image_bytes
                relationship_id = 'ComposedInlineFrameImage'
                assert not any(item.get('Id') == relationship_id for item in relationships)
                E.SubElement(relationships, rel + 'Relationship', Id=relationship_id,
                    Type=image.get('Type'), Target='media/composed-inline-frame.png')
                node.set(r + 'embed', relationship_id)
            entries[target_relationship_part] = E.tostring(relationships, encoding='utf-8', xml_declaration=True)
        tree = E.fromstring(entries[part])
        if part == 'word/document.xml':
            body = tree.find(w + 'body')
            body.insert(list(body).index(body.find(w + 'tbl')), paragraph)
        else:
            # Keep the existing trailing paragraph and all original table content.
            tree.insert(max(0, len(tree) - 1), paragraph)
        entries[part] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
        pins.append({'Input': source['Input'], 'Sha256': source['Sha256'], 'DestinationPart': part})
    return pins
