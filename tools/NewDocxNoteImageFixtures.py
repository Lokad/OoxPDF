"""Public inline-image/math controls for direct normal note stories."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
from copy import deepcopy
import hashlib, json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
pkg = '{http://schemas.openxmlformats.org/package/2006/relationships}'
types = '{http://schemas.openxmlformats.org/package/2006/content-types}'
source = cases / 'Composition/report.docx'
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == '10F63D162D0D46DBE8CD682BB50452DB6672104706F116EA50CDCAFEE054575F'
with ZipFile(source) as archive:
    document = E.fromstring(archive.read('word/document.xml'))
    drawing = next(document.iter(w + 'drawing'))
    rels = E.fromstring(archive.read('word/_rels/document.xml.rels'))
    image_id = next(node for node in drawing.iter() if node.tag.endswith('}blip')).get(r + 'embed')
    image = next(node for node in rels if node.get('Id') == image_id)
    png = archive.read('word/' + image.get('Target'))
sources = json.loads((cases / 'CompositionProbes/math-note-content.json').read_text(encoding='utf-8'))
manifest = []
for kind in ['footnote', 'endnote']:
    for flow in ['inline', 'wrapped']:
        source = next(row for row in sources if row['Id'] == 'math-note-content-' + kind + '-' + flow)
        path = root / source['Input']
        assert hashlib.sha256(path.read_bytes()).hexdigest().upper() == source['Sha256']
        with ZipFile(path) as archive:
            entries = {name:archive.read(name) for name in archive.namelist()}
        notes = E.fromstring(entries['word/' + kind + 's.xml'])
        note = next(node for node in notes if int(node.get(w + 'id', '0')) > 0)
        paragraph = note.find(w + 'p')
        at = list(paragraph).index(next(paragraph.iter(m + 'oMath')))
        run = E.Element(w + 'r'); picture = deepcopy(drawing)
        next(node for node in picture.iter() if node.tag.endswith('}blip')).set(r + 'embed', 'NoteImage')
        if flow == 'wrapped':
            for node in picture.iter():
                if 'cx' in node.attrib and 'cy' in node.attrib:
                    for dimension in ['cx', 'cy']: node.set(dimension, str(int(node.get(dimension)) * 2))
        run.append(picture); paragraph.insert(at, run)
        entries['word/' + kind + 's.xml'] = E.tostring(notes, encoding='utf-8', xml_declaration=True)
        entries['word/media/note-composition.png'] = png
        part_rels = 'word/_rels/' + kind + 's.xml.rels'
        assert part_rels not in entries
        rels = E.Element(pkg + 'Relationships')
        E.SubElement(rels, pkg + 'Relationship', {'Id':'NoteImage', 'Type':r[1:-1] + '/image', 'Target':'media/note-composition.png'})
        entries[part_rels] = E.tostring(rels, encoding='utf-8', xml_declaration=True)
        content_types = E.fromstring(entries['[Content_Types].xml'])
        if not any(node.get('Extension') == 'png' for node in content_types):
            E.SubElement(content_types, types + 'Default', {'Extension':'png', 'ContentType':'image/png'})
        entries['[Content_Types].xml'] = E.tostring(content_types, encoding='utf-8', xml_declaration=True)
        identifier = kind + '-image-math' + ('-wrapped' if flow == 'wrapped' else '')
        target = cases / ('CompositionProbes/' + identifier + '.docx')
        with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
            for name,data in sorted(entries.items()):
                info = ZipInfo(name, (2026,10,16,0,0,0)); info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        manifest.append(dict(Id=identifier, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), HeldOut=False,
            ExpectedEquationCount=4, ExpectedInlineImageCount=1,
            ExpectedText=['Public calculation note.', 'End of supported note.'],
            Features=[kind + ' content', flow + ' prose', 'two note fractions', 'bookmark',
                      'inline PNG between note prose and math', 'two body fractions']))
(cases / 'CompositionProbes/math-note-images.json').write_text(json.dumps(manifest, indent=2)+'\n', encoding='utf-8')
print('Pinned four direct normal-note image/math controls; wrapped notes use a larger image.')
