"""Verify authored-row hyperlink controls; --write publishes source-exact inputs."""
import argparse
import copy
import hashlib
import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parents[1]
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
source = cases / 'docx-omml-isolated-inline-body-fraction-left.docx'
sha = lambda value: hashlib.sha256(value).hexdigest().upper()
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--write', action='store_true')
parser.add_argument('--containers', action='store_true')
parser.add_argument('--pictures', action='store_true', help='Picture navigation encodings; authored math controls stay unchanged.')
parser.add_argument('--compound-pictures', action='store_true', help='Add only canonical click metadata to the held-out composed report.')
parser.add_argument('--footnote-pictures', action='store_true', help='Add canonical clicks to the existing mixed-footnote holdout.')
args = parser.parse_args()
if sum((args.pictures, args.containers, args.compound_pictures, args.footnote_pictures)) > 1:
    parser.error('Select one probe profile')

if args.compound_pictures or args.footnote_pictures:
    directory = cases / ('CompositionProbes' if args.footnote_pictures else 'Composition')
    if args.footnote_pictures:
        golden = json.loads((directory / 'linked-mixed-note-office.json').read_text(encoding='utf-8-sig'))
        parent = dict(Id='linked-mixed-note-report', Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/linked-mixed-note-report.docx',
            Sha256=golden['InputSha256'], HeldOut=True, ExpectedInlineImageCount=golden['ImagePlacements'],
            ExpectedText=['Public continuation note line 01:', 'Public continuation note line 64:', 'Second public continuation note.'])
    else:
        parent = json.loads((directory / 'composed-inline-frame-report.json').read_text())[0]
    original = (root / parent['Input']).read_bytes()
    assert sha(original) == parent['Sha256']
    with ZipFile(BytesIO(original)) as archive:
        infos = archive.infolist()
        parts = {entry.filename: archive.read(entry.filename) for entry in infos}
    if args.footnote_pictures:
        parent['ExpectedEquationCount'] = sum(sum(1 for _ in E.fromstring(payload).iter(m + 'oMath'))
            for name, payload in parts.items() if name.startswith('word/') and name.endswith('.xml'))
    changed, owners = parts.copy(), []
    a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    wp = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
    pic = '{http://schemas.openxmlformats.org/drawingml/2006/picture}'
    rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    office_rel = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    for name in sorted(parts):
        if not name.startswith('word/') or not name.endswith('.xml'): continue
        tree = E.fromstring(parts[name])
        pictures = [item for item in tree.iter(wp + 'inline') if item.find('.//' + pic + 'pic') is not None]
        if not pictures: continue
        before = E.tostring(tree)
        rel_name = str(Path(name).parent / '_rels' / (Path(name).name + '.rels')).replace('\\', '/')
        relationships = E.fromstring(parts[rel_name])
        before_relationships = E.tostring(relationships)
        added = []
        for index, picture in enumerate(pictures):
            properties = picture.find(wp + 'docPr')
            assert properties is not None and properties.find(a + 'hlinkClick') is None
            relation_id = 'rIdPublicCompoundPicture' + str(index + 1)
            assert all(item.get('Id') != relation_id for item in relationships)
            uri = 'https://example.invalid/' + ('footnote-picture/' if args.footnote_pictures else 'composed-picture/') + Path(name).stem + '/' + properties.get('id')
            click = E.SubElement(properties, a + 'hlinkClick', {office_rel + 'id': relation_id})
            relation = E.SubElement(relationships, rel + 'Relationship', {'Id': relation_id,
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
                'Target': uri, 'TargetMode': 'External'})
            added.append((properties, click, relation))
            owners.append({'Part': name, 'PictureId': properties.get('id'), 'TargetUri': uri,
                'ExpectedNativeNavigation': True,
                'Scope': 'Complete inline picture; container owns physical placement and clip'})
        changed[name] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
        changed[rel_name] = E.tostring(relationships, encoding='utf-8', xml_declaration=True)
        for properties, click, relation in added:
            properties.remove(click)
            relationships.remove(relation)
        assert E.tostring(tree) == before and E.tostring(relationships) == before_relationships
    assert len(owners) == (2 if args.footnote_pictures else 5) and all(item['ExpectedNativeNavigation'] for item in owners)
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as archive:
        for entry in infos: archive.writestr(copy.copy(entry), changed[entry.filename])
    generated = buffer.getvalue()
    identifier = 'linked-footnote-picture-report' if args.footnote_pictures else 'composed-picture-navigation-report'
    row = copy.deepcopy(parent)
    row.update(Id=identifier, Input=(directory.relative_to(root) / (identifier + '.docx')).as_posix(),
        Sha256=sha(generated), SourceInput=parent['Input'], SourceSha256=parent['Sha256'], PictureLinkOwners=owners,
        Change=('Add only canonical picture click targets/relationships; preserve all visible source content and geometry.' if args.footnote_pictures else
            'Add only five canonical picture click targets/relationships; preserve all visible source content and geometry.'))
    jobs = [(directory / (identifier + '.docx'), generated),
        (directory / ('footnote-picture-links.json' if args.footnote_pictures else identifier + '.json'),
            (json.dumps([row], indent=2) + '\n').replace('\n', '\r\n').encode())]
    for target, generated in jobs:
        if args.write:
            stage = target.with_suffix(target.suffix + '.staging')
            with stage.open('xb') as output: output.write(generated)
            stage.replace(target)
        else: assert target.read_bytes() == generated, target.name
    print(str(len(owners)) + ' complete report picture owners verified; visible source content unchanged.')
    raise SystemExit(0)

if args.pictures:
    directory = cases / 'InlineBoxMath'
    parents = json.loads((directory / 'inline-box-ordinary-pictures.json').read_text())
    a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    wp = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
    pic = '{http://schemas.openxmlformats.org/drawingml/2006/picture}'
    rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    office_rel = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    uri = 'https://example.invalid/public-picture'
    rows, jobs = [], []
    for variant, parent_id in [('wrapper-body', 'inline-box-picture-body'),
                               ('properties-body', 'inline-box-picture-body'),
                               ('properties-cell', 'inline-box-picture-cell')]:
        parent = next(item for item in parents if item['Id'] == parent_id)
        original = (root / parent['Input']).read_bytes()
        assert sha(original) == parent['Sha256']
        with ZipFile(BytesIO(original)) as archive:
            infos = archive.infolist()
            parts = {entry.filename: archive.read(entry.filename) for entry in infos}
        tree = E.fromstring(parts['word/document.xml'])
        before = E.tostring(tree)
        run = next(node for node in tree.iter(w + 'r') if node.find('.//' + pic + 'pic') is not None)
        owner = next(node for node in tree.iter() if run in list(node))
        position = list(owner).index(run)
        relation_id = 'rIdPublicPictureNavigation'
        if variant.startswith('wrapper'):
            owner.remove(run)
            wrapper = E.Element(w + 'hyperlink', {office_rel + 'id': relation_id})
            wrapper.append(run)
            owner.insert(position, wrapper)
        else:
            properties = run.find('.//' + wp + 'docPr')
            assert properties is not None and properties.find(a + 'hlinkClick') is None
            click = E.SubElement(properties, a + 'hlinkClick', {office_rel + 'id': relation_id})
        changed = E.tostring(tree, encoding='utf-8', xml_declaration=True)
        if variant.startswith('wrapper'):
            owner.remove(wrapper)
            owner.insert(position, run)
        else:
            properties.remove(click)
        assert E.tostring(tree) == before
        relationships = E.fromstring(parts['word/_rels/document.xml.rels'])
        original_relationships = E.tostring(relationships)
        assert all(item.get('Id') != relation_id for item in relationships)
        added = E.SubElement(relationships, rel + 'Relationship', {'Id': relation_id,
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            'Target': uri, 'TargetMode': 'External'})
        changed_relationships = E.tostring(relationships, encoding='utf-8', xml_declaration=True)
        relationships.remove(added)
        assert E.tostring(relationships) == original_relationships
        buffer = BytesIO()
        with ZipFile(buffer, 'w') as archive:
            for entry in infos:
                payload = changed if entry.filename == 'word/document.xml' else (
                    changed_relationships if entry.filename == 'word/_rels/document.xml.rels' else parts[entry.filename])
                archive.writestr(copy.copy(entry), payload)
        generated = buffer.getvalue()
        identifier = 'inline-picture-link-' + variant
        row = copy.deepcopy(parent)
        row.update(Id=identifier, Input='tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath/' + identifier + '.docx',
                   Sha256=sha(generated), SourceInput=parent['Input'], SourceSha256=parent['Sha256'],
                   Variant=variant, TargetUri=uri, HeldOut=True,
                   Change='Add only the picture hyperlink metadata/relationship; retain visible source content and geometry.')
        rows.append(row)
        jobs.append((directory / (identifier + '.docx'), generated))
    jobs.append((directory / 'inline-picture-links.json', (json.dumps(rows, indent=2) + '\n').replace('\n', '\r\n').encode()))
    for target, generated in jobs:
        if args.write:
            stage = target.with_suffix(target.suffix + '.staging')
            with stage.open('xb') as output:
                output.write(generated)
            stage.replace(target)
        else:
            assert target.read_bytes() == generated, target.name
    print('Three source-preserving picture navigation controls verified.')
    raise SystemExit(0)
data = source.read_bytes()
assert sha(data) == '96A12C11555C1CF0C9360F0939551E77BE12180A69411CC534D56297E7B25B4F'
with ZipFile(BytesIO(data)) as archive:
    infos = archive.infolist()
    parts = {entry.filename: archive.read(entry.filename) for entry in infos}
base = E.fromstring(parts['word/document.xml'])
formula = next(base.iter(m + 'oMath'))
assert len(list(base.iter(m + 'oMath'))) == 1
source_math = E.tostring(formula)
rows, jobs = [], []
if args.containers:
    parent = next(row for row in json.loads((cases / 'CompositionProbes/math-authored-row-links.json').read_text())
        if row['Variant'] == 'linked-prose')
    data = (root / parent['Input']).read_bytes()
    assert sha(data) == parent['Sha256']
    with ZipFile(BytesIO(data)) as archive:
        infos = archive.infolist()
        parts = {entry.filename: archive.read(entry.filename) for entry in infos}
    for container in ['cell', 'header', 'footer']:
        tree = E.fromstring(parts['word/document.xml'])
        body = tree.find(w + 'body')
        paragraph = next(p for p in body.findall(w + 'p') if list(p.iter(m + 'oMath')))
        paragraph_hash = sha(E.tostring(paragraph))
        index = list(body).index(paragraph)
        body.remove(paragraph)
        changed = parts.copy()
        part = 'word/document.xml'
        if container == 'cell':
            table = E.Element(w + 'tbl')
            properties = E.SubElement(table, w + 'tblPr')
            E.SubElement(properties, w + 'tblW', {w + 'w': '6000', w + 'type': 'dxa'})
            E.SubElement(properties, w + 'tblLayout', {w + 'type': 'fixed'})
            margins = E.SubElement(properties, w + 'tblCellMar')
            for side, value in [('top', '0'), ('bottom', '0'), ('left', '108'), ('right', '108')]:
                E.SubElement(margins, w + side, {w + 'w': value, w + 'type': 'dxa'})
            E.SubElement(E.SubElement(table, w + 'tblGrid'), w + 'gridCol', {w + 'w': '6000'})
            cell = E.SubElement(E.SubElement(table, w + 'tr'), w + 'tc')
            E.SubElement(E.SubElement(cell, w + 'tcPr'), w + 'tcW', {w + 'w': '6000', w + 'type': 'dxa'})
            cell.append(paragraph)
            body.insert(index, table)
        else:
            part = 'word/math-row-' + container + '.xml'
            story = E.Element(w + ('hdr' if container == 'header' else 'ftr'))
            story.append(paragraph)
            changed[part] = E.tostring(story, encoding='utf-8', xml_declaration=True)
            section = body.find(w + 'sectPr')
            assert section.find(w + container + 'Reference') is None
            section.find(w + 'pgMar').set(w + container, '180')
            rid = 'rMathRow' + container.title()
            E.SubElement(section, w + container + 'Reference', {w + 'type': 'default',
                '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': rid})
            rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
            links = E.fromstring(parts['word/_rels/document.xml.rels'])
            E.SubElement(links, rel + 'Relationship', {'Id': rid,
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + container,
                'Target': 'math-row-' + container + '.xml'})
            changed['word/_rels/document.xml.rels'] = E.tostring(links, encoding='utf-8', xml_declaration=True)
            ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
            types = E.fromstring(parts['[Content_Types].xml'])
            E.SubElement(types, ct + 'Override', {'PartName': '/' + part,
                'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + container + '+xml'})
            changed['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
        assert sha(E.tostring(paragraph)) == paragraph_hash
        changed['word/document.xml'] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
        buffer = BytesIO()
        with ZipFile(buffer, 'w') as result:
            for entry in infos:
                result.writestr(copy.copy(entry), changed[entry.filename])
            for name in sorted(set(changed) - {entry.filename for entry in infos}):
                from zipfile import ZipInfo, ZIP_DEFLATED
                info = ZipInfo(name, (2026, 10, 19, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                result.writestr(info, changed[name])
        generated = buffer.getvalue()
        name = 'math-authored-row-linked-prose-' + container
        rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + name + '.docx',
            Sha256=sha(generated), SourceInput=parent['Input'], SourceSha256=parent['Sha256'],
            SourceParagraphSha256=paragraph_hash, MathPart=part, Variant='linked-prose', Container=container,
            HeldOut=True, ExpectedEquationCount=1, ExpectedRenderedEquationCount=1,
            ExpectedText=['Before public math table.', 'Reference row.', 'After public math table.'],
            Change='Move the unchanged linked-prose/implicit-math paragraph into a ' + container + ' source frame.'))
        jobs.append((cases / 'CompositionProbes' / (name + '.docx'), generated))
    assert len({row['SourceParagraphSha256'] for row in rows}) == 1
    jobs.append((cases / 'CompositionProbes/math-authored-row-link-containers.json',
        (json.dumps(rows, indent=2) + '\n').replace('\n', '\r\n').encode()))
    for target, generated in jobs:
        if args.write:
            stage = target.with_suffix(target.suffix + '.staging')
            with stage.open('xb') as output:
                output.write(generated)
            stage.replace(target)
        else:
            assert target.read_bytes() == generated, target.name
    print('Three source-exact authored-row hyperlink container controls verified.')
    raise SystemExit(0)

for variant in ['unlinked', 'linked-prose', 'linked-math']:
    tree = copy.deepcopy(base)
    paragraph = next(p for p in tree.iter(w + 'p') if list(p.iter(m + 'oMath')))
    formula = next(paragraph.iter(m + 'oMath'))
    assert formula in list(paragraph), 'The control requires implicit inline source math.'
    at = list(paragraph).index(formula)
    paragraph.remove(formula)
    text = E.Element(w + 'r')
    E.SubElement(text, w + 't').text = 'Reference row.'
    link = E.Element(w + 'hyperlink', {w + 'anchor': 'RowCalculation'})
    link.append(text if variant == 'linked-prose' else formula)
    line_break = E.Element(w + 'r')
    E.SubElement(line_break, w + 'br')
    content = [E.Element(w + 'bookmarkStart', {w + 'id': '108', w + 'name': 'RowCalculation'}),
        E.Element(w + 'bookmarkEnd', {w + 'id': '108'}),
        link if variant == 'linked-prose' else text,
        line_break, link if variant == 'linked-math' else formula]
    for offset, child in enumerate(content):
        paragraph.insert(at + offset, child)
    assert E.tostring(next(tree.iter(m + 'oMath'))) == source_math
    changed = parts.copy()
    changed['word/document.xml'] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as result:
        for entry in infos:
            result.writestr(copy.copy(entry), changed[entry.filename])
    generated = buffer.getvalue()
    name = 'math-authored-row-' + variant
    rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + name + '.docx',
        Sha256=sha(generated), SourceInput=source.relative_to(root).as_posix(), SourceSha256=sha(data),
        SourceMathSha256=sha(source_math), Variant=variant, HeldOut=variant != 'unlinked',
        ExpectedEquationCount=1, ExpectedRenderedEquationCount=1, ExpectedText=['Reference row.'],
        Change='Add the same authored text row before an unchanged implicit equation; hyperlink ownership is ' + variant + '.'))
    jobs.append((cases / 'CompositionProbes' / (name + '.docx'), generated))
jobs.append((cases / 'CompositionProbes/math-authored-row-links.json',
    (json.dumps(rows, indent=2) + '\n').replace('\n', '\r\n').encode()))
for target, generated in jobs:
    if args.write:
        stage = target.with_suffix(target.suffix + '.staging')
        with stage.open('xb') as output:
            output.write(generated)
        stage.replace(target)
    else:
        assert target.read_bytes() == generated, target.name
print('Three source-exact authored-row hyperlink controls verified.')
