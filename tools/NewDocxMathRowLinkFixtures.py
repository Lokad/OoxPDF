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
args = parser.parse_args()
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
