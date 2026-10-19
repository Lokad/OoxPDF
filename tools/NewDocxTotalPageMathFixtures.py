"""Verify total-page controls; --write publishes inputs through atomic replacement."""
import argparse
import copy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--write', action='store_true')
probes = parser.add_mutually_exclusive_group()
probes.add_argument('--compound-report-probe', action='store_true')
probes.add_argument('--spanning-cell-probe', action='store_true')
args = parser.parse_args()
sha = lambda value: hashlib.sha256(value).hexdigest().upper()


def publish(jobs):
    for target, data in jobs:
        if args.write:
            stage = target.with_suffix(target.suffix + '.staging')
            with stage.open('xb') as output:
                output.write(data)
            stage.replace(target)
        else:
            assert target.read_bytes() == data, target.name


if args.spanning_cell_probe:
    parent = next(row for row in json.loads((cases.parent / 'InlineBoxMath/inline-box-total-counter-containers.json').read_text())
        if row['Variant'] == 'cell')
    source = root / parent['Input']
    original = source.read_bytes()
    assert sha(original) == parent['Sha256']
    with ZipFile(BytesIO(original)) as archive:
        infos = archive.infolist()
        parts = {entry.filename: archive.read(entry.filename) for entry in infos}
    tree = E.fromstring(parts['word/document.xml'])
    cell = next(cell for cell in tree.iter(w + 'tc')
        if any(field.get(w + 'instr', '').strip() == 'NUMPAGES' for field in cell.iter(w + 'fldSimple')))
    children = [E.tostring(child) for child in cell]
    assert all(field.get(w + 'instr', '').strip() == 'NUMPAGES' for field in cell.iter(w + 'fldSimple'))
    assert not list(tree.iter(w + 'cantSplit'))
    for index in range(1, 85):
        paragraph = E.SubElement(cell, w + 'p')
        E.SubElement(E.SubElement(paragraph, w + 'pPr'), w + 'spacing',
            {w + 'after': '0', w + 'line': '240', w + 'lineRule': 'auto'})
        run = E.SubElement(paragraph, w + 'r')
        props = E.SubElement(run, w + 'rPr')
        E.SubElement(props, w + 'rFonts', {w + 'ascii': 'Calibri', w + 'hAnsi': 'Calibri'})
        E.SubElement(props, w + 'sz', {w + 'val': '20'})
        E.SubElement(run, w + 't').text = 'Planning row ' + str(index) + ' stays visible.'
    assert [E.tostring(child) for child in list(cell)[:len(children)]] == children
    following = next(paragraph for paragraph in tree.find(w + 'body').findall(w + 'p')
        if ''.join(node.text or '' for node in paragraph.iter(w + 't')) == 'After public math table.')
    properties = following.find(w + 'pPr')
    if properties is None:
        properties = E.Element(w + 'pPr')
        following.insert(0, properties)
    E.SubElement(properties, w + 'pageBreakBefore')
    parts['word/document.xml'] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as archive:
        for entry in infos:
            archive.writestr(copy.copy(entry), parts[entry.filename])
    data = buffer.getvalue()
    name = 'spanning-cell-total-count'
    row = dict(parent, Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + name + '.docx',
        Sha256=sha(data), SourceInput=parent['Input'], SourceSha256=parent['Sha256'],
        SourceCellChildrenSha256=sha(b''.join(children)), HeldOut=True,
        AddedCellParagraphs=84, FieldInstruction='NUMPAGES',
        ExpectedText=parent['ExpectedText'] + ['Planning row 1 stays visible.', 'Planning row 84 stays visible.'],
        Change='Append ordinary source paragraphs to the accepted NUMPAGES-only boxed-math cell; preserve original cell items and move following prose to a new page.')
    publish([(cases / (name + '.docx'), data), (cases / (name + '.json'),
        (json.dumps([row], indent=2) + '\n').replace('\n', '\r\n').encode())])
    print('Source-exact spanning total-count cell verified; original math/fields unchanged.')
    raise SystemExit(0)

if args.compound_report_probe:
    manifest = cases.parent / 'Composition/composed-inline-frame-report.json'
    parent = json.loads(manifest.read_text())[0]
    source = root / parent['Input']
    original = source.read_bytes()
    assert sha(original) == parent['Sha256'] == '24982F90C9E3AD52FEA6ED01EE0BABB819E6C0F6FF00EBC28F72F41404129C32'
    with ZipFile(BytesIO(original)) as archive:
        infos = archive.infolist()
        parts = {entry.filename: archive.read(entry.filename) for entry in infos}
    document = E.fromstring(parts['word/document.xml'])
    title = document.find(w + 'body').find(w + 'p')
    assert ''.join(node.text or '' for node in title.iter(w + 't')) == 'Quarterly operations report'
    children = [E.tostring(child) for child in title]
    E.SubElement(E.SubElement(title, w + 'r'), w + 't',
        {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = ' '
    field = E.SubElement(title, w + 'fldSimple', {w + 'instr': 'NUMPAGES'})
    E.SubElement(E.SubElement(field, w + 'r'), w + 't').text = '1'
    assert [E.tostring(child) for child in list(title)[:-2]] == children
    parts['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as archive:
        for entry in infos:
            archive.writestr(copy.copy(entry), parts[entry.filename])
    data = buffer.getvalue()
    name = 'composed-inline-frame-total-count'
    row = dict(parent, Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/Composition/' + name + '.docx',
        Sha256=sha(data), SourceInput=parent['Input'], SourceSha256=parent['Sha256'],
        SourceTitleChildrenSha256=sha(b''.join(children)), FieldInstruction='NUMPAGES', CachedFieldValue='1',
        Change='Append one bare total-page field to the ordinary title; preserve existing source items and every other package part.')
    publish([(cases.parent / 'Composition' / (name + '.docx'), data),
        (cases.parent / 'Composition' / (name + '.json'),
            (json.dumps([row], indent=2) + '\n').replace('\n', '\r\n').encode())])
    print('Source-exact compound total-count control verified; prior title items/all other parts unchanged.')
    raise SystemExit(0)

source_pins = {row['Id']: row['SourceSha256'] for row in json.loads((cases / 'math-total-pages.json').read_text())}
rows = []
jobs = []
for form in ['simple', 'packed']:
    for container in ['body', 'cell']:
        source = cases / f'math-fields-{form}-{container}.docx'
        source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
        assert source_sha == source_pins[f'math-total-pages-{form}-{container}']
        with ZipFile(source) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        document = E.fromstring(entries['word/document.xml'])
        for field in document.iter(w + 'fldSimple'):
            field.set(w + 'instr', ' NUMPAGES ')
            for text in field.iter(w + 't'):
                text.text = '777'
        for field in document.iter(w + 'instrText'):
            field.text = ' NUMPAGES '
        for run in document.iter(w + 'r'):
            if run.find(w + 'instrText') is not None:
                for text in run.iter(w + 't'):
                    text.text = '777'
        following = next(p for p in document.iter(w + 'p')
            if any(t.text == 'Following prose.' for t in p.iter(w + 't')))
        props = following.find(w + 'pPr')
        if props is None:
            props = E.Element(w + 'pPr')
            following.insert(0, props)
        E.SubElement(props, w + 'pageBreakBefore')
        entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
        target = cases / f'math-total-pages-{form}-{container}.docx'
        buffer = BytesIO()
        with ZipFile(buffer, 'w', ZIP_DEFLATED) as archive:
            for name, data in entries.items():
                info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        data = buffer.getvalue()
        jobs.append((target, data))
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=sha(data),
            SourceSha256=source_sha, HeldOut=False, ExpectedEquationCount=2,
            ExpectedText=['Following prose.', 'Start', 'middle', 'End.']))
jobs.append((cases / 'math-total-pages.json', (json.dumps(rows, indent=2) + '\n').replace('\n', '\r\n').encode()))
publish(jobs)
print('Verified', len(rows), 'total-page/math controls')
