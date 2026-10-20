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
probes.add_argument('--spanning-frame-probes', action='store_true')
probes.add_argument('--parallel-cell-probes', action='store_true')
probes.add_argument('--parallel-cell-bordered-probe', action='store_true')
probes.add_argument('--parallel-cell-short-probe', action='store_true')
probes.add_argument('--parallel-cell-box-probe', action='store_true')
probes.add_argument('--parallel-cell-inset-probes', action='store_true')
probes.add_argument('--parallel-cell-bottom-inset-probe', action='store_true')
probes.add_argument('--parallel-cell-box-inset-probe', action='store_true')
probes.add_argument('--parallel-cell-border-probes', action='store_true')
probes.add_argument('--parallel-cell-trailing-row-probe', action='store_true')
probes.add_argument('--boxed-inset-trailing-row-probe', action='store_true')
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


if args.boxed_inset_trailing_row_probe:
    parent = json.loads((cases / 'parallel-cell-box-inset-probe.json').read_text())[0]
    original = (root / parent['Input']).read_bytes()
    assert sha(original) == parent['Sha256']
    with ZipFile(BytesIO(original)) as archive:
        infos = archive.infolist()
        parts = {entry.filename: archive.read(entry.filename) for entry in infos}
    tree = E.fromstring(parts['word/document.xml'])
    table = next(tree.iter(w + 'tbl'))
    source_row = table.find(w + 'tr')
    source_row_bytes = E.tostring(source_row)
    tail = copy.deepcopy(source_row)
    for side, cell in zip(('Left', 'Right'), tail.findall(w + 'tc')):
        ordinary = copy.deepcopy(cell.findall(w + 'p')[-1])
        assert not any(node.tag.endswith(('oMath', 'oMathPara', 'drawing', 'pict', 'fldSimple', 'instrText')) for node in ordinary.iter())
        for child in list(cell):
            if child.tag != w + 'tcPr': cell.remove(child)
        ordinary.find('.//' + w + 't').text = side + ' terminal control.'
        cell.append(ordinary)
    table.append(tail)
    assert E.tostring(source_row) == source_row_bytes
    parts['word/document.xml'] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as archive:
        for entry in infos: archive.writestr(copy.copy(entry), parts[entry.filename])
    data = buffer.getvalue()
    row = copy.deepcopy(parent)
    identifier = 'parallel-cell-boxed-inset-tail'
    row.update(Id=identifier, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + identifier + '.docx',
        Sha256=sha(data), CompositionParentInput=parent['Input'], CompositionParentSha256=parent['Sha256'],
        UnchangedFirstRowSha256=sha(source_row_bytes), LongRowSourceRowsByCell=dict(Left=120, Right=120),
        ExpectedText=parent['ExpectedText']+['Left terminal control.','Right terminal control.'],
        Change='Append one ordinary row to accepted shaded/inset parallel prose with boxed math and NUMPAGES; the complete first-row XML is unchanged.')
    row.pop('SourceRowsPerCell', None)
    manifest = (json.dumps([row], indent=2) + '\n').replace('\n', '\r\n').encode()
    publish([(cases / (identifier + '.docx'), data), (cases / 'parallel-cell-boxed-tail-probe.json', manifest)])
    print('Boxed math/count/inset trailing-row control verified; source first row unchanged.')
    raise SystemExit(0)


parallel_group = next((group for enabled, group in [
    (args.parallel_cell_probes, ('parallel-cell-probes.json', ('aligned', 'shifted', 'mixed'))),
    (args.parallel_cell_bordered_probe, ('parallel-cell-bordered-probe.json', ('bordered',))),
    (args.parallel_cell_short_probe, ('parallel-cell-short-probe.json', ('short-left',))),
    (args.parallel_cell_box_probe, ('parallel-cell-box-probe.json', ('boxed-math',))),
    (args.parallel_cell_inset_probes, ('parallel-cell-inset-probes.json', ('uniform-insets', 'unequal-insets'))),
    (args.parallel_cell_bottom_inset_probe, ('parallel-cell-bottom-inset-probe.json', ('bottom-insets',))),
    (args.parallel_cell_box_inset_probe, ('parallel-cell-box-inset-probe.json', ('boxed-insets',))),
    (args.parallel_cell_border_probes, ('parallel-cell-border-probes.json', ('bordered-aligned', 'bordered-thick'))),
    (args.parallel_cell_trailing_row_probe, ('parallel-cell-trailing-row-probe.json', ('bordered-tail',))),
] if enabled), None)
if parallel_group:
    parent = next(row for row in json.loads((cases / 'spanning-frame-probes.json').read_text())
        if row['Id'] == 'spanning-frame-mixed-prose')
    original = (root / parent['Input']).read_bytes()
    assert sha(original) == parent['Sha256']
    with ZipFile(BytesIO(original)) as archive:
        infos = archive.infolist()
        original_parts = {entry.filename: archive.read(entry.filename) for entry in infos}
    rows, jobs = [], []
    manifest_name, variants = parallel_group
    source_insets = {
        'uniform-insets': dict(Left=(3, 3), Right=(3, 3)),
        'unequal-insets': dict(Left=(3, 1), Right=(6, 3)),
        'bottom-insets': dict(Left=(3, 18), Right=(3, 3)),
        'boxed-insets': dict(Left=(3, 3), Right=(3, 3)),
    }
    source_borders = {'bordered': (8, True), 'bordered-aligned': (8, False), 'bordered-thick': (16, True), 'bordered-tail': (8, True)}
    for variant in variants:
        parts = dict(original_parts)
        tree = E.fromstring(parts['word/document.xml'])
        table = next(tree.iter(w + 'tbl'))
        widths = (240, 160) if variant in ('boxed-math', 'boxed-insets') else (160, 160)
        table.find(w + 'tblPr/' + w + 'tblW').set(w + 'w', str(sum(widths) * 20))
        if variant in source_borders:
            E.SubElement(table.find(w + 'tblPr'), w + 'tblInd', {w + 'w': '0', w + 'type': 'dxa'})
        grid = table.find(w + 'tblGrid')
        grid.clear()
        for width in widths: E.SubElement(grid, w + 'gridCol', {w + 'w': str(width * 20)})
        row_node = table.find(w + 'tr')
        source_cell = row_node.find(w + 'tc')
        for cell in row_node.findall(w + 'tc'): row_node.remove(cell)
        for side, width in zip(('Left', 'Right'), widths):
            cell = copy.deepcopy(source_cell)
            cell.find(w + 'tcPr/' + w + 'tcW').set(w + 'w', str(width * 20))
            if variant in source_borders:
                borders = E.SubElement(cell.find(w + 'tcPr'), w + 'tcBorders')
                for edge in ('top', 'left', 'bottom', 'right'):
                    E.SubElement(borders, w + edge, {w + 'val': 'single', w + 'sz': str(source_borders[variant][0]), w + 'color': '000000'})
            if variant in source_insets:
                properties = cell.find(w + 'tcPr')
                margins = properties.find(w + 'tcMar')
                if margins is None: margins = E.SubElement(properties, w + 'tcMar')
                top, bottom = source_insets[variant][side]
                for edge, points in [('top', top), ('bottom', bottom)]:
                    margin = margins.find(w + edge)
                    if margin is None: margin = E.SubElement(margins, w + edge)
                    margin.set(w + 'w', str(points * 20))
                    margin.set(w + 'type', 'dxa')
                E.SubElement(properties, w + 'shd', {w + 'val': 'clear', w + 'fill': 'DDEBF7' if side == 'Left' else 'E2F0D9'})
            paragraphs = cell.findall(w + 'p')
            prefix = paragraphs[0]
            prefix.find('.//' + w + 't').text = side + ' ordinary introduction.'
            properties = prefix.find(w + 'pPr')
            if properties is None:
                properties = E.Element(w + 'pPr')
                prefix.insert(0, properties)
            E.SubElement(properties, w + 'spacing', {w + 'after': '320' if side == 'Right' and (variant == 'shifted' or source_borders.get(variant, (0, False))[1]) else '160'})
            for index, paragraph in enumerate(paragraphs[1:], 1):
                paragraph.find('.//' + w + 't').text = side + ' source row ' + str(index) + '.'
                paragraph.find('.//' + w + 'sz').set(w + 'val',
                    '24' if side == 'Right' and variant in ('mixed', 'short-left') and index % 2 else '20')
            if side == 'Left' and variant == 'short-left':
                for paragraph in paragraphs[21:]: cell.remove(paragraph)
            if side == 'Left' and variant in ('boxed-math', 'boxed-insets'):
                box_parent = json.loads((cases / 'spanning-cell-total-count.json').read_text())[0]
                box_bytes = (root / box_parent['Input']).read_bytes()
                assert sha(box_bytes) == box_parent['Sha256']
                with ZipFile(BytesIO(box_bytes)) as archive:
                    box_tree = E.fromstring(archive.read('word/document.xml'))
                box_paragraph = copy.deepcopy(next(box_tree.iter(w + 'tc')).find(w + 'p'))
                cell.insert(list(cell).index(prefix), box_paragraph)
                cell.remove(prefix)
            row_node.append(cell)
        if variant == 'bordered-tail':
            composition_parent = cases / 'parallel-cell-bordered.docx'
            with ZipFile(composition_parent) as archive:
                parent_row = next(E.fromstring(archive.read('word/document.xml')).iter(w + 'tr'))
            assert E.tostring(row_node) == E.tostring(parent_row), 'The complete accepted source row must remain unchanged'
            source_row_sha = sha(E.tostring(row_node))
            tail = copy.deepcopy(row_node)
            for side, cell in zip(('Left', 'Right'), tail.findall(w + 'tc')):
                paragraphs = cell.findall(w + 'p')
                for paragraph in paragraphs[1:]: cell.remove(paragraph)
                paragraphs[0].find('.//' + w + 't').text = side + ' terminal control.'
            table.append(tail)
        parts['word/document.xml'] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
        buffer = BytesIO()
        with ZipFile(buffer, 'w') as archive:
            for entry in infos: archive.writestr(copy.copy(entry), parts[entry.filename])
        data = buffer.getvalue()
        name = 'parallel-cell-' + variant
        rows.append(dict(Id=name,Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + name + '.docx',
            Sha256=sha(data),SourceInput=parent['Input'],SourceSha256=parent['Sha256'],HeldOut=True,
            ExpectedEquationCount=0,ExpectedRenderedEquationCount=0,SourceRowsPerCell=120,
            ExpectedText=['Before public math table.','Left source row 1.','Right source row 1.',
                'Left source row ' + ('20.' if variant == 'short-left' else '120.'),'Right source row 120.','After public math table.'],
            Change='Parallel ordinary cells: aligned source metrics, different introduction spacing, or different line metrics. No candidate-derived geometry.'))
        if variant == 'bordered-tail':
            rows[-1].update(CompositionParentInput=str(composition_parent.relative_to(root)).replace('\\', '/'),
                CompositionParentSha256=sha(composition_parent.read_bytes()),UnchangedFirstRowSha256=source_row_sha,
                ExpectedText=rows[-1]['ExpectedText']+['Left terminal control.','Right terminal control.'],
                Change='Append one ordinary row while preserving the complete accepted spanning row byte-for-byte as XML; no candidate-derived geometry.')
        if variant == 'short-left':
            del rows[-1]['SourceRowsPerCell']
            rows[-1]['SourceRowsByCell'] = dict(Left=20, Right=120)
        if variant in ('boxed-math', 'boxed-insets'):
            rows[-1].update(ExpectedEquationCount=1,ExpectedRenderedEquationCount=1,
                BoxSourceInput=box_parent['Input'],BoxSourceSha256=box_parent['Sha256'],
                SourceBoxParagraphSha256=sha(E.tostring(box_paragraph)),FieldInstruction='NUMPAGES',
                ExpectedText=rows[-1]['ExpectedText']+['Before inline','after equation.'],
                Change='Preserve an accepted boxed equation/NUMPAGES prefix in the left cell beside ordinary parallel prose; use source-authored widths.')
        if variant in ('bordered-aligned', 'bordered-thick'):
            rows[-1].update(BorderWidthPoints=source_borders[variant][0] / 8,
                RightIntroductionAfterPoints=16 if source_borders[variant][1] else 8,
                Change='Independently vary border thickness or introduction spacing; preserve explicit zero table indentation and source paragraph content.')
        if variant in source_insets:
            rows[-1].update(CellInsets={side:dict(Top=points[0],Bottom=points[1]) for side,points in source_insets[variant].items()},
                CellFills=dict(Left='DDEBF7',Right='E2F0D9'),
                Change='Explicit cell-owned vertical insets with ordinary prose; pale source fills expose shared physical fragments without border rules.')
        jobs.append((cases / (name + '.docx'), data))
    jobs.append((cases / manifest_name,
        (json.dumps(rows, indent=2) + '\n').replace('\n', '\r\n').encode()))
    publish(jobs)
    print(str(len(rows)) + ' ordinary parallel-cell controls verified; earlier inputs unchanged.')
    raise SystemExit(0)


if args.spanning_frame_probes:
    parent = json.loads((cases / 'spanning-cell-total-count.json').read_text())[0]
    original = (root / parent['Input']).read_bytes()
    assert sha(original) == parent['Sha256']
    with ZipFile(BytesIO(original)) as archive:
        infos = archive.infolist()
        original_parts = {entry.filename: archive.read(entry.filename) for entry in infos}
    wp = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
    a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    rows, jobs = [], []
    for variant, height in [('short-box', 42), ('tall-box', 78), ('mixed-prose', None)]:
        parts = dict(original_parts)
        tree = E.fromstring(parts['word/document.xml'])
        cell = next(tree.iter(w + 'tc'))
        if height is not None:
            for tag in (wp + 'extent', a + 'ext'):
                node = next(cell.iter(tag))
                node.set('cy', str(height * 12700))
            count = 84
        else:
            prefix = cell.find(w + 'p')
            prefix.clear()
            E.SubElement(E.SubElement(prefix, w + 'r'), w + 't').text = 'Ordinary cell introduction.'
            sample = cell.findall(w + 'p')[1]
            for paragraph in cell.findall(w + 'p')[1:]:
                cell.remove(paragraph)
            count = 120
            for index in range(1, count + 1):
                paragraph = copy.deepcopy(sample)
                paragraph.find('.//' + w + 't').text = 'Planning row ' + str(index) + ' stays visible.'
                paragraph.find('.//' + w + 'sz').set(w + 'val', '24' if index % 2 else '20')
                cell.append(paragraph)
        parts['word/document.xml'] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
        buffer = BytesIO()
        with ZipFile(buffer, 'w') as archive:
            for entry in infos:
                archive.writestr(copy.copy(entry), parts[entry.filename])
        data = buffer.getvalue()
        name = 'spanning-frame-' + variant
        row = dict(Id=name,Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + name + '.docx',
            Sha256=sha(data),SourceInput=parent['Input'],SourceSha256=parent['Sha256'],HeldOut=True,
            ExpectedEquationCount=0 if height is None else 1,ExpectedRenderedEquationCount=0 if height is None else 1,
            PlanningParagraphs=count,BoxHeightPoints=height,
            ExpectedText=['Before public math table.','Planning row 1 stays visible.',
                'Planning row ' + str(count) + ' stays visible.','After public math table.'],
            Change='Vary the measured prefix frame or ordinary paragraph metrics; no source-count eligibility rule.')
        rows.append(row)
        jobs.append((cases / (name + '.docx'), data))
    jobs.append((cases / 'spanning-frame-probes.json',
        (json.dumps(rows, indent=2) + '\n').replace('\n', '\r\n').encode()))
    publish(jobs)
    print('Three spanning-frame controls verified; accepted total-count input unchanged.')
    raise SystemExit(0)


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
