"""Verify bounded picture/box wrapping controls; --write publishes atomically."""
import argparse
import copy
import hashlib
import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
wp = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
pic = '{http://schemas.openxmlformats.org/drawingml/2006/picture}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--write', action='store_true')
args = parser.parse_args()
parent = next(row for row in json.loads((cases / 'static-inline-footer-tables.json').read_text())
              if row['Id'] == 'static-inline-footer-table-picture')
source = (root / parent['Input']).read_bytes()
assert sha(source) == parent['Sha256']
with ZipFile(BytesIO(source)) as archive:
    infos = archive.infolist()
    parts = {entry.filename: archive.read(entry.filename) for entry in infos}
jobs, rows = [], []
for variant in ['picture-height', 'grid-wide']:
    tree = E.fromstring(parts[parent['MathPart']])
    original_xml = E.tostring(tree)
    changes = []
    if variant == 'picture-height':
        drawing = next(node for node in tree.iter(wp + 'inline') if node.find('.//' + pic + 'pic') is not None)
        nodes = [drawing.find(wp + 'extent'), drawing.find('.//' + a + 'xfrm/' + a + 'ext')]
        for node in nodes:
            assert node is not None and node.get('cy') == str(18 * 12700)
            changes.append((node, 'cy', node.get('cy'), str(30 * 12700)))
    else:
        table = tree.find(w + 'tbl')
        assert table is not None
        nodes = [table.find(w + 'tblPr/' + w + 'tblW'), table.find(w + 'tblGrid/' + w + 'gridCol'),
                 table.find(w + 'tr/' + w + 'tc/' + w + 'tcPr/' + w + 'tcW')]
        for node in nodes:
            assert node is not None and node.get(w + 'w') == '6000'
            changes.append((node, w + 'w', node.get(w + 'w'), '8800'))
    for node, attribute, _, value in changes:
        node.set(attribute, value)
    changed_xml = E.tostring(tree, encoding='utf-8', xml_declaration=True)
    for node, attribute, old, _ in changes:
        node.set(attribute, old)
    assert E.tostring(tree) == original_xml
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as archive:
        for entry in infos:
            archive.writestr(copy.copy(entry), changed_xml if entry.filename == parent['MathPart'] else parts[entry.filename])
    data = buffer.getvalue()
    identifier = 'footer-cell-sibling-' + variant
    row = copy.deepcopy(parent)
    row.update(Id=identifier, Input='tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath/' + identifier + '.docx',
               Sha256=sha(data), SourceInput=parent['Input'], SourceSha256=parent['Sha256'], Variant=variant,
               Change='Change only the picture height18-to30pt in its paired extents.' if variant == 'picture-height' else
               'Change only the table width300-to440pt in its matched preferred/grid/cell widths; preserve all source items.')
    rows.append(row)
    jobs.append((cases / (identifier + '.docx'), data))
jobs.append((cases / 'footer-cell-sibling-probes.json', (json.dumps(rows, indent=2) + '\n').replace('\n', '\r\n').encode()))
for target, data in jobs:
    if args.write:
        stage = target.with_suffix(target.suffix + '.staging')
        with stage.open('xb') as output:
            output.write(data)
        stage.replace(target)
    else:
        assert target.read_bytes() == data, target.name
print('Two picture/box sibling controls verified; source recovered after reverting declared geometry attributes.')
