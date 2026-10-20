"""Verify four source-preserving inline metadata wrappers; --write publishes atomically."""
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
sha = lambda data: hashlib.sha256(data).hexdigest().upper()
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--write', action='store_true')
args = parser.parse_args()
parents = []
for manifest, identifier in [('inline-sibling-frames.json', 'inline-sibling-mixed'),
                             ('inline-box-ordinary-pictures.json', 'inline-box-picture-cell')]:
    parents.append(next(row for row in json.loads((cases / manifest).read_text()) if row['Id'] == identifier))
jobs, rows = [], []
for parent in parents:
    source = (root / parent['Input']).read_bytes()
    assert sha(source) == parent['Sha256']
    with ZipFile(BytesIO(source)) as archive:
        infos = archive.infolist()
        parts = {entry.filename: archive.read(entry.filename) for entry in infos}
    part = parent['MathPart']
    for kind in ['customXml', 'smartTag']:
        tree = E.fromstring(parts[part])
        original_xml = E.tostring(tree)
        paragraph = next(node for node in tree.iter(w + 'p') if node.find('.//' + w + 'txbxContent') is not None)
        content = [node for node in paragraph if node.tag != w + 'pPr']
        assert content and any(node.find('.//' + w + 'txbxContent') is not None for node in content)
        wrapper = E.Element(w + kind, {w + 'uri': 'urn:lokad:public-composition', w + 'element': 'calculation'})
        for node in content:
            paragraph.remove(node)
            wrapper.append(node)
        paragraph.append(wrapper)
        changed = E.tostring(tree, encoding='utf-8', xml_declaration=True)
        paragraph.remove(wrapper)
        paragraph.extend(list(wrapper))
        assert E.tostring(tree) == original_xml
        output = BytesIO()
        with ZipFile(output, 'w') as archive:
            for entry in infos:
                archive.writestr(copy.copy(entry), changed if entry.filename == part else parts[entry.filename])
        data = output.getvalue()
        variant = 'body-mixed' if parent['Id'] == 'inline-sibling-mixed' else 'cell-ordinary'
        identifier = 'inline-wrapper-' + kind.lower() + '-' + variant
        row = copy.deepcopy(parent)
        row.update(Id=identifier, Input='tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath/' + identifier + '.docx',
                   Sha256=sha(data), SourceInput=parent['Input'], SourceSha256=parent['Sha256'], HeldOut=True,
                   Variant=variant, Wrapper=kind,
                   Change='Wrap only the existing paragraph inline children in a metadata container. All source owners, formatting, properties and other package parts stay unchanged.')
        rows.append(row)
        jobs.append((cases / (identifier + '.docx'), data))
jobs.append((cases / 'inline-metadata-wrappers.json', (json.dumps(rows, indent=2) + '\n').replace('\n', '\r\n').encode()))
for target, data in jobs:
    if args.write:
        staging = target.with_suffix(target.suffix + '.staging')
        with staging.open('xb') as stream:
            stream.write(data)
        staging.replace(target)
    else:
        assert target.read_bytes() == data, target.name
print('Four inline metadata wrappers verified; removing only the wrapper recovers the source XML.')
