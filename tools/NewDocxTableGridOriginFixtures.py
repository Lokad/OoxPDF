"""Regenerate independently reviewed table-grid origin controls byte-identically."""
import argparse
import copy
import hashlib
import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parent.parent
parser = argparse.ArgumentParser(description=__doc__)
mode = parser.add_mutually_exclusive_group()
mode.add_argument('--authored-zero-probes', action='store_true')
mode.add_argument('--implicit-border-probes', action='store_true')
parser.add_argument('--write', action='store_true')
args = parser.parse_args()
manifest = root / ('tests/Lokad.OoxPdf.Tests/Cases/TableGridOrigins/' +
    ('table-implicit-border-origins.json' if args.implicit_border_probes else
     'table-zero-indents.json' if args.authored_zero_probes else 'table-grid-origins.json'))
source = root / ('tests/Lokad.OoxPdf.Tests/Cases/TableGridOrigins/table-origin-bordered.docx' if args.implicit_border_probes
    else 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/unmerged-math-cell-control.docx')
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()
source_bytes = source.read_bytes()
with ZipFile(BytesIO(source_bytes)) as archive:
    infos = archive.infolist()
    original = {info.filename: archive.read(info.filename) for info in infos}


def set_margins(owner, value):
    for side in ['left', 'right']:
        entry = owner.find(w + side)
        if entry is None:
            entry = E.SubElement(owner, w + side)
        entry.attrib.update({w + 'w': str(value), w + 'type': 'dxa'})


prepared = []
if args.implicit_border_probes:
    cases = []
    for variant in ['margin0', 'margin200', 'cell200', 'width8', 'width24']:
        identifier = 'table-start-border-' + variant
        cases.append(dict(Id=identifier, Input='tests/Lokad.OoxPdf.Tests/Cases/TableGridOrigins/' + identifier + '.docx',
            Source=str(source.relative_to(root)).replace('\\', '/'), SourceSha256=sha(source_bytes), HeldOut=True,
            ExpectedEquationCount=1, ExpectedRenderedEquationCount=1,
            ExpectedText=['Before public math table.', 'Before inline', 'after equation.', 'After public math table.'],
            Variant='implicit-' + variant,
            Change='Authored table/cell margins with absent indentation and solid borders. Thick-border Word offset remains unresolved.'))
elif args.authored_zero_probes:
    cases = []
    for variant, margin in [('zero-bordered', 108), ('zero-bordered-wide', 200)]:
        identifier = 'table-origin-' + variant
        cases.append(dict(Id=identifier, Input='tests/Lokad.OoxPdf.Tests/Cases/TableGridOrigins/' + identifier + '.docx',
            Source=str(source.relative_to(root)).replace('\\', '/'), SourceSha256=sha(source_bytes), HeldOut=True,
            ExpectedEquationCount=1, ExpectedRenderedEquationCount=1,
            ExpectedText=['Before public math table.', 'Before inline', 'after equation.', 'After public math table.'],
            Variant=variant, IndentTwips=0, HorizontalMarginPoints=margin / 20,
            Change='Explicit zero indentation with a complete mathematical cell and ordinary/wide margins; preserve implicit-border controls.'))
else:
    cases = json.loads(manifest.read_text())
for case in cases:
    assert sha(source_bytes) == case['SourceSha256'], case['Id']
    parts = original.copy()
    tree = E.fromstring(parts['word/document.xml'])
    table = tree.find('.//' + w + 'tbl')
    properties = table.find(w + 'tblPr')
    margins = properties.find(w + 'tblCellMar')
    variant = case['Variant']
    if variant.startswith('implicit-width'):
        for edge in properties.find(w + 'tblBorders'):
            edge.set(w + 'sz', variant.removeprefix('implicit-width'))
    if variant.startswith('implicit-margin'):
        set_margins(margins, int(variant.removeprefix('implicit-margin')))
    if variant == 'implicit-cell200':
        set_margins(E.SubElement(table.find(w + 'tr').find(w + 'tc').find(w + 'tcPr'), w + 'tcMar'), 200)
    if variant in ['direct-zero', 'direct-wide', 'style-108-direct-zero']:
        set_margins(margins, 0 if 'zero' in variant else 200)
    if variant == 'cell-wide':
        cell = table.find(w + 'tr').find(w + 'tc').find(w + 'tcPr')
        set_margins(E.SubElement(cell, w + 'tcMar'), 200)
    if variant.startswith('style-'):
        styles = E.fromstring(parts['word/styles.xml'])
        assert not any(style.get(w + 'type') == 'table' for style in styles.findall(w + 'style'))
        style = E.SubElement(styles, w + 'style', {w + 'type': 'table', w + 'default': '1', w + 'styleId': 'TableNormal'})
        E.SubElement(style, w + 'name', {w + 'val': 'Normal Table'})
        style_margins = E.SubElement(E.SubElement(style, w + 'tblPr'), w + 'tblCellMar')
        set_margins(style_margins, 0 if variant == 'style-zero' else 200 if variant == 'style-200-direct-108' else 108)
        if variant == 'style-108':
            properties.remove(margins)
        parts['word/styles.xml'] = E.tostring(styles, encoding='utf-8', xml_declaration=True)
    if variant == 'zero-bordered-wide':
        set_margins(margins, 200)
    if variant in ['indent-zero', 'zero-bordered', 'zero-bordered-wide']:
        E.SubElement(properties, w + 'tblInd', {w + 'w': '0', w + 'type': 'dxa'})
    if variant == 'compat-15':
        settings = E.Element(w + 'settings')
        compat = E.SubElement(settings, w + 'compat')
        E.SubElement(compat, w + 'compatSetting', {w + 'name': 'compatibilityMode',
            w + 'uri': 'http://schemas.microsoft.com/office/word', w + 'val': '15'})
        parts['word/settings.xml'] = E.tostring(settings, encoding='utf-8', xml_declaration=True)
        relationships = E.fromstring(parts['word/_rels/document.xml.rels'])
        E.SubElement(relationships, rel + 'Relationship', {'Id': 'rOriginSettings',
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'Target': 'settings.xml'})
        parts['word/_rels/document.xml.rels'] = E.tostring(relationships, encoding='utf-8', xml_declaration=True)
        types = E.fromstring(parts['[Content_Types].xml'])
        E.SubElement(types, ct + 'Override', {'PartName': '/word/settings.xml',
            'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'})
        parts['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
    if variant == 'auto-width':
        properties.find(w + 'tblW').attrib.update({w + 'w': '0', w + 'type': 'auto'})
        properties.remove(properties.find(w + 'tblLayout'))
    if variant in ['bordered', 'top-border', 'right-border', 'inside-border', 'zero-bordered', 'zero-bordered-wide']:
        borders = E.SubElement(properties, w + 'tblBorders')
        sides = (['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] if variant in ['bordered', 'zero-bordered', 'zero-bordered-wide']
            else [{'top-border': 'top', 'right-border': 'right', 'inside-border': 'insideV'}[variant]])
        for side in sides:
            E.SubElement(borders, w + side, {w + 'val': 'single', w + 'sz': '4', w + 'color': '000000'})
    parts['word/document.xml'] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
    buffer = BytesIO()
    with ZipFile(buffer, 'w', ZIP_DEFLATED) as archive:
        for info in infos:
            archive.writestr(copy.copy(info), parts[info.filename])
        for name in sorted(set(parts) - {info.filename for info in infos}):
            info = ZipInfo(name, (2026, 10, 19, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, parts[name])
    data = buffer.getvalue()
    if args.authored_zero_probes or args.implicit_border_probes:
        case['Sha256'] = sha(data)
    else:
        assert sha(data) == case['Sha256'], case['Id']
    prepared.append((root / case['Input'], data))

if args.authored_zero_probes or args.implicit_border_probes:
    prepared.append((manifest, (json.dumps(cases, indent=2) + '\n').replace('\n', '\r\n').encode()))
# Default verification never writes through frozen input links. Explicit publication
# replaces directory entries after all source/output bytes have been constructed.
for target, data in prepared:
    if args.write:
        stage = target.with_suffix(target.suffix + '.stage')
        with stage.open('xb') as output:
            output.write(data)
        stage.replace(target)
    else:
        assert target.read_bytes() == data, target.name
print(str(len(cases)) + ' table-grid origin controls verified; atomic publication requires --write.')
