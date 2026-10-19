"""Regenerate independently reviewed table-grid origin controls byte-identically."""
import copy
import hashlib
import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parent.parent
manifest = root / 'tests/Lokad.OoxPdf.Tests/Cases/TableGridOrigins/table-grid-origins.json'
source = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/unmerged-math-cell-control.docx'
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
for case in json.loads(manifest.read_text()):
    assert sha(source_bytes) == case['SourceSha256'], case['Id']
    parts = original.copy()
    tree = E.fromstring(parts['word/document.xml'])
    table = tree.find('.//' + w + 'tbl')
    properties = table.find(w + 'tblPr')
    margins = properties.find(w + 'tblCellMar')
    variant = case['Variant']
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
    if variant == 'indent-zero':
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
    if variant in ['bordered', 'top-border', 'right-border', 'inside-border']:
        borders = E.SubElement(properties, w + 'tblBorders')
        sides = (['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] if variant == 'bordered'
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
    assert sha(data) == case['Sha256'], case['Id']
    prepared.append((root / case['Input'], data))

# Validate every expected source/output before replacing any tracked fixture.
for target, data in prepared:
    target.write_bytes(data)
print('Fourteen table-grid origin controls regenerated byte-identically.')
