"""Public image-revision composition and owning-run formatting controls."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json
root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
output = cases / 'PictureRevisions'; output.mkdir(exist_ok=True)
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
E.register_namespace('w', w[1:-1]); E.register_namespace('r', r[1:-1])
previous = json.loads((cases / 'ReviewAuthors/nontext-authors.json').read_text(encoding='utf-8'))
rows = [row for row in previous if row['Kind'] in ['image-insertion', 'image-before-text', 'image-after-text']]
rows += [row for row in json.loads((cases / 'ReviewAuthors/revision-text-ink.json').read_text(encoding='utf-8')) if row['Id'] == 'revision-ink-deletion-after-image']
source = cases / 'ReviewAuthors/nontext-author-image-insertion.docx'
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == next(row['Sha256'] for row in rows if row['Id'] == 'nontext-author-image-insertion')
for variant in ['cell', 'header', 'footer', 'arial14', 'calibri18', 'authored-double-green', 'shared-run-body', 'shared-run-cell', 'shared-run-header', 'shared-run-footer']:
    with ZipFile(source) as archive:
        entries = {part: archive.read(part) for part in archive.namelist()}
    document = E.fromstring(entries['word/document.xml']); body = document.find(w + 'body')
    paragraph = body.find(w + 'p'); section = body.find(w + 'sectPr')
    insertion = next(paragraph.iter(w + 'ins')); picture_run = insertion.find(w + 'r')
    placement = variant.replace('shared-run-', '') if variant.startswith('shared-run-') else variant
    if variant.startswith('shared-run-'):
        drawing = deepcopy(picture_run.find(w + 'drawing'))
        wp = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
        drawing.find('.//' + wp + 'docPr').set('id', '2')
        picture_run.append(drawing)
        E.SubElement(picture_run, w + 't', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = ' Attached source.'
    if variant in ['arial14', 'calibri18', 'authored-double-green']:
        props = E.SubElement(picture_run, w + 'rPr'); picture_run.remove(props); picture_run.insert(0, props)
        if variant in ['arial14', 'calibri18']:
            family, size = ('Arial', '28') if variant == 'arial14' else ('Calibri', '36')
            E.SubElement(props, w + 'rFonts', {w + 'ascii': family, w + 'hAnsi': family})
            E.SubElement(props, w + 'sz', {w + 'val': size})
        else:
            E.SubElement(props, w + 'u', {w + 'val': 'double', w + 'color': '00A000'})
    if placement == 'cell':
        body.remove(paragraph)
        table = E.Element(w + 'tbl'); props = E.SubElement(table, w + 'tblPr')
        E.SubElement(props, w + 'tblW', {w + 'w': '9360', w + 'type': 'dxa'})
        E.SubElement(props, w + 'tblLayout', {w + 'type': 'fixed'})
        E.SubElement(E.SubElement(table, w + 'tblGrid'), w + 'gridCol', {w + 'w': '9360'})
        cell = E.SubElement(E.SubElement(table, w + 'tr'), w + 'tc')
        E.SubElement(E.SubElement(cell, w + 'tcPr'), w + 'tcW', {w + 'w': '9360', w + 'type': 'dxa'})
        cell.append(paragraph); body.insert(0, table)
    if placement in ['header', 'footer']:
        body.remove(paragraph)
        story = E.Element(w + ('hdr' if placement == 'header' else 'ftr')); story.append(paragraph)
        entries['word/' + placement + '1.xml'] = E.tostring(story, encoding='utf-8', xml_declaration=True)
        entries['word/_rels/' + placement + '1.xml.rels'] = entries['word/_rels/document.xml.rels']
        E.SubElement(section, w + placement + 'Reference', {w + 'type': 'default', r + 'id': placement})
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    target = output / ('picture-revision-' + variant + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
        for part, data in sorted(entries.items()):
            info = ZipInfo(part, (2026, 10, 17, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(), Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), SourceSha256=source_sha, Variant=variant, HeldOut=False, ExpectedEquationCount=0, ExpectedText=['Public owner paragraph.', 'Public comment anchor.']))
(output / 'picture-revisions.json').write_text(json.dumps(rows, indent=2)+'\n', encoding='utf-8')
print('Pinned fourteen image-revision controls; four prior sources reused')
