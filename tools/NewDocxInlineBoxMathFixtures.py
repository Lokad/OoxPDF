"""Verify derived inline-box math fixtures; --write publishes validated bytes.

The Word-accepted boxed-equation fixture is the canonical frame input. Child
paragraphs and image extents remain source-exact when sizing its outer frame.
"""
import argparse
import copy
import hashlib
import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parent.parent
directory = root / 'tests/Lokad.OoxPdf.Tests/Cases/InlineBoxMath'
cases = json.loads((directory / 'inline-box-math.json').read_text())
by_variant = {case['Variant']: case for case in cases}
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
wp = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
wps = '{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}'
rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
sha = lambda data: hashlib.sha256(data).hexdigest().upper()


def unpack(data):
    with ZipFile(BytesIO(data)) as archive:
        infos = archive.infolist()
        return infos, {info.filename: archive.read(info.filename) for info in infos}


def pack(infos, parts):
    buffer = BytesIO()
    with ZipFile(buffer, 'w', ZIP_DEFLATED) as archive:
        for info in infos:
            archive.writestr(copy.copy(info), parts[info.filename])
        for name in sorted(set(parts) - {info.filename for info in infos}):
            info = ZipInfo(name, (2026, 10, 19, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, parts[name])
    return buffer.getvalue()


seed_case = by_variant['boxed-equation']
seed = (root / seed_case['Input']).read_bytes()
assert sha(seed) == seed_case['Sha256']
seed_infos, seed_parts = unpack(seed)
seed_frame = E.fromstring(seed_parts['word/document.xml']).find(w + 'body').findall(w + 'p')[1]
prepared = []
for case in cases:
    variant = case['Variant']
    if variant == 'boxed-equation':
        continue  # Canonical input, verified above.
    infos, parts = seed_infos, seed_parts.copy()
    tree = E.fromstring(parts['word/document.xml'])
    body = tree.find(w + 'body')
    outer = body.findall(w + 'p')[1]
    content = outer.find('.//' + w + 'txbxContent')
    if variant in ['body-equation', 'body-equation-inline-box']:
        paragraph = copy.deepcopy(content.find(w + 'p'))
        if variant == 'body-equation-inline-box':
            drawing_run = copy.deepcopy(outer.find(w + 'r'))
            box_content = drawing_run.find('.//' + w + 'txbxContent')
            box_content.clear()
            E.SubElement(E.SubElement(E.SubElement(box_content, w + 'p'), w + 'r'), w + 't').text = 'Inline box content.'
            paragraph.append(drawing_run)
        body.remove(outer)
        body.insert(1, paragraph)
    elif variant == 'boxed-text':
        paragraph = content.find(w + 'p')
        for node in list(paragraph):
            if node.tag in [m + 'oMath', m + 'oMathPara']:
                paragraph.remove(node)
    elif variant == 'cell':
        body.remove(outer)
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
        cell.append(outer)
        body.insert(1, table)
    elif variant in ['header', 'footer']:
        body.remove(outer)
        name = 'word/inline-box-' + variant + '.xml'
        story = E.Element(w + ('hdr' if variant == 'header' else 'ftr'))
        story.append(outer)
        parts[name] = E.tostring(story, encoding='utf-8', xml_declaration=True)
        section = body.find(w + 'sectPr')
        section.find(w + 'pgMar').set(w + variant, '180')
        rid = 'rInlineBox' + variant.title()
        E.SubElement(section, w + variant + 'Reference', {w + 'type': 'default',
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': rid})
        links = E.fromstring(parts['word/_rels/document.xml.rels'])
        E.SubElement(links, rel + 'Relationship', {'Id': rid,
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + variant,
            'Target': 'inline-box-' + variant + '.xml'})
        parts['word/_rels/document.xml.rels'] = E.tostring(links, encoding='utf-8', xml_declaration=True)
        types = E.fromstring(parts['[Content_Types].xml'])
        E.SubElement(types, ct + 'Override', {'PartName': '/' + name,
            'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + variant + '+xml'})
        parts['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
    elif variant.startswith('footer-'):
        footer_source = (root / by_variant['footer']['Input']).read_bytes()
        assert sha(footer_source) == case['SourceSha256']
        infos, parts = unpack(footer_source)
        story = E.fromstring(parts['word/inline-box-footer.xml'])
        paragraph = story.find(w + 'p')
        if variant == 'footer-plain':
            child = paragraph.find('.//' + w + 'txbxContent').find(w + 'p')
            for node in list(child):
                if node.tag in [m + 'oMath', m + 'oMathPara']:
                    child.remove(node)
        elif variant == 'footer-pair':
            other = copy.deepcopy(paragraph)
            other.find('.//' + wp + 'docPr').set('id', '91')
            story.append(other)
        elif variant == 'footer-tail':
            E.SubElement(E.SubElement(E.SubElement(story, w + 'p'), w + 'r'), w + 't').text = 'Footer continuation.'
        else:
            raise ValueError(variant)
        parts['word/inline-box-footer.xml'] = E.tostring(story, encoding='utf-8', xml_declaration=True)
    elif variant == 'mixed-source-report':
        source = (root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/math-composition-unnumbered.docx').read_bytes()
        assert sha(source) == case['SourceSha256']
        infos, parts = unpack(source)
        tree = E.fromstring(parts['word/document.xml'])
        body = tree.find(w + 'body')
        frame = copy.deepcopy(seed_frame)
        box_content = frame.find('.//' + w + 'txbxContent')
        box_content.clear()
        for paragraph in body.findall(w + 'p')[1:]:
            body.remove(paragraph)
            box_content.append(paragraph)
        inline = frame.find(w + 'r').find(w + 'drawing').find(wp + 'inline')
        inline.find(wp + 'extent').attrib.update({'cx': '4572000', 'cy': '1524000'})
        # Only this frame's transform changes. Descendant pictures retain theirs.
        own_shape = inline.find(a + 'graphic').find(a + 'graphicData').find(wps + 'wsp')
        own_shape.find(wps + 'spPr').find(a + 'xfrm').find(a + 'ext').attrib.update({'cx': '4572000', 'cy': '1524000'})
        body.insert(1, frame)
    else:
        raise ValueError(variant)
    if not variant.startswith('footer-'):
        parts['word/document.xml'] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
    data = pack(infos, parts)
    assert sha(data) == case['Sha256'], case['Id']
    prepared.append((root / case['Input'], data))

review_cases = json.loads((directory / 'inline-box-review-frames.json').read_text())
for case in review_cases:
    variant = case['Variant']
    source_case = by_variant['boxed-equation' if variant == 'body' else
        'cell' if variant == 'static-cell' else variant]
    source = (root / source_case['Input']).read_bytes()
    assert sha(source) == case['SourceSha256']
    infos, parts = unpack(source)
    tree = E.fromstring(parts['word/document.xml'])
    body = tree.find(w + 'body')
    if variant == 'static-cell':
        table = body.find(w + 'tbl')
        body.remove(table)
        story = E.Element(w + 'hdr')
        story.append(table)
        parts['word/inline-box-review-header.xml'] = E.tostring(story, encoding='utf-8', xml_declaration=True)
        section = body.find(w + 'sectPr')
        section.find(w + 'pgMar').set(w + 'header', '180')
        E.SubElement(section, w + 'headerReference', {w + 'type': 'default',
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id': 'rBoxReviewHeader'})
        links = E.fromstring(parts['word/_rels/document.xml.rels'])
        E.SubElement(links, rel + 'Relationship', {'Id': 'rBoxReviewHeader',
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
            'Target': 'inline-box-review-header.xml'})
        parts['word/_rels/document.xml.rels'] = E.tostring(links, encoding='utf-8', xml_declaration=True)
        types = E.fromstring(parts['[Content_Types].xml'])
        E.SubElement(types, ct + 'Override', {'PartName': '/word/inline-box-review-header.xml',
            'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'})
        parts['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
    paragraph = body.find(w + 'p')
    paragraph.insert(0, E.Element(w + 'commentRangeStart', {w + 'id': '51'}))
    paragraph.append(E.Element(w + 'commentRangeEnd', {w + 'id': '51'}))
    E.SubElement(E.SubElement(paragraph, w + 'r'), w + 'commentReference', {w + 'id': '51'})
    parts['word/document.xml'] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
    comments = E.Element(w + 'comments')
    comment = E.SubElement(comments, w + 'comment', {w + 'id': '51', w + 'author': 'Public Reviewer',
        w + 'initials': 'PR', w + 'date': '2026-10-19T00:00:00Z'})
    E.SubElement(E.SubElement(E.SubElement(comment, w + 'p'), w + 'r'), w + 't').text = 'Review the public calculation frame.'
    parts['word/comments.xml'] = E.tostring(comments, encoding='utf-8', xml_declaration=True)
    links = E.fromstring(parts['word/_rels/document.xml.rels'])
    assert not any(link.get('Type', '').endswith('/comments') for link in links)
    E.SubElement(links, rel + 'Relationship', {'Id': 'rBoxReviewComment',
        'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'Target': 'comments.xml'})
    parts['word/_rels/document.xml.rels'] = E.tostring(links, encoding='utf-8', xml_declaration=True)
    types = E.fromstring(parts['[Content_Types].xml'])
    E.SubElement(types, ct + 'Override', {'PartName': '/word/comments.xml',
        'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'})
    parts['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
    data = pack(infos, parts)
    assert sha(data) == case['Sha256'], case['Id']
    prepared.append((root / case['Input'], data))

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--write', action='store_true')
if parser.parse_args().write:
    for target, data in prepared:
        temporary = target.with_suffix('.docx.staging')
        with temporary.open('xb') as output:
            output.write(data)
        temporary.replace(target)  # Replace the entry; never write through frozen links.
print('Canonical frame verified; ten ordinary and five review fixtures reproduce byte-identically.')
