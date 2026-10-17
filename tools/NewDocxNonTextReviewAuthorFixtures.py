"""Isolate actual nontext revision owners beside an ordinary source comment."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
source = cases / 'ReviewAuthors/review-authors-no-revision.docx'
source_sha = hashlib.sha256(source.read_bytes()).hexdigest().upper()
assert source_sha == '201EEFC0969A39E392FE50CB94FA6D7402DB65A212A3EF102A32345B6A29E569'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
E.register_namespace('w', w[1:-1])
rows = []
for kind in ['none', 'numbered-mark', 'paragraph-mark', 'run-format', 'image-insertion', 'empty-insertion', 'text-insertion', 'same-author-mark', 'mark-with-text', 'image-before-text', 'image-after-text', 'format-before-text', 'format-after-text']:
    variant = kind
    kind = {'mark-with-text': 'numbered-mark', 'image-before-text': 'image-insertion', 'image-after-text': 'image-insertion', 'format-before-text': 'run-format', 'format-after-text': 'run-format'}.get(variant, variant)
    with ZipFile(source) as archive:
        entries = {part: archive.read(part) for part in archive.namelist()}
    document = E.fromstring(entries['word/document.xml'])
    body = document.find(w + 'body')
    section = deepcopy(body.find(w + 'sectPr'))
    for child in list(section):
        if child.tag in [w + 'headerReference', w + 'footerReference']:
            section.remove(child)
    label = deepcopy(next(body.iter(w + 'numPr')))
    drawing_run = deepcopy(next(run for run in body.iter(w + 'r') if run.find(w + 'drawing') is not None))
    body.clear()
    owner = E.SubElement(body, w + 'p')
    properties = E.SubElement(owner, w + 'pPr')
    revision = {w + 'id': '400', w + 'author': 'Nontext Reviewer', w + 'date': '2026-10-17T00:00:00Z'}
    if kind in ['numbered-mark', 'same-author-mark']:
        properties.append(label)
    if kind in ['numbered-mark', 'paragraph-mark', 'same-author-mark']:
        E.SubElement(E.SubElement(properties, w + 'rPr'), w + 'ins', revision)
    run = E.SubElement(owner, w + 'r')
    if kind == 'run-format':
        rpr = E.SubElement(run, w + 'rPr')
        E.SubElement(rpr, w + 'b')
        E.SubElement(E.SubElement(rpr, w + 'rPrChange', revision), w + 'rPr')
    E.SubElement(run, w + 't').text = 'Public owner paragraph.'
    if kind in ['image-insertion', 'empty-insertion', 'text-insertion']:
        insertion = E.SubElement(owner, w + 'ins', revision)
        if kind == 'image-insertion':
            insertion.append(drawing_run)
        else:
            E.SubElement(E.SubElement(insertion, w + 'r'), w + 't').text = 'Inserted phrase.' if kind == 'text-insertion' else ''
    if variant in ['mark-with-text', 'image-before-text', 'image-after-text', 'format-before-text', 'format-after-text']:
        other_revision = dict(revision); other_revision[w + 'author'] = 'Other Reviewer'; other_revision[w + 'id'] = '401'
        other = E.SubElement(owner, w + 'ins', other_revision)
        E.SubElement(E.SubElement(other, w + 'r'), w + 't').text = 'Other inserted phrase.'
        if variant == 'image-after-text':
            owner.remove(insertion); owner.append(insertion)
        if variant == 'format-after-text':
            owner.remove(run); owner.append(run)
    paragraph = E.SubElement(body, w + 'p')
    E.SubElement(paragraph, w + 'commentRangeStart', {w + 'id': '0'})
    E.SubElement(E.SubElement(paragraph, w + 'r'), w + 't').text = 'Public comment anchor.'
    E.SubElement(paragraph, w + 'commentRangeEnd', {w + 'id': '0'})
    E.SubElement(E.SubElement(paragraph, w + 'r'), w + 'commentReference', {w + 'id': '0'})
    body.append(section)
    comments = E.Element(w + 'comments')
    comment = E.SubElement(comments, w + 'comment', {w + 'id': '0', w + 'author': 'Nontext Reviewer' if variant in ['same-author-mark', 'mark-with-text', 'image-before-text', 'image-after-text', 'format-before-text', 'format-after-text'] else 'Comment Reviewer', w + 'initials': 'CR'})
    E.SubElement(E.SubElement(E.SubElement(comment, w + 'p'), w + 'r'), w + 't').text = 'Check source owner.'
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    entries['word/comments.xml'] = E.tostring(comments, encoding='utf-8', xml_declaration=True)
    target = cases / ('ReviewAuthors/nontext-author-' + variant + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
        for part, data in sorted(entries.items()):
            info = ZipInfo(part, (2026, 10, 17, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(), Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), SourceSha256=source_sha, Kind=variant, HeldOut=False, ExpectedEquationCount=0, ExpectedText=['Public owner paragraph.', 'Public comment anchor.']))
(cases / 'ReviewAuthors/nontext-authors.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'nontext revision-author controls')
