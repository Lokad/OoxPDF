"""Controlled numbering-instance identity, continuation and restart inputs."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
from copy import deepcopy
import json
import hashlib

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
source = cases / 'list-decoration-mark-inserted.docx'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = []
names = ['shared-return', 'override-return', 'override-continue', 'distinct-definitions',
    'distinct-nsids', 'shared-nsid', 'level-format-override',
    'shared-return-unmarked', 'override-return-unmarked', 'zero-nsid', 'case-nsid', 'shared-nsid-variation', 'shared-nsid-reverse-definitions', 'shared-nsid-reverse-references']
for name in names:
    with ZipFile(source) as z:
        entries = {part: z.read(part) for part in z.namelist()}
    footer = E.fromstring(entries['word/footer1.xml'])
    paragraphs = footer.findall(w + 'p')
    numbering = E.fromstring(entries['word/numbering.xml'])
    abstract = numbering.find(w + 'abstractNum')
    first_number = numbering.find(w + 'num')
    second_number = E.SubElement(numbering, w + 'num', {w + 'numId': '2'})
    E.SubElement(second_number, w + 'abstractNumId', {w + 'val': '0'})
    for paragraph in paragraphs[1:]:
        paragraph.find(w + 'pPr/' + w + 'numPr/' + w + 'numId').set(w + 'val', '2')
    if 'return' in name:
        paragraphs[2].find(w + 'pPr/' + w + 'numPr/' + w + 'numId').set(w + 'val', '1')
    if name.startswith('override-'):
        owner = first_number if name == 'override-continue' else second_number
        level = E.SubElement(owner, w + 'lvlOverride', {w + 'ilvl': '0'})
        E.SubElement(level, w + 'startOverride', {w + 'val': '9'})
    if name in ['distinct-definitions', 'distinct-nsids', 'shared-nsid', 'zero-nsid', 'case-nsid', 'shared-nsid-variation', 'shared-nsid-reverse-definitions', 'shared-nsid-reverse-references']:
        other = deepcopy(abstract)
        other.set(w + 'abstractNumId', '1')
        numbering.insert(1, other)
        second_number.find(w + 'abstractNumId').set(w + 'val', '1')
        if name != 'distinct-definitions':
            first_nsid = '00000000' if name == 'zero-nsid' else 'AA11BB22' if name == 'case-nsid' else '11111111'
            second_nsid = '22222222' if name == 'distinct-nsids' else 'aa11bb22' if name == 'case-nsid' else first_nsid
            abstract.insert(0, E.Element(w + 'nsid', {w + 'val': first_nsid}))
            other.insert(0, E.Element(w + 'nsid', {w + 'val': second_nsid}))
        if name in ['shared-nsid-variation', 'shared-nsid-reverse-definitions', 'shared-nsid-reverse-references']:
            other.find(w + 'lvl/' + w + 'lvlText').set(w + 'val', '%1)')
        if name == 'shared-nsid-reverse-definitions':
            numbering.remove(other)
            numbering.insert(0, other)
        if name == 'shared-nsid-reverse-references':
            for index, paragraph in enumerate(paragraphs):
                paragraph.find(w + 'pPr/' + w + 'numPr/' + w + 'numId').set(w + 'val', '2' if index == 0 else '1')
    if name == 'level-format-override':
        override = E.SubElement(second_number, w + 'lvlOverride', {w + 'ilvl': '0'})
        level = deepcopy(abstract.find(w + 'lvl'))
        level.find(w + 'lvlText').set(w + 'val', '%1)')
        override.append(level)
    if name.endswith('-unmarked'):
        mark = paragraphs[0].find(w + 'pPr/' + w + 'rPr')
        mark.remove(mark.find(w + 'ins'))
    entries['word/footer1.xml'] = E.tostring(footer, encoding='utf-8', xml_declaration=True)
    entries['word/numbering.xml'] = E.tostring(numbering, encoding='utf-8', xml_declaration=True)
    target = cases / ('numbering-continuation-' + name + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as z:
        for part, data in entries.items():
            info = ZipInfo(part, (2026, 10, 17, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=1, HeldOut=False, ExpectedText=['share the same input series.']))
(cases / 'numbering-continuations.json').write_text(json.dumps(rows, indent=2)+'\n', encoding='utf-8')
print('Pinned', len(rows), 'continuation/restart identity controls')
