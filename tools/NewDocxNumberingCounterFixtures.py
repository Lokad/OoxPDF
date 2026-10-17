"""Independent counter-order, digit-width and suffix controls for inserted marks."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
source = cases / 'list-decoration-mark-inserted.docx'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = []
for name in ['second-mark', 'two-marks', 'start-nine', 'space-suffix', 'nothing-suffix', 'custom-tab', 'independent-list', 'custom-tab-wide', 'restart-list']:
    with ZipFile(source) as z:
        entries = {part: z.read(part) for part in z.namelist()}
    footer = E.fromstring(entries['word/footer1.xml'])
    paragraphs = footer.findall(w + 'p')
    first = paragraphs[0].find(w + 'pPr/' + w + 'rPr')
    insertion = first.find(w + 'ins')
    if name in ['second-mark', 'two-marks']:
        mark = E.SubElement(paragraphs[1].find(w + 'pPr'), w + 'rPr')
        E.SubElement(mark, w + 'ins', dict(insertion.attrib) | {w + 'id': '311'})
        if name == 'second-mark': first.remove(insertion)
    numbering = E.fromstring(entries['word/numbering.xml'])
    level = numbering.find(w + 'abstractNum/' + w + 'lvl')
    if name in ['start-nine', 'custom-tab-wide', 'restart-list']: level.find(w + 'start').set(w + 'val', '9')
    elif name.endswith('-suffix'):
        suffix = level.find(w + 'suff')
        if suffix is None: suffix = E.SubElement(level, w + 'suff')
        suffix.set(w + 'val', name.split('-')[0])
    if name.startswith('custom-tab'):
        settings = E.fromstring(entries['word/settings.xml'])
        E.SubElement(settings, w + 'defaultTabStop', {w + 'val': '480'})
        entries['word/settings.xml'] = E.tostring(settings, encoding='utf-8', xml_declaration=True)
    if name in ['independent-list', 'restart-list']:
        number = E.SubElement(numbering, w + 'num', {w + 'numId': '2'})
        E.SubElement(number, w + 'abstractNumId', {w + 'val': '0'})
        if name == 'restart-list':
            override = E.SubElement(number, w + 'lvlOverride', {w + 'ilvl': '0'})
            E.SubElement(override, w + 'startOverride', {w + 'val': '1'})
        for paragraph in paragraphs[1:]:
            paragraph.find(w + 'pPr/' + w + 'numPr/' + w + 'numId').set(w + 'val', '2')
    entries['word/footer1.xml'] = E.tostring(footer, encoding='utf-8', xml_declaration=True)
    entries['word/numbering.xml'] = E.tostring(numbering, encoding='utf-8', xml_declaration=True)
    target = cases / ('numbering-counter-' + name + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as z:
        for part, data in entries.items():
            info = ZipInfo(part, (2026, 10, 17, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=1, HeldOut=False, ExpectedText=['share the same input series.']))
(cases / 'numbering-counters.json').write_text(json.dumps(rows, indent=2)+'\n', encoding='utf-8')
print('Pinned', len(rows), 'counter-order/width/suffix controls')
