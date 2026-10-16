"""Static-story transformations of the public formatting/math controls."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
rows = []
for name in ['footer-distance-18', 'footer-distance-72', 'footer-second-paragraph',
             'footer-plain-first', 'footer-plain-second', 'header-second-paragraph',
             'header-plain-first', 'header-plain-second']:
    story = 'header' if name.startswith('header') else 'footer'
    source = cases / ('math-run-format-' + story + '-bold.docx')
    with ZipFile(source) as z:
        entries = {part: z.read(part) for part in z.namelist()}
    part = 'word/' + story + '1.xml'
    xml = E.fromstring(entries[part])
    if name.endswith('second-paragraph') or name.endswith('plain-second'):
        paragraphs = xml.findall(w + 'p')
        before = paragraphs[0].find(w + 'r/' + w + 'rPr')
        after = paragraphs[1].find(w + 'r')
        after.insert(0, deepcopy(before))
        before.clear()
    if 'plain' in name:
        for parent in xml.iter():
            for child in list(parent):
                if child.tag == m + 'oMathPara':
                    index = list(parent).index(child)
                    run = E.Element(w + 'r')
                    E.SubElement(run, w + 't').text = 'formula'
                    parent.remove(child)
                    parent.insert(index, run)
                elif child.tag == w + 'drawing':
                    parent.remove(child)
    if 'distance' in name:
        document = E.fromstring(entries['word/document.xml'])
        for margins in document.iter(w + 'pgMar'):
            margins.set(w + 'footer', str(int(name.rsplit('-', 1)[1]) * 20))
        entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    else:
        entries[part] = E.tostring(xml, encoding='utf-8', xml_declaration=True)
    target = cases / ('static-frame-' + name + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as z:
        for part, data in entries.items():
            info = ZipInfo(part, (2026, 10, 16, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=name, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=0 if 'plain' in name else 4, HeldOut=False,
        ExpectedText=['share the same input series.']))
(cases / 'static-source-frames.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'static story transformations')
