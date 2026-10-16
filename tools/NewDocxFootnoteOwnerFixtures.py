"""Typography variants of the pinned public four-reference/page-restart control."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import json
import hashlib

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
source = cases / 'footnote-page-owners-calibri-12-zero.docx'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
variants = [('calibri-12-zero', None, None, None), ('calibri-12-met', 'Calibri', 12, '16'),
            ('arial-11-zero', 'Arial', 11, '0'), ('times-14-zero', 'Times New Roman', 14, '0')]
rows = []
for name, family, size, threshold in variants:
    target = cases / ('footnote-page-owners-' + name + '.docx')
    if family is not None:
        with ZipFile(source) as z:
            entries = {name: z.read(name) for name in z.namelist()}
        for part in ['word/document.xml', 'word/footnotes.xml', 'word/styles.xml']:
            xml = E.fromstring(entries[part])
            for fonts in xml.iter(w + 'rFonts'):
                for attr in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
                    if w + attr in fonts.attrib:
                        fonts.set(w + attr, family)
            for properties in xml.iter(w + 'rPr'):
                for element in properties:
                    if element.tag in [w + 'sz', w + 'szCs']:
                        element.set(w + 'val', str(size * 2))
                kern = properties.find(w + 'kern')
                if kern is not None:
                    kern.set(w + 'val', threshold)
            entries[part] = E.tostring(xml, encoding='utf-8', xml_declaration=True)
        with ZipFile(target, 'w', ZIP_DEFLATED) as z:
            for part, data in entries.items():
                info = ZipInfo(part, (2026, 10, 16, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=0, HeldOut=False, ExpectedText=['Public destination Alpha.']))
(cases / 'footnote-page-owners.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned', len(rows), 'complete marker/note-owner controls')

# One-page control whose parsed labels already match the complete measured page.
# Keep it outside the four independent Word typography comparisons above.
with ZipFile(source) as z:
    entries = {name: z.read(name) for name in z.namelist()}
document = E.fromstring(entries['word/document.xml'])
for parent in document.iter():
    for child in list(parent):
        if child.tag == w + 'br' and child.get(w + 'type') == 'page':
            parent.remove(child)
entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
settings = E.fromstring(entries['word/settings.xml'])
settings.find(w + 'footnotePr/' + w + 'numFmt').set(w + 'val', 'decimalZero')
settings.find(w + 'footnotePr/' + w + 'numStart').set(w + 'val', '9')
entries['word/settings.xml'] = E.tostring(settings, encoding='utf-8', xml_declaration=True)
target = cases / 'footnote-page-owners-single-page-unchanged.docx'
with ZipFile(target, 'w', ZIP_DEFLATED) as z:
    for part, data in entries.items():
        info = ZipInfo(part, (2026, 10, 16, 0, 0, 0))
        info.compress_type = ZIP_DEFLATED
        z.writestr(info, data)
print('Pinned already-correct single-page control', hashlib.sha256(target.read_bytes()).hexdigest().upper())
