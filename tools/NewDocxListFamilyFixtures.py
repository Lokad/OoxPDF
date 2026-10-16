"""Independent numbering-level/paragraph-mark font-family controls."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
source = cases / 'static-frame-footer-plain-first.docx'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = []

def set_family(properties, family):
    fonts = properties.find(w + 'rFonts')
    if fonts is None:
        fonts = E.SubElement(properties, w + 'rFonts')
    fonts.attrib.clear()
    fonts.set(w + 'ascii', family)
    fonts.set(w + 'hAnsi', family)

for name in ['run-arial', 'mark-times', 'level-courier', 'normal-times', 'defaults-arial', 'no-family-default']:
    with ZipFile(source) as z:
        entries = {part: z.read(part) for part in z.namelist()}
    footer = E.fromstring(entries['word/footer1.xml'])
    first = footer.find(w + 'p')
    run_properties = first.find(w + 'r/' + w + 'rPr')
    for change in list(run_properties.findall(w + 'rPrChange')):
        run_properties.remove(change)
    set_family(run_properties, 'Times New Roman' if name == 'defaults-arial' else 'Arial')
    if name in ['mark-times', 'level-courier']:
        mark = E.SubElement(first.find(w + 'pPr'), w + 'rPr')
        set_family(mark, 'Times New Roman')
    if name == 'level-courier':
        numbering = E.fromstring(entries['word/numbering.xml'])
        level = numbering.find(w + 'abstractNum/' + w + 'lvl')
        set_family(E.SubElement(level, w + 'rPr'), 'Courier New')
        entries['word/numbering.xml'] = E.tostring(numbering, encoding='utf-8', xml_declaration=True)
    styles = E.fromstring(entries['word/styles.xml'])
    if name == 'normal-times':
        normal = E.SubElement(styles, w + 'style', {w + 'type': 'paragraph', w + 'styleId': 'Normal', w + 'default': '1'})
        set_family(E.SubElement(normal, w + 'rPr'), 'Times New Roman')
    if name == 'defaults-arial':
        set_family(styles.find(w + 'docDefaults/' + w + 'rPrDefault/' + w + 'rPr'), 'Arial')
    if name == 'no-family-default':
        properties = styles.find(w + 'docDefaults/' + w + 'rPrDefault/' + w + 'rPr')
        properties.remove(properties.find(w + 'rFonts'))
    entries['word/styles.xml'] = E.tostring(styles, encoding='utf-8', xml_declaration=True)
    entries['word/footer1.xml'] = E.tostring(footer, encoding='utf-8', xml_declaration=True)
    target = cases / ('list-family-' + name + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as z:
        for part, data in entries.items():
            info = ZipInfo(part, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        ExpectedEquationCount=0, HeldOut=False, ExpectedText=['share the same input series.']))
(cases / 'list-families.json').write_text(json.dumps(rows, indent=2)+'\n', encoding='utf-8')
print('Prepared',len(rows),'public family controls')
