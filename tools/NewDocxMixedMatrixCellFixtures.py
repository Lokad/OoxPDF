from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json

root = Path(__file__).resolve().parent.parent
source = root / 'tests/Lokad.OoxPdf.Tests/Cases/docx-omml-matrix-inline-literal.docx'
output = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
output.mkdir(exist_ok=True)
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == '17DDB69EE5AE5DC9E84E1A1707BA368044A880D7E747547E70659EF1C58D4BFC'
with ZipFile(source) as archive:
    original = {name: archive.read(name) for name in archive.namelist()}
template = E.fromstring(original['word/document.xml']).find('.//' + m + 'r')

def run(text):
    child = deepcopy(template); child.find(m + 't').text = text
    return child

def control(pr):
    rp = E.SubElement(E.SubElement(pr, m + 'ctrlPr'), w + 'rPr')
    E.SubElement(rp, w + 'rFonts', {w + 'ascii': 'Cambria Math', w + 'hAnsi': 'Cambria Math'})
    E.SubElement(rp, w + 'sz', {w + 'val': '24'})

def fraction():
    child = E.Element(m + 'f'); control(E.SubElement(child, m + 'fPr'))
    E.SubElement(child, m + 'num').append(run('p'))
    E.SubElement(child, m + 'den').append(run('q'))
    return child

def script():
    child = E.Element(m + 'sSup'); control(E.SubElement(child, m + 'sSupPr'))
    E.SubElement(child, m + 'e').append(run('u'))
    E.SubElement(child, m + 'sup').append(run('2'))
    return child

rows = []
for variant in ['fraction', 'script', 'ordered-batch']:
    entries = dict(original); document = E.fromstring(entries['word/document.xml'])
    matrix = document.find('.//' + m + 'm'); cells = matrix.findall(m + 'mr/' + m + 'e')
    changes = [(cells[0], script() if variant == 'script' else fraction())]
    if variant == 'ordered-batch': changes.append((cells[2], script()))
    for cell, child in changes:
        cell[:] = [run('x+'), child, run('+y')]
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    name = 'matrix-mixed-cell-' + variant; dest = output / (name + '.docx')
    with ZipFile(dest, 'w', ZIP_DEFLATED) as archive:
        for path, data in sorted(entries.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + dest.name,
        Sha256=hashlib.sha256(dest.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=1, ExpectedText=['Following prose.'], Variant=variant))
(output / 'mixed-matrix-cell.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Three bounded mixed matrix-cell controls generated; public inputs unchanged.')
