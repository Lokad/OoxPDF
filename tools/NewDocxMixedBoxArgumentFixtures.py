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
template = E.fromstring(original['word/document.xml'])
matrix = template.find('.//' + m + 'm')
template_run = matrix.find('.//' + m + 'r')

def run(text):
    item = deepcopy(template_run); item.find(m + 't').text = text
    return item

def control(pr):
    rp = E.SubElement(E.SubElement(pr, m + 'ctrlPr'), w + 'rPr')
    E.SubElement(rp, w + 'rFonts', {w + 'ascii': 'Cambria Math', w + 'hAnsi': 'Cambria Math'})
    E.SubElement(rp, w + 'sz', {w + 'val': '24'})

def fraction():
    item = E.Element(m + 'f'); control(E.SubElement(item, m + 'fPr'))
    E.SubElement(item, m + 'num').append(run('a'))
    E.SubElement(item, m + 'den').append(run('b'))
    return item

def script():
    item = E.Element(m + 'sSup'); control(E.SubElement(item, m + 'sSupPr'))
    E.SubElement(item, m + 'e').append(run('c'))
    E.SubElement(item, m + 'sup').append(run('2'))
    return item

rows = []
for variant, owner, child in [
    ('radical-fraction', 'rad', 'fraction'),
    ('radical-matrix', 'rad', 'matrix'),
    ('nary-script', 'nary', 'script'),
    ('nary-matrix', 'nary', 'matrix'),
    ('delimiter-fraction-script', 'd', 'both'),
    ('function-fraction', 'func', 'fraction'),
]:
    entries = dict(original); document = E.fromstring(entries['word/document.xml'])
    equation = document.find('.//' + m + 'oMath'); equation.clear()
    container = E.SubElement(equation, m + owner)
    props = E.SubElement(container, m + owner + 'Pr')
    if owner == 'rad':
        E.SubElement(props, m + 'degHide', {m + 'val': '1'}); E.SubElement(container, m + 'deg')
    elif owner == 'nary':
        E.SubElement(props, m + 'chr', {m + 'val': '\u2211'})
        E.SubElement(props, m + 'limLoc', {m + 'val': 'subSup'})
        E.SubElement(container, m + 'sub').append(run('i'))
        E.SubElement(container, m + 'sup').append(run('n'))
    elif owner == 'func':
        name = run('sin'); E.SubElement(name.find(m + 'rPr'), m + 'nor')
        E.SubElement(container, m + 'fName').append(name)
    control(props)
    arg = E.SubElement(container, m + 'e'); arg.append(run('x+'))
    arg.append(fraction() if child in ('fraction', 'both') else script() if child == 'script' else deepcopy(matrix))
    if child == 'both': arg.append(run('+')); arg.append(script())
    arg.append(run('+y'))
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    name = 'mixed-box-argument-' + variant; dest = output / (name + '.docx')
    with ZipFile(dest, 'w', ZIP_DEFLATED) as archive:
        for path, data in sorted(entries.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + dest.name,
        Sha256=hashlib.sha256(dest.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=1, ExpectedText=['Following prose.'],
        Owner=owner, Child=child))
(output / 'mixed-box-argument.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Six bounded mixed box-argument probes generated; existing public inputs unchanged.')
