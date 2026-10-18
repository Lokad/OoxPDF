from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
source = cases / 'Composition/composed-operator-matrix-specification.docx'
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == 'D7B869C53D801B3C861513B0489A5B45DF646751C4938710D1D8408CBBF4614E'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
with ZipFile(source) as archive:
    entries = {name: archive.read(name) for name in archive.namelist()}
document = E.fromstring(entries['word/document.xml']); body = document.find(w + 'body')
direct = [equation for paragraph in body.findall(w + 'p') for equation in paragraph.iter(m + 'oMath')]
cell = body.find('.//' + w + 'tbl//' + m + 'oMath')
assert len(direct) == 4 and cell is not None
for equation in [direct[0], direct[1], cell]:
    container = equation.find(m + 'rad')
    if container is None: container = equation.find(m + 'nary')
    argument = container.find(m + 'e')
    template = argument.find('.//' + m + 'r')
    def run(text):
        child = deepcopy(template); child.find(m + 't').text = text
        return child
    argument.insert(0, run('x+')); argument.append(run('+y'))
for target, name in [(direct[2], 'function-fraction'), (direct[3], 'delimiter-fraction-script')]:
    path = cases / 'CompositionProbes' / ('mixed-box-argument-' + name + '.docx')
    with ZipFile(path) as archive:
        equation = E.fromstring(archive.read('word/document.xml')).find('.//' + m + 'oMath')
    target[:] = [deepcopy(child) for child in equation]
entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
name = 'composed-mixed-argument-specification'
output = cases / 'Composition'
output.mkdir(exist_ok=True)
destination = output / (name + '.docx')
with ZipFile(destination, 'w', ZIP_DEFLATED) as archive:
    for path, data in sorted(entries.items()):
        info = ZipInfo(path, (2026, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
parent = json.loads((cases / 'Composition/operator-matrix-specification.json').read_text(encoding='utf-8'))[0]
row = dict(parent, Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/Composition/' + destination.name,
    Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
    SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
    Features=parent['Features'] + ['mixed-box-arguments', 'function', 'delimiter'])
(output / 'mixed-argument-specification.json').write_text(json.dumps([row], indent=2) + '\n', encoding='utf-8')
print('One composite mixed-argument specification generated; fixed portfolio inputs unchanged.')
