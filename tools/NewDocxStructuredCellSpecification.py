from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
parent = json.loads((cases / 'Composition/mixed-argument-specification.json').read_text(encoding='utf-8'))[0]
source = root / parent['Input']
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == parent['Sha256']
matrix_source = cases / 'CompositionProbes/matrix-mixed-cell-ordered-batch.docx'
with ZipFile(source) as archive:
    entries = {name: archive.read(name) for name in archive.namelist()}
with ZipFile(matrix_source) as archive:
    matrix_doc = E.fromstring(archive.read('word/document.xml'))
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
matrix = matrix_doc.find('.//' + m + 'm')
document = E.fromstring(entries['word/document.xml'])
owners = [(owner, child) for owner in document.iter() for child in list(owner) if child.tag == m + 'm']
assert len(owners) == 3
for owner, child in owners:
    index = list(owner).index(child); owner.remove(child); owner.insert(index, deepcopy(matrix))
entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
name = 'composed-structured-cell-specification'
output = cases / 'Composition'
output.mkdir(exist_ok=True)
destination = output / (name + '.docx')
with ZipFile(destination, 'w', ZIP_DEFLATED) as archive:
    for path, data in sorted(entries.items()):
        info = ZipInfo(path, (2026, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
row = dict(parent, Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/Composition/' + destination.name,
    Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(), SourceSha256=parent['Sha256'],
    MatrixSourceSha256=hashlib.sha256(matrix_source.read_bytes()).hexdigest().upper(),
    Features=parent['Features'] + ['mixed-matrix-cells'])
(output / 'structured-cell-specification.json').write_text(json.dumps([row], indent=2) + '\n', encoding='utf-8')
print('One structured-cell/container specification generated; fixed/held-out inputs unchanged.')
