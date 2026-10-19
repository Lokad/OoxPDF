"""Verify authored-row hyperlink controls; --write publishes source-exact inputs."""
import argparse
import copy
import hashlib
import json
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as E

root = Path(__file__).resolve().parents[1]
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
source = cases / 'docx-omml-isolated-inline-body-fraction-left.docx'
sha = lambda value: hashlib.sha256(value).hexdigest().upper()
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--write', action='store_true')
args = parser.parse_args()
data = source.read_bytes()
assert sha(data) == '96A12C11555C1CF0C9360F0939551E77BE12180A69411CC534D56297E7B25B4F'
with ZipFile(BytesIO(data)) as archive:
    infos = archive.infolist()
    parts = {entry.filename: archive.read(entry.filename) for entry in infos}
base = E.fromstring(parts['word/document.xml'])
formula = next(base.iter(m + 'oMath'))
assert len(list(base.iter(m + 'oMath'))) == 1
source_math = E.tostring(formula)
rows, jobs = [], []
for variant in ['unlinked', 'linked-prose', 'linked-math']:
    tree = copy.deepcopy(base)
    paragraph = next(p for p in tree.iter(w + 'p') if list(p.iter(m + 'oMath')))
    formula = next(paragraph.iter(m + 'oMath'))
    assert formula in list(paragraph), 'The control requires implicit inline source math.'
    at = list(paragraph).index(formula)
    paragraph.remove(formula)
    text = E.Element(w + 'r')
    E.SubElement(text, w + 't').text = 'Reference row.'
    link = E.Element(w + 'hyperlink', {w + 'anchor': 'RowCalculation'})
    link.append(text if variant == 'linked-prose' else formula)
    line_break = E.Element(w + 'r')
    E.SubElement(line_break, w + 'br')
    content = [E.Element(w + 'bookmarkStart', {w + 'id': '108', w + 'name': 'RowCalculation'}),
        E.Element(w + 'bookmarkEnd', {w + 'id': '108'}),
        link if variant == 'linked-prose' else text,
        line_break, link if variant == 'linked-math' else formula]
    for offset, child in enumerate(content):
        paragraph.insert(at + offset, child)
    assert E.tostring(next(tree.iter(m + 'oMath'))) == source_math
    changed = parts.copy()
    changed['word/document.xml'] = E.tostring(tree, encoding='utf-8', xml_declaration=True)
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as result:
        for entry in infos:
            result.writestr(copy.copy(entry), changed[entry.filename])
    generated = buffer.getvalue()
    name = 'math-authored-row-' + variant
    rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + name + '.docx',
        Sha256=sha(generated), SourceInput=source.relative_to(root).as_posix(), SourceSha256=sha(data),
        SourceMathSha256=sha(source_math), Variant=variant, HeldOut=variant != 'unlinked',
        ExpectedEquationCount=1, ExpectedRenderedEquationCount=1, ExpectedText=['Reference row.'],
        Change='Add the same authored text row before an unchanged implicit equation; hyperlink ownership is ' + variant + '.'))
    jobs.append((cases / 'CompositionProbes' / (name + '.docx'), generated))
jobs.append((cases / 'CompositionProbes/math-authored-row-links.json',
    (json.dumps(rows, indent=2) + '\n').replace('\n', '\r\n').encode()))
for target, generated in jobs:
    if args.write:
        stage = target.with_suffix(target.suffix + '.staging')
        with stage.open('xb') as output:
            output.write(generated)
        stage.replace(target)
    else:
        assert target.read_bytes() == generated, target.name
print('Three source-exact authored-row hyperlink controls verified.')
