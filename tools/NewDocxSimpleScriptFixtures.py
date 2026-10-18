from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
output = cases / 'CompositionProbes'
output.mkdir(exist_ok=True)
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = []
for context, source_name, replacement in [
    ('latin12', 'ordinary-minimum-line-20.docx', 'Alph\u00e1'),
    ('greek12', 'ordinary-minimum-line-20.docx', '\u0391\u03bb\u03c6\u03b1'),
    ('cyrillic12', 'ordinary-minimum-line-20.docx', '\u0410\u043b\u044c\u0444\u0430'),
    ('latin11', 'CompositionProbes/minimum-bookmark-ordinary11-control.docx', 'Alph\u00e1'),
    ('latin-noted12', 'note-marked-minimum-line-20.docx', 'Alph\u00e1'),
]:
    source = cases / source_name
    expected = {
        'ordinary-minimum-line-20.docx': 'CA9514199B2B8BD44BD0EFF7FF7CAE637AD3A3A0DECDA6ADF584D304B6B4FE74',
        'CompositionProbes/minimum-bookmark-ordinary11-control.docx': 'F8AF84291A8A82232E0BF16DADE46540F8B2B7F4C4847F345954AC375EAE8B3B',
        'note-marked-minimum-line-20.docx': '7E81696732717C13824B757E4721B1DF5EE414A35205BBE136F5806ABB4D65E4',
    }
    assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == expected[source_name]
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    doc = E.fromstring(entries['word/document.xml'])
    paragraphs = doc.find(w + 'body').findall(w + 'p')
    text = paragraphs[1].find('.//' + w + 't')
    assert text is not None and 'Alpha' in text.text
    text.text = text.text.replace('Alpha', replacement)
    entries['word/document.xml'] = E.tostring(doc, encoding='utf-8', xml_declaration=True)
    name = 'minimum-simple-script-' + context
    destination = output / (name + '.docx')
    with ZipFile(destination, 'w', ZIP_DEFLATED) as archive:
        for path, data in sorted(entries.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + destination.name,
        Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False, ExpectedEquationCount=0,
        ExpectedText=['Public minimum Top.', 'Public minimum AfterDelta.'],
        Context=context, Replacement=replacement))
(output / 'minimum-simple-script.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Five simple-script neighbors generated; minimum paragraphs unchanged.')
