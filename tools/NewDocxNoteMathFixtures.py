"""Public note-story composition controls from pinned automatic-note inputs."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
from copy import deepcopy
import hashlib, json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
E.register_namespace('w', w[1:-1])
manifest = []
for kind in ['footnote', 'endnote']:
    sources = json.loads((cases / ('math-' + kind + 's.json')).read_text(encoding='utf-8'))
    source = next(row for row in sources if row['Id'] == 'math-' + kind + '-before-body')
    path = root / source['Input']
    assert hashlib.sha256(path.read_bytes()).hexdigest().upper() == source['Sha256']
    with ZipFile(path) as archive:
        original = {name:archive.read(name) for name in archive.namelist()}
    for flow in ['inline', 'wrapped']:
        entries = dict(original)
        document = E.fromstring(entries['word/document.xml'])
        equation = next(document.iter(m + 'oMath'))
        part = 'word/' + kind + 's.xml'
        notes = E.fromstring(entries[part])
        note = next(node for node in notes if int(node.get(w + 'id', '0')) > 0)
        paragraph = note.find(w + 'p')
        def text(value):
            run = E.SubElement(paragraph, w + 'r')
            E.SubElement(run, w + 't', {'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text = value
        if flow == 'wrapped':
            text(' The supporting calculation retains every source item when the note wraps onto another line. ' * 3)
        else:
            text(' Calculation: ')
        E.SubElement(paragraph, w + 'bookmarkStart', {w + 'id':'55001', w + 'name':'NoteCalculation'})
        paragraph.append(deepcopy(equation))
        E.SubElement(paragraph, w + 'bookmarkEnd', {w + 'id':'55001'})
        text(' and ')
        paragraph.append(deepcopy(equation))
        text(' End of supported note.')
        entries[part] = E.tostring(notes, encoding='utf-8', xml_declaration=True)
        target = cases / ('math-note-content-' + kind + '-' + flow + '.docx')
        with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
            for name,data in sorted(entries.items()):
                info = ZipInfo(name, (2026,10,16,0,0,0)); info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        manifest.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), HeldOut=False,
            ExpectedEquationCount=4, ExpectedInlineImageCount=0,
            ExpectedText=['Public calculation note.', 'End of supported note.'],
            Features=[kind + ' content', flow + ' prose', 'two note fractions', 'bookmark', 'two body fractions']))
(cases / 'math-note-content.json').write_text(json.dumps(manifest, indent=2)+'\n', encoding='utf-8')
print('Pinned four normal-note story controls with two supported note fractions each.')
