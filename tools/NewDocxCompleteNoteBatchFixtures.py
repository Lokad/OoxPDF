"""Public automatic-marker/batch controls for the shared complete-note frame."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
from copy import deepcopy
import json, hashlib

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
output = cases / 'CompositionProbes'
output.mkdir(exist_ok=True)
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = []
for name, plain, after, count in [
    ('compound-automatic-note-after12', False, 12, 1),
    ('plain-automatic-note-after6', True, 6, 1),
    ('plain-automatic-note-after12', True, 12, 1),
    ('plain-automatic-batch-after12', True, 12, 2),
]:
    source = cases / ('Composition/composed-operator-matrix-specification.docx' if plain else
                      'CompositionProbes/compact-note-no-header-note-after12.docx')
    expected = ('D7B869C53D801B3C861513B0489A5B45DF646751C4938710D1D8408CBBF4614E' if plain else
                next(row['Sha256'] for row in json.loads((cases / 'CompositionProbes/compact-note-spacing.json').read_text(encoding='utf-8'))
                     if row['Id'] == 'compact-note-no-header-note-after12'))
    assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == expected
    with ZipFile(source) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    document = E.fromstring(entries['word/document.xml'])
    body = document.find(w + 'body')
    if plain:
        section = deepcopy(body.find(w + 'sectPr'))
        for ref in section.findall(w + 'headerReference'):
            section.remove(ref)
        section.find(w + 'pgSz').set(w + 'h', '16000')
        body.clear()
        def paragraph(text):
            p = E.Element(w + 'p')
            E.SubElement(E.SubElement(p, w + 'r'), w + 't').text = text
            return p
        # Round 800pt page and ordinary inherited paragraph spacing; one/two
        # notes bracket a compact boundary without fitting a production offset.
        for index in range(1, 33 - count):
            body.append(paragraph(f'Public body entry {index}. Check shared note capacity.'))
        last = paragraph('Public conclusion. Supporting note')
        for note_id in range(1, count + 1):
            if note_id > 1:
                E.SubElement(E.SubElement(last, w + 'r'), w + 't',
                             {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = ' Additional note'
            E.SubElement(E.SubElement(last, w + 'r'), w + 'footnoteReference', {w + 'id': str(note_id)})
        body.extend([last, section])
        header = E.Element(w + 'hdr')
        header.append(paragraph('Public plain frame header'))
        entries['word/header1.xml'] = E.tostring(header, encoding='utf-8', xml_declaration=True)
    notes = E.fromstring(entries['word/footnotes.xml'])
    first = next(note for note in notes.findall(w + 'footnote') if note.get(w + 'id') == '1')
    if count == 2:
        second = deepcopy(first)
        second.set(w + 'id', '2')
        second.find('.//' + w + 't').text = 'Public second note remains visible in the shared batch.'
        notes.append(second)
    for note in notes.findall(w + 'footnote'):
        if note.get(w + 'type') is not None:
            continue
        p = note.find(w + 'p')
        pr = p.find(w + 'pPr')
        if pr is None:
            pr = E.Element(w + 'pPr'); p.insert(0, pr)
        spacing = pr.find(w + 'spacing')
        if spacing is None:
            spacing = E.SubElement(pr, w + 'spacing')
        spacing.set(w + 'after', str((6 if count == 2 and note.get(w + 'id') == '1' else after) * 20))
        mark = E.Element(w + 'r')
        E.SubElement(E.SubElement(mark, w + 'rPr'), w + 'vertAlign', {w + 'val': 'superscript'})
        E.SubElement(mark, w + 'footnoteRef')
        gap = E.Element(w + 'r')
        E.SubElement(gap, w + 't', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = ' '
        p.insert(1, mark); p.insert(2, gap)
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    entries['word/footnotes.xml'] = E.tostring(notes, encoding='utf-8', xml_declaration=True)
    destination = output / (name + '.docx')
    with ZipFile(destination, 'w', ZIP_DEFLATED) as archive:
        for path, data in sorted(entries.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    rows.append(dict(Id=name,Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/'+destination.name,
        Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
        SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
        HeldOut=False,ExpectedEquationCount=0 if plain else 6,
        ExpectedText=['Public conclusion.' if plain else 'Review conclusion:',
                      'All quantities in this synthetic example are illustrative.'],
        PlainFrame=plain,NoteCount=count,FinalNoteAfterPoints=after))
(cases/'CompositionProbes/complete-note-batch.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Four public note-target/batch controls generated; no existing inputs changed.')
