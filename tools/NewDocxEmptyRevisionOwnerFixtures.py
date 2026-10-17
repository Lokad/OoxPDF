"""Public empty-run revision ownership, ordering and math composition controls."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib, json
from xml.etree import ElementTree as E
root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/ReviewAuthors'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
E.register_namespace('w', w[1:-1]); E.register_namespace('m', m[1:-1])
existing = json.loads((cases/'nontext-authors.json').read_text(encoding='utf-8'))
control = next(row for row in existing if row['Kind'] == 'empty-insertion')
rows = [dict(control, ExpectedEmptyOwners=1)]
for variant in ['before-text', 'after-text', 'with-math', 'plain-math']:
    source_row = next(row for row in existing if row['Kind'] == ('none' if variant == 'plain-math' else 'empty-insertion'))
    source = root/source_row['Input']; assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == source_row['Sha256']
    with ZipFile(source) as archive: entries = {part: archive.read(part) for part in archive.namelist()}
    document = E.fromstring(entries['word/document.xml']); paragraph = document.find(w+'body').find(w+'p')
    empty = paragraph.find(w+'ins')
    if variant in ['before-text', 'after-text']:
        other = E.SubElement(paragraph, w+'ins', {w+'id':'401', w+'author':'Other Reviewer', w+'date':'2026-10-17T00:00:00Z'})
        E.SubElement(E.SubElement(other, w+'r'), w+'t').text = 'Other inserted phrase.'
        if variant == 'after-text': paragraph.remove(empty); paragraph.append(empty)
        comments = E.fromstring(entries['word/comments.xml']); comments.find(w+'comment').set(w+'author','Nontext Reviewer')
        entries['word/comments.xml'] = E.tostring(comments, encoding='utf-8', xml_declaration=True)
    if variant in ['with-math','plain-math']:
        math = E.SubElement(paragraph, m+'oMath'); fraction = E.SubElement(math,m+'f')
        for name,text in [('num','1'),('den','2')]: E.SubElement(E.SubElement(E.SubElement(fraction,m+name),m+'r'),m+'t').text=text
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    target = cases/('empty-revision-'+variant+'.docx')
    with ZipFile(target,'w',ZIP_DEFLATED) as archive:
        for part,data in sorted(entries.items()):
            info=ZipInfo(part,(2026,10,17,0,0,0));info.compress_type=ZIP_DEFLATED;archive.writestr(info,data)
    rows.append(dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),SourceSha256=source_row['Sha256'],Kind=variant,HeldOut=False,ExpectedEquationCount=1 if variant in ['with-math','plain-math'] else 0,ExpectedEmptyOwners=0 if variant=='plain-math' else 1,ExpectedText=['Public owner paragraph.','Public comment anchor.']+(['Other inserted phrase.'] if variant in ['before-text','after-text'] else [])))
(cases/'empty-revisions.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Pinned five empty-revision ownership controls; one original source reused')
