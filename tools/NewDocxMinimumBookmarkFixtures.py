"""Neutral bookmark controls for existing source-qualified flow frames."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import json, hashlib

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
output = cases / 'CompositionProbes'
output.mkdir(exist_ok=True)
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rows = []
for context, font_size, source_name in [
    ('ordinary12', 12, 'ordinary-minimum-line-20.docx'),
    ('ordinary11', 11, 'ordinary-minimum-line-20.docx'),
    ('noted12', 12, 'note-marked-minimum-line-20.docx')
]:
    source = cases / source_name
    with ZipFile(source) as archive:
        original = {name: archive.read(name) for name in archive.namelist()}
    expected = {'ordinary-minimum-line-20.docx': 'CA9514199B2B8BD44BD0EFF7FF7CAE637AD3A3A0DECDA6ADF584D304B6B4FE74', 'note-marked-minimum-line-20.docx': '7E81696732717C13824B757E4721B1DF5EE414A35205BBE136F5806ABB4D65E4'}
    assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == expected[source_name]
    for bookmarked in [False, True]:
        entries = dict(original)
        document = E.fromstring(entries['word/document.xml'])
        paragraphs = document.find(w + 'body').findall(w + 'p')
        if font_size == 11:
            # The existing eleven-point qualification control uses this same
            # inherited contextual frame and explicit visible-run sizes.
            for paragraph in paragraphs:
                for run in paragraph.findall(w + 'r'):
                    pr = run.find(w + 'rPr')
                    if pr is None:
                        pr = E.Element(w + 'rPr'); run.insert(0, pr)
                    size = pr.find(w + 'sz')
                    if size is None:
                        size = E.SubElement(pr, w + 'sz')
                    size.set(w + 'val', '22')
                    text = run.find(w + 't')
                    if text is not None and text.text:
                        text.text += ' AB.'
                pr = paragraph.find(w + 'pPr')
                spacing = pr.find(w + 'spacing')
                spacing.set(w + 'before', '80'); spacing.set(w + 'after', '80')
                E.SubElement(pr, w + 'contextualSpacing')
        if bookmarked:
            # Enclose an ordinary automatic-height neighbor, not the minimum
            # paragraphs whose visible baselines must remain unchanged.
            paragraph = paragraphs[1]
            first_run = paragraph.find(w + 'r')
            index = list(paragraph).index(first_run)
            paragraph.insert(index, E.Element(w + 'bookmarkStart', {w + 'id': '91', w + 'name': 'PublicNeutralBody'}))
            paragraph.append(E.Element(w + 'bookmarkEnd', {w + 'id': '91'}))
        name = 'minimum-bookmark-' + context + ('-bookmarked' if bookmarked else '-control')
        destination = source if font_size == 12 and not bookmarked else output / (name + '.docx')
        if font_size == 12 and not bookmarked:
            pass  # Reuse the existing public control bytes.
        else:
            entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
            with ZipFile(destination, 'w', ZIP_DEFLATED) as archive:
                for path, data in sorted(entries.items()):
                    info = ZipInfo(path, (2026, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
                    archive.writestr(info, data)
        rows.append(dict(Id=name,Input='tests/Lokad.OoxPdf.Tests/Cases/'+destination.relative_to(cases).as_posix(),
            Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            HeldOut=False,ExpectedEquationCount=0,ExpectedText=['Public minimum Top.', 'Public minimum AfterDelta.'],
            Context=context,FontPoints=font_size,Bookmarked=bookmarked))
(cases/'CompositionProbes/minimum-bookmark.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Six public minimum-leading/bookmark cases generated; existing inputs unchanged.')

# Retain the independently exported neighboring static-frame control unchanged.
source = cases / 'CompositionProbes/plain-automatic-batch-after12.docx'
expected = '2986F0A186FBFA3CC557F3980244F06FF045E5C4EA17A959A2C3C4AECA52BB51'
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == expected
with ZipFile(source) as archive:
    entries = {name: archive.read(name) for name in archive.namelist()}
footer = E.fromstring(entries['word/footer1.xml'])
p = footer.find(w + 'p'); index = list(p).index(p.find(w + 'r'))
p.insert(index, E.Element(w + 'bookmarkStart', {w + 'id': '92', w + 'name': 'PublicNeutralFooter'}))
p.append(E.Element(w + 'bookmarkEnd', {w + 'id': '92'}))
entries['word/footer1.xml'] = E.tostring(footer, encoding='utf-8', xml_declaration=True)
destination = output / 'plain-batch-footer-bookmark.docx'
with ZipFile(destination, 'w', ZIP_DEFLATED) as archive:
    for path, data in sorted(entries.items()):
        info = ZipInfo(path, (2026, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
static_rows = []
for name, path in [('plain-automatic-batch-after12', source), ('plain-batch-footer-bookmark', destination)]:
    static_rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/' + path.relative_to(cases).as_posix(),
        Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(), HeldOut=False, ExpectedEquationCount=0,
        ExpectedText=['Public conclusion.', 'Public second note remains visible']))
(cases / 'CompositionProbes/minimum-bookmark-static.json').write_text(json.dumps(static_rows, indent=2)+'\n', encoding='utf-8')
print('Two public static bookmark cases generated; earlier control preserved.')
