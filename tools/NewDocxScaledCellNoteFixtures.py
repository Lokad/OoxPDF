"""A scaled ordinary cell strut and its custom-note clickable paragraph slot."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import hashlib
import json

root = Path(__file__).resolve().parent.parent
output = root / "tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes"
w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
rel = "http://schemas.openxmlformats.org/package/2006/relationships"
ct = "http://schemas.openxmlformats.org/package/2006/content-types"
rows = []
for size in [11, 12]:
    properties = f'<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="{size * 2}"/><w:kern w:val="0"/>'
    spacing = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
    def run(text, extra=""):
        return f'<w:r><w:rPr>{properties}{extra}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
    paragraph = f'<w:p><w:pPr>{spacing}<w:rPr>{properties}</w:rPr></w:pPr>' + run('Public cell before note ')
    paragraph += f'<w:r><w:rPr>{properties}<w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2" w:customMarkFollows="1"/><w:t>*</w:t></w:r>'
    paragraph += run(' and after note.') + '</w:p>'
    table = '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar>'
    table += ''.join(f'<w:{edge} w:w="0" w:type="dxa"/>' for edge in ['top', 'left', 'bottom', 'right'])
    table += '</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>' + paragraph + '</w:tc></w:tr></w:tbl>'
    anchor = f'<w:p><w:pPr>{spacing}<w:rPr>{properties}</w:rPr></w:pPr><w:commentRangeStart w:id="9"/>' + run('Public review anchor') + '<w:commentRangeEnd w:id="9"/><w:r><w:commentReference w:id="9"/></w:r></w:p>'
    section = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>'
    entries = {
        '_rels/.rels': f'<Relationships xmlns="{rel}"><Relationship Id="document" Type="{r}/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/document.xml': f'<w:document xmlns:w="{w}"><w:body>{table}{anchor}{section}</w:body></w:document>',
        'word/styles.xml': f'<w:styles xmlns:w="{w}"><w:docDefaults><w:rPrDefault><w:rPr>{properties}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>{spacing}</w:pPr></w:pPrDefault></w:docDefaults></w:styles>',
        'word/footnotes.xml': f'<w:footnotes xmlns:w="{w}"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:pPr>{spacing}</w:pPr>' + run('*', '<w:vertAlign w:val="superscript"/>') + run(' Public note target.') + '</w:p></w:footnote></w:footnotes>',
        'word/comments.xml': f'<w:comments xmlns:w="{w}"><w:comment w:id="9" w:author="Public author" w:initials="PA"><w:p>{run("Public frame check.")}</w:p></w:comment></w:comments>',
        'word/_rels/document.xml.rels': f'<Relationships xmlns="{rel}">' + ''.join(f'<Relationship Id="{kind}" Type="{r}/{kind}" Target="{name}"/>' for kind, name in [('styles', 'styles.xml'), ('footnotes', 'footnotes.xml'), ('comments', 'comments.xml')]) + '</Relationships>',
        '[Content_Types].xml': f'<Types xmlns="{ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + ''.join(f'<Override PartName="/word/{name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>' for kind, name in [('document.main', 'document.xml'), ('styles', 'styles.xml'), ('footnotes', 'footnotes.xml'), ('comments', 'comments.xml')]) + '</Types>',
    }
    target = output / f'scaled-cell-note-calibri-{size}.docx'
    with ZipFile(target, 'w', ZIP_DEFLATED) as z:
        for name, text in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 18, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, '<?xml version="1.0" encoding="UTF-8"?>' + text)
    rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
        Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(), FontSize=size,
        HeldOut=False, ExpectedEquationCount=0, ExpectedImagePlacements=0,
        ExpectedText=['Public cell before note', 'and after note.', 'Public review anchor', 'Public note target.']))
(output / 'scaled-cell-notes.json').write_text(json.dumps(rows, indent=2) + '\n', encoding='utf-8')
print('Pinned two ordinary-cell note controls; comment markup selects scaled Review geometry')
