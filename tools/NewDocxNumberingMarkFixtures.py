"""Place actual inserted numbering marks beside plain and mixed source content."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib
import json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
rows = []
for kind in ['plain', 'math']:
    for story in ['body', 'cell', 'header', 'footer']:
        source = cases / ('list-decoration-mark-inserted.docx' if kind == 'plain'
            else 'math-run-format-' + story + '-bold.docx')
        with ZipFile(source) as z:
            entries = {part: z.read(part) for part in z.namelist()}
        if kind == 'plain' and story == 'footer':
            target = source  # Reuse its two independently exported C73 references.
        else:
            if kind == 'plain':
                content = E.fromstring(entries['word/footer1.xml'])
                paragraphs = list(content.findall(w + 'p'))
                assert len(paragraphs) == 3
                document = E.fromstring(entries['word/document.xml'])
                body = document.find(w + 'body')
                section = body.find(w + 'sectPr')
                for reference in list(section.findall(w + 'footerReference')):
                    section.remove(reference)
                if story == 'header':
                    header = E.Element(w + 'hdr')
                    header.extend(paragraphs)
                    entries['word/header1.xml'] = E.tostring(header, encoding='utf-8', xml_declaration=True)
                    entries['word/_rels/header1.xml.rels'] = entries['word/_rels/footer1.xml.rels']
                    E.SubElement(section, w + 'headerReference', {w + 'type': 'default', r + 'id': 'header'})
                elif story == 'body':
                    for paragraph in paragraphs:
                        body.insert(list(body).index(section), paragraph)
                else:
                    table = E.Element(w + 'tbl')
                    properties = E.SubElement(table, w + 'tblPr')
                    E.SubElement(properties, w + 'tblW', {w + 'w': '9360', w + 'type': 'dxa'})
                    E.SubElement(properties, w + 'tblLayout', {w + 'type': 'fixed'})
                    E.SubElement(E.SubElement(table, w + 'tblGrid'), w + 'gridCol', {w + 'w': '9360'})
                    cell = E.SubElement(E.SubElement(table, w + 'tr'), w + 'tc')
                    E.SubElement(E.SubElement(cell, w + 'tcPr'), w + 'tcW', {w + 'w': '9360', w + 'type': 'dxa'})
                    cell.extend(paragraphs)
                    body.insert(list(body).index(section), table)
                entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
            else:
                part = 'word/' + story + '1.xml' if story in ['header', 'footer'] else 'word/document.xml'
                document = E.fromstring(entries[part])
                paragraph = next(p for p in document.iter(w + 'p') if list(p.iter(m + 'oMath')))
                properties = paragraph.find(w + 'pPr')
                assert properties is not None and properties.find(w + 'numPr') is not None
                mark = properties.find(w + 'rPr')
                if mark is None:
                    mark = E.SubElement(properties, w + 'rPr')
                E.SubElement(mark, w + 'ins', {w + 'id': '310', w + 'author': 'Public Reviewer', w + 'date': '2026-10-17T00:00:00Z'})
                entries[part] = E.tostring(document, encoding='utf-8', xml_declaration=True)
            target = cases / ('numbering-mark-' + kind + '-' + story + '.docx')
            with ZipFile(target, 'w', ZIP_DEFLATED) as z:
                for part, data in entries.items():
                    info = ZipInfo(part, (2026, 10, 17, 0, 0, 0))
                    info.compress_type = ZIP_DEFLATED
                    z.writestr(info, data)
        rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
            Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            Kind=kind, Story=story,
            ExpectedEquationCount=sum(len(list(E.fromstring(data).iter(m + 'oMath')))
                for part, data in entries.items() if part.endswith('.xml')),
            SelectedSourceEquations=4 if kind == 'math' else 0,
            HeldOut=False, ExpectedText=['share the same input series.']))
(cases / 'numbering-marks.json').write_text(json.dumps(rows, indent=2)+'\n', encoding='utf-8')
print('Pinned',len(rows),'numbering-mark placement controls; one existing source reused')
