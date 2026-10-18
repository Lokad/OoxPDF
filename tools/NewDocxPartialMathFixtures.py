from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json

root = Path(__file__).resolve().parent.parent
source = root / 'tests/Lokad.OoxPdf.Tests/Cases/docx-omml-fraction-inline.docx'
output = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
output.mkdir(exist_ok=True)
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == 'FFFCECCC3DD4F72C7013F6E8760CF376B673D6FE5E6AC1D47199DC566CAC9213'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
pkg = '{http://schemas.openxmlformats.org/package/2006/relationships}'
ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
with ZipFile(source) as archive:
    original = {name: archive.read(name) for name in archive.namelist()}

def text(value):
    run = E.Element(w + 'r'); E.SubElement(run, w + 't', {'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text = value
    return run

rows = []
for container in ['body', 'cell', 'header']:
    for unsupported in [False, True]:
        entries = dict(original); document = E.fromstring(entries['word/document.xml'])
        body = document.find(w + 'body'); paragraph = body.find(w + 'p')
        equation = paragraph.find(m + 'oMath'); missing = deepcopy(equation)
        fraction = missing.find(m + 'f'); pr = E.Element(m + 'fPr'); fraction.insert(0, pr)
        E.SubElement(pr, m + 'type', {m + 'val':'skw'})
        template = fraction.find('.//' + m + 'r')
        for role, value in [('num','p'), ('den','q')]:
            argument = fraction.find(m + role); run = deepcopy(template); run.find(m + 't').text = value; argument[:] = [run]
        paragraph.append(text(' Between formulae: '))
        if unsupported: paragraph.append(missing)
        paragraph.append(text(' after formula.'))
        if container == 'cell':
            table = E.Element(w + 'tbl'); props = E.SubElement(table, w + 'tblPr')
            E.SubElement(props, w + 'tblW', {w + 'w':'9360', w + 'type':'dxa'})
            margins = E.SubElement(props, w + 'tblCellMar')
            for edge in ['top','left','bottom','right']: E.SubElement(margins, w + edge, {w + 'w':'0', w + 'type':'dxa'})
            E.SubElement(E.SubElement(table, w + 'tblGrid'), w + 'gridCol', {w + 'w':'9360'})
            cell = E.SubElement(E.SubElement(table, w + 'tr'), w + 'tc')
            E.SubElement(E.SubElement(cell, w + 'tcPr'), w + 'tcW', {w + 'w':'9360', w + 'type':'dxa'})
            body.remove(paragraph); cell.append(paragraph); body.insert(0, table)
        elif container == 'header':
            header = E.Element(w + 'hdr'); header.append(deepcopy(paragraph))
            entries['word/headerPartial.xml'] = E.tostring(header, encoding='utf-8', xml_declaration=True)
            relation_path = 'word/_rels/document.xml.rels'
            relations = E.fromstring(entries[relation_path]) if relation_path in entries else E.Element(pkg + 'Relationships')
            assert all(rel.get('Id') != 'rIdPartialHeader' for rel in relations)
            E.SubElement(relations, pkg + 'Relationship', {'Id':'rIdPartialHeader', 'Type':r[1:-1]+'/header', 'Target':'headerPartial.xml'})
            entries[relation_path] = E.tostring(relations, encoding='utf-8', xml_declaration=True)
            types = E.fromstring(entries['[Content_Types].xml'])
            E.SubElement(types, ct + 'Override', {'PartName':'/word/headerPartial.xml', 'ContentType':'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'})
            entries['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
            section = body.find(w + 'sectPr'); section.insert(0, E.Element(w + 'headerReference', {w + 'type':'default', r + 'id':'rIdPartialHeader'}))
            section.find(w + 'pgMar').set(w + 'header', '720')
            paragraph[:] = [text('Body probe.')]
        entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
        name = 'partial-math-' + container + ('-unsupported' if unsupported else '-control'); dest = output / (name + '.docx')
        with ZipFile(dest, 'w', ZIP_DEFLATED) as archive:
            for path, data in sorted(entries.items()):
                info = ZipInfo(path, (2026, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + dest.name,
            Sha256=hashlib.sha256(dest.read_bytes()).hexdigest().upper(), SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            HeldOut=False, ExpectedEquationCount=2 if unsupported else 1, ExpectedText=['Before formula:', 'Between formulae:', 'after formula.', 'Following prose.'],
            Container=container, UnsupportedNeighbor=unsupported))
(output / 'partial-math.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Six partial-math body/cell/header probes generated; existing public inputs unchanged.')
