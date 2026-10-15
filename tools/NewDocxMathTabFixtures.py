"""Public mixed-math controls for shared default/left tab placement in four containers."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
source = cases / 'CompositionProbes/math-printed-frame-fraction-anchor12-multiple.docx'
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == 'EA7FB764C259DEA27489D7BBDA18BE64CC1AC37C6C091656715B6FC7897C0AFC'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
pkg = '{http://schemas.openxmlformats.org/package/2006/relationships}'
types = '{http://schemas.openxmlformats.org/package/2006/content-types}'
for prefix, namespace in [('w', w), ('m', m), ('r', r)]:
    E.register_namespace(prefix, namespace[1:-1])

def run(parent, value=None):
    node = E.SubElement(parent, w+'r')
    props = E.SubElement(node, w+'rPr')
    E.SubElement(props, w+'rFonts', {w+'ascii':'Arial', w+'hAnsi':'Arial'})
    E.SubElement(props, w+'sz', {w+'val':'24'})
    if value is not None:
        E.SubElement(node, w+'t', {'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text = value
    return node

manifest = []
for stop in ['default18', 'left72']:
    for container in ['body', 'cell', 'header', 'footer']:
        with ZipFile(source) as archive:
            entries = {name:archive.read(name) for name in archive.namelist()}
        document = E.fromstring(entries['word/document.xml'])
        body = document.find(w+'body')
        paragraph = next(p for p in body.findall(w+'p') if p.find(m+'oMath') is not None)
        equations = [deepcopy(node) for node in paragraph.findall(m+'oMath')]
        original_props = paragraph.find(w+'pPr')
        props = deepcopy(original_props) if original_props is not None else E.Element(w+'pPr')
        paragraph.clear()
        paragraph.append(props)
        tabs = E.SubElement(props, w+'tabs')
        if stop == 'left72':
            for position in ['1440', '2880']:
                E.SubElement(tabs, w+'tab', {w+'val':'left', w+'pos':position})
        run(paragraph, 'Start ')
        paragraph.append(equations[0])
        E.SubElement(run(paragraph), w+'tab')
        run(paragraph, 'Middle ')
        paragraph.append(equations[1])
        E.SubElement(run(paragraph), w+'tab')
        run(paragraph, 'End.')
        settings = E.fromstring(entries['word/settings.xml'])
        default = settings.find(w+'defaultTabStop')
        if default is None:
            default = E.SubElement(settings, w+'defaultTabStop')
        default.set(w+'val', '360')
        entries['word/settings.xml'] = E.tostring(settings, encoding='utf-8', xml_declaration=True)
        if container != 'body':
            body.remove(paragraph)
        if container == 'cell':
            table = E.Element(w+'tbl')
            table_props = E.SubElement(table, w+'tblPr')
            E.SubElement(table_props, w+'tblW', {w+'w':'9360', w+'type':'dxa'})
            E.SubElement(table_props, w+'tblLayout', {w+'type':'fixed'})
            borders = E.SubElement(table_props, w+'tblBorders')
            for side in ['top','bottom','left','right','insideV']:
                E.SubElement(borders, w+side, {w+'val':'single', w+'sz':'4'})
            grid = E.SubElement(table, w+'tblGrid')
            row = E.SubElement(table, w+'tr')
            for index in range(2):
                E.SubElement(grid, w+'gridCol', {w+'w':'4680'})
                cell = E.SubElement(row, w+'tc')
                E.SubElement(E.SubElement(cell, w+'tcPr'), w+'tcW', {w+'w':'4680', w+'type':'dxa'})
                if index:
                    cell.append(paragraph)
                else:
                    run(E.SubElement(cell, w+'p'), 'Observed series.')
            body.insert(1, table)
        elif container in ['header','footer']:
            name = 'word/math-tabs-'+container+'.xml'
            part = E.Element(w+('hdr' if container == 'header' else 'ftr'))
            part.append(paragraph)
            entries[name] = E.tostring(part, encoding='utf-8', xml_declaration=True)
            relationships = E.fromstring(entries['word/_rels/document.xml.rels'])
            E.SubElement(relationships, pkg+'Relationship', {'Id':'mathTabsStory','Type':r[1:-1]+'/'+container,'Target':name[5:]})
            entries['word/_rels/document.xml.rels'] = E.tostring(relationships, encoding='utf-8', xml_declaration=True)
            section = body.find(w+'sectPr')
            section.insert(0, E.Element(w+container+'Reference', {w+'type':'default',r+'id':'mathTabsStory'}))
            content_types = E.fromstring(entries['[Content_Types].xml'])
            E.SubElement(content_types, types+'Override', {'PartName':'/'+name,'ContentType':'application/vnd.openxmlformats-officedocument.wordprocessingml.'+container+'+xml'})
            entries['[Content_Types].xml'] = E.tostring(content_types, encoding='utf-8', xml_declaration=True)
        entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
        output = cases / ('CompositionProbes/math-tabs-'+stop+'-'+container+'.docx')
        with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
            for name, data in sorted(entries.items()):
                info = ZipInfo(name, (2026,10,16,0,0,0)); info.compress_type = ZIP_DEFLATED
                archive.writestr(info, data)
        manifest.append(dict(Id=output.stem,Input=output.relative_to(root).as_posix(),Sha256=hashlib.sha256(output.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=2,ExpectedText=['Start','Middle','End.','Following prose.']))
(cases/'CompositionProbes/math-tabs.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
print('Pinned',len(manifest),'mixed-math tab controls')
