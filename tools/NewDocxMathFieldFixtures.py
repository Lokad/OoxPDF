"""Public cached-field/math controls; keep dynamic page fields outside this slice."""
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

def run(parent, text=None):
    node = E.SubElement(parent, w+'r')
    props = E.SubElement(node, w+'rPr')
    E.SubElement(props, w+'rFonts', {w+'ascii':'Arial', w+'hAnsi':'Arial'})
    E.SubElement(props, w+'sz', {w+'val':'24'})
    if text is not None:
        E.SubElement(node, w+'t', {'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text = text
    return node

def field(parent, form, value):
    instruction = ' REF Cache'+value+' '
    if form == 'simple':
        run(E.SubElement(parent, w+'fldSimple', {w+'instr':instruction}), value)
    else:
        node = run(parent)
        E.SubElement(node, w+'fldChar', {w+'fldCharType':'begin'})
        E.SubElement(node, w+'instrText', {'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text = instruction
        E.SubElement(node, w+'fldChar', {w+'fldCharType':'separate'})
        E.SubElement(node, w+'t').text = value
        E.SubElement(node, w+'fldChar', {w+'fldCharType':'end'})

manifest = []
for form in ['simple', 'packed']:
    for container in ['body','cell','header','footer']:
        with ZipFile(source) as archive:
            entries = {name:archive.read(name) for name in archive.namelist()}
        document = E.fromstring(entries['word/document.xml'])
        body = document.find(w+'body')
        heading = body.find(w+'p')
        props = deepcopy(heading.find(w+'pPr'))
        heading.clear()
        if props is not None: heading.append(props)
        for index, value in enumerate(['North','South']):
            if index: run(heading, ' ')
            E.SubElement(heading, w+'bookmarkStart', {w+'id':str(310+index), w+'name':'Cache'+value})
            run(heading, value)
            E.SubElement(heading, w+'bookmarkEnd', {w+'id':str(310+index)})
        paragraph = next(p for p in body.findall(w+'p') if p.find(m+'oMath') is not None)
        equations = [deepcopy(node) for node in paragraph.findall(m+'oMath')]
        props = deepcopy(paragraph.find(w+'pPr'))
        paragraph.clear()
        if props is not None: paragraph.append(props)
        run(paragraph, 'Start '); field(paragraph, form, 'North'); run(paragraph, ' ')
        paragraph.append(equations[0]); run(paragraph, ' middle ')
        field(paragraph, form, 'South'); run(paragraph, ' ')
        paragraph.append(equations[1]); run(paragraph, ' End.')
        if container != 'body': body.remove(paragraph)
        if container == 'cell':
            table = E.Element(w+'tbl'); props = E.SubElement(table,w+'tblPr')
            E.SubElement(props,w+'tblW',{w+'w':'9360',w+'type':'dxa'})
            E.SubElement(props,w+'tblLayout',{w+'type':'fixed'})
            borders = E.SubElement(props,w+'tblBorders')
            for side in ['top','bottom','left','right','insideV']:
                E.SubElement(borders,w+side,{w+'val':'single',w+'sz':'4'})
            grid = E.SubElement(table,w+'tblGrid'); row = E.SubElement(table,w+'tr')
            for index in range(2):
                E.SubElement(grid,w+'gridCol',{w+'w':'4680'})
                cell = E.SubElement(row,w+'tc')
                E.SubElement(E.SubElement(cell,w+'tcPr'),w+'tcW',{w+'w':'4680',w+'type':'dxa'})
                if index: cell.append(paragraph)
                else: run(E.SubElement(cell,w+'p'),'Observed series.')
            body.insert(1,table)
        elif container in ['header','footer']:
            name = 'word/math-fields-'+container+'.xml'
            part = E.Element(w+('hdr' if container == 'header' else 'ftr')); part.append(paragraph)
            entries[name] = E.tostring(part,encoding='utf-8',xml_declaration=True)
            relationships = E.fromstring(entries['word/_rels/document.xml.rels'])
            E.SubElement(relationships,pkg+'Relationship',{'Id':'mathFieldsStory','Type':r[1:-1]+'/'+container,'Target':name[5:]})
            entries['word/_rels/document.xml.rels'] = E.tostring(relationships,encoding='utf-8',xml_declaration=True)
            body.find(w+'sectPr').insert(0,E.Element(w+container+'Reference',{w+'type':'default',r+'id':'mathFieldsStory'}))
            content_types = E.fromstring(entries['[Content_Types].xml'])
            E.SubElement(content_types,types+'Override',{'PartName':'/'+name,'ContentType':'application/vnd.openxmlformats-officedocument.wordprocessingml.'+container+'+xml'})
            entries['[Content_Types].xml'] = E.tostring(content_types,encoding='utf-8',xml_declaration=True)
        entries['word/document.xml'] = E.tostring(document,encoding='utf-8',xml_declaration=True)
        output = cases / ('CompositionProbes/math-fields-'+form+'-'+container+'.docx')
        with ZipFile(output,'w',ZIP_DEFLATED) as archive:
            for name,data in sorted(entries.items()):
                info=ZipInfo(name,(2026,10,16,0,0,0));info.compress_type=ZIP_DEFLATED
                archive.writestr(info,data)
        manifest.append(dict(Id=output.stem,Input=output.relative_to(root).as_posix(),Sha256=hashlib.sha256(output.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=2,ExpectedText=['Start','North','South','middle','End.','Following prose.']))
(cases/'CompositionProbes/math-fields.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
print('Pinned',len(manifest),'cached-field composition controls')
