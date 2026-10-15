"""Public equation/comment composition controls with prose, formula and mixed anchors."""
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
E.register_namespace('w', w[1:-1]); E.register_namespace('m', m[1:-1])

def run(parent, value):
    node = E.SubElement(parent, w+'r')
    E.SubElement(node, w+'t', {'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text = value
    return node

manifest = []
for anchor in ['prose', 'between', 'formula', 'mixed']:
    for container in ['body', 'cell']:
        with ZipFile(source) as archive:
            entries = {name:archive.read(name) for name in archive.namelist()}
        document = E.fromstring(entries['word/document.xml'])
        for parent in document.iter():
            for node in list(parent):
                if node.tag in [w+'commentRangeStart', w+'commentRangeEnd', w+'commentReference']:
                    parent.remove(node)
        body = document.find(w+'body')
        paragraph = next(p for p in body.findall(w+'p') if p.find(m+'oMath') is not None)
        equations = [deepcopy(node) for node in paragraph.findall(m+'oMath')]
        paragraph.clear()
        run(paragraph, 'Start ')
        paragraph.append(equations[0])
        run(paragraph, ' middle ')
        paragraph.append(equations[1])
        run(paragraph, ' End.')
        children = list(paragraph)
        start, end = {'prose':(0,1), 'between':(2,3), 'formula':(1,2), 'mixed':(0,5)}[anchor]
        paragraph.insert(end, E.Element(w+'commentRangeEnd', {w+'id':'0'}))
        paragraph.insert(start, E.Element(w+'commentRangeStart', {w+'id':'0'}))
        marker = E.SubElement(paragraph, w+'r')
        E.SubElement(marker, w+'commentReference', {w+'id':'0'})
        comments = E.Element(w+'comments')
        comment = E.SubElement(comments, w+'comment', {w+'id':'0',w+'author':'Public reviewer',w+'initials':'RV'})
        run(E.SubElement(comment,w+'p'), 'Check composed calculation.')
        entries['word/comments.xml'] = E.tostring(comments,encoding='utf-8',xml_declaration=True)
        if container == 'cell':
            body.remove(paragraph)
            table = E.Element(w+'tbl')
            props = E.SubElement(table,w+'tblPr')
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
        entries['word/document.xml'] = E.tostring(document,encoding='utf-8',xml_declaration=True)
        output = cases / ('CompositionProbes/math-comment-'+anchor+'-'+container+'.docx')
        with ZipFile(output,'w',ZIP_DEFLATED) as archive:
            for name,data in sorted(entries.items()):
                info=ZipInfo(name,(2026,10,16,0,0,0));info.compress_type=ZIP_DEFLATED
                archive.writestr(info,data)
        manifest.append(dict(Id=output.stem,Input=output.relative_to(root).as_posix(),Sha256=hashlib.sha256(output.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=2,ExpectedText=['Start','middle','End.','Following prose.']))
(cases/'CompositionProbes/math-comments.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
print('Pinned',len(manifest),'equation/comment controls')
