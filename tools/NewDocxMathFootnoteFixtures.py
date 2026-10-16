"""Public continuous-footnote controls around mixed normal-text fractions."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json, argparse

parser=argparse.ArgumentParser()
parser.add_argument('--without-comments',action='store_true')
args=parser.parse_args()

root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases'
source=cases/'CompositionProbes/math-printed-frame-fraction-anchor12-multiple.docx'
assert hashlib.sha256(source.read_bytes()).hexdigest().upper()=='EA7FB764C259DEA27489D7BBDA18BE64CC1AC37C6C091656715B6FC7897C0AFC'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
pkg='{http://schemas.openxmlformats.org/package/2006/relationships}'
rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
types='{http://schemas.openxmlformats.org/package/2006/content-types}'
E.register_namespace('w',w[1:-1]);E.register_namespace('m',m[1:-1])
manifest=[]
for position in ['before','after']:
    for container in ['body','cell']:
        with ZipFile(source) as archive:
            entries={name:archive.read(name) for name in archive.namelist()}
        doc=E.fromstring(entries['word/document.xml']);body=doc.find(w+'body')
        if args.without_comments:
            for parent in doc.iter():
                for node in list(parent):
                    if node.tag in [w+'commentRangeStart',w+'commentRangeEnd',w+'commentReference']:
                        parent.remove(node)
            entries.pop('word/comments.xml')
        paragraph=next(p for p in body.findall(w+'p') if p.find(m+'oMath') is not None)
        marker=E.Element(w+'r');props=E.SubElement(marker,w+'rPr')
        E.SubElement(props,w+'rFonts',{w+'ascii':'Arial',w+'hAnsi':'Arial'})
        E.SubElement(props,w+'sz',{w+'val':'24'})
        E.SubElement(props,w+'vertAlign',{w+'val':'superscript'})
        E.SubElement(marker,w+'footnoteReference',{w+'id':'1'})
        equation=paragraph.find(m+'oMath')
        paragraph.insert(list(paragraph).index(equation) if position=='before' else len(paragraph),marker)
        notes=E.Element(w+'footnotes')
        note=E.SubElement(notes,w+'footnote',{w+'id':'1'})
        note_paragraph=E.SubElement(note,w+'p')
        E.SubElement(E.SubElement(note_paragraph,w+'r'),w+'footnoteRef')
        E.SubElement(E.SubElement(note_paragraph,w+'r'),w+'t',{'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text='Public calculation note.'
        entries['word/footnotes.xml']=E.tostring(notes,encoding='utf-8',xml_declaration=True)
        relationships=E.fromstring(entries['word/_rels/document.xml.rels'])
        if args.without_comments:
            for node in list(relationships):
                if node.attrib.get('Type')==rel+'comments':relationships.remove(node)
        E.SubElement(relationships,pkg+'Relationship',{'Id':'mathFootnotes','Type':rel+'footnotes','Target':'footnotes.xml'})
        entries['word/_rels/document.xml.rels']=E.tostring(relationships,encoding='utf-8',xml_declaration=True)
        ct=E.fromstring(entries['[Content_Types].xml'])
        if args.without_comments:
            for node in list(ct):
                if node.attrib.get('PartName')=='/word/comments.xml':ct.remove(node)
        E.SubElement(ct,types+'Override',{'PartName':'/word/footnotes.xml','ContentType':'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'})
        entries['[Content_Types].xml']=E.tostring(ct,encoding='utf-8',xml_declaration=True)
        section=body.find(w+'sectPr');settings=E.SubElement(section,w+'footnotePr')
        E.SubElement(settings,w+'numRestart',{w+'val':'continuous'})
        if container=='cell':
            body.remove(paragraph);table=E.Element(w+'tbl');props=E.SubElement(table,w+'tblPr')
            E.SubElement(props,w+'tblW',{w+'w':'9360',w+'type':'dxa'});E.SubElement(props,w+'tblLayout',{w+'type':'fixed'})
            grid=E.SubElement(table,w+'tblGrid');row=E.SubElement(table,w+'tr')
            for index in range(2):
                E.SubElement(grid,w+'gridCol',{w+'w':'4680'});cell=E.SubElement(row,w+'tc')
                E.SubElement(E.SubElement(cell,w+'tcPr'),w+'tcW',{w+'w':'4680',w+'type':'dxa'})
                if index:cell.append(paragraph)
                else:E.SubElement(E.SubElement(E.SubElement(cell,w+'p'),w+'r'),w+'t').text='Observed series.'
            body.insert(1,table)
        entries['word/document.xml']=E.tostring(doc,encoding='utf-8',xml_declaration=True)
        output=cases/('CompositionProbes/math-footnote-'+position+'-'+container+('-plain' if args.without_comments else '')+'.docx')
        with ZipFile(output,'w',ZIP_DEFLATED) as archive:
            for name,data in sorted(entries.items()):
                info=ZipInfo(name,(2026,10,16,0,0,0));info.compress_type=ZIP_DEFLATED;archive.writestr(info,data)
        manifest.append(dict(Id=output.stem,Input=output.relative_to(root).as_posix(),Sha256=hashlib.sha256(output.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=2,ExpectedText=['Before formula:','Following prose.','Public calculation note.']))
(cases/('CompositionProbes/math-footnotes'+('-plain' if args.without_comments else '')+'.json')).write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
print('Pinned',len(manifest),'continuous footnote composition controls')
