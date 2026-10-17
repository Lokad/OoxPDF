"""Paired ordinary paragraphs differing only by an actual empty insertion."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json
root=Path(__file__).resolve().parent.parent
output=root/'tests/Lokad.OoxPdf.Tests/Cases/RevisionNeutrality';output.mkdir(exist_ok=True)
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
E.register_namespace('w',w[1:-1]);rows=[]
for size,minimum in [(11,16),(12,18)]:
 for variant in ['control','empty','empty-style']:
    document=E.Element(w+'document');body=E.SubElement(document,w+'body')
    for index,text in enumerate(['Public first paragraph AVATAR To Wa.','Public second paragraph.']):
        paragraph=E.SubElement(body,w+'p');pp=E.SubElement(paragraph,w+'pPr')
        E.SubElement(pp,w+'spacing',{w+'before':'0',w+'after':'0',w+'line':str(minimum*20),w+'lineRule':'atLeast'})
        run=E.SubElement(paragraph,w+'r');rp=E.SubElement(run,w+'rPr')
        E.SubElement(rp,w+'rFonts',{w+'ascii':'Calibri',w+'hAnsi':'Calibri'})
        E.SubElement(rp,w+'sz',{w+'val':str(size*2)});E.SubElement(rp,w+'kern',{w+'val':'0'})
        E.SubElement(run,w+'t').text=text
        if variant.startswith('empty') and index==0:
            insertion=E.SubElement(paragraph,w+'ins',{w+'id':'7',w+'author':'Empty Reviewer',w+'date':'2026-10-17T00:00:00Z'})
            empty_run=E.SubElement(insertion,w+'r')
            if variant=='empty-style':
                rp=E.SubElement(empty_run,w+'rPr');E.SubElement(rp,w+'rFonts',{w+'ascii':'Arial',w+'hAnsi':'Arial'})
                E.SubElement(rp,w+'sz',{w+'val':'36'})
            E.SubElement(empty_run,w+'t').text=''
    section=E.SubElement(body,w+'sectPr');E.SubElement(section,w+'pgSz',{w+'w':'12240',w+'h':'15840'})
    E.SubElement(section,w+'pgMar',{w+'top':'1440',w+'bottom':'1440',w+'left':'1440',w+'right':'1440'})
    styles=E.Element(w+'styles');defaults=E.SubElement(E.SubElement(E.SubElement(styles,w+'docDefaults'),w+'rPrDefault'),w+'rPr')
    E.SubElement(defaults,w+'rFonts',{w+'ascii':'Calibri',w+'hAnsi':'Calibri'});E.SubElement(defaults,w+'sz',{w+'val':str(size*2)})
    parts={
      '[Content_Types].xml':b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
      '_rels/.rels':b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
      'word/_rels/document.xml.rels':b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
      'word/document.xml':E.tostring(document,encoding='utf-8',xml_declaration=True),
      'word/styles.xml':E.tostring(styles,encoding='utf-8',xml_declaration=True)}
    target=output/('revision-neutrality-'+str(size)+'-'+variant+'.docx')
    with ZipFile(target,'w',ZIP_DEFLATED) as archive:
        for name,data in sorted(parts.items()):
            info=ZipInfo(name,(2026,10,17,0,0,0));info.compress_type=ZIP_DEFLATED;archive.writestr(info,data)
    rows.append(dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),Pair='minimum-'+str(size),Variant=variant,HeldOut=False,ExpectedEquationCount=0,ExpectedText=['Public first paragraph AVATAR To Wa.','Public second paragraph.']))
(output/'revision-neutrality.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Pinned six ordinary minimum-height controls; paired source differs only by empty insertion')
