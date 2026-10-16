"""Independent small page-admission controls for an indivisible footnote head."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
from copy import deepcopy
import json, hashlib
root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
source=cases/'math-footnote-after-cell-plain.docx'
with ZipFile(source) as z: base={n:z.read(n) for n in z.namelist()}
rows=[]
for container in ['cell','body']:
    entries=base.copy();doc=E.fromstring(entries['word/document.xml']);body=doc.find(w+'body')
    sect=deepcopy(body.find(w+'sectPr'));table=deepcopy(body.find(w+'tbl'))
    for node in list(body):body.remove(node)
    for cell in table.findall('.//'+w+'tc'):
        for p in cell.findall(w+'p'):
            for node in list(p):
                if node.tag in [m+'oMath',m+'oMathPara']:p.remove(node)
    marker=table.find('.//'+w+'footnoteReference')
    owner=next(r for r in table.iter(w+'r') if marker in list(r))
    if container=='body':
        for p in table.iter(w+'p'):
            if owner in list(p):p.remove(owner)
    body.append(table)
    for i in range(6):
        p=E.SubElement(body,w+'p');pr=E.SubElement(p,w+'pPr')
        E.SubElement(pr,w+'spacing',{w+'before':'0',w+'after':'80',w+'line':'240',w+'lineRule':'exact'})
        E.SubElement(pr,w+'widowControl',{w+'val':'0'})
        for j in range(3):
            r=E.SubElement(p,w+'r');rp=E.SubElement(r,w+'rPr')
            E.SubElement(rp,w+'rFonts',{w+'ascii':'Arial',w+'hAnsi':'Arial'})
            E.SubElement(rp,w+'sz',{w+'val':'22'})
            E.SubElement(r,w+'t').text=f'Section {i+1} line {j+1}.'
            if j<2:E.SubElement(r,w+'br')
        if i==0 and container=='body':p.append(owner)
    size=sect.find(w+'pgSz');size.set(w+'h','7440');size.set(w+'w','8640')
    margins=sect.find(w+'pgMar')
    for edge in ['top','bottom','left','right']:margins.set(w+edge,'720')
    body.append(sect);entries['word/document.xml']=E.tostring(doc,encoding='utf-8',xml_declaration=True)
    notes=E.fromstring(entries['word/footnotes.xml'])
    sep=E.SubElement(notes,w+'footnote',{w+'type':'separator',w+'id':'-1'})
    p=E.SubElement(sep,w+'p');pr=E.SubElement(p,w+'pPr');E.SubElement(pr,w+'spacing',{w+'after':'160'})
    E.SubElement(E.SubElement(p,w+'r'),w+'separator')
    for p in notes.findall('./'+w+'footnote[@'+w+'id="1"]/'+w+'p'):
        pr=E.SubElement(p,w+'pPr');E.SubElement(pr,w+'spacing',{w+'after':'120'})
    entries['word/footnotes.xml']=E.tostring(notes,encoding='utf-8',xml_declaration=True)
    path=cases/f'complete-note-{container}.docx'
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for name,data in sorted(entries.items()):
            info=ZipInfo(name,(2026,10,16,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
    rows.append(dict(Id=path.stem,Input=path.relative_to(root).as_posix(),Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=0,ExpectedText=['Section 6 line 3.','Public calculation note.']))
(cases/'complete-notes.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Pinned two independent complete-note page-admission controls')
