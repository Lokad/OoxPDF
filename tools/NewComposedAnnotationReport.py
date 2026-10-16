"""Pin a complete acceptance report combining independently supported inline features.

Keep this document outside derivation: reduce failures to separate small controls.
"""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json

root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases'
source=cases/'Composition/report.docx'
assert hashlib.sha256(source.read_bytes()).hexdigest().upper()=='10F63D162D0D46DBE8CD682BB50452DB6672104706F116EA50CDCAFEE054575F'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
r='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
for prefix,ns in [('w',w),('m',m),('r',r)]:E.register_namespace(prefix,ns[1:-1])
with ZipFile(source) as archive:entries={name:archive.read(name) for name in archive.namelist()}
document=E.fromstring(entries['word/document.xml']);body=document.find(w+'body')
parents={child:parent for parent in document.iter() for child in parent}
note=next(document.iter(w+'footnoteReference'));note_run=deepcopy(parents[note]);parents[note].remove(note)
for child in list(note_run):
    if child.tag not in [w+'rPr',w+'footnoteReference']:note_run.remove(child)
for parent in document.iter():
    for child in list(parent):
        if child.tag in [w+'commentRangeStart',w+'commentRangeEnd',w+'commentReference']:parent.remove(child)
paragraph=next(p for p in body.findall(w+'p') if p.find('.//'+w+'drawing') is not None)
equation=paragraph.find(m+'oMath');assert equation is not None
bookmark_id=str(max([int(node.get(w+'id')) for node in document.iter(w+'bookmarkStart')]+[0])+1)
bookmark_name='ComposedCalculation'
index=list(paragraph).index(equation)
paragraph.insert(index,E.Element(w+'bookmarkStart',{w+'id':bookmark_id,w+'name':bookmark_name}))
paragraph.insert(index+2,E.Element(w+'bookmarkEnd',{w+'id':bookmark_id}))
existing_link=next(node for node in document.iter(w+'hyperlink') if node.get(r+'id'))
text_run=next(node for node in paragraph.findall(w+'r') if node.find(w+'t') is not None and node.find(w+'drawing') is None)
index=list(paragraph).index(text_run);paragraph.remove(text_run)
link=E.Element(w+'hyperlink',{r+'id':existing_link.get(r+'id')});link.append(text_run);paragraph.insert(index,link)
internal=E.SubElement(paragraph,w+'hyperlink',{w+'anchor':bookmark_name})
E.SubElement(E.SubElement(internal,w+'r'),w+'t',{'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text=' See calculation.'
paragraph.insert(1 if paragraph.find(w+'pPr') is not None else 0,E.Element(w+'commentRangeStart',{w+'id':'0'}))
E.SubElement(paragraph,w+'commentRangeEnd',{w+'id':'0'})
E.SubElement(E.SubElement(paragraph,w+'r'),w+'commentReference',{w+'id':'0'})
cell_paragraph=next(p for cell in body.iter(w+'tc') for p in cell.findall(w+'p') if p.find(m+'oMath') is not None)
cell_paragraph.append(note_run)
entries['word/document.xml']=E.tostring(document,encoding='utf-8',xml_declaration=True)
output=cases/'Composition/composed-annotations-report.docx'
with ZipFile(output,'w',ZIP_DEFLATED) as archive:
    for name,data in sorted(entries.items()):
        info=ZipInfo(name,(2026,10,16,0,0,0));info.compress_type=ZIP_DEFLATED;archive.writestr(info,data)
manifest=[dict(Id='composed-annotations-report',Input=output.relative_to(root).as_posix(),Sha256=hashlib.sha256(output.read_bytes()).hexdigest().upper(),HeldOut=True,ExpectedEquationCount=6,ExpectedInlineImageCount=1,ExpectedText=['Quarterly operations report','See calculation.','Review section 12.','All quantities in this synthetic example are illustrative.'],Features=['image/math/comment/link/bookmark in one paragraph','math/footnote reference in a table cell','numbering','header equation','footer','continuous footnote','wrapped prose'])]
(cases/'Composition/composed-annotations-report.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
print(manifest[0]['Sha256'])
