"""Combine independent public note and minimum-footer sources across body/empty tails."""
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib,json
from argparse import ArgumentParser
root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
parser=ArgumentParser();parser.add_argument('--output-directory',type=Path,default=cases)
out=parser.parse_args().output_directory.resolve();out.mkdir(parents=True,exist_ok=True)
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';r='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';rel='{http://schemas.openxmlformats.org/package/2006/relationships}';ct='{http://schemas.openxmlformats.org/package/2006/content-types}'
with ZipFile(cases/'minimum-note-page-boundary.docx') as z:original={n:z.read(n) for n in z.namelist()}
with ZipFile(cases/'minimum-static-footer.docx') as z:footer=z.read('word/footerMinimumPlain.xml')
rows=[]
for tail in [False,True]:
 parts=dict(original);doc=E.fromstring(parts['word/document.xml']);body=doc.find(w+'body');section=body.find(w+'sectPr');section.insert(0,E.Element(w+'footerReference',{w+'type':'default',r+'id':'rIdNoteFooter'}));section.find(w+'pgMar').set(w+'footer','720')
 if tail:
  following=next(p for p in body if 'Following prose.' in ''.join(p.itertext()));leading=next(p for p in body if 'Leading paragraph.' in ''.join(p.itertext()))
  for child in list(following):
   if child.tag in [w+'commentRangeStart',w+'commentRangeEnd'] or child.find(w+'commentReference') is not None: leading.append(child)
  body.remove(following)
 parts['word/document.xml']=E.tostring(doc,encoding='utf-8',xml_declaration=True);parts['word/footerNoteFrame.xml']=footer
 relationships=E.fromstring(parts['word/_rels/document.xml.rels']);E.SubElement(relationships,rel+'Relationship',dict(Id='rIdNoteFooter',Type=r[1:-1]+'/footer',Target='footerNoteFrame.xml'));parts['word/_rels/document.xml.rels']=E.tostring(relationships,encoding='utf-8',xml_declaration=True)
 types=E.fromstring(parts['[Content_Types].xml']);E.SubElement(types,ct+'Override',dict(PartName='/word/footerNoteFrame.xml',ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'));parts['[Content_Types].xml']=E.tostring(types,encoding='utf-8',xml_declaration=True)
 name='note-footer-frame'+('-empty-tail' if tail else '-body-tail');path=out/(name+'.docx')
 with ZipFile(path,'w',ZIP_DEFLATED) as z:
  for n,data in sorted(parts.items()):
   info=ZipInfo(n,(2026,1,1,0,0,0));info.create_system=0;info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
 rows.append(dict(Id=name,Input=path.relative_to(root).as_posix(),Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=4,ExpectedInlineImageCount=0,ExpectedText=['Public calculation note.','End of supported note.','Public review footer line 3.']+([] if tail else ['Following prose.'])))
(out/'note-footer-frames.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Two independent source-backed note/footer frame probes pinned.')
