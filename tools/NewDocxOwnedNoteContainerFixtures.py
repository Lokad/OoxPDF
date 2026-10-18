# Generate bounded intermediate and dedicated owned-note container controls.
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import json,hashlib
root=Path('.').resolve();source=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/note-footer-frame-body-tail.docx';out=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes';w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';xml='{http://www.w3.org/XML/1998/namespace}'
with ZipFile(source) as z:original={n:z.read(n) for n in z.namelist()}
rows=[]
for dedicated in [False,True]:
 parts=dict(original);notes=E.fromstring(parts['word/footnotes.xml']);p=next(n for n in notes if n.get(w+'id')=='1').find(w+'p');sp=p.find(w+'pPr').find(w+'spacing');sp.attrib.pop(w+'line',None);sp.attrib.pop(w+'lineRule',None)
 repeat=150 if dedicated else 80
 E.SubElement(E.SubElement(p,w+'r'),w+'t',{xml+'space':'preserve'}).text=' Continued public note content remains present.'*repeat+' Final tail.'
 parts['word/footnotes.xml']=E.tostring(notes,encoding='utf-8',xml_declaration=True);doc=E.fromstring(parts['word/document.xml']);body=doc.find(w+'body');following=next(p for p in body if 'Following prose.' in ''.join(p.itertext()))
 if dedicated:
  leading=next(p for p in body if 'Leading paragraph.' in ''.join(p.itertext()))
  for child in list(following):
   if child.tag in [w+'commentRangeStart',w+'commentRangeEnd'] or child.find(w+'commentReference') is not None:leading.append(child)
  body.remove(following)
 else:
  pr=following.find(w+'pPr')
  if pr is None:pr=E.Element(w+'pPr');following.insert(0,pr)
  sp=pr.find(w+'spacing')
  if sp is None:sp=E.SubElement(pr,w+'spacing')
  sp.set(w+'line','1600');sp.set(w+'lineRule','exact')
 parts['word/document.xml']=E.tostring(doc,encoding='utf-8',xml_declaration=True);name='owned-note-container-'+('dedicated' if dedicated else 'intermediate');path=out/(name+'.docx')
 with ZipFile(path,'w',ZIP_DEFLATED) as z:
  for n,data in sorted(parts.items()):
   info=ZipInfo(n,(2026,1,1,0,0,0));info.create_system=0;info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
 rows.append(dict(Id=name,Input=path.relative_to(root).as_posix(),Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(),SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=4,ExpectedInlineImageCount=0,ExpectedText=['Public calculation note.','Final tail.','Public review footer line 3.']+([] if dedicated else ['Following prose.']),RepeatedSentences=repeat,DeferredBodyHeight=0 if dedicated else 80))
(out/'owned-note-containers.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Two minimal single-note intermediate/dedicated controls pinned.')
