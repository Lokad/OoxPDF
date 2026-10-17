"""Finite revision-role and metric-only empty argument controls."""
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib,json
root=Path(__file__).resolve().parent.parent;cases=root/'tests/Lokad.OoxPdf.Tests/Cases';output=cases/'CompositionProbes'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
def source(role):return cases/('docx-omml-matrix-columns-inline-defaults.docx' if role=='matrix' else 'docx-omml-isolated-inline-body-'+role+'.docx')
def package(path):
 with ZipFile(path) as z:return {n:z.read(n) for n in z.namelist()}
def publish(stem,entries,base,role,kind):
 target=output/(stem+'.docx')
 with ZipFile(target,'w',ZIP_DEFLATED) as z:
  for name,data in sorted(entries.items()):
   info=ZipInfo(name,(2026,10,17,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
 return dict(Id=stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),SourceSha256=hashlib.sha256(base.read_bytes()).hexdigest().upper(),Role=role,Kind=kind,HeldOut=False,ExpectedEquationCount=1,ExpectedText=[])
# Countercontrols keep the original independent probe serialization: before
# registering authoring prefixes, ElementTree uses its stable ns0/ns1 spelling.
counters=[]
for role in ['script','matrix']:
 base=source(role);entries=package(base);xml=E.fromstring(entries['word/document.xml']);eq=next(xml.iter(m+'oMath'))
 owner=eq.find(m+'sSubSup').find(m+'sup') if role=='script' else eq.find(m+'m').find(m+'mr').find(m+'e')
 owner.remove(owner.find(m+'r'));entries['word/document.xml']=E.tostring(xml,encoding='utf-8',xml_declaration=True)
 counters.append(publish('math-role-'+role+'-authored-empty',entries,base,role,'empty'))
base=source('script');entries=package(base);xml=E.fromstring(entries['word/document.xml']);owner=next(xml.iter(m+'sub'));run=owner.find(m+'r');owner.remove(run)
wrapper=E.SubElement(owner,w+'del',{w+'id':'902',w+'author':'Math Role Reviewer',w+'date':'2026-10-17T00:00:00Z'});wrapper.append(run);entries['word/document.xml']=E.tostring(xml,encoding='utf-8',xml_declaration=True)
counters.append(publish('math-role-script-sub-del',entries,base,'script','sub-deletion'))
for kind,suffix in [('off','-off'),('omitted','-omitted'),('hidden','')]:
 base=cases/('docx-omml-matrix-empty-display-diagonal'+suffix+'.docx')
 counters.append(dict(Id='math-role-matrix-empty-'+kind,Input=base.relative_to(root).as_posix(),Sha256=hashlib.sha256(base.read_bytes()).hexdigest().upper(),Role='matrix',Kind=kind,HeldOut=False,ExpectedEquationCount=1,ExpectedText=[]))
E.register_namespace('w',w[1:-1]);E.register_namespace('m',m[1:-1]);rows=[]
for role in ['literal','script','root','matrix']:
 base=source(role)
 for kind in ['ins','del']:
  entries=package(base);xml=E.fromstring(entries['word/document.xml']);eq=next(xml.iter(m+'oMath'));run=eq.find(m+'sSubSup').find(m+'sup').find(m+'r') if role=='script' else next(eq.iter(m+'r'))
  owner=next(parent for parent in eq.iter() if run in list(parent));i=list(owner).index(run);owner.remove(run)
  wrapper=E.Element(w+kind,{w+'id':'901',w+'author':'Math Role Reviewer',w+'date':'2026-10-17T00:00:00Z'});wrapper.append(run);owner.insert(i,wrapper);entries['word/document.xml']=E.tostring(xml,encoding='utf-8',xml_declaration=True)
  rows.append(publish('math-role-'+role+'-'+kind,entries,base,role,kind))
(output/'math-roles.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8');(output/'math-role-counters.json').write_text(json.dumps(counters,indent=2)+'\n',encoding='utf-8')
print('Pinned eight role controls and six countercontrols; three existing inputs reused')
