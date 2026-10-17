"""Function/operator remainders after deleting a literal argument or body."""
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import json,hashlib
root=Path(__file__).resolve().parent.parent;cases=root/'tests/Lokad.OoxPdf.Tests/Cases';output=cases/'CompositionProbes'
m='{http://schemas.openxmlformats.org/officeDocument/2006/math}';w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';E.register_namespace('m',m[1:-1]);E.register_namespace('w',w[1:-1]);rows=[]
roles=[('function-argument','docx-omml-function-display-sin-literal','func'),('nary-stacked-body','docx-omml-nary-display-sum-undOvr','nary'),('nary-side-body','docx-omml-nary-default-inline-sum-setting-side','nary')]
for role,base,node in roles:
 for kind in ['del','empty']:
  source=cases/(base+'.docx')
  with ZipFile(source) as z:entries={name:z.read(name) for name in z.namelist()}
  xml=E.fromstring(entries['word/document.xml']);structure=next(xml.iter(m+node));owner=structure.find(m+'e');run=owner.find(m+'r');owner.remove(run)
  if kind=='del':E.SubElement(owner,w+'del',{w+'id':'931',w+'author':'Math Body Reviewer',w+'date':'2026-10-17T00:00:00Z'}).append(run)
  entries['word/document.xml']=E.tostring(xml,encoding='utf-8',xml_declaration=True);target=output/('math-body-'+role+'-'+kind+'.docx')
  with ZipFile(target,'w',ZIP_DEFLATED) as z:
   for name,data in sorted(entries.items()):
    info=ZipInfo(name,(2026,10,17,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
  rows.append(dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),Role=role,Kind=kind,HeldOut=False,ExpectedEquationCount=1,ExpectedText=[]))
(output/'math-body.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8');print('Pinned six empty argument/body controls')
