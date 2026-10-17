"""Paint owners survive empty function names and structural-only math."""
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import json,hashlib
root=Path(__file__).resolve().parent.parent;cases=root/'tests/Lokad.OoxPdf.Tests/Cases';output=cases/'CompositionProbes'
m='{http://schemas.openxmlformats.org/officeDocument/2006/math}';w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';E.register_namespace('m',m[1:-1]);E.register_namespace('w',w[1:-1]);rows=[]
roles=[('function-name','docx-omml-function-display-sin-literal','func',['fName']),('root-body','docx-omml-isolated-inline-body-root','rad',['e']),('delimiter-argument','docx-omml-delimiter-display-brackets','d',['e']),('nary-only','docx-omml-nary-display-sum-undOvr','nary',['sub','sup','e'])]
families=[('math-source',roles,941,'Math Source Reviewer'),
 ('math-degree',[(profile,'docx-omml-root-degree-'+profile,'rad',['deg']) for profile in
  ['display-cube','inline-cube','display-large-control','inline-large-control']],951,'Math Degree Reviewer')]
for prefix,roles,first_id,author in families:
 rows=[]
 for role,base,node,args in roles:
  for kind in ['del','empty']:
   source=cases/(base+'.docx')
   with ZipFile(source) as z:entries={name:z.read(name) for name in z.namelist()}
   xml=E.fromstring(entries['word/document.xml']);structure=next(xml.iter(m+node))
   for i,arg in enumerate(args):
    owner=structure.find(m+arg);runs=list(owner.findall(m+'r'));assert runs
    for run in runs:owner.remove(run)
    if kind=='del':E.SubElement(owner,w+'del',{w+'id':str(first_id+i),w+'author':author,w+'date':'2026-10-17T00:00:00Z'}).extend(runs)
   entries['word/document.xml']=E.tostring(xml,encoding='utf-8',xml_declaration=True);target=output/(prefix+'-'+role+'-'+kind+'.docx')
   with ZipFile(target,'w',ZIP_DEFLATED) as z:
    for name,data in sorted(entries.items()):
     info=ZipInfo(name,(2026,10,17,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
   rows.append(dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),Role=role,Kind=kind,HeldOut=False,ExpectedEquationCount=1,ExpectedText=[]))
 (output/(prefix+'.json')).write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8');print('Pinned '+str(len(rows))+' '+prefix+' controls')
