"""Independent empty-fraction/deletion structure and size controls."""
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import json,hashlib
root=Path(__file__).resolve().parent.parent;cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
E.register_namespace('w',w[1:-1]);E.register_namespace('m',m[1:-1])
source=cases/'equation-revision-body-del.docx';sourcehash=hashlib.sha256(source.read_bytes()).hexdigest().upper();rows=[]
for size,kind in [(11,'removed'),(11,'empty-row'),(11,'empty-fraction'),(9,'deleted'),(9,'empty-fraction'),(18,'deleted'),(18,'empty-fraction')]:
 with ZipFile(source) as z:entries={n:z.read(n) for n in z.namelist()}
 xml=E.fromstring(entries['word/document.xml']);p=next(p for p in xml.iter(w+'p') if p.find(w+'del') is not None);wrapper=p.find(w+'del');i=list(p).index(wrapper)
 if kind!='deleted':
  p.remove(wrapper)
  if kind!='removed':
   para=E.Element(m+'oMathPara');eq=E.SubElement(para,m+'oMath')
   if kind=='empty-fraction':
    f=E.SubElement(eq,m+'f');E.SubElement(f,m+'num');E.SubElement(f,m+'den')
   p.insert(i,para)
 entries['word/document.xml']=E.tostring(xml,encoding='utf-8',xml_declaration=True)
 if size!=11:
  styles=E.fromstring(entries['word/styles.xml']);styles.find(w+'docDefaults').find(w+'rPrDefault').find(w+'rPr').find(w+'sz').set(w+'val',str(size*2));entries['word/styles.xml']=E.tostring(styles,encoding='utf-8',xml_declaration=True)
 stem='final-equation-flow-'+kind+('' if size==11 else '-'+str(size));target=cases/(stem+'.docx')
 with ZipFile(target,'w',ZIP_DEFLATED) as z:
  for name,data in sorted(entries.items()):
   info=ZipInfo(name,(2026,10,17,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
 rows.append(dict(Id=stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),SourceSha256=sourcehash,FontSize=size,Kind=kind,HeldOut=False,ExpectedEquationCount=3 if kind=='removed' else 4,ExpectedText=['share the same input series.']))
(cases/'empty-fractions.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8');print('Pinned',len(rows),'empty fraction controls')
