from pathlib import Path
from xml.etree import ElementTree as E
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
import hashlib,json
root=Path(__file__).resolve().parent.parent; source=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/partial-math-body-unsupported.docx'; output=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'; output.mkdir(exist_ok=True)
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == '2637A68BD493B5537B292F58F9BDF458F046D35A3CBD9B0FFDB5843E45ADE07B'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
rows=json.loads((output/'partial-math.json').read_text(encoding='utf-8'))
for kind,bodySize,childSize,controlSize in [('small8',8,8,None),('control20',11,12,20)]:
 with ZipFile(source) as archive: entries={n:archive.read(n) for n in archive.namelist()}
 doc=E.fromstring(entries['word/document.xml']); p=doc.find('.//'+w+'body/'+w+'p'); p.remove(p.find(m+'oMath'))
 for run in doc.iter(w+'r'):
  pr=run.find(w+'rPr')
  if pr is None:pr=E.Element(w+'rPr');run.insert(0,pr)
  sz=pr.find(w+'sz')
  if sz is None:sz=E.SubElement(pr,w+'sz')
  sz.set(w+'val',str(2*bodySize))
 for run in doc.iter(m+'r'):
  pr=run.find(w+'rPr')
  if pr is None:pr=E.Element(w+'rPr');run.insert(0,pr)
  sz=pr.find(w+'sz')
  if sz is None:sz=E.SubElement(pr,w+'sz')
  sz.set(w+'val',str(2*childSize))
 pr=p.find(w+'pPr')
 if pr is None:pr=E.Element(w+'pPr');p.insert(0,pr)
 mark=pr.find(w+'rPr')
 if mark is None:mark=E.SubElement(pr,w+'rPr')
 sz=mark.find(w+'sz')
 if sz is None:sz=E.SubElement(mark,w+'sz')
 sz.set(w+'val',str(2*bodySize))
 if controlSize:
  ctrl=E.SubElement(doc.find('.//'+m+'fPr'),m+'ctrlPr'); cp=E.SubElement(ctrl,w+'rPr')
  E.SubElement(cp,w+'rFonts',{w+'ascii':'Cambria Math',w+'hAnsi':'Cambria Math'});E.SubElement(cp,w+'sz',{w+'val':str(2*controlSize)});E.SubElement(cp,w+'color',{w+'val':'D02020'})
 entries['word/document.xml']=E.tostring(doc,encoding='utf-8',xml_declaration=True)
 name='skewed-fraction-'+kind; dest=output/(name+'.docx')
 with ZipFile(dest,'w',ZIP_DEFLATED) as a:
  for n,data in sorted(entries.items()):
   info=ZipInfo(n,(2026,1,1,0,0,0));info.compress_type=ZIP_DEFLATED;a.writestr(info,data)
 rows.append(dict(Id=name,Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/'+dest.name,Sha256=hashlib.sha256(dest.read_bytes()).hexdigest().upper(),SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=1,ExpectedText=['Before formula:','Between formulae:','after formula.','Following prose.'],Container='body',BodySize=bodySize,ChildSize=childSize,ControlSize=controlSize or childSize))
(output/'skewed-fraction.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Two source-size/control-owner contrasts generated.')
