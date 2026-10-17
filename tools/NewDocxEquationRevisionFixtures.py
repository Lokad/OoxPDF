"""Source-owned equation revisions beside unchanged composed content."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json
root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
E.register_namespace('w',w[1:-1]); E.register_namespace('m',m[1:-1])
rows=[]
for story,part in [('body','word/document.xml'),('cell','word/document.xml'),('header','word/header1.xml'),('footer','word/footer1.xml')]:
 source=cases/('math-numbered-display-'+story+'.docx')
 for kind in ['ins','del']:
  with ZipFile(source) as z: entries={name:z.read(name) for name in z.namelist()}
  xml=E.fromstring(entries[part]);paragraph=next(p for p in xml.iter(w+'p') if list(p.iter(m+'oMath')))
  equation=next(e for e in paragraph if e.tag in [m+'oMath',m+'oMathPara'])
  index=list(paragraph).index(equation);paragraph.remove(equation)
  revision=E.Element(w+kind,{w+'id':'871',w+'author':'Equation Reviewer',w+'date':'2026-10-17T00:00:00Z'})
  revision.append(equation);paragraph.insert(index,revision)
  entries[part]=E.tostring(xml,encoding='utf-8',xml_declaration=True)
  target=cases/('equation-revision-'+story+'-'+kind+'.docx')
  with ZipFile(target,'w',ZIP_DEFLATED) as z:
   for name,data in sorted(entries.items()):
    info=ZipInfo(name,(2026,10,17,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
  rows.append(dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),Story=story,Kind=kind,HeldOut=False,ExpectedEquationCount=4,ExpectedText=['share the same input series.']))
(cases/'equation-revisions.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Pinned',len(rows),'source-owned equation revision controls')
