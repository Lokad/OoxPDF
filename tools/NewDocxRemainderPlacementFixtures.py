"""Bounded empty-base/operator composition across shared story containers."""
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib,json
root=Path(__file__).resolve().parent.parent;cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
E.register_namespace('w',w[1:-1]);E.register_namespace('m',m[1:-1]);rows=[]
for role,source,node in [('script-base','math-remaining-script-base-del','sSubSup'),('operator-limit','math-remaining-nary-stacked-upper-del','nary')]:
 with ZipFile(cases/(source+'.docx')) as z:
  fragment=E.tostring(next(E.fromstring(z.read('word/document.xml')).iter(m+node)))
 for story,part in [('body','word/document.xml'),('cell','word/document.xml'),('header','word/header1.xml'),('footer','word/footer1.xml')]:
  base=cases/('math-numbered-display-'+story+'.docx')
  with ZipFile(base) as z:entries={name:z.read(name) for name in z.namelist()}
  xml=E.fromstring(entries[part]);equation=next(xml.iter(m+'oMath'));equation.clear();equation.append(E.fromstring(fragment));entries[part]=E.tostring(xml,encoding='utf-8',xml_declaration=True);target=cases/('remainder-placement-'+role+'-'+story+'.docx')
  with ZipFile(target,'w',ZIP_DEFLATED) as z:
   for name,data in sorted(entries.items()):
    info=ZipInfo(name,(2026,10,17,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
  rows.append(dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),SourceSha256=hashlib.sha256(base.read_bytes()).hexdigest().upper(),FragmentSourceSha256=hashlib.sha256((cases/(source+'.docx')).read_bytes()).hexdigest().upper(),Role=role,Story=story,HeldOut=False,ExpectedEquationCount=4,ExpectedText=['share the same input series.']))
(cases/'remainder-placement.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8');print('Pinned eight shared-story remainder controls')
