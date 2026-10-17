"""Tracked literal math runs inside a supported fraction, across containers."""
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import json,hashlib
root=Path(__file__).resolve().parent.parent;cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
E.register_namespace('w',w[1:-1]);E.register_namespace('m',m[1:-1]);rows=[]
for story,part in [('body','word/document.xml'),('cell','word/document.xml'),('header','word/header1.xml'),('footer','word/footer1.xml')]:
 source=cases/('math-numbered-display-'+story+'.docx')
 for kind in ['ins','del']:
  with ZipFile(source) as z:entries={n:z.read(n) for n in z.namelist()}
  xml=E.fromstring(entries[part]);eq=next(xml.iter(m+'oMath'));num=eq.find(m+'f').find(m+'num');run=num.find(m+'r');i=list(num).index(run);num.remove(run)
  wrapper=E.Element(w+kind,{w+'id':'891',w+'author':'Math Run Reviewer',w+'date':'2026-10-17T00:00:00Z'});wrapper.append(run);num.insert(i,wrapper);entries[part]=E.tostring(xml,encoding='utf-8',xml_declaration=True)
  target=cases/('math-run-revision-'+story+'-'+kind+'.docx')
  with ZipFile(target,'w',ZIP_DEFLATED) as z:
   for name,data in sorted(entries.items()):
    info=ZipInfo(name,(2026,10,17,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
  rows.append(dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),Story=story,Kind=kind,HeldOut=False,ExpectedEquationCount=4,ExpectedText=['share the same input series.']))
(cases/'math-run-revisions.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8');print('Pinned',len(rows),'inner math revision controls')
