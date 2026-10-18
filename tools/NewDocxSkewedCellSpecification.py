from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib,json
root=Path(__file__).resolve().parent.parent;cases=root/'tests/Lokad.OoxPdf.Tests/Cases'
parent=json.loads((cases/'Composition/structured-cell-specification.json').read_text(encoding='utf-8'))[0];source=root/parent['Input'];assert hashlib.sha256(source.read_bytes()).hexdigest().upper()==parent['Sha256']
with ZipFile(source) as archive:entries={n:archive.read(n) for n in archive.namelist()}
m='{http://schemas.openxmlformats.org/officeDocument/2006/math}';document=E.fromstring(entries['word/document.xml']);matrices=list(document.iter(m+'m'));assert len(matrices)==3
for matrix in matrices:
 fraction=matrix.find('.//'+m+'f');pr=fraction.find(m+'fPr');assert pr is not None
 typ=pr.find(m+'type')
 if typ is None:typ=E.Element(m+'type');pr.insert(0,typ)
 typ.set(m+'val','skw')
entries['word/document.xml']=E.tostring(document,encoding='utf-8',xml_declaration=True);name='composed-skewed-cell-specification';output=cases/'Composition';output.mkdir(exist_ok=True);destination=output/(name+'.docx')
with ZipFile(destination,'w',ZIP_DEFLATED) as archive:
 for path,data in sorted(entries.items()):
  info=ZipInfo(path,(2026,1,1,0,0,0));info.compress_type=ZIP_DEFLATED;archive.writestr(info,data)
row=dict(parent,Id=name,Input='tests/Lokad.OoxPdf.Tests/Cases/Composition/'+destination.name,Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),SourceSha256=parent['Sha256'],Features=parent['Features']+['skewed-fractions'])
(output/'skewed-cell-specification.json').write_text(json.dumps([row],indent=2)+'\n',encoding='utf-8');print('One complete source changed only three matrix fraction types; fixed/held-out inputs unchanged.')
