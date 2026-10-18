from pathlib import Path
from copy import deepcopy
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib,json
root=Path(__file__).resolve().parent.parent;source=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/radical-matrix-inline.docx'
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == 'A6E14611378485FF726EB85BE4664FAE30F4281EF777A3472BAEA1EC50D2F17D'
with ZipFile(source) as a:entries={n:a.read(n) for n in a.namelist()}
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';m='{http://schemas.openxmlformats.org/officeDocument/2006/math}';doc=E.fromstring(entries['word/document.xml']);rad=doc.find('.//'+m+'rad');template=deepcopy(rad.find('.//'+m+'r'));template.find(m+'t').text='A'
mp=template.find(m+'rPr')
if mp is None:mp=E.Element(m+'rPr');template.insert(0,mp)
mp[:]=[E.Element(m+'nor',{m+'val':'1'})]
wp=template.find(w+'rPr')
if wp is None:wp=E.Element(w+'rPr');template.insert(1,wp)
wp[:]=[E.Element(w+'rFonts',{w+'ascii':'Cambria Math',w+'hAnsi':'Cambria Math'}),E.Element(w+'sz',{w+'val':'400'})]
rad.find(m+'e')[:]=[template];pr=rad.find(m+'radPr')
if pr is None:pr=E.Element(m+'radPr');rad.insert(0,pr)
ctrl=pr.find(m+'ctrlPr')
if ctrl is not None:pr.remove(ctrl)
cp=E.SubElement(E.SubElement(pr,m+'ctrlPr'),w+'rPr');E.SubElement(cp,w+'rFonts',{w+'ascii':'Cambria Math',w+'hAnsi':'Cambria Math'});E.SubElement(cp,w+'sz',{w+'val':'24'})
entries['word/document.xml']=E.tostring(doc,encoding='utf-8',xml_declaration=True);name='radical-assembly-child200';output=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes';output.mkdir(exist_ok=True);dest=output/(name+'.docx')
with ZipFile(dest,'w',ZIP_DEFLATED) as a:
 for p,data in sorted(entries.items()):
  info=ZipInfo(p,(2026,1,1,0,0,0));info.compress_type=ZIP_DEFLATED;a.writestr(info,data)
row=dict(Id=name,Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/'+dest.name,Sha256=hashlib.sha256(dest.read_bytes()).hexdigest().upper(),SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=1,ExpectedText=['Following prose.'],ControlSize=12,ChildSize=200)
(output/'radical-assembly.json').write_text(json.dumps([row],indent=2)+'\n',encoding='utf-8');print('One authored-control/large-child assembly probe generated; public sources unchanged.')
