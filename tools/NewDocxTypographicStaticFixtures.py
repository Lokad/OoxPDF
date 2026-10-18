"""Exercise declared static minimums with a face requesting typographic metrics."""
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import json,hashlib
from argparse import ArgumentParser
root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
parser=ArgumentParser();parser.add_argument('--output-directory',type=Path,default=cases)
out=parser.parse_args().output_directory.resolve();out.mkdir(parents=True,exist_ok=True)
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';rows=[]
for kind in ['header','footer']:
 source=cases/('minimum-static-'+kind+'.docx')
 with ZipFile(source) as z:parts={n:z.read(n) for n in z.namelist()}
 part='word/'+kind+'MinimumPlain.xml';story=E.fromstring(parts[part])
 for fonts in story.iter(w+'rFonts'):
  fonts.set(w+'ascii','Cambria Math');fonts.set(w+'hAnsi','Cambria Math')
 parts[part]=E.tostring(story,encoding='utf-8',xml_declaration=True);name='minimum-static-typographic-'+kind;path=out/(name+'.docx')
 with ZipFile(path,'w',ZIP_DEFLATED) as z:
  for n,data in sorted(parts.items()):
   info=ZipInfo(n,(2026,1,1,0,0,0));info.create_system=0;info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
 rows.append(dict(Id=name,Input=path.relative_to(root).as_posix(),Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(),SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=0,ExpectedInlineImageCount=0,ExpectedText=['Leading paragraph.','Following prose.']+['Public review '+kind+' line '+str(i)+'.' for i in range(1,4)],Font='Cambria Math',MinimumLineHeight=20,Part=part))
(out/'typographic-static-frames.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Two typographic-metric static controls pinned; only the declared face changes.')
