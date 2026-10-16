"""Independent review-pane controls across body margins and paper shapes."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import json, hashlib
root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
with ZipFile(cases/'math-comment-prose-body.docx') as z:base={n:z.read(n) for n in z.namelist()}
rows=[]
variants=[('letter-36-36',12240,15840,720,720),('letter-72-72',12240,15840,1440,1440),('letter-144-144',12240,15840,2880,2880),('letter-36-456',12240,15840,720,9120),('letter-144-36',12240,15840,2880,720),('letter-36-144',12240,15840,720,2880),('a4-72-72',11906,16838,1440,1440),('landscape-72-72',15840,12240,1440,1440)]
for name,width,height,top,bottom in variants:
    data=base.copy();doc=E.fromstring(data['word/document.xml'])
    section=doc.find('.//'+w+'sectPr');size=section.find(w+'pgSz')
    size.set(w+'w',str(width));size.set(w+'h',str(height));size.set(w+'orient','landscape' if width>height else 'portrait')
    margins=section.find(w+'pgMar');margins.set(w+'top',str(top));margins.set(w+'bottom',str(bottom))
    data['word/document.xml']=E.tostring(doc,encoding='utf-8',xml_declaration=True)
    path=cases/('comment-pane-'+name+'.docx')
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for n,b in sorted(data.items()):
            i=ZipInfo(n,(2026,10,16,0,0,0));i.compress_type=ZIP_DEFLATED;z.writestr(i,b)
    rows.append(dict(Id=path.stem,Input=path.relative_to(root).as_posix(),Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=2,ExpectedText=[]))
path=cases.parent/'docx-markup-margin-mirrored.docx'
rows.append(dict(Id='comment-pane-mirrored',Input=path.relative_to(root).as_posix(),Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=0,ExpectedText=[]))
(cases/'comment-panes.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Pinned eight paper/margin controls plus the existing mirrored reference')
