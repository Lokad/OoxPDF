"""Add a public comment lane to the independent complete-note controls."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
from copy import deepcopy
import json, hashlib
root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
comment_rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
with ZipFile(cases/'math-footnote-after-cell.docx') as z: seed={n:z.read(n) for n in z.namelist()}
base={x['Id']:x for x in json.loads((cases/'complete-notes.json').read_text(encoding='utf-8'))}
rows=[]
for kind in ['body','cell']:
    source=cases/f'complete-note-{kind}.docx'
    assert hashlib.sha256(source.read_bytes()).hexdigest().upper()==base[source.stem]['Sha256']
    with ZipFile(source) as z:data={n:z.read(n) for n in z.namelist()}
    doc=E.fromstring(data['word/document.xml'])
    # Word's review exporter centers custom short sheets in a Letter print frame.
    # Use a standard sheet and the same 300pt body height, keeping all ink in-page.
    section=doc.find('.//'+w+'sectPr')
    section.find(w+'pgSz').set(w+'h','15840')
    section.find(w+'pgMar').set(w+'bottom','9120')
    p=doc.find(w+'body').find(w+'p')
    p.insert(1,E.Element(w+'commentRangeStart',{w+'id':'0'}))
    E.SubElement(p,w+'commentRangeEnd',{w+'id':'0'})
    E.SubElement(E.SubElement(p,w+'r'),w+'commentReference',{w+'id':'0'})
    data['word/document.xml']=E.tostring(doc,encoding='utf-8',xml_declaration=True)
    data['word/comments.xml']=seed['word/comments.xml']
    rels=E.fromstring(data['word/_rels/document.xml.rels'])
    rels.append(deepcopy(next(x for x in E.fromstring(seed['word/_rels/document.xml.rels']) if x.get('Type')==comment_rel)))
    data['word/_rels/document.xml.rels']=E.tostring(rels,encoding='utf-8',xml_declaration=True)
    types=E.fromstring(data['[Content_Types].xml'])
    types.append(deepcopy(next(x for x in E.fromstring(seed['[Content_Types].xml']) if x.get('PartName')=='/word/comments.xml')))
    data['[Content_Types].xml']=E.tostring(types,encoding='utf-8',xml_declaration=True)
    path=cases/f'scaled-note-{kind}.docx'
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for name,content in sorted(data.items()):
            info=ZipInfo(name,(2026,10,16,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,content)
    rows.append(dict(Id=path.stem,Input=path.relative_to(root).as_posix(),Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=0,ExpectedText=['Section 6 line 3.','Public calculation note.']))
(cases/'scaled-notes.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Pinned two scaled complete-note controls')
