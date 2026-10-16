"""Bordered independent math/note cells, with nominal-size and script variants."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json
root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
source=cases/'math-footnote-after-cell.docx'
pin=next(x for x in json.loads((cases/'math-footnotes.json').read_text()) if x['Id']==source.stem)
assert hashlib.sha256(source.read_bytes()).hexdigest().upper()==pin['Sha256']
with ZipFile(source) as z:base={n:z.read(n) for n in z.namelist()}
rows=[]
for size in [11,12]:
    for alignment in ['baseline','superscript']:
        data=base.copy();doc=E.fromstring(data['word/document.xml'])
        table=doc.find('.//'+w+'tbl');p=table.findall('.//'+w+'tc')[1].find(w+'p')
        equation=p.find(m+'oMath')
        marker=next(run for run in p.findall(w+'r') if run.find(w+'footnoteReference') is not None)
        for node in list(p):
            if node.tag!=w+'pPr':p.remove(node)
        p.append(equation);p.append(marker)
        if alignment=='baseline':
            for props in marker.findall(w+'rPr'):
                for va in props.findall(w+'vertAlign'):props.remove(va)
        for node in doc.iter():
            if node.tag in [w+'sz',w+'szCs']:node.set(w+'val',str(size*2))
        borders=E.SubElement(table.find(w+'tblPr'),w+'tblBorders')
        for edge in ['top','bottom','left','right','insideH','insideV']:
            E.SubElement(borders,w+edge,{w+'val':'single',w+'sz':'4',w+'color':'000000'})
        data['word/document.xml']=E.tostring(doc,encoding='utf-8',xml_declaration=True)
        path=cases/f'math-note-border-{size}-{alignment}.docx'
        with ZipFile(path,'w',ZIP_DEFLATED) as z:
            for name,content in sorted(data.items()):
                info=ZipInfo(name,(2026,10,16,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,content)
        rows.append(dict(Id=path.stem,Input=path.relative_to(root).as_posix(),Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(),HeldOut=False,ExpectedEquationCount=1,ExpectedText=['Observed series.','Public calculation note.']))
(cases/'math-note-borders.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8')
print('Pinned four bordered math/note controls')
