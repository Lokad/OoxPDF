"""Regenerate independent row-floor controls, preserving source/picture identities."""
import copy,hashlib,json,zipfile
from io import BytesIO
from pathlib import Path
import xml.etree.ElementTree as E
W=Path(__file__).resolve().parent.parent;P=W/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes';D=P
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';a='{http://schemas.openxmlformats.org/drawingml/2006/main}';r='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}';rel='{http://schemas.openxmlformats.org/package/2006/relationships}';ct='{http://schemas.openxmlformats.org/package/2006/content-types}'
source=P/'merged-span-page-boundary.docx';image_source=P/'justified-linked-image-cell.docx';sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest().upper()
with zipfile.ZipFile(source) as z:infos=z.infolist();original={i.filename:z.read(i.filename) for i in infos}
with zipfile.ZipFile(image_source) as z:
 drawing=copy.deepcopy(next(E.fromstring(z.read('word/document.xml')).iter(w+'drawing')));picture=z.read('word/media/math-row-trend.png')
 next(drawing.iter(a+'blip')).set(r+'embed','rIdFloorImage')
expected=json.loads((D/'merge-source-row-floors.json').read_text())
assert all(sha(source)==row['SourceSha256'] for row in expected)
assert sha(image_source)==next(row['PictureSourceSha256'] for row in expected if row['Variant']=='picture-wrap')
manifest=[]
variants=['short','wrap','picture-wrap','wrap-mark','wrap-run','short-mark','paragraphs','paragraphs-mark']
for variant in variants:
 parts=original.copy();tree=E.fromstring(parts['word/document.xml']);body=tree.find(w+'body')
 for p in list(body.findall(w+'p')):
  if any(t.text=='Public page-boundary spacer.' for t in p.iter(w+'t')):body.remove(p)
 cell=body.find(w+'tbl').find(w+'tr').find(w+'tc')
 for p in list(cell.findall(w+'p')):cell.remove(p)
 strings=['Public first merged paragraph.','Public second merged paragraph.'] if variant.startswith('paragraphs') else ['Public merged prose.' if variant.startswith('short') else 'Public merged prose keeps its source ordered across a shared row frame. A second sentence continues through the same merged source.']
 for index,text in enumerate(strings):
  p=E.SubElement(cell,w+'p')
  if 'mark' in variant and index==len(strings)-1:
   E.SubElement(E.SubElement(E.SubElement(p,w+'pPr'),w+'rPr'),w+'sz',{w+'val':'40'})
  run=E.SubElement(p,w+'r');E.SubElement(E.SubElement(run,w+'rPr'),w+'sz',{w+'val':'40' if variant=='wrap-run' else '22'})
  if variant=='picture-wrap':
   split=text.index(' A second');E.SubElement(run,w+'t',{'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text=text[:split]+' '
   E.SubElement(p,w+'r').append(copy.deepcopy(drawing));E.SubElement(E.SubElement(p,w+'r'),w+'t',{'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text=text[split:]
  else:E.SubElement(run,w+'t').text=text
 parts['word/document.xml']=E.tostring(tree,encoding='utf-8',xml_declaration=True)
 if variant=='picture-wrap':
  relationships=E.fromstring(parts['word/_rels/document.xml.rels']);E.SubElement(relationships,rel+'Relationship',{'Id':'rIdFloorImage','Type':r[1:-1]+'/image','Target':'media/row-floor-trend.png'});parts['word/_rels/document.xml.rels']=E.tostring(relationships,encoding='utf-8',xml_declaration=True)
  types=E.fromstring(parts['[Content_Types].xml']);assert not any(t.get('Extension')=='png' for t in types);E.SubElement(types,ct+'Default',{'Extension':'png','ContentType':'image/png'});parts['[Content_Types].xml']=E.tostring(types,encoding='utf-8',xml_declaration=True)
 target=D/('merge-row-floor-'+variant+'.docx')
 buffer=BytesIO()
 with zipfile.ZipFile(buffer,'w',zipfile.ZIP_DEFLATED) as z:
  for info in infos:z.writestr(copy.copy(info),parts[info.filename])
  if variant=='picture-wrap':
   info=zipfile.ZipInfo('word/media/row-floor-trend.png',(2026,10,19,0,0,0));info.compress_type=zipfile.ZIP_DEFLATED;z.writestr(info,picture)
 data=buffer.getvalue();assert hashlib.sha256(data).hexdigest().upper()==next(row['Sha256'] for row in expected if row['Id']==target.stem)
 target.write_bytes(data)
 manifest.append(dict(Id=target.stem,Input=target.relative_to(W).as_posix(),Sha256=sha(target),SourceSha256=sha(source),HeldOut=True,ExpectedEquationCount=0,ExpectedInlineImageCount=1 if variant=='picture-wrap' else 0,ExpectedText=['Before public math table.','After public math table.','Side row 1.','Side row 2.']+(['Public first merged paragraph.','Public second merged paragraph.'] if variant.startswith('paragraphs') else ['Public merged prose']),Variant=variant,Change='Independent source-row-floor probe: short/wrapped prose, image, source run/paragraph mark, or multiple paragraphs; no equation or page spacer.'))
expected=json.loads((D/'merge-source-row-floors.json').read_text())
for row in manifest:
 case=next(c for c in expected if c['Id']==row['Id']);assert row['Sha256']==case['Sha256'] and row['SourceSha256']==case['SourceSha256']
print('Eight merge source-row-floor controls regenerated byte-identically.')
