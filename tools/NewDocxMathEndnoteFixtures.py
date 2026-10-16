"""Public continuous-endnote controls derived from pinned footnote combinations."""
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib,json

root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
sources=json.loads((cases/'math-footnotes.json').read_text(encoding='utf-8'))
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
types='{http://schemas.openxmlformats.org/package/2006/content-types}'
rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
E.register_namespace('w',w[1:-1]);manifest=[]
for case in sources:
    source=root/case['Input'];assert hashlib.sha256(source.read_bytes()).hexdigest().upper()==case['Sha256']
    with ZipFile(source) as archive:entries={name:archive.read(name) for name in archive.namelist()}
    document=E.fromstring(entries['word/document.xml'])
    for node in document.iter(w+'footnoteReference'):node.tag=w+'endnoteReference'
    for section in document.iter(w+'sectPr'):
        setting=section.find(w+'footnotePr')
        if setting is not None:
            assert section.find(w+'endnotePr') is None
            setting.tag=w+'endnotePr'
    entries['word/document.xml']=E.tostring(document,encoding='utf-8',xml_declaration=True)
    notes=E.fromstring(entries.pop('word/footnotes.xml'));notes.tag=w+'endnotes'
    for node in notes.iter():
        if node.tag==w+'footnote':node.tag=w+'endnote'
        elif node.tag==w+'footnoteRef':node.tag=w+'endnoteRef'
    if 'word/endnotes.xml' in entries:
        existing=E.fromstring(entries['word/endnotes.xml'])
        assert all(int(node.get(w+'id','0'))<=0 for node in existing)
        for node in existing:notes.insert(0,node)
    entries['word/endnotes.xml']=E.tostring(notes,encoding='utf-8',xml_declaration=True)
    relationships=E.fromstring(entries['word/_rels/document.xml.rels'])
    has_endnotes=any(node.get('Type')==rel+'endnotes' for node in relationships)
    for node in list(relationships):
        if node.get('Type')==rel+'footnotes':
            if has_endnotes:relationships.remove(node)
            else:node.set('Type',rel+'endnotes');node.set('Target','endnotes.xml')
    entries['word/_rels/document.xml.rels']=E.tostring(relationships,encoding='utf-8',xml_declaration=True)
    content_types=E.fromstring(entries['[Content_Types].xml'])
    has_endnotes=any(node.get('PartName')=='/word/endnotes.xml' for node in content_types)
    for node in list(content_types):
        if node.get('PartName')=='/word/footnotes.xml':
            if has_endnotes:content_types.remove(node)
            else:node.set('PartName','/word/endnotes.xml');node.set('ContentType','application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml')
    entries['[Content_Types].xml']=E.tostring(content_types,encoding='utf-8',xml_declaration=True)
    target=cases/(source.stem.replace('math-footnote-','math-endnote-')+'.docx')
    with ZipFile(target,'w',ZIP_DEFLATED) as archive:
        for name,data in sorted(entries.items()):
            info=ZipInfo(name,(2026,10,16,0,0,0));info.compress_type=ZIP_DEFLATED;archive.writestr(info,data)
    manifest.append(case|dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper()))
(cases/'math-endnotes.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
print('Pinned four continuous endnote/body/cell controls; default lower-Roman numbering retained.')
