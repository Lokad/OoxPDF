"""Mechanically add cached fields to the held-out composed report, preserving text."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib, json

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/Composition'
source = cases / 'composed-annotations-report.docx'
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == 'CD239F3752E1F34C801530C0E93FA1501E0A4D6A823828E6A18463B4F0CD9598'
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
xml = '{http://www.w3.org/XML/1998/namespace}'
for prefix, namespace in [('w',w),('m',m)]: E.register_namespace(prefix,namespace[1:-1])
with ZipFile(source) as archive:
    entries = {name:archive.read(name) for name in archive.namelist()}
document = E.fromstring(entries['word/document.xml'])
visible = ''.join(node.text or '' for node in document.iter(w+'t'))
equations = [E.tostring(node) for node in document.iter(m+'oMath')]

def replace_word(scope, word, transform, last=False):
    parents = {child:parent for parent in document.iter() for child in parent}
    matches = [(node, (node.text or '').find(word)) for node in scope.iter(w+'t') if word in (node.text or '')]
    node, offset = matches[-1 if last else 0]
    run = parents[node]; parent = parents[run]
    assert run.tag == w+'r' and all(child.tag in [w+'rPr',w+'t'] for child in run)
    text = node.text
    def piece(value):
        clone = deepcopy(run); item = clone.find(w+'t'); item.text = value; item.set(xml+'space','preserve')
        return clone
    replacement = []
    if offset: replacement.append(piece(text[:offset]))
    replacement.extend(transform(piece(word)))
    if offset + len(word) < len(text): replacement.append(piece(text[offset+len(word):]))
    index = list(parent).index(run); parent.remove(run)
    for child in reversed(replacement): parent.insert(index,child)

def bookmark(name, number):
    return lambda run: [E.Element(w+'bookmarkStart',{w+'id':str(number),w+'name':name}),run,
                        E.Element(w+'bookmarkEnd',{w+'id':str(number)})]

paragraphs = list(document.iter(w+'p'))
summary = next(p for p in paragraphs if ''.join(t.text or '' for t in p.iter(w+'t')).startswith('Summary:'))
mixed = next(p for p in paragraphs if ''.join(t.text or '' for t in p.iter(w+'t')).startswith('The trend marker'))
replace_word(summary,'the',bookmark('CacheArticle',34580))
replace_word(mixed,'calculation',bookmark('CacheCalculation',34581))

def packed(run):
    text = run.find(w+'t'); run.remove(text)
    E.SubElement(run,w+'fldChar',{w+'fldCharType':'begin'})
    E.SubElement(run,w+'instrText',{xml+'space':'preserve'}).text = ' REF CacheArticle '
    E.SubElement(run,w+'fldChar',{w+'fldCharType':'separate'});run.append(text)
    E.SubElement(run,w+'fldChar',{w+'fldCharType':'end'})
    return [run]

def simple(run):
    field = E.Element(w+'fldSimple',{w+'instr':' REF CacheCalculation '});field.append(run)
    return [field]

replace_word(mixed,'the',packed)
replace_word(mixed,'calculation',simple,last=True)
assert ''.join(node.text or '' for node in document.iter(w+'t')) == visible
assert [E.tostring(node) for node in document.iter(m+'oMath')] == equations
entries['word/document.xml'] = E.tostring(document,encoding='utf-8',xml_declaration=True)
output = cases / 'composed-fields-report.docx'
with ZipFile(output,'w',ZIP_DEFLATED) as archive:
    for name,data in sorted(entries.items()):
        info=ZipInfo(name,(2026,10,16,0,0,0));info.compress_type=ZIP_DEFLATED;archive.writestr(info,data)
baseline = json.loads((cases/'composed-annotations-report.json').read_text(encoding='utf-8'))[0]
manifest = [baseline | dict(Id=output.stem,Input=output.relative_to(root).as_posix(),
    Sha256=hashlib.sha256(output.read_bytes()).hexdigest().upper(),
    ExpectedText=[text.removesuffix('.') if text == 'See calculation.' else text for text in baseline['ExpectedText']],
    Features=baseline['Features']+['cached field results beside math/image/comment and inside an internal hyperlink'])]
(cases/'composed-fields-report.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
print('Pinned one held-out cached-field report; visible prose, equations and other package parts preserved.')
