"""Visible script/operator remainders after deleting a base or a limit."""
from pathlib import Path
from zipfile import ZipFile,ZipInfo,ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib,json
root=Path(__file__).resolve().parent.parent;cases=root/'tests/Lokad.OoxPdf.Tests/Cases';output=cases/'CompositionProbes'
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}';m='{http://schemas.openxmlformats.org/officeDocument/2006/math}'
E.register_namespace('w',w[1:-1]);E.register_namespace('m',m[1:-1]);rows=[]
roles=[('script-base','docx-omml-isolated-inline-body-script','e',921),('nary-stacked-lower','docx-omml-nary-display-sum-undOvr','sub',921),('nary-side-lower','docx-omml-nary-default-inline-sum-setting-side','sub',921),('nary-stacked-upper','docx-omml-nary-display-sum-undOvr','sup',922),('nary-side-upper','docx-omml-nary-default-inline-sum-setting-side','sup',922)]
for role,file,argument,identifier in roles:
 source=cases/(file+'.docx')
 for kind in (['ins','del','empty'] if identifier==921 else ['del','empty']):
  with ZipFile(source) as z:entries={name:z.read(name) for name in z.namelist()}
  xml=E.fromstring(entries['word/document.xml']);eq=next(xml.iter(m+'oMath'));structure=eq.find(m+('sSubSup' if role=='script-base' else 'nary'));owner=structure.find(m+argument);run=owner.find(m+'r');owner.remove(run)
  if kind!='empty':E.SubElement(owner,w+kind,{w+'id':str(identifier),w+'author':'Math Argument Reviewer',w+'date':'2026-10-17T00:00:00Z'}).append(run)
  entries['word/document.xml']=E.tostring(xml,encoding='utf-8',xml_declaration=True);target=output/('math-remaining-'+role+'-'+kind+'.docx')
  with ZipFile(target,'w',ZIP_DEFLATED) as z:
   for name,data in sorted(entries.items()):
    info=ZipInfo(name,(2026,10,17,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,data)
  rows.append(dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),Role=role,Kind=kind,HeldOut=False,ExpectedEquationCount=1,ExpectedText=[]))
(output/'math-remaining.json').write_text(json.dumps(rows,indent=2)+'\n',encoding='utf-8');print('Pinned thirteen base/limit controls')

# Blank scripts retain authored slots and an implicit metric owner. Literal
# deletion and empty insertion challenge source ownership without new geometry.
source = output / 'math-control-revision-script.docx'
with ZipFile(source) as z:
    original = {name: z.read(name) for name in z.namelist()}
empty_rows = []
for role, kind in [('sub', 'empty'), ('sup', 'empty'), ('combined', 'empty'),
                   ('combined', 'literal-del'), ('combined', 'empty-ins')]:
    entries = dict(original)
    xml = E.fromstring(entries['word/document.xml'])
    equation = next(xml.iter(m + 'oMath'))
    wrapper = equation.find(w + 'ins')
    script = wrapper.find(m + 'sSubSup')
    wrapper.remove(script)
    equation.remove(wrapper)
    if role == 'sub':
        script.tag = m + 'sSub'
        script.remove(script.find(m + 'sup'))
    if role == 'sup':
        script.tag = m + 'sSup'
        script.remove(script.find(m + 'sub'))
    for index, parent in enumerate(list(script.iter())):
        for run in list(parent.findall(m + 'r')):
            parent.remove(run)
            if kind == 'literal-del':
                E.SubElement(parent, w + 'del', {w + 'id': str(1101 + index), w + 'author': 'Script Frame Reviewer',
                    w + 'date': '2026-10-17T00:00:00Z'}).append(run)
    if kind == 'empty-ins':
        E.SubElement(equation, w + 'ins', {w + 'id': '1100', w + 'author': 'Script Frame Reviewer',
            w + 'date': '2026-10-17T00:00:00Z'}).append(script)
    else:
        equation.append(script)
    entries['word/document.xml'] = E.tostring(xml, encoding='utf-8', xml_declaration=True)
    target = output / ('empty-script-' + role + '-' + kind + '.docx')
    with ZipFile(target, 'w', ZIP_DEFLATED) as z:
        for name, data in sorted(entries.items()):
            info = ZipInfo(name, (2026, 10, 17, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            z.writestr(info, data)
    empty_rows.append(dict(Id=target.stem, Input=target.relative_to(root).as_posix(),
                           Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                           SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
                           Role=role, Kind=kind, HeldOut=False, ExpectedEquationCount=4,
                           ExpectedText=['share the same input series.'], ProsePrefixes=['Summary:', 'governs']))
(output / 'empty-scripts.json').write_text(json.dumps(empty_rows, indent=2) + '\n', encoding='utf-8')
print('Pinned five empty-script frame controls')
