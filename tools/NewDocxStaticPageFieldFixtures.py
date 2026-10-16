"""Pinned multi-page static stories combining PAGE/NUMPAGES with fractions."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import hashlib,json

root=Path(__file__).resolve().parent.parent
cases=root/'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
sources={row['Id']:row for row in json.loads((cases/'math-fields.json').read_text(encoding='utf-8'))}
w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
E.register_namespace('w',w[1:-1])
manifest=[]
for form in ['simple','packed']:
    for story in ['header','footer']:
        source=cases/('math-fields-'+form+'-'+story+'.docx')
        assert hashlib.sha256(source.read_bytes()).hexdigest().upper()==sources[source.stem]['Sha256']
        with ZipFile(source) as archive:entries={name:archive.read(name) for name in archive.namelist()}
        name='word/math-fields-'+story+'.xml'
        part=E.fromstring(entries[name])
        for field in part.iter(w+'fldSimple'):
            field.set(w+'instr',' PAGE ' if 'North' in field.get(w+'instr') else ' NUMPAGES ')
            field.find('.//'+w+'t').text='1' if field.get(w+'instr').strip()=='PAGE' else '12'
        for run in part.iter(w+'r'):
            if run.find(w+'instrText') is not None:
                instruction=run.find(w+'instrText');page='North' in instruction.text
                instruction.text=' PAGE ' if page else ' NUMPAGES '
                run.find(w+'t').text='1' if page else '12'
        entries[name]=E.tostring(part,encoding='utf-8',xml_declaration=True)
        doc=E.fromstring(entries['word/document.xml']);body=doc.find(w+'body');section=body.find(w+'sectPr')
        for number in range(2,13):
            p=E.Element(w+'p');E.SubElement(E.SubElement(p,w+'pPr'),w+'pageBreakBefore')
            run=E.SubElement(p,w+'r');props=E.SubElement(run,w+'rPr')
            E.SubElement(props,w+'rFonts',{w+'ascii':'Arial',w+'hAnsi':'Arial'});E.SubElement(props,w+'sz',{w+'val':'24'})
            E.SubElement(run,w+'t').text=f'Public page {number}.'
            body.insert(list(body).index(section),p)
        entries['word/document.xml']=E.tostring(doc,encoding='utf-8',xml_declaration=True)
        target=cases/('math-page-fields-'+form+'-'+story+'.docx')
        with ZipFile(target,'w',ZIP_DEFLATED) as archive:
            for name,data in sorted(entries.items()):
                info=ZipInfo(name,(2026,10,16,0,0,0));info.compress_type=ZIP_DEFLATED;archive.writestr(info,data)
        manifest.append(dict(Id=target.stem,Input=target.relative_to(root).as_posix(),Sha256=hashlib.sha256(target.read_bytes()).hexdigest().upper(),
                             HeldOut=False,ExpectedEquationCount=2,ExpectedText=['North','South','Start','middle','End.','Public page 12.']))
(cases/'math-page-fields.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf-8')
print('Pinned four twelve-page static-field controls; both sides of the 9/10 boundary are retained.')
