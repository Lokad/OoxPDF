"""Write independent public multi-section records using Python's standard library."""
from pathlib import Path
from xml.sax.saxutils import escape
import hashlib, json, struct, zipfile, zlib

root = Path(__file__).resolve().parents[1]
destination = root / 'tests/Lokad.OoxPdf.Tests/Cases/SectionRecords'
destination.mkdir(exist_ok=True)
w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
relationships = 'http://schemas.openxmlformats.org/package/2006/relationships'
ns = (f'xmlns:w="{w}" xmlns:r="{office[:-1]}" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')

def run(text, properties=''):
    return f'<w:r>{properties}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'

def paragraph(text='', extra='', properties=''):
    return f'<w:p><w:pPr>{properties}</w:pPr>{run(text) if text else ""}{extra}</w:p>'

def math(text):
    return f'<m:oMath><m:r><m:rPr><m:nor/></m:rPr><m:t>{escape(text)}</m:t></m:r></m:oMath>'

def reference(number):
    return f'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="{number * 10}"/></w:r>'

def field(instruction):
    return f'<w:fldSimple w:instr=" {instruction} ">{run("0")}</w:fldSimple>'

def section(number, page_size, margin):
    width,height=page_size
    return (f'<w:sectPr><w:headerReference w:type="default" r:id="header{number}"/>'
            f'<w:footerReference w:type="default" r:id="footer{number}"/><w:type w:val="nextPage"/>'
            f'<w:pgSz w:w="{width}" w:h="{height}"/>'
            f'<w:pgMar w:top="{margin}" w:right="{margin}" w:bottom="{margin}" w:left="{margin}" w:header="360" w:footer="360"/>'
            '<w:footnotePr><w:numFmt w:val="decimal"/><w:numStart w:val="1"/><w:numRestart w:val="continuous"/></w:footnotePr></w:sectPr>')

def picture(width, height):
    x,y=width*12700,height*12700
    return (f'<w:r><w:drawing><wp:inline><wp:extent cx="{x}" cy="{y}"/><wp:docPr id="1" name="Public service marker"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            '<pic:nvPicPr><pic:cNvPr id="1" name="Public service marker"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="marker"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{x}" cy="{y}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')

def marker_png():
    def chunk(kind,data):
        return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data))
    pixels=b''.join(b'\0'+b''.join(bytes((30,120,190) if x<16 else (30,160,90)) for x in range(32)) for _ in range(16))
    return b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',32,16,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(pixels))+chunk(b'IEND',b'')

def table(number, label):
    cells=[paragraph('Reading '+label,extra=math('r+'+str(number))),paragraph('Table-owned verification',extra=reference(number))]
    return ('<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid><w:tr>'+
            ''.join('<w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr>'+cell+'</w:tc>' for cell in cells)+'</w:tr></w:tbl>')

records=[
    dict(Id='field-brief',Title='Field service brief',Pages=['Intake','Review'],Size=(12240,15840),Margin=1080,PicturePage=2,PictureSize=(48,24),HeldOut=False,
         Prose=['Record the initial condition before starting work. All readings in this public service example are illustrative.',
                'Retain the final observations, the verification note and the signed handover record.']),
    dict(Id='visit-record',Title='Site visit record',Pages=['Preparation','Observation','Handover'],Size=(11906,16838),Margin=920,PicturePage=1,PictureSize=(36,18),HeldOut=True,
         Prose=['Review the access instructions and identify the equipment before entering the work area.',
                'Keep each observation with its source reading. A longer statement here checks normal prose wrapping without changing note or section ownership.',
                'Hand over the completed visit record and retain the supporting observations for the next planned inspection.'])]

expected=[]
for record in records:
    body=[];page_contracts=[];parts={};notes=[];overrides=[]
    for number,label in enumerate(record['Pages'],1):
        last=number==len(record['Pages']);header=record['Title']+' / '+label;footer=label+' ledger / Page '
        title=record['Title'] if number==1 else label+' findings'
        bookmark=('<w:bookmarkStart w:id="7" w:name="FinalDetails"/>' if last else '')
        bookmark_end='<w:bookmarkEnd w:id="7"/>' if last else ''
        body.append(paragraph(title,extra=bookmark+bookmark_end,properties='<w:spacing w:after="160"/>'))
        body.append(paragraph(record['Prose'][number-1]))
        text=[title,record['Prose'][number-1]]
        if number==1:
            body.append(paragraph('Navigate to the final section: ',extra='<w:hyperlink w:anchor="FinalDetails">'+run('Final details')+'</w:hyperlink>'))
            body.append(paragraph('Public supporting information: ',extra='<w:hyperlink r:id="publicLink">'+run('Service reference')+'</w:hyperlink>'))
            text.extend(['Navigate to the final section: Final details','Public supporting information: Service reference'])
        if number==2:
            body.append(table(number,label));text.extend(['Reading '+label,'Table-owned verification'])
        else:
            body.append(paragraph('Source-owned verification',extra=reference(number)));text.append('Source-owned verification')
        if number==record['PicturePage']:
            body.append(paragraph(extra=picture(*record['PictureSize'])));body.append(paragraph('Service marker retained.'));text.append('Service marker retained.')
        closing='End of '+label.lower()+'.';text.append(closing)
        body.append(paragraph(closing,properties='' if last else section(number,record['Size'],record['Margin'])))
        header_math='h+'+str(number)
        parts[f'word/header{number}.xml']=f'<w:hdr {ns}>'+paragraph(header+' ',extra=math(header_math))+'</w:hdr>'
        parts[f'word/footer{number}.xml']=f'<w:ftr {ns}>'+paragraph(footer,extra=field('PAGE')+run(' of ')+field('NUMPAGES'))+'</w:ftr>'
        for kind in ['header','footer']:
            overrides.append(f'<Override PartName="/word/{kind}{number}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>')
        note_text=label+' note: the corresponding source record is retained.';note_math='n+'+str(number)
        note_marker='<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
        notes.append(f'<w:footnote w:id="{number*10}">'+paragraph(extra=note_marker+run(' '+note_text+' ')+math(note_math))+'</w:footnote>')
        page_contracts.append(dict(Page=number,Header=header,Footer=footer+str(number)+' of '+str(len(record['Pages'])),Text=text,Note=note_text,Mark=str(number),Math=[header_math,note_math]+(['r+'+str(number)] if number==2 else []),Picture=number==record['PicturePage']))
    body.append(section(len(record['Pages']),record['Size'],record['Margin']))
    parts['word/document.xml']=f'<w:document {ns}><w:body>'+''.join(body)+'</w:body></w:document>'
    parts['word/footnotes.xml']=f'<w:footnotes {ns}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'+''.join(notes)+'</w:footnotes>'
    parts['word/styles.xml']=f'<w:styles xmlns:w="{w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="100" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>'
    parts['word/settings.xml']=f'<w:settings xmlns:w="{w}"><w:footnotePr><w:numFmt w:val="decimal"/><w:numStart w:val="1"/><w:numRestart w:val="continuous"/></w:footnotePr></w:settings>'
    parts['word/_rels/document.xml.rels']=f'<Relationships xmlns="{relationships}">'+''.join(f'<Relationship Id="{kind}" Type="{office}{kind}" Target="{kind}.xml"/>' for kind in ['styles','settings','footnotes'])+''.join(f'<Relationship Id="{kind}{number}" Type="{office}{kind}" Target="{kind}{number}.xml"/>' for number in range(1,len(record['Pages'])+1) for kind in ['header','footer'])+f'<Relationship Id="marker" Type="{office}image" Target="media/marker.png"/><Relationship Id="publicLink" Type="{office}hyperlink" Target="https://example.invalid/public-service-record" TargetMode="External"/></Relationships>'
    parts['_rels/.rels']=f'<Relationships xmlns="{relationships}"><Relationship Id="document" Type="{office}officeDocument" Target="word/document.xml"/></Relationships>'
    types=['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>','<Default Extension="xml" ContentType="application/xml"/>','<Default Extension="png" ContentType="image/png"/>']
    for part,kind in [('document','document.main'),('styles','styles'),('settings','settings'),('footnotes','footnotes')]:
        types.append(f'<Override PartName="/word/{part}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>')
    parts['[Content_Types].xml']=f'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+''.join(types+overrides)+'</Types>'
    parts['word/media/marker.png']=marker_png()
    path=destination/(record['Id']+'.docx')
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as package:
        for part,data in sorted(parts.items()):
            entry=zipfile.ZipInfo(part,(2026,11,15,0,0,0));entry.compress_type=zipfile.ZIP_DEFLATED
            package.writestr(entry,data.encode('utf-8') if isinstance(data,str) else data)
    expected.append(dict(Id=record['Id'],HeldOut=record['HeldOut'],InputSha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(),PageSize=[v/20 for v in record['Size']],PictureSize=record['PictureSize'],Pages=page_contracts))
(destination/'expected.json').write_text(json.dumps(expected,indent=2)+'\n',encoding='utf-8')
print('Generated two independent section/note/static-story records (two and three authored pages).')
