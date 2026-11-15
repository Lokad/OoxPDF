"""Generate independent public DOCX list compositions using only Python's stdlib."""
from pathlib import Path
import hashlib, json, struct, zlib, zipfile

root = Path(__file__).resolve().parents[1]
destination = root / 'tests/Lokad.OoxPdf.Tests/Cases/HierarchicalLists'
destination.mkdir(exist_ok=True)
w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'

def paragraph(text, level=None, number=1, extra=''):
    properties = '' if level is None else f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{number}"/></w:numPr>'
    return f'<w:p><w:pPr>{properties}</w:pPr><w:r><w:t>{text}</w:t></w:r>{extra}</w:p>'

def level(index, form, start, pattern, legal=''):
    return (f'<w:lvl w:ilvl="{index}"><w:start w:val="{start}"/><w:numFmt w:val="{form}"/>'
            f'<w:lvlText w:val="{pattern}"/><w:suff w:val="tab"/>{legal}'
            f'<w:pPr><w:tabs><w:tab w:val="num" w:pos="{2400+index*240}"/></w:tabs>'
            f'<w:ind w:left="{2400+index*240}" w:hanging="2160"/></w:pPr></w:lvl>')

def png():
    def chunk(kind, data):
        return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data))
    pixels = b''.join(b'\0'+b''.join(bytes((20,150,50) if x<12 else (30,90,220)) for x in range(24)) for _ in range(12))
    return b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',24,12,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(pixels))+chunk(b'IEND',b'')

picture = '''<w:r><w:drawing><wp:inline><wp:extent cx="304800" cy="152400"/><wp:docPr id="1" name="Public status marker"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Public status marker"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="304800" cy="152400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'''
math = '<m:oMath><m:r><m:rPr><m:nor/></m:rPr><m:t>x+1</m:t></m:r></m:oMath>'
link = '<w:hyperlink r:id="link"><w:r><w:t> Public navigation record</w:t></w:r></w:hyperlink>'

levels = level(0,'upperRoman',4,'%1.')+level(1,'lowerLetter',26,'%1.%2)')+level(2,'decimalZero',9,'%1.%2.%3)')
numbering = f'<w:abstractNum w:abstractNumId="7">{levels}</w:abstractNum><w:abstractNum w:abstractNumId="8">{level(0,"upperLetter",52,"(%1)")}{level(1,"lowerRoman",8,"%1-%2")}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="7"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="8"/></w:num>'
body = paragraph('Mixed-format project checklist')+paragraph('Scope and assumptions',0,extra=math+picture+link)
body += paragraph('Validate input ownership and preserve cancellation throughout conversion.',1)
body += paragraph('Preserve visible content when the list description wraps to its continuation line; this deliberately longer item also checks the position of the following table.',1)
body += paragraph('Verify source labels and neighboring mathematical content.',2)
body += '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc>'+paragraph('Table-owned continuation item',2)+paragraph('Verified locally')+'</w:tc></w:tr></w:tbl>'
body += paragraph('Delivery checkpoint',0)+paragraph('Retain output publication guarantees.',1)
body += paragraph('Portfolio record',0,2)+paragraph('Held-out group',1,2)+paragraph('All public source content remains visible.')
rows = [dict(Id='mixed-formats', Body=body, Numbering=numbering, Labels=['IV.','IV.z)','IV.aa)','IV.aa.09)','IV.aa.10)','V.','V.z)','(ZZ)','ZZ-viii'], Picture=True, Math='x+1', Text=['Scope and assumptions','Table-owned continuation item','All public source content remains visible.'])]

levels = level(0,'upperLetter',1,'%1.')+level(1,'lowerRoman',9,'%1.%2', '<w:isLgl/>')
numbering = f'<w:abstractNum w:abstractNumId="20">{levels}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="20"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="12"/>{level(0,"upperRoman",1,"%1.")}</w:lvlOverride></w:num>'
body = paragraph('Public legal-numbering schedule')+paragraph('Effective concrete format owns this section.',0)
body += paragraph('Legal numbering uses decimal values for both levels.',1)+paragraph('The next legal item retains its own counter.',1)
body += paragraph('A new section restarts its subordinate level.',0)+paragraph('The subordinate legal counter starts again.',1)
body += paragraph('End of the independently structured schedule.')
rows.append(dict(Id='legal-overrides',Body=body,Numbering=numbering,Labels=['XII.','12.9','12.10','XIII.','13.9'],Picture=False,Math='',Text=['Effective concrete format owns this section.','The subordinate legal counter starts again.','End of the independently structured schedule.']))

expected=[]
for row in rows:
    namespaces=f'xmlns:w="{w}" xmlns:r="{office[:-1]}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
    parts={'[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
        '_rels/.rels':f'<Relationships xmlns="{rel}"><Relationship Id="document" Type="{office}officeDocument" Target="word/document.xml"/></Relationships>',
        'word/_rels/document.xml.rels':f'<Relationships xmlns="{rel}"><Relationship Id="numbering" Type="{office}numbering" Target="numbering.xml"/><Relationship Id="styles" Type="{office}styles" Target="styles.xml"/>'+ (f'<Relationship Id="picture" Type="{office}image" Target="media/status.png"/><Relationship Id="link" Type="{office}hyperlink" Target="https://example.invalid/public-list-contract" TargetMode="External"/>' if row['Picture'] else '')+'</Relationships>',
        'word/styles.xml':f'<w:styles xmlns:w="{w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>',
        'word/numbering.xml':f'<w:numbering xmlns:w="{w}">{row["Numbering"]}</w:numbering>',
        'word/document.xml':f'<w:document {namespaces}><w:body>{row["Body"]}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="720" w:right="720" w:top="720" w:bottom="720"/></w:sectPr></w:body></w:document>'}
    if row['Picture']:parts['word/media/status.png']=png()
    path=destination/(row['Id']+'.docx')
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as package:
        for name,data in sorted(parts.items()):
            entry=zipfile.ZipInfo(name,(2026,11,15,0,0,0));entry.compress_type=zipfile.ZIP_DEFLATED
            package.writestr(entry,data.encode('utf-8') if isinstance(data,str) else data)
    expected.append({key:value for key,value in row.items() if key not in ['Body','Numbering']} | dict(InputSha256=hashlib.sha256(path.read_bytes()).hexdigest().upper()))
(destination/'expected.json').write_text(json.dumps(expected,indent=2)+'\n',encoding='utf-8')
print('Generated two independent public hierarchical-list sources.')
