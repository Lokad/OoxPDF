"""Generate independent public table-flow documents with Python's standard library."""
from pathlib import Path
from xml.sax.saxutils import escape
import hashlib, json, struct, zlib, zipfile

root = Path(__file__).resolve().parents[1]
destination = root / 'tests/Lokad.OoxPdf.Tests/Cases/MixedTables'
destination.mkdir(exist_ok=True)
w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
relationships = 'http://schemas.openxmlformats.org/package/2006/relationships'

def paragraph(text, before=0, after=0):
    return (f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}"/></w:pPr>'
            f'<w:r><w:t>{escape(text)}</w:t></w:r></w:p>')

def picture(width, height):
    x, y = width * 12700, height * 12700
    return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="80"/></w:pPr><w:r><w:drawing>'
            f'<wp:inline><wp:extent cx="{x}" cy="{y}"/><wp:docPr id="1" name="Public table marker"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Public table marker"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="marker"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{x}" cy="{y}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
            '</wp:inline></w:drawing></w:r></w:p>')

def equation(label, display):
    math = f'<m:oMath><m:r><m:rPr><m:nor/></m:rPr><m:t>{label}</m:t></m:r></m:oMath>'
    return '<w:p>'+('<m:oMathPara>'+math+'</m:oMathPara>' if display else '<w:r><w:t>Rate </w:t></w:r>'+math)+'</w:p>'

def marker_png():
    def chunk(kind, data):
        return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data))
    pixels=b''.join(b'\0'+b''.join(bytes((40,110,180) if x<16 else (210,140,30)) for x in range(32)) for _ in range(16))
    return b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',32,16,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(pixels))+chunk(b'IEND',b'')

expectations=[]
for name, row_count, column_count, width, height, minimum in [
        ('maintenance-matrix',4,3,48,24,False), ('capacity-schedule',5,4,36,18,True)]:
    prose=[]
    rows=[]
    for row in range(row_count):
        cells=[]
        for column in range(column_count):
            label=f'{"Check" if not minimum else "Stage"} {row+1}-{column+1}'
            prose.append(label)
            content=paragraph(label)
            if (row,column)==(1,0):
                content+=picture(width,height)+paragraph('Below the marker', before=60)
                prose.append('Below the marker')
            elif (row,column)==(2,1):
                content+=equation('a+b',display=True)
                content+='<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>'+paragraph('Nested record')+'</w:tc></w:tr></w:tbl>'
                content+=paragraph('After nested record')
                prose.extend(['Nested record','After nested record'])
            elif (row,column)==(row_count-1,column_count-1):
                content+=equation('v+2',display=False)+paragraph('End of schedule')
                prose.extend(['Rate ', 'End of schedule'])
            elif (row,column)==(0,1):
                content+=paragraph('A longer independent description wraps inside its own cell and preserves the neighboring record.')
                prose.append('A longer independent description wraps inside its own cell and preserves the neighboring record.')
            cells.append('<w:tc><w:tcPr><w:tcW w:w="'+str(10080//column_count)+'" w:type="dxa"/></w:tcPr>'+content+'</w:tc>')
        row_properties='<w:trPr><w:trHeight w:val="640" w:hRule="atLeast"/></w:trPr>' if minimum else ''
        rows.append('<w:tr>'+row_properties+''.join(cells)+'</w:tr>')
    borders=''.join(f'<w:{side} w:val="single" w:sz="4" w:color="888888"/>' for side in ['top','left','bottom','right','insideH','insideV'])
    margins=''.join(f'<w:{side} w:w="{value}" w:type="dxa"/>' for side,value in [('top',80),('bottom',80),('left',100),('right',100)])
    table='<w:tbl><w:tblPr><w:tblW w:w="10080" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>'+borders+'</w:tblBorders><w:tblCellMar>'+margins+'</w:tblCellMar></w:tblPr><w:tblGrid>'+''.join(f'<w:gridCol w:w="{10080//column_count}"/>' for _ in range(column_count))+'</w:tblGrid>'+''.join(rows)+'</w:tbl>'
    title='Maintenance matrix' if not minimum else 'Capacity schedule'
    body=paragraph(title,after=160)+table+paragraph('Public table-flow record complete.',before=120)
    prose=[title]+prose+['Public table-flow record complete.']
    namespaces=f'xmlns:w="{w}" xmlns:r="{office[:-1]}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
    parts={
        '[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
        '_rels/.rels':f'<Relationships xmlns="{relationships}"><Relationship Id="document" Type="{office}officeDocument" Target="word/document.xml"/></Relationships>',
        'word/_rels/document.xml.rels':f'<Relationships xmlns="{relationships}"><Relationship Id="styles" Type="{office}styles" Target="styles.xml"/><Relationship Id="marker" Type="{office}image" Target="media/marker.png"/></Relationships>',
        'word/styles.xml':f'<w:styles xmlns:w="{w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="40" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>',
        'word/document.xml':f'<w:document {namespaces}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr></w:body></w:document>',
        'word/media/marker.png':marker_png()}
    path=destination/(name+'.docx')
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as package:
        for part,data in sorted(parts.items()):
            entry=zipfile.ZipInfo(part,(2026,11,15,0,0,0)); entry.compress_type=zipfile.ZIP_DEFLATED
            package.writestr(entry,data.encode('utf-8') if isinstance(data,str) else data)
    expectations.append(dict(Id=name,Rows=row_count,Columns=column_count,PictureWidth=width,PictureHeight=height,
                             Text=prose,Math=['a+b','v+2'],InputSha256=hashlib.sha256(path.read_bytes()).hexdigest().upper()))
(destination/'expected.json').write_text(json.dumps(expectations,indent=2)+'\n',encoding='utf-8')
print('Generated two independent mixed-table sources (automatic and minimum-height rows).')
