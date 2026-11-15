"""Generate independent public composition decks; no Office or external libraries."""
from pathlib import Path
from xml.sax.saxutils import escape
import hashlib, json, struct, zipfile, zlib

root=Path(__file__).resolve().parents[1]
destination=root/'tests/Lokad.OoxPdf.Tests/Cases/OperationsDecks'
destination.mkdir(exist_ok=True)
a='http://schemas.openxmlformats.org/drawingml/2006/main'
p='http://schemas.openxmlformats.org/presentationml/2006/main'
c='http://schemas.openxmlformats.org/drawingml/2006/chart'
r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
rels='http://schemas.openxmlformats.org/package/2006/relationships'
ns=f'xmlns:a="{a}" xmlns:p="{p}" xmlns:r="{r}" xmlns:c="{c}"'
group='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
colors=['2878B4','28A05A','D98B25','9261A8','BD4B50','547581']

def xfrm(x,y,width,height,tag='a:xfrm'):
    return f'<{tag}><a:off x="{round(x*12700)}" y="{round(y*12700)}"/><a:ext cx="{round(width*12700)}" cy="{round(height*12700)}"/></{tag}>'

def text(identifier,text,x,y,width,height,size=18,link=False):
    hlink='<a:hlinkClick r:id="schedule"/>' if link else ''
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{identifier}" name="Text {identifier}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr>'+xfrm(x,y,width,height)+'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>'
            '<p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="square"/><a:lstStyle/><a:p><a:pPr/>'
            f'<a:r><a:rPr lang="en-US" sz="{size*100}"><a:solidFill><a:srgbClr val="243442"/></a:solidFill><a:latin typeface="DejaVu Sans"/>{hlink}</a:rPr><a:t>{escape(text)}</a:t></a:r>'
            f'<a:endParaRPr lang="en-US" sz="{size*100}"/></a:p></p:txBody></p:sp>')

def picture(identifier,x,y,width,height):
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{identifier}" name="Public operations marker"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
            '<p:blipFill><a:blip r:embed="marker"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>'+xfrm(x,y,width,height)+
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')

def chart_frame(x,y,width,height):
    return ('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="10" name="Weekly throughput chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'+
            xfrm(x,y,width,height,'p:xfrm')+f'<a:graphic><a:graphicData uri="{c}"><c:chart r:id="chart"/></a:graphicData></a:graphic></p:graphicFrame>')

def tx_properties(size):
    return f'<a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{size*100}"><a:latin typeface="DejaVu Sans"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p>'

def chart(names,legend):
    series=[]
    for index,name in enumerate(names):
        series.append(f'<c:ser><c:idx val="{index}"/><c:order val="{index}"/><c:tx><c:v>{name}</c:v></c:tx>'
                      f'<c:spPr><a:solidFill><a:srgbClr val="{colors[index]}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>'
                      '<c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>Week A</c:v></c:pt><c:pt idx="1"><c:v>Week B</c:v></c:pt></c:strLit></c:cat>'
                      f'<c:val><c:numLit><c:formatCode>0</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>{index+2}</c:v></c:pt><c:pt idx="1"><c:v>{index+3}</c:v></c:pt></c:numLit></c:val></c:ser>')
    title='<c:title><c:tx><c:rich>'+tx_properties(14).replace('<a:endParaRPr lang="en-US"/>','<a:r><a:rPr sz="1400"><a:latin typeface="DejaVu Sans"/></a:rPr><a:t>Weekly throughput</a:t></a:r>')+'</c:rich></c:tx><c:overlay val="0"/></c:title>'
    cat='<c:catAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="200"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>'
    value='<c:valAx><c:axId val="200"/><c:scaling><c:orientation val="minMax"/><c:max val="10"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="0" sourceLinked="0"/><c:majorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="100"/><c:crosses val="autoZero"/><c:crossBetween val="between"/><c:majorUnit val="2"/></c:valAx>'
    return (f'<c:chartSpace xmlns:c="{c}" xmlns:a="{a}" xmlns:r="{r}"><c:lang val="en-US"/><c:chart>'+title+
            '<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'+''.join(series)+
            '<c:gapWidth val="150"/><c:overlap val="0"/><c:axId val="100"/><c:axId val="200"/></c:barChart>'+cat+value+
            f'</c:plotArea><c:legend><c:legendPos val="{legend}"/><c:overlay val="0"/><c:txPr>'+tx_properties(11)+'</c:txPr></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr>'+tx_properties(10)+'</c:txPr></c:chartSpace>')

def marker_png():
    def chunk(kind,data):
        return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data))
    pixels=b''.join(b'\0'+b''.join(bytes((40,120,180) if x<16 else (40,160,90)) for x in range(32)) for _ in range(16))
    return b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',32,16,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(pixels))+chunk(b'IEND',b'')

def relationships(entries):
    return f'<Relationships xmlns="{rels}">'+''.join(f'<Relationship Id="{identifier}" Type="{r}/{kind}" Target="{target}"{mode}/>' for identifier,kind,target,mode in entries)+'</Relationships>'

theme_colors=['000000','FFFFFF','243442','EEEEEE']+colors+['0563C1','954F72']
scheme=''.join(f'<a:{name}><a:srgbClr val="{color}"/></a:{name}>' for name,color in zip(['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'],theme_colors))
font='<a:latin typeface="DejaVu Sans"/><a:ea typeface=""/><a:cs typeface=""/>'
theme=f'<a:theme xmlns:a="{a}" name="Public operations"><a:themeElements><a:clrScheme name="Operations">{scheme}</a:clrScheme><a:fontScheme name="Operations"><a:majorFont>{font}</a:majorFont><a:minorFont>{font}</a:minorFont></a:fontScheme><a:fmtScheme name="Operations"><a:fillStyleLst>'+('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'*3)+'</a:fillStyleLst><a:lnStyleLst>'+('<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'*3)+'</a:lnStyleLst><a:effectStyleLst>'+('<a:effectStyle><a:effectLst/></a:effectStyle>'*3)+'</a:effectStyleLst><a:bgFillStyleLst>'+('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'*3)+'</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'
master=f'<p:sldMaster {ns}><p:cSld><p:spTree>{group}</p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>'
layout=f'<p:sldLayout {ns} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>{group}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'

expected=[]
for identifier,width,height,names,legend,held_out in [
        ('service-operations',800,600,['Intake','Inspect','Clean','Calibrate','Test','Handover'],'r',False),
        ('survey-operations',1000,500,['Survey','Prepare','Measure','Adjust','Verify','Archive'],'b',True)]:
    title='Service operations' if not held_out else 'Survey operations'
    chart_box=[30,120,740,390] if not held_out else [36,105,610,320]
    first=text(2,title,32,22,width-140,40,28)+text(3,'Illustrative weekly work volume',32,72,width-100,30,16)
    first+=picture(4,width-80,24,48,24)+chart_frame(*chart_box)+text(5,'All values are synthetic.',32,height-36,width-100,24,12)
    if held_out:first+=text(6,'Keep all six named work streams visible.',700,170,260,100,18)
    second=text(2,'Record checklist',32,24,width-80,45,26)+text(3,'Confirm that every label remains visible.',40,110,width-260,56,20)
    second+=text(4,'Keep the status marker with the record.',40,190,width-260,56,20)+text(5,'Public schedule',40,280,220,35,18,True)
    second+=picture(6,width-210,150,120,60)+text(7,'Retain the public record.',32,height-44,width-100,28,14)
    parts={
        '_rels/.rels':relationships([('document','officeDocument','ppt/presentation.xml','')]),
        'ppt/presentation.xml':f'<p:presentation {ns}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="master"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="slide1"/><p:sldId id="257" r:id="slide2"/></p:sldIdLst><p:sldSz cx="{width*12700}" cy="{height*12700}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
        'ppt/_rels/presentation.xml.rels':relationships([('master','slideMaster','slideMasters/slideMaster1.xml',''),('slide1','slide','slides/slide1.xml',''),('slide2','slide','slides/slide2.xml','')]),
        'ppt/slideMasters/slideMaster1.xml':master,
        'ppt/slideMasters/_rels/slideMaster1.xml.rels':relationships([('layout','slideLayout','../slideLayouts/slideLayout1.xml',''),('theme','theme','../theme/theme1.xml','')]),
        'ppt/slideLayouts/slideLayout1.xml':layout,
        'ppt/slideLayouts/_rels/slideLayout1.xml.rels':relationships([('master','slideMaster','../slideMasters/slideMaster1.xml','')]),
        'ppt/theme/theme1.xml':theme,'ppt/charts/chart1.xml':chart(names,legend),'ppt/media/marker.png':marker_png()}
    for number,shapes in enumerate([first,second],1):
        parts[f'ppt/slides/slide{number}.xml']=f'<p:sld {ns}><p:cSld><p:spTree>{group}{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
        entries=[('layout','slideLayout','../slideLayouts/slideLayout1.xml',''),('marker','image','../media/marker.png','')]
        entries+= [('chart','chart','../charts/chart1.xml','')] if number==1 else [('schedule','hyperlink','https://example.invalid/public-operations-schedule',' TargetMode="External"')]
        parts[f'ppt/slides/_rels/slide{number}.xml.rels']=relationships(entries)
    overrides={'ppt/presentation.xml':'presentationml.presentation.main','ppt/slideMasters/slideMaster1.xml':'presentationml.slideMaster','ppt/slideLayouts/slideLayout1.xml':'presentationml.slideLayout','ppt/theme/theme1.xml':'theme','ppt/charts/chart1.xml':'chart'}
    overrides.update({f'ppt/slides/slide{i}.xml':'presentationml.slide' for i in [1,2]})
    types='<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
    for part,kind in overrides.items():
        content='application/vnd.openxmlformats-officedocument.'+kind+'+xml'
        if kind in ['theme','chart']:content='application/vnd.openxmlformats-officedocument.drawingml.'+kind+'+xml'
        types+=f'<Override PartName="/{part}" ContentType="{content}"/>'
    parts['[Content_Types].xml']=f'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">{types}</Types>'
    path=destination/(identifier+'.pptx')
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as package:
        for part,data in sorted(parts.items()):
            entry=zipfile.ZipInfo(part,(2026,11,15,0,0,0));entry.compress_type=zipfile.ZIP_DEFLATED
            package.writestr(entry,data.encode('utf-8') if isinstance(data,str) else data)
    first_text=[title,'Illustrative weekly work volume','Weekly throughput','Week A','Week B','All values are synthetic.']+names
    if held_out:first_text+=['Keep all six named work streams visible.']
    expected.append(dict(Id=identifier,HeldOut=held_out,InputSha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(),Size=[width,height],LegendNames=names,ChartBox=chart_box,
                         Slides=[dict(Page=1,Text=first_text,Picture=[width-80,24,48,24],Links=0),dict(Page=2,Text=['Record checklist','Confirm that every label remains visible.','Keep the status marker with the record.','Public schedule','Retain the public record.'],Picture=[width-210,150,120,60],Links=1)]))
(destination/'expected.json').write_text(json.dumps(expected,indent=2)+'\n',encoding='utf-8')
print('Generated independent 4:3 and 2:1 operations decks with charts, six-name legends, prose, pictures and a link.')
