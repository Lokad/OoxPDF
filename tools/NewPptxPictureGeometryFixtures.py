"""Public synthetic picture-geometry controls, derived from the existing crop fixture."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
import json
import sys

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'tests/Lokad.OoxPdf.Tests/Cases/PictureGeometry'
OUT.mkdir(parents=True, exist_ok=True)
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
for prefix, uri in [('p', P), ('a', A), ('r', R)]:
    E.register_namespace(prefix, uri)

source = ROOT / 'tests/Lokad.OoxPdf.Tests/Cases/pptx-ladder-07-image-crop.pptx'
with ZipFile(source) as archive:
    parts = {name: archive.read(name) for name in archive.namelist()}

# Retain rendering resources without copying source document properties,
# author metadata or thumbnails into the synthetic input.
for name in list(parts):
    if name.startswith('docProps/'):
        del parts[name]
package_relationships = E.fromstring(parts['_rels/.rels'])
for relationship in list(package_relationships):
    if 'docProps/' in relationship.get('Target', ''):
        package_relationships.remove(relationship)
parts['_rels/.rels'] = E.tostring(package_relationships, encoding='utf-8', xml_declaration=True)

controls = [
    dict(name='pic-rect', preset='rect'),
    dict(name='pic-ellipse', preset='ellipse'),
    dict(name='pic-roundRect', preset='roundRect'),
    dict(name='pic-roundRect-adjusted', preset='roundRect', adjustment=40000),
    dict(name='pic-triangle', preset='triangle'),
    dict(name='pic-ellipse-sourcecrop', preset='ellipse', crop=True),
    dict(name='pic-ellipse-rotated-flipped', preset='ellipse', transformed=True),
    dict(name='shape-ellipse-picture-fill', preset='ellipse', shape=True),
    dict(name='shape-roundRect-adjusted-picture-fill', preset='roundRect', shape=True, adjustment=40000),
]

if '--extended' in sys.argv:
    controls = [
        dict(name='pic-ellipse-nonsquare-following', preset='ellipse', nonsquare=True, following=True),
        dict(name='pic-ellipse-grouped-following', preset='ellipse', nonsquare=True, grouped=True, following=True),
        dict(name='pic-svg-ellipse-following', preset='ellipse', nonsquare=True, svg=True, following=True),
        dict(name='pic-ellipse-fill-alpha-crop', preset='ellipse', nonsquare=True, crop=True, fillalpha=True, following=True),
    ]

if '--effects' in sys.argv:
    controls = [
        dict(name='pic-ellipse-outline', preset='ellipse', nonsquare=True, outline=True, following=True),
        dict(name='pic-roundRect-outline-adjusted', preset='roundRect', nonsquare=True, adjustment=40000, outline=True, following=True),
        dict(name='pic-ellipse-outline-transformed', preset='ellipse', nonsquare=True, outline=True, transformed=True, following=True),
        dict(name='pic-ellipse-shadow', preset='ellipse', nonsquare=True, shadow=True, following=True),
        dict(name='pic-roundRect-shadow-adjusted', preset='roundRect', nonsquare=True, adjustment=40000, shadow=True, following=True),
        dict(name='pic-ellipse-blurred-shadow', preset='ellipse', nonsquare=True, shadow=True, blur=True, following=True),
    ]

def xml(element):
    return E.tostring(element, encoding='utf-8', xml_declaration=True)

presentation = E.fromstring(parts['ppt/presentation.xml'])
ids = presentation.find(f'{{{P}}}sldIdLst')
ids.clear()
relationships = E.fromstring(parts['ppt/_rels/presentation.xml.rels'])
for relationship in list(relationships):
    if relationship.get('Type', '').endswith('/slide'):
        relationships.remove(relationship)
types = E.fromstring(parts['[Content_Types].xml'])
for override in list(types):
    if override.get('PartName', '').startswith(('/ppt/slides/slide', '/docProps/')):
        types.remove(override)
source_slide = parts['ppt/slides/slide1.xml']
source_slide_rels = parts['ppt/slides/_rels/slide1.xml.rels']

for number, control in enumerate(controls, 1):
    slide = E.fromstring(source_slide)
    tree = slide.find(f'{{{P}}}cSld/{{{P}}}spTree')
    picture = tree.find(f'{{{P}}}pic')
    picture.find(f'{{{P}}}nvPicPr/{{{P}}}cNvPr').set('name', control['name'])
    fill = picture.find(f'{{{P}}}blipFill')
    crop = fill.find(f'{{{A}}}srcRect')
    if not control.get('crop'):
        fill.remove(crop)
    properties = picture.find(f'{{{P}}}spPr')
    geometry = properties.find(f'{{{A}}}prstGeom')
    geometry.set('prst', control['preset'])
    if 'adjustment' in control:
        E.SubElement(geometry.find(f'{{{A}}}avLst'), f'{{{A}}}gd', name='adj', fmla=f"val {control['adjustment']}")
    xfrm = properties.find(f'{{{A}}}xfrm')
    xfrm.find(f'{{{A}}}ext').attrib.update(cx='2743200', cy='2743200')
    if control.get('nonsquare'):
        xfrm.find(f'{{{A}}}ext').attrib.update(cx='3657600', cy='1828800')
    if control.get('fillalpha'):
        rect = fill.find(f'{{{A}}}stretch/{{{A}}}fillRect')
        rect.attrib.update(l='10000', t='15000', r='-10000')
        E.SubElement(fill.find(f'{{{A}}}blip'), f'{{{A}}}alphaModFix', amt='60000')
    if control.get('transformed'):
        xfrm.attrib.update(rot='1500000', flipH='1')
    line = E.SubElement(properties, f'{{{A}}}ln')
    if control.get('outline'):
        line.set('w', '50800')
        E.SubElement(E.SubElement(line, f'{{{A}}}solidFill'), f'{{{A}}}srgbClr', val='000000')
    else:
        E.SubElement(line, f'{{{A}}}noFill')
    if control.get('shadow'):
        shadow=E.SubElement(E.SubElement(properties,f'{{{A}}}effectLst'),f'{{{A}}}outerShdw',blurRad='63500' if control.get('blur') else '0',dist='91440',dir='2700000',rotWithShape='1')
        E.SubElement(E.SubElement(shadow,f'{{{A}}}srgbClr',val='444444'),f'{{{A}}}alpha',val='50000')
    if control.get('shape'):
        shape = E.Element(f'{{{P}}}sp')
        nonvisual = E.SubElement(shape, f'{{{P}}}nvSpPr')
        E.SubElement(nonvisual, f'{{{P}}}cNvPr', id='3', name=control['name'])
        E.SubElement(nonvisual, f'{{{P}}}cNvSpPr')
        E.SubElement(nonvisual, f'{{{P}}}nvPr')
        shape.append(deepcopy(properties))
        shape_properties = shape.find(f'{{{P}}}spPr')
        shape_fill = deepcopy(fill)
        shape_fill.tag = f'{{{A}}}blipFill'
        shape_properties.insert(2, shape_fill)
        tree.remove(picture)
        tree.append(shape)
    slide_rels = E.fromstring(source_slide_rels)
    if control.get('svg'):
        fill.find(f'{{{A}}}blip').set(f'{{{R}}}embed', 'rIdMaskSvg')
        E.SubElement(slide_rels, f'{{{REL}}}Relationship', Id='rIdMaskSvg', Type=f'{R}/image', Target='../media/mask.svg')
        parts['ppt/media/mask.svg'] = b'<svg xmlns="http://www.w3.org/2000/svg" width="400" height="200" viewBox="0 0 400 200"><rect width="400" height="200" fill="red"/><rect x="100" width="100" height="200" fill="lime"/><rect x="200" width="100" height="200" fill="blue"/><rect x="300" width="100" height="200" fill="yellow"/></svg>'
        if not any(item.get('Extension') == 'svg' for item in types):
            E.SubElement(types, f'{{{CT}}}Default', Extension='svg', ContentType='image/svg+xml')
    if control.get('grouped'):
        group = E.Element(f'{{{P}}}grpSp')
        nv = E.SubElement(group, f'{{{P}}}nvGrpSpPr')
        E.SubElement(nv, f'{{{P}}}cNvPr', id='30', name='Transformed picture group')
        E.SubElement(nv, f'{{{P}}}cNvGrpSpPr'); E.SubElement(nv, f'{{{P}}}nvPr')
        props = E.SubElement(group, f'{{{P}}}grpSpPr')
        gx = E.SubElement(props, f'{{{A}}}xfrm', rot='600000', flipH='1')
        for tag, attrs in [('off',dict(x='914400', y='685800')),('ext',dict(cx='6096000', cy='3429000')),('chOff',dict(x='0', y='0')),('chExt',dict(cx='4876800', cy='2743200'))]:
            E.SubElement(gx, f'{{{A}}}{tag}', **attrs)
        tree.remove(picture);group.append(picture);tree.append(group)
    if control.get('following'):
        shape=E.SubElement(tree, f'{{{P}}}sp')
        nv=E.SubElement(shape, f'{{{P}}}nvSpPr');E.SubElement(nv,f'{{{P}}}cNvPr',id='40',name='Following corner marker');E.SubElement(nv,f'{{{P}}}cNvSpPr');E.SubElement(nv,f'{{{P}}}nvPr')
        props=E.SubElement(shape,f'{{{P}}}spPr');tx=E.SubElement(props,f'{{{A}}}xfrm')
        E.SubElement(tx,f'{{{A}}}off',x='1828800',y='1828800');E.SubElement(tx,f'{{{A}}}ext',cx='254000',cy='254000')
        E.SubElement(props,f'{{{A}}}prstGeom',prst='rect')
        E.SubElement(E.SubElement(props,f'{{{A}}}solidFill'),f'{{{A}}}srgbClr',val='AA00AA')
        E.SubElement(E.SubElement(props,f'{{{A}}}ln'),f'{{{A}}}noFill')
    parts[f'ppt/slides/slide{number}.xml'] = xml(slide)
    parts[f'ppt/slides/_rels/slide{number}.xml.rels'] = xml(slide_rels)
    E.SubElement(ids, f'{{{P}}}sldId', id=str(255 + number), attrib={f'{{{R}}}id': f'rIdControl{number}'})
    E.SubElement(relationships, f'{{{REL}}}Relationship', Id=f'rIdControl{number}', Type=f'{R}/slide', Target=f'slides/slide{number}.xml')
    E.SubElement(types, f'{{{CT}}}Override', PartName=f'/ppt/slides/slide{number}.xml', ContentType='application/vnd.openxmlformats-officedocument.presentationml.slide+xml')

parts['ppt/presentation.xml'] = xml(presentation)
parts['ppt/_rels/presentation.xml.rels'] = xml(relationships)
parts['[Content_Types].xml'] = xml(types)
stem = 'picture-geometry-effects' if '--effects' in sys.argv else 'picture-geometry-extended' if '--extended' in sys.argv else 'picture-geometry-controls'
with ZipFile(OUT / (stem + '.pptx'), 'w', ZIP_DEFLATED) as archive:
    for name, data in parts.items():
        info = ZipInfo(name, (2026, 10, 16, 0, 0, 0))
        info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
(OUT / (stem + '.json')).write_text(json.dumps(controls, indent=2) + '\n', encoding='utf-8')
print(f'Wrote {len(controls)} controls to {OUT}')
