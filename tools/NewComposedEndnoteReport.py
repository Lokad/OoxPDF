"""Mechanically replace the held-out composed-field report's footnote with an endnote."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
from copy import deepcopy
import argparse, hashlib, json
from DocxFixtureNotes import replace_footnotes_with_endnotes, w

root = Path(__file__).resolve().parent.parent
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--include-math', action='store_true', help='Add two supported fractions inside the endnote')
parser.add_argument('--include-image', action='store_true', help='Add the source PNG beside the two note fractions')
parser.add_argument('--include-page-fields', action='store_true', help='Add PAGE fields to the complete image/math report body, link and table')
parser.add_argument('--include-static-cells', action='store_true', help='Place existing headers/footers in simple cells and add mixed footer content')
args = parser.parse_args()
if args.include_static_cells: args.include_page_fields = True
if args.include_page_fields: args.include_image = True
if args.include_image: args.include_math = True
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases/Composition'
baseline = json.loads((cases / 'composed-fields-report.json').read_text(encoding='utf-8'))[0]
source = root / baseline['Input']
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == baseline['Sha256']
with ZipFile(source) as archive:
    original = {name:archive.read(name) for name in archive.namelist()}
entries = replace_footnotes_with_endnotes(original)
before = E.fromstring(original['word/document.xml'])
after = E.fromstring(entries['word/document.xml'])
m = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
assert ''.join(node.text or '' for node in before.iter(w + 't')) == ''.join(node.text or '' for node in after.iter(w + 't'))
assert [E.tostring(node) for node in before.iter(m + 'oMath')] == [E.tostring(node) for node in after.iter(m + 'oMath')]
assert sum(len(list(E.fromstring(data).iter(m + 'oMath'))) for name,data in entries.items()
           if name.startswith('word/') and name.endswith('.xml')) == baseline['ExpectedEquationCount']
changed = {'word/document.xml','word/footnotes.xml','word/endnotes.xml','word/_rels/document.xml.rels','[Content_Types].xml'}
assert all(entries[name] == data for name,data in original.items() if name not in changed)
assert all(entries[name] == data for name,data in original.items() if name.startswith('word/media/'))
if args.include_math:
    notes = E.fromstring(entries['word/endnotes.xml'])
    note = next(node for node in notes if int(node.get(w + 'id', '0')) > 0)
    paragraph = note.find(w + 'p')
    def text(value):
        run = E.SubElement(paragraph, w + 'r')
        E.SubElement(run, w + 't', {'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text = value
    text(' Calculation: ')
    E.SubElement(paragraph, w + 'bookmarkStart', {w + 'id':'55003', w + 'name':'NoteCalculation'})
    paragraph.append(deepcopy(next(after.iter(m + 'oMath'))))
    E.SubElement(paragraph, w + 'bookmarkEnd', {w + 'id':'55003'})
    text(' and ')
    paragraph.append(deepcopy(next(after.iter(m + 'oMath'))))
    entries['word/endnotes.xml'] = E.tostring(notes, encoding='utf-8', xml_declaration=True)
if args.include_image:
    relationships = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    package_relationships = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    drawing = deepcopy(next(after.iter(w + 'drawing')))
    for node in drawing.iter():
        if node.tag.endswith('}docPr'):
            node.set('id', '55004'); node.set('name', 'NoteCompositionImage')
    blip = next(node for node in drawing.iter() if node.tag.endswith('}blip'))
    image_id = blip.get(relationships + 'embed')
    main_rels = E.fromstring(entries['word/_rels/document.xml.rels'])
    image = next(node for node in main_rels if node.get('Id') == image_id)
    assert image.get('Type') == relationships[1:-1] + '/image'
    blip.set(relationships + 'embed', 'NoteImage')
    run = E.Element(w + 'r'); run.append(drawing)
    paragraph.insert(list(paragraph).index(next(paragraph.iter(m + 'oMath'))), run)
    entries['word/endnotes.xml'] = E.tostring(notes, encoding='utf-8', xml_declaration=True)
    assert 'word/_rels/endnotes.xml.rels' not in entries
    rels = E.Element(package_relationships + 'Relationships')
    E.SubElement(rels, package_relationships + 'Relationship',
                 {'Id':'NoteImage', 'Type':image.get('Type'), 'Target':image.get('Target')})
    entries['word/_rels/endnotes.xml.rels'] = E.tostring(rels, encoding='utf-8', xml_declaration=True)
if args.include_page_fields:
    simple = next(after.iter(w + 'fldSimple'))
    parents = {child:parent for parent in after.iter() for child in parent}
    link = parents[simple]
    assert link.tag == w + 'hyperlink'
    body_paragraph = parents[link]
    table_paragraph = next(p for table in after.iter(w + 'tbl') for p in table.iter(w + 'p')
                           if next(p.iter(m + 'oMath'), None) is not None)
    packed = next(run for run in body_paragraph.iter(w + 'r') if run.find(w + 'instrText') is not None)
    template = simple.find(w + 'r')
    def page_text(value):
        run = deepcopy(template)
        run.find(w + 't').text = value
        run.find(w + 't').set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return run
    def page_field():
        field = deepcopy(simple)
        field.set(w + 'instr', ' PAGE ')
        field.find('.//' + w + 't').text = '1'
        return field
    # Keep existing cached fields and all prose; add page references rather
    # than replacing a meaningful cached word with a page number.
    for owner in [link, table_paragraph]:
        owner.extend([page_text(' on page '), page_field()])
    packed_page = deepcopy(packed)
    packed_page.find(w + 'instrText').text = ' PAGE '
    packed_page.find(w + 't').text = '1'
    body_paragraph.extend([page_text(' (page '), packed_page, page_text(')')])
    entries['word/document.xml'] = E.tostring(after, encoding='utf-8', xml_declaration=True)
if args.include_static_cells:
    probes = root / 'tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes'
    control = probes / 'math-static-cell-composition.docx'
    assert hashlib.sha256(control.read_bytes()).hexdigest().upper() == '4A3BFC92282618F52BFB9B62070A4FBD8BA6AEA11C69EAF5E49714413D5B7696'
    with ZipFile(control) as archive:
        footer_math = E.fromstring(archive.read('word/static-cell-footer.xml')).find('.//' + w + 'p')
        entries['word/media/static-cell-image.png'] = archive.read('word/media/leading-image.png')
    for owner in ['header', 'footer']:
        part = 'word/' + owner + '1.xml'
        story = E.fromstring(entries[part])
        contents = list(story)
        story.clear()
        table = E.SubElement(story, w + 'tbl')
        properties = E.SubElement(table, w + 'tblPr')
        E.SubElement(properties, w + 'tblW', {w + 'w':'9360', w + 'type':'dxa'})
        E.SubElement(properties, w + 'tblLayout', {w + 'type':'fixed'})
        E.SubElement(E.SubElement(table, w + 'tblGrid'), w + 'gridCol', {w + 'w':'9360'})
        cell = E.SubElement(E.SubElement(table, w + 'tr'), w + 'tc')
        E.SubElement(E.SubElement(cell, w + 'tcPr'), w + 'tcW', {w + 'w':'9360', w + 'type':'dxa'})
        cell.extend(contents)
        if owner == 'footer':
            cell.append(deepcopy(footer_math))
            rels = E.Element(package_relationships + 'Relationships')
            E.SubElement(rels, package_relationships + 'Relationship',
                {'Id':'leadingImage', 'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                 'Target':'media/static-cell-image.png'})
            assert 'word/_rels/footer1.xml.rels' not in entries
            entries['word/_rels/footer1.xml.rels'] = E.tostring(rels, encoding='utf-8', xml_declaration=True)
        E.SubElement(story, w + 'p')
        entries[part] = E.tostring(story, encoding='utf-8', xml_declaration=True)
output = cases / ('composed-static-cell-report.docx' if args.include_static_cells else
                  'composed-page-counter-report.docx' if args.include_page_fields else
                  'composed-note-image-math-report.docx' if args.include_image else
                  'composed-note-math-report.docx' if args.include_math else 'composed-endnotes-report.docx')
with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
    for name,data in sorted(entries.items()):
        info = ZipInfo(name, (2026,10,16,0,0,0))
        info.compress_type = ZIP_DEFLATED
        archive.writestr(info, data)
manifest = [baseline | dict(Id=output.stem, Input=output.relative_to(root).as_posix(),
    Sha256=hashlib.sha256(output.read_bytes()).hexdigest().upper(),
    ExpectedEquationCount=baseline['ExpectedEquationCount'] + (2 if args.include_math else 0) + (2 if args.include_static_cells else 0),
    ExpectedInlineImageCount=baseline['ExpectedInlineImageCount'] + (1 if args.include_image else 0) + (1 if args.include_static_cells else 0),
    ExpectedText=baseline['ExpectedText'] + (['Calculation:'] if args.include_math else []) + (['Footer:'] if args.include_static_cells else []),
    Features=[item.replace('footnote', 'endnote') for item in baseline['Features']] +
        (['two supported fractions and a bookmark inside the endnote'] if args.include_math else []) +
        (['source PNG also embedded beside the endnote fractions'] if args.include_image else []) +
        (['simple/packed PAGE fields beside cached fields/math/images, inside an internal hyperlink and in a math/note table cell'] if args.include_page_fields else []) +
        (['existing header/footer prose and equation preserved in simple cells; mixed footer fractions/image/PAGE/NUMPAGES'] if args.include_static_cells else []))]
(cases / (output.stem + '.json')).write_text(json.dumps(manifest, indent=2)+'\n', encoding='utf-8')
print('Pinned one held-out endnote report; existing prose, six equations, media and unrelated package parts preserved' +
      ('; two note fractions added' if args.include_math else '') +
      ('; source PNG also embedded in the note.' if args.include_image else '.'))
