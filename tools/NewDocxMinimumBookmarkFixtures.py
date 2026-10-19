"""Neutral source/container composition controls for minimum body frames."""
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.etree import ElementTree as E
from io import BytesIO
import argparse, copy, json, hashlib

def package_bytes(entries):
    stream = BytesIO()
    with ZipFile(stream, 'w', ZIP_DEFLATED) as archive:
        for path, data in sorted(entries.items()):
            info = ZipInfo(path, (2026, 1, 1, 0, 0, 0)); info.compress_type = ZIP_DEFLATED
            archive.writestr(info, data)
    return stream.getvalue()

def publish(path, data):
    staging = path.with_suffix(path.suffix + '.staging')
    with staging.open('xb') as stream:
        stream.write(data)
    staging.replace(path)

def manifest_bytes(rows):
    return (json.dumps(rows, indent=2) + '\n').replace('\n', '\r\n').encode('utf-8')

parser = argparse.ArgumentParser(description=__doc__)
passive = parser.add_mutually_exclusive_group()
passive.add_argument('--passive-headers', action='store_true')
passive.add_argument('--passive-footers', action='store_true')
passive.add_argument('--page-field-probes', action='store_true')
passive.add_argument('--inactive-static-probes', action='store_true')
passive.add_argument('--inactive-note-probe', action='store_true')
passive.add_argument('--field-code-probes', action='store_true')
passive.add_argument('--simple-field-mark-probes', action='store_true')
passive.add_argument('--simple-field-format-probes', action='store_true')
parser.add_argument('--write', action='store_true', help='Publish selected composition controls; otherwise verify them.')
args = parser.parse_args()

root = Path(__file__).resolve().parent.parent
cases = root / 'tests/Lokad.OoxPdf.Tests/Cases'
output = cases / 'CompositionProbes'
output.mkdir(exist_ok=True)
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

if args.field_code_probes or args.simple_field_mark_probes or args.simple_field_format_probes:
    source = next(row for row in json.loads((output / 'minimum-passive-page-fields.json').read_text())
        if row['FieldInstruction'] == 'PAGE')
    original = (root / source['Input']).read_bytes()
    assert hashlib.sha256(original).hexdigest().upper() == source['Sha256']
    prepared, rows = [], []
    for code_points in [11, 13]:
        with ZipFile(BytesIO(original)) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        body = entries['word/document.xml']
        part = 'word/passive-body-header.xml'
        story = E.fromstring(entries[part])
        paragraph = story.find(w + 'p')
        field = paragraph.find(w + 'fldSimple')
        result = copy.deepcopy(field.find(w + 'r'))
        visible = ''.join(story.itertext())
        if args.simple_field_format_probes:
            field.set(w + 'instr', 'PAGE \\* MERGEFORMAT')
            field.find(w + 'r').find(w + 'rPr').find(w + 'sz').set(w + 'val', str(code_points * 2))
        elif args.simple_field_mark_probes:
            properties = paragraph.find(w + 'pPr')
            assert properties.find(w + 'rPr') is None
            mark = E.SubElement(properties, w + 'rPr')
            E.SubElement(mark, w + 'sz', {w + 'val': str(code_points * 2)})
        else:
            paragraph.remove(field)
            def field_run(tag, value):
                run = E.SubElement(paragraph, w + 'r')
                run.append(copy.deepcopy(result.find(w + 'rPr')))
                node = E.SubElement(run, w + tag)
                if tag == 'fldChar':
                    node.set(w + 'fldCharType', value)
                else:
                    run.find(w + 'rPr').find(w + 'sz').set(w + 'val', str(code_points * 2))
                    node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    node.text = value
            field_run('fldChar', 'begin')
            field_run('instrText', ' PAGE ')
            field_run('fldChar', 'separate')
            paragraph.append(result)
            field_run('fldChar', 'end')
        assert ''.join(node.text or '' for node in story.iter(w + 't')) == visible
        entries[part] = E.tostring(story, encoding='utf-8', xml_declaration=True)
        assert entries['word/document.xml'] == body
        prefix = 'mergeformat-' if args.simple_field_format_probes else 'mark-' if args.simple_field_mark_probes else 'code-'
        name = 'minimum-page-' + prefix + str(code_points)
        data = package_bytes(entries)
        rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + name + '.docx',
            Sha256=hashlib.sha256(data).hexdigest().upper(), SourceId=source['Id'], SourceSha256=source['Sha256'],
            SourceBodySha256=source['SourceBodySha256'], HeldOut=True, ExpectedEquationCount=0,
            ExpectedText=source['ExpectedText'], FieldInstruction='PAGE', CodeFontPoints=code_points,
            VisibleFontPoints=11, Context='ordinary12'))
        if args.simple_field_mark_probes:
            rows[-1]['ParagraphMarkPoints'] = rows[-1].pop('CodeFontPoints')
        elif args.simple_field_format_probes:
            rows[-1]['CachedResultFontPoints'] = rows[-1].pop('CodeFontPoints')
            rows[-1]['VisibleFontPoints'] = code_points
            rows[-1]['FieldInstruction'] = field.get(w + 'instr')
        prepared.append((output / (name + '.docx'), data))
    manifest_name = 'minimum-field-format.json' if args.simple_field_format_probes else 'minimum-field-mark.json' if args.simple_field_mark_probes else 'minimum-field-code.json'
    prepared.append((output / manifest_name, manifest_bytes(rows)))
    for path, data in prepared:
        if args.write:
            publish(path, data)
        else:
            assert path.read_bytes() == data, path.name
    print('Two PAGE metric controls verified; cached result formatting, visible text and body unchanged.')
    raise SystemExit(0)

if args.inactive_static_probes or args.inactive_note_probe:
    r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
    math = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
    rich_cases = json.loads((cases / 'InlineBoxMath/static-table-origins.json').read_text())
    prepared, rows = [], []
    controls = [('header', 'noted12', 'first')] if args.inactive_note_probe else [
        ('header', 'ordinary12', 'first'), ('footer', 'ordinary11', 'even')]
    for kind, context, variant in controls:
        source_manifest = 'minimum-bookmark.json' if args.inactive_note_probe else 'minimum-passive-' + kind + 's.json'
        source = next(row for row in json.loads((output / source_manifest).read_text()) if row['Context'] == context and not row.get('Bookmarked', False))
        rich = next(row for row in rich_cases if row['Id'] == 'static-table-origin-' + kind + '-wide')
        original = (root / source['Input']).read_bytes()
        rich_data = (root / rich['Input']).read_bytes()
        assert hashlib.sha256(original).hexdigest().upper() == source['Sha256']
        assert hashlib.sha256(rich_data).hexdigest().upper() == rich['Sha256']
        with ZipFile(BytesIO(original)) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        with ZipFile(BytesIO(rich_data)) as archive:
            story = E.fromstring(archive.read('word/inline-box-review-header.xml'))
        assert story.tag == w + ('hdr' if kind == 'header' else 'ftr')
        assert not any(node.get(r + key) is not None for node in story.iter() for key in ['id', 'embed', 'link'])
        document = E.fromstring(entries['word/document.xml'])
        body = document.find(w + 'body')
        before = b''.join(E.tostring(node) for node in body if node.tag != w + 'sectPr')
        section = body.find(w + 'sectPr')
        assert section.find(w + 'titlePg') is None
        selected = section.find(w + kind + 'Reference')
        assert selected is None if args.inactive_note_probe else selected.get(w + 'type') == 'default'
        settings = E.fromstring(entries['word/settings.xml'])
        assert settings.find(w + 'evenAndOddHeaders') is None
        part = 'word/inactive-' + variant + '-' + kind + '.xml'
        assert part not in entries
        entries[part] = E.tostring(story, encoding='utf-8', xml_declaration=True)
        relationships = E.fromstring(entries['word/_rels/document.xml.rels'])
        relationship_id = 'rInactiveStatic'
        assert all(node.get('Id') != relationship_id for node in relationships)
        E.SubElement(relationships, rel + 'Relationship', Id=relationship_id, Target=part.removeprefix('word/'), Type=r[1:-1] + '/' + kind)
        section.insert(0, E.Element(w + kind + 'Reference', {w + 'type': variant, r + 'id': relationship_id}))
        entries['word/_rels/document.xml.rels'] = E.tostring(relationships, encoding='utf-8', xml_declaration=True)
        entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
        assert before == b''.join(E.tostring(node) for node in E.fromstring(entries['word/document.xml']).find(w + 'body') if node.tag != w + 'sectPr')
        types = E.fromstring(entries['[Content_Types].xml'])
        E.SubElement(types, ct + 'Override', PartName='/' + part, ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + '+xml')
        entries['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
        name = 'minimum-inactive-' + variant + '-' + kind + ('-noted12' if args.inactive_note_probe else '')
        data = package_bytes(entries)
        rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + name + '.docx',
            Sha256=hashlib.sha256(data).hexdigest().upper(), SourceId=source['Id'], SourceSha256=source['Sha256'],
            SourceBodySha256=hashlib.sha256(before).hexdigest().upper(), RichSourceSha256=rich['Sha256'], RichPart=part,
            HeldOut=True, ExpectedEquationCount=sum(1 for _ in story.iter(math + 'oMath')), ExpectedRenderedEquationCount=0,
            ExpectedText=source['ExpectedText'], Context=context, FontPoints=source['FontPoints'], Variant=variant, StoryKind=kind))
        prepared.append((output / (name + '.docx'), data))
    manifest_name = 'minimum-inactive-notes.json' if args.inactive_note_probe else 'minimum-inactive-static.json'
    prepared.append((output / manifest_name, manifest_bytes(rows)))
    for path, data in prepared:
        if args.write:
            publish(path, data)
        else:
            assert path.read_bytes() == data, path.name
    print(str(len(controls)) + ' inactive rich variants verified; selected stories and body contents/properties unchanged.')
    raise SystemExit(0)

if args.page_field_probes:
    prepared, rows = [], []
    for kind, context, instruction in [('header', 'ordinary12', 'PAGE'), ('footer', 'ordinary11', 'NUMPAGES')]:
        source = next(row for row in json.loads((output / ('minimum-passive-' + kind + 's.json')).read_text()) if row['Context'] == context)
        original = (root / source['Input']).read_bytes()
        assert hashlib.sha256(original).hexdigest().upper() == source['Sha256']
        with ZipFile(BytesIO(original)) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        body = entries['word/document.xml']
        part = 'word/passive-body-' + kind + '.xml'
        story = E.fromstring(entries[part])
        paragraph = story.find(w + 'p')
        run = paragraph.find(w + 'r')
        text = run.find(w + 't')
        text.text += ' '
        text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        field = E.SubElement(paragraph, w + 'fldSimple', {w + 'instr': instruction})
        result = E.SubElement(field, w + 'r')
        result.append(copy.deepcopy(run.find(w + 'rPr')))
        E.SubElement(result, w + 't').text = '1'
        entries[part] = E.tostring(story, encoding='utf-8', xml_declaration=True)
        assert entries['word/document.xml'] == body
        name = 'minimum-passive-' + kind + '-' + instruction.lower()
        data = package_bytes(entries)
        rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + name + '.docx',
            Sha256=hashlib.sha256(data).hexdigest().upper(), SourceId=source['Id'], SourceSha256=source['Sha256'],
            BaseSourceId=source['SourceId'], BaseSourceSha256=source['SourceSha256'], SourceBodySha256=source['SourceBodySha256'],
            HeldOut=True, ExpectedEquationCount=0, ExpectedText=source['ExpectedText'], Context=context,
            FontPoints=source['FontPoints'], FieldInstruction=instruction))
        prepared.append((output / (name + '.docx'), data))
    prepared.append((output / 'minimum-passive-page-fields.json', manifest_bytes(rows)))
    for path, data in prepared:
        if args.write:
            publish(path, data)
        else:
            assert path.read_bytes() == data, path.name
    print('Two source-exact passive PAGE/NUMPAGES controls verified; document part unchanged.')
    raise SystemExit(0)

if args.passive_headers or args.passive_footers:
    kind = 'header' if args.passive_headers else 'footer'
    r = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
    sources = json.loads((output / 'minimum-bookmark.json').read_text())
    prepared, rows = [], []
    for context in ['ordinary12', 'ordinary11']:
        source = next(row for row in sources if row['Context'] == context and not row['Bookmarked'])
        original = (root / source['Input']).read_bytes()
        assert hashlib.sha256(original).hexdigest().upper() == source['Sha256']
        with ZipFile(BytesIO(original)) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        document = E.fromstring(entries['word/document.xml'])
        body = document.find(w + 'body')
        before = b''.join(E.tostring(node) for node in body if node.tag != w + 'sectPr')
        section = body.find(w + 'sectPr')
        assert section is not None and section.find(w + kind + 'Reference') is None
        part = 'word/passive-body-' + kind + '.xml'
        assert part not in entries
        header = E.Element(w + ('hdr' if kind == 'header' else 'ftr'))
        paragraph = E.SubElement(header, w + 'p')
        properties = E.SubElement(paragraph, w + 'pPr')
        E.SubElement(properties, w + 'spacing', {w + 'before': '0', w + 'after': '0', w + 'line': '240', w + 'lineRule': 'auto'})
        run = E.SubElement(paragraph, w + 'r')
        properties = E.SubElement(run, w + 'rPr')
        E.SubElement(properties, w + 'rFonts', {w + 'ascii': 'Calibri', w + 'hAnsi': 'Calibri'})
        E.SubElement(properties, w + 'sz', {w + 'val': '22'})
        E.SubElement(run, w + 't').text = 'Public passive ' + kind + '.'
        entries[part] = E.tostring(header, encoding='utf-8', xml_declaration=True)
        relationships = E.fromstring(entries['word/_rels/document.xml.rels'])
        relationship_id = 'rPassiveBody' + kind.title()
        assert all(node.get('Id') != relationship_id for node in relationships)
        E.SubElement(relationships, rel + 'Relationship', Id=relationship_id, Target=part.removeprefix('word/'),
            Type=r[1:-1] + '/' + kind)
        section.insert(0, E.Element(w + kind + 'Reference', {w + 'type': 'default', r + 'id': relationship_id}))
        entries['word/_rels/document.xml.rels'] = E.tostring(relationships, encoding='utf-8', xml_declaration=True)
        entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
        after = b''.join(E.tostring(node) for node in E.fromstring(entries['word/document.xml']).find(w + 'body') if node.tag != w + 'sectPr')
        assert before == after
        types = E.fromstring(entries['[Content_Types].xml'])
        E.SubElement(types, ct + 'Override', PartName='/' + part,
            ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + '+xml')
        entries['[Content_Types].xml'] = E.tostring(types, encoding='utf-8', xml_declaration=True)
        name = 'minimum-passive-' + kind + '-' + context
        data = package_bytes(entries)
        rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/CompositionProbes/' + name + '.docx',
            Sha256=hashlib.sha256(data).hexdigest().upper(), SourceId=source['Id'], SourceSha256=source['Sha256'],
            SourceBodySha256=hashlib.sha256(before).hexdigest().upper(), HeldOut=True, ExpectedEquationCount=0,
            ExpectedText=source['ExpectedText'] + ['Public passive ' + kind + '.'], Context=context, FontPoints=source['FontPoints']))
        prepared.append((output / (name + '.docx'), data))
    prepared.append((output / ('minimum-passive-' + kind + 's.json'), manifest_bytes(rows)))
    for path, data in prepared:
        if args.write:
            publish(path, data)
        else:
            assert path.read_bytes() == data, path.name
    print('Two ordinary minimum-body/passive-' + kind + ' controls verified; body contents/properties unchanged.')
    raise SystemExit(0)

rows = []
for context, font_size, source_name in [
    ('ordinary12', 12, 'ordinary-minimum-line-20.docx'),
    ('ordinary11', 11, 'ordinary-minimum-line-20.docx'),
    ('noted12', 12, 'note-marked-minimum-line-20.docx')
]:
    source = cases / source_name
    with ZipFile(source) as archive:
        original = {name: archive.read(name) for name in archive.namelist()}
    expected = {'ordinary-minimum-line-20.docx': 'CA9514199B2B8BD44BD0EFF7FF7CAE637AD3A3A0DECDA6ADF584D304B6B4FE74', 'note-marked-minimum-line-20.docx': '7E81696732717C13824B757E4721B1DF5EE414A35205BBE136F5806ABB4D65E4'}
    assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == expected[source_name]
    for bookmarked in [False, True]:
        entries = dict(original)
        document = E.fromstring(entries['word/document.xml'])
        paragraphs = document.find(w + 'body').findall(w + 'p')
        if font_size == 11:
            # The existing eleven-point qualification control uses this same
            # inherited contextual frame and explicit visible-run sizes.
            for paragraph in paragraphs:
                for run in paragraph.findall(w + 'r'):
                    pr = run.find(w + 'rPr')
                    if pr is None:
                        pr = E.Element(w + 'rPr'); run.insert(0, pr)
                    size = pr.find(w + 'sz')
                    if size is None:
                        size = E.SubElement(pr, w + 'sz')
                    size.set(w + 'val', '22')
                    text = run.find(w + 't')
                    if text is not None and text.text:
                        text.text += ' AB.'
                pr = paragraph.find(w + 'pPr')
                spacing = pr.find(w + 'spacing')
                spacing.set(w + 'before', '80'); spacing.set(w + 'after', '80')
                E.SubElement(pr, w + 'contextualSpacing')
        if bookmarked:
            # Enclose an ordinary automatic-height neighbor, not the minimum
            # paragraphs whose visible baselines must remain unchanged.
            paragraph = paragraphs[1]
            first_run = paragraph.find(w + 'r')
            index = list(paragraph).index(first_run)
            paragraph.insert(index, E.Element(w + 'bookmarkStart', {w + 'id': '91', w + 'name': 'PublicNeutralBody'}))
            paragraph.append(E.Element(w + 'bookmarkEnd', {w + 'id': '91'}))
        name = 'minimum-bookmark-' + context + ('-bookmarked' if bookmarked else '-control')
        destination = source if font_size == 12 and not bookmarked else output / (name + '.docx')
        if font_size == 12 and not bookmarked:
            pass  # Reuse the existing public control bytes.
        else:
            entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
            publish(destination, package_bytes(entries))
        rows.append(dict(Id=name,Input='tests/Lokad.OoxPdf.Tests/Cases/'+destination.relative_to(cases).as_posix(),
            Sha256=hashlib.sha256(destination.read_bytes()).hexdigest().upper(),
            SourceSha256=hashlib.sha256(source.read_bytes()).hexdigest().upper(),
            HeldOut=False,ExpectedEquationCount=0,ExpectedText=['Public minimum Top.', 'Public minimum AfterDelta.'],
            Context=context,FontPoints=font_size,Bookmarked=bookmarked))
publish(output / 'minimum-bookmark.json', manifest_bytes(rows))
print('Six public minimum-leading/bookmark cases generated; existing inputs unchanged.')

# Retain the independently exported neighboring static-frame control unchanged.
source = cases / 'CompositionProbes/plain-automatic-batch-after12.docx'
expected = '2986F0A186FBFA3CC557F3980244F06FF045E5C4EA17A959A2C3C4AECA52BB51'
assert hashlib.sha256(source.read_bytes()).hexdigest().upper() == expected
with ZipFile(source) as archive:
    entries = {name: archive.read(name) for name in archive.namelist()}
footer = E.fromstring(entries['word/footer1.xml'])
p = footer.find(w + 'p'); index = list(p).index(p.find(w + 'r'))
p.insert(index, E.Element(w + 'bookmarkStart', {w + 'id': '92', w + 'name': 'PublicNeutralFooter'}))
p.append(E.Element(w + 'bookmarkEnd', {w + 'id': '92'}))
entries['word/footer1.xml'] = E.tostring(footer, encoding='utf-8', xml_declaration=True)
destination = output / 'plain-batch-footer-bookmark.docx'
publish(destination, package_bytes(entries))
static_rows = []
for name, path in [('plain-automatic-batch-after12', source), ('plain-batch-footer-bookmark', destination)]:
    static_rows.append(dict(Id=name, Input='tests/Lokad.OoxPdf.Tests/Cases/' + path.relative_to(cases).as_posix(),
        Sha256=hashlib.sha256(path.read_bytes()).hexdigest().upper(), HeldOut=False, ExpectedEquationCount=0,
        ExpectedText=['Public conclusion.', 'Public second note remains visible']))
publish(output / 'minimum-bookmark-static.json', manifest_bytes(static_rows))
print('Two public static bookmark cases generated; earlier control preserved.')
