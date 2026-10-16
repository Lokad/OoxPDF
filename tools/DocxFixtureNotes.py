"""Shared package transformation for public footnote-to-endnote fixture variants."""
from xml.etree import ElementTree as E

w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
E.register_namespace('w', w[1:-1])


def replace_footnotes_with_endnotes(entries):
    """Return changed package parts; leave unrelated input bytes untouched."""
    entries = dict(entries)
    document = E.fromstring(entries['word/document.xml'])
    for node in document.iter(w + 'footnoteReference'):
        node.tag = w + 'endnoteReference'
    for section in document.iter(w + 'sectPr'):
        setting = section.find(w + 'footnotePr')
        if setting is not None:
            assert section.find(w + 'endnotePr') is None
            setting.tag = w + 'endnotePr'
    entries['word/document.xml'] = E.tostring(document, encoding='utf-8', xml_declaration=True)
    notes = E.fromstring(entries.pop('word/footnotes.xml'))
    notes.tag = w + 'endnotes'
    for node in notes.iter():
        if node.tag == w + 'footnote': node.tag = w + 'endnote'
        elif node.tag == w + 'footnoteRef': node.tag = w + 'endnoteRef'
    if 'word/endnotes.xml' in entries:
        existing = E.fromstring(entries['word/endnotes.xml'])
        assert all(int(node.get(w + 'id', '0')) <= 0 for node in existing)
        for node in existing: notes.insert(0, node)
    entries['word/endnotes.xml'] = E.tostring(notes, encoding='utf-8', xml_declaration=True)
    relationships = E.fromstring(entries['word/_rels/document.xml.rels'])
    has_endnotes = any(node.get('Type') == rel + 'endnotes' for node in relationships)
    for node in list(relationships):
        if node.get('Type') == rel + 'footnotes':
            if has_endnotes: relationships.remove(node)
            else:
                node.set('Type', rel + 'endnotes')
                node.set('Target', 'endnotes.xml')
    entries['word/_rels/document.xml.rels'] = E.tostring(relationships, encoding='utf-8', xml_declaration=True)
    content_types = E.fromstring(entries['[Content_Types].xml'])
    has_endnotes = any(node.get('PartName') == '/word/endnotes.xml' for node in content_types)
    for node in list(content_types):
        if node.get('PartName') == '/word/footnotes.xml':
            if has_endnotes: content_types.remove(node)
            else:
                node.set('PartName', '/word/endnotes.xml')
                node.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml')
    entries['[Content_Types].xml'] = E.tostring(content_types, encoding='utf-8', xml_declaration=True)
    return entries
