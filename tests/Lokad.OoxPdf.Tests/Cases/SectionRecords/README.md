# Independent section records

`tools/NewSectionCompositionFixtures.py` writes these public synthetic service
records from handwritten OOXML, without copying the existing report template.
`field-brief.docx` is a two-section Letter brief. Held-out `visit-record.docx` is a
three-section A4 log with different margins, picture placement and prose wrapping.
The small opaque PNG is generated locally; all content is illustrative.

Each section requests a new page and owns its distinct default header/footer,
PAGE/NUMPAGES fields, one footnote and its source marker. Section two owns its
note inside a table; other sections own theirs in body paragraphs. Header and
note mathematics, a picture, an external link and a forward bookmark link must
survive together. `expected.json` specifies content and page ownership before
output is examined. Each note belongs with its reference, each header/footer
with its section, and each whole picture with its authored page.

These contracts establish visible content and ownership, not exact Office
geometry. No Office export or application-view calibration is supplied here.
