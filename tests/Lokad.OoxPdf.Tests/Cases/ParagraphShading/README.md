# Paragraph shading controls

Independent Word16 exports cover font sizes, wrapping, first/hanging indents,
spacing, style inheritance, empty paragraphs, cells with different margins,
headers/footers, inline math, pictures and four-page paragraph fragments. `office.json`
pins the source/reference hashes and measured Word colors/rectangles.

`CLEAR`/`CELLCLEAR` label the `w:shd w:val="nil"` probes. Word retains their
inherited paragraph shading; these labels do not imply a transparent override.
Application-view calibration remains open. Exact geometry/visual failures remain
in qualification evidence; color agreement alone is not Office parity.

The shading inheritance model follows the existing paragraph style cascade;
see the [Microsoft Open XML shading reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.shading).

`spacing.docx` and held-out `spacing-containers.docx` independently check background
boundaries in body, cells, header and footer. After-spacing is painted between
adjacent shaded paragraphs, including different colors; it stays clear before an
unshaded paragraph or at the end of a container. `spacing-office.json` pins twelve
Word fill samples. `before-office.json` adds independent before-spacing controls,
including `before-boundaries.docx` and held-out `before-containers.docx`. Excess
before-spacing joins shaded paragraphs through the first placed line; page and
container starts keep their clear gaps. `before-edges.docx` distinguishes matching
background edges from changes in paragraph indentation; text alignment may differ.
Image thresholds remain unchanged.
