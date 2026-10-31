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
