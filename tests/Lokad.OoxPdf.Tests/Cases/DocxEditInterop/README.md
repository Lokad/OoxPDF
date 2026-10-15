# DocxEdit interoperability fixtures

These public synthetic documents use DocxEdit 0.2.0 at commit
`67e303117f3e51465f3da3d4971c357c091d407c`. DocxEdit is fixture tooling;
the OoxPdf library has no dependency on it.

Build that revision's `src/Lokad.DocxEdit.Cli` project in Release, then run:

```powershell
pwsh tools/NewDocxEditInteropFixtures.ps1 -ToolDll <path-to-Lokad.DocxEdit.Cli.dll> -OutputDirectory artifacts/docxedit-fixtures
```

The recipes create a document with the producer's default Arial 11 pt styles,
then author a display fraction, an inline Greek script, a fenced matrix, SVG
placements, cropping/resizing, repeated placements, and an equation with an SVG.
The fraction is replaced through its result alias. Crop edits address one image
placement independently. The SVG has opaque, public geometric shapes and no
raster preview. Regeneration can change ZIP metadata and document hashes.

The committed DOCX files sit in the parent Cases folder. Independent Word/PDFium
comparisons use the same visual thresholds as existing math and SVG cases.
Passing producer validation or preserving media bytes does not establish visual
parity. Fenced-matrix optical/baseline differences remain explicit review items.
The consecutive body image control exposed duplicated padding, now removed from
body advance and its keep estimate. Table/story image flow remains separately
qualified; none of these fixtures relaxes the thresholds.
