
## .NET Commands

Use `--tl:off` to avoid dynamic terminal logger output.

```powershell
dotnet restore Foo.slnx --tl:off -v minimal
dotnet build   Foo.slnx --tl:off --nologo -v minimal
dotnet test    tests/Foo.Tests/Foo.Tests.csproj --tl:off --nologo -v minimal
dotnet pack    src/Foo/Foo.csproj --tl:off --nologo -v minimal --no-restore
```

NuGet packages are written to ignored `artifacts/nuget/`.

## Repo Map

- `src/Lokad.OoxPdf`: dependency-free library and PDF/OOXML renderers.
- `src/Lokad.OoxPdf.Cli`: local CLI wrapper.
- `tests/Lokad.OoxPdf.Tests`: dependency-free console test runner and fixtures.
- `tools/`: validation scripts and helper tools.
- `tools/vendor/`: local validation binaries only, never library dependencies.
- `visual-cases/`: public visual case manifests.
- `private-cases/`: ignored local-only private manifests and inputs.
- `artifacts/`: ignored generated validation output.
- `docs/`: public documentation.
- `EXECPLAN.md`: active execution plan and validation log.

## Tools

- `tools/CheckVisualCase.ps1`: run a public visual case.
- `tools/CheckVisualFamily.ps1`: run or list a public visual capability family.
- `tools/CompareVisualReports.ps1`: compare two visual family reports for regressions.
- `tools/ValidateVisualCases.ps1`: validate public visual case/family naming, coverage, and manifest integrity.
- `tools/CheckPrivateCase.ps1`: run a private, ignored visual case.
- `tools/SummarizePrivateCase.ps1`: summarize ignored private-case metrics without document content.
- `tools/InventoryPptxSlides.ps1`: produce private-safe PPTX slide feature counts.
- `tools/RenderReference.ps1`: render Office reference output through COM.
- `tools/RasterizePdf.ps1`: rasterize PDFs through local PDFium.
- `tools/InspectPdf.ps1`: inspect Office/candidate PDF objects and streams.
- `tools/InspectPptxText.ps1`: inspect PPTX glyph-run layout/emission summaries without text by default.
- `tools/ComparePdfTextOperations.ps1`: compare inspected PDF text matrices and spacing.
- `tools/ComparePptxTextEmission.ps1`: compare Office PDF text operations to candidate PPTX glyph-run emission.
- `tools/SummarizePptxFontEmissionProbe.ps1`: join public-safe PPTX textbox geometry with Office text-operation font sizes.
- `tools/ComparePdfGraphicsOperations.ps1`: compare inspected PDF path/clip/stroke/fill geometry.
- `tools/ClassifyPdfChartGraphics.ps1`: classify inspected PDF graphics operations into chart-like structures.
- `tools/ClassifyPdfChartText.ps1`: classify inspected PDF text operations relative to derived chart plot boxes.
- `tools/NewOfficeVisualFixtures.ps1`: regenerate Office-authored public fixtures.
- `tools/NewSyntheticFixtures.ps1`: regenerate synthetic fixtures.
- `tools/NewVisualCase.ps1`: scaffold a visual case.
- `tools/Lokad.OoxPdf.VisualDiff`: compare reference and candidate PNGs.
- `tools/Lokad.OoxPdf.PdfiumRasterizer`: dependency-free PDFium P/Invoke rasterizer.
- `tools/Lokad.OoxPdf.PdfInspect`: dependency-free PDF object/stream inspector.
- `tools/Lokad.OoxPdf.PptxInspect`: dependency-free PPTX text/glyph-run inspector.

Unit tests can be filtered by capability group with `--group`, for example:
`dotnet run --project tests/Lokad.OoxPdf.Tests --tl:off --nologo -v minimal -- --group pptx-typography --skip-slow`.

## Private Mode

Use private mode for documents that must not be versioned or made public. Put manifests and inputs under ignored `private-cases/`, run `pwsh tools/CheckPrivateCase.ps1 -Case private-cases/<case>.json`, and review ignored outputs under `artifacts/private-visual/`. The script rejects tracked files, paths outside `private-cases/`, and unsafe case IDs. Public notes must be anonymized: record feature gaps, not private text, screenshots, filenames, or document contents.

## ExecPlans

When writing complex features or significant refactors, use an ExecPlan (as described in PLANS.md) from design to implementation.

## Autonomy Policy

If you're working towards goals, do NOT end your turn. This allows for continuous autonomous work.

The user will interrupt you when required, but they will mostly provide steering messages.

Do not pester the user by ending your turn after a unit of work, as that requires them to keep nudging you to keep working.

You MUST continue working autonomously towards any known objectives until the user interrupts you. Do NOT end your turn until there is absolutely nothing left to do.
