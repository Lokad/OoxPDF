
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
- `tools/CheckPrivateCase.ps1`: run a private, ignored visual case.
- `tools/SummarizePrivateCase.ps1`: summarize ignored private-case metrics without document content.
- `tools/RenderReference.ps1`: render Office reference output through COM.
- `tools/RasterizePdf.ps1`: rasterize PDFs through local PDFium.
- `tools/NewOfficeVisualFixtures.ps1`: regenerate Office-authored public fixtures.
- `tools/NewSyntheticFixtures.ps1`: regenerate synthetic fixtures.
- `tools/NewVisualCase.ps1`: scaffold a visual case.
- `tools/Lokad.OoxPdf.VisualDiff`: compare reference and candidate PNGs.
- `tools/Lokad.OoxPdf.PdfiumRasterizer`: dependency-free PDFium P/Invoke rasterizer.

## Private Mode

Use private mode for documents that must not be versioned or made public. Put manifests and inputs under ignored `private-cases/`, run `pwsh tools/CheckPrivateCase.ps1 -Case private-cases/<case>.json`, and review ignored outputs under `artifacts/private-visual/`. The script rejects tracked files, paths outside `private-cases/`, and unsafe case IDs. Public notes must be anonymized: record feature gaps, not private text, screenshots, filenames, or document contents.

## ExecPlans

When writing complex features or significant refactors, use an ExecPlan (as described in PLANS.md) from design to implementation.

## Autonomy Policy

If you're working towards goals, do NOT end your turn. This allows for continuous autonomous work.

The user will interrupt you when required, but they will mostly provide steering messages.

Do not pester the user by ending your turn after a unit of work, as that requires them to keep nudging you to keep working.

You MUST continue working autonomously towards any known objectives until the user interrupts you. Do NOT end your turn until there is absolutely nothing left to do.
