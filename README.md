# Lokad.OoxPdf

`Lokad.OoxPdf` is a dependency-free .NET OOXML-to-PDF renderer for `.pptx` and `.docx` files. The library reads Office Open XML packages directly, builds static PDF pages, and does not call Microsoft Office, PDFium, PowerShell, external executables, or third-party packages at runtime.

The current renderer is intended for simple corporate decks and documents. It supports common text, shapes, images, tables, page setup, headers and footers, and emits diagnostics for unsupported features that cannot be represented faithfully in a static PDF.

## Library Usage

```csharp
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;

var diagnostics = new List<OoxPdfDiagnostic>();

OoxPdfConverter.Convert(
    "input.docx",
    "output.pdf",
    new OoxPdfOptions
    {
        DiagnosticSink = diagnostics.Add,
        Deterministic = true
    });
```

Use `OoxPdfInputKind.Pptx` or `OoxPdfInputKind.Docx` in `OoxPdfOptions.InputKind` to override extension-based detection.

## CLI Usage

```powershell
dotnet build src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj --tl:off --nologo -v minimal
dotnet src/Lokad.OoxPdf.Cli/bin/Debug/net10.0/Lokad.OoxPdf.Cli.dll convert input.pptx output.pdf --diagnostics diagnostics.json
```

Add `--strict` to make the command return exit code `3` when conversion succeeds but warnings or errors were emitted.

Exit codes:

- `0`: conversion succeeded without strict-mode warnings or errors.
- `1`: conversion failed.
- `2`: invalid CLI arguments or unsupported input selection.
- `3`: conversion succeeded, but `--strict` saw diagnostics with warning or error severity.

## Visual Validation

Visual validation compares an Office-rendered reference against the candidate PDF rasterized with PDFium:

```powershell
pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-blank/case.json
```

The command writes timestamped artifacts under `artifacts/visual/<case-id>/<run-id>/`, including reference PNGs, candidate PDF and PNGs, diagnostics JSON, pixel metrics, a comparison HTML file, and an assessment template.

The validation tools require Windows with Office installed for reference rendering. PDF rasterization uses `tools/vendor/pdfium/win-x64/bin/pdfium.dll`, retrieved from `https://github.com/bblanchon/pdfium-binaries/releases/latest/download/pdfium-win-x64.tgz`. These tools are not used by the library.

## Capabilities And Limits

See [docs/Capabilities.md](docs/Capabilities.md) for the supported, partial, approximated, and unsupported PPTX/DOCX feature set.

See [docs/Diagnostics.md](docs/Diagnostics.md) for diagnostic code conventions and emitted unsupported-feature warnings.

See [docs/RenderingModel.md](docs/RenderingModel.md) for the package, layout, page, and PDF writer architecture.

See [docs/PrivateValidation.md](docs/PrivateValidation.md) for local-only validation of documents that must not be versioned or made public.

## Development

```powershell
dotnet restore Lokad.OoxPdf.slnx --tl:off -v minimal
dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal
dotnet run --project tests/Lokad.OoxPdf.Tests --tl:off
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore
```

Packages are written to `artifacts/nuget/`.

`src/Lokad.OoxPdf` is the NuGet library and must remain free of package references. Office automation and PDFium are isolated under `tools/` for validation only.
