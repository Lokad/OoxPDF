# Lokad.OoxPdf

Convert `.pptx` presentations and `.docx` documents to PDF in **.NET 10**.
The library is safe managed code with no runtime dependencies, Office automation
or native libraries. Office and PDFium are used only by separate validation tools.

[Lokad.OoxPdf 0.2.0](https://www.nuget.org/packages/Lokad.OoxPdf/0.2.0) is available
on NuGet. Read the
[changelog](CHANGELOG.md), [migration guide](docs/PublicApiChanges.md) and
[supported scope and limitations](docs/PreviewRelease.md).

## Usage

Install with `dotnet add package Lokad.OoxPdf --version 0.2.0`.

```csharp
using Lokad.OoxPdf;

OoxPdfConverter.Convert("input.docx", "output.pdf", new OoxPdfOptions
{
    DiagnosticSink = diagnostic => Console.Error.WriteLine(
        $"{diagnostic.Id}: {diagnostic.Message}")
});
```

The same API accepts `.pptx`. Unsupported content may be simplified or omitted;
inspect diagnostics when accepting a conversion. Complex layout, typography and
charts remain approximate rather than pixel-identical to Office.

File and stream overloads support cooperative cancellation. File output publishes
atomically; caller-owned streams stay open. Discard partial output after a
failure. Stream input requires an explicit `OoxPdfOptions.InputKind`.

## Fonts

The default resolver reads installed Windows fonts. For Linux or a portable
deployment, supply `OoxPdfOptions.FontResolver` with packaged or custom fonts.
Missing fonts affect glyph coverage, wrapping and pagination. See
[font setup and HTTP font packs](docs/Fonts.md).

## CLI

From a checkout with the SDK pinned in `global.json`:

```powershell
dotnet run --project src/Lokad.OoxPdf.Cli -f net11.0 -- convert input.docx output.pdf --diagnostics diagnostics.json --strict
```

With `--strict`, exit code `3` means the PDF was produced with warnings or errors.
See [diagnostics and exit codes](docs/Diagnostics.md)
and [DOCX markup modes](docs/DocxMarkupModes.md).

## Documentation

- [Capabilities and fallbacks](docs/Capabilities.md)
- [Rendering architecture](docs/RenderingModel.md)
- [Resource limits and hosting](docs/OperatingEnvelope.md)
- [Qualification results and known failures](docs/PreviewQualification.md)
- [Visual validation](docs/VisualValidation.md) and [private documents](docs/PrivateValidation.md)

## Development

The 0.3.0 branch targets .NET 10 and .NET 11. Install the pinned .NET 11 SDK
(currently RC3, a prerelease) and .NET 10 SDK 10.0.403, which supplies the
compatibility runtime and reference packs. Published 0.2.0 targets
.NET 10. Use .NET 11 for development and extensive tests; select `net10.0`
explicitly for compatibility checks.

```powershell
dotnet build Lokad.OoxPdf.slnx -c Release
dotnet run --no-build --project tests/Lokad.OoxPdf.Tests -c Release -f net11.0 -- --skip-slow
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj -c Release --no-build
```

Packages go directly to ignored `artifacts/nuget/`. CI runs on Windows and Ubuntu.
See [AGENTS.md](AGENTS.md) for focused
test selection and reuse of qualified build artifacts.
