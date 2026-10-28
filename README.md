# Lokad.OoxPdf

Convert `.pptx` presentations and `.docx` documents to PDF in **.NET 10**.
The library is safe managed code with no runtime dependencies, Office automation
or native libraries. Office and PDFium are used only by separate validation tools.

The next release is **0.2.0**, currently unreleased. Read the
[changelog](CHANGELOG.md), [migration guide](docs/PublicApiChanges.md) and
[supported scope and limitations](docs/PreviewRelease.md).

## Usage

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
[font setup and HTTP font packs](https://github.com/Lokad/OoxPDF/blob/master/docs/Fonts.md).

## CLI

From a checkout with the .NET 10 SDK:

```powershell
dotnet run --project src/Lokad.OoxPdf.Cli -- convert input.docx output.pdf --diagnostics diagnostics.json --strict
```

With `--strict`, exit code `3` means the PDF was produced with warnings or errors.
See [diagnostics and exit codes](https://github.com/Lokad/OoxPDF/blob/master/docs/Diagnostics.md)
and [DOCX markup modes](https://github.com/Lokad/OoxPDF/blob/master/docs/DocxMarkupModes.md).

## Documentation

- [Capabilities and fallbacks](https://github.com/Lokad/OoxPDF/blob/master/docs/Capabilities.md)
- [Rendering architecture](https://github.com/Lokad/OoxPDF/blob/master/docs/RenderingModel.md)
- [Resource limits and hosting](https://github.com/Lokad/OoxPDF/blob/master/docs/OperatingEnvelope.md)
- [Qualification results and known failures](https://github.com/Lokad/OoxPDF/blob/master/docs/PreviewQualification.md)
- [Visual validation](https://github.com/Lokad/OoxPDF/blob/master/docs/VisualValidation.md) and [private documents](https://github.com/Lokad/OoxPDF/blob/master/docs/PrivateValidation.md)

## Development

```powershell
dotnet build Lokad.OoxPdf.slnx -c Release
dotnet run --no-build --project tests/Lokad.OoxPdf.Tests -c Release -- --skip-slow
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj -c Release --no-build
```

Packages go to ignored `artifacts/nuget/`. CI runs on Windows and Ubuntu.
See [AGENTS.md](https://github.com/Lokad/OoxPDF/blob/master/AGENTS.md) for focused
test selection and reuse of qualified build artifacts.
