# Lokad.OoxPdf

`Lokad.OoxPdf` is a dependency-free .NET OOXML-to-PDF renderer for `.pptx` and `.docx` files. The library reads Office Open XML packages directly, builds static PDF pages, and does not call Microsoft Office, PDFium, PowerShell, external executables, or third-party packages at runtime.

The current renderer is intended for corporate decks and documents. PPTX rendering is the more mature path; DOCX rendering supports text, page setup, headers and footers, images, tables, numbering, and diagnostics, with layout coverage still being expanded bottom-up.

Release notes are tracked in [CHANGELOG.md](CHANGELOG.md).

## Runtime Model

The converter runtime is pure managed .NET:

- no Office automation;
- no PDFium;
- no PowerShell scripts;
- no native image, font, browser, or PDF libraries;
- no NuGet package dependencies in `src/Lokad.OoxPdf`.

The runtime inputs are the OOXML file or stream, the output PDF path or stream, options, diagnostics callbacks, and resolved font programs. Validation tools under `tools/` use Office and PDFium, but those tools are not part of library conversion.

Rendering follows this broad pipeline:

```text
OOXML package -> typed document/scene model -> layout -> PDF pages/resources
```

Font handling follows a storage-neutral pipeline:

```text
FontRequest -> FontFaceResolution -> IFontProgramSource -> OpenTypeFont -> embedded PDF font
```

## Fonts And Portability

OOXPDF embeds resolved OpenType font programs in generated PDFs when a font source is available. Font access is intentionally not hard-wired to the filesystem: an `IFontResolver` returns a `FontFaceResolution`, and that resolution carries an `IFontProgramSource`.

An `IFontProgramSource` can be backed by files, memory, blob storage, memory-mapped files, embedded resources, or another deployment-specific source:

```csharp
public interface IFontProgramSource
{
    string StableId { get; }
    ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct = default);
}
```

The built-in `WindowsFontResolver` is a file-backed resolver for local Windows font directories. For deterministic Linux operation, `OoxPdfFontPackResolver` can load a packaged font manifest over HTTP(S) and lazily download only the font programs that conversion actually uses:

```csharp
using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;

using var httpClient = new HttpClient();
IFontResolver fontResolver = await OoxPdfFontPackResolver.CreateHttpAsync(
    "office-compatible-v1",
    new Uri("https://example.test/ooxpdf-fonts/"),
    httpClient,
    cancellation.Token);

await OoxPdfConverter.ConvertAsync(
    "input.pptx",
    "output.pdf",
    new OoxPdfOptions { FontResolver = fontResolver },
    cancellation.Token);
```

Each pack is served from `<source>/<packId>/manifest.json`. Font files are addressed by relative paths in the manifest, and each file is checked against its declared byte size and SHA-256 hash before use:

```json
{
  "packId": "office-compatible-v1",
  "files": [
    {
      "relativePath": "files/aptos.ttf",
      "byteSize": 131072,
      "sha256": "0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF"
    }
  ],
  "families": [
    {
      "requestedFamily": "Aptos",
      "resolvedFamily": "Aptos",
      "relativeFontFile": "files/aptos.ttf",
      "weight": 400,
      "italic": false,
      "faceIndex": 0,
      "hasMathTable": false
    }
  ],
  "fallbacks": [
    { "family": "Aptos" }
  ]
}
```

Without resolved font programs, conversion may still produce a PDF, but typography, wrapping, glyph coverage, and visual fidelity are degraded.

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
        DocxMarkupMode = OoxPdfDocxMarkupMode.Final
    });
```

Conversion is cooperatively cancellable. Use the async API when the host already has an async request
lifetime, or the token-aware synchronous overload when the caller is synchronous:

```csharp
using var cancellation = new CancellationTokenSource();

await OoxPdfConverter.ConvertAsync(
    "input.docx",
    "output.pdf",
    cancellation.Token);
```

The same token is passed through package loading, OOXML reading, rendering, font loading, and PDF writing.
Custom `IFontProgramSource.GetBytesAsync(CancellationToken)` implementations should observe the token.

Use `OoxPdfInputKind.Pptx` or `OoxPdfInputKind.Docx` in `OoxPdfOptions.InputKind` to override extension-based detection.

Stream conversion avoids materializing temporary input and output files. For stream input, `InputKind` must be `Pptx` or `Docx` because there is no filename extension to inspect. The converter leaves caller-owned streams open:

```csharp
using var input = new MemoryStream(ooxmlBytes, writable: false);
using var output = new MemoryStream();

await OoxPdfConverter.ConvertAsync(
    input,
    output,
    new OoxPdfOptions
    {
        InputKind = OoxPdfInputKind.Pptx
    },
    cancellation.Token);

byte[] pdfBytes = output.ToArray();
```

Advanced deployments can provide their own font resolver:

```csharp
using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;

public sealed class MyFontResolver : IFontResolver
{
    public FontFaceResolution Resolve(FontRequest request)
    {
        byte[] bytes = LoadFontBytesFromPackageOrBlob(request);

        return new FontFaceResolution(
            request.FamilyName,
            "Resolved Family",
            new FontStyleKey(request.Bold, request.Italic),
            new MemoryFontProgramSource("font-pack:resolved-family:regular", bytes),
            IsFallback: false);
    }
}

OoxPdfConverter.Convert(
    "input.pptx",
    "output.pdf",
    new OoxPdfOptions
    {
        FontResolver = new MyFontResolver()
    });
```

## CLI Usage

```powershell
dotnet build src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj
dotnet src/Lokad.OoxPdf.Cli/bin/Debug/net10.0/Lokad.OoxPdf.Cli.dll convert input.pptx output.pdf --diagnostics diagnostics.json
dotnet src/Lokad.OoxPdf.Cli/bin/Debug/net10.0/Lokad.OoxPdf.Cli.dll convert input.docx output.pdf --docx-markup all --diagnostics diagnostics.json
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

See [docs/DocxMarkupModes.md](docs/DocxMarkupModes.md) for DOCX final/original/simple/all markup mode semantics, CLI usage, and current approximations.

See [docs/RenderingModel.md](docs/RenderingModel.md) for the package, layout, page, and PDF writer architecture.

See [docs/PrivateValidation.md](docs/PrivateValidation.md) for local-only validation of documents that must not be versioned or made public.

## Development

```powershell
dotnet restore Lokad.OoxPdf.slnx
dotnet build Lokad.OoxPdf.slnx
dotnet run --project tests/Lokad.OoxPdf.Tests -- --skip-slow
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --no-restore
```

Packages are written to `artifacts/nuget/`.

`src/Lokad.OoxPdf` is the NuGet library and must remain free of package references. Office automation and PDFium are isolated under `tools/` for validation only.
