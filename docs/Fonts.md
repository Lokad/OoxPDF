# Fonts and portability

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

using var cancellation = new CancellationTokenSource();
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

Custom resolvers return a `FontFaceResolution` containing the requested and
resolved families, style, collection face index and an `IFontProgramSource`.
An explicit resolver defines the complete allowed font universe: the library
does not add system fonts when a custom resolver lacks a catalog. Built-in font
packs retain fallback among their supplied faces. Missing glyphs use the existing
diagnosed replacement; supply the necessary coverage for faithful typography.
Omitting `FontResolver` selects local Windows discovery.
Use `MemoryFontProgramSource` for programs held in memory or implement the source
interface for other storage. Different font programs need different stable IDs;
published bytes must stay immutable and byte-identical for each ID. Resolvers
and sources must support concurrent callers and observe cancellation.

Keep files unchanged while file-backed sources are in use. See the
[font lifetime and migration contract](PublicApiChanges.md) before replacing
fonts or invalidating discovery caches.
