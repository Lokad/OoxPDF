# Fonts and portability

OOXPDF embeds resolved OpenType font programs in generated PDFs when a font source is available. Font access is intentionally not hard-wired to the filesystem: an `IFontResolver` returns a `FontFaceResolution`, and that resolution carries an `IFontProgramSource`.

An `IFontProgramSource` can be backed by files, memory, blob storage, memory-mapped files, embedded resources, or another deployment-specific source:

```csharp
public interface IFontProgramSource
{
    string StableId { get; }
    ValueTask<IOoxPdfFontProgramLease> AcquireAsync(
        IMemoryBudget? operationBudget, CancellationToken ct = default);
}

public interface IOoxPdfFontProgramLease : IDisposable
{
    ReadOnlyMemory<byte> Bytes { get; }
}
```

The built-in `WindowsFontResolver` is a file-backed resolver for local Windows font directories. For deterministic Linux operation, `OoxPdfFontPackResolver` can load a packaged font manifest over HTTP(S) and lazily download only the font programs that conversion actually uses:

```csharp
using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;

using var cancellation = new CancellationTokenSource();
using var httpClient = new HttpClient();
using var fontResolver = await OoxPdfFontPackResolver.CreateHttpAsync(
    "office-compatible-v1",
    new Uri("https://example.test/ooxpdf-fonts/"),
    httpClient,
    cancellationToken: cancellation.Token);

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
For DOCX runs without a declared or theme family, the library asks the resolver
for `OOXPDF_DOCUMENT_DEFAULT`. A resolver can supply that request explicitly;
otherwise the library tries Times New Roman, Aptos, Calibri and Arial in order,
then uses the resolver's fallback. This stable policy can differ from Word's
application default. Declare the document's font or bind that resolver request
when matching a particular Office environment: different font metrics change
table widths, line heights and pagination.

Use `MemoryFontProgramSource` for programs held in memory or implement the source
interface for other storage. Different font programs need different stable IDs;
published bytes must stay immutable and byte-identical for each ID. Resolvers
and sources must support concurrent callers and observe cancellation.

Keep each acquired lease alive through its final reader, then dispose it. The
parser makes an owned font copy, so its source lease ends after copying/parsing.
`MemoryFontProgramSource` borrows caller-owned immutable memory without copying
or charging it again; keep that storage and its funding alive for the source.
Custom sources fund significant acquisition buffers before allocation and finish
outstanding reads before refunding them, including when cancellation is requested.

Pass a longer-lived `cacheBudget` to `CreateHttpAsync` or `FileFontProgramSource`
to fund shared cached programs. Without one, operations with a memory budget fund
private leases rather than lend their short-lived budget to cache retention.
Ungoverned local convenience retains its bounded cache. Dispose the pack resolver
or an owned file source to end retention and refuse new acquisition. Existing
leases and pending I/O keep their admission until their final user finishes;
keep the cache budget alive until those users complete. `HttpClient` stays
caller-owned. Admission refusal is terminal, with `SharedResource` scope for a
cache budget; it is never a transport retry or a missing-font fallback.

Keep files unchanged while file-backed sources are in use. See the
[font lifetime and migration contract](PublicApiChanges.md) before replacing
fonts or invalidating discovery caches.
