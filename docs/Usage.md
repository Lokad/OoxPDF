# Using Lokad.OoxPdf

This guide describes the current source API. The library converts DOCX and PPTX
to PDF without Office, native libraries or runtime dependencies.

## Conversion

```csharp
using Lokad.OoxPdf;

await OoxPdfConverter.ConvertAsync("input.docx", "output.pdf",
    new OoxPdfOptions
    {
        DiagnosticSink = d => Console.Error.WriteLine($"{d.Id}: {d.Message}")
    }, cancellationToken);
```

The same overload accepts PPTX. File output is staged and atomically published;
a failed conversion preserves the existing destination. Stream overloads leave
caller streams open. Specify `InputKind` for stream input, use distinct input and
output streams, and provide an empty output positioned at zero when seekable.
Discard any output prefix after a stream conversion fails.

The asynchronous API awaits input, font acquisition, staging and output I/O.
Layout and drawing execute synchronously between awaits. Cancellation is
cooperative. A custom resolver supplies the complete allowed font universe;
see [fonts](Fonts.md) and [hosting](Hosting.md).

`FixedCreationDate` optionally supplies PDF creation/modification dates. Without
it, dates are omitted. Stable input, fonts and options produce deterministic
output; different font installations can change layout and PDF bytes.

## DOCX review views

| `DocxMarkupMode` | Visible content |
|---|---|
| `Final` (default) | Final text; no printed comments or revision indicators. |
| `Original` | Original text; insertions and moved-to text hidden. |
| `SimpleMarkup` | Final text with lightweight comment/change indicators. |
| `AllMarkup` | Insertions, deletions and moves, with supported revision styling, comments and balloons. |

`DocxMarkupGeometryMode` selects `PreserveDocumentLayout` (default),
`ReserveMarkupMargin`, or `WordCompatibleAllMarkup`. The last uses an approximate
Word print-review profile. Review pagination, balloon collision handling and
custom-paper placement remain limited; see [capabilities](Capabilities.md).
Markup options are ignored for PPTX. Document visibility settings still apply.

## Diagnostics and failures

A successful conversion can contain approximated or omitted content. Review
`DiagnosticSink` when deciding whether the result is acceptable. Diagnostic
records include an ID, severity, message and available part/page/feature/fallback
metadata. An observer exception aborts conversion.

| Prefix | Area |
|---|---|
| `OOXML_` | Package, relationships, XML and dialect handling |
| `DOCX_`, `PPTX_` | Format-specific omissions or approximations |
| `FONT_` | Substitution, coverage, parsing and unsupported outlines |
| `IMAGE_` | Missing, malformed or unsupported pictures/SVG |
| `PDF_` | PDF generation |
| `CONVERSION_` | Optional resource summary |

`OoxPdfLimitExceededException` identifies `Kind` and `Scope`. Memory refusal and
cancellation terminate conversion; they do not request lower-quality output.
Unsupported content can instead use a diagnosed fallback. Hidden comments do
not generate a warning merely because their source parts exist. A diagnostic
list is not a complete inventory of the source document.

## CLI

```powershell
dotnet run --project src/Lokad.OoxPdf.Cli -f net11.0 -- convert input.docx output.pdf --diagnostics diagnostics.json --strict
```

Options include `--docx-markup final|original|simple|all` and
`--docx-markup-geometry preserve|reserve-margin|word-compatible`.

| Exit | Meaning |
|---|---|
| 0 | Conversion succeeded under the selected diagnostic policy |
| 1 | Conversion or diagnostics publication failed |
| 2 | Invalid command or arguments |
| 3 | PDF produced, but `--strict` found warnings or errors |

Strict diagnostic exit handling belongs to the CLI. Library hosts choose their
own policy. Review [scope](Capabilities.md) before treating success as fidelity.
