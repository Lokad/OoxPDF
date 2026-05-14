# Lokad.OoxPdf — stand-alone engineering specification

**Draft version:** 1.0
**Date:** 2026-05-14
**Primary audience:** coding agent implementing the project
**Project name:** `Lokad.OoxPdf`
**Output:** C#/.NET library, published as a NuGet package, converting `.pptx` and `.docx` to `.pdf` without runtime dependencies beyond .NET standard libraries.

OOXML is a suitable input target because the Office Open XML file formats are standardized as ECMA-376 / ISO/IEC 29500 and are based on ZIP and XML. Microsoft’s Open XML documentation also describes the formats as ZIP/XML-based and distinguishes word processing and presentation packages. ([Ecma International][1]) The Library of Congress overview identifies DOCX as WordprocessingML and PPTX as PresentationML packages. ([The Library of Congress][2])

---

## 1. Product intent

`Lokad.OoxPdf` is a small, dependency-free Office renderer for corporate documents. Its job is to read a `.pptx` or `.docx` package and produce a visually faithful static `.pdf`.

The expected document domain is ordinary corporate material: decks, reports, proposals, invoices, agendas, internal memos, exported templates, tables, charts, logos, headers, footers, and images.

The project explicitly **does not** attempt to reproduce Office as a full application. It renders the final static visual appearance that would appear in a PDF export. Interactive or dynamic Office features are ignored or approximated with diagnostics.

The project must also include a **visual validation harness** that lets a coding agent compare `Lokad.OoxPdf` output against Office 365 output. The agent can inspect `.png` or `.jpg` images, so the harness must produce side-by-side visual artifacts for self-assessment.

---

## 2. Non-goals

The following are out of scope for the library:

1. Running Microsoft Office, LibreOffice, browser engines, printer drivers, or external processes during conversion.
2. Depending on Open XML SDK, PdfSharp, SkiaSharp, ImageSharp, System.Drawing.Common, Aspose, iText, HarfBuzz, FreeType, Ghostscript, LibreOffice, Chromium, or any other third-party package.
3. Supporting video playback, audio playback, macros, VBA execution, animation timelines, slide transitions, embedded ActiveX controls, or live OLE object execution.
4. Preserving editability of Office content in the PDF.
5. Producing tagged PDF, PDF/A, accessibility structures, comments, reviewer markup, form fields, or digital signatures in the first release.
6. Achieving byte-identical output with Microsoft Office.
7. Achieving pixel-perfect agreement in antialiasing. Layout, object placement, colors, typography, and missing elements matter more than subpixel rasterization differences.

The validation harness may use external tools under `tools/`, including Office 365 automation and a vendored Windows `pdfium_test.exe`, but these must never become runtime dependencies of `Lokad.OoxPdf`.

---

## 3. Hard constraints

### 3.1 Library dependency policy

The NuGet package must declare **zero `PackageReference` dependencies**.

Allowed APIs include .NET base class libraries such as:

* `System.IO`
* `System.IO.Compression.ZipArchive`
* `System.Xml`
* `System.Xml.Linq`
* `System.Text.Json`
* `System.Buffers`
* `System.Numerics`
* `System.Security.Cryptography`
* `System.IO.Compression.ZLibStream` when targeting a modern .NET runtime

`ZipArchive` is the standard .NET API to read and create ZIP archives, and `XmlReader` provides forward-only XML reading from standard libraries. ([Microsoft Learn][3])

The core library must not use `System.Drawing.Common`. Microsoft documents `System.Drawing.Common` as Windows-specific starting with .NET 6 and not supported on non-Windows platforms. ([Microsoft Learn][4])

### 3.2 Office automation policy

Office automation is allowed **only in the validation harness** and only on a Windows development machine where Office 365 is installed and licensed.

It must not appear in:

* `src/Lokad.OoxPdf`
* the public NuGet package
* library tests that are expected to run cross-platform
* production conversion code

Microsoft states that Office automation is not recommended or supported from unattended, non-interactive applications because Office may show dialogs, hang, or deadlock. Therefore, Office automation scripts in this repo are development/reference tools, not production components. ([Microsoft Learn][5])

### 3.3 Target framework

Initial target:

```xml
<TargetFramework>net10.0</TargetFramework>
```

Rationale: modern standard library availability, no dependency requirement, simpler initial implementation.

A later version may multi-target, but only if it does not introduce dependencies.

### 3.4 Determinism

The converter must support deterministic output mode:

* stable object ordering in PDF
* stable resource names
* no ambient timestamps unless explicitly requested
* fixed PDF metadata when `OoxPdfOptions.FixedCreationDate` is supplied
* diagnostics emitted in stable order where possible

---

## 4. Repository layout

Required layout:

```text
Lokad.OoxPdf/
  README.md
  LICENSE
  Directory.Build.props
  src/
    Lokad.OoxPdf/
      Lokad.OoxPdf.csproj
      OoxPdfConverter.cs
      OoxPdfOptions.cs
      Diagnostics/
      Ooxml/
      Pptx/
      Docx/
      Layout/
      Pdf/
      Fonts/
      Imaging/
    Lokad.OoxPdf.Cli/
      Lokad.OoxPdf.Cli.csproj
      Program.cs
  tests/
    Lokad.OoxPdf.Tests/
      Lokad.OoxPdf.Tests.csproj
      Program.cs
      Cases/
  tools/
    RenderReference.ps1
    RasterizePdf.ps1
    CheckVisualCase.ps1
    NewVisualCase.ps1
    vendor/
      pdfium/
        win-x64/
          pdfium_test.exe
          LICENSE.txt
          VERSION.txt
    Lokad.OoxPdf.VisualDiff/
      Lokad.OoxPdf.VisualDiff.csproj
      Program.cs
  visual-cases/
    README.md
    cases/
      sample-pptx-basic/
        case.json
      sample-docx-basic/
        case.json
  artifacts/
    .gitkeep
  docs/
    Capabilities.md
    RenderingModel.md
    Diagnostics.md
    VisualValidation.md
```

`artifacts/` must be ignored by Git except for `.gitkeep`.

The NuGet package should be created from an SDK-style project using `dotnet pack`; Microsoft recommends SDK-style projects for modern NuGet package creation. ([Microsoft Learn][6])

---

## 5. Public API

The public API must be small and stable.

```csharp
namespace Lokad.OoxPdf;

public static class OoxPdfConverter
{
    public static void Convert(
        string inputPath,
        string outputPdfPath,
        OoxPdfOptions? options = null);

    public static void Convert(
        Stream input,
        string inputNameOrExtension,
        Stream outputPdf,
        OoxPdfOptions? options = null);
}
```

```csharp
namespace Lokad.OoxPdf;

public sealed class OoxPdfOptions
{
    public OoxPdfInputKind InputKind { get; init; } = OoxPdfInputKind.Auto;

    public IFontResolver? FontResolver { get; init; }

    public Action<OoxPdfDiagnostic>? DiagnosticSink { get; init; }

    public bool FailOnUnsupportedFeature { get; init; } = false;

    public bool EnableExternalRelationships { get; init; } = false;

    public string? ExternalRelationshipBaseDirectory { get; init; }

    public DateTimeOffset? FixedCreationDate { get; init; }

    public bool Deterministic { get; init; } = true;
}
```

```csharp
namespace Lokad.OoxPdf;

public enum OoxPdfInputKind
{
    Auto = 0,
    Pptx = 1,
    Docx = 2
}
```

```csharp
namespace Lokad.OoxPdf;

public interface IFontResolver
{
    FontResolution? Resolve(FontRequest request);
}
```

```csharp
namespace Lokad.OoxPdf;

public sealed record OoxPdfDiagnostic(
    string Id,
    OoxPdfSeverity Severity,
    string Message,
    string? PartName = null,
    string? XPath = null,
    int? PageIndex = null,
    int? SlideIndex = null,
    string? Feature = null,
    string? Fallback = null);
```

No public API should expose Open XML SDK types, Office COM types, PDFium types, or any external library type.

---

## 6. Command-line tool

The CLI is a development convenience and dogfooding target. It may be published later as a separate tool, but the core deliverable is the library.

Required CLI shape:

```powershell
dotnet run --project src/Lokad.OoxPdf.Cli -- convert input.pptx output.pdf --diagnostics diagnostics.json
dotnet run --project src/Lokad.OoxPdf.Cli -- convert input.docx output.pdf --strict
```

CLI requirements:

* Exit code `0`: success.
* Exit code `1`: conversion failed.
* Exit code `2`: unsupported input or invalid arguments.
* Exit code `3`: conversion succeeded but strict mode failed due to warnings or unsupported features.
* Emit JSON diagnostics when `--diagnostics` is supplied.
* Do not call Office, PDFium, or external tools.

---

## 7. Internal architecture

The implementation must be layered.

```text
OOXML package reader
  -> semantic document model
    -> layout / page scene model
      -> PDF writer
```

### 7.1 Package reader

Namespace: `Lokad.OoxPdf.Ooxml`

Responsibilities:

* Open OOXML ZIP package with `ZipArchive`.
* Read `[Content_Types].xml`.
* Read package relationships from `_rels/.rels`.
* Resolve part relationships from `*.rels`.
* Normalize part paths.
* Prevent path traversal.
* Reject or cap suspicious ZIP packages.
* Expose parts as streams or parsed XML.
* Handle `mc:AlternateContent` using a deterministic policy:

  * Prefer known `mc:Choice` namespaces.
  * Otherwise use `mc:Fallback`.
  * Otherwise emit diagnostic and skip.

Do not assume fixed file paths except for initial relationship discovery. Always resolve through relationships where the format expects relationships.

### 7.2 Semantic model

Namespaces:

* `Lokad.OoxPdf.Pptx`
* `Lokad.OoxPdf.Docx`

Responsibilities:

* Convert raw XML to stable semantic objects.
* Resolve style inheritance.
* Resolve theme colors and fonts.
* Resolve media references.
* Preserve source locations for diagnostics.
* Avoid PDF-specific assumptions.

### 7.3 Layout/page scene model

Namespace: `Lokad.OoxPdf.Layout`

All renderable content should become a page-level scene graph:

```csharp
internal sealed class PageScene
{
    public double WidthPt { get; }
    public double HeightPt { get; }
    public IReadOnlyList<SceneNode> Nodes { get; }
}
```

Scene coordinates:

* Unit: PDF point.
* Origin: top-left.
* Positive X: right.
* Positive Y: down.

PDF output is bottom-left coordinate based; the coordinate transform must happen only in the PDF writer.

Scene node types:

* `TextRunNode`
* `TextBoxNode`
* `PathNode`
* `ImageNode`
* `TableNode` or table decomposed into fills, borders, and text
* `GroupNode`
* `ClipNode`
* `UnsupportedPlaceholderNode`

### 7.4 PDF writer

Namespace: `Lokad.OoxPdf.Pdf`

Responsibilities:

* Write valid PDF files directly.
* Support pages, page tree, resources, content streams, xref table, trailer.
* Support vector paths, fills, strokes, transforms, clipping.
* Support images as XObjects.
* Support embedded TrueType/OpenType fonts.
* Support Unicode text through Type0/CID fonts and ToUnicode CMaps.
* Keep output deterministic.

Initial PDF version: PDF 1.7.

The PDF writer must not rely on an existing PDF library.

---

## 8. Units and geometry

Implement a single central unit conversion utility.

Required constants:

```csharp
public static class OoxUnits
{
    public const double PointsPerInch = 72.0;
    public const double EmuPerInch = 914400.0;
    public const double TwipsPerPoint = 20.0;

    public static double EmuToPt(long emu) => emu * PointsPerInch / EmuPerInch;
    public static double TwipsToPt(long twips) => twips / TwipsPerPoint;
    public static double HalfPointsToPt(long halfPoints) => halfPoints / 2.0;
}
```

All internal layout should use `double` points.

---

## 9. Font system

Font handling is critical for visual correctness.

### 9.1 Requirements

Implement `Lokad.OoxPdf.Fonts` with:

* system font discovery
* font family/style matching
* OOXML theme font resolution
* fallback font selection
* TrueType/OpenType table parsing
* glyph advance measurement
* simple Latin text shaping
* PDF font embedding
* ToUnicode CMap generation

### 9.2 Minimum font table support

The font parser must support at least:

* `cmap`
* `head`
* `hhea`
* `hmtx`
* `maxp`
* `name`
* `OS/2`
* `post`

Optional but important:

* `kern`
* `GPOS`
* `GSUB`

Phase 1 may ignore full OpenType shaping, but it must emit diagnostics when complex shaping is likely required.

### 9.3 Font fallback policy

Font resolution order:

1. Embedded font in OOXML package, if legally and technically usable.
2. User-provided `IFontResolver`.
3. OS-installed font.
4. Built-in substitute mapping.
5. Last-resort PDF base font or bundled vector fallback, with warning.

The library must never redistribute Microsoft fonts.

### 9.4 Text measurement

Initial implementation:

* map Unicode scalar to glyph ID through `cmap`
* use `hmtx` advance widths
* apply font size and character spacing
* support bold/italic only if matching face exists; synthetic bold/italic is optional and must be diagnostic-backed
* implement greedy line breaking
* implement tabs and bullet indentation for common cases

Known limitation: complex scripts, ligatures, bidirectional text, emoji, and advanced OpenType shaping may be incomplete at first.

---

## 10. Image system

Namespace: `Lokad.OoxPdf.Imaging`

Required support:

* JPEG: read dimensions, embed original bytes as PDF `/DCTDecode`.
* PNG: parse IHDR/IDAT/IEND, support 8-bit RGB and RGBA, support alpha as PDF soft mask.
* PNG color types required initially: 2 and 6.
* PNG color types to add next: 0, 3, 4.
* SVG: phase 2 subset parser for common logos and icons.
* GIF: static first-frame support optional.
* BMP, TIFF, EMF, WMF: unsupported initially unless fallback image is present.

For unsupported image types, render a visible placeholder only in debug mode; otherwise omit with diagnostic. Production default should not draw ugly placeholders unless configured.

---

## 11. PPTX rendering requirements

Namespace: `Lokad.OoxPdf.Pptx`

### 11.1 Package parts

Required parsing:

* presentation part
* slide list
* slide size
* slide parts
* slide relationships
* slide layout parts
* slide master parts
* theme parts
* image/media parts
* notes ignored
* comments ignored

### 11.2 Inheritance and cascade

Implement the PowerPoint inheritance chain:

```text
theme
  -> slide master
    -> slide layout
      -> slide
```

For each slide:

1. Resolve background.
2. Resolve master/layout placeholders.
3. Resolve shape tree.
4. Preserve z-order.
5. Apply group transforms.
6. Apply theme colors and fonts.
7. Render final static slide only.

### 11.3 Slide size

Use presentation slide size from `ppt/presentation.xml`.

PDF page size must match slide size in points.

Visual reference PNG size should use:

```text
pixelWidth  = round(slideWidthPt  * dpi / 72)
pixelHeight = round(slideHeightPt * dpi / 72)
```

### 11.4 Shapes

P0 support:

* rectangle
* rounded rectangle
* ellipse
* line
* arrow line
* triangle
* diamond
* basic freeform paths where explicitly present
* solid fill
* no fill
* solid stroke
* stroke width
* dash style
* transparency
* rotation
* flip horizontal/vertical

P1 support:

* gradients
* shadows, approximated
* common preset geometries
* picture fills
* crop and stretch
* grouped shapes

P2 support:

* glow/reflection/soft edges approximations
* 3D effects ignored with diagnostic
* complex custom geometries

### 11.5 Text boxes

P0 support:

* paragraphs
* runs
* font family, size, bold, italic, underline
* font color
* paragraph alignment
* line spacing
* text box margins
* word wrapping
* vertical anchoring: top, middle, bottom
* bullets with common bullet characters
* numbered lists approximated

P1 support:

* tabs
* hanging indents
* autofit text
* vertical text
* text rotation
* superscript/subscript

### 11.6 Tables

P0 support:

* table grid
* merged cells where straightforward
* cell fills
* borders
* text in cells
* row heights and column widths

P1 support:

* theme table styles
* banded rows/columns
* per-side border precedence

### 11.7 Charts

Corporate decks often contain charts. The implementation must treat charts as a first-class roadmap item.

P0:

* detect chart parts
* emit diagnostics
* render fallback if a cached image or embedded preview is present

P1:

* render basic bar, column, line, and pie charts from cached chart data
* support title, legend, axes, labels, series colors

P2:

* stacked charts
* combo charts
* secondary axes
* data labels and advanced formatting

### 11.8 Unsupported PPTX features

These must not crash conversion:

* animations
* transitions
* videos
* audio
* macros
* ActiveX
* embedded OLE live rendering
* SmartArt without fallback
* embedded Excel live recalculation
* comments
* speaker notes

Each must produce a diagnostic if encountered.

---

## 12. DOCX rendering requirements

Namespace: `Lokad.OoxPdf.Docx`

Word rendering is harder than slide rendering because pagination is implicit. The implementation must be incremental and validation-driven.

### 12.1 Package parts

Required parsing:

* main document part
* styles
* numbering
* settings
* theme
* font table
* headers
* footers
* footnotes/endnotes initially ignored or approximated
* image parts
* relationships

Microsoft’s WordprocessingML documentation notes that a `.docx` can be inspected as a ZIP package and that the main document part contains the body content. ([Microsoft Learn][7])

### 12.2 Page setup

Support:

* section properties
* page size
* orientation
* margins
* header/footer distance
* columns initially single-column only
* page breaks
* section breaks approximated

### 12.3 Paragraphs

P0 support:

* paragraph style resolution
* run style resolution
* document defaults
* font family and size
* bold, italic, underline
* text color
* alignment
* spacing before/after
* line spacing
* left/right/first-line/hanging indent
* page breaks before
* manual line breaks

P1 support:

* keep with next
* keep lines together
* widow/orphan control
* tabs
* borders and shading
* drop caps ignored with diagnostic

### 12.4 Numbering and bullets

P0:

* unordered bullets
* decimal numbering
* indentation based on numbering level

P1:

* roman numerals
* letters
* restart rules
* multi-level numbering formats

### 12.5 Tables

P0:

* fixed-width tables
* rows
* cells
* grid columns
* cell margins
* cell fill
* borders
* horizontal merge
* vertical merge
* text layout inside cells

P1:

* autofit approximation
* repeating header rows
* nested tables
* table styles

### 12.6 Images and drawings

P0:

* inline images
* image sizing
* aspect ratio preservation
* simple anchored images when relative to page or margin

P1:

* floating text wrapping
* behind/in-front-of-text layering
* crop
* rotation

### 12.7 Headers and footers

P0:

* default header
* default footer
* page number fields approximated
* first page header/footer
* odd/even headers optional

### 12.8 Fields

P0:

* page number
* total pages after layout pass
* date fields as static if cached text exists

P1:

* TOC from cached text
* hyperlinks visual text only

### 12.9 Unsupported DOCX features

These must not crash conversion:

* macros
* tracked changes beyond accepting final visible text
* comments
* complex fields without cached text
* equations initially as placeholder unless DrawingML fallback exists
* embedded OLE live content
* content controls beyond visible text
* complex scripts and bidirectional text until shaping is implemented
* multi-column layout until implemented
* footnotes/endnotes until implemented

Each unsupported feature must produce diagnostics.

---

## 13. PDF output requirements

### 13.1 Basic PDF structure

The writer must emit:

* header
* indirect objects
* catalog
* pages tree
* page objects
* content streams
* resource dictionaries
* font objects
* image XObjects
* xref table
* trailer
* `startxref`

### 13.2 Graphics

Support:

* `q` / `Q` graphics state save/restore
* `cm` transforms
* path construction
* fill/stroke
* clipping
* line width
* line caps and joins
* dash patterns
* RGB colors
* alpha through ExtGState

### 13.3 Text

Support:

* embedded TrueType/OpenType fonts
* Unicode text
* ToUnicode CMaps
* text positioning
* character spacing
* line placement
* clipping within text boxes

Do not rely on PDF base fonts except as emergency fallback.

### 13.4 Images

Support:

* JPEG passthrough
* PNG decoded to RGB/alpha image XObjects
* alpha masks
* image scaling and clipping
* crop rectangles

### 13.5 Metadata

Default deterministic metadata:

* `/Producer (Lokad.OoxPdf)`
* no creation date unless supplied
* no author/title unless explicitly supplied later

---

## 14. Diagnostics

Diagnostics are a core feature, not an afterthought. They allow the coding agent to identify what to implement next.

Diagnostic severity:

```csharp
public enum OoxPdfSeverity
{
    Info,
    Warning,
    Error
}
```

Required diagnostic fields:

* `Id`: stable machine-readable code, e.g. `PPTX_UNSUPPORTED_SMARTART`
* `Severity`
* `Message`
* `PartName`
* `XPath`
* `PageIndex`
* `SlideIndex`
* `Feature`
* `Fallback`

Examples:

```json
{
  "id": "PPTX_UNSUPPORTED_VIDEO",
  "severity": "Warning",
  "message": "Video media was ignored; static PDF output does not support video playback.",
  "partName": "/ppt/slides/slide4.xml",
  "slideIndex": 3,
  "feature": "video",
  "fallback": "ignored"
}
```

```json
{
  "id": "DOCX_FLOATING_IMAGE_WRAP_APPROX",
  "severity": "Warning",
  "message": "Floating image wrap type 'tight' was approximated as square.",
  "partName": "/word/document.xml",
  "pageIndex": 5,
  "feature": "floating-image-wrap",
  "fallback": "square-wrap"
}
```

Rules:

* Unsupported features must be diagnostic-visible.
* Diagnostics must not spam duplicate messages; aggregate repeated identical feature warnings where possible.
* Strict mode must fail if any warning or error is emitted.
* Non-strict mode should produce the best possible PDF.

---

## 15. Visual validation harness

The validation harness is mandatory. Its purpose is to let the coding agent compare Office reference output with `Lokad.OoxPdf` output visually.

### 15.1 Reference rendering strategy

For `.pptx`:

```text
PowerPoint COM -> slide PNG images
```

PowerPoint exposes `Presentation.Export` and `Slide.Export` methods that export slides with a graphics filter and optional pixel dimensions; Microsoft’s docs show `FilterName := "png"` and `ScaleWidth` / `ScaleHeight` parameters. ([Microsoft Learn][8])

For `.docx`:

```text
Word COM -> reference PDF -> PDFium -> page PNG images
```

Word’s `Document.ExportAsFixedFormat` saves a document as PDF or XPS, and `wdExportFormatPDF` has value `17`. ([Microsoft Learn][9])

For candidate output:

```text
Lokad.OoxPdf -> candidate PDF -> PDFium -> page PNG images
```

PDFium’s standalone `pdfium_test` supports rasterizing PDF pages to PNG or PPM output, and its help includes `--png`, `--ppm`, and `--scale=<number>`. ([PDFium][10])

### 15.2 Visual case artifact layout

Each visual run must produce:

```text
artifacts/visual/<case-id>/<run-id>/
  input/
    original.pptx|original.docx
  reference/
    page-001.png
    page-002.png
    ...
    reference.pdf              # docx only, optional for pptx
    manifest.json
  candidate/
    output.pdf
    page-001.png
    page-002.png
    ...
    diagnostics.json
  comparison/
    index.html
    page-001-side-by-side.png
    page-001-diff.png
    page-001-report.json
    ...
    metrics.json
  assessment.md
```

The coding agent should inspect `comparison/index.html` and the `page-###-side-by-side.png` files.

### 15.3 Visual case manifest

Each case must have a manifest:

```json
{
  "id": "sample-pptx-basic",
  "kind": "pptx",
  "input": "../../../corpus/pptx/sample-basic.pptx",
  "dpi": 144,
  "tags": ["pptx", "text", "shapes", "images"],
  "expected": {
    "minAgentRating": 4,
    "pageCountMustMatch": true,
    "dimensionsMustMatch": true
  },
  "allowedUnsupportedFeatures": [
    "animation",
    "speaker-notes"
  ]
}
```

### 15.4 Visual comparison metrics

Metrics are advisory at first. The agent’s visual judgment is authoritative.

For each page:

```json
{
  "page": 1,
  "referenceWidth": 1920,
  "referenceHeight": 1080,
  "candidateWidth": 1920,
  "candidateHeight": 1080,
  "meanAbsoluteError": 2.71,
  "rootMeanSquaredError": 7.84,
  "changedPixelRatioAtThreshold16": 0.0182,
  "changedPixelRatioAtThreshold32": 0.0061,
  "dimensionsMatch": true
}
```

Metric interpretation:

* DOCX comparisons can be stricter because both reference and candidate PDFs are rasterized by PDFium.
* PPTX comparisons are looser because reference images come directly from PowerPoint while candidate images come from PDFium rasterization of OoxPdf’s PDF.
* Page count and dimensions are hard gates.
* Pixel metrics are trend indicators, not proof of correctness.

### 15.5 Agent visual rating rubric

The coding agent must record a visual rating in `assessment.md`.

```text
5 = visually indistinguishable except tiny antialiasing differences
4 = good; minor spacing, kerning, or antialiasing differences only
3 = usable; visible but non-blocking differences
2 = poor; major layout defects or missing important elements
1 = severe; page/slide mostly wrong
0 = conversion failed or page missing
```

Required `assessment.md` template:

```markdown
# Visual assessment: <case-id>

Input: <file>
Kind: pptx|docx
Run: <run-id>
Agent rating: <0-5>

## Summary

<One paragraph summary of visual fidelity.>

## Pages/slides reviewed

- Page 1: <rating>, <main differences>
- Page 2: <rating>, <main differences>

## Major defects

1. <Defect, page, suspected feature>
2. <Defect, page, suspected feature>

## Diagnostics reviewed

<Note important warnings/errors from diagnostics.json.>

## Next implementation target

<The smallest renderer improvement likely to improve this case.>
```

### 15.6 Visual gates

A visual case passes when:

* conversion succeeds
* page/slide count matches
* page/slide dimensions match
* no unexpected diagnostics with severity `Error`
* agent rating is at least `expected.minAgentRating`
* known allowed unsupported features are the only relevant unsupported features

---

## 16. Required PowerShell scripts

The scripts below are implementation templates. The coding agent may refactor them, but the behavior and outputs must remain.

### 16.1 `tools/RenderReference.ps1`

Purpose: produce Office reference images.

```powershell
param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath,

    [Parameter(Mandatory = $true)]
    [string] $OutputDir,

    [ValidateSet("Auto", "Pptx", "Docx")]
    [string] $Kind = "Auto",

    [int] $Dpi = 144
)

$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object] $Object)
    if ($null -ne $Object) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

$fullInput = (Resolve-Path $InputPath).Path
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

if ($Kind -eq "Auto") {
    $ext = [System.IO.Path]::GetExtension($fullInput).ToLowerInvariant()
    if ($ext -eq ".pptx" -or $ext -eq ".pptm") {
        $Kind = "Pptx"
    } elseif ($ext -eq ".docx" -or $ext -eq ".docm") {
        $Kind = "Docx"
    } else {
        throw "Unsupported extension: $ext"
    }
}

if ($Kind -eq "Pptx") {
    $powerPoint = $null
    $presentation = $null

    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application

        # Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
        $presentation = $powerPoint.Presentations.Open($fullInput, $true, $false, $false)

        $widthPt = [double] $presentation.PageSetup.SlideWidth
        $heightPt = [double] $presentation.PageSetup.SlideHeight
        $widthPx = [int] [Math]::Round($widthPt * $Dpi / 72.0)
        $heightPx = [int] [Math]::Round($heightPt * $Dpi / 72.0)

        $slideCount = [int] $presentation.Slides.Count

        for ($i = 1; $i -le $slideCount; $i++) {
            $slide = $presentation.Slides.Item($i)
            $file = Join-Path $OutputDir ("page-{0:D3}.png" -f $i)
            $slide.Export($file, "PNG", $widthPx, $heightPx)
            Release-ComObject $slide
        }

        @{
            kind = "pptx"
            source = $fullInput
            dpi = $Dpi
            slideCount = $slideCount
            widthPt = $widthPt
            heightPt = $heightPt
            widthPx = $widthPx
            heightPx = $heightPx
        } | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 (Join-Path $OutputDir "manifest.json")
    }
    finally {
        if ($null -ne $presentation) {
            $presentation.Close()
            Release-ComObject $presentation
        }
        if ($null -ne $powerPoint) {
            $powerPoint.Quit()
            Release-ComObject $powerPoint
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    exit 0
}

if ($Kind -eq "Docx") {
    $word = $null
    $document = $null

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        # Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        $document = $word.Documents.Open($fullInput, $false, $true, $false)

        $pdfPath = Join-Path $OutputDir "reference.pdf"
        $wdExportFormatPDF = 17

        $document.ExportAsFixedFormat($pdfPath, $wdExportFormatPDF)

        & (Join-Path $PSScriptRoot "RasterizePdf.ps1") `
            -PdfPath $pdfPath `
            -OutputDir $OutputDir `
            -Dpi $Dpi `
            -Prefix "page"

        @{
            kind = "docx"
            source = $fullInput
            dpi = $Dpi
            referencePdf = $pdfPath
        } | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 (Join-Path $OutputDir "manifest.json")
    }
    finally {
        if ($null -ne $document) {
            $document.Close($false)
            Release-ComObject $document
        }
        if ($null -ne $word) {
            $word.Quit()
            Release-ComObject $word
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    exit 0
}
```

### 16.2 `tools/RasterizePdf.ps1`

Purpose: convert PDF pages to PNG using vendored PDFium.

```powershell
param(
    [Parameter(Mandatory = $true)]
    [string] $PdfPath,

    [Parameter(Mandatory = $true)]
    [string] $OutputDir,

    [int] $Dpi = 144,

    [string] $Prefix = "page"
)

$ErrorActionPreference = "Stop"

$fullPdf = (Resolve-Path $PdfPath).Path
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$pdfium = Join-Path $PSScriptRoot "vendor/pdfium/win-x64/pdfium_test.exe"
if (-not (Test-Path $pdfium)) {
    throw "Missing PDFium binary: $pdfium"
}

$scale = [Math]::Round($Dpi / 72.0, 6)

# pdfium_test writes files next to the PDF as <pdf-name>.<zero-based-page>.png
Push-Location (Split-Path $fullPdf)
try {
    & $pdfium --png "--scale=$scale" $fullPdf
    if ($LASTEXITCODE -ne 0) {
        throw "pdfium_test failed with exit code $LASTEXITCODE"
    }
}
finally {
    Pop-Location
}

$escaped = [Regex]::Escape([System.IO.Path]::GetFileName($fullPdf))
$sourceDir = Split-Path $fullPdf
$files = Get-ChildItem $sourceDir -File |
    Where-Object { $_.Name -match "^$escaped\.(\d+)\.png$" } |
    Sort-Object {
        if ($_.Name -match "\.(\d+)\.png$") { [int] $Matches[1] } else { 0 }
    }

$page = 1
foreach ($file in $files) {
    $target = Join-Path $OutputDir ("{0}-{1:D3}.png" -f $Prefix, $page)
    Move-Item -Force $file.FullName $target
    $page++
}

@{
    pdf = $fullPdf
    dpi = $Dpi
    scale = $scale
    pageCount = $files.Count
} | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 (Join-Path $OutputDir "rasterize-manifest.json")
```

### 16.3 `tools/CheckVisualCase.ps1`

Purpose: run the full comparison.

Required behavior:

```powershell
pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/sample-pptx-basic/case.json
```

Pipeline:

1. Read `case.json`.
2. Copy or reference input.
3. Run `RenderReference.ps1`.
4. Run `Lokad.OoxPdf.Cli` to produce candidate PDF and diagnostics.
5. Run `RasterizePdf.ps1` for candidate PDF.
6. Run `Lokad.OoxPdf.VisualDiff`.
7. Create `comparison/index.html`.
8. Create `assessment.md` from template.
9. Exit nonzero if conversion fails, page count mismatches, or dimensions mismatch.

---

## 17. VisualDiff tool

Namespace/project: `tools/Lokad.OoxPdf.VisualDiff`

This tool must remain dependency-free.

Required command:

```powershell
dotnet run --project tools/Lokad.OoxPdf.VisualDiff -- `
  --reference artifacts/visual/case/run/reference `
  --candidate artifacts/visual/case/run/candidate `
  --output artifacts/visual/case/run/comparison
```

Required outputs:

* `metrics.json`
* one side-by-side PNG per page
* one diff PNG per page
* `index.html`

Implementation options:

1. Implement a minimal PNG reader/writer in C# using standard libraries.
2. Reuse the core library’s PNG parser/writer if cleanly separated.
3. If PNG writing is not ready, emit `index.html` with image tags first; side-by-side PNGs and diff PNGs are still required before release.

Diff image policy:

* Same-size images: compute per-pixel absolute RGB difference.
* Different-size images: create a canvas large enough for both and mark mismatch in `page-###-report.json`.
* Amplify visible differences by a factor of 4 for diff images.
* Ignore alpha for metric computation unless alpha differs visibly.

---

## 18. Agent development workflow

The coding agent must follow this loop for each feature:

1. Select a visual case or create a small synthetic case.
2. Run the visual harness.
3. Inspect `comparison/index.html` and side-by-side PNGs.
4. Identify the first major visible divergence.
5. Implement the smallest renderer improvement that addresses it.
6. Re-run the same visual case.
7. Update `assessment.md`.
8. Add or update unit-level parser/layout tests.
9. Update `docs/Capabilities.md`.

The agent must not mark a feature complete unless:

* it has diagnostics for unsupported branches
* it has at least one visual case
* it does not regress existing visual cases
* it does not add any package dependency
* it does not call Office or PDFium from library code

---

## 19. Test strategy

### 19.1 Unit tests

Because the repo is dependency-free, tests may be implemented as a simple console test runner instead of xUnit/NUnit.

Required areas:

* ZIP package path resolution
* relationship resolution
* content type parsing
* XML namespace parsing
* unit conversions
* theme color resolution
* font table parsing
* basic text measurement
* PDF object writing
* PNG/JPEG parsing
* PPTX shape parsing
* DOCX paragraph style cascade
* DOCX numbering resolution

### 19.2 Golden structure tests

These tests inspect produced PDF structure without rasterization:

* PDF has expected page count.
* Page sizes match expected points.
* Font objects exist for text.
* Image XObjects exist for embedded images.
* Diagnostics match expected warnings.

### 19.3 Visual tests

Visual tests are the primary quality driver.

Corpus categories:

```text
visual-cases/
  cases/
    pptx-basic-text/
    pptx-shapes/
    pptx-images/
    pptx-tables/
    pptx-corporate-theme/
    pptx-chart-basic/
    docx-basic-paragraphs/
    docx-tables/
    docx-images/
    docx-headers-footers/
    docx-numbering/
    docx-corporate-report/
```

Case tiers:

* **Bronze:** small synthetic cases, one feature per document.
* **Silver:** realistic documents with multiple common features.
* **Gold:** representative corporate documents that define release quality.

Private or confidential documents must not be committed. Use anonymized or synthetic equivalents.

---

## 20. Security requirements

The converter must treat input as untrusted.

Required protections:

* Disable XML DTD processing.
* Do not resolve external XML entities.
* Cap uncompressed ZIP size.
* Cap number of ZIP entries.
* Cap individual part size.
* Normalize and validate relationship targets.
* Disable external relationships by default.
* Reject absolute file paths unless explicitly enabled.
* Reject network URLs unless explicitly enabled in a future feature.
* Avoid recursive XML traversal without depth limits.
* Avoid integer overflow in unit and image calculations.
* Avoid allocating based solely on untrusted dimensions.
* Emit diagnostics for ignored external resources.

Default behavior: render only resources inside the OOXML package.

---

## 21. Performance requirements

Initial performance target:

* PPTX: render a 50-slide corporate deck in under 30 seconds on a typical developer workstation after warm font cache.
* DOCX: render a 50-page report in under 30 seconds after warm font cache.
* Memory: avoid loading all decoded images for large documents simultaneously.
* Streaming: output PDF to a stream.
* Caching: cache font metadata per process.
* Threading: conversion objects should be independent; no mutable global conversion state.

These targets are not release blockers for the first prototype, but architectural choices must not make them impossible.

---

## 22. Capability documentation

`docs/Capabilities.md` must be updated continuously.

Required format:

```markdown
# Lokad.OoxPdf capabilities

## PPTX

| Feature | Status | Notes | Visual cases |
|---|---:|---|---|
| Slide size | Supported | Uses presentation page setup | pptx-basic-text |
| Text runs | Partial | Latin text only | pptx-basic-text |
| Images PNG/JPEG | Supported | Alpha for PNG | pptx-images |
| Charts | Partial | Basic bar/line planned | pptx-chart-basic |
| Animations | Ignored | Static final slide only | n/a |

## DOCX

| Feature | Status | Notes | Visual cases |
|---|---:|---|---|
| Paragraphs | Partial | Basic flow layout | docx-basic-paragraphs |
| Tables | Partial | Fixed-width tables | docx-tables |
| Headers/footers | Partial | Default header/footer | docx-headers-footers |
| Floating images | Partial | Limited wrapping | docx-images |
```

Status values:

* `Supported`
* `Partial`
* `Approximated`
* `Ignored`
* `Unsupported`

---

## 23. Implementation phases

### Phase 0 — skeleton and guardrails

Deliverables:

* solution structure
* dependency-free csproj
* core public API
* CLI
* diagnostics model
* package reader
* PDF writer with blank pages
* visual harness scripts
* PDFium vendoring instructions
* first PPTX and DOCX visual cases

Acceptance:

* `dotnet build` succeeds
* `dotnet pack` produces a package with no dependencies
* visual harness can create Office references
* CLI can create a blank-page PDF with correct page count and dimensions

### Phase 1 — PPTX basic renderer

Deliverables:

* slide size
* slide backgrounds
* text boxes
* basic shapes
* PNG/JPEG images
* solid fills/strokes
* group transforms
* theme colors/fonts
* diagnostics for unsupported features

Acceptance:

* bronze PPTX cases rating ≥ 4
* no crash on silver PPTX cases
* page dimensions match Office references

### Phase 2 — DOCX basic renderer

Deliverables:

* page setup
* paragraphs
* runs
* basic font metrics
* line breaking
* page breaking
* tables
* inline images
* headers/footers
* numbering basics

Acceptance:

* bronze DOCX cases rating ≥ 4
* no crash on silver DOCX cases
* page count matches simple documents
* diagnostics explain unsupported layout features

### Phase 3 — corporate fidelity

Deliverables:

* better font embedding/subsetting
* PPTX tables and charts
* DOCX table styles
* floating images
* gradients and shadows approximation
* SVG subset
* improved visual diff

Acceptance:

* silver cases rating ≥ 4
* gold cases rating ≥ 3
* no unexplained missing major elements

### Phase 4 — release candidate

Deliverables:

* stable API review
* NuGet metadata
* README examples
* security review
* performance pass
* deterministic output tests
* corpus summary

Acceptance:

* no library dependencies
* gold cases rating ≥ 4 except documented unsupported cases
* all unsupported features are diagnostic-visible
* `dotnet pack -c Release` creates the publishable NuGet package

---

## 24. Definition of done

A pull request is done only when:

1. It builds without adding any package dependency to `src/Lokad.OoxPdf`.
2. It includes unit or structure tests for parser/layout logic.
3. It includes or updates at least one visual case when visual output changes.
4. It updates diagnostics for unsupported or approximated behavior.
5. It updates `docs/Capabilities.md`.
6. It produces no unexpected regression in existing visual cases.
7. It keeps Office/PDFium usage confined to `tools/`.

A release is done only when:

1. `Lokad.OoxPdf` converts representative `.pptx` and `.docx` files without crashing.
2. The NuGet package has no dependencies.
3. Visual validation artifacts exist for the release corpus.
4. Known limitations are documented.
5. The public API is stable.
6. The package can be consumed by a separate sample project.

---

## 25. Minimal README example

The repository README must contain this style of example:

```csharp
using Lokad.OoxPdf;

OoxPdfConverter.Convert(
    inputPath: "quarterly-report.docx",
    outputPdfPath: "quarterly-report.pdf",
    options: new OoxPdfOptions
    {
        DiagnosticSink = diagnostic =>
        {
            Console.Error.WriteLine($"{diagnostic.Severity} {diagnostic.Id}: {diagnostic.Message}");
        }
    });
```

CLI example:

```powershell
dotnet run --project src/Lokad.OoxPdf.Cli -- convert deck.pptx deck.pdf --diagnostics deck.diagnostics.json
```

Visual validation example:

```powershell
pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-basic-text/case.json
```

---

## 26. Critical engineering guidance for the coding agent

Start with vertical slices. Do not attempt full OOXML support before producing PDFs and visual comparisons.

Recommended first vertical slice:

1. Read `.pptx`.
2. Resolve slide count and slide size.
3. Emit one blank PDF page per slide.
4. Render slide background.
5. Render one text box.
6. Render one rectangle.
7. Add visual case.
8. Compare against PowerPoint PNG.
9. Iterate.

Recommended second vertical slice:

1. Read `.docx`.
2. Resolve page size and margins.
3. Lay out plain paragraphs.
4. Emit PDF pages.
5. Add Word/PDFium reference comparison.
6. Iterate.

The main measure of progress is not the number of OOXML elements parsed. The main measure is improved visual agreement on representative documents, with clear diagnostics for everything not yet supported.

[1]: https://ecma-international.org/publications-and-standards/standards/ecma-376/ "ECMA-376 - Ecma International"
[2]: https://www.loc.gov/preservation/digital/formats/fdd/fdd000395.shtml "OOXML Format Family -- ISO/IEC 29500 and ECMA 376"
[3]: https://learn.microsoft.com/en-us/dotnet/api/system.io.compression.ziparchive?view=net-10.0 "ZipArchive Class (System.IO.Compression) | Microsoft Learn"
[4]: https://learn.microsoft.com/en-us/dotnet/core/compatibility/core-libraries/6.0/system-drawing-common-windows-only "Breaking change: System.Drawing.Common only supported on Windows - .NET | Microsoft Learn"
[5]: https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa "Considerations for unattended automation of Office in the Microsoft 365 for unattended RPA environment | Microsoft Learn"
[6]: https://learn.microsoft.com/en-us/dotnet/standard/library-guidance/nuget "NuGet and .NET libraries - .NET | Microsoft Learn"
[7]: https://learn.microsoft.com/en-us/office/open-xml/word/structure-of-a-wordprocessingml-document "Structure of a WordprocessingML document | Microsoft Learn"
[8]: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.export "Presentation.Export method (PowerPoint) | Microsoft Learn"
[9]: https://learn.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat "Document.ExportAsFixedFormat method (Word) | Microsoft Learn"
[10]: https://pdfium.googlesource.com/pdfium/%2B/master/README.md "PDFium - PDFium"
