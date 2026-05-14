# Build Lokad.OoxPdf as a dependency-free OOXML-to-PDF renderer with visual validation

This ExecPlan is a living document. The sections `Progress`, `Surprises & Discoveries`, `Decision Log`, and `Outcomes & Retrospective` must be kept up to date as work proceeds.

If a `PLANS.md` file exists in the repository, maintain this document in accordance with that file. This ExecPlan is self-contained and should be sufficient for a coding agent or novice contributor to continue from a clean working tree.

## Purpose / Big Picture

The goal is to create `Lokad.OoxPdf`, a C#/.NET library that converts `.pptx` and `.docx` files into `.pdf` without runtime dependencies beyond the .NET standard libraries. After this work, a user should be able to reference the library from a .NET project, call `OoxPdfConverter.Convert("input.pptx", "output.pdf")` or `OoxPdfConverter.Convert("input.docx", "output.pdf")`, and receive a valid static PDF.

The project also needs a visual validation harness. A coding agent running on Windows with Office 365 installed must be able to render an Office-produced reference image, render the `Lokad.OoxPdf` candidate PDF to images, compare both outputs, and inspect side-by-side PNGs. This matters because PDF rendering quality is visual: compilation alone does not prove that corporate slides and documents look right.

This plan intentionally starts with a minimal vertical slice that produces valid PDFs, then adds PPTX and DOCX support incrementally. Each progress item is designed to be small enough to deserve its own commit.

## Progress

- [x] (2026-05-14) ExecPlan authored from the product requirements and validation strategy.
- [x] (2026-05-14) Create repository skeleton with `src/`, `tests/`, `tools/`, `docs/`, `visual-cases/`, and `artifacts/`.
- [x] (2026-05-14) Add `.gitignore`, `Directory.Build.props`, `.slnx` solution file, README placeholder, LICENSE placeholder, and NuGet metadata placeholders.
- [x] (2026-05-14) Create `src/Lokad.OoxPdf/Lokad.OoxPdf.csproj` targeting `net10.0` with no `PackageReference`.
- [x] (2026-05-14) Create `src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj` referencing the library project only.
- [x] (2026-05-14) Create `tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj` as a dependency-free console test runner.
- [x] (2026-05-14) Create `tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj` as a dependency-free console tool.
- [x] (2026-05-14) Implement public API shells: `OoxPdfConverter`, `OoxPdfOptions`, `OoxPdfInputKind`, diagnostics types, and font resolver types.
- [x] (2026-05-14) Implement CLI argument parsing for `convert input output`, `--diagnostics`, and `--strict`.
- [x] (2026-05-14) Implement minimal diagnostics collection and JSON writing in the CLI.
- [x] (2026-05-14) Implement test runner infrastructure with assertions and clear pass/fail exit codes.
- [x] (2026-05-14) Add a unit test proving the public API can be called and rejects missing files predictably.
- [x] (2026-05-14) Implement OOXML ZIP opening with central package part lookup.
- [x] (2026-05-14) Add ZIP safety limits for entry count, individual part size, and total uncompressed size.
- [x] (2026-05-14) Implement safe package part path normalization and path traversal rejection.
- [x] (2026-05-14) Implement XML reader settings that disable DTDs and external entity resolution.
- [x] (2026-05-14) Parse `[Content_Types].xml`.
- [x] (2026-05-14) Parse package-level `_rels/.rels`.
- [x] (2026-05-14) Parse part-level relationship files such as `ppt/slides/_rels/slide1.xml.rels`.
- [x] (2026-05-14) Add unit tests for content types, relationships, and relationship target resolution.
- [x] (2026-05-14) Implement central unit conversion helpers for EMU, twips, half-points, inches, and PDF points.
- [x] (2026-05-14) Implement minimal PDF writer that emits one blank page with deterministic objects.
- [x] (2026-05-14) Add PDF structure tests for header, catalog, page tree, xref, trailer, and one page.
- [x] (2026-05-14) Extend PDF writer to support multiple blank pages with specified width and height.
- [x] (2026-05-14) Add structure tests proving page count and page sizes are encoded in produced PDFs.
- [x] (2026-05-14) Implement PPTX package discovery: presentation part, slide list, slide parts, slide size.
- [x] (2026-05-14) Wire PPTX conversion to emit one blank PDF page per slide with correct slide dimensions.
- [x] (2026-05-14) Add a small synthetic PPTX fixture or fixture generator for slide count and size tests.
- [x] (2026-05-14) Implement DOCX package discovery: main document part and basic section page size.
- [x] (2026-05-14) Wire DOCX conversion to emit one blank PDF page with default or discovered page dimensions.
- [x] (2026-05-14) Add a small synthetic DOCX fixture or fixture generator for basic document discovery tests.
- [x] (2026-05-14) Retrieve `pdfium-win-x64.tgz` from bblanchon PDFium binaries and add `tools/RasterizePdf.ps1` using a dependency-free PDFium DLL rasterizer tool.
- [x] (2026-05-14) Add `tools/RenderReference.ps1` for PowerPoint `.pptx` to slide PNG export.
- [x] (2026-05-14) Extend `tools/RenderReference.ps1` for Word `.docx` to reference PDF and PNG rasterization.
- [x] (2026-05-14) Add `tools/CheckVisualCase.ps1` orchestration script.
- [x] (2026-05-14) Implement first minimal `VisualDiff` command that discovers reference and candidate PNGs and writes `metrics.json`.
- [x] (2026-05-14) Extend `VisualDiff` to write `index.html` with side-by-side image tags.
- [x] (2026-05-14) Implement minimal PNG reader for truecolor and truecolor-alpha PNGs in `VisualDiff`.
- [x] (2026-05-14) Implement minimal PNG writer or reusable image output path for generated side-by-side and diff images.
- [x] (2026-05-14) Extend `VisualDiff` to compute dimensions, mean absolute error, root mean squared error, and changed pixel ratios.
- [x] (2026-05-14) Create `visual-cases/README.md` documenting how to run visual validation.
- [x] (2026-05-14) Create first PPTX visual case manifest for a simple slide deck.
- [x] (2026-05-14) Create first DOCX visual case manifest for a simple one-page document.
- [x] (2026-05-14) Validate that a blank-page PPTX conversion can be visually compared against PowerPoint output.
- [x] (2026-05-14) Validate that a blank-page DOCX conversion can be visually compared against Word output.
- [x] (2026-05-14) Implement PDF graphics state, RGB fill, RGB stroke, rectangles, lines, and simple paths.
- [x] (2026-05-14) Add PDF structure tests for drawing operators.
- [x] (2026-05-14) Parse PPTX slide background fill and render solid color backgrounds.
- [x] (2026-05-14) Parse PPTX solid-fill rectangles and render them.
- [x] (2026-05-14) Parse PPTX lines and render them.
- [x] (2026-05-14) Parse PPTX ellipse geometry and render it with Bézier approximation.
- [x] (2026-05-14) Parse PPTX rotation and flip transforms for basic shapes.
- [x] (2026-05-14) Add visual case for PPTX shapes and update assessment template.
- [x] (2026-05-14) Implement font discovery on Windows from `C:\Windows\Fonts` with a deterministic cache.
- [x] (2026-05-14) Implement TrueType/OpenType table directory parsing.
- [x] (2026-05-14) Parse required font tables: `cmap`, `head`, `hhea`, `hmtx`, `maxp`, `name`, `OS/2`, and `post`.
- [x] (2026-05-14) Implement simple Unicode-to-glyph mapping for common Microsoft fonts.
- [x] (2026-05-14) Implement glyph advance measurement from `hmtx`.
- [x] (2026-05-14) Add font parser and measurement unit tests using installed fonts when available and skipped tests when absent.
- [x] (2026-05-14) Implement PDF embedded TrueType/CID font objects sufficient for Unicode text output.
- [x] (2026-05-14) Implement ToUnicode CMap generation for embedded text.
- [x] (2026-05-14) Add PDF text structure tests proving text content uses embedded fonts and ToUnicode maps.
- [x] (2026-05-14) Parse PPTX text boxes, paragraphs, runs, run font size, color, bold, italic, underline, and alignment.
- [x] (2026-05-14) Implement simple Latin line breaking and text placement inside PPTX text boxes.
- [x] (2026-05-14) Render PPTX text boxes into PDF using embedded fonts.
- [x] (2026-05-14) Add PPTX text visual case and compare against PowerPoint.
- [x] (2026-05-14) Parse PPTX theme colors and resolve common scheme colors to RGB.
- [x] (2026-05-14) Parse PPTX theme fonts and use them when a text run asks for theme fonts.
- [x] (2026-05-14) Parse PPTX slide master and slide layout background and shape inheritance for common cases.
- [x] (2026-05-14) Render PPTX layout/master placeholders where visible and not overridden.
- [ ] Add PPTX corporate-theme visual case.
- [x] (2026-05-14) Implement JPEG dimension parsing and PDF JPEG passthrough as `/DCTDecode` image XObjects.
- [x] (2026-05-14) Implement PNG parsing for color types 2 and 6, including alpha as PDF soft mask.
- [x] (2026-05-14) Parse PPTX picture relationships and render JPEG/PNG images.
- [x] (2026-05-14) Implement PPTX image sizing, aspect ratio preservation, and basic cropping.
- [x] (2026-05-14) Add PPTX images visual case.
- [x] (2026-05-14) Parse PPTX grouped shapes and apply nested transforms.
- [x] (2026-05-14) Add PPTX group transform unit and visual tests.
- [x] (2026-05-14) Parse PPTX tables into fills, borders, and text cells for simple grid tables.
- [x] (2026-05-14) Render PPTX tables with fixed row and column geometry.
- [x] (2026-05-14) Add PPTX table visual case.
- [x] (2026-05-14) Detect PPTX charts, SmartArt, videos, audio, OLE objects, transitions, and animations.
- [x] (2026-05-14) Emit stable diagnostics for unsupported PPTX features without crashing.
- [x] (2026-05-14) Add tests proving unsupported PPTX features are diagnostic-visible.
- [x] (2026-05-14) Implement DOCX style part parsing for document defaults, paragraph styles, and character styles.
- [x] (2026-05-14) Parse DOCX body paragraphs and runs with basic run formatting.
- [x] (2026-05-14) Implement DOCX page setup from section properties: page size, margins, and orientation.
- [x] (2026-05-14) Implement DOCX paragraph layout with margins, spacing before/after, alignment, and line spacing.
- [x] (2026-05-14) Implement DOCX text line breaking and page breaking.
- [x] (2026-05-14) Render simple DOCX paragraphs into PDF.
- [x] (2026-05-14) Add DOCX basic paragraphs visual case.
- [x] (2026-05-14) Parse DOCX numbering part for simple bullets and decimal numbering.
- [x] (2026-05-14) Render DOCX bullets and decimal lists.
- [x] (2026-05-14) Add DOCX numbering visual case.
- [x] (2026-05-14) Parse DOCX inline images and render JPEG/PNG images.
- [x] (2026-05-14) Add DOCX inline images visual case.
- [x] (2026-05-14) Parse DOCX fixed-width tables, rows, cells, fills, and text.
- [x] (2026-05-14) Render DOCX fixed-width tables with text inside cells.
- [x] (2026-05-14) Add DOCX table visual case.
- [x] (2026-05-14) Parse DOCX headers and footers and render default header/footer content.
- [x] (2026-05-14) Implement page number field approximation for headers and footers.
- [x] (2026-05-14) Add DOCX headers/footers visual case.
- [x] (2026-05-14) Detect DOCX comments, tracked changes, complex fields, equations, OLE objects, floating wrap types, footnotes, endnotes, multi-column sections, and macros.
- [x] (2026-05-14) Emit stable diagnostics for unsupported DOCX features without crashing.
- [x] (2026-05-14) Add tests proving unsupported DOCX features are diagnostic-visible.
- [x] (2026-05-14) Implement deterministic output mode for PDF object ordering, resource names, and metadata.
- [x] (2026-05-14) Add deterministic-output test that converts the same fixture twice and compares bytes.
- [x] (2026-05-14) Implement strict mode behavior where warnings or errors produce CLI exit code 3.
- [x] (2026-05-14) Add tests for CLI exit codes 0, 1, 2, and 3.
- [ ] Add `docs/Capabilities.md` and update it with all implemented PPTX and DOCX capabilities.
- [ ] Add `docs/RenderingModel.md` explaining package, layout, scene, and PDF layers.
- [ ] Add `docs/Diagnostics.md` with diagnostic code conventions.
- [ ] Add `docs/VisualValidation.md` explaining Office reference rendering, PDFium rasterization, and agent assessment.
- [ ] Add README quick start with library, CLI, and visual validation examples.
- [x] (2026-05-14) Add NuGet package metadata and verify `dotnet pack -c Release` creates a dependency-free package.
- [ ] Run full build and tests from clean checkout.
- [x] (2026-05-14) Run at least one PPTX visual case and record `assessment.md`.
- [x] (2026-05-14) Run at least one DOCX visual case and record `assessment.md`.
- [ ] Complete first retrospective entry with remaining gaps and next implementation priorities.

## Surprises & Discoveries

- Observation: The .NET 10 SDK in this workspace creates `Lokad.OoxPdf.slnx` for `dotnet new sln -n Lokad.OoxPdf`, not `Lokad.OoxPdf.sln`.
  Evidence: Running `dotnet new sln -n Lokad.OoxPdf` produced `Lokad.OoxPdf.slnx`; `dotnet sln Lokad.OoxPdf.slnx add ...` accepted all four projects, and `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` succeeded with 0 warnings and 0 errors.

- Observation: The initial dependency-free test runner and public API shell build on the installed .NET 10 SDK.
  Evidence: `dotnet run --project tests/Lokad.OoxPdf.Tests --tl:off --no-build` printed `PASS PublicApiRejectsMissingInput`, `PASS AutoDetectsPptxExtension`, `PASS AutoDetectsDocxExtension`, and `3 passed, 0 failed`.

- Observation: The first OOXML package layer can be tested with synthetic ZIP packages rather than checked-in Office binaries.
  Evidence: `OoxmlTests.ParsesContentTypesAndRelationships` builds an in-memory ZIP with `[Content_Types].xml`, `_rels/.rels`, and presentation parts; the full test run now prints `7 passed, 0 failed`.

- Observation: The minimal PDF writer can produce deterministic single-page and multi-page blank PDFs without external libraries.
  Evidence: `PdfWriterTests.WritesSingleBlankPagePdfStructure` and `PdfWriterTests.WritesMultipleBlankPagesWithPageSizes` validate header, catalog, page tree, xref, trailer, page count, and media boxes; the full test run now prints `9 passed, 0 failed`.

- Observation: The first PPTX and DOCX conversions can be validated without Office by generating minimal OOXML ZIP packages in tests.
  Evidence: `PptxSyntheticTwoSlidesProducesTwoPdfPages` verifies a two-slide synthetic PPTX produces two PDF pages with a 960 by 540 media box, and `DocxSyntheticDocumentProducesOnePdfPage` verifies a synthetic DOCX produces one 612 by 792 point page; the full test run now prints `11 passed, 0 failed`.

- Observation: The minimal visual comparison command can create browsable artifacts before pixel-level PNG parsing exists.
  Evidence: Running `dotnet tools/Lokad.OoxPdf.VisualDiff/bin/Debug/net10.0/Lokad.OoxPdf.VisualDiff.dll <reference> <candidate> <comparison>` against temporary PNG filenames wrote `metrics.json` and `index.html`.

- Observation: VisualDiff can now read ordinary RGB/RGBA PNGs and compute pixel metrics without `System.Drawing.Common` or third-party packages.
  Evidence: A smoke run using two temporary 2 by 1 PNGs wrote `ReferenceWidth: 2`, `ReferenceHeight: 1`, `MeanAbsoluteError: 0`, `RootMeanSquaredError: 0`, and `DimensionsMatch: true` in `metrics.json`.

- Observation: The renderer now emits non-blank PDF graphics for simple PPTX slides.
  Evidence: `PptxSyntheticShapesProduceDrawingOperators` converts a synthetic slide with a white background, red rectangle with blue stroke, green ellipse, and black line, then verifies the expected PDF drawing operators; the full test run now prints `13 passed, 0 failed`.

- Observation: The requested bblanchon PDFium archive provides `bin/pdfium.dll`, headers, libraries, and license files, but not `pdfium_test.exe`.
  Evidence: After extracting `https://github.com/bblanchon/pdfium-binaries/releases/latest/download/pdfium-win-x64.tgz`, `tools/vendor/pdfium/win-x64/bin/pdfium.dll` exists and no `.exe` file is present. A new `tools/Lokad.OoxPdf.PdfiumRasterizer` console tool P/Invokes the DLL and successfully produced `page-001.png` from a small hand-written PDF.

- Observation: PowerPoint exports simple blank slides as indexed PNGs, so VisualDiff needs palette PNG support in addition to RGB/RGBA.
  Evidence: The first `pptx-blank` visual run failed on a reference PNG with `bitDepth=1 colorType=3`; after adding palette/grayscale support, `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-blank/case.json` completed successfully.

- Observation: Visual validation now works end to end for blank PPTX, blank DOCX, and simple PPTX shapes on this Windows machine with Office installed.
  Evidence: `pptx-blank` run `20260514-132845` has matching 1920 by 1080 reference/candidate images and zero pixel error; `docx-blank` run `20260514-132733` has matching 1224 by 1584 images and zero pixel error; `pptx-shapes` run `20260514-132846` has matching 1920 by 1080 images with mean absolute error `0.23382511091820987` and changed pixel ratio at threshold 16 of `0.0032373649691358024`.

- Observation: Common Windows TrueType fonts are sufficient to validate the first font parser slice.
  Evidence: `WindowsFontResolverFindsInstalledFonts` discovers installed fonts through `C:\Windows\Fonts`, and `OpenTypeParserMapsBasicLatinGlyphs` loads `arial.ttf`, reads the family name, required table records, `maxp` glyph count, selected `OS/2` metrics, selected `post` metrics, maps `A` through `cmap`, and reads a positive advance from `hmtx`; the full test run now prints `16 passed, 0 failed`.

- Observation: Full-font TrueType embedding is enough for the first readable PPTX text slice, though layout remains approximate.
  Evidence: `WritesEmbeddedTrueTypeFontObjects` validates Type0, CIDFontType2, FontDescriptor, FontFile2, and ToUnicode objects; `PptxSyntheticTextBoxEmbedsFontAndDrawsGlyphs` verifies PPTX text emits embedded font resources and glyph drawing commands. The full test run now prints `18 passed, 0 failed`. `pptx-text` visual run `20260514-134307` has matching 1920 by 1080 dimensions with mean absolute error `1.3818904320987655` and an assessment rating of 3.

- Observation: PPTX bold, italic, underline, and paragraph alignment can be represented with simple PDF approximations before full font-face selection exists.
  Evidence: `PptxSyntheticStyledTextProducesStyleOperators` verifies italic text matrices, underline line drawing, run color, and the bold double-draw approximation; the full test run now prints `19 passed, 0 failed`.

- Observation: Common PPTX theme colors and theme Latin font references can be resolved without implementing full master/layout inheritance.
  Evidence: `PptxSyntheticThemeColorsAndFontsResolve` adds a synthetic theme relationship, resolves `schemeClr` values such as `accent1` and `dk1`, resolves `+mn-lt` through the theme font scheme, and the full test run now prints `20 passed, 0 failed`.

- Observation: Basic slide master and layout inheritance can be handled by following slide-to-layout and layout-to-master relationships and rendering inherited shape trees before the slide tree.
  Evidence: `PptxSyntheticLayoutAndMasterShapesRender` builds a synthetic slide with master, layout, and slide-local rectangles and verifies all three PDF fill colors are emitted; the full test run now prints `21 passed, 0 failed`.

- Observation: The first image rendering path can use direct PDF image XObjects without a general raster library.
  Evidence: `JpegInfoReadsDimensions` validates JPEG SOF dimension parsing, and `PptxSyntheticPngPictureRendersImageXObject` builds a synthetic PNG picture relationship and verifies `/Subtype /Image`, `/XObject`, `/Im1 Do`, and image dimensions in the candidate PDF; the full test run now prints `23 passed, 0 failed`.

- Observation: The first Office-authored PPTX image visual case renders end to end with usable fidelity.
  Evidence: `pptx-images` run `20260514-135714` has matching 1920 by 1080 reference/candidate images, mean absolute error `1.1591259162808643`, changed pixel ratio at threshold 16 of `0.013763020833333334`, and an assessment rating of 4.

- Observation: Basic PPTX image cropping can be represented with PDF clipping and an adjusted image transform.
  Evidence: `PptxSyntheticCroppedPictureUsesClipping` verifies that a picture with `a:srcRect` emits clipping (`re W n`) and still draws the image XObject; the full test run now prints `24 passed, 0 failed`.

- Observation: Common grouped shapes can be flattened by composing group child coordinate transforms before rendering child shapes.
  Evidence: `PptxSyntheticGroupedShapeAppliesTransform` verifies a child rectangle inside a scaled group emits the expected transformed PDF rectangle; the full test run now prints `25 passed, 0 failed`.

- Observation: Simple PPTX tables can be rendered as decomposed cell rectangles and text runs.
  Evidence: `PptxSyntheticTableRendersGridAndText` builds a synthetic `p:graphicFrame` with a 2 by 2 DrawingML table, verifies cell fill and border operators, and verifies embedded text output; the full test run now prints `26 passed, 0 failed`.

- Observation: Inherited slide master placeholders can contain editable template text that PowerPoint does not render on slides.
  Evidence: The first `pptx-table` visual run showed `Click to edit Master text styles` over the slide. `PptxSyntheticInheritedPlaceholderTextIsSkipped` now verifies inherited placeholder text does not create PDF font resources, and the full test run now prints `27 passed, 0 failed`.

- Observation: The first Office-authored PPTX table visual case renders the core table content but exposes simplified border and cell text layout.
  Evidence: `pptx-table` run `20260514-140923` has matching 1920 by 1080 reference/candidate images, mean absolute error `2.446560329861111`, changed pixel ratio at threshold 16 of `0.019048996913580248`, and an assessment rating of 4. The assessment records darker black grid borders and approximate vertical text placement as the main defects.

- Observation: Unsupported PPTX feature detection can run before rendering and aggregate repeated feature warnings per slide.
  Evidence: `PptxUnsupportedFeaturesEmitDiagnostics` builds a synthetic slide containing transition, animation, video, audio, OLE, chart, and SmartArt markers, then verifies stable `PPTX_UNSUPPORTED_*` warning diagnostics scoped to slide 1; the full test run now prints `28 passed, 0 failed`.

- Observation: The first DOCX text slice can reuse the existing embedded TrueType PDF path for readable one-page paragraph output.
  Evidence: `DocxSyntheticParagraphRendersText` builds a synthetic DOCX with page margins, a centered paragraph, run font size, color, bold, and underline, then verifies embedded font output, glyph drawing, red fill color, and underline stroke; the full test run now prints `29 passed, 0 failed`.

- Observation: Common DOCX style defaults and simple paragraph/character styles can be resolved during parsing before rendering.
  Evidence: `DocxSyntheticStylesApplyToParagraphText` builds a synthetic styles part with document defaults, a paragraph style, and a character style, then verifies the resolved font size, blue color, italic approximation, and underline stroke in PDF output; the full test run now prints `30 passed, 0 failed`.

- Observation: Simple DOCX paragraph rendering can continue onto additional PDF pages instead of dropping overflow text.
  Evidence: `DocxSyntheticParagraphsBreakAcrossPages` builds a synthetic DOCX with 45 paragraphs and verifies the produced PDF page tree contains two pages; the full test run now prints `31 passed, 0 failed`.

- Observation: The first Office-authored DOCX paragraph visual case renders with usable text fidelity but exposes layout differences.
  Evidence: `docx-basic-paragraphs` run `20260514-142144` has matching 1224 by 1584 reference/candidate images, mean absolute error `1.0340011635967519`, changed pixel ratio at threshold 16 of `0.011712818544926389`, and an assessment rating of 4. The assessment records top-baseline and paragraph-spacing differences as the main defects.

- Observation: Simple DOCX numbering can be resolved from `numbering.xml` and rendered as generated paragraph label prefixes.
  Evidence: `DocxSyntheticNumberingRendersListLabels` builds a synthetic numbering part and verifies two numbered paragraphs render through the embedded text path; the full test run now prints `32 passed, 0 failed`.

- Observation: The first Office-authored DOCX numbering visual case is readable but needs hanging-indent and level-text cleanup.
  Evidence: `docx-numbering` run `20260514-142632` has matching 1224 by 1584 reference/candidate images, mean absolute error `1.0996359376031557`, changed pixel ratio at threshold 16 of `0.012073863636363636`, and an assessment rating of 3. The assessment records labels like `1 .` and missing hanging indents as the main defects.

- Observation: DOCX inline DrawingML images can reuse the same PNG/JPEG PDF image XObject path as PPTX.
  Evidence: `DocxSyntheticInlinePngRendersImageXObject` builds a synthetic DOCX with `w:drawing/wp:inline`, resolves the image relationship, and verifies `/Subtype /Image`, `/XObject`, `/Im1 Do`, and image dimensions in the candidate PDF; the full test run now prints `33 passed, 0 failed`.

- Observation: The first Office-authored DOCX inline image visual case renders the bitmap accurately with approximate surrounding paragraph spacing.
  Evidence: `docx-images` run `20260514-143302` has matching 1224 by 1584 reference/candidate images, mean absolute error `0.5217354302832244`, changed pixel ratio at threshold 16 of `0.0071971760084505185`, and an assessment rating of 4. The assessment records paragraph/image baseline spacing as the main defect.

- Observation: Simple DOCX fixed-grid tables can be decomposed into cell fills, black borders, and embedded text drawing.
  Evidence: `DocxSyntheticTableRendersCellsAndText` builds a synthetic fixed-width table with `w:tblGrid`, shaded cells, and cell text, then verifies fill, stroke, embedded font, and text drawing operators; the full test run now prints `34 passed, 0 failed`.

- Observation: The first Office-authored DOCX table visual case is readable but does not yet preserve table cell text styling.
  Evidence: `docx-tables` run `20260514-143739` has matching 1224 by 1584 reference/candidate images, mean absolute error `2.0042449876625734`, changed pixel ratio at threshold 16 of `0.028418374925727866`, and an assessment rating of 3. The assessment records missing white bold header text and approximate row/cell spacing as the main defects.

- Observation: Basic DOCX default headers and footers can be resolved from section relationships and drawn on each generated page.
  Evidence: `DocxSyntheticHeaderAndFooterRenderOnPage` builds a synthetic DOCX with default header and footer relationships and verifies body, header, and footer text are all drawn through embedded font output; the full test run now prints `35 passed, 0 failed`.

- Observation: Common DOCX `PAGE` fields in headers and footers can be approximated during static header/footer rendering.
  Evidence: `DocxSyntheticFooterPageFieldUsesGeneratedPageNumbers` builds a two-page synthetic DOCX with a footer `PAGE` field and verifies footer text is drawn on each generated page; the full test run now prints `36 passed, 0 failed`.

- Observation: Word can store both a `PAGE` field instruction and its cached numeric result, so field parsing must suppress the cached result when substituting generated page numbers.
  Evidence: The first `docx-headers-footers` visual run rendered `Page 11`; after suppressing digit-only runs after a `PAGE` instruction, rerun `20260514-144543` renders `Page 1`, has matching 1224 by 1584 dimensions, mean absolute error `1.3533170501997096`, changed pixel ratio at threshold 16 of `0.014745597312999273`, and an assessment rating of 4.

- Observation: Unsupported DOCX feature detection can run before rendering and aggregate repeated warnings per document.
  Evidence: `DocxUnsupportedFeaturesEmitDiagnostics` builds a synthetic DOCX containing comments, tracked changes, non-page complex fields, equations, OLE, floating drawing, footnote/endnote references, multi-column section markup, and a VBA project, then verifies stable `DOCX_UNSUPPORTED_*` warning diagnostics; the full test run now prints `37 passed, 0 failed`.

- Observation: Current PDF output is deterministic for a stable OOXML package because object ordering, resource names, and metadata are stable.
  Evidence: `DeterministicConversionProducesStableBytes` converts the same synthetic DOCX twice with `Deterministic = true` and verifies byte-for-byte equality; the full test run now prints `38 passed, 0 failed`.

- Observation: CLI exit codes are now covered through process-level tests instead of only through direct API calls.
  Evidence: `CliConvertReturnsZeroOnSuccess`, `CliReturnsOneForConversionFailure`, `CliReturnsTwoForInvalidArguments`, and `CliStrictReturnsThreeWhenWarningsAreEmitted` invoke `src/Lokad.OoxPdf.Cli` and verify exit codes `0`, `1`, `2`, and `3`; the full test run now prints `42 passed, 0 failed`.

Examples of discoveries that belong here include: Office COM automation requiring a visible desktop session, PDFium output naming differing from expectations, a Microsoft font using an unexpected `cmap` format, a PPTX fixture storing shape colors through a theme rather than direct RGB, or Word producing an extra blank page due to section breaks.

## Decision Log

- Decision: Build the library with no `PackageReference` dependencies and isolate all external tools under `tools/`.
  Rationale: The user explicitly wants a NuGet library with no dependencies beyond standard libraries. Office and PDFium are allowed for validation only, not for production conversion.
  Date/Author: 2026-05-14 / ExecPlan author.

- Decision: Target `net10.0` initially.
  Rationale: `net10.0` gives modern .NET standard-library APIs while keeping the implementation simple. Multi-targeting can be added later only if it does not introduce dependencies.
  Date/Author: 2026-05-14 / ExecPlan author.

- Decision: Use a vertical-slice implementation order: blank PDF pages first, then PPTX basics, then DOCX basics, then fidelity improvements.
  Rationale: Visual rendering projects need observable output early. A blank but correctly sized PDF proves package parsing and PDF writing before the harder layout work begins.
  Date/Author: 2026-05-14 / ExecPlan author.

- Decision: Implement a simple console-based test runner instead of xUnit or NUnit.
  Rationale: Standard test frameworks are third-party packages. The dependency policy applies most strongly to the library, but keeping tests dependency-free avoids accidental dependency drift and makes the repository easier for a coding agent to bootstrap.
  Date/Author: 2026-05-14 / ExecPlan author.

- Decision: Treat pixel metrics as advisory and human or agent visual inspection as authoritative.
  Rationale: Office and PDFium use different rasterizers. Small antialiasing differences can look large numerically while being visually harmless. The harness should expose evidence, not pretend that a single number proves correctness.
  Date/Author: 2026-05-14 / ExecPlan author.

- Decision: Do not use `System.Drawing.Common`.
  Rationale: It is not a reliable cross-platform rendering foundation in modern .NET, and the project needs predictable dependency-free behavior. PNG and JPEG parsing should be implemented directly where required.
  Date/Author: 2026-05-14 / ExecPlan author.

- Decision: Use the bblanchon `pdfium.dll` binary with a local dependency-free P/Invoke rasterizer tool instead of expecting `pdfium_test.exe`.
  Rationale: The requested latest `pdfium-win-x64.tgz` package contains the PDFium DLL but no `pdfium_test.exe`. A small tool under `tools/` keeps PDFium isolated to validation while preserving the no-runtime-dependency rule for `src/Lokad.OoxPdf`.
  Date/Author: 2026-05-14 / Codex.

## Outcomes & Retrospective

- Outcome: Phase 0, blank-page conversion, visual comparison scaffolding, and first simple PPTX shape rendering are implemented. The repository builds with `Lokad.OoxPdf.slnx`, the library has the planned public API shell, the CLI can produce PDFs for recognized PPTX and DOCX inputs, and the visual harness creates Office reference PNGs, candidate PDFs, PDFium candidate PNGs, comparison metrics, HTML indexes, and assessment files. VisualDiff writes `metrics.json` and `index.html`, reads common grayscale, indexed, RGB, and RGBA PNGs, and computes dimensions plus simple pixel metrics.
  Validation: `dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal` succeeds with 0 warnings and 0 errors. `dotnet run --project tests/Lokad.OoxPdf.Tests --tl:off` prints `42 passed, 0 failed`. `dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore` succeeds. `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-blank/case.json`, `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/docx-blank/case.json`, `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-shapes/case.json`, `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-text/case.json`, `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-images/case.json`, `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-table/case.json`, `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/docx-basic-paragraphs/case.json`, `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/docx-numbering/case.json`, `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/docx-images/case.json`, `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/docx-tables/case.json`, and `pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/docx-headers-footers/case.json` all complete successfully on this machine.
  Remaining gaps: Rendering covers simple PPTX solid backgrounds, rectangles, lines, ellipses, basic rotation/flip transforms, simple Latin text runs with basic style approximations, common theme color/font references, common master/layout inheritance, JPEG/PNG pictures with basic cropping, grouped shape coordinate transforms, fixed-grid tables with simple fills, black borders, and text, warning diagnostics for common unsupported PPTX slide features, and DOCX paragraphs with style defaults, paragraph/character styles, spacing, alignment, basic run formatting, page breaking, simple bullet/decimal list labels, inline JPEG/PNG images, fixed-width tables with simple fills, borders, and text, default header/footer text, page number field approximation, warning diagnostics for common unsupported DOCX features, deterministic output for stable inputs, and CLI exit-code behavior. Documentation remains incomplete.
  Next target: Fill in release-readiness documentation, starting with rendering model, visual validation, and README examples.

## Context and Orientation

This repository is expected to start either empty or nearly empty. Create all paths named in this plan relative to the repository root.

`Lokad.OoxPdf` is the core library. It must live in `src/Lokad.OoxPdf`. It is the only project intended for NuGet packaging. It must not call Office, PDFium, PowerShell, external executables, or third-party packages.

`Lokad.OoxPdf.Cli` is a small command-line wrapper in `src/Lokad.OoxPdf.Cli`. It references the library project and lets contributors run conversions from a terminal. It is useful for validation and examples.

`Lokad.OoxPdf.Tests` is a console test runner in `tests/Lokad.OoxPdf.Tests`. A console test runner is just a program that calls test methods, reports pass/fail lines, and exits with code `0` on success and nonzero on failure. This avoids test framework dependencies.

`Lokad.OoxPdf.VisualDiff` is a tool in `tools/Lokad.OoxPdf.VisualDiff`. It compares PNG images produced from Office and from the candidate PDF. It writes JSON metrics, side-by-side images, diff images, and an `index.html` that an agent can inspect.

`tools/RenderReference.ps1` uses Office COM automation. COM automation means controlling installed Microsoft Office applications through Windows automation objects. It is allowed only as a development reference mechanism. Never call it from the library.

`tools/RasterizePdf.ps1` uses a vendored `pdfium_test.exe`. PDFium is an external PDF rasterizer. It is allowed only in the visual harness to turn PDFs into PNGs. Never call it from the library.

OOXML means Office Open XML. `.pptx` and `.docx` files are ZIP packages containing XML parts and binary media such as images. A “part” is one file inside the ZIP package. A “relationship” is an XML entry that points from one part to another, such as from a slide to an image. A “content type” identifies what kind of part a ZIP entry represents.

PDF points are the coordinate unit used by PDF. One inch is 72 points. EMU is a common Office drawing unit; one inch is 914400 EMUs. Twips are Word layout units; one point is 20 twips.

A scene graph is an internal list of visual objects to draw on a page, such as text runs, paths, images, and filled rectangles. The implementation should parse OOXML into semantic information, convert that into a page-level scene graph, then write PDF commands from the scene graph.

A diagnostic is a structured warning, error, or informational message. Diagnostics are essential because unsupported Office features must not silently disappear. If the renderer ignores a video, chart, SmartArt object, unsupported image type, or complex Word field, it must emit a stable diagnostic code.

## Target Repository Layout

Create this layout. Some directories will be empty at first, but they should exist so future commits have clear destinations.

    Lokad.OoxPdf/
      README.md
      LICENSE
      .gitignore
      Directory.Build.props
      Lokad.OoxPdf.slnx
      src/
        Lokad.OoxPdf/
          Lokad.OoxPdf.csproj
          OoxPdfConverter.cs
          OoxPdfOptions.cs
          OoxPdfInputKind.cs
          Diagnostics/
            OoxPdfDiagnostic.cs
            OoxPdfSeverity.cs
            DiagnosticCollector.cs
          Ooxml/
            OoxPackage.cs
            OoxPart.cs
            OoxRelationship.cs
            OoxContentTypes.cs
            OoxPath.cs
            SafeXml.cs
            OoxUnits.cs
          Pdf/
            PdfDocumentWriter.cs
            PdfObjectWriter.cs
            PdfPage.cs
            PdfRectangle.cs
          Layout/
            PageScene.cs
            SceneNode.cs
          Pptx/
            PptxDocument.cs
            PptxReader.cs
            PptxRenderer.cs
          Docx/
            DocxDocument.cs
            DocxReader.cs
            DocxRenderer.cs
          Fonts/
            IFontResolver.cs
            FontRequest.cs
            FontResolution.cs
          Imaging/
            JpegInfo.cs
            PngImage.cs
        Lokad.OoxPdf.Cli/
          Lokad.OoxPdf.Cli.csproj
          Program.cs
      tests/
        Lokad.OoxPdf.Tests/
          Lokad.OoxPdf.Tests.csproj
          Program.cs
          TestAssert.cs
          OoxmlTests.cs
          PdfWriterTests.cs
          PptxTests.cs
          DocxTests.cs
          TestFixtures.cs
          Cases/
      tools/
        RenderReference.ps1
        RasterizePdf.ps1
        CheckVisualCase.ps1
        NewVisualCase.ps1
        vendor/
          pdfium/
            win-x64/
              README.md
              pdfium_test.exe
              LICENSE.txt
              VERSION.txt
        Lokad.OoxPdf.VisualDiff/
          Lokad.OoxPdf.VisualDiff.csproj
          Program.cs
      visual-cases/
        README.md
        cases/
          pptx-blank/
            case.json
          docx-blank/
            case.json
      artifacts/
        .gitkeep
      docs/
        Capabilities.md
        RenderingModel.md
        Diagnostics.md
        VisualValidation.md

Add `artifacts/**` to `.gitignore` but keep `artifacts/.gitkeep` tracked.

## Interfaces and Dependencies

The core library must expose these public types.

In `src/Lokad.OoxPdf/OoxPdfConverter.cs`, define:

    namespace Lokad.OoxPdf;

    public static class OoxPdfConverter
    {
        public static void Convert(string inputPath, string outputPdfPath, OoxPdfOptions? options = null);

        public static void Convert(Stream input, string inputNameOrExtension, Stream outputPdf, OoxPdfOptions? options = null);
    }

In `src/Lokad.OoxPdf/OoxPdfOptions.cs`, define:

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

In `src/Lokad.OoxPdf/OoxPdfInputKind.cs`, define:

    namespace Lokad.OoxPdf;

    public enum OoxPdfInputKind
    {
        Auto = 0,
        Pptx = 1,
        Docx = 2
    }

In `src/Lokad.OoxPdf/Diagnostics/OoxPdfSeverity.cs`, define:

    namespace Lokad.OoxPdf;

    public enum OoxPdfSeverity
    {
        Info,
        Warning,
        Error
    }

In `src/Lokad.OoxPdf/Diagnostics/OoxPdfDiagnostic.cs`, define:

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

In `src/Lokad.OoxPdf/Fonts/IFontResolver.cs`, define the public font resolver:

    namespace Lokad.OoxPdf;

    public interface IFontResolver
    {
        FontResolution? Resolve(FontRequest request);
    }

In `src/Lokad.OoxPdf/Fonts/FontRequest.cs`, define:

    namespace Lokad.OoxPdf;

    public sealed record FontRequest(
        string FamilyName,
        bool Bold,
        bool Italic,
        string? Script = null);

In `src/Lokad.OoxPdf/Fonts/FontResolution.cs`, define:

    namespace Lokad.OoxPdf;

    public sealed record FontResolution(
        string FontPath,
        int FaceIndex = 0);

Keep implementation classes internal unless they are intentionally part of the API. Prefer `internal sealed` classes for package readers, renderers, and PDF writer internals.

The library project file must have no `PackageReference`. The CLI, tests, and tools should also avoid package references unless a future decision log explicitly records an exception. The expected project references are: CLI references library, tests reference library, VisualDiff references nothing or references shared source if later refactored.

## Plan of Work

### Milestone 0: Repository skeleton and dependency guardrails

Create the repository structure, solution, projects, and basic documentation. The user-visible result is that `dotnet build` works and `dotnet pack` can run on the library project without any dependencies. This milestone does not render anything yet.

Start by creating `Directory.Build.props` with common settings: nullable enabled, implicit usings enabled, deterministic builds enabled, latest major language version or C# language version compatible with `net10.0`, and warnings treated seriously but not necessarily as errors until the codebase stabilizes.

Create `src/Lokad.OoxPdf/Lokad.OoxPdf.csproj` as an SDK-style library targeting `net10.0`. Add package metadata placeholders such as `PackageId`, `Authors`, `Description`, `RepositoryUrl`, and `GeneratePackageOnBuild` set to false. Do not add package references.

Create `src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj` as an executable targeting `net10.0` with a project reference to the library.

Create `tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj` as an executable targeting `net10.0` with a project reference to the library.

Create `tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj` as an executable targeting `net10.0`.

Create `Lokad.OoxPdf.slnx` and add all projects. With the .NET 10 SDK used in this workspace, `dotnet new sln -n Lokad.OoxPdf` creates an `.slnx` file.

Validation for this milestone is:

    dotnet restore Lokad.OoxPdf.slnx --tl:off -v minimal
    dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal
    dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj -c Release

The pack command should produce a `.nupkg` under `src/Lokad.OoxPdf/bin/Release/` and the generated NuGet package should not list dependencies.

### Milestone 1: Public API, CLI shell, and diagnostics

Add the public API files listed in the Interfaces and Dependencies section. Implement `OoxPdfConverter.Convert` so it validates arguments, opens streams, detects input kind from extension when `InputKind.Auto` is used, and then dispatches to placeholder PPTX or DOCX conversion methods. At this milestone, the placeholder methods may throw `NotImplementedException` or produce a blank PDF once the PDF writer exists. Prefer returning a controlled `InvalidOperationException` with a diagnostic for unsupported or invalid input rather than crashing with null-reference errors.

Implement `DiagnosticCollector` as an internal helper that accepts diagnostics, forwards them to `OoxPdfOptions.DiagnosticSink`, and remembers them for the CLI. The collector should preserve insertion order.

Implement CLI argument parsing manually in `src/Lokad.OoxPdf.Cli/Program.cs`. Do not add a command-line package. The CLI must support:

    dotnet run --project src/Lokad.OoxPdf.Cli -- convert input.pptx output.pdf
    dotnet run --project src/Lokad.OoxPdf.Cli -- convert input.docx output.pdf --diagnostics diagnostics.json
    dotnet run --project src/Lokad.OoxPdf.Cli -- convert input.docx output.pdf --strict

Exit codes are: `0` success, `1` conversion failed, `2` unsupported input or invalid arguments, and `3` conversion succeeded but strict mode saw warnings or errors.

At this milestone, conversion may fail with a controlled error because the package reader and PDF writer are not done. The observable result is that help, invalid arguments, diagnostics output, and exit codes behave predictably.

### Milestone 2: Dependency-free test runner

Implement a small test runner in `tests/Lokad.OoxPdf.Tests/Program.cs`. It should register named test methods, run them, print one line per test, and exit nonzero if any test fails. Add `TestAssert.cs` with helpers such as `Equal`, `True`, `False`, `Contains`, and `Throws`.

Add initial tests for argument validation and input-kind detection. To test input-kind detection without needing a full OOXML file yet, isolate the detection logic in an internal method and expose it to tests through `InternalsVisibleTo` or test through public error messages.

The command:

    dotnet run --project tests/Lokad.OoxPdf.Tests

should print a short transcript similar to:

    PASS PublicApiRejectsMissingInput
    PASS AutoDetectsPptxExtension
    PASS AutoDetectsDocxExtension
    3 passed, 0 failed

### Milestone 3: Safe OOXML package reader

Implement `OoxPackage` over `System.IO.Compression.ZipArchive`. It should open a stream, index entries by normalized package path, expose `TryGetPart`, `OpenPart`, `ReadXml`, and relationship lookup methods.

Implement `OoxPath.NormalizePartName`. A package part name should always use forward slashes, should start with `/` internally, should not contain `..`, should not be absolute filesystem paths, and should not contain empty path segments except the leading slash.

Implement `SafeXml.CreateReaderSettings` with `DtdProcessing.Prohibit`, `XmlResolver = null`, and reasonable limits if available. Use this settings object everywhere XML is read.

Implement `OoxContentTypes` to parse defaults and overrides from `[Content_Types].xml`.

Implement `OoxRelationship` and relationship parsing. Relationship targets can be relative to the source part. For example, a relationship in `/ppt/slides/_rels/slide1.xml.rels` with target `../media/image1.png` should resolve to `/ppt/media/image1.png`.

Add tests for safe path normalization, path traversal rejection, content type lookup, package relationship parsing, and part relationship resolution.

### Milestone 4: Minimal deterministic PDF writer

Implement a direct PDF writer in `src/Lokad.OoxPdf/Pdf`. It must write enough PDF syntax to create valid blank pages. Use ASCII output for PDF operators and invariant culture formatting for numbers.

At this stage, support:

- PDF header `%PDF-1.7`.
- Indirect objects with deterministic object numbers.
- Catalog object.
- Pages tree.
- Page objects with `/MediaBox`.
- Empty content streams.
- Cross-reference table.
- Trailer with `/Root`.
- `startxref`.

The writer should accept a list of page sizes in points and write to a stream. A point is 1/72 inch. For deterministic output, object order must be stable.

Add structure tests that read the produced PDF bytes as Latin-1 or ASCII-compatible text and assert that `%PDF-1.7`, `/Type /Catalog`, `/Type /Pages`, `/Count 2`, `/MediaBox [0 0 612 792]`, `xref`, `trailer`, and `startxref` appear.

### Milestone 5: Blank-page PPTX and DOCX conversion

Implement `PptxReader` to discover the presentation part from package relationships, read slide size from `ppt/presentation.xml`, read slide IDs, resolve slide relationships, and return a `PptxDocument` containing slide count, slide dimensions in points, and slide part names.

If slide size is absent, use a reasonable default 10 by 7.5 inches. Emit an informational diagnostic when a default is used.

Wire PPTX conversion so it writes one blank PDF page per slide with the discovered slide dimensions.

Implement `DocxReader` to discover the main document part from package relationships, read the body, and look for section page size and margins. At this milestone, it can return a one-page `DocxDocument` with discovered or default page dimensions. Use Letter or A4 based on the document settings when obvious; otherwise choose Letter initially and record the decision if this is later changed.

Wire DOCX conversion so it writes one blank PDF page with discovered dimensions.

Add simple synthetic fixture generation in tests. Since `.pptx` and `.docx` are ZIP files, create test packages in memory using `ZipArchive` with the minimal XML parts needed. Do not commit binary fixtures unless the XML package generation becomes too cumbersome.

Validation is that converting the synthetic PPTX creates a PDF with one page per slide and converting the synthetic DOCX creates a one-page PDF.

### Milestone 6: Visual harness foundation

Add `tools/RasterizePdf.ps1`. It should accept `-PdfPath`, `-OutputDir`, `-Dpi`, and `-Prefix`. It should locate `tools/vendor/pdfium/win-x64/pdfium_test.exe`, call it with PNG output and a scale equal to `Dpi / 72`, and rename generated files to `page-001.png`, `page-002.png`, and so on. It should write `rasterize-manifest.json`.

Add `tools/RenderReference.ps1`. For PPTX, use PowerPoint COM automation to export each slide as PNG at the requested DPI. For DOCX, use Word COM automation to export a reference PDF, then call `RasterizePdf.ps1`.

Add `tools/CheckVisualCase.ps1`. It should read a case manifest, create a timestamped run directory under `artifacts/visual/<case-id>/<run-id>`, run the reference renderer, run the CLI conversion, rasterize the candidate PDF, run `VisualDiff`, and create an `assessment.md` from a template.

Add `visual-cases/README.md` with exact commands and environmental assumptions: Windows, PowerShell 7 or Windows PowerShell, Office 365 installed, and `pdfium_test.exe` placed under `tools/vendor/pdfium/win-x64`.

Do not commit Office-generated visual artifacts under `artifacts/`; that directory is for local runs.

### Milestone 7: VisualDiff foundation

Implement `tools/Lokad.OoxPdf.VisualDiff/Program.cs` to parse:

    --reference <directory>
    --candidate <directory>
    --output <directory>

Discover `page-*.png` files in both directories. Initially, write `metrics.json` with page counts and file names. Then write `index.html` that shows each reference PNG beside its candidate PNG using relative paths.

Next, implement enough PNG parsing to read width, height, bit depth, color type, and decompressed pixel data for common PNGs. Support color type 2, which is RGB, and color type 6, which is RGBA. Use `System.IO.Compression.ZLibStream` for IDAT decompression. Implement PNG filters None, Sub, Up, Average, and Paeth.

Then implement PNG writing for RGB or RGBA images. If writing PNG proves unexpectedly time-consuming, temporarily generate only `index.html` and record the gap in `Surprises & Discoveries`, but complete PNG writing before release.

Compute per-page metrics: width and height match, mean absolute RGB error, root mean squared RGB error, changed pixel ratio at threshold 16, and changed pixel ratio at threshold 32. Write one `page-###-report.json` per page and combined `metrics.json`.

Generate `page-###-side-by-side.png` with reference on the left, candidate on the right, and a small separator. Generate `page-###-diff.png` where absolute RGB differences are multiplied by 4 and clamped to 255.

Validation is that running VisualDiff on two identical PNGs gives zero error, matching dimensions, and produces side-by-side and diff files.

### Milestone 8: Basic PDF graphics and PPTX solid shapes

Extend the PDF writer to support content stream operations. Add a small internal page canvas or content builder that can write `q`, `Q`, `cm`, path commands, `re`, `m`, `l`, `c`, `h`, `f`, `S`, `B`, line width, dash pattern, RGB fill, and RGB stroke.

Implement a scene model with `PageScene` and `SceneNode`. At this stage, support rectangles, lines, ellipses, and filled backgrounds.

Parse PPTX slide background solid fills. Parse basic shapes from slide XML, including position and size from DrawingML transforms. DrawingML often uses EMUs, so convert EMUs to PDF points. Respect z-order by drawing shapes in the same order they appear in the slide shape tree.

Render rectangles and lines first. Then render ellipses using four cubic Bézier curves. Add rounded rectangles after normal rectangles and ellipses work.

Add visual case `pptx-shapes` containing a slide with a colored background, rectangle, line, and ellipse. Run `CheckVisualCase.ps1` and record an assessment. Early ratings may be low; the important result is that visible shapes appear in the right approximate positions.

### Milestone 9: Font parsing, PDF font embedding, and PPTX text

Implement Windows font discovery by scanning `C:\Windows\Fonts` when running on Windows. Keep the API flexible so other platforms can later provide an `IFontResolver`. The default resolver should not throw if the font directory is missing; it should emit diagnostics and fall back.

Implement TrueType/OpenType parsing for table directory and the minimum required tables. The goal is simple Latin text measurement and embedding, not full professional text shaping.

For text measurement, map Unicode scalar values to glyph IDs through `cmap`, use horizontal advances from `hmtx`, scale by font size, and optionally apply simple kerning later. For this milestone, support ordinary Latin text, numbers, punctuation, and spaces.

Implement PDF font embedding using a Type0/CID font approach with a ToUnicode CMap. If full subsetting is too much initially, embed the whole font file for development, then optimize later. Record this decision if made because full embedding can produce large PDFs.

Parse PPTX text boxes, paragraphs, and runs. Support font family, font size, bold, italic, underline if simple, text color, paragraph alignment, margins, and word wrapping. Implement greedy line breaking: add words to a line until the measured width would exceed the text box, then start a new line.

Render text into the PDF. Add a PPTX visual case with a title and body text. The acceptance is that text is readable, in approximately the right location, with the right color and size.

### Milestone 10: PPTX themes, masters, layouts, images, groups, and tables

Implement theme color resolution. A theme color is a named color such as accent1 or tx1 that resolves through the presentation theme. Parse the theme part and map common scheme colors to RGB.

Implement theme font resolution. If a run asks for a major or minor theme font, resolve it through the theme.

Implement slide master and slide layout inheritance. In PowerPoint, visible slide content can come from the theme, the slide master, the slide layout, and the slide itself. The renderer should draw background and inherited layout/master shapes before slide-local shapes. Handle common placeholders such as title and body.

Implement JPEG passthrough into PDF. Parse JPEG dimensions from markers and embed the original bytes with `/DCTDecode`.

Implement PNG image embedding for color type 2 and 6. For RGBA PNGs, put RGB in the image XObject and alpha in a soft mask image XObject.

Parse PPTX picture relationships and render images at the specified position and size. Add basic crop support if it is straightforward; otherwise emit a diagnostic and render uncropped as a fallback.

Implement grouped shape transforms. A group transform means child coordinates are expressed inside a parent coordinate system. Compose transforms carefully and add tests for nested translation and scaling.

Implement simple PPTX tables by decomposing each cell into a fill rectangle, border lines, and text. Support fixed grids first.

Add visual cases for corporate theme, images, grouped shapes, and tables. Update `docs/Capabilities.md` after each feature.

### Milestone 11: DOCX paragraph layout and text rendering

Implement DOCX style parsing. Word styles can define default paragraph and run properties. Parse document defaults, paragraph styles, and character styles from `word/styles.xml`. Apply the cascade in this order: document defaults, paragraph style, run style, direct paragraph formatting, direct run formatting.

Implement DOCX page setup. Read page size, orientation, and margins from section properties. Convert twips to points. One point is 20 twips.

Implement paragraph layout. For each paragraph, compute available line width from page width minus margins and paragraph indents. Use the same simple font measurement as PPTX. Support spacing before, spacing after, line spacing, left alignment, center alignment, right alignment, and justified text only as left alignment with a diagnostic until true justification is implemented.

Implement page breaking. If the next line would exceed the page content bottom margin, start a new page. Preserve manual page breaks. Add a second pass later for total page count fields.

Render DOCX paragraphs to PDF. Add a DOCX visual case with headings, paragraphs, bold, italic, colors, and page breaks. Acceptance is that the page count and visible text layout roughly match Word for simple documents.

### Milestone 12: DOCX numbering, images, tables, headers, and footers

Parse `word/numbering.xml` for simple bullets and decimal numbering. Support indentation from numbering levels. Render bullet characters and decimal prefixes as text.

Parse inline images from DrawingML. Resolve relationships to image parts and reuse the image rendering code from PPTX. Support image sizing and aspect ratio preservation.

Parse fixed-width tables. Support rows, cells, grid columns, cell margins, fills, borders, horizontal merges, and vertical merges where common. Render tables by laying out cell boxes and text inside each cell.

Parse headers and footers. Render default header and default footer on every page. Add support for first-page header/footer if the section requests it. Approximate page number fields by writing the current page number during rendering.

Add visual cases for numbering, inline images, fixed-width tables, and headers/footers. Acceptance is that simple corporate reports with paragraphs, lists, images, tables, and page numbers are usable and visually close enough for ratings of 3 or better before deeper fidelity work.

### Milestone 13: Unsupported feature diagnostics and strict mode

Add feature detection and diagnostics for unsupported PPTX features: animations, transitions, videos, audio, macros, ActiveX, OLE live objects, SmartArt without fallback, and charts without fallback. The renderer must not crash when it sees these features.

Add feature detection and diagnostics for unsupported DOCX features: comments, tracked changes beyond final visible text, complex fields without cached text, equations without fallback drawing, OLE live content, complex floating wrap types, footnotes, endnotes, multi-column layout, macros, bidirectional text, and complex scripts.

Implement duplicate diagnostic aggregation where useful. For example, a document with 50 unsupported animation elements should not produce 50 identical noisy warnings unless the locations are important.

Implement strict mode in the CLI. If conversion succeeds but diagnostics include warnings or errors, exit with code `3`. Write diagnostics JSON either way when requested.

Add tests for diagnostic codes and strict mode exit behavior.

### Milestone 14: Determinism, packaging, documentation, and release readiness

Implement deterministic PDF output. Convert the same fixture twice in the same process and compare bytes. Resource names, object ordering, metadata, and creation dates must be stable when `OoxPdfOptions.Deterministic` is true. If `FixedCreationDate` is set, include that value in PDF metadata; otherwise omit creation date in deterministic mode.

Update `docs/Capabilities.md` continuously. At the end of this milestone it should describe PPTX and DOCX features as Supported, Partial, Approximated, Ignored, or Unsupported.

Update `README.md` with a library example, CLI example, visual validation example, dependency statement, and known limitations.

Run clean build, tests, pack, and at least one PPTX and one DOCX visual case. Record assessment files under the local artifacts directory and summarize outcomes in this ExecPlan. Do not commit generated artifacts unless they are intentionally small fixtures.

## Concrete Steps

Start from the repository root. If the repository already contains files, inspect them first and adapt the steps while preserving the public API and layout required by this plan.

Create the solution and projects:

    dotnet new sln -n Lokad.OoxPdf
    mkdir src
    mkdir tests
    mkdir tools
    mkdir docs
    mkdir visual-cases
    mkdir artifacts
    dotnet new classlib -n Lokad.OoxPdf -o src/Lokad.OoxPdf -f net10.0
    dotnet new console -n Lokad.OoxPdf.Cli -o src/Lokad.OoxPdf.Cli -f net10.0
    dotnet new console -n Lokad.OoxPdf.Tests -o tests/Lokad.OoxPdf.Tests -f net10.0
    dotnet new console -n Lokad.OoxPdf.VisualDiff -o tools/Lokad.OoxPdf.VisualDiff -f net10.0
    dotnet sln Lokad.OoxPdf.slnx add src/Lokad.OoxPdf/Lokad.OoxPdf.csproj
    dotnet sln Lokad.OoxPdf.slnx add src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj
    dotnet sln Lokad.OoxPdf.slnx add tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj
    dotnet sln Lokad.OoxPdf.slnx add tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj
    dotnet add src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj reference src/Lokad.OoxPdf/Lokad.OoxPdf.csproj
    dotnet add tests/Lokad.OoxPdf.Tests/Lokad.OoxPdf.Tests.csproj reference src/Lokad.OoxPdf/Lokad.OoxPdf.csproj

Create or edit `Directory.Build.props`:

    <Project>
      <PropertyGroup>
        <Nullable>enable</Nullable>
        <ImplicitUsings>enable</ImplicitUsings>
        <Deterministic>true</Deterministic>
        <ContinuousIntegrationBuild Condition="'$(CI)' == 'true'">true</ContinuousIntegrationBuild>
        <LangVersion>latestMajor</LangVersion>
      </PropertyGroup>
    </Project>

Edit `src/Lokad.OoxPdf/Lokad.OoxPdf.csproj` so it contains no package references and includes NuGet metadata:

    <Project Sdk="Microsoft.NET.Sdk">
      <PropertyGroup>
        <TargetFramework>net10.0</TargetFramework>
        <PackageId>Lokad.OoxPdf</PackageId>
        <Authors>Lokad</Authors>
        <Description>Dependency-free OOXML to PDF renderer for PPTX and DOCX corporate documents.</Description>
        <PackageReadmeFile>README.md</PackageReadmeFile>
        <RepositoryType>git</RepositoryType>
        <GeneratePackageOnBuild>false</GeneratePackageOnBuild>
      </PropertyGroup>
      <ItemGroup>
        <None Include="../../README.md" Pack="true" PackagePath="/" Condition="Exists('../../README.md')" />
      </ItemGroup>
    </Project>

Add `.gitignore`:

    bin/
    obj/
    .vs/
    .idea/
    artifacts/**
    !artifacts/.gitkeep
    *.user
    *.suo

Create `artifacts/.gitkeep`.

After the skeleton exists, run:

    dotnet restore Lokad.OoxPdf.slnx --tl:off -v minimal
    dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal

Expected result:

    Build succeeded.
        0 Warning(s)
        0 Error(s)

After adding the public API and CLI shell, run:

    dotnet run --project src/Lokad.OoxPdf.Cli -- convert

Expected result:

    Usage: Lokad.OoxPdf.Cli convert <input.pptx|input.docx> <output.pdf> [--diagnostics <file>] [--strict]

and the process should exit with code `2`.

After implementing tests, run:

    dotnet run --project tests/Lokad.OoxPdf.Tests

Expected result after the initial tests:

    PASS PublicApiRejectsMissingInput
    PASS AutoDetectsPptxExtension
    PASS AutoDetectsDocxExtension
    3 passed, 0 failed

After implementing blank-page conversion, create synthetic fixtures through tests rather than manually at first. Then run:

    dotnet run --project tests/Lokad.OoxPdf.Tests

Expected result should include tests similar to:

    PASS PptxSyntheticTwoSlidesProducesTwoPdfPages
    PASS DocxSyntheticDocumentProducesOnePdfPage

After adding the visual harness, place `pdfium_test.exe` under:

    tools/vendor/pdfium/win-x64/pdfium_test.exe

Then run a local visual case on Windows:

    pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-blank/case.json

Expected result is an artifacts directory like:

    artifacts/visual/pptx-blank/20260514-.../
      reference/page-001.png
      candidate/output.pdf
      candidate/page-001.png
      comparison/index.html
      comparison/metrics.json
      assessment.md

Open `comparison/index.html` and verify that both reference and candidate images are visible.

As each milestone completes, run the standard validation set:

    dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal
    dotnet run --project tests/Lokad.OoxPdf.Tests
    dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore

Before declaring any visual feature complete, run the relevant visual case with:

    pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/<case-id>/case.json

Then inspect:

    artifacts/visual/<case-id>/<run-id>/comparison/index.html
    artifacts/visual/<case-id>/<run-id>/assessment.md

Update the assessment with a rating from 0 to 5:

    5 = visually indistinguishable except tiny antialiasing differences
    4 = good; minor spacing, kerning, or antialiasing differences only
    3 = usable; visible but non-blocking differences
    2 = poor; major layout defects or missing important elements
    1 = severe; page or slide mostly wrong
    0 = conversion failed or page missing

## Validation and Acceptance

The project is acceptable at the end of Phase 0 when a clean checkout can run:

    dotnet build Lokad.OoxPdf.slnx --tl:off --nologo -v minimal
    dotnet run --project tests/Lokad.OoxPdf.Tests
    dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj --tl:off --nologo -v minimal --no-restore

and all commands succeed. The generated NuGet package must not list third-party dependencies.

The blank conversion milestone is acceptable when a synthetic PPTX with two slides produces a PDF whose structure reports two page objects with the expected slide dimensions, and a synthetic DOCX produces a one-page PDF with expected document dimensions.

The visual harness is acceptable when, on a Windows machine with Office 365 and PDFium available, this command:

    pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-blank/case.json

creates reference PNGs, candidate PDF, candidate PNGs, metrics JSON, index HTML, and assessment template in `artifacts/visual`.

PPTX basic rendering is acceptable when the `pptx-shapes`, `pptx-text`, and `pptx-images` visual cases run end-to-end and receive an agent rating of at least 3, with no unexplained missing major elements. A lower rating is acceptable during early commits but must be recorded honestly in `assessment.md`.

DOCX basic rendering is acceptable when the `docx-basic-paragraphs`, `docx-numbering`, `docx-images`, and `docx-tables` visual cases run end-to-end and receive an agent rating of at least 3, with page count matching simple documents.

Strict mode is acceptable when a document containing a known unsupported feature converts in non-strict mode with a warning diagnostic, and the same document exits with code `3` under `--strict`.

Determinism is acceptable when converting the same fixture twice with deterministic options produces byte-identical PDFs.

The release-candidate state is acceptable when:

- The library project has no `PackageReference`.
- Office and PDFium are used only in `tools/`.
- PPTX and DOCX representative cases convert without crashing.
- Unsupported static features emit diagnostics.
- `docs/Capabilities.md` accurately describes supported and unsupported behavior.
- The README contains working library, CLI, and visual validation examples.
- `dotnet pack -c Release` produces a consumable NuGet package.

## Idempotence and Recovery

All build and test commands are safe to rerun. The visual harness writes to timestamped directories under `artifacts/visual`, so repeated runs should not overwrite prior runs unless a script explicitly accepts a `-RunId` and the same run ID is reused.

If `dotnet new` is accidentally run over existing files, inspect the diff before committing. Prefer small commits so accidental generated changes can be reverted easily.

If Office COM automation leaves a hidden PowerPoint or Word process running after a script failure, close it manually through Task Manager or run a cleanup command only after confirming no important Office documents are open. The scripts should call `Quit()` and release COM objects in `finally` blocks to reduce this risk.

If `pdfium_test.exe` is missing, `RasterizePdf.ps1` should fail with a clear message. Do not commit random downloaded binaries without adding `LICENSE.txt` and `VERSION.txt` under `tools/vendor/pdfium/win-x64`.

If a visual case uses confidential material, do not commit the input or generated artifacts. Create a synthetic anonymized equivalent and commit only that.

If a parser encounters unexpected XML, prefer emitting a diagnostic and rendering the rest of the document over throwing. Throw only when the package is invalid in a way that prevents safe processing.

If a feature implementation causes broad regressions, keep the failing visual artifacts under `artifacts/` locally, revert the commit, and add a note to `Surprises & Discoveries` describing the failure and evidence.

## Artifacts and Notes

Use this assessment template in every visual run. `CheckVisualCase.ps1` should create it automatically as `assessment.md`:

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

Use this shape for visual case manifests:

    {
      "id": "pptx-blank",
      "kind": "pptx",
      "input": "../../../tests/Lokad.OoxPdf.Tests/Cases/pptx-blank.pptx",
      "dpi": 144,
      "tags": ["pptx", "blank", "smoke"],
      "expected": {
        "minAgentRating": 3,
        "pageCountMustMatch": true,
        "dimensionsMustMatch": true
      },
      "allowedUnsupportedFeatures": []
    }

Use this shape for diagnostics JSON entries:

    {
      "id": "PPTX_UNSUPPORTED_VIDEO",
      "severity": "Warning",
      "message": "Video media was ignored; static PDF output does not support video playback.",
      "partName": "/ppt/slides/slide4.xml",
      "slideIndex": 3,
      "feature": "video",
      "fallback": "ignored"
    }

Use this shape for visual metrics:

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

Use stable diagnostic code prefixes:

- `OOXML_` for package-level and XML-level diagnostics.
- `PDF_` for PDF writer diagnostics.
- `PPTX_` for PowerPoint-specific diagnostics.
- `DOCX_` for Word-specific diagnostics.
- `FONT_` for font resolution, parsing, and embedding diagnostics.
- `IMAGE_` for image parsing and rendering diagnostics.

Examples:

- `OOXML_EXTERNAL_RELATIONSHIP_IGNORED`
- `PPTX_UNSUPPORTED_SMARTART`
- `PPTX_CHART_FALLBACK_MISSING`
- `DOCX_FLOATING_IMAGE_WRAP_APPROX`
- `DOCX_COMPLEX_FIELD_USING_CACHED_TEXT`
- `FONT_COMPLEX_SCRIPT_SHAPING_UNSUPPORTED`
- `IMAGE_UNSUPPORTED_FORMAT`

## Granular Commit Guidance

Each item in `Progress` is intended to be a separate commit when practical. Use commit messages that describe observable behavior, not only internal files. Good examples are:

    Add dependency-free solution skeleton
    Add public converter API and CLI shell
    Add safe OOXML package path normalization
    Write deterministic blank PDF pages
    Convert PPTX slide count into blank PDF pages
    Add PowerPoint reference rendering script
    Add VisualDiff metrics and comparison index
    Render PPTX solid rectangles
    Embed TrueType fonts for basic Unicode text
    Render DOCX basic paragraphs

Avoid large commits named only `work in progress`. If a task becomes too large, split it and update `Progress` immediately with completed and remaining subitems.

## Change Log for This ExecPlan

- 2026-05-14: Initial ExecPlan created. It translates the product goal into a granular implementation plan, defines repository layout, public APIs, validation commands, visual harness behavior, diagnostics policy, and milestone acceptance criteria.
