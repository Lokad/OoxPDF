# Visual Validation

Visual validation exists because a valid PDF structure does not prove visual fidelity. The harness renders an Office-produced reference, renders the candidate PDF with PDFium, compares PNG pages, and creates files that an agent or contributor can inspect.

## Requirements

- Windows with Microsoft Office installed.
- PowerShell.
- .NET SDK.
- `tools/vendor/pdfium/win-x64/bin/pdfium.dll`.
- .NET restore needs network access to nuget.org (or a local feed such as the ignored `artifacts/localfeed/` workaround); without it only prebuilt binaries under `bin/` can run.
- Reference rendering drives Office through COM: PowerPoint closes other presentations and Word runs headless (`DisplayAlerts = 0`); large decks export silently for many minutes with no progress output, and a missing `artifacts/reference-cache/` entry fails cache-only private runs instead of invoking Office.

The repository vendors the exact rasterizer binary (`tools/vendor/pdfium/win-x64/bin/pdfium.dll`, PDFium 150.0.7834.0 per `VERSION`); checkouts must use it as-is. Do not replace it from `releases/latest`: a different build silently changes raster output, and `tools/RasterizePdf.ps1` verifies the binary against the pinned SHA256 in `tools/vendor/pdfium/win-x64/pdfium.sha256` before every run. To upgrade, fetch the versioned release matching `VERSION`, replace the DLL, and update the pin file in the same commit.

## Running A Case

Pinned DocxEdit 0.2.0 producer fixtures cover authored/edited equations and SVG
placements, including crop, repeated images and mixed equation/image flow.
Recipes and regeneration instructions live in
[Cases/DocxEditInterop](../tests/Lokad.OoxPdf.Tests/Cases/DocxEditInterop/README.md).
The `docx-math-producer-*` and `docx-svg-producer-*` cases retain existing visual
thresholds. Matrix fence/baseline residuals remain review items. Consecutive body
pictures share their paragraph spacing with the keep estimate; table/story image
flow remains separately qualified. DocxEdit is fixture tooling only.

For affected DOCX qualification, use `tools/RunDocxAffectedTests.ps1` with the
frozen runner and a fresh output directory. It runs overlapping DOCX/balloon
methods once and gives each runner invocation the output directory's
`temporary-files/` path through process-local `TMP` and `TEMP`, restoring both
variables afterward. This avoids slow fixture creation in a crowded Windows
temp directory. Keep the generated files with the ignored validation output.
For other test batches, use a fresh ignored temp directory in the launching
process; retain the same frozen binaries, fixtures and checks.

```powershell
pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/docx-basic-paragraphs/case.json
```

The script:

1. Builds the CLI and converts the test input to `candidate/output.pdf`.
2. Uses PowerPoint or Word COM automation to render reference output.
3. Rasterizes the candidate PDF through the local PDFium rasterizer.
4. Runs `Lokad.OoxPdf.VisualDiff`.
5. Writes an `assessment.md` template.

Outputs are timestamped under:

```text
artifacts/visual/<case-id>/<run-id>/
```

Important files:

- `reference/page-001.png`: Office reference page.
- `candidate/output.pdf`: generated candidate PDF.
- `candidate/page-001.png`: PDFium rasterization of the candidate.
- `candidate/diagnostics.json`: emitted renderer diagnostics.
- `comparison/metrics.json`: dimensions and pixel-difference metrics.
- `comparison/index.html`: side-by-side browser review.
- `assessment.md`: agent review notes and rating.

## Assessment

Pixel metrics are advisory. Office and PDFium can differ in antialiasing, font hinting, and rasterization details even when the PDF is acceptable. The assessment rating is the authoritative result:

- `5`: visually indistinguishable except tiny antialiasing differences.
- `4`: good; minor spacing, kerning, or antialiasing differences only.
- `3`: usable; visible but non-blocking differences.
- `2`: poor; major layout defects or missing important elements.
- `1`: severe; page or slide mostly wrong.
- `0`: conversion failed or page missing.

Do not commit generated visual artifacts unless they are intentionally small fixtures.

`tools/ValidateVisualCases.ps1` also checks public PPTX character-property
highlight ordering. DrawingML puts `highlight` before font and underline children;
this targeted check is not a complete OOXML schema validator. The
[Open XML SDK content model](https://github.com/dotnet/Open-XML-SDK/blob/main/data/schemas/schemas_openxmlformats_org_drawingml_2006_main.json)
records the sequence. PowerPoint 16 discarded five late highlights in four public
typography probes, both in PDF and PNG preview; corrected ordering survives import.
Those fixtures and their generator definitions are corrected, including garbled
accents in the boundary-invariance fixture. The renderer is unchanged. That case's
strict text-operation/line-start gates remain partial: Office now splits highlighted
text into seven operations while the candidate coalesces it into four. All text is
present and line origins differ by at most 0.02pt; its gates remain unchanged.
Original inputs, eight original/ordered Office controls, the corrected accent
reference and failed strict comparisons are retained under ignored
`artifacts/plan-revision-20261005/rv08-t3/`. Windows and Linux validate all 345 cases.

## Family Parity Targets (Q07 triage, 2026-09-22)

Each family has an explicit target classification so needs-review triage has a
destination. Counts below track the validated manifest state as of 2026-10-05
(118 locked, 9 locked-text-ops, 216 approximate, 0 needs-review across 343
cases; family rows overlap on shared patterns so their totals exceed 343).

| Family | Cases | Target | Rationale |
|---|---|---|---|
| pptx-images | 13 / 0 / 3 / 0 | Embedding locked; SVG paints approximate | Radial/stroke probe records remaining paint and geometry residuals |
| pptx-charts | 28 / 0 / 34 / 0 | Tier-1 ports locked; rest approximate with recorded gaps | Matches the Tier-1/Tier-2 split in Capabilities.md; triage complete 2026-09-22 |
| pptx-typography | 15 / 9 / 71 / 0 | Approximate by default; exact ports lock | Font-metric approximations are structural; text-ops locks pin emission; run-merging fix aligns op grouping with Office |
| pptx-tables | 9 / 0 / 7 / 0 | Mixed; first-pass built-ins lock, rich styles approximate | Per Capabilities table-style scope; triage complete 2026-09-22 |
| pptx-shapes | 21 / 0 / 12 / 0 | Approximate; small preset geometry only | Preset-geometry scope in Capabilities.md; triage complete 2026-09-22 |
| pptx-composition | 10 / 0 / 3 / 0 | Mixed; master/layout inheritance locks case by case | Triage complete 2026-09-22 |
| pptx-effects | 5 / 0 / 4 / 0 | Mixed; raster shadows/glows and gradient approximations render, other effects approximate | Per Capabilities effects scope; triage complete 2026-09-22 |
| pptx-smoke | 6 / 0 / 0 / 0 | Locked | Blank/size discovery must stay exact |
| docx-layout | 11 / 0 / 49 / 0 | Mixed; greedy-wrap approximations stay approximate | Latin greedy wrapping scope in Capabilities.md; markup cases tracked under docx-markup |
| docx-markup | 0 / 0 / 34 / 0 (gated via reference cache) | Approximate with margin modes tracked | Reference-cache workflow below |

Columns are locked / locked-text-ops / approximate / needs-review.

### Approximate-family limitations

Named limitation, intended next improvement, and gate for each approximate family. Counts are the current validated state.

- pptx-images: SVG vectors stay approximate (strip tessellation, solid path strokes with mapped dash/cap/join/miter, non-scaling strokes, CSS colors, inherited group paint, basic shape elements, evenodd rules and radial gradients). Radial pad fills now extend the final stop outside the outer circle. The Office-authored `pptx-ladder-07-svg-stroke-radial` probe covers caps, dashes, miter limits, scaled/non-scaling strokes and radial fills; generate it with `tools/NewSvgStrokeRadialFixtures.ps1`. Its reference uses explicit user-space gradient coordinates because this PowerPoint import flattens the object-bounding-box variants to the final color. Focal points remain diagnosed and ignored, so this case is approximate. On the PowerPoint 16 validation host, six independent focal controls (numeric, percentage, px, zero focal radius, outside focus and centered) export centered. Reproduce that Office observation with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet focal-controls`, which writes to ignored artifacts by default. Centered fallback preserves that Office appearance; it does not implement [SVG focal-gradient semantics](https://www.w3.org/TR/SVG2/pservers.html#RadialGradientElement). User-space percentage and px gradient coordinates remain unsupported by the converter. Representable radial pad fills use smooth PDF shading when the radius and stop intervals survive PDF number precision. The stroke/radial probe improves MAE from 1.659 to 1.526 and shrinks from 12,406 to 1,649 PDF bytes. Five independent opaque controls improve, covering source-unit rescaling, three stops, endpoint extension and a small radius; the opaque pad milestone retained transparent and hard-stop controls. Regenerate these nine controls with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet radial-paint-controls`. Representable repeat/reflect fills also use bounded stitched radial functions through the full shape, rather than stopping at the first circle. Ten independent Office controls (two/three stops, shifted centers, small radii and rescaled source units) improve MAE from 17-23 to 0.27-1.87 and foreground recall from 2-25% to 100%. Regenerate them with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet radial-spread-controls`. Native repetition is limited to 128 cycles and 256 expanded normalized stops. Hard/near-stop and dense-repetition fills retain sampled rendering; their Office differences and stroke-position/width residuals remain open. Picture viewport stretching now preserves directional stroke widths, caps and dashes through a normalized PDF transform. Three independent Office stretch controls improve MAE 0.78 to 0.33, 1.52 to 1.18 (dashes), and 0.82 to 0.34 (square caps); seven uniform/group/rotation/shear/non-scaling controls retain identical rasters. Reproduce these ten controls with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet stroke-transform-controls`. Transform coefficients with more than 0.1% rounding error retain the prior scalar approximation; source-coordinate rescaling and extreme-aspect fallback are covered by production regressions. PowerPoint 16 uses a uniform stroke width scaled by the largest singular value of the element transform. Eleven independent stretch/contraction/rotation/shear/reflection and filled-shape controls qualify that rule; regenerate them with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet stroke-element-controls`. The nine unfilled controls improve MAE 0.12-0.88 to 0.06-0.16; the filled controls improve 0.71 to 0.17 and 0.20 to 0.09, and the earlier group stretch and shear controls improve 1.18 to 0.26 and 0.38 to 0.18. Seven earlier controls and five held-out PDFs retain their bytes. The saved Office deck removes vector-effect from the two non-scaling controls; direct production regressions separately preserve non-scaling widths, including CSS. Combined solid fill/stroke paint uses the same viewport map while retaining one PDF paint operation and the separate fill/stroke alpha state. Eight independent Office controls (opaque, evenodd, fill/stroke/path/CSS alpha, square caps and non-scaling) improve MAE 0.65-1.97 to 0.17-0.87; generator rasters and five held-out PDFs remain identical. Regenerate them with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet stroke-filled-controls`. Source-unit rescaling, paint-operation count, alpha state and extreme-aspect fallback have production regressions. Dash offset follows Office stroke-width units while dash-array lengths remain source geometry. Eighteen independent width/negative/positive offset controls improve 16 cases (MAE 0.31-2.49 to 0.17-0.36); two offsets are cycle-equivalent. All eighteen generator rasters reproduce. The SVG probe improves MAE 1.282 to 1.072 and SSIM 0.979 to 0.983, with four other held-out PDF identities. Regenerate the width/phase matrix with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet stroke-dash-phase-controls` and the eighteen cap/phase controls with `-ProbeSet stroke-dash-controls`. Source-unit rescaling and composed offset-overflow regressions pass. Square-cap dash footprints compensate the PDF cap extent when every normalized painted segment exceeds the stroke width. All six earlier square-cap controls improve MAE 0.63-1.25 to 0.21-0.31; the other twelve cap controls retain their rasters. Twenty of twenty-six independent width/phase/odd/multi-pair controls improve, and six equal/short/zero boundary controls retain identical approximate rasters. Regenerate this matrix with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet stroke-square-dash-controls`. Production regressions cover painted footprint, source-unit rescaling, normalized odd/multi-pair cycles, composed gap overflow with retained fill/sibling paint and the boundary fallback. Equal/short/zero square dashes remain partial; zero centerline dash segments worsened Office parity and are excluded from compensation. Standard SVG affine stroke outlines, standard dash-offset semantics and path-level opacity compositing remain scoped residuals. The approximate gate does not establish full SVG parity. Qualified fill-only radial transparency applies node/fill alpha once through one clipped radial shading. Eighteen fill-only PowerPoint preview controls improve MAE 1.46-14.61 to 0.046-0.586 (minimum SSIM 0.999211), covering pad/repeat/reflect and attribute/CSS alpha. Regenerate them with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet radial-opacity-controls`; native PNG previews stay in a separate `office-preview/` directory from PDF rasters. PowerPoint 16 PDF export tiles circular transparency regions and leaves holes, so the user-selected target is the smooth preview. Group/root/stop alpha and fill/stroke opacity isolation are outside this qualification. Unused shapes under `defs` are excluded before path parsing, following [SVG definition semantics](https://www.w3.org/TR/SVG2/struct.html#DefsElement). A production byte-invariance regression covers every basic shape, unsupported path commands and unused gradient alpha; referenced radial paint still resolves. Five held-out PDFs retain their bytes. Partial CSS/percentage container alpha and varying, CSS or percentage stop opacity remain unsupported and produce explicit warnings for rendered containers or referenced gradients. RV07-E5 qualifies uniform numeric stop alpha and numeric zero-opacity containers; RV07-E6 qualifies numeric partial outer-SVG/group alpha through bounded isolated composition. Group compositing requires an isolated result, so applying parent alpha independently to overlapping children would not implement [SVG group opacity](https://www.w3.org/TR/SVG2/render.html#ObjectAndGroupOpacity). [Gradient stop opacity](https://www.w3.org/TR/SVG2/pservers.html#StopOpacityProperty) is a separate unimplemented paint property.

Qualified opaque round-dashed strokes on separate straight subpaths use a native flat-cap dashed body and outward semicircles. A visible first/last dash end rounds only when it lies within half a stroke width of an original endpoint; the outward half avoids filling the neighboring gap when the terminal painted segment is shorter than the radius. Every subpath must exceed the stroke width and either contain multiple visible dashes or have a single visible dash with at least one end outside the endpoint region; the normalized two-slot cycle must have draw length greater than the width and a positive gap. Qualification is bounded to 128 subpaths and at most 1e12 cycles per subpath. Source geometry locates cycle boundaries before the rounded viewport transform introduces drift. Twenty-eight of thirty-two independent uniform/stretched width/endpoint/phase controls improve; four short-line rasters retain their bytes. Regenerate the matrix with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet stroke-round-dash-controls`. Production regressions cover source-unit rescaling, negative phase, initial/trailing gaps, exact cycle boundaries, independent viewport axes, endpoint-region eligibility, outward partial caps and native fallback for alpha, curves, joins, closed/odd/short/zero/dense/oversized/incomplete paths. Combined solid fill/stroke paths retain their current cap behavior. Separate body/cap paint retains small antialias and placement differences. Office uses a solid rounded fallback when an entire opaque straight path falls in dash gaps. The bounded two-slot preflight covers positive draw/gap lengths and at most 128 single-line subpaths; a visible dash anywhere in the path suppresses this fallback, while separate path elements qualify independently. Twelve of fifty-two short/compound/element controls improve and forty rasters retain their bytes; all thirty-two endpoint-control rasters are unchanged. Regenerate with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet stroke-short-dash-controls`. Production regressions cover eighteen source-unit/width/length variants, wholly unpainted versus mixed paths and alpha fallback. Eight of the same fifty-two controls improve when a single visible dash has an end outside the Office endpoint region; forty-four short-control and all thirty-two endpoint-control rasters retain their bytes. A single visible dash with both rounded ends already matching keeps native paint. A production regression covers thirty source-unit, viewport and endpoint variants. Short painted paths at or below the stroke width and equal/short draw slots retain native approximation. The approximate pixel gate does not establish parity for these residuals.
- pptx-charts: Tier-2 behavior stays approximate (trendlines, secondary axes, leader-line labels, percent stacking) with calibrated gallery-line styling (raw 5/7pt strokes, Hermite smoothing, gradient markers, nominal-5 sizeless defaults); next is other chart families and styles with held-out Office calibration. Gate: family 62/62, pptx-charts group, structure, color, and label checks plus pixels.
- pptx-typography: font metrics stay structural approximations with autofit and overflow handling approximate; next is held-out Office calibration for the remaining approximate ports. Gate: family 95/95, pptx-typography group, text operations and line starts plus pixels.
- pptx-tables: rich table styles stay approximate while mixed-run word wrap is fixed; next is held-out Office calibration for rich-style cells. Gate: pptx-tables group, text operations and line starts plus pixels.
- pptx-shapes: small preset geometry only with scene-only custom geometry; next is held-out Office calibration for further geometry kinds. Gate: family 33/33, pptx-shapes group, path-operation assertions plus pixels.
- pptx-composition: master and layout inheritance locks case by case; next is extending locked inheritance coverage. Gate: family gates plus pixels.
- pptx-effects: raster shadows and glows approximate, other effects unsupported with diagnostics; next is held-out Office calibration for further effects. Gate: family 9/9 plus pixels.
- docx-layout: Latin greedy wrapping with mid-line images across body, table-cell, related-story and static paths plus approximate columns, notes, and floating wrap; Office calibration of image baseline, line growth, and justification is complete; remaining scope is explicit non-goals. Gate: family 60/60, docx-text groups, words and line starts plus pixels.
- docx-markup: margin modes tracked with diagnosed fallbacks including word-compatible text. On the validation workstation, all 34 public manifests have cached references covering 26 distinct input/view identities. Office parity remains partial; scoped repairs and remaining layout/balloon/text gates are recorded in `docs/DocxMarkupFidelityTriage.md`. Gate: cached Office gates plus layout snapshots.

SVG gradient geometry is checked after transforms and coordinate mapping, before
painting. Finite input values whose spans, sampling projection or radial bounds
overflow are diagnosed and their fills omitted. Usable solid strokes and adjacent
shapes remain. This is an extreme-value fallback, not arbitrary-precision SVG
geometry support; ordinary SVG reference-case PDF bytes remain unchanged.

### Lock policy

- New locks require a reviewed agent rating of 4 or 5, tight pixel gates at
  roughly 2-3x observed headroom (MAE, changed-pixel ratio, foreground recall),
  empty diagnostics where the manifest demands it, and passing family support
  gates (SSIM/histogram correlation where the family defines them).
- The 2026-09-21 corpus predates this bar (uniform rating 3, recall thresholds
  often absent); those locks are grandfathered because their pixel gates do the
  real work. A 2026-09-22 audit confirmed all 59 then-locked cases reference
  committed inputs.
- Promotions in this round: `pptx-ladder-11-chart-bar-clustered-port`,
  `pptx-ladder-11-chart-pie-5-categories-port`,
  `pptx-ladder-11-chart-column-clustered-port`, and
  `pptx-ladder-11-chart-bar-stacked-port` to locked (indistinguishable renders,
  all metrics tight including family gates, empty diagnostics);
  `pptx-ladder-11-chart-line-markers-port` to approximate (marker-fill shade
  gap from the documented 0.975x unstyled-series approximation; fails family
  histogram correlation at 0.77); `pptx-ladder-11-chart-line-3series-port` to
  approximate (same unstyled-series shade family, no markers; histcorr 0.79).

## Allocation and peak-memory probes

`tools/MeasureAllocation.ps1 -FontPath <font.ttf>` supplies one explicit TrueType
face for all requested families. The underlying AllocProbe `--font-file` option
also applies to isolated inputs, stage measurements, concurrency and self-tests.
Reports pin the actual font-program SHA-256 and size, source lifetime, input hash,
runtime and revision. Per-conversion metrics count embedded TrueType resources;
the explicit-font self-test rejects fallback-only output. This probe resolver
uses a regular face and shares its lazy file source across warm conversions.
It measures that controlled workload rather than Office font matching.

Allocation volume is attributed to the calling thread. Retained managed heap,
sampled managed/private/working-set peaks and process lifetime high-water marks
are separate metrics. Use `--output-mode file` to exclude caller output buffering,
or `--output-mode forward-only` for a strict non-seekable caller stream backed by
a temporary file. The latter exercises stream conversion with seek/position/length
rejected and checks that caller streams remain open. Self-tests require identical
PDF bytes across all three output modes. Use `--concurrency 1`, `2`, or `4` for batch peaks. `--isolate` uses fresh processes
per input and cannot be combined with concurrency.

## Cached DOCX Markup References

Cached markup comparisons build the CLI and visual diff tool through
`tools/EnsureDotnetBuild.ps1`, which delegates dependency and imported-property
tracking to MSBuild and serializes shared builds. A failed build stops the
comparison. Each run records `cli-build-info.json` and, when raster comparison
is enabled, `visual-diff-build-info.json` with the selected output hash.
Run `pwsh tools/TestCachedMarkupBuild.ps1` to check property-driven refresh,
incremental output stability, provenance and rejection of an existing DLL
after a failed build. These checks require .NET and no Office installation.

`InspectPdf.ps1` and `InspectPptxText.ps1` use the same shared MSBuild gate.
When an output directory is provided, PDF inspection records
`pdf-inspect-build-info.json`; PPTX inspection records
`pptx-inspect-build-info.json`. These identify the selected inspector DLL and
build inputs. Run `pwsh tools/TestInspectorBuild.ps1` to verify both production
build calls against changed parent properties, unchanged output, provenance
and failed-build rejection.

DOCX markup parity work uses cache-only reference comparisons so autonomous runs do not launch Word or COM. Generate a reference request for the public DOCX markup cases:

```powershell
pwsh tools/NewDocxMarkupReferenceRequest.ps1 -OutputDirectory artifacts/docx-markup-reference-requests/public
```

Use `-MissingOnly` when refreshing the handoff after some cached references already exist; the request summary records ready and missing counts without launching Office.

After trusted Office PDFs have been produced on a controlled setup, import them with `tools/ImportDocxMarkupReferenceCache.ps1`. Then audit or run the cached gate:

```powershell
pwsh tools/RunDocxMarkupReferenceGate.ps1 -CacheStatusOnly
pwsh tools/RunDocxMarkupReferenceGate.ps1 -FailOnDeltas
```

The cache-status output includes `missing-import-commands.ps1`, which lists only the references still absent from `artifacts/reference-cache/`.

`CacheVariant` identifies a cached PDF; it does not set Word's persistent view
options. `RenderCachedReference.ps1` serves verified markup hits, but refuses
generic rendering of missing or corrupt DOCX markup variants. Export with all
view settings explicit, then import against the case manifest. Corrupt markup
entries are preserved for inspection. Ordinary DOCX/PPTX cache fills are unchanged.

The markup baseline and merged-line text geometry gates ignore decoded
whitespace-only lines. Runs merge first, so spaces between words still contribute
to line advances. Raw operator counts and undecoded runs remain available; blank
paragraph pagination must be assessed through layout snapshots and pixels.
The standalone `ComparePdfTextOperations.ps1` keeps its default full comparison;
use `-MergeSameLineOperations -IgnoreWhitespaceOnlyLines` for this geometry policy.
`tools/TestMarkupTextGeometry.ps1` and the text comparison adversarial checks cover
missing/moved lines, inline spaces and undecoded runs without Office or a cache.

### Reference revision-view matching

Word-compatible all-markup rendering declares a RevisionsMode-1-like view:
insertions and deletions render inline with revision styling while comments and
formatting revisions balloon. A reference exported with RevisionsMode=0 instead
balloons deletions (`Deleted:` title rows) and hides them from body lines, so
the gate compares different views: body wraps, op pairings, and page flow all
shift for non-layout reasons. Export word-compatible all-markup references with
`RevisionsMode=1` (explicit `ExportDocxMarkupReference.ps1` knobs; the remaining
four knobs stay pinned as usual) and record the full ExportSettings string at
import time; re-import view-mismatched references with `-Force`.

Verify the view before importing: a RevisionsMode-1 reference carries deleted
text inline at body size with no `Deleted:`/`Inserted:` title operations, while
a RevisionsMode-0 reference moves that text into balloon rows. Fixtures without
revision markup (`w:ins`/`w:del`/moves in `document.xml`) are view-agnostic and
need no re-export. Precedents: margin-dense-revisions and margin-landscape were re-imported
view-matched (record the full ExportSettings string, back up the replaced entry,
re-import with `-Force`). Comment-long and comment-unresolved carry no revision
markup; the all reference already matches the inline-deletion class;
links-fields-final is a final-mode case whose ShowRevisions=False reference
already matches. Check the case manifest markup/geometry before importing: a
mismatched variant key silently orphans the entry and no gate resolves it.

### Reference field-result matching

Word can refresh unlocked `REF` fields during PDF export. Lokad.OoxPdf preserves
their stored results, including revision markup. Compare the resulting text
before attributing a different wrap or downstream balloon position to layout.
The public `docx-markup-links-fields-locked-all` companion locks its five REF
fields and keeps their stored results and revisions. Regenerate just that
companion with `pwsh tools/NewDocxMarkupLinkFieldFixtures.ps1 -LockedOnly`;
the existing unlocked and note fixtures retain their bytes. Keep both cases:
the unlocked case measures the field-refresh policy difference, while the
locked case isolates the remaining spacing, table and pixel differences.

## RV07-E5: bounded SVG opacity (2026-10-10)

Uniform numeric `stop-opacity` attributes now multiply the existing fill/node alpha.
Numeric zero-opacity root/group containers do not paint or activate diagnostics
for their invisible contents. Varying stops, percentage/CSS stop alpha and partial
container opacity retain their prior fallback. The selected reference remains
PowerPoint16's smooth PNG preview; its PDF-export transparency holes are outside
the target. Office rewrites SVG styles and percentages during import, so evidence
pins the original source separately from the SVG actually stored in the PPTX.

Frozen runtime **e9ff75a5** passes one unfiltered original catalogue invocation:
**2250 registered /2249 passed /0 failed /1 optional private-document skip**.
It includes **110 SVG /131 images /959 DOCX** methods and the prior L127 note
fixes. Eight targeted Linux checks pass; Windows builds are clean and the
git-free Linux archive retains two expected SourceLink warnings. Two unchanged
final regressions fail against the exact L127 library; the fallback guard passes.
The exact local **0.1.5** package smoke is `90c728feb5974e92b81cf3a8c6fdaf68`.

Seventeen new Office slides yield **ten paired MAE/SSIM improvements** and seven
unchanged guard rasters. The two later controls independently cover a vertical
three-color linear gradient and a shifted reflecting three-color radial gradient.
Fresh frozen PDFs match the separately compared prototype PDFs, transferring
pinned preview/raster/graphics evidence with its original library provenance.
Radial MAE falls **11.412→0.054**; the independent reflecting case falls
**17.181→0.236**. Linear sampling remains approximate (vertical MAE **0.327**,
SSIM **0.987976**); opacity support does not widen the gradient geometry claim.

Across **39 prior SVG decks**, 38 retain exact PDF bytes. One older opacity deck
changes its uniform-stop slide: MAE **11.421→0.051**, SSIM
**0.799984→0.999999** against its original Office preview. Its other five rasters
retain bytes. All **2911 inherited/table PDFs** retain accepted L127 bytes.
Eighty fresh DOCX comparison PDFs also retain bytes; cached markup failures stay
**508**, with original source-layout/structural/raster provenance preserved.

The caught test-helper compile error and initial corpus identity stop remain
archived. Partial group/root opacity, varying stop alpha, overlapping composition,
gradient strokes and earlier chart/text/DOCX residuals remain. Version stays
0.1.5; release preparation is deferred. Evidence:
`artifacts/plan-revision-20261005/rv07-e5/`. Frozen runtime, package and subsequent
documentation integration identities remain distinct.

## RV08-T1: bounded PPTX column word splitting (2026-10-10)

Emergency unbroken first-token chunks in horizontal two/three-column frames now
use the measured column edge. A font-sized fit allowance admitted an extra
character and displaced following lines in the public unspaced-column case.
Its MAE improves **10.423→1.395**, SSIM **0.373800→0.959693**. Its existing
manifest now requires MAE≤2, SSIM≥0.95 and foreground recall≥0.90.
Single/four-plus columns, rotated paths and no-wrap retain their prior fallbacks.
Font measurements, column balancing and other typography ports remain approximate.

Frozen runtime **01aadc1c** passes one unfiltered original catalogue invocation:
**2253 registered /2252 passed /0 failed /1 optional private-document skip**,
including **158 typography /959 DOCX** methods. Ten targeted Linux regressions
pass. Windows builds are clean; the git-free Linux archive retains two expected
SourceLink warnings. Two unchanged final regressions fail against exact accepted
E5; the single/four-column/no-wrap guard passes. Exact local **0.1.5** package
smoke: `b427062c511347249f46f27420b8ca2d`.

Twenty-nine independent Office controls yield **ten paired MAE/SSIM improvements**
and nineteen identical guard rasters. Across 95 public typography cases, 94 PDFs
retain accepted bytes and the unspaced-column case improves. Every fresh parent
PDF matches its original passing gate output. Fresh frozen candidate PDFs pin
the separately compared prototype evidence with its original library provenance.
All 39 prior SVG decks retain E5 PDF bytes. All 2911 inherited DOCX PDFs and 80
comparison PDFs retain bytes; the 34 cached markup cases retain 508 failed gates.

The broader four-column admission worsened one later Cambria Math control and is
excluded. Fifteen additional Office controls show that explicit zero disables
emitted kerning, but changing the shared layout style regressed three wrapped
controls; that combined source and its comparisons are archived as rejected.
It requires a separate wrapping/emission design and is outside this milestone.
The build-server access failure, excluded stale-binary comparison run and corrected
synthetic-font pair setup remain recorded. Evidence:
`artifacts/plan-revision-20261005/rv08-t1/`. Version remains 0.1.5; release
preparation is deferred. Frozen tested/package and documentation integration
identities remain distinct.

## RV07-E6: isolated SVG container opacity (2026-10-10)

Numeric partial opacity on the outer SVG and group containers now composites
overlapping supported children once. Each container paints into an isolated,
non-knockout DeviceRGB PDF Form, then applies alpha at its invocation. The Form
uses the picture viewport and the existing outer crop/rotation/flip transform.
Nesting is bounded at 32. Partial CSS/percentage container alpha, nested SVG
viewport alpha and excluded deeper containers retain their diagnostic fallbacks.
Malformed attributes retain their prior behavior. Path-level fill/stroke
composition, varying stop alpha and gradient strokes remain partial.

The selected reference remains PowerPoint's smooth PNG preview. Twenty-eight
independent Office controls produce **24 paired MAE/SSIM improvements and four
guard raster identities**. Basic overlapping group/root controls improve MAE
**11.537→0.017**, SSIM **0.796701→0.999992**. Nested/sibling containers, child
alpha, gradients, strokes, picture transforms, crop, stretch, colored backdrops
and viewport overflow are covered. Both rotated-picture controls improve, but
their existing clipping error remains approximate. Five original-syntax guards
retain parent rasters. Actual slide relationships pin imported SVGs separately;
Office deduplicates 20 slides into 19 SVG media parts and normalizes some styles.
Imported numeric alpha does not establish direct CSS/percentage syntax support.

Forms use local resource namespaces and the existing shared content spill store.
Their payloads admit once before retained string snapshots and stream through
64KiB scratch. Page and group streams share the resident window. Staged descriptors
are blanked; weak source identities avoid keeping previous pages' payload owners
alive. The retained strong-map prototype fails the same ownership regression,
whose embedded PDB source checksum matches the frozen infrastructure test file.
Its prototype test binary remains pinned separately from the frozen test assembly.
Rollback removes content and group resources together. Structural validation,
exact/zero content and output limits, spill identity, cancellation cleanup and
independent resource names are covered by fourteen new infrastructure regressions.

Frozen runtime **f637d302** passes one unfiltered complete catalogue invocation:
**2270 registered /2269 passed /0 failed /1 optional private-document skip**,
including **79 PDF /113 SVG /134 images /158 typography /959 DOCX** methods.
Linux passes 73 PDF methods and thirteen targeted methods, including all fourteen
new infrastructure regressions; six existing PDF checks skip for unavailable
Windows Arial files. The git-free Linux archive retains two expected SourceLink
warnings. Three unchanged frozen SVG regressions fail against exact accepted T1;
the unqualified-alpha guard passes. Windows Release/inspector builds are clean.
Exact local **0.1.5** package smoke: `da07dc0897b146289ca633e2845e65a6`.

Across 41 prior SVG decks, 39 retain parent PDF bytes. Two earlier opacity decks
improve on four affected pages; their other seventeen rasters retain bytes.
All 95 typography, 2911 inherited DOCX and 80 DOCX comparison PDFs freshly retain
accepted bytes. The 34 cached markup cases retain 508 failed gates. Fresh frozen
PDF identities pin independently compared prototype evidence with its original
DLL, raster and Office preview provenance. PDF graphics-inspector traversal of
the new general Forms is not claimed. Input/reference/report hashes, rejected
assertion/filter attempts and the strong-map negative proof remain under
`artifacts/plan-revision-20261005/rv07-e6/`. The public inventory retains 344
manifests across ten families. Version remains 0.1.5; release preparation is deferred.

## RV07-E7: SVG picture viewport clipping (2026-10-10)

The SVG picture viewport now clips after its picture transform. Previously the
clip stayed in page coordinates while the picture rotated, losing valid content
outside its original rectangle. The same viewport rectangle, crop, paint, alpha
and PDF resources remain. This follows the established raster-picture clipping
order. It resolves the rotated-picture residual measured during E6 qualification.

Twenty independent Office controls produce **thirteen paired MAE/SSIM improvements
and seven guard raster identities**. E6's 30° control improves MAE **2.591→0.016**,
SSIM **0.851956→0.999887**; its 90° control improves MAE **4.771→0.021**, SSIM
**0.676703→0.999630**. Additional opaque/isolated rotations, portrait pictures,
combined crop/flip/rotation and PowerPoint groups improve. Half-turn, plain
flip/crop/stretch and backdrop/gradient guards retain their rasters. The selected
reference remains smooth Office PNG preview, with trusted input/reference hashes.

Two portable regressions independently interpret the emitted device-space clip,
checking the quarter-turn footprint and oblique centre/projected extent for
opaque and isolated paint. Both unchanged frozen regressions fail against exact
accepted E6; the unqualified-alpha guard passes. Frozen runtime **f5e3efa1** passes
one unfiltered complete catalogue invocation: **2272 registered /2271 passed
/0 failed /1 optional private-document skip**, including **79 PDF /115 SVG
/136 images /158 typography /959 DOCX** methods. Linux passes 73 PDF methods
and fifteen targeted methods; six existing PDF checks require Windows Arial files.
Windows builds are clean; the git-free Linux archive has two expected SourceLink
warnings. Exact local **0.1.5** package smoke: `f91e1cf331d44f418bec4e5e2afaf6ab`.

Across 44 prior SVG decks, 43 retain accepted PDF bytes. The earlier eight-slide
E6 picture-control deck improves on its two rotated pages; the other six rasters
retain identity. Those pages are counted once in the twenty-control summary.
All 95 typography, 2911 inherited DOCX and 80 DOCX comparison PDFs freshly retain
accepted bytes; inherited page counts do not change and cached markup failures
remain 508 across 34 cases. Fresh frozen Office PDFs match independently compared
prototype bytes, pinning the original comparison DLL and raster provenance.
The public inventory remains 344 manifests across ten families. Evidence:
`artifacts/plan-revision-20261005/rv07-e7/`. Other SVG alpha/composition/stroke
limits remain documented. Version0.1.5 and deferred release preparation remain.

## RV07-E8: bounded varying SVG radial alpha (2026-10-10)

Numeric varying stop-opacity on admitted radial pad fills uses one vector
luminosity mask. Equal RGB components interpolate eight-bit alpha; fill/node
alpha multiplies that mask. Its local Form binding shares existing validation,
depth32, content admission, resident spill window, chunked output and rollback.
Graphics-state descriptors retain local names; weak staging identities release
earlier source payload owners. No public API or dependency changes.

Admission covers 2..256 usable stops with representable native radii/intervals,
identity gradient transforms, no off-center focus and identity user-space path
transforms. The mask inherits the picture/group transform when its state is
established, before color-shading transforms. It restores before independent
strokes and later paint. At container depth32, retain the existing container
composition and diagnose the omitted mask. Transformed/focal, linear, sampled,
repeated/reflected, malformed/CSS/percentage and other excluded alpha retain
fallback. Existing mapped color geometry and path-level composition remain partial.

Thirty-one Office controls produce **22 paired MAE/SSIM improvements and nine
guard raster identities**. Basic alpha improves MAE **4.528→0.141**, SSIM
**0.957268→0.999988**; transparent edges improve MAE **16.938→0.149**, SSIM
**0.426412→0.999882**. Backdrops, interior stops, padded endpoints, user-space,
crop/flip/rotation, stretched pictures, path clipping, nested overlap, strokes
and later paint are covered. The selected reference remains smooth Office PNG
preview. Original syntax repeats these controls with the same 22 improvements
and nine guards; 29 rasters match imported syntax. Two source-geometry variants
retain their own comparisons. These are the same controls, not 31 more admissions.

The broader transformed-gradient prototype worsens MAE/SSIM and is rejected.
Initial black/color controls inherited a white master background; corrected
references disable FollowMasterBackground and verify their corner pixels. Both
failed attempts remain archived. Varying focal alpha produces an off-center
Office preview; the chosen centered policy and that alpha fallback are retained.

Six new portable infrastructure methods cover independent namespaces, missing or
ambiguous mask bindings, exact content/output limits, rollback, large spills,
streaming owner release and cancellation cleanup. Four SVG methods cover mask
stops/products, state restoration, fallbacks and stop/depth boundaries, including
mixed admitted/excluded uses. Two unchanged frozen semantic regressions fail
against exact accepted E7; the geometry fallback guard passes on both.

Frozen runtime **d2c7f1cc** passes one unfiltered complete
catalogue invocation: **2282 registered /2281 passed /0 failed /1 optional
private-document skip**, including **85 PDF /119 SVG /140 images /158 typography
/959 DOCX** methods. Linux passes **79 PDF +19 targeted methods** with six existing
Windows-Arial skips and two expected git-free SourceLink warnings. Windows builds
are clean. Exact local **0.1.5** package smoke: `d1bd15b06b624e2ea7b3df9a6054c686`.

Across 45 prior SVG decks, 44 retain PDF bytes. One earlier 15-page deck improves
on two numeric-alpha pages and retains thirteen guard rasters. Its basic varying
control repeats the new set and is counted once there; another Office-normalized
numeric-alpha page also improves. Direct CSS/percentage syntax stays excluded.
All 95 typography, 2911 inherited DOCX and 80 DOCX comparison PDFs freshly retain
accepted bytes; page counts stay unchanged and the 34 cached markup cases retain
508 failed gates. Fresh frozen Office PDF identities pin the original prototype
DLL/Office-preview/raster evidence. General Form graphics-inspector traversal is
not claimed. Public inventory remains 344 manifests across ten families.

Evidence: `artifacts/plan-revision-20261005/rv07-e8/`. Keep0.1.5; release preparation
remains deferred. Other documented renderer gaps remain.

## RV07-E9: bounded varying SVG linear alpha (2026-10-10)

Admitted numeric varying linear-pad fills use native axial color and a vector
luminosity mask. Uniform/opaque gradients retain prior sampled/scalar bytes.
The existing E8 Form validation, depth, admission, spill, chunked output, rollback
and weak ownership remain unchanged; no PDF writer/public API/dependency changes.
Both color and equal-channel eight-bit alpha use the same native [0,1] axis.
Fill/node alpha multiplies the mask; state restores before stroke/later paint.

Object-box gradients project in normalized box units; user-space gradients retain
source units. Their vector and perpendicular axis map through the viewport,
including anisotropic scale and Y inversion. Admission uses actual serialized
three-decimal matrix/corner values and bounds inverse-projection error at all
four path-bounds corners to **0.001**. That error is affine, so corner checks bound
the rectangle. Reject nonfinite/collapsed/imprecise matrices. Admit2..256usable
stops with representable distinct intervals/endpoints, forward vector components,
identity gradient transforms, identity user-space or positive axis box path
transforms, and total Form/mask depth32. Reverse vectors, rotated/sheared/reflected
box paths, nonidentity user-space paths, repeat/reflect, near/duplicate stops,
excess count/depth and unqualified syntax retain prior diagnostic behavior.

Thirty-seven Office PNG-preview controls produce **24 paired MAE/SSIM improvements
and thirteen guard raster identities**. Horizontal alpha improves MAE
**8.610→0.120**, SSIM **0.876283→0.999984**; diagonal alpha improves MAE
**9.228→0.117**, SSIM **0.884345→0.999982**. Stretched diagonals, user units,
transparent/interior/padded stops, black/color/internal backdrops, positive box
scales, clipped paths, picture crop/flip/rotation, nested overlap, strokes and
later paint are covered. Near-zero alpha matches the Office raster exactly.
Original syntax repeats those37controls with24improvements/thirteen guards;
34rasters match imported syntax. Three source-geometry variants retain their own
comparisons. These are the same controls, not37additional admissions.

The first source-space projection prototype leaves diagonal MAE2.322/SSIM.985199
and is excluded. Its source patch, comparison-library hashes and PDF/raster
outputs remain archived; the original prototype DLL was replaced by a corrected
build. The reversed-gradient Office preview is flat; native SVG reversal differs.
That case retains prior sampling pending a separately qualified Office rule.

Five portable methods cover native shading/mask/product, serialized box/user-unit
projection under stretch, state restoration, syntax/transform/precision fallbacks,
stop/depth boundaries and mixed uses. Three unchanged frozen semantic regressions
fail exact accepted E8; the geometry fallback guard passes on both.

Frozen runtime **769ce13f** passes one unfiltered complete
catalogue invocation: **2287 registered /2286 passed /0 failed /1 optional
private-document skip**, including **85 PDF /124 SVG /145 images /158 typography
/959 DOCX** methods. Linux passes **79 PDF +24 targeted methods**, with six existing
Windows-Arial skips and two expected git-free SourceLink warnings. Windows builds
are clean. Exact local **0.1.5** package smoke: `ccdfb1a98e71464ca3c4b8e105d4a9c2`.

All49priorSVG/95typography/2911inheritedDOCX/80DOCXcomparison PDFs freshly retain
accepted bytes. Inherited page counts remain;34cached markup cases retain508failed
gates. Frozen Office PDF identities preserve original prototype/library/Office
PNG/raster provenance. General Form graphics-inspector traversal is not claimed.
Public inventory remains344manifests/ten families. Evidence:
`artifacts/plan-revision-20261005/rv07-e9/`. Keep0.1.5;release preparation deferred.
Other mapped geometry, stroke/path-level composition and renderer gaps remain.

## RV08-T2: bounded PPTX no-wrap zero kerning (2026-10-10)

Explicit/inherited `kern="0"` now disables pair kerning for plain left-aligned,
non-bullet shape paragraphs in unrotated horizontal, single-column,
`wrap="none"`/`noAutofit` frames. Previously it enabled pair adjustments, so
nominal rows contracted toward the positively kerned row. Both XML and scene-fed
run cascades now resolve the admitted setting before measurement and emission.
Positive thresholds and absence retain their rules. Wrapped, centered/right/
justified, table/bullet, autofit, rotated, vertical and multiple-column contexts
retain their prior approximation. No emission flag, PDF writer, public API or
dependency change survives the investigation.

Fifty-four PowerPoint PDF controls produce **20 paired MAE/SSIM improvements and
34 guard raster identities**. They cover9..24point text, five font families,
tracking, inherited/default overrides, fragmented and mixed runs, manual breaks,
underline/highlight and per-paragraph alignment guards. The24pt primary control's
2400 threshold is active at equality; it is an active guard. Twenty-four wrapped
triplet pages and29earlier T1 control pages retain whole PDF bytes separately.

Broad emission-only separation regresses Arial9 even when its line breaks stay
unchanged. Broad no-wrap alignment regresses centered Cambria. Both rejected
source patches, exact binaries/PDBs and original comparisons are archived.
The observed Office residual spacing cannot be identified with pair kerning
alone. Font advances, baselines and other text layout remain approximate.

New reproducible public case
`pptx-ladder-04-typography-no-wrap-zero-kerning-probe` improves MAE
**1.508→0.617**, SSIM **.715052→.921390**, foreground recall **.705270→.898226**.
Its approximate gate requiresMAE≤.8/SSIM≥.90/recall≥.88. Three portable synthetic
font regressions cover nominal/active glyph advances, cascade precedence,
per-paragraph admission and thirteen exclusion byte guards. Two unchanged
frozen semantics fail exact accepted E9; the exclusion guard passes.

Frozen runtime **2e8a0409** passes one unfiltered catalogue:
**2290 registered /2289 passed /0 failed /1 optional private-document skip**,
including **85PDF /124SVG /145images /161typography /959DOCX** methods.
Linux passes **79PDF +27targeted =106methods**, with six existing Windows-Arial
skips and two expected git-free SourceLink warnings. Windows builds are clean;
exact local0.1.5 package smoke `205a6398a22b4888b05dac23ae7d6c5c` passes.

All53priorSVG/95typography/2911inheritedDOCX/80DOCXcomparison PDFs freshly retain
bytes. No inherited page-count changes;34cached markup cases retain508failed
gates. Original Office/raster/comparison-library evidence transfers only through
fresh exact frozen PDF identity. Inventory:345public manifests/ten families.
The initial generated table URI and synthetic disabled-threshold errors remain
retained excluded attempts; corrected controls pass. Accidental generator-prelude
fixture rewrites were restored in the owned worktree before freezing.
Evidence: `artifacts/plan-revision-20261005/rv08-t2/`. Keep0.1.5;release deferred.

## RV07-E10: bounded SVG linear gradient strokes (2026-10-10)

Supported gradient rims previously disappeared with an unsupported-stroke warning.
Uniformly numeric linear-pad strokes now use native axial PDF shading patterns.
Their matrices map into the parent stream's initial user space independently of
later path transforms, following PDF 32000-1 section 8.7.2. Page and isolated Form
pattern resources share validation, snapshot/rollback, shading value deduplication,
weak staging, the page spill window and existing output/admission limits.
The public API, dependencies and version remain unchanged.

Admission requires unrotated/unflipped pictures, forward axes, identity gradient
transforms, 2..256 usable stops, represented intervals at least .001, and serialized
projection error at most .001. User-space paths require identity transforms;
positive axis box transforms remain approximate. Plain numeric node/stroke/stop
alpha is supported. Partial CSS/percentage container opacity, nested SVG viewports,
excess depth and filled elements with partial node opacity retain diagnosed omission.
Radial, varying-alpha, reverse and repeat/reflect strokes also retain omission.
Supported root/group isolation uses local pattern resources; fill masks restore
before strokes and later paint. General Form graphics-inspector traversal is not added.

Forty-six independent Office controls produce **34 imported paired MAE/SSIM
improvements and twelve raster identities**. Original syntax produces **32 paired
improvements and fourteen identities**. Office normalizes percentage stops and CSS
root opacity; imported and original SVGs are pinned separately through slide
relationships. These are 46 controls, not 92 independent examples. Basic rectangle
MAE improves **4.428→.029**, SSIM **.045320→.999986**. Ellipses, opaque filled shapes,
curves, joins, forward/diagonal axes, user units, viewport stretch, uniform alpha,
group overlap, fill-mask restoration, later solid paint and dash caps/phase are covered.
Scaled box-path source syntax retains MAE up to **1.126 / SSIM .766568**; negative
square dash phase retains **SSIM .980362**. These material improvements do not
establish general SVG stroke parity. Transparent zero-alpha paint retains a raster;
whole-PDF identity is not inferred from that result.

Initial solid square-cap compensation worsens gradient gaps and is now limited to
solid paint. A combined fill/stroke B attempt worsens the opaque filled inner rim
from MAE .029 / SSIM .999712 to 1.147/.884237. Separate paint remains; translucent
filled-node isolation needs another design. The rejected source, exact binaries,
comparisons and narrower opacity attempts remain archived.

The reproducible one-page public case is
`pptx-ladder-07-svg-linear-gradient-stroke`; it repeats the primary rectangle.
Regenerate with `tools/NewSvgStrokeRadialFixtures.ps1 -ProbeSet linear-stroke-controls
-OutputPath tests/Lokad.OoxPdf.Tests/Cases/pptx-svg-linear-gradient-stroke.pptx`.
Its approximate gate requires MAE≤.1 / SSIM≥.999 / histogram≥.995 / recall≥.99.
It passes Office preview (.029/.999986) and PDF export (.038/.999931), recall 1.
PNG previews remain separate from Office PDF rasters, preserving the selected policy.

Six portable PDF methods cover local bindings, shading deduplication, rollback,
defensive snapshots, invalid names/printed geometry, spill, exact/zero admission,
cancellation cleanup and release of prior Form owners. Six portable SVG methods
cover alpha, source-unit/stretch projection, nested Forms, state restoration,
square/round dash paint, eighteen fallback byte controls and the 256-stop boundary.
All six unchanged native semantic methods fail exact accepted T3/T2 runtime;
the exclusion checks reach the final native boundary assertion on that parent.

Frozen runtime **0d3523a1** passes one unfiltered catalogue:
**2302 registered /2301 passed /0 failed /1 optional private-document skip**,
including **91 PDF /130 SVG /151 images /161 typography /959 DOCX** methods.
Linux passes **85 PDF +33 targeted =118 methods**, with six existing Windows-Arial
skips and two expected git-free SourceLink warnings. Windows builds are clean.
Exact local **0.1.5** package smoke `0222aff23de74b35ad6ab427d77171bd` passes.

All **53 prior SVG /96 current typography /2911 inherited DOCX /80 DOCX comparison**
PDFs freshly retain accepted bytes. The typography set includes all four corrected
T3 fixtures and the T2 zero-kerning case. Inherited page counts remain; 34 cached
markup cases retain 508 failed gates through exact PDF identity, without rerunning
their reference gates. The T3 boundary raw-operation gates remain partial.
Fresh frozen PDF identities transfer the seven Office/public deck comparisons while
preserving their original source/library/Office/raster provenance. Inventory:
346 public manifests across ten families, validated on Windows and Linux.
Evidence: `artifacts/plan-revision-20261005/rv07-e10/`. Keep 0.1.5; release deferred.

## RV08-T4: PowerPoint numbering and label clearance (2026-10-10)

Wide numbering labels previously overlapped the paragraph body, and each explicit
start value restarted independently. Bounded known-format lists now continue when
the effective format/start settings match, keep per-level sequences, and reset on
changed settings or plain paragraphs. Advancing a parent clears deeper sequences.
Empty numbered paragraphs update settings without emitting or advancing a label;
space-only paragraphs retain numbering. The same resolved model drives XML and scene
text. Qualified first fragments use the body font and nominal label advances,
including authored trailing character spacing. Continuation lines keep their margin.

Numbering admission covers the eleven existing Arabic/alphabetic/Roman format
spellings, levels0..8 and valid starts1..32767, default text-sized/font-following
labels outside tables. Explicit bullet-size/font overrides and unsupported settings
retain fallback. Label-clearance admission additionally requires left alignment,
horizontal single-column non-table text, no tabs or normal autofit, no rotation or
flip, nonpositive hanging indent and a label that fits the frame. Other frames keep
their label-placement approximation even when their supported sequence is corrected.

Seventy-five independent public Office controls, covering36 primary,31 held-out and
eight empty/space variants, are pinned in both original and imported syntax. Each
syntax yields72 paired MAE/SSIM improvements, two override raster identities and
one mixed-metric residual. The extreme50pt frame corrects numbering and lowers MAE
1.1824 to1.1538, while SSIM declines.317261 to.315023; its width remains excluded
from the new clearance rule. This result establishes no general wrapping parity.
Positive indents, manual-break pitch, Calibri metrics, overrides and other constrained
layout remain partial. Rotated Office text is judged through rasters because the
page-level inspector does not traverse its Form text.

Across96 current public typography cases,92 PDFs retain exact bytes. All four
changed cases improve both raster metrics. Arabic numbering improves MAE.9575 to
.0422 and SSIM.484789 to.992995; alphabetic and Roman lists reach.996440 and.996562.
Those three existing approximate gates are strengthened to MAE<=.1 and SSIM>=.99,
without changing fixtures or classifications. The fourth change is a small
improvement in the synthetic-bold numbered probe.

Six portable synthetic-font methods cover sequence/settings, nested and inherited
lists, empty paragraphs, nominal label advances and spacing, continuation indent,
XML/scene agreement and fallback. Five unchanged methods fail the exact parent;
the fallback method passes. Initial kerned-label and pre-boundary prototypes retain
their exact source and binaries; no failing model is admitted incidentally.

Frozen Windows qualification passes2307 methods, zero failures and one optional
private-document skip, from2308 registered methods. It includes91 PDF,130 SVG,
151 image,167 typography and959 DOCX checks. Linux passes85 PDF plus39 targeted
checks, with six existing Windows-Arial skips and two expected archive SourceLink
warnings. Windows/Linux inventories validate346 cases across ten families.
Fresh conversions retain60 SVG PDFs, five related typography-control PDFs,
2911 inherited DOCX PDFs and80 DOCX comparison PDFs, with no inherited page changes.
The34 cached DOCX cases and508 failed gates transfer through fresh PDF identity;
no new reference-cache gate is claimed. The0.1.5 package smoke succeeds. Dependencies,
public API and version are unchanged; release preparation remains deferred.

## RV08-T5: generator-proven accent fixture repairs (2026-10-10)

Dense-column and whitespace-controls each contained one text node with UTF-8
accent bytes decoded as Windows-1252. Their existing typography generator authors
the intended French accents. Those two nodes now match that text. The whitespace
probe retains its intact nonbreaking and narrow spaces. All other ZIP entries and
layout/style settings are unchanged; generator and runtime source are unchanged.
Regenerating only those two probes reproduces every ZIP payload after XML line-ending
normalization. Raw payload identities differ only in platform line endings.

Two fresh Office references are pinned along with independently imported inputs.
Original and imported corrected candidate PDFs are identical per case. The dense
probe retains nine text operations, MAE1.570659 and SSIM.558875; whitespace-controls
retains thirteen, MAE.419306 and SSIM.953478. Their decoded accents are correct, but
spacing and baseline differences remain. The new Office targets supersede malformed
fixture references; these metrics are no claim of a renderer improvement across
different inputs. Existing approximate gates/classifications are unchanged.

The167 typography checks pass against the corrected working fixtures using the
previously qualified runtime/test assemblies. Windows and Linux validate346 public
cases across ten families. RV08-T4's full runtime/Linux/package qualification
remains applicable through unchanged source and assembly identities. No full suite
or package rerun is needed for these two text-only fixture changes. Version0.1.5
and deferred release preparation are retained.

## RV08-T6: bounded PowerPoint no-wrap distribution (2026-10-10)

Plain distributed text in a no-wrap/noAutofit frame previously stretched across
the frame. Office retains nominal advances for the qualified horizontal,
single-column, unrotated/unflipped, non-table ASCII paragraphs. The resolved
paragraph rule now suppresses stretching without changing active kerning,
normal justification or wrapping. Bullets, tabs, manual breaks, non-ASCII text,
other frame modes and other alignment values retain existing behavior.

Thirty primary and36 independent font/size/spacing/inheritance/width/exclusion
controls are pinned in original and Office-imported syntax. Each syntax yields
19 paired MAE/SSIM improvements and47 raster identities, with no mixed results.
The reproducible public no-wrap case repeats the primary geometry: MAE improves
.8009 to.0298 and SSIM.411048 to.995125. Its approximate gate requires MAE<=.1,
SSIM>=.99 and foreground recall>=.95. Measured recall.957897 failed the provisional
.98 bound before qualification; glyph-edge differences remain. This case is not
counted as an additional independent control or exact font parity.

Four portable synthetic-font methods cover nominal PDF placement and active
kerning, inheritance and mixed paragraphs, XML/scene agreement, and excluded
frame/script/manual-break fallback. Three unchanged semantic methods fail the
exact T5 parent; fallback passes. The rejected broader word-spacing model retains
its exact source, binary, PDB and comparisons. Grouped Office text operations can
contain per-glyph TJ adjustments; operation counts alone establish no spacing rule.

Frozen Windows clean Release passes2311 methods, zero failures and one optional
private-document skip from2312 registered methods, including91 PDF,130 SVG,
151 image,171 typography and959 DOCX checks. Linux passes85 PDF plus43 targeted
checks, with six existing Windows-Arial skips and two archive SourceLink warnings.
Both inventories validate347 cases across ten families. Fresh conversions retain
all96 existing typography PDFs,60 SVG PDFs,11 related control-deck PDFs,
2911 inherited DOCX PDFs and80 DOCX comparison PDFs, with no inherited page changes.
The34 cached DOCX cases and508 failed gates transfer through fresh PDF identity;
no new reference-cache gate is claimed. All six candidate Office deck PDFs retain
their compared bytes after the source freeze. Packed0.1.5 smoke passes. API,
dependencies and version stay unchanged; release preparation remains deferred.

## RV08-T7: numbering overrides, alphabetic labels and body pitch (2026-10-10)

Mixed-level lists with bullet font/size overrides previously shared a global
counter, producing1–7 where Office uses parent/child sequences1/1/2/2/1/1/3.
Known-format valid-level/start non-table numbering now uses per-level sequences
independently from label geometry. Overridden font/size/kerning and first-fragment
X remain; default label-clearance admission is unchanged. Run-free plain paragraphs
clear the qualified sequence map. A larger qualified number label no longer enlarges
body-font paragraph pitch:36pt labels over24pt body retain28.8pt advances, replacing
43.2pt advances. Character/table/unsupported numbering retains its prior height rule.

Office alphabetic labels repeat letters (z,aa,bb; zz,aaa,bbb), with a1..30repeat cycle.
Controls pin the779/780/781,806/807 and1560/1561 boundaries, larger1024/2048/4096/
8192/16384 starts and32767 continuation. Formatting allocates at most30letters in
the qualified branch. Table/malformed fallback and shared DOCX alphabetic formatting
remain unchanged. Earlier spreadsheet-style alpha expectations are corrected against
these independent Office controls. Counter-only and pre-empty-boundary prototypes
retain exact source/binaries and comparisons; neither is qualified incidentally.

86 authored controls include one duplicate singleton32767 boundary, so85 distinct
controls are counted per syntax. Original syntax gives80paired MAE/SSIM improvements,
three raster identities and two metric tradeoffs; import gives79improvements, four
identities and the same tradeoffs. All84inspectable text sequences match Office;
rotated text has no page-level Office operations and improves by raster. The Times
equal-start case raises MAE1.35982→1.36326 and SSIM.389113→.390548. The30letter Y case
raises MAE1.43121→1.59600 and SSIM.267517→.303779 while correcting its label. No
both-metric regression remains. Wrapping, clipping, normal autofit and centered/bottom
anchor controls improve. Label-clearance/font residuals remain: in the36pt-label probe,
Office body starts120.02/140.06pt, while candidate retains108pt; baselines now agree
within.04pt. This batch establishes no complete override-label layout parity.

The existing public mixed-list case now matches decoded labels and improves
MAE.4991→.4622, SSIM.962934→.968632. Its approximate gate strengthens to MAE<=.6 and
SSIM>=.965; fixture and classification stay unchanged. All96other public PDFs retain
bytes. Five new portable methods cover counters, inheritance/style/empty transitions,
authored label geometry, alphabetic cycles and body pitch. Four unchanged new methods
fail exact T6 parent; the label-style/placement guard passes.

Frozen Windows clean Release/full2317registered/2316passed/0failed/1optional skip
includes91PDF/130SVG/151images/176typography/959DOCX. Linux85PDF+48targeted=133passes,
six existing Windows-Arial skips, two archive SourceLink warnings. Both inventories
validate347cases/10families. Ten new control-deck PDFs retain their compared bytes
after freeze. Fresh60SVG/2911inheritedDOCX/80comparison PDFs retain bytes, with no
inherited page changes. Fourteen of17related guard decks retain bytes; three older
numbering decks improve on five pages with98other raster identities. The34cached
DOCXcases/508failed gates transfer through fresh PDF identity; no new cache gate is
claimed. Packed0.1.5smoke passes; API/dependencies/version stay unchanged and release
preparation remains deferred. The content audit explicitly replaces four stale T3
pre-maintenance reference paths with their pinned qualified repaired targets.

## RV08-T8: automatic-number font and first-fragment clearance (2026-10-10)

Font/size-overridden number labels previously left the first body fragment at
the authored margin even when it overlapped the label. Independent Office16
controls also show that automatic labels follow the first body run font despite
an authored bullet font; size-only labels follow that face too. The bounded path
now uses the body face and nominal glyph advances, preserves authored label size,
and includes trailing character spacing in clearance. A36pt “100.” over24pt Arial
moves the candidate body from108pt to160.066pt, compared with Office159.98pt.
Roman IV controls independently distinguish nominal advances from pair kerning.

The existing left/horizontal/single-column/unrotated/unflipped/non-table/no-tab/
no-normal-autofit/nonpositive-hanging-indent/label-fits-frame bounds remain. For
overrides the first drawable body fragment must also fit after the label. The
initial label-only-fit prototype regresses a50pt frame on both metrics and is
retained with exact source/DLL/PDB/report. Default-label admission retains bytes.
Continuation margins and body-font paragraph pitch stay intact. Character bullets,
table/unsupported numbering and excluded frames retain font/kerning/placement
fallback. Wrapped and manual-break flow remain partial even where metrics improve.

48new independently authored controls plus85priorT7controls give133distinct
controls per syntax. One repeated singleton boundary remains counted once.
Original and Office-imported syntax each yield87paired MAE/SSIM improvements and
46raster identities, with no metric tradeoffs or both-metric regression. All131
inspectable text sequences match; two rotated references lack page-level text
operations and retain rasters. Four new portable methods cover body-face/size,
nominal advances/trailing spacing, continuation/scene agreement and fallback.
Three unchanged new methods fail exact T7 parent; the fallback guard passes.

The public mixed-list fixture retains its input and decoded text, improving
MAE.4622→.4194 and SSIM.968632→.973070. Its existing approximate gate strengthens
to MAE≤.45/SSIM≥.97; the changed-pixel bound remains. Classification stays approximate.
The other96current typography PDFs freshly retain accepted bytes. After freeze,
all14compared control-deck PDFs,60priorSVG and97current typography PDFs reproduce
their qualified bytes. Fifteen of17older guard decks retain bytes; two numbering
decks improve on four pages, with68other raster identities.

Frozen Windows clean Release/full2321registered/2320passed/0failed/1optional skip
includes91PDF/130SVG/151images/180typography/959DOCX. Linux85PDF+52targeted=137passes,
six existing Windows-Arial skips, two archive SourceLink warnings. Both inventories
validate347cases/10families. Fresh2911inheritedDOCX and80comparison PDFs retain
bytes, with no inherited page changes. The34cachedDOCXcases/508failed gates transfer
only through fresh PDF identity; no new cache gate is claimed. Packed0.1.5smoke
passes. API/dependencies/version stay unchanged; release preparation is deferred.

## RV08-T9: bounded automatic-number word wrapping (2026-10-10)

The general bullet allowance lets words exceed the authored line width by 20% of
font size. In a 190pt numbered textbox this joins “compare stock” where Office
breaks after “compare”, omitting a 28.8pt line and moving later paragraphs upward.
Already-admitted automatic labels now use coordinate word-fit tolerance for
wrapping with explicit noAutofit and printable ASCII text runs without fields or
manual breaks. The first drawable body fragment must fit after its actual origin.
Existing frame, indentation, font/size, label-clearance and counter rules remain.
Character bullets retain the general allowance. Non-ASCII, fields/manual breaks,
insufficient first-word space and excluded frame/autofit modes retain fallback.

64 new controls comprise 42 width/font/format/split/guard cases, 16 heldout font-
size/first-fragment boundaries and six additional first-word edges. Together with
133 prior controls, there are 197 independent controls per syntax; a repeated
singleton boundary is counted once. Original and Office-imported syntax each
give 16 paired MAE/SSIM improvements and 181 raster identities, with no metric
tradeoffs. Every changed word-line sequence matches Office; 194 inspectable full
text sequences match. Three rotated references lack page-level operations and
retain rasters. The initial strict rule regresses five first-word edges per syntax
by adding an empty body line; exact source/DLL/PDB/reports retain that failed model.
The first-fragment fit bound removes those regressions. Four new portable methods
cover line boundaries, split words/scene agreement, counters/body pitch and
character/excluded/oversized-first-word fallback. Three unchanged methods fail
exact T8 parent; the fallback guard passes.

All 97 prior public typography PDFs retain bytes. A new approximate public probe
duplicates one primary geometry and is not counted as another independent control.
Original/imported comparisons both improve MAE 0.6935→0.0923 and SSIM
0.746999→0.985514, with foreground recall 0.970963. Its gates are MAE≤0.12,
SSIM≥0.98 and recall≥0.96. Only this fixture is generated; existing inputs and
gates are unchanged. The three-paragraph control improves MAE 2.7255→0.2125 and
SSIM 0.539201→0.984790; general shaping and manual-break baseline parity remain
outside this rule.

Frozen clean Windows Release and one full catalogue invocation pass 2325 registered
/2324 passed /0 failed /1 optional private-document skip, including 91 PDF, 130 SVG,
151 images, 184 typography and 959 DOCX methods. Linux passes 85 PDF +56 targeted
=141 checks, with six existing Windows-Arial skips and two archive SourceLink
warnings. Both inventories validate 348 cases across ten families. After freeze,
20 control decks plus two public variants reproduce their compared PDFs; 60 SVG
and 98 current typography PDFs reproduce qualified bytes. Fifteen of 17 older
guard decks retain PDFs; two numbering decks improve on two pages, with 70 other
raster identities. Fresh 2911 inherited DOCX and 80 comparison PDFs retain bytes,
with no inherited page changes. The 34 cached DOCX cases and 508 failed gates
transfer only through fresh PDF identity; no new cache gate is claimed. Packed
0.1.5 smoke passes. API, dependencies and version stay unchanged; release
preparation is deferred.


## RV06-L128: section occurrence counts for DOCX notes (2026-10-10)

Word 16 controls establish section-local starts/formats and continuous versus
`eachSect` numbering for footnotes and endnotes. A continuous section adds its
start to the number of earlier automatic occurrences: two earlier notes followed
by start 9 produce 11 and 12, even when those earlier sections restarted at 4.
`eachSect` counts from the current section boundary. Omitted section properties
use decimal footnotes/lowerRoman endnotes and start 1, without inheriting document
starts/formats/restarts. Authored document settings remain inspectable.

Admission covers an entire note kind's simple section sequence independently of
the other kind: valid starts up to 32767, the six existing decimal/Roman/letter
formats, and omitted/continuous/eachSect restarts. Closing section properties apply
to preceding body elements, including table references and paragraphs split at
manual breaks. Custom marks do not advance automatic counters. Each-page settings,
unsupported/malformed section settings, revision bodies and hidden/nested section
boundaries keep the prior fallback. The existing single-section partial override
policy remains available outside this admission.

Fifty independent public controls give 44 supported visible-label and decoded-word
matches, 43 case-average MAE/SSIM improvements and six excluded-control PDF
identities. Word does not encode every paragraph separator as whitespace; decoded
word audits explicitly separate pages and text baselines. Four individual pages
trade pixel metrics despite better case averages. A same-page footnote section
retains the previous one-page candidate versus two Word pages; its labels match.
Layout, font metrics and note spacing remain approximate.

Eight new portable methods pass; seven fail the exact previous runtime and one
fallback guard passes both. Repeated labels keep distinct destinations. Across
52 authored controls, 186 candidate links resolve; 176 links across 49 controls
match actual Office source/destination pages without positional ambiguity. The
same-page pagination case keeps two unmapped reference links. Each custom-mark
control keeps three candidate links versus four Word links; positional matches
are not claimed for those two controls.

Two new approximate public probes duplicate primary geometry and are not extra
independent controls. They have fresh Office references and pass per-page numeric
gates: footnotes MAE <=.42/SSIM >=.80/recall >=.80; endnotes MAE <=.95/SSIM >=.60/
recall >=.56, plus page/dimension and changed-pixel checks. Each gate requires
only its existing DOCX_APPROXIMATED_FOOTNOTE/ENDNOTE warning; approximation
warnings remain enabled because note layout is partial. Regenerating
the fixtures reproduces every ZIP entry payload; ZIP timestamps/order are excluded.

Frozen runtime **e90bac15** passes clean Windows Release and one full catalogue:
2333 registered/2332 passed/0 failed/1 optional private-document skip, including
91 PDF, 130 SVG, 151 images, 184 typography and 967 methods with Docx in their name.
Linux passes 85 PDF +71 targeted checks, with six existing Windows-Arial skips and
two archive SourceLink warnings. Both inventories validate 350 cases in ten families.
All 2911 inherited DOCX PDFs and 251 transfer PDFs retain bytes: 80 DOCX comparisons,
60 SVG, 98 current typography and 13 prior single-section note controls. There are
no inherited page changes. The 34 cached cases/508 failed gates transfer solely
through fresh PDF identity; no new cache gate is claimed. Packed 0.1.5 smoke passes.
API, dependencies and version remain unchanged; release preparation is deferred.


## RV06-L129: simple DOCX section page boundaries (2026-10-10)

Word 16 controls establish that a section's type describes how that section
started: the following section owns the boundary after a closing sectPr.
Omitted types mean nextPage. In simple, equal-geometry, one-column paragraph
sequences, nextPage/continuous boundaries use that following type without
rewriting the authored model. Current-page markers for unplaced, fitting
page-bottom/default footnotes complete the page at a continuous boundary.
Beneath-text notes and endnotes do not force that completion; notes already
placed on an earlier page cannot add another page.

Admission excludes tables, static stories, floats, rich/revision content,
lists, fields, comments, images, geometry changes, multiple columns, nondefault
spacing scales, terminal section elements, parity/column breaks and split or
overflowing notes. Each-page numbering and other note limitations remain.
The general continuous-section and footnote approximation diagnostics remain.

44 authored controls cover 43 independent geometries, including one repeated
L128 geometry counted once. All 40 admitted controls match Word's page count;
14 change page counts, of which 13 are independent. Thirty PDFs retain bytes,
including four table/column/margin/terminal guards with the previous one versus
two Word pages. All 44 decoded word multisets match with explicit page/baseline
separators. Missing/extra baseline pages cannot establish a paired case-average
MAE/SSIM improvement; page recovery/removal and matched-page metrics are recorded.

Five new portable methods pass; three fail the exact previous runtime and two
guard/earlier-page methods pass both. The initial two-column fallback expectation
was corrected from one page to the exact parent's two, without a runtime change.
109 candidate note links resolve across 47 authored controls. The 43 supported
annotation controls, including two public duplicates and the prior recovery,
match all 95 Office source/destination pages; four excluded guards remain partial.

Two public fixtures duplicate primary geometry and have fresh printer-restored
Word references. The continuous-footnote gate requires the existing footnote
and continuous-section warnings, MAE <=.42/SSIM >=.80/recall >=.80 and changed
ratio <=.0055. The boundary-only gate requires no diagnostics, MAE <=.60/
SSIM >=.74/recall >=.76 and changed ratio <=.007. Both require matching page
counts/dimensions. Default and pagination generation reproduces all ZIP entry
payloads; container order/timestamps are excluded.

Frozen runtime **4c5e097f** passes clean Windows Release and one full catalogue:
2338 registered/2337 passed/0 failed/1 optional private-document skip, including
91 PDF, 130 SVG, 151 images, 184 typography and 972 methods containing Docx.
Linux passes 85 PDF +76 targeted checks, with six existing Windows-Arial skips
and two source-archive SourceLink warnings. Both inventories validate 352 cases
in ten families. Fresh PDF identities transfer 44 new comparisons, two public
gates and 51 prior note controls. The remaining prior note control recovers
Word's two-page layout and repeats primary geometry; it adds no independent case.
All 2911 inherited DOCX PDFs and 251 broader transfers retain bytes with no
inherited page changes. The 34 cached cases/508 failed gates transfer solely
through fresh identity; no new cache gate is claimed. Packed 0.1.5 smoke passes.
API, dependencies and version remain unchanged; release preparation is deferred.


## RV06-L130: bounded custom-mark note navigation (2026-10-10)

Explicit `0`/`false`/`off` custom-mark flags restore automatic labels and counters;
raw flags remain inspectable and unknown values retain the previous fallback.
True custom marks retain their complete authored text and do not count toward
automatic labels. A source run with exactly one reference derives its marker
from the first non-whitespace BMP character in the immediately following w:t.
Annotations require exact emitted source run/offset ownership, measurable glyph
advance and a rendered note destination. A prefix in the same run is supported;
the hit area covers the first character even when more authored characters follow.
Tables and wrapped markers use the same actual emission mapping. Automatic
annotations keep their existing behavior.

Empty/whitespace/surrogate, separate-run, before-reference, multiple-reference
and revision/field/comment marks retain the old annotation fallback. Word emits
only one custom link for two references in one XML run. The initial prototype's
four links versus three Word links are retained as rejected evidence; the
qualified guard keeps two candidate links versus three Word links. Other note
flow, each-page numbering and after-spacing hit-area limits remain approximate.

30 primary, 14 held-out and two public controls cover 42 independent controls:
two automatic baselines repeat previous geometry and the public cases duplicate
primary geometry. Thirty independent supported controls plus four duplicates
match all 136 actual Office source/destination pages. Six explicit-false variants
improve case-average MAE/SSIM. Twenty-six annotation-only outputs retain raster
bytes; twelve guards and two automatic baselines retain complete PDFs. All
candidate/parent/Office page counts agree. Forty-four decoded word multisets
match with page/baseline separators; two revision guards retain L128's document
starts (105/106/107) versus Word's section labels (iv/09/10).

Five portable methods pass; four fail exact L129 and one fallback guard passes
both. Two public probes have fresh printer-restored Word references and pass
numeric/page/dimension gates requiring only the corresponding approximation
diagnostic. Footnote gates are MAE <=.42/SSIM >=.80/recall >=.80/changed <=.0055;
endnotes use .95/.60/.56/.010. Their four links also match actual Word source/
destination pages in the direct PDF annotation audit. Default, pagination and
custom generation reproduce every ZIP entry payload; container order/timestamps
are excluded.

Frozen runtime **b64c1fe5** passes clean Windows Release and one full catalogue:
2343 registered/2342 passed/0 failed/1 optional private-document skip, including
91 PDF,130 SVG,151 images,184 typography and977 methods containing Docx.
Linux passes85PDF+81targeted checks, six existing Windows-Arial skips and two
archive SourceLink warnings. Both inventories validate354cases/10families.
Fresh identities transfer46new comparisons and96prior note outputs. Two prior
custom-mark controls gain their fourth link with exact raster identities. Across
48authored annotation controls,178candidate links resolve;36supported controls
including duplicates match all144Office source/destination pages.
All2911inheritedDOCX and251broader transfer PDFs retain bytes with no inherited
page changes.34cached cases/508failed gates transfer only through fresh identity;
no new cache gate is claimed. Packed0.1.5smoke passes. API/dependencies/version
remain unchanged; release preparation is deferred.


## RV06-L131: single-line note hit areas through after-spacing (2026-10-10)

Word16 expands a one-line paragraph's note link through after-spacing; multiline
paragraphs retain line-only slots even when the mark is on their last line.
Actual rendered lines are counted across the full layout by paragraph identity.
Ordinary unscaled, one-column body paragraphs with automatic/exact line boxes,
no before/contextual/auto/line-unit spacing and a fitting lower bound add actual
after-spacing to the existing hit-area height. Rendered glyphs, horizontal bounds
and note destinations are unchanged. Paragraphs without note destinations skip
the line count. Tables, static content, floats, mirror margins, rich/revision
content, lists, hyperlinks and manual/page-break layouts retain prior behavior.

24primary +22held-out +2public controls cover42independent geometries. Four
zero-spacing baselines repeat earlier geometry and public probes duplicate primary
geometry. Twelve independent extended-slot controls plus the public duplicates
match56actualOffice link source/destination pages and vertical bounds within.005pt.
Thirty guards/four baselines retain PDFs; all48parent/candidate rasters match.
Allpage counts agree.46decoded-word multisets match with page/baseline separators;
two revision guards retain L128 document-start numbering rather than Word section
labels. Before-spacing12 expands Word's slot upward and downward; that rule and
auto/line-unit/table/revision bounds remain partial pending separate admission.

An initial null-source-paragraph-index check admitted no ordinary body paragraphs
(their index is0). Its source/library, no-op comparisons and two semantic failures
are retained. Corrected four portable methods pass; two fail exact L130 and two
guard methods pass both. Default/pagination/custom/spacing fixture regeneration
reproduces every ZIP payload; container timestamps/order are excluded.
Two public gates retain existing note approximation warnings and per-page numeric
limits, matching pages/dimensions and separately audited vertical annotations.
Footnote gates are MAE<=.42/SSIM>=.80/recall>=.80/changed<=.0055; endnotes use
.95/.60/.56/.010. Pixel gates establish no new painting improvement claim.

Frozen runtime **0f115006** passes clean Windows Release and one full catalogue:
2347registered/2346passed/0failed/1optionalprivate-document skip, including91PDF,
130SVG,151images,184typography and981methods containingDocx. Linux passes85PDF+
85targeted checks, six existing WindowsArial skips and two archive SourceLink
warnings. Both inventories validate356cases/10families. Fresh identities transfer
48current comparisons and144prior note outputs through192freshPDFidentities.
All2911inheritedDOCX PDFs retain bytes with no inherited page changes. Of251broader
comparisons,248 retain PDFs:77DOCX comparisons,60SVG,98typography,13prior note
controls. Three older direct24/direct8/inherited8pt note controls improve their
actual Office vertical bounds from24/8pt error to.002pt with exact parent rasters,
unchanged destinations/pages and one link each. They add no newly authored controls.
The original strict-transfer failure is retained; qualified changed-output bindings
are separate.34cached cases/508failed gates transfer only through fresh exact
unchangedPDFidentity; no new cache gate is claimed. Packed0.1.5smoke passes.
API/dependencies/version remain unchanged; release preparation is deferred.


## RV06-L132: note hit areas own the unconsumed before-gap (2026-10-10)

Word16 assigns a note-bearing paragraph only the before-gap remaining beyond its
predecessor's after-spacing. With 12pt before-spacing, a preceding 6pt after-gap
leaves 6pt owned above the line; preceding 24/36pt leaves none. Direct/inherited
settings, automatic/exact slots, section starts and empty predecessors are
qualified in ordinary unscaled, one-column paragraphs rendered on one body line.
The rectangle expands upward and retains the qualified after-spacing below.
After-only upper-bound behavior remains unchanged. Wrapped/contextual/rich
predecessors and multiline/contextual/auto/line-unit/scaled/static/table/column/
manual/page-break layouts keep fallback. Glyphs, horizontal bounds and actual
note destinations remain unchanged.

28 primary +12 held-out +2 public controls cover 38 independent geometries. Two
primary cases reuse exact L131 before12 inputs and references; public probes
duplicate primary geometry. Twenty independent controls and four duplicates
change 24 rectangles. All 96 links in those supported controls match Word's
vertical bounds within 0.005pt, including unchanged links. Eighteen guards retain
PDF bytes; all 42 retain rasters, decoded-word multisets and page counts. Four
portable methods pass; three fail exact L131 and one guard passes both. The same
test source is pinned for both builds. The qualified prototype source/DLL/PDB
remains archived. An accidental unchanged-root build is retained and excluded
from prototype qualification. All ten generator ZIP payloads reproduce.

Two new public cases require matching pages/dimensions, only the corresponding
note approximation warning and per-page numeric gates: footnotes use
MAE<=.42/SSIM>=.80/recall>=.80/changed<=.0055; endnotes use .95/.60/.56/.010.
Annotations are audited separately; no new painting improvement is claimed.

Frozen runtime **71de81e8** passes clean Windows Release and one full catalogue:
2351 registered /2350 passed /0 failed /1 optional private-document skip, including
91 PDF, 130 SVG, 151 images, 184 typography and 985 methods containing Docx. Linux
passes 85 PDF +89 targeted methods, with six existing Arial skips and two archive
SourceLink warnings. Both inventories validate 358 cases across ten families.
Fresh PDF identity transfers 42 current comparisons, 206 unchanged prior outputs
and two requalified identical before12 inputs through 250 conversions. All 2911
inherited DOCX and 251 broader comparison PDFs retain bytes and page counts.
The 34 cached markup cases /508 failed gates transfer only through fresh exact
unchanged PDF identity. Local 0.1.5 package smoke passes. API/dependencies/version
remain unchanged; release preparation is deferred.


## RV06-L133: first-line note hit areas in multiline paragraphs (2026-10-10)

In ordinary unscaled, one-column multiline paragraphs, Word assigns a positive
unconsumed before-gap to a note mark on the first rendered line. Marks on later
lines retain their line-only rectangles. No multiline mark includes after-spacing.
The resolved paragraph spacing map retains before/after values for single lines
and only the positive before-gap for multiple lines. Automatic/exact line boxes,
predecessor admission and existing contextual/auto/line-unit/scaled/static/table/
column/manual-break exclusions remain. Glyphs, horizontal bounds and actual note
destinations are unchanged.

24 primary +8 held-out +2 public controls cover 30 independent geometries. Two
primary cases reuse exact L132 multiline inputs and Office references; public
cases duplicate primary geometry. Sixteen independent controls and four repeats
correct 20 rectangles. All 80 links in those supported controls match Word's
vertical bounds within 0.005pt, including unchanged links. Fourteen guards retain
PDF bytes; all 34 retain parent rasters, decoded-word multisets and page counts.
Four portable methods pass; two fail exact L132 and two guards pass both. The
identical test source and qualified prototype source/DLL/PDB are retained. All
twelve generator ZIP payloads reproduce.

The two new public cases remain approximate for text painting. Their unchanged
first-page MAE/SSIM are 2.170/.587753 for footnotes and 2.091/.540705 for endnotes.
They use MAE<=2.30, SSIM>=.57/.53, recall>=.65/.56 and changed<=.024, matching
pages/dimensions and only the corresponding note approximation warning. Separate
annotation audits establish navigation geometry. An initial new-case endnote
SSIM floor .57 rejected the measured .5407; its scaffold/reason are retained.
Existing public inputs and gates stay unchanged.

Frozen runtime **144b818b** passes clean Windows Release and one full catalogue:
2355 registered /2354 passed /0 failed /1 optional private-document skip, including
91 PDF, 130 SVG, 151 images, 184 typography and 989 methods containing Docx. Linux
passes 85 PDF +93 targeted methods, with six existing Arial skips and two archive
SourceLink warnings. Both inventories validate 360 cases across ten families.
Fresh identity transfers 34 current comparisons, 248 unchanged prior outputs and
two requalified identical multiline inputs through 284 conversions. All 2911
inherited DOCX and 251 broader PDFs retain bytes and page counts. The 34 cached
markup cases /508 failed gates transfer only through fresh exact unchanged PDF
identity. Local 0.1.5 package smoke passes. API/dependencies/version remain
unchanged; release preparation is deferred.


## RV06-L134: contextual paragraph contributions (2026-10-10)

Word collapses adjacent numeric spacing before removing the same-style contextual
contributions. The predecessor owns its after-spacing; the current paragraph owns
only max(0,before-pendingAfter). Suppression does not recompute that collapse.
Plain finite numeric body/table spacing and keep estimates share the resolver;
authored snapshots remain unchanged. Single-line note rectangles use the resolved
gap and omit current contextual after-spacing before a plain same-style successor.
Headers/footers and note-story text retain their previous branch: attempted
extensions worsened independent raster comparisons, despite improved relative
spacing. The rejected binaries/comparisons are retained. Table end-of-cell after
spacing and existing multiline contextual note hit areas remain partial.

32 primary +12 held-out +4 additional +2 public controls cover46 independent
geometries. Two primary controls repeat L132 inputs and public fixtures duplicate
primary geometry. Twelve PDFs and16 rasters retain bytes;34 rasters improve in
aggregate MAE/SSIM. All50 decoded-word multisets and page counts match Word.
Sixty-two rectangles change. In44 ordinary controls all176 links match Word
vertically within.005pt; table/header controls retain existing annotation residuals.
Five portable methods pass;three fail exact L133 andtwo guards pass both. Fourteen
generator ZIP payloads reproduce. The public footnote gates are MAE<=.50,
SSIM>=.80,recall>=.80,changed<=.0065; endnotes use.95/.60/.56/.010. Both require
matching pages/dimensions and only the corresponding approximation warning.
Existing gates and inputs are unchanged.

Frozen runtime **f4da5e20** passes clean Windows Release:2360 registered,2359 passed,
zero failed,one optional private-document skip (91 PDF,130 SVG,151 images,
184 typography,994 methods containing Docx). Linux passes85 PDF +98 targeted,
with six existing Arial skips andtwo archive SourceLink warnings. Both inventories
validate362 cases across ten families. Fresh334 visual conversions reproduce50
current comparisons and276 unchanged prior PDFs; eight older contextual controls
are independently requalified against pinned Word references and all improve.
They add14 changed rectangles;two multiline cases retain a6pt rectangle residual.
All2911 inherited DOCX and251 broader PDFs retain bytes and page counts. The34
cached markup cases/508 failed gates transfer only through fresh unchanged PDF
identities. Local0.1.5 package smoke passes. API/dependencies/version remain
unchanged; release preparation is deferred.


## RV06-L135: multiline note bounds after contextual predecessors (2026-10-10)

Resolve the predecessor's spacing contribution before filtering multiline note
paragraphs. A first-line mark owns a positive unconsumed before-gap after a plain,
nonempty contextual predecessor rendered on one line. Different styles retain
ordinary collapse ownership. Later-line marks and multiline after-spacing retain
line-only behavior. Empty/wrapped/rich/auto/line-unit predecessors, current-contextual
multiline paragraphs and existing layout exclusions retain fallback. Painting,
horizontal bounds and actual note destinations remain unchanged.

48 primary +12 held-out +2 public controls cover 56 independent geometries. Two
primary controls reuse exact historical inputs/references; two manual-page guards
also repeat prior inputs, and public cases duplicate supported geometry. Sixteen
independent and four repeated controls correct 20 rectangles. All 80 links in the
supported controls match Word vertically within 0.005pt. Forty-two guards retain
PDF bytes; all 62 retain rasters, decoded-word multisets and page counts. Four
portable methods pass; two fail exact L134 and two guards pass both. All 16 generator
ZIP payloads reproduce. Qualified prototype source/DLL/PDB remain archived.

Public multiline painting remains approximate: first-page MAE/SSIM are 2.254/.595424
for footnotes and 2.175/.552240 for endnotes. The new gates use MAE<=2.30,
SSIM>=.57/.53, recall>=.65/.56 and changed<=.024, matching pages/dimensions and only
the corresponding approximation warning. Separate annotation audits qualify bounds.
Existing inputs and gates are unchanged.

Frozen runtime **c5f412da** passes clean Windows Release: 2364 registered /2363
passed /0 failed /1 optional private-document skip, including 91 PDF, 130 SVG,
151 images, 184 typography and 998 methods containing Docx. Linux passes 85 PDF
+102 targeted methods with six existing Arial skips and two archive SourceLink
warnings. Both inventories validate 364 cases across ten families. Fresh visual
proof reproduces 62 current and 332 unchanged prior PDFs, plus two historical
identical inputs requalified by current Word controls, through 396 conversions.
Those two historical 6pt rectangle residuals are resolved. All 2911 inherited
DOCX and 251 broader PDFs retain bytes and page counts. The 34 cached markup cases
/508 failed gates transfer only through fresh unchanged PDF identity. Local 0.1.5
package smoke passes. API/dependencies/version stay unchanged; release deferred.


## RV06-L136: current-contextual multiline note bounds (2026-10-10)

First-line note marks now retain a positive resolved before-gap when the current
multiline paragraph uses contextual spacing. Section starts and different-style
predecessors can own such gaps. Matching styles keep suppression; later-line marks
remain line-only and multiline marks omit after-spacing. Plain successor, numeric
spacing, nonempty one-line predecessor and existing layout guards remain. Painting,
horizontal bounds and note destinations are unchanged.

64 controls cover 62 independent geometries. Twenty-four independent and two
public duplicate controls qualify 26 rectangles; all 104 links in those controls
match Word vertically within 0.005pt. Thirty-six guards retain PDF bytes. All 64
retain rasters, decoded-word multisets and page counts. Two independent omitted/
explicit default-style alias controls add two partial rectangle changes, retaining
their 6pt layout/bounds residual without worsening the maximum error. They add no
Office-bound admissions. Four portable methods pass; two fail exact L135 and two
guards pass both. All 18 generator ZIP payloads reproduce. Qualified source/DLL/PDB
remain archived. Public multiline painting and its new visual gates remain approximate.

Frozen runtime **6511b496** passes clean Windows Release: 2368 registered /2367
passed /0 failed /1 optional private-document skip, including 91 PDF, 130 SVG,
151 images, 184 typography and 1002 methods containing Docx. Linux passes 85 PDF
+106 targeted methods with six existing Arial skips and two archive SourceLink
warnings. Both inventories validate 366 cases across ten families. Fresh visual
proof reproduces 64 current and 396 unchanged prior PDFs through 460 conversions.
All 2911 inherited DOCX and 251 broader PDFs retain bytes and page counts. The 34
cached markup cases /508 failed gates transfer only through fresh unchanged PDF
identity. Local 0.1.5 package smoke passes. API/dependencies/version stay unchanged;
release preparation remains deferred.


## RV06-L137: declared default contextual style identity (2026-10-10)

Plain numeric body/table boundaries now match an omitted paragraph style with the
unambiguous declared default paragraph style for contextual spacing. The existing
note-bound rules use the same identity. Authored identifiers, property inheritance
and existing raw matches remain intact. Missing/off/multiple declarations,
malformed style elements, rich/auto/line-unit boundaries and static/related stories
retain fallback. The declared default's identifier can differ from Normal.

64 Office controls plus two public duplicates cover 62 independent geometries.
Thirty-eight independent and two public controls improve both aggregate MAE and
SSIM. Twenty-six controls retain PDF/raster identities. All 66 match Word's decoded
word multisets and page counts. Seventy changed ordinary-control rectangles are
qualified; all 152 links in 38 ordinary supported controls match Word vertically
within 0.005pt, including two supported PDF identities. Four table controls correct
their first boundary and improve visually, changing eight rectangles while
retaining at most 0.008pt vertical bounds error in the final qualified controls.
The earlier 24.007pt trailing-cell claim was corrected during L139. Two historical undeclared-default
aliases retain exact PDFs and their 6pt residual. Partial cases add no full
annotation admission. Horizontal bounds and actual destinations remain unchanged.

Four portable methods pass; two fail exact L136 and two preservation/guard methods
pass both. All 20 generator ZIP payloads reproduce. Public multiline painting
remains approximate under the new MAE<=2.30, SSIM>=.57/.53, recall>=.65/.56 and
changed<=.024 gates. First-boundary flow is checked absolutely; wrapped after-gap
contributions are checked against matching flags-off Word controls, separating
the retained cumulative line-height drift. Initial invalid table inputs and cloned
duplicate style names are retained excluded fixture attempts; accepted inputs
have valid section placement and distinct style names. Qualified source/DLL/PDB
remain archived.

Frozen runtime **4ba1ca56** passes clean Windows Release: 2372 registered /2371
passed /0 failed /1 optional private-document skip, including 91 PDF, 130 SVG,
151 images, 184 typography and 1006 methods containing Docx. Linux passes 85 PDF
+110 targeted methods, with six existing Arial skips and two archive SourceLink
warnings. Both inventories validate 368 cases across ten families. Fresh visual
proof reproduces 66 current and 460 unchanged prior PDFs through 526 conversions.
All 2911 inherited DOCX and 251 broader PDFs retain bytes and page counts. The 34
cached markup cases /508 failed gates transfer only through fresh unchanged PDF
identity. Local 0.1.5 package smoke passes. API/dependencies/version stay unchanged;
release preparation remains deferred.


## RV06-L138: Office implicit Normal contextual identity (2026-10-10)

For a loaded styles part without a declared paragraph default, plain numeric
contextual spacing now matches omitted styles with explicit Normal. This includes
absent paragraph-style definitions and Normal marked default-off. One declared
default retains precedence; ambiguity, missing styles parts and complex/static/
related boundaries keep fallback. Authored identifiers and property inheritance
remain intact. The existing body/table layout and note-bound rules share identity.

48 Office plus two public controls cover 42 independent geometries. Twenty-eight
independent and eight repeated/duplicate controls improve aggregate MAE and SSIM.
Fourteen guards retain PDFs/rasters. All 50 match Word words and pages. Seventy-two
changed rectangles are qualified; all 144 links in the 36 admitted controls match
Word vertically within 0.004pt. Horizontal bounds and actual destinations remain
unchanged. Four exact-slot controls and two named-default precedence guards
supplement the direction/flag/definition matrix. Six exact historical inputs and
references resolve four 12pt and two 6pt residual controls; these add no independent
admission. The four earlier final table controls retain at most 0.008pt vertical
bounds error; L139 corrects the stale 24.007pt assessment.

Four new portable methods pass; two fail exact L137 and two preservation/precedence
guards pass both. Earlier declared-default and contribution methods also pass.
All 22 generator ZIP payloads reproduce. Public multiline painting remains
approximate under the new MAE<=2.30, SSIM>=.57/.53, recall>=.65/.56 and changed<=.024
gates. Historical default-line-height and exact-slot source-layout residuals of
0.048/0.128pt remain; their numeric spacing corrections are checked against the
independent Word flag matrix. Wrapped after-gap contributions use matching
flags-off controls. Annotation and raster/text/page gates are separate. Qualified
prototype source/DLL/PDB are retained.

Frozen runtime **13b9a708** passes clean Windows Release: 2376 registered /2375
passed /0 failed /1 optional private-document skip, including 91 PDF, 130 SVG,
151 images, 184 typography and 1010 methods containing Docx. Linux passes 85 PDF
+114 targeted methods, with six existing Arial skips and two archive SourceLink
warnings. Both inventories validate 370 cases across ten families. Fresh visual
proof reproduces 50 current and 518 unchanged prior PDFs, plus eight prior entries
sharing six exact historical inputs requalified by current Word controls, through
576 conversions. All 2911 inherited DOCX and 251 broader PDFs retain bytes/pages.
The 34 cached markup cases /508 failed gates transfer only through fresh unchanged
PDF identity. Local 0.1.5 package smoke passes. API/dependencies/version stay
unchanged; release preparation remains deferred.

## RV06-L139: first-line table note before-spacing (2026-10-10)

Unsplit, top-aligned one-row/one-cell tables with two plain numeric paragraphs
now give a multiline first-line note marker its remaining collapsed before-space
after a nonempty, single-line predecessor. Automatic and exact slots and custom
first-character marks share the existing resolved contextual/style contributions.
Later-line markers remain line-only. Final-cell after-spacing stays in row height
and does not enlarge these multiline note links. Ordinary body paragraphs in table
documents retain their previous fallback. Rich/auto/line-unit/at-least spacing,
static/scaled/merged/nested/multiple-row and fragmented content remain partial.

96 independent Word controls plus two public geometry duplicates qualify
50 corrected rectangles. Changed rectangle heights match Word exactly; the
existing absolute vertical residual remains at most 0.042pt. This qualifies spacing
ownership, not general table typography. Forty-eight guards retain exact PDFs.
All 98 retain raster, inspected text and graphics identities against the parent;
Word word multisets and page counts match. Horizontal bounds and actual note
destinations remain unchanged. Unsupported guard residuals remain explicit.

The prior L137/L138 notes carried a stale 24.007pt trailing-cell claim. L137's
final frozen annotation proof and fresh L138 conversions instead establish a
maximum 0.008pt for those four corrected table controls. Their exact PDFs also
reproduce in this batch. Earlier artifacts are retained; the current assessment
is corrected. No row-height change was needed for that claim.

Four portable methods pass; two unchanged semantic checks fail against the exact
accepted L138 DLL and two preservation/fallback checks pass both. All 24 generator
ZIP payloads reproduce. The two new public cases remain approximate with the
existing MAE<=2.30, SSIM>=.57/.53, recall>=.65/.56 and changed<=.024 gates.

Frozen runtime **f7ae9380** passes clean Windows Release: 2380 registered /2379
passed /0 failed /1 optional private-document skip, including 91 PDF, 130 SVG,
151 images, 184 typography and 1014 methods containing Docx. Linux passes 85 PDF
+118 targeted methods, with six existing Arial skips and two archive SourceLink
warnings. Both inventories validate 372 cases across ten families. Fresh visual
proof reproduces 98 current and 576 unchanged prior PDFs through 674 conversions.
All 2911 inherited DOCX and 251 broader PDFs retain bytes and page counts. The
34 cached markup cases /508 failed gates transfer only through fresh unchanged
PDF identity. Local 0.1.5 package smoke passes. API/dependencies/version stay
unchanged; release preparation remains deferred.

## RV06-L140: minimum-height placement in bounded body tables

Runtime `7df5677d` corrects the excess minimum-height space above the first
baseline of a wrapped paragraph in an unsplit, top-aligned one-row/one-cell
body table. One or two plain numeric paragraphs with uniform font state and
full vertical metrics qualify; a predecessor must render as one line. Cursor
advances, row height, following flow, wrapping, glyph advances and graphics
remain unchanged. Related/static/nested, rich, merged, scaled, fragmented and
multiple-row paths retain fallback.

The actual slot/inset follows the corrected lines so first-line note links
retain minimum height and receive their remaining numeric before contribution.
After-spacing stays outside these multiline links. Automatic/exact slots keep
their previous PDF bytes.

Seventy-four independent Word controls, two historical repetitions and two
public duplicates cover three 12pt faces, five minima, first/following positions,
10pt/16pt extensions, custom marks and complex guards. All 78 retain Word word
multisets and page counts. Forty-six changed raster comparisons improve both
MAE and SSIM; 32 retain PNG identities. Twenty guards retain PDF bytes and
44 first-line rectangles gain the remaining gap. Qualified note vertical error
is at most 0.085pt, height error 0.002pt and first-baseline error 0.132pt.
Natural-height rounding can move a floor case by 0.002pt while retaining PNGs;
unsupported controls keep their inherited residuals. This is bounded admission.

Four semantic checks give two failures on the exact preceding DLL and two
preservation passes. Frozen validation: 2384 Windows methods, 2383 passes,
one optional private skip; Linux 85 PDF and 122 targeted passes with six existing
Arial skips. Fresh 2911 inherited and 251 broader outputs retain PDF/page
identities. Among 674 prior visual outputs, 672 retain bytes; the two original
minimum-height probes are freshly requalified against their exact Word
references. All 752 compared outputs reproduce in the integrated checkout.
The two public cases remain approximate. The local 0.1.5 package smoke passes;
release preparation remains deferred.

## RV06-L141: minimum-height placement in plain table grids

Runtime `a886860b` extends the established excess-above-baseline correction per
eligible body cell in multirow/multicell grids. All cell, row, font metrics,
fragmentation and related/static guards remain. Row heights, following flow,
wrapping, font state, glyph advances and graphics remain unchanged. Header,
declared-height and complex cells retain their existing model.

Fifty-four independent Word controls cover 1x2, 2x1 and 2x2 grids, first/following
paragraphs, Calibri/Arial minima18/24/36, Times New Roman24 and mixed row/column
face/size state. Two public duplicates and two historical note grids are counted
separately. All 58 match Word word multisets and page counts. Fifty changed
raster comparisons improve both MAE and SSIM; eight guards retain exact PDFs
and PNGs. The 136 improved cell baselines retain at most0.214pt absolute error,
including inherited row/border rounding. Historical note links and destinations
remain unchanged; their missing6pt before contribution remains partial.

Three semantic checks give one failure on the exact preceding DLL and two
preservation passes. The public generator reproduces its two new ZIP payloads;
the earlier26 reproducibility results transfer through unchanged generator,
helper and fixture Git blobs. New public thresholds cover both measured pages;
the inherited second-page text remains unchanged and approximate.

Frozen validation:2387 Windows methods,2386 passes,one optional private skip;
Linux85 PDF and125 targeted passes with six existing Arial skips. Fresh2911
inherited and251 broader outputs retain PDF/page identities. Of752 prior visual
outputs,750 retain bytes; two old note grids are freshly requalified against
their exact Word references. All810 qualified outputs reproduce in root. The
local0.1.5 package smoke passes; release preparation remains deferred.

## RV06-L142: minimum-height placement across table paragraphs

Runtime `891d05ae` extends the qualified excess-above-baseline rule to first,
middle, final and single-line minimum-height paragraphs in eligible plain body
table cells. Automatic/exact neighbors, row heights, following flow, wrapping,
glyph advances and graphics retain their prior geometry. Uniform font/full
metrics, top-alignment, numeric-spacing, unsplit/unscaled body guards remain.
Note before/after ownership retains its earlier bounded admission.

113 independent Word 16 controls plus three public duplicates cover fonts,
sizes, minimum heights, paragraph positions and mixed spacing transitions.
98 changed raster comparisons improve both MAE and SSIM; 18 retain images and
16 retain PDF bytes. 233 paragraph baselines improve, with maximum admitted
absolute error 0.173pt. Natural-height floors include two raster-identical
PDF-only rounding changes. All controls retain Word words/page counts and
parent graphics; neighboring automatic/exact text remains fixed. Eight complex
guards retain PDF and PNG identities. Three public cases remain approximate.

Four semantic checks fail twice on the exact accepted parent and preserve flow
and excluded paths there. Windows passes 2390 of 2391 registered methods with
one optional private skip; Linux passes 85 PDF plus 129 targeted methods with
six existing Arial skips and two archive SourceLink warnings. Fresh 2911
inherited DOCX and 251 broader PDFs retain bytes/pages. All 926 frozen visual
PDF identities transfer compared evidence, including all 810 prior controls.
The generator reproduces five current payloads and transfers 26 earlier note
payload results through unchanged source/helper/fixture Git blobs.

A temporary prototype audit with two reused output paths was superseded by a
fresh audit with unique namespaces and all final files rehashed. Its provenance
remains archived; accepted comparisons and outputs have independent paths.
Cached markup failures and broader complex table/note compositions remain
partial. Keep version 0.1.5; release preparation remains deferred.

## RV06-L143: minimum-height paragraphs in declared table rows

Runtime `f0ce24e4` applies the existing excess-above-baseline rule to eligible
plain body cells with positive exact/minimum row height, including an omitted
minimum rule. Explicit automatic row height ignores its authored value and
receives the same paragraph correction. Row geometry, following flow, ordinary
and exact neighbors, wrapping, glyph advances and graphics remain unchanged.
Headers, complex/scaled/fragmented cells and malformed height rules retain their
earlier admission. Note rectangle ownership remains bounded as before.

80 independent Word controls plus two public duplicates and three historical
repetitions produce 77 paired MAE/SSIM improvements. Six complex guards retain
PDF/PNG bytes, while two fully clipped variants retain images. 144 paragraph
baselines improve with maximum 0.146pt absolute error. All page counts match and
84 of 85 Word text multisets match; an unchanged nested guard retains a clipped
text omission. Source-backed diagnostics match 144 of 148 cell clips within 1.110pt;
four historical minimum-row clips and 175 unclassified reference clips remain
explicit. These supplemental diagnostics do not change existing graphics gates.

Windows passes 2393 of 2394 methods with one optional private skip. Linux passes
85 PDF plus 132 targeted methods, with six existing Arial skips and two archive
SourceLink warnings. Fresh 2911 inherited and 251 broader PDFs retain bytes/pages.
Among 926 prior visuals, 923 retain bytes and three declared-row controls reproduce
fresh Word comparisons. All 1011 qualified outputs retain their compared identities
in root. Seven generator payloads reproduce; 26 earlier results transfer through
unchanged source/helper/fixture Git blobs. Each inherited input has a separate
output path, including the two inputs that previously shared a case ID.

The separate cleanup commit archives 425 root-level agent scratch files with
verified hashes and ignores local validation snapshots. Keep 0.1.5; release
preparation remains deferred. Cached markup and broader composed-layout gaps
remain partial.

## RV06-L144: minimum table spacing across plain paragraph font changes

Runtime **7c000558** corrects minimum-height paragraph placement when font family,
size or bold/italic state changes between plain paragraphs in eligible body
table cells. Each paragraph places its own excess minimum-height slot above
its baselines. Rich runs within a paragraph, repeated headers, complex/scaled/
fragmented cells and unsupported stories retain fallback. The existing
automatic/exact baseline model, wrapping, glyph advances, font resources, row
height, following content and graphics stay unchanged.

The qualification covers 70 independent Word 16 controls, two public generator
cases, five repeated prior controls and three repeated complex guards. All 73
changed image comparisons improve both MAE and SSIM; seven guards retain exact
PDF/PNG bytes. All text multisets and page counts match Word. The 82 improved
first baselines have at most 0.271pt absolute error; mixed-size transitions
retain their existing automatic baseline residuals.

Eight note rectangles in six independent and two historical controls improve
their bounded first-line before-gap coverage to exact 30pt heights, with at
most 0.043pt vertical error. Horizontal bounds, annotation counts, destinations
and following paragraphs stay unchanged. Broader table note ownership and
after-spacing remain partial. Supplemental source mapping resolves all 83 cell
clips within 1.093pt; raw graphics gates are unchanged and absolute geometry
parity remains unproven.

The frozen runtime passes 2,397 registered Windows checks (2,396 pass, one
optional private skip) and 220 Linux checks (six existing Arial skips).
Windows Release builds have zero warnings/errors; Linux retains two archive
SourceLink warnings. All 2,911 inherited DOCX and 251 broader PDFs retain bytes
and page counts. Of 1,011 prior visual PDFs, 1,006 retain bytes and five are
freshly requalified; all 1,091 visual outputs reproduce their compared bytes.
The generator reproduces nine current payloads, with 26 earlier note payloads
transferred through unchanged source/helper/fixture Git blobs. The local
0.1.5 package smoke passes against the exact frozen runtime. Release
preparation is deferred.

## RV06-L145: minimum paragraph spacing in vertically aligned table cells

Runtime **89633277** corrects plain minimum-height paragraph slots in eligible
centered/bottom body table cells with positive exact/minimum row heights,
including an omitted minimum rule. The alignment model uses actual laid-out
content height and the terminal bottom-border advance. Ordinary/exact neighbors
move with the aligned content. Row height, following flow, wrapping, glyph
advances, font resources, horizontal bounds, graphics and destinations stay
unchanged. Automatic rows, positive vertical margins, rich runs, headers and
complex/scaled/fragmented cells retain their prior non-top fallback.

The qualification covers 70 independent Word 16 controls, two public generator
cases and four historical controls. All 69 admitted comparisons improve both
MAE and SSIM; seven guards retain exact PDF/PNG bytes. All 76 text multisets
and page counts match Word. The 264 improved in-cell paragraph baselines have
at most 0.212pt admitted first-baseline error. The rejected model for positive
vertical margins remains archived because both image metrics did not improve.

Supplemental source mapping resolves all 103 cell clips. The 96 admitted cells
retain up to 2.04pt absolute bounds error, and guarded margin cases retain up
to 6.99pt. Raw graphics gates are unchanged; mapping does not establish absolute
geometry parity. Broader table note ownership and after-spacing remain partial.

The frozen runtime passes 2,400 registered Windows checks (2,399 pass, one
optional private skip) and 223 Linux checks (six existing Arial skips).
Windows Release builds have zero warnings/errors; Linux retains two archive
SourceLink warnings. All 2,911 inherited DOCX and 251 broader PDFs retain bytes
and page counts. Of 1,091 prior visual PDFs, 1,089 retain bytes and two outputs
of one centered control are freshly requalified; all 1,167 qualified outputs
reproduce their compared bytes. The generator reproduces 11 current payloads,
with 26 earlier note payloads transferred through unchanged Git blobs.
The local 0.1.5 package smoke passes against the exact frozen runtime. Release
preparation is deferred.

## RV06-L146: declared-row margins and aligned minimum slots

Runtime **0c98e859** includes the maximum bottom margin in eligible positive
exact/minimum body-row advancement, outside the paragraph-slot extent used for
centered/bottom alignment. An omitted minimum rule is included. Plain cells
require minimum-height paragraphs, uniform font state within each paragraph and
full vertical metrics. Complete tables must fit the current body frame.
Measurement, keep estimates and recorded cell admission share the rule;
content-dominated minimum rows retain fallback even when final heights coincide.
Automatic/invalid rows, headers, rich cells and complex/scaled/fragmented paths
retain their separate fallback. Static/related/nested frame paths retain theirs.

40 independent Word 16 controls, two historical repeats and two public cases
qualify 36 paired MAE/SSIM improvements and eight exact PDF/PNG guards. All 44
text multisets and page counts match Word. The 147 improved in-cell first
baselines are within 0.187pt; 34 following-body baselines improve within 0.310pt.
Both public cases pass their unchanged numeric image gates. Horizontal text and
graphics geometry, glyph/font resources and actual destinations stay unchanged.
Admitted rows, vertical graphics and following flow include the bottom margin.

Supplemental mapping resolves eight of 59 source clips and leaves 51 unresolved:
49 admitted margin clips and two automatic-height guards. Clip parity remained
unproven at L146; L147 qualifies the admitted bounds below. The rejected
alignment-only model and its failing public image gates remain archived.

The frozen runtime passes 2,404 registered Windows checks (2,403 pass, one
optional private skip) and 227 Linux checks (six existing Arial skips). Windows
Release builds have zero warnings/errors; Linux retains two archive SourceLink
warnings. All 2,911 inherited DOCX and 251 broader PDFs retain bytes and pages.
Of 1,167 prior visual PDFs, 1,165 retain bytes and two historical margin controls
are freshly requalified; all 1,211 qualified outputs reproduce their compared
bytes. The generator reproduces 13 payloads, with 26 earlier note payloads
transferred through unchanged Git blobs. The exact local 0.1.5 package smoke
passes. Release preparation is deferred.

## RV06-L147: bounded clipping inside positive-margin cells

Runtime **50576e78** clips admitted body-cell content inside row-maximum
top/bottom margins, full top/terminal-bottom borders and half side borders.
Single and absent borders are included; compound borders, nonpositive content
rectangles and all unqualified L146 layouts retain their full-cell mask.
The same helper supplies effective bounds to layout snapshots. Supplemental
mapping consumes those bounds with unchanged 0.01pt/8pt search tolerances.

24 independent Word 16 controls, 44 trusted repeats and two new public overflow
cases qualify 18 paired MAE/SSIM improvements and 52 exact preview identities.
All 70 text multisets and page counts match Word. All text operations,
non-clip graphics operations and actual link rectangles/destinations stay
unchanged. 73 admitted clips map to Word within 0.190pt. Across 85 source cells,
81 map; four unresolved clips belong to automatic-height/compound-border
guards outside admission. Mapped fallback clips retain up to 6.99pt residual.
Both public cases pass the inherited numeric image gates with empty diagnostics.

Windows passes 2,407 registered checks (2,406 pass, one optional private skip)
with zero Release warnings/errors. Linux passes 230 checks, retaining six
existing Arial skips and two source-archive SourceLink warnings. Exact-parent
regressions fail twice and retain the fallback pass; the prototype passes
35 related checks. 2,911 inherited DOCX and 251 broader PDFs retain bytes/pages.
Of 1,211 earlier visual PDFs, 1,175 retain bytes and 36 changed masks are freshly
requalified through exact preview/text/non-clip-graphics identities. The frozen
runtime reproduces all 4,443 qualified outputs, including 1,281 visual PDFs.
The local dependency-free 0.1.5 package smoke passes; release preparation remains
deferred. The generator reproduces 15 payloads and transfers 26 earlier note
payloads through unchanged Git blobs. Malformed preliminary generator trials
are retained locally and excluded from qualification.

## RV06-L148: minimum slots in bounded automatic-height cells

Runtime **39758d08** corrects paragraph minimum-height excess in plain centered/bottom one-row/one-cell body tables that fit their current frame. Absent heights and explicit automatic heights are included; authored automatic values are ignored. Natural-height floors retain exact previous emission. Row geometry, ordinary/exact neighbors, following flow, graphics operations and all link rectangles/destinations remain unchanged. Broader grids, rich/static/related/scaled/fragmented and frame-boundary paths retain fallback.

50 independent Word16 controls, seven historical repeats, 12 trusted guards and two public cases qualify 55 paired MAE/SSIM improvements and 16 exact PDF/preview identities. All 71 text multisets/pages match Word. 57 first baselines improve; non-compound residual is at most 0.267pt. Two compound-border controls improve but retain 2.043pt error. Supplemental mapping resolves 35/83 source clips and leaves 48 automatic/compound clips unresolved; clip parity is not admitted.

The initial floor model fails the natural-floor regression; the guarded prototype passes 39 related checks. Exact-parent tests fail once and retain three passes. The Calibri public trial improves relative metrics but misses the unchanged 0.84 SSIM gate at 0.8377. That failure is archived; both Arial companion cases pass inherited numeric/diagnostic gates.

Windows 2411 registered/2410 passes/one optional private skip, with zero Release warnings/errors. Linux 234 passes with six existing Arial skips and two source-archive SourceLink warnings. 2911 inherited DOCX and 251 broader PDFs retain bytes/pages. Of 1281 earlier visual outputs, 1272 retain bytes and nine outputs of seven inputs receive fresh independent comparisons. All 4514 qualified corpus outputs, including 1352 visual PDFs, reproduce frozen bytes. The generator reproduces 17 payloads and transfers 26 earlier note payloads through unchanged Git blobs. Local 0.1.5 package smoke passes; release preparation remains deferred.

An existing spill-file directory-snapshot assertion failed in the initial full catalogue. The same frozen binaries pass its isolated check and the complete rerun, with no spill files remaining. The cause is unestablished; both reports are preserved.

## RV06-L149: content clipping in bounded automatic-height cells

Runtime **9372c29c** clips plain centered/bottom one-row/one-cell automatic-height body tables inside vertical margins and single/absent border insets. It reuses L148's complete-current-frame admission. Absent heights and explicit automatic heights are included; authored automatic values are ignored. The same helper supplies PDF masks and effective snapshot bounds. Text operations, row/following-flow geometry, non-clip graphics, actual links and page counts remain unchanged. Compound borders, nonpositive content rectangles and broader grids/complex/scaled/fragmented/frame-boundary paths retain fallback.

24 independent Word controls, 71 trusted repeats and two new public overflow cases qualify 26 paired MAE/SSIM improvements and 71 exact preview identities. All 97 text multisets and page counts match Word. 83 masks change. Supplemental mapping resolves 105 of 109 source cells; all 99 admitted clips map within 0.22pt using unchanged 0.01pt candidate and 8pt reference search tolerances. Four compound-border guards remain unresolved. Mapped fallback clips retain up to 6.99pt error, and the inherited compound-border first-baseline residual remains 2.043pt.

The initial public mask-only trials retain exact previews and are preserved with input snapshots and Office references. Final public cases expose exact-line descender overflow and pass inherited image gates with empty diagnostics. Three new regressions fail twice on the exact parent and retain the fallback pass; all 42 related prototype checks pass.

Windows has 2414 registered methods, 2413 passes and one optional private skip, with zero Release warnings/errors. Linux passes 237 checks with six existing Arial skips and two source-archive SourceLink warnings. All 2911 inherited DOCX and 251 broader PDFs retain bytes/pages. Of 1352 earlier visual outputs, 1286 retain bytes; 66 outputs of 57 inputs receive fresh comparisons proving exact preview/text/non-clip-graphics/link identities. All 4611 qualified corpus outputs, including 1449 visual PDFs, reproduce frozen bytes. The generator reproduces 19 payloads and transfers 26 earlier note payloads through unchanged Git blobs. Local 0.1.5 package smoke passes; release preparation remains deferred.

## RV06-L150: bounded uniform double-border bands

Runtime **54559e42** matches Word's double-border band in unscaled one-column body tables with uniform, equal, positive authored widths. Two authored-width strokes and an equal gap supply layout, painting and existing content-mask admissions. Mixed/unequal widths, missing/invalid/nonpositive sizes, revised/nested/merged/legacy/cell-spaced and multicolumn/scaled paths retain prior behavior.

50 independent Word controls, four trusted compound repeats and two new public overflow cases qualify 56 paired MAE/SSIM improvements with matching text contents/font state and unchanged actual navigation. The four earlier unresolved compound masks map within 0.144pt; their first-baseline residual improves from 2.043pt to 0.123pt. Across 57 controls, supplemental mapping resolves 36 of 76 source cells; all 34 admitted masks map within 0.41pt using unchanged 0.01pt candidate and 8pt reference search tolerances. Broader top/grid/header/rich masks remain partial, with 40 unresolved cells and matched fallback error up to 7.59pt. Broader grid/header/rich first baselines retain up to 10.074pt error and following flow up to 12.89pt; measured positions do not worsen.

A global width prototype was rejected because the public mixed-border review fixture's MAE worsened and emitted characters changed. The bounded final model restores exact parent PDF and preview bytes for that fixture, including nine inherited broader outputs. This guard remains partial against Word (783 versus 622 nonspace characters); its identity does not establish Office parity. An intermediate bounded model also failed the missing/invalid/nonpositive-width guard; the final admission retains those paths. Raw rejected models and comparisons remain archived. Four regressions fail twice on the exact parent, retain two fallback passes, and pass in the final 46-check focused qualification.

Both new public cases pass unchanged inherited numeric gates with empty diagnostics. The generator reproduces 21 current payloads and transfers 26 earlier note payloads through unchanged Git blobs. Of 1449 earlier visual outputs, 1439 retain bytes; ten outputs of four inputs receive fresh independent comparisons. All 2911 inherited DOCX and 251 broader outputs retain bytes/pages. All 4668 frozen corpus outputs, including 1506 visual PDFs, reproduce qualified bytes/pages.

Windows has 2418 registered methods, 2417 passes and one optional private skip, with zero Release warnings/errors. Linux passes 241 checks with six existing Arial skips and two archive SourceLink warnings. Local 0.1.5 package smoke passes; release preparation remains deferred. Earlier cached markup failures and unrelated geometry gaps remain explicit.

## RV06-L151: automatic-height grid slots, alignment and content masks

Runtime **a2f2d207** extends minimum-slot placement to plain centered/bottom automatic-height body grids that fit their complete current single-column frame. Absent heights and explicit automatic heights are included. Excess minimum height sits above the first baseline. Shorter cells align their complete slot beside taller cells using the row maximum bottom margin, including natural-height floors. Eligible PDF masks and snapshot content bounds share that maximum margin. Row heights, borders, following body flow and actual navigation retain prior geometry. Mixed automatic/declared heights, top-aligned masks and complex/static/related/scaled/fragmented/frame-boundary paths retain fallback.

56 independent Word controls, five trusted repeats and two new public cases qualify 59 paired MAE/SSIM improvements and four exact PDF/preview identities. All 63 text multisets and page counts match Word; contents/font state and actual navigation stay unchanged. Qualified first baselines map within 0.723pt, including double-width grids; following body flow retains up to 1.25pt. The mixed-height automatic fallback retains a 9.384pt baseline residual. Supplemental source mapping resolves 192 of 204 cells, including all 192 admitted masks within 0.77pt. Eight top-aligned cells and four mixed-height automatic cells remain unresolved. The mapper canonicalizes signed-zero translation buckets, removing false candidate ambiguity without changing its 0.01pt candidate or 8pt reference tolerances or raw visual gates.

Six regressions fail three times on the exact parent and retain three compatibility passes; all 52 focused checks pass. The first frozen grid build changed four inherited single-cell exact/minimum overflow PDFs and was rejected. A regression with fractional font metrics and wide borders reproduces the failure; restricting the new automatic alignment calculation to grids restores all four accepted PDFs exactly. The rejected freeze and outputs remain archived. Intermediate placement and own-margin models fail the new unequal-cell and row-maximum semantic guards, respectively; those binaries and comparisons remain archived. The initial dense four-cell public trial improves both metrics but misses the unchanged 2.0 MAE gate at 2.146. Its input snapshots, generator, Office references and raw comparisons remain preserved. Reduced fixtures retain the minimum/exact descender-overflow defect and pass the same inherited gates with empty diagnostics.

Of 1506 earlier visual PDFs, 1501 retain bytes; five inputs receive fresh independent comparisons, including a former L141 centered-grid fallback. All 2911 inherited DOCX and 251 broader outputs retain bytes/pages. All 4731 frozen corpus outputs, including 1569 visual PDFs, reproduce final qualified bytes/pages. The generator reproduces 23 payloads and transfers 26 earlier note payloads through unchanged Git blobs.

Windows has 2424 registered methods, 2423 passes and one optional private skip, with zero Release warnings/errors. Linux passes 247 checks with six existing Arial skips and two archive SourceLink warnings. Local 0.1.5 package smoke passes; release preparation remains deferred. Earlier cached markup failures, review-text omissions and unrelated geometry gaps remain explicit.

## RV06-L152: automatic rows in mixed-height body grids

Runtime **623582ae** admits plain centered/bottom two-by-two body grids that combine automatic/absent heights with positive exact/atLeast heights and fit their complete current single-column frame. Automatic-row minimum slots and content masks use their row's height rule; declared rows retain existing geometry and emission, including natural-height floors. Row height, borders, following flow and actual navigation stay unchanged. Other mixed shapes, default/invalid declared rules and complex/static/related/scaled/fragmented/frame-boundary paths retain fallback.

42 independent Word controls, two trusted repeats and two public cases qualify 42 paired MAE/SSIM improvements and four exact rich/header PDF/preview guards. All 46 text multisets/page counts match Word; contents/font state and actual navigation retain identity. Outside the guards, measured baseline/following positions improve from up to 19.85pt to at most 0.53pt. Guard residuals remain explicit: rich automatic slots retain 9.384pt; headers retain up to 32.78pt. Source mapping resolves 176 of 184 cells, including all 172 admitted masks within 0.29pt. Eight guard masks remain unresolved, and four matched fallback cells retain 6.99pt error. Mapping tolerances and raw gates are unchanged.

Four regressions fail twice on the exact parent and retain two compatibility passes; all 56 focused checks pass. Both public cases pass inherited thresholds with empty diagnostics. The generator reproduces 25 payloads and transfers 26 earlier note payloads through unchanged Git blobs.

Of 1569 earlier visual PDFs, 1567 retain bytes; two previously partial mixed-height inputs receive fresh independent comparisons. All 2911 inherited DOCX and 251 broader PDFs retain bytes/pages. All 4777 frozen corpus outputs, including 1615 visual PDFs, reproduce final qualified bytes/pages. Windows passes 2427 of 2428 registered methods, with one optional private skip and zero Release warnings/errors. Linux passes 251 checks, with six existing Arial skips and two archive SourceLink warnings. Local 0.1.5 package smoke passes; release preparation remains deferred. L151's rejected single-cell freeze, dense public trial and older cached markup failures and review-text omissions remain preserved.

## RV06-L153: top-aligned automatic table content masks

Runtime **2f089522** admits inner content masks for plain top-aligned automatic/absent-height body tables with one or two rows and columns. Each cell contains minimum-height spacing with full vertical metrics; the complete table must fit its current single-column frame. The mask uses vertical margins, qualified single/double border insets and the terminal bottom-border advance. It preserves exact text operations and existing slot placement. Declared/mixed-height top rows, larger shapes and complex/static/related/scaled/fragmented/frame-boundary paths retain fallback.

46 independent Word controls, two trusted repeats and two public cases qualify three paired MAE/SSIM overflow improvements, 47 exact preview identities and four exact rich/header/alignment/mixed-height PDF guards. All 50 text multisets and page counts match Word; every parent/candidate text-operation file retains exact bytes, and actual navigation stays unchanged. Source mapping resolves 114 of 128 cells, including all 112 admitted masks within 0.81pt. Fourteen guard masks remain unresolved; two matched fallback cells retain 6.99pt error. Raw gates and mapping tolerances are unchanged.

Three focused regressions retain two compatibility passes and fail once on the exact parent; all 59 focused checks pass. Both public cases pass inherited thresholds with empty diagnostics. The generator reproduces 27 payloads and transfers 26 earlier note payloads through unchanged Git blobs.

Of 1615 earlier visual PDFs, 1343 retain bytes. The remaining 272 outputs map to 266 verified input/view identities and receive fresh Office raster, text and graphics inspection with source-backed cell mapping. All 266 preserve exact parent/candidate text operations and match Word text/page counts; 266 retain preview bytes and 0 improve both raster metrics. Their 425 admitted masks all resolve against Word, within 0.770pt. Evidence transfer does not rely on preview identity alone.

All 2911 inherited DOCX and 251 broader PDFs retain bytes/pages. All 4827 frozen corpus outputs, including 1665 visual PDFs, reproduce final qualified bytes/pages. Windows passes 2430 of 2431 registered methods, with one optional private skip and zero Release warnings/errors. Linux passes 254 checks, with six existing Arial skips and two archive SourceLink warnings. Local 0.1.5 package smoke passes; release preparation remains deferred. Prior rejected freezes, public trials, cached markup failures and review-text omissions remain preserved.

## RV06-L154: declared top-row margin advance and content masks

Runtime **b1c9750a** includes the row maximum bottom margin in the advance of qualified top-aligned rows, outside their content slot. Independent Word measurements showed the missing advance displaced following rows and body flow by 12–30pt. Inner masks alone could not correct that geometry. Admission covers plain one/two-row and column body tables with positive explicit exact/atLeast heights, minimum-height paragraphs, full metrics and qualified single/double border bands. The complete table must fit its current single-column frame. Exact/atLeast mixtures qualify; each content-dominated minimum row retains fallback. Mixed automatic/default rules, unsupported compound borders, larger shapes and complex/static/related/scaled/fragmented/frame-boundary paths retain fallback.

54 independent Word controls, two trusted repeats and two public cases qualify 46 paired MAE/SSIM improvements, 12 exact preview identities and seven exact excluded-layout PDF guards. All 58 text multisets and page counts match Word. Text contents, font state, horizontal placement and actual navigation stay unchanged; qualified following-row/body Y positions move with the corrected advance. Twelve controls retain exact parent/candidate text-operation bytes. Measured admitted baseline/body-position errors fall from a maximum 30.19pt to 1.27pt; guards retain up to 24.31pt. Source mapping resolves 136 of 152 cells, including all 122 admitted masks within 0.79pt. Sixteen guard masks remain unresolved; fourteen matched fallback cells retain up to 7.11pt error. Raw gates and mapping tolerances are unchanged. A representative two-row border comparison also confirms the corrected row advance and terminal paint geometry, within 0.07pt vertically.

Three focused regressions retain two compatibility passes and fail once on the exact parent; all 62 focused checks pass. Both public cases pass inherited thresholds with empty diagnostics. The generator reproduces 29 payloads and transfers 26 earlier note payloads through unchanged Git blobs. The initial mask-only analytical model missed later-row geometry by 12.79pt. A subsequent compound-border regression failed on the prototype; the final admission guard preserves its fallback. Original sources, binaries and failed reports remain archived.

Of 1665 earlier visual PDFs, 1559 retain bytes. The remaining 106 outputs map to 97 verified input/view identities and receive fresh Office raster, text and graphics inspection with source-backed mapping. All 97 match Word text/page counts and retain contents, font state and horizontal placement; 90 retain exact text-operation and preview bytes, and seven improve both raster metrics. All 137 admitted prior masks resolve within 0.41pt. Eight changed outputs arise from two earlier top-row fallback inputs which are now explicitly admitted and independently compared. Nine repeated outputs map to verified identical input/view and parent/candidate hashes. Changed PDF evidence does not transfer through preview identity alone.

All 2911 inherited DOCX and 251 broader PDFs retain bytes/pages. All 4885 frozen corpus outputs, including 1723 visual PDFs, reproduce final qualified bytes/pages. Windows passes 2433 of 2434 registered methods, with one optional private skip and zero Release warnings/errors. Linux passes 257 checks, with six existing Arial skips and two archive SourceLink warnings. Local 0.1.5 package smoke passes; release preparation remains deferred. Prior rejected freezes, public trials, cached markup failures and review-text omissions remain preserved.

## RV06-L155: mixed automatic/declared top table rows

Runtime **94c2cf36** admits plain top-aligned two-by-two body grids combining automatic/absent and positive explicit exact/atLeast rows. Independent Word measurements qualify both orders, single/double border bands, margin variants and minimum-spacing floors. Each row reuses its accepted uniform-table geometry: automatic rows retain natural advance, while declared rows include the maximum bottom margin outside the content slot. Inner masks use qualified border/margin insets. Full metrics and minimum-height paragraphs are required, and the complete table must fit its current single-column frame. Content-dominated minimum rows retain per-row fallback. Other mixed shapes, default/invalid rules, unsupported compound borders and complex/static/related/scaled/fragmented/frame-boundary paths retain fallback.

57 independent controls, two trusted earlier guards and two public cases qualify 46 paired MAE/SSIM improvements, 15 exact preview identities and nine exact excluded-layout PDF guards. All 61 text multisets/page counts match Word. Contents, font state, horizontal placement and navigation remain unchanged; qualified following-row/body Y positions move with the declared margin advance. Fifteen controls retain exact text-operation bytes. Measured admitted position errors fall from 18.29pt to 1.27pt; guards retain up to 20.21pt. Source mapping resolves 220 of 246 cells, including all 204 admitted masks within 0.77pt. Twenty-six fallback masks remain unresolved; sixteen matched fallback cells retain up to 7.13pt error. Raw gates and mapping tolerances are unchanged. Two new regressions fail on the exact parent while one compatibility method passes; all 65 focused checks pass. Both public cases pass inherited thresholds with empty diagnostics. The generator reproduces 31 payloads and transfers 26 earlier note payloads through unchanged Git blobs.

Of 1723 earlier visual PDFs, 1721 retain bytes. The two changed outputs are earlier mixed-top fallback inputs, now explicitly admitted and freshly compared against Word. Both improve raster metrics and preserve Word text/page counts; all eight source masks resolve within 0.15pt. Fresh comparisons include text and graphics inspection rather than relying on preview similarity. A source-inspection helper initially omitted the markup-view arguments; original records and failed output are archived, and the corrected metadata receives a successful rerun.

All 2911 inherited DOCX and 251 broader PDFs retain bytes/pages. All 4946 frozen outputs, including 1784 visual PDFs, reproduce final qualified bytes/pages. Windows passes 2436 of 2437 registered methods, with one optional private skip and zero Release warnings/errors. Linux passes 260 checks, with six existing Arial skips and two archive SourceLink warnings. Local 0.1.5 package smoke passes; release preparation remains deferred. Earlier rejected models/freezes, public trials, cached markup failures and review-text omissions remain preserved.

## RV06-L156: two-row/one-column mixed top tables

Runtime **091e03a3** admits plain two-row/one-column top-aligned body tables combining automatic/absent and positive explicit exact/atLeast rows. Independent Word measurements qualify both orders, single/double border bands, margins and minimum-spacing floors. Each row reuses its accepted uniform-table geometry: automatic rows retain natural advance, while declared rows include the maximum bottom margin outside their slot. Inner masks use qualified border/margin insets. Full metrics and minimum-height paragraphs are required, and the complete table must fit its current single-column frame. Content-dominated minimum rows retain per-row fallback. Existing two-by-two admission stays unchanged. Larger mixed shapes, default/invalid rules, unsupported compound borders and complex/static/related/scaled/fragmented/frame-boundary paths retain fallback.

57 independent controls, one trusted earlier guard and two public cases qualify 45 paired MAE/SSIM improvements, 15 exact preview identities and nine exact excluded-layout PDF guards. All 60 text multisets/page counts match Word. Contents, font state, horizontal placement and navigation remain unchanged; qualified following-row/body Y positions move with the declared margin advance. Fifteen controls retain exact text-operation bytes. Measured admitted position errors fall from 18.29pt to 1.25pt; guards retain up to 20.21pt. Source mapping resolves 110 of 127 cells, including all 100 admitted masks within 0.79pt. Seventeen fallback masks remain unresolved; ten matched fallback cells retain up to 7.13pt error. Raw gates and mapping tolerances are unchanged. Two new regressions fail on the exact parent while one compatibility method passes; all 68 focused checks pass. Both public cases pass inherited numerical image thresholds and emit empty diagnostics. The generator reproduces 33 payloads and transfers 26 earlier note payloads through unchanged Git blobs.

Of 1784 earlier visual PDFs, 1783 retain bytes. The changed output is the earlier one-column mixed-top fallback input, now explicitly admitted and freshly compared against Word. It improves both raster metrics and matches Word text/page counts; its two source masks resolve within 0.15pt. Fresh comparison includes text and graphics inspection rather than relying on preview similarity.

All 2911 inherited DOCX and 251 broader PDFs retain bytes/pages. All 5006 frozen outputs, including 1844 visual PDFs, reproduce final qualified bytes/pages. Windows passes 2439 of 2440 registered methods, with one optional private skip and zero Release warnings/errors. Linux passes 263 checks, with six existing Arial skips and two archive SourceLink warnings. Local 0.1.5 package smoke passes; release preparation remains deferred. Earlier rejected models/freezes, validation-helper failures, public trials, cached markup failures and review-text omissions remain preserved.

## RV06-L157: bounded mixed three-row top tables

Runtime **280bd0e6** admits plain mixed three-row top-aligned body tables with one/two columns, automatic/absent and positive explicit exact/atLeast rows. Full metrics, minimum-height paragraphs and a complete fit in the current single-column frame are required. Single bands up to 0.96pt, double bands up to 2.88pt and absent borders are qualified. Each row retains its accepted natural/declared advance and inner mask. A regression exposed repeated whole-table fit checks against space remaining after earlier rows; the corrected implementation retains one frame only for admitted complete three-row tables. Existing one/two-row behavior stays unchanged. Uniform three-row tables, wider three-row borders, larger mixed shapes, default/invalid rules and complex/frame-boundary paths retain fallback. Content-dominated minimum rows retain per-row fallback.

115 independent controls, two trusted earlier guards and two focused public cases qualify 80 paired MAE/SSIM improvements, 39 exact preview/text-operation identities and 37 exact excluded-layout PDF guards. All 119 text multisets/page counts match Word. Contents, font state, horizontal placement and navigation stay unchanged; qualified later-row/body Y positions include the declared margin advances. Source mapping resolves 385 of 559 cells, including all 366 admitted masks within 0.43pt. The 174 unresolved fallback masks and 19 matched fallback cells, with up to 6.99pt error, remain explicit. Raw gates and mapping tolerances are unchanged. All 71 focused checks pass; two regressions fail on the exact parent and one compatibility method passes. Two single-page public cases pass unchanged numerical image thresholds and empty diagnostics. The generator reproduces 35 payloads and transfers 26 earlier note payloads through unchanged Git blobs.

The original broad model reaches 1.15pt mask and 1.63pt text errors for wide double borders, which remain excluded. The rejected initial frame prototype and its failed witness remain archived. An initial four-row guard retains exact parent/candidate PDF bytes while exposing an inherited two-versus-three-page Word mismatch; a smaller four-row fallback guard receives a fresh comparison. Initial public table pages pass, while unchanged second pages miss SSIM/foreground recall. Those trials remain archived. Focused public cases remove the later text and its section break; the intermediate empty-section trial remains archived. Helper count/array corrections, a stale prior-summary attempt and an incorrect concurrent output prefix are preserved; final evidence uses completed consumers, matching job/result output paths and separate source-mapping directories.

Of 1844 earlier visual PDFs, 1842 retain bytes. Both changed PDFs are earlier mixed three-row fallback inputs, freshly compared against Word with paired improvements, matching text/page counts and all nine source masks within 0.19pt. All 2911 inherited DOCX and 251 broader PDFs retain bytes/pages. All 5125 frozen outputs, including 1963 visual PDFs, reproduce qualified bytes/pages. Windows passes 2442 of 2443 registered methods, with one optional private skip and zero Release warnings/errors. Linux passes 266 checks, with six existing Arial skips and two archive SourceLink warnings. Local 0.1.5 package smoke passes; release preparation remains deferred. Earlier rejected models/freezes, cached markup failures and review-text omissions remain explicit.

## RV06-L158: bounded uniform three-row top tables

Runtime **02caf773** admits plain uniform three-row top-aligned body tables with one/two columns and automatic/absent or positive explicit exact/atLeast rows. Full metrics, minimum-height paragraphs and a complete fit in the current single-column frame are required. Single bands up to 0.96pt, double bands up to 2.88pt and absent borders remain qualified. Existing natural/declared slots and masks are reused; declared rows include the maximum bottom margin outside their slots. Stable frame admission now also covers declared-only three-row tables. Accepted one/two-row and mixed three-row behavior retains fresh corpus identity. Content-dominated minimum rows retain per-row fallback. Wider three-row borders, larger shapes, default/invalid rules and complex/frame-boundary paths retain fallback.

86 independent controls and two focused public cases qualify 32 paired MAE/SSIM improvements, 56 raster identities, 57 exact text-operation identities and 26 exact excluded-layout PDF guards. All 88 text multisets/page counts match Word. Text contents, font state, horizontal placement and navigation remain unchanged. Source mapping resolves 295 of 437 cells, including all 276 admitted masks within 0.70pt. The 142 unresolved fallback masks and 19 matched fallback cells, with up to 6.99pt error, remain explicit. Raw gates and mapping tolerances are unchanged. All 74 focused checks pass; two regressions fail on the exact parent and one compatibility method passes. Both single-section public cases pass existing numerical image thresholds and empty diagnostics. The generator reproduces 37 payloads and transfers 26 earlier note payloads through unchanged Git blobs.

The initial independent broad hypothesis reaches 1.18pt mask and 1.656pt text error on wider double borders. Narrow controls measure 126 masks within 0.46pt and predicted text positions within 0.70pt; wider bands retain fallback. One initial boundary fixture incorrectly excluded a natural automatic table that still fitted its 160pt page. Its original source and failed report remain archived; a 100pt page supplies a genuine boundary test without changing runtime. Earlier three-row fallback assertions were superseded by four-row guards and retained with their reports. The inherited L157 four-row pagination and second-page typography gaps remain explicit.

Of 1963 earlier visual PDFs, 1960 retain bytes. The three changed outputs are earlier uniform three-row fallback inputs with fresh Word raster/text/graphics comparisons and source-backed mapping. All match text/page counts; one retains preview/text-operation identity and two improve both raster metrics. All 18 admitted prior masks resolve within 0.22pt. All 2911 inherited DOCX and 251 broader PDFs retain bytes/pages. All 5213 frozen outputs, including 2051 visual PDFs, reproduce qualified bytes/pages. Windows passes 2445 of 2446 registered methods, with one optional private skip and zero Release warnings/errors. Linux passes 269 checks, with six existing Arial skips and two archive SourceLink warnings. Both inventories validate 411 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Earlier rejected models/freezes, cached markup failures and review-text omissions remain explicit.

## RV06-L159: bounded each-page footnote numbering

Runtime **381a39af** numbers automatic footnotes from their actual physical pages in plain single-section, single-column ASCII body flow. Explicit section `eachPage`, decimal/decimalZero formats, positive starts and measured section defaults qualify. Markers must own their source runs, use a single text line and share a page with a complete one-line note. Corrected labels receive fresh font subsets and measured advances; one retry is accepted only with stable marker/note pages. Authored digits, source offsets, settings and reader inspection remain intact. Typed page-break spacing rows remain in PDF output but do not count as wrapped marker lines. Document-level defaults, custom marks, revisions, tables, headers/footers, columns, wrapped content, unavailable full fonts and unstable-page paths retain fallback.

42 independently authored controls and two reproducible public counterparts measure 208 automatic body/note labels and physical pages matching Word. All 104 links contain their rendered markers and target distinct matching notes. Destination positions differ from Word by at most 0.998pt and horizontal link bounds by at most 1.398pt. Parent link vertical bounds and destinations remain identical. All 52 admitted raster pages improve MAE, SSIM and foreground recall; all 18 excluded PDFs retain parent bytes. Authored text, font sizes and actual pages retain parent identity. Natural controls preserve inherited 64pt page-two body drift; after-spacing controls preserve up to 12.003pt body drift. Those numbering and navigation admissions do not establish complete visual fidelity. The two explicit-break public cases pass existing note-section numeric thresholds and the expected `DOCX_APPROXIMATED_FOOTNOTE` diagnostic. Their generator reproduces all 14 package-entry payloads; agent-rating metadata is not an independent human rating.

The initial prototype corrected paragraph-break and natural numbering but retained fallback for run-break controls because real spacing rows shared paragraph provenance. Its source, failed-to-admit measurements and typed witness remain archived. Validation helpers initially read diagnostic `Code` rather than `Id` and mixed DOCX group/name counting predicates. Their failures and corrections remain archived without a runtime or threshold change. Five focused semantic regressions cover authored-digit preservation, spacing versus wrapping, moved/partial/duplicate notes, excluded settings and emitted labels/navigation together.

All 5213 accepted-parent PDFs retain bytes. All 5257 frozen outputs, including 2095 visual PDFs, reproduce qualified bytes/pages. Windows passes 2450 of 2451 registered methods, with one optional private skip and zero Release warnings/errors. Linux passes 274 checks, with six existing missing-Arial skips and two archive SourceLink warnings. Both inventories validate 413 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Earlier rejected models/freezes, cached markup failures, stored REF policies, review-text omissions and unrelated gaps remain explicit.

## RV06-L160: shared complete-note reservation

Runtime **480fea20** reclaims unused reserve for complete one-line footnotes in qualified natural pagination. The earlier reserve repeated separator spacing and a continuation allowance for each note. The bounded model sums complete note line boxes and one shared separator line while retaining existing note placement. It requires requested Office-compatible output, unit print scale, full font measurement, plain single-section/single-column ASCII flow, an authored blank separator and explicit each-page or continuous section numbering. Every marker/note owner and the page count must retain the original layout. Missing/nonblank separators, wrapped/continued notes, revisions, tables, headers/footers, columns, keep chains, unavailable full fonts and changed owners retain fallback.

Eighteen independent Word controls comprise nine admissions and nine parent-PDF exclusions. All nine admissions match Word spacer capacity: the retained 31-versus-35-line gap becomes 35/1, and the three-initial-note variant changes from 27/9 to Word's 34/2. Body baselines differ by at most 0.22pt. Three unchanged-input references from L159 independently requalify the earlier natural cases, whose body baselines now differ by at most 0.14pt. Across those controls and two public counterparts, 112 generated labels/pages and 56 actual links resolve to matching Word notes. Link bounds remain approximate, with maximum 1.398pt displacement; destinations differ by at most 0.945pt. Authored text and font sizes are preserved even when spacer paragraphs change pages.

All eighteen admitted independent raster pages improve SSIM and foreground recall; nine improve MAE. Visible spacer controls retain low absolute pixel similarity. The two-note public case passes the existing note-section thresholds (MAE <=.42, SSIM >=.8, recall >=.8, changed-pixel ratio <=.0055). The three-note counterpart retains three page-one failures: MAE .6167, SSIM .7825 and changed-pixel ratio .00762. Its page-two metrics improve and its capacity, labels, owners and navigation qualify, but full raster fidelity is not established. These partial results remain public with unchanged thresholds and the expected `DOCX_APPROXIMATED_FOOTNOTE` diagnostic. Agent-rating metadata is not an independent human review. The generator reproduces all fourteen package-entry payloads.

Three semantic regressions qualify reclaimed rows with stable actual destinations, unqualified story fallbacks and rejection of complete pairs moving together across pages. All five earlier numbering regressions pass. The first prototype checked effective rather than requested geometry and admitted no changes. The next prototype moved the single-initial-note boundary pair to page one, although Word carries that note to page two; requiring original owner-page stability retains the accepted fallback. Both rejected source/measurement sets and the CRLF archive-helper failure remain recorded.

Of 5257 earlier PDFs, 5254 retain bytes and only three targeted natural cases change; all three have fresh comparisons to trusted unchanged-input Word references. All 5277 frozen outputs, including 2115 visual PDFs, reproduce qualified bytes/pages. Windows passes 2453 of 2454 methods with one optional private skip and no Release warnings/errors. Linux passes 277 checks with six existing missing-Arial skips and two archive SourceLink warnings. Both inventories validate 415 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Earlier after-spacing, composed-layout and cached markup gaps remain explicit.

## RV06-L161: plain-footnote kerning thresholds

Runtime **262067f2** retains the raw `w:kern` cascade through defaults, paragraph styles, character styles and direct overrides, including explicit zero and malformed values. In qualified requested Office-compatible each-page footnote documents, regular 12pt Calibri, Arial and Times New Roman ASCII runs use nominal advances when the threshold is absent, zero or above the font size. Positive thresholds met by the font size retain the existing pair positioning. A derived rendering document shares the policy across layout measurement, PDF emission and glyph signatures; authored settings remain unchanged. Full font measurement, unit print scale, one-line body flow, complete one-line note owners and unchanged page/owner identities bound the retry. Preserve and established review markup retain their measured behavior. Other sizes/styles, tracking, malformed thresholds, wrapped/revised/table/static/column/complex flow, unavailable full fonts and unstable owners retain fallback. Keep chains are not independently qualified for kerning.

Forty-five independent Word controls cross three faces, three sizes and five absent/explicit thresholds. The 135 selected decoded text operations resolve. Twelve 12pt admissions remove the previously unconditional pair adjustment; maximum nominal advance residual is 0.144pt. Thirty-three size/enabled-threshold exclusions retain exact parent PDF bytes. Fifteen further independent controls qualify inherited and direct overrides: five admissions and ten exclusions, with the wrapped-body selector remaining unresolved only in an unchanged exclusion. All seventeen admitted raster pages improve MAE, SSIM and foreground recall. Body baselines differ from Word by at most 0.14pt; link bounds by at most 0.555pt and actual destinations by at most 0.945pt. The seventeen admissions and five trusted earlier counterparts resolve 176 generated labels/pages and 88 actual links to matching Word notes.

All four existing public cases pass their unchanged image and diagnostic gates. The three-note complete-reserve case improves page-one MAE from .6167 to .2014, SSIM from .7825 to .9494 and changed-pixel ratio from .00762 to .00421. Page two has MAE .0505, SSIM .9854 and changed ratio .00131. The L160 partial recorded above is historical. The expected `DOCX_APPROXIMATED_FOOTNOTE` diagnostic remains; metadata ratings do not establish independent human review. Natural visible-spacer controls still have page-one similarity near .608, and 10/16pt controls retain separate font-export residuals. No global width-grid or vertical-phase fit is admitted.

Of 5277 earlier PDFs, 5249 retain exact bytes. All 28 changed outputs have fresh comparisons against trusted unchanged-input Word references, and preserve every decoded text operation, font size, page, baseline, transform and actual note destination from the parent. All 56 changed raster pages improve MAE, SSIM and recall, but absolute fidelity remains partial in the earlier natural/spacing controls. One single-initial-note control retains its known Word body/label pagination mismatch: 27 of 28 controls fully match Word labels/pages, while all 28 preserve their parent identities and actual candidate note targets. Three earlier complete-reserve exclusions change only the independently qualified glyph policy; their reservation fallback remains. These distinct scopes are recorded separately.

Four semantic regressions cover threshold cascade, measurement/emission/navigation coherence, malformed/unqualified fallbacks and preserved review positioning. Five earlier numbering and three complete-reserve regressions pass. Trimmed-tail selector and runtime-report flag helper failures are archived; the runtime itself did not change during qualification. All 5337 frozen PDFs, including 2175 qualified visual outputs, reproduce bytes/pages. Windows passes 2457 of 2458 registered methods with one optional private skip and a clean Release build. Linux passes 281 checks with six existing missing-Arial skips and two archive SourceLink warnings. Both inventories validate 415 public cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred.

## RV06-L162: minimum-height body line leading

Runtime **880c023a** places positive minimum-height body leading above the baseline using the measured natural single-line box. The adjusted paragraph top remains fixed and pagination does not rerun. The bounded admission requires requested Office-compatible plain each-page footnote flow, full font measurement, unit print scale, complete stable note owners and an unwrapped, unmarked body paragraph with zero before/after spacing. Uniform regular 12pt Calibri, Arial and Times New Roman runs in 14-24pt minimum slots qualify; Times slots below 16pt remain guarded. Exact/auto rules, paragraph defaults, minimum note markers, keep chains, rich/mixed/other-size/font, wrapped/revised/table/static/column/non-ASCII or scaled/fallback flow retain the original layout. Glyph plans, line pitch and actual note annotations remain unchanged.

Fifty-two independent Word controls qualify eleven admissions and 41 exact parent-PDF exclusions. The eleven admissions resolve 44 minimum baselines within 0.188pt and improve MAE, SSIM and foreground recall on all eleven raster pages. The primary 44-control inspection resolves 174 operations; two wrapped/table exclusion anchors remain unresolved, with their PDFs unchanged. Eight inherited/direct/composed counterparts resolve 32 operations. The three trusted natural inputs resolve all 120 body/spacer operations by actual occurrence and keep baselines within .212pt. Their page-one SSIM improves from roughly .608 to .984 and MAE falls to .3746-.3770; page two has MAE .1021-.1084 and SSIM .9821-.9844. Capacity, fonts, advances and note destinations retain their accepted identities.

Two public fixtures reproduce all fourteen package-entry payloads with `tools/NewDocxMinimumBodyLineFixtures.ps1`. Both resolve Word pages, labels, minimum baselines and actual note targets, and improve all measured raster metrics. Their existing note-section thresholds remain MAE <=.42, SSIM >=.8, recall >=.8 and changed-pixel ratio <=.0055. Both remain partial on MAE and changed pixels: the 16pt counterpart has MAE .4757, SSIM .9546 and changed ratio .01010; the 20pt counterpart has MAE .5222, SSIM .9471 and changed ratio .01062. The expected `DOCX_APPROXIMATED_FOOTNOTE` diagnostic remains; metadata ratings do not establish independent human review. Across eleven independent admissions, three trusted inputs and two public counterparts, 128 generated labels/pages and 64 actual links resolve to matching Word notes. Bounds remain approximate at up to .555pt, with destinations differing by up to .945pt.

Three semantic regressions qualify metric-derived placement with stable paragraph top/capacity, guarded rules/markers/defaults/styles and coherent emission/preserve geometry with actual destinations. Twelve earlier kerning/numbering/reservation regressions pass. The first built prototype was narrowed before consumption after document-default probes exposed cumulative automatic-pitch differences. The first consumed model worsened all three raster metrics at the 0.2pt Times minimum boundary without improving maximum baseline error; that boundary is rejected. Both source/report sets remain archived. Paragraph defaults need a separate source model: an explicit auto override can retain an inherited point-valued slot. No global font-export or pixel-phase correction is admitted.

Of 5337 earlier PDFs, 5323 retain bytes. All fourteen changed outputs have fresh comparisons against trusted unchanged-input Word references. Every decoded text operation, font size, page, horizontal advance/transform, graphics operation and actual annotation retains its parent identity; only measured minimum baselines change. Twenty-four of 28 raster pages improve all three metrics and four remain equal. Thirteen controls fully match Word labels/pages; the single-initial-note control retains its inherited marker pagination/label mismatch. Earlier reservation, after-spacing, composed-layout and cached markup limitations remain explicit.

All 5391 frozen PDFs, including 2229 qualified visual outputs, reproduce bytes/pages. Windows passes 2460 of 2461 methods with one optional private skip and a clean Release build. Linux passes 284 checks with six existing missing-Arial skips and two archive SourceLink warnings. Both inventories validate 417 public cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred.

## RV06-L163: numeric line-spacing overrides through inheritance

Runtime **08399875** corrects paragraph inheritance when a local numeric line value changes automatic/point units. Previously, an automatic child could retain a parent's 16pt point height, and a line-only child could retain the inherited exact/minimum rule in its effective metadata. Numeric overrides now replace both derived units together; a missing local rule uses automatic spacing. Document defaults, named paragraph styles and based-on chains share this resolution. Structure mode counts agree with the effective units. Source packages, numbering settings and other paragraph properties remain authored.

Sixty-two independent Word controls qualify 29 admissions and 33 exact parent-PDF exclusions. Fifty-one note controls resolve all 204 selected body baselines; eleven ordinary/table/bold-14pt controls resolve 44. Table displacement falls from up to 38.546pt to .056pt, while the bold counterpart stays within .221pt. All 29 admitted raster pages improve MAE, SSIM and foreground recall. Numeric mode changes preserve every decoded text operation, font size, page, horizontal position/advance and transform. Baselines, note regions, table line geometry and actual annotation coordinates follow corrected units. Three default/minimum counterparts retain a roughly 1.5pt minimum-body baseline gap because that leading correction remains separately guarded.

Two public fixtures reproduce all fourteen ZIP entry payloads with `tools/NewDocxLineSpacingInheritanceFixtures.ps1`. Both improve and retain expected approximate-footnote diagnostics, but remain partial at unchanged .42/.8/.8/.0055 note-section thresholds. The default-auto case has MAE .5133, SSIM .9479 and changed ratio .01040; the style line-only case has MAE .5351, SSIM .9450 and changed ratio .01073. Both fail MAE and changed pixels. Metadata ratings do not establish human review.

Of 5391 earlier PDFs, 5388 retain exact bytes and three earlier default probes have fresh parent/candidate comparisons against qualified unchanged-input Word references. All three preserve pages, decoded text/font/X/advance/transform sequences and improve all three raster metrics. Current and earlier admissions resolve 200 labels/pages and 100 actual note links to matching Word notes, with four distinct destinations per note input. Word link bounds differ by at most .042pt and destinations by .945pt. The earlier default/minimum baselines remain partial near 1.5pt; no global pixel-phase correction is admitted.

Four new reader-to-layout/navigation regressions and fifteen earlier minimum-leading, kerning, numbering and reservation checks pass. Independent Word measurements rejected recomputing units from an attribute-wise merged line/rule tuple: a line-only override is automatic, while a standalone rule does not supply a new numeric setting. A consumed prototype that normalized standalone rules activated the retained ordinary minimum-baseline defect and was narrowed. Standalone rule tokens retain their earlier metadata/geometry. The first full Windows run also caught an outdated metadata assertion expecting an inherited exact rule under a direct line-only override. That assertion now checks automatic units and structure counts; the failed run, successful first PDF/Linux evidence, assemblies and source remain archived with an explicit path map.

The final source reproduces all 5455 frozen PDF bytes/pages, including 2293 qualified visual outputs. Windows passes 2464 of 2465 methods, with one optional private skip and zero Release warnings/errors. Linux passes 289 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventories validate 419 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Default/ordinary minimum leading, font export/pixel phase, after-spacing, composed layouts and cached markup retain separate limitations.

## RV06-L164: minimum body leading in defaults and ordinary flow

Runtime **223474f7** extends the qualified minimum-line metric model to coherent paragraph defaults and ordinary body flow without requiring unused footnote settings. Positive minimum-height excess moves above the baseline while preserving the paragraph top, slot pitch, capacity, glyph advances and actual note geometry. Requested Word-compatible output, full fonts, unit scale, simple single-line flow, regular 12pt Calibri/Arial/Times and 14..24pt minimum slots bound the change. Times below 16pt, exact/automatic spacing, keep rules, painted/rich/wrapped text and composed layouts retain fallback behavior. The shared plain source predicate preserves earlier note-numbering eligibility.

Seventy independent Word 16 controls cover 44 ordinary font/slot/rule and rich/wrapped/keep/table/revision counterparts, 18 inherited/direct note-default counterparts and eight note-default exclusions. Twenty-eight admissions resolve 112 selected baselines within .188pt of Word; 42 exclusions preserve exact parent PDF bytes. All 28 admitted raster pages improve MAE, SSIM and foreground recall. Decoded text, font sizes, pages, horizontal positions/advances, transforms, graphics and actual annotations retain parent identity. No global glyph or raster-phase tuning is admitted.

Two public controls reproduce thirteen ZIP payloads through `tools/NewDocxMinimumBodyFlowFixtures.ps1`. Both improve but remain partial at unchanged .42/.8/.8/.0055 thresholds. The note-default 16pt case has MAE .4757, SSIM .9546, recall .9356 and changed ratio .01010, failing MAE and changed pixels with the expected approximate-footnote diagnostic. The ordinary 20pt case has MAE 1.1869, SSIM .7004, recall .7391 and changed ratio .01341, failing all four image gates without unsupported-feature diagnostics. This exposes remaining ordinary glyph-spacing/font-export gaps; metadata ratings do not establish visual review.

Of 5455 prior PDFs, 5448 preserve exact bytes and seven earlier default/minimum controls have fresh comparisons against qualified unchanged-input Word references. All seven improve all three raster metrics, preserve pages/advances/graphics/actual annotations and resolve 28 selected baselines within .163pt. Current/public and earlier admissions cover 37 nonoverlapping raster pages, 148 selected baselines, 208 matching labels/pages and 104 actual note links, with four distinct destinations per note input. Word link bounds differ by at most .555pt and destinations by .945pt. Three new semantic regressions and nineteen earlier inheritance/minimum/kerning/numbering/reservation checks pass.

The frozen source reproduces 5527 PDF bytes/pages, including 2365 qualified visual outputs. Windows passes 2467 of 2468 methods with one optional private skip and zero Release warnings/errors. Linux passes 292 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventories validate 421 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Ordinary glyph spacing, font export/pixel phase, after-spacing, broader note leading, composed layouts and cached markup retain separate limits.

## RV06-L165: inherited kerning in ordinary body flow

Runtime **7a0ce2ee** extends the plain-note kerning policy to ordinary body flow without requiring unused note settings. In requested Word-compatible output, absent/zero thresholds and thresholds above the nominal font size use nominal advances; positive thresholds met by the font size retain pair positioning. Regular 12pt Calibri/Arial/Times, full fonts, unit scale, no related stories or note references, and simple single-line unpainted flow bound the extension. Rich text, other sizes, keep rules, wrapped text, columns, static content, tables and malformed thresholds retain fallback. Raw reader metadata and Preserve output remain unchanged. Minimum-line slots, page capacity and qualified minimum baselines retain their prior geometry.

Ninety-five independent Word 16 controls cover font/size/threshold counterparts, defaults, paragraph/character styles, based-on chains, explicit zero, minimum-line interactions and excluded layouts. Thirty-five admissions resolve 212 lines and improve MAE, SSIM and foreground recall on all 35 raster pages. Sixty exclusions preserve exact parent PDF bytes. Three of these verify that the threshold is inclusive at exactly 12pt, matching the enabled 8pt counterpart for each qualified font. Text/font/page/baseline/natural-width/transform sequences, graphics, font resources and actual annotation arrays retain parent identity; admitted parent/candidate/Word arrays are empty.

Two new public controls reproduce twelve ZIP payloads through `tools/NewDocxOrdinaryKerningFixtures.ps1`. Both pass unchanged .42/.8/.8/.0055 gates, each with MAE .2036, SSIM .9514, foreground recall .9333 and changed ratio .00394, without unsupported-feature diagnostics. Metadata ratings do not establish independent human review.

Of 5527 prior PDFs, 5484 retain exact bytes and 43 changed ordinary/default/minimum controls have fresh comparisons against qualified unchanged-input Word references. All 43 improve every measured raster metric and preserve page counts. Current/public and earlier admissions cover 80 nonoverlapping raster pages and 564 resolved lines. Visible advances differ from Word by at most .219pt; the previously quoted complete-line .658pt residual includes invisible paragraph-end whitespace operations and is not a visible-font-width measurement. Baselines differ by at most .188pt in the new independent controls and .198pt in earlier controls. The earlier ordinary-minimum 20pt public case now has MAE .3766, SSIM .9389, recall .9237 and changed ratio .00694: it passes three image gates but remains partial against the unchanged .0055 changed-pixel limit.

Three new semantic regressions and 22 earlier flow/inheritance/kerning/numbering/reservation checks pass. Helper whitespace-font filtering and a corrected test compile failure remain in ignored evidence. The frozen source reproduces 5624 PDF bytes/pages, including 2462 qualified visual outputs. Windows passes 2470 of 2471 methods with one optional private skip and zero Release warnings/errors. Linux passes 295 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventories validate 423 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Marked minimum leading, after-spacing, broader note flow, composed layouts and cached markup remain separate limitations.

## RV06-L166: marked minimum leading and after-spacing

Runtime **5fe5b93d** extends minimum-height leading to plain each-page footnote flow containing source-owned superscript automatic markers and numeric after-spacing up to 8pt. Requested Word-compatible output, full fonts, unit scale, complete one-line body/note owners, regular 12pt Calibri/Arial/Times, 16..24pt marked minimum slots and zero before-spacing bound the extension. The automatic marker must match its actual reference/run ownership, offset and numeric display text; arbitrary superscripts do not qualify. Positive actual-font-metric excess moves body and marker baselines together. Paragraph tops, slot pitch, capacity and note regions remain fixed. Actual note annotations use the same paragraph slot top and extend downward through admitted after-spacing. Raw source labels and Preserve behavior remain unchanged.

Fifty-four independent Word 16 controls cross three fonts, 16/20/24pt minimum slots and 0/4/8pt after-spacing, exact/auto counterparts, paragraph/default inheritance, source/layout exclusions and three two-page owner/capacity cases. Thirty-seven admissions resolve 377 body lines, including every capacity spacer, 148 markers and 148 actual links. Seventeen exclusions retain exact parent PDF bytes. Body baselines differ by at most .308pt; Times markers retain up to .740pt of existing vertical residual. All 40 admitted raster pages improve MAE, SSIM and foreground recall. Text/font/page/X/advance/transform sequences, font resources, graphics, source-backed placed notes, actual footer baselines and note destinations retain parent identity. Actual link bounds differ by at most .555pt and destinations by .945pt, with four distinct destinations per admitted input.

Two public controls reproduce fourteen ZIP payloads through `tools/NewDocxMarkedMinimumFixtures.ps1`. Both improve but remain partial at unchanged .42/.8/.8/.0055 gates. The zero-after 20pt case has MAE .7277, SSIM .9126, recall .8986 and changed ratio .01265; the 8pt-after case has MAE .9460, SSIM .8764, recall .8665 and changed ratio .01490. Both fail MAE and changed pixels, with only the required approximate-footnote diagnostic. Metadata ratings do not establish independent human review.

Of 5624 prior PDFs, 5623 retain exact bytes and the retained L162 minimum-marker case has a fresh comparison against its qualified unchanged-input Word reference. It improves all three raster metrics, preserves page count and resolves eight body lines within .163pt and four markers within .180pt. Combined current/public/earlier admissions cover 43 nonoverlapping raster pages, 401 body lines, 160 markers, 320 matching labels/pages and 160 actual links. Four new semantic regressions and 29 earlier checks pass. Test snapshot/fixture assumptions and a PowerShell wrapper's unset native exit status are retained as helper failures; no consumed runtime model was rejected.

The frozen source reproduces 5680 PDF bytes/pages, including 2518 qualified visual outputs. Windows passes 2474 of 2475 methods with one optional private skip and zero Release warnings/errors. Linux passes 299 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventories validate 425 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Before-spacing, ordinary after-spacing, broader marker metrics, larger slots, continued notes, composed layouts and cached markup retain separate limits.

## RV06-L167: minimum note body before-spacing

Runtime **2f3fa461** extends minimum body leading to numeric 0..8pt before-spacing in the existing plain each-page footnote scope. Requested Word-compatible output, full fonts, unit scale, complete single-line body/note owners, regular 12pt Calibri/Arial/Times and 16..24pt spaced/marked minimum slots bound the extension. Source-owned superscript automatic markers remain required. Positive actual-font-metric excess shifts body and marker baselines together; authored spacing, paragraph tops, pitch, capacity and note placement stay fixed. The existing qualified minimum-slot annotation model includes only the remaining before-gap after predecessor after-spacing and admitted current after-spacing. Raw reader metadata and Preserve behavior remain unchanged.

Eighty-seven independent Word 16 controls cross three fonts, 16/20/24pt slots, 4/8pt before-spacing and 0/4/8pt after-spacing, with style/default inheritance, mixed gap sequences, first-note page tops, three two-page capacity controls, exact/auto counterparts and source/layout exclusions. Sixty-nine admissions resolve 630 body lines, 276 markers and 276 actual links; all 72 raster pages improve MAE, SSIM and foreground recall. Eighteen exclusions retain exact parent PDF bytes. Body baselines differ by at most .298pt; Times markers retain up to .740pt of existing vertical residual. Actual link bounds differ by at most .555pt and destinations by .945pt, with four distinct targets per input. Text/font/page/X/advance/transform sequences, font resources, graphics, source-backed placed notes, actual footer baselines and note destination parameters retain parent identity.

Two public counterparts reproduce fourteen ZIP payloads through `tools/NewDocxMinimumBeforeFixtures.ps1`, including a document-default first-note case with explicit zero-before note paragraphs. Both remain partial at unchanged .42/.8/.8/.0055 gates, failing MAE and changed pixels with only the required approximate-footnote diagnostic. `docx-ladder-04-note-minimum-before-after-4`: MAE 0.9460, SSIM 0.8764, recall 0.8665, changed ratio 0.01490; `docx-ladder-04-note-minimum-before-inherited-first`: MAE 0.8233, SSIM 0.8890, recall 0.8774, changed ratio 0.01331. Metadata ratings do not establish independent human review.

Of 5680 prior PDFs, 5679 retain exact bytes and 1 have fresh comparisons against qualified unchanged-input Word references; every change improves all three raster metrics without changing pages. Combined admissions cover 75 nonoverlapping raster pages, 653 body lines, 288 markers, 576 matching labels/pages and 288 actual links. Four new semantic regressions and 33 earlier checks pass. Failed synthetic-font page-count assumptions and the note-default isolation correction remain archived; no consumed runtime model was rejected.

The frozen source reproduces 5769 PDF bytes/pages, including 2607 qualified visual outputs. Windows passes 2478 of 2479 methods with one optional private skip and zero Release warnings/errors. Linux passes 303 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventories validate 427 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Ordinary before/after-spacing, larger slots, broader marker metrics, continued notes, composed layouts and cached markup retain separate limits.

## RV06-L168: ordinary minimum body before/after-spacing

Runtime **fb0e2851** (implementation `499c580e`, nullable parameter correction `fb0e2851`) extends ordinary minimum body leading to numeric 0..8pt before/after spacing. Requested Word-compatible output, full fonts, unit scale, complete one-column single-line body flow, one uniform visible regular 12pt Calibri/Arial/Times family and 16..24pt spaced minimum slots bound the extension. New spaced admissions qualify numeric spacing and visible runs across the entire body once. Legacy zero-spacing and note admissions keep their existing gates. Positive actual-font-metric excess shifts only minimum baselines/insets; automatic neighbors, authored paragraph tops, spacing, pitch, capacity and glyph plans stay fixed. Raw metadata and Preserve output remain unchanged, with actual annotation arrays empty.

One hundred two independent Word 16 controls cross three fonts, 16/20/24pt slots and eight positive numeric before/after combinations, with style/default inheritance, mixed collapsed gaps, three two-page capacity cases, exact/auto counterparts and twelve source/layout guards. Eighty-four admissions resolve 753 body lines: 417 minimum lines and 336 automatic neighbors. All 87 admitted raster pages improve MAE, SSIM and foreground recall. Eighteen exclusions retain exact parent PDF bytes. Every baseline is within .305pt of Word. Text/font/page/X/advance/transform sequences, font resources, graphics, automatic-neighbor baselines and page counts retain parent identity; only selected minimum baselines change.

Two public counterparts reproduce twelve ZIP payloads through `tools/NewDocxOrdinaryMinimumSpacingFixtures.ps1`. Both improve but retain unchanged .42/.8/.8/.0055 gates and empty diagnostics. The direct before-4/after-4 case has MAE .3766, SSIM .9389, recall .9237 and changed ratio .00694; it fails changed pixels. The inherited before-8/after-4 case has MAE .4730, SSIM .9152, recall .9037 and changed ratio .00778; it fails MAE and changed pixels. Metadata ratings do not establish independent human review. Combined admissions cover 89 raster pages, 769 body lines and 425 minimum lines. Four new semantic regressions and 37 earlier checks pass.

All 5769 prior PDFs retain exact bytes and pages. The frozen source reproduces 5873 PDFs, including 2711 qualified visual outputs. Its fresh exact PDF identities transfer the original Word/raster reports after the nullable-only correction. The original consumed prototype, failed warning build, original design and failed newline-converting design reconstruction remain archived. The correction changes only one parameter annotation; no runtime branch/model or image threshold is changed by that follow-up.

Windows passes 2482 of 2483 methods with one optional private skip and zero final Release warnings/errors. Linux passes 307 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventories validate 429 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Contextual minimum spacing, broader fonts/sizes, custom/continued notes, composed layouts and cached markup retain separate limits.

## RV06-L169: contextual minimum body leading and actual note gaps

Runtime **5fddb2ca** extends plain ordinary and complete each-page note minimum leading to contextual spacing with trustworthy style identities. Requested Word-compatible output, full fonts, unit scale, complete one-column single-line body flow, one uniform visible regular 12pt Calibri/Arial/Times family, numeric 0..8pt before/after spacing and 16..24pt spaced/marked slots bound the extension. The complete visible body qualifies once. Explicit named styles must resolve; omitted identifiers require a loaded implicit Normal or unambiguous declared default. Only actual source-owned automatic superscript note marks qualify. Missing, ambiguous and unresolved identities remain excluded. Raw reader metadata, Preserve output, legacy zero-spacing and non-contextual gates keep their earlier behavior.

One hundred six independent Word 16 controls cover three fonts, 16/20/24pt slots, direct numeric gaps, same/different named styles, omitted/explicit declared defaults, implicit Normal, paragraph-style/document-default contextual flags, current/predecessor contribution transitions, mixed gaps and six two-page counterparts. Seventy-four admissions resolve 754 body lines, including 606 minimum lines and 148 source-owned markers. All 80 admitted raster pages improve MAE, SSIM and foreground recall. Thirty-two source/rule/layout exclusions retain exact parent PDFs. Ten unresolved body mappings remain reported in excluded wrapped/column controls; no admitted owner remains unresolved. An analysis helper was corrected to use the explicit complete tag/owner map; the early minimum-baseline results remain unchanged. A Word 255-character export-path failure is retained; completed profile-verified references were reused and the remaining references use shorter output paths.

Body baselines remain within .275pt of Word; existing Times marker residuals remain within .740pt. All 148 actual note rectangles resolve within .555pt and destinations within .945pt. The predecessor owns its remaining after-contribution, the current paragraph owns only its remaining before-gap, and same-style contextual successors suppress current after-spacing. Minimum baselines/insets and qualified marked slot metadata change; authored paragraph tops, spacing, pitch, capacity, automatic-neighbor baselines, complete note pages, footer baselines, actual destinations, text/font/page/X/advance/transform sequences, font resources and graphics retain parent identity. Ordinary actual annotation arrays remain empty. Four new semantic regressions and 41 earlier focused checks pass.

Two public counterparts reproduce all thirteen ZIP payloads in two independent generations through `tools/NewDocxContextualMinimumFixtures.ps1`; a final reproduction also matches. Both retain unchanged .42/.8/.8/.0055 gates. The ordinary case has MAE .3766, SSIM .9389, recall .9237 and changed ratio .00694, failing changed pixels with empty diagnostics. The alternating-style note case has MAE .9460, SSIM .8764, recall .8665 and changed ratio .01490, failing MAE and changed pixels with the existing approximate-footnote diagnostic. Metadata ratings do not establish independent human review.

Of 5873 earlier outputs, 5871 retain exact bytes; two earlier contextual guard inputs improve against their hash-verified Word references through fresh parent/prototype comparisons. All page counts remain fixed. Combined new and earlier admissions qualify 84 raster pages, 786 body lines, 630 minimum lines, 156 markers, 312 body/footer labels/pages and 156 actual links. The frozen source reproduces 5981 PDFs, including 2819 qualified visual outputs, transferring comparison evidence only through fresh exact PDF identity.

Windows passes 2486 of 2487 registered methods with one optional private skip and zero Release warnings/errors. Linux passes 311 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventory validates 431 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Broader fonts/sizes, automatic/line-unit spacing, custom/continued notes, composed layouts and cached markup retain separate limits.

## RV06-L170: regular 11pt body kerning and minimum leading

Runtime **75cf87c2** qualifies regular 11pt ordinary and complete each-page note body flow. Requested Word-compatible output, full fonts, unit scale, complete left-aligned one-column single-line flow, trusted named or loaded default styles, one uniform Calibri/Arial/Times family, numeric 0..8pt before/after spacing and 16..24pt explicit pitches bound the extension. Automatic counterparts use single spacing. Only actual source-owned automatic superscript markers qualify. Note-story letters retain regular 12pt in the same family. Legacy 12pt admissions retain their independent gates; raw reader metadata and Preserve remain unchanged.

One hundred forty-six independent Word 16 controls cross three faces, 16/20/24pt minimum slots, contextual/numeric/zero gaps, known alternating styles, document/paragraph font inheritance, kerning thresholds, exact/automatic counterparts and six actual two-page capacity cases. Eighty-four admissions resolve 834 body lines, including 594 minimum lines, and 168 source-owned markers. Sixty-two exclusions retain exact parent PDFs, including twelve centered/right-aligned counterparts outside the existing plain frame. Excluded unresolved mappings remain in the reports; no admitted body, marker or note label remains unresolved.

The existing emitter already matches Word's 11.04pt body and 6.96pt body-marker sizes. The source threshold policy uses the logical 11pt size: absent, zero and above-threshold kerning disables pairs; a met positive threshold retains them. All six face/flow pairs verify identical Word spans for absent/zero/above-threshold controls and changed spans for met thresholds. Positive minimum excess moves the visible baseline while retaining paragraph tops, pitch, spacing contributions and capacity. All 90 admitted raster pages strictly improve MAE, SSIM and foreground recall. Body baselines remain within .228pt and markers within .194pt. All 168 actual note rectangles resolve within .566pt and destinations within .945pt; complete note placement, footer anchor positions and actual destination parameters retain parent identity. Font resources, graphics, text/font/page/transform sequences and automatic/exact baselines retain identity. Kerning advances and body-marker X change as measured; qualified 12pt footer advances also normalize.

The 11pt nominal export grid remains a separate approximation: visible-span errors reach .7443pt. Fifty-six isolated capacity spans have up to .3411pt more error than the parent, while their complete pages improve on all three raster metrics. The initial assumption that every individual width would improve is retained with the diagnostic report and source. The accepted scope corrects the measured source kerning policy and minimum leading together; it does not claim complete horizontal parity. Existing public gates and 12pt thresholds are unchanged.

Five new semantic regressions and 45 earlier focused checks pass. Initial authored test setup left inherited 12pt runs and lacked the synthetic face's AB pair; its three failures and original source are retained, and corrected 11pt inputs pass. A marker-analysis helper was corrected before broad consumption to apply excess only to the minimum rule. One Word activation failure is retained; checkpointed references were reused and the resumed batch completed without terminating existing Word processes.

Two public counterparts reproduce all thirteen ZIP payloads twice through `tools/NewDocxElevenPointFixtures.ps1`; final reproduction also matches. Both remain partial on MAE and changed pixels at unchanged .42/.8/.8/.0055 gates. The ordinary case has empty diagnostics, and the note case retains its approximate-footnote diagnostic. Metadata ratings do not establish independent human review. Combined public and independent admissions qualify 92 raster pages, 850 body lines, 606 minimum lines, 172 markers, 344 body/footer labels/pages and 172 actual links.

All 5981 earlier PDFs retain exact bytes and page counts. The final frozen source reproduces 6129 PDF byte/page identities, including 2967 qualified visual outputs; fresh identity transfers the Word evidence to the committed assembly. Windows passes 2491 of 2492 registered methods with one optional private skip and zero Release warnings/errors. Linux passes 316 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventory validates 433 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Export-grid compensation, broader sizes/faces, other alignments, richer/composed layouts and custom/continued notes remain separate work.


## RV06-L171: nominal 11pt glyph and automatic-marker advances

Runtime **4c847e03** keeps the accepted regular 11pt frame boundaries and physical 11.04pt body glyph size. Nominal baseline runs advance at logical 11pt metrics through per-glyph TJ compensation, retaining the last natural physical glyph width. Met source thresholds keep their accepted pairs. Exact source-owned automatic superscript marks keep physical 6.96pt emission and use unkerned pairs only when every visible baseline body run has an absent, zero or unmet threshold. Related 12pt stories and all excluded frames retain their qualifications. Source metadata and Preserve remain unchanged.

Eighteen independent repeated-glyph Word controls cross Calibri, Arial and Times, absent/zero/above/met thresholds and 12/14pt guards. Embedded Word and local hmtx metrics agree. Per-glyph 600dpi rounding was rejected; actual glyph inspection qualifies logical advances. Nine nominal controls improve every complete page on MAE, SSIM and foreground recall; nine guards retain exact PDFs. Three independent Calibri repeated-i spans have a small worsening despite page improvements. Original hypotheses and diagnostics remain recorded.

Fresh qualification of all 148 accepted source counterparts admits 80 nominal corrections and preserves 68 exact guards, including met-body thresholds. All 86 changed pages improve on all three raster metrics. Individual glyph inspection resolves 802 visible body spans and 33,252 starts, trimming invisible trailing spaces before measuring endpoints. The earlier nominal-span helper included terminal whitespace for inherited paragraph fonts; those historical measurements remain retained and are distinguished from visible glyph endpoints. Twelve source spans have at most .0403pt additional error, while the maximum visible span falls from .6956pt to .2157pt. Every admitted span and relative start stays within .219pt; absolute starts remain below .25pt.

Nine additional automatic-marker controls cross the three faces and raw 6/7/8pt thresholds. Word leaves their automatic number pairs unkerned. Six Arial/Times controls improve on all three raster metrics, while three Calibri controls keep exact PDFs. A broader marker-only change in a met-body Times control worsened all three metrics and was rejected. Its source, consumed second prototype and reports remain available; the final scope keeps that combination's accepted output. The final prototype transfers measured raster/operation evidence only through 6,158 fresh exact PDF identities and 177 fresh layout inspections.

Two new public cases reproduce all thirteen ZIP payloads twice and in the final frozen preparation through `tools/NewDocxNominalGlyphFixtures.ps1`. The repeated-glyph case passes the existing .42/.8/.8/.0055 gates with empty diagnostics. The Arial note case remains partial on MAE and changed pixels and retains its approximate-footnote diagnostic. Two earlier public counterparts now remain partial only on changed pixels. Gates remain unchanged; metadata ratings do not establish independent human review.

Combined qualification covers **177 controls /97 admissions /80 exact guards /103 improving raster pages /938 body lines /702 minimum lines /188 markers and actual links /376 body/footer labels and pages**. Maximum body baseline error is .228pt, marker baseline .194pt, actual note bounds .051pt and destinations .945pt. Font resources, graphics, text/font/page/transform sequences and vertical positions retain identity. Paragraph pitch and spacing, complete note placement, footer anchors and destination parameters remain fixed. Source-owned marker line widths change only by the independently checked difference between logical 7pt pair advances and physical 6.96pt emission.

Of 6,129 earlier PDFs, **6,049 remain exact**; all 80 changed PDFs have fresh appropriate Word comparisons, with fixed page counts. Frozen source reproduces **6,158 PDF byte/page identities**, including **2,996 qualified visual outputs**, with twelve source Git blob identities. Windows passes 2,496 of 2,497 registered methods with one optional private skip and zero Release warnings/errors. Linux passes 321 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventory validates 435 cases across ten families. Local 0.1.5 package smoke passes; release preparation remains deferred. Broader sizes/faces, other alignments, richer/composed/custom/continued note flow and remaining met-threshold horizontal residuals require separate qualification.

## RV06-L172: bounded 11pt met-threshold glyph and pair advances

Runtime **be05fb98**, integrated runtime root **77672224**, extends the strict regular 11pt frame to ordinary Times and complete Calibri/Arial note flow only when every nonempty baseline body run shares one positive source threshold at most 22 half points. Preceding body advances and pairs use logical 11pt metrics while visible body/marker sizes remain 11.04/6.96pt and final natural widths remain physical. Exact source-owned automatic markers are unkerned within the same frame. Other met face/flow combinations, mixed thresholds and related 12pt stories retain accepted policies. Source metadata and Preserve remain unchanged.

Twenty-one independent Word controls resolve 168 body runs and 6,225 glyph starts. Actual embedded Word/local hmtx widths match for 755 mapped entries. A broader consumed prototype improved all 15 affected pages but failed unchanged glyph bounds in eight controls, including 41 visible-glyph failures. Its pin, source and reports remain retained. Rounded 1000em alternatives were rejected. The initial note model incorrectly carried corrections across independently anchored operations; corrected source anchors reset them per operation and change later marker anchors only by the measured logical marker pair-width delta. Neither rejected model establishes admission.

The narrower prototype qualifies seven independent admissions and fourteen exact guards. Four earlier source counterparts receive fresh Word comparisons; the remaining 173 accepted source controls retain exact PDFs. Two new public counterparts reproduce all thirteen ZIP payloads twice through `tools/NewDocxMetGlyphFixtures.ps1`. The Times ordinary case passes unchanged .42/.8/.8/.0055 gates with empty diagnostics; the Arial note case remains partial on MAE and changed pixels with its approximate-footnote diagnostic. No independent human rating is supplied.

Combined qualification covers **200 controls /13 admissions /187 exact guards /13 improving pages /104 body lines /84 minimum lines /28 markers and actual links /56 body/footer labels and pages**. All admitted pages improve MAE, SSIM and recall. Individual inspection resolves 4,163 admitted glyph starts with maximum .170pt visible-span, .215pt relative-start and .239pt absolute-start errors. Body baselines remain within .148pt, markers .188pt, actual link bounds .051pt and destinations .945pt. Trailing whitespace is measured separately. Text/font/page/transform sequences, graphics and font resources retain identity. Complete note placement, paragraph pitch/spacing, footer anchors and actual destination parameters remain fixed. Source-owned marker line widths change only by the independently checked logical 7pt pair-width delta relative to physical 6.96pt emission.

A label helper initially assumed older public anchor wording and left the new pair probes unresolved. Exact source-owned pre-marker text, note IDs and source blocks resolve all labels without changing the runtime, inputs or references; original reports and helper are retained. Fifty-nine focused checks pass, including positive-threshold boundaries and mixed-threshold fallback.

Of **6,158 earlier PDFs**, **6,154 remain exact** and four have fresh appropriate Word qualification, with no page-count changes. Frozen source reproduces **6,181 PDF byte/page identities**, including **3,019 qualified visual outputs**. Windows passes 2,500 of 2,501 registered methods with one optional private skip and zero Release warnings/errors. Linux passes 325 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventory validates 437 cases across ten families. Root passes 273 focused checks and reproduces 3,019 fresh PDF identities with eleven source Git blob identities. Local 0.1.5 package smoke passes; release preparation remains deferred. Broader fonts, sizes, mixed/rich/composed/continued note flow and full horizontal parity remain separate work.

## RV06-L173: exact source owners for wrapped each-page numbering

Runtime **0d71da89**, integrated runtime root **c83c87d1**, corrects each-page automatic footnote labels for wrapped regular 12pt Arial/Times body paragraphs. Plain single-section frames, zero numeric spacing, single automatic pitch, full fonts and unit scale bound the retry. Each marked paragraph must remain on one physical page, with one exact source run/offset/display-text segment and one complete one-line footnote. Source-authored digits and metadata stay unchanged. Existing kerning, minimum-leading and complete-reserve callers retain their one-line owner gate; only the explicit numbering retry uses wrapped ownership.

Twenty-four independent Word controls cross three faces, early/later-line marks, run/paragraph page breaks and decimal/decimalZero numbering near ten. Initial output has 192 mismatched body/footer labels across 48 pages, with all source owners and links resolved. The broad consumed prototype fixes those labels and improves all pages on MAE, SSIM and recall, but Calibri retains horizontal marker errors up to 60.471pt for later marks and 1.204pt for early marks. Its source, binary, reports and two changed earlier PDFs remain archived. The narrower scope qualifies sixteen Arial/Times admissions and eight exact Calibri guards.

Two public counterparts reproduce all fourteen ZIP payloads twice through `tools/NewDocxWrappedFootnoteFixtures.ps1`. Both remain partial at unchanged .42/.8/.8/.0055 gates: Arial fails changed pixels; Times fails MAE and changed pixels. Both retain the approximate-footnote diagnostic. Independent human ratings and full Office parity are not claimed.

Combined qualification covers **26 controls /18 admissions /8 exact guards /36 improving pages /72 source-owned body marks /72 footer labels /144 matching labels and pages /72 actual links**. Actual rectangles differ from Word by at most .046pt and destinations .799pt. Body wrapping, vertical positions, pitch/spacing, complete notes/reserve, footer baselines, graphics and source full-font metrics retain parent identity. Authored body/footer text and digits retain identity after removing only generated numeric operations. Marker physical sizes and baselines stay fixed; Times retains its inherited .660pt superscript-baseline residual. The geometry helper initially applied an absolute-baseline assumption to that unchanged residual. Its failed assumption is retained; final qualification requires zero parent/current baseline delta and records the remaining Word error explicitly. The .6pt link and 1pt destination gates are unchanged.

Four new semantic methods and 59 earlier focused checks pass. They exercise exact early/later segments, source digits and thresholds, explicit legacy gates, missing/shifted/duplicate owners, same-paragraph page splits, unsupported font/style/spacing/keep/custom/column/continued scopes, stable complete-note pages and distinct actual destinations. An initial source-block helper omitted the separate reader element inserted for pageBreakBefore; source XML and renderer indices are now recorded separately. A non-idempotent temporary-directory retry stopped before overwriting reports and resumed with a new prefix. Original helpers/reports and preparation corrections remain retained.

All **6,181 earlier PDFs** retain exact bytes and page counts. Frozen source reproduces **6,207 PDF byte/page identities**, including **3,045 qualified visual outputs**. Windows passes 2,504 of 2,505 registered methods with one optional private skip and zero Release warnings/errors. Linux passes 329 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventory validates 439 cases across ten families. Root passes 277 focused checks and reproduces 3,045 fresh PDFs with ten source Git blob identities. Local 0.1.5 package smoke passes; release preparation remains deferred. Calibri advance/wrapping residuals and broader composed or continued-note flow require separate qualification.

## RV06-L174: nominal advances in complete wrapped Calibri note frames

Runtime **90b34391**, integrated runtime root **3f72d4d4**, corrects regular 12pt Calibri body advances before each-page automatic footnote numbering in plain complete-note frames. Raw body kerning thresholds must be absent, zero or above 24 half points. Zero numeric spacing, single automatic pitch, full fonts, unit scale, exact source run/offset/display-text ownership and complete one-line notes bound the retry. The original frame must already wrap; newly wrapped output from a former one-line retry retains fallback. Source metadata, authored digits, physical font sizes and Preserve remain unchanged. Existing minimum-leading and complete-reserve callers keep their one-line gates.

Read-only measurements cover 680 actual Word body operations and 816 embedded/local glyph widths without a mismatch. Nominal advances fit absent/zero/unmet thresholds; positive met thresholds retain a separate pair policy. A 600dpi per-glyph quantization model was rejected because it worsens starts/spans. Equal embedded font widths do not establish equal Word PDF positions. Consumed source, pins, raw/rounded/pair/grid hypotheses and helper corrections remain in ignored evidence.

Sixteen fresh held-out Word controls cross two page-break forms, early/later marks, nominal zero/unmet-26 thresholds, met-24 thresholds and 14pt guards. Eight retained Calibri controls, two public counterparts and two earlier changed wrapped-note cases complete qualification. **28 controls /20 corrections /8 exact guards /39 strictly improving pages /80 source-owned body marks /80 footer labels /160 matching labels and pages /80 actual links** qualify. All 456 actual Word body-line sequences match over 31,314 visible characters. Remaining absolute glyph-start residual reaches .352pt and visible span .328pt; body baselines differ by at most .193pt. Physical marker baselines retain parent identity, with .233pt remaining Word residual. Link bounds are within .048pt and destinations .945pt. Complete note/reserve placement, vertical geometry, source full-font metrics, graphics, authored text/digits and actual destinations retain identity. Complete Office glyph/baseline parity is not claimed.

The public zero-early and unmet-later fixtures reproduce all fourteen ZIP payloads twice through `tools/NewDocxWrappedCalibriFixtures.ps1`. Both retain the approximate-footnote diagnostic and remain partial at unchanged .42/.8/.8/.0055 image gates, failing MAE and changed pixels. No independent human rating is supplied. Of 6,207 earlier outputs, 6,197 remain exact and ten change with unchanged page counts: eight retained Calibri controls and two older wrapped-note cases. All ten are independently requalified against Word. The two older cases receive fresh Office exports under the same verified Word 16/printer profile.

Four new semantic methods and 63 earlier focused checks pass. They cover nominal/met/unmet/malformed thresholds, exact source ownership, unsupported fonts/sizes/style/spacing/keeps/custom notes, original one-line fallback, retained reserve/minimum-leading boundaries, source immutability, Preserve and actual destination pages. One helper initially removed authored digit-only operations after PDF splitting changed; it now excludes only the eight generated 8.04pt numeric labels. Another assigned a whitespace-only operation to a visible body line; it now compares only visible glyphs. Original assumptions remain retained and no consumed PDF/source pin was overwritten.

Frozen source reproduces **6,233 PDF byte/page identities**, including **3,071 qualified visual outputs**. Windows passes 2,508 of 2,509 methods with one optional private skip and zero Release warnings/errors. Linux passes 333 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventory validates 441 cases across ten families. Root passes 281 focused checks and reproduces 3,071 fresh PDFs with twelve source Git blob identities. Local 0.1.5 package smoke passes; release preparation remains deferred. Met-threshold wrapped Calibri numbering and broader composed/continued flow remain separate qualifications.

## RV06-L175: met-threshold wrapped Calibri numbering

Runtime **456ced88**, integrated runtime root **087bf35e**, extends each-page automatic footnote numbering to complete regular 12pt Calibri wrapped frames with one uniform positive raw body threshold at most 24 half points. Existing pair advances remain unchanged. Plain single-section frames, zero numeric spacing, single automatic pitch, full fonts, unit scale, exact source segment ownership and complete one-line notes bound the retry. Only explicit Word-compatible numbering/stability calls enable this extension. Preserve, nominal L174 admission and legacy one-line kerning/minimum-leading/reserve gates retain their policies.

Thirty-six fresh independent Word controls cross early/later marks, run/paragraph page breaks, thresholds 1/22/24, absent/zero/unmet counterparts, 14pt and Arial/Times guards. Eight retained met/size controls and two reproducible public counterparts complete **46 controls /18 corrections /28 exact guards /36 strictly improving pages /72 body marks /72 footer labels /144 matching labels and pages /72 actual links**. All admitted labels/pages match Word. Actual Word breaks match across 432 lines and 29,838 visible characters. Remaining absolute glyph-start/span errors reach .320/.339pt, body baselines .193pt and unchanged marker baselines .207pt. Actual link rectangles are within .047pt and destinations .648pt. Body wrapping, all vertical geometry, complete notes/reserve, source font metrics, graphics, authored text/digits, source pairs/thresholds, physical font sizes and destinations retain parent identity. Full glyph/baseline parity is not claimed.

The public threshold-one early and threshold-limit later fixtures reproduce fourteen ZIP payloads twice through `tools/NewDocxMetWrappedCalibriFixtures.ps1`. Both retain the approximate-footnote diagnostic and remain partial at unchanged .42/.8/.8/.0055 gates, failing MAE and changed pixels. No independent human rating is supplied. Of 6,233 earlier outputs, 6,229 remain exact and only the four retained met controls change; all four receive independent Word qualification with unchanged page counts. A public glyph-map helper initially flattened its single PowerShell tuple, failing before producing maps. Its original form remains retained; only that helper resumed, preserving existing Word/candidate maps and consumed source pins.

Four new semantic methods and 67 earlier focused checks pass, covering uniform thresholds, nominal/malformed/mixed exclusions, unsupported formatting/notes, shifted/duplicate marker segments, source immutability, exact stable pages, explicit default/Preserve fallback and actual destination pages. Frozen source reproduces **6,279 PDF byte/page identities**, including **3,117 qualified visual outputs**. Windows passes 2,512 of 2,513 methods with one optional private skip and zero Release warnings/errors. Linux passes 337 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventory validates 443 cases across ten families. Root passes 285 focused checks and reproduces 3,117 fresh PDFs with nine source Git blob identities. Local 0.1.5 package smoke passes; release preparation remains deferred. Mixed-threshold advances and richer composed/continued flow require separate qualifications.

## RV06-L176: shared wrapped-footnote policy and mixed thresholds

Runtime **142716fd**, integrated runtime root **ad5caaff**, consolidates regular wrapped-footnote eligibility and raw-threshold classification in `DocxWrappedFootnoteFrame`. Explicit source/normalized policies replace the expanding Calibri boolean combinations. Existing nominal-only and uniform-met qualifications remain separate purposes; the mixed Word-compatible purpose derives baseline pairs per retained raw run threshold. The shared complete-frame and exact source-owner rules retain physical font sizes, authored digits/metadata, Preserve and the legacy one-line spacing/minimum-leading/reserve callers. Malformed, other-size, richer, scaled, kept, split/continued/custom or ambiguous flow retains fallback.

Thirty-six independent paragraph/run, early/later and run/paragraph-break controls plus twelve retained nominal/uniform-met/14pt guards and two public counterparts qualify **50 controls /34 corrections /16 exact guards /68 improving pages /136 body marks /136 footer labels /136 actual links**. Labels, pages and actual Word breaks match across **816 lines /56,372 visible characters**. Remaining glyph-start/span errors reach .320/.339pt, body baseline .193pt and unchanged marker baseline .207pt. Link bounds are within .048pt and destinations .648pt. All body verticals, complete notes/reserve, full-font metrics, graphics, authored text and actual destinations retain parent identity. Full glyph/baseline parity is unclaimed.

The two public fixtures reproduce fourteen ZIP payloads twice through `tools/NewDocxMixedWrappedCalibriFixtures.ps1`. Both retain the approximate-footnote diagnostic and remain partial at unchanged .42/.8/.8/.0055 image gates, failing MAE and changed pixels on both pages. No human rating is supplied. All **6,279** earlier PDFs retain fresh exact byte/page identity; frozen source reproduces **6,329** outputs, including **3,167** qualified visual outputs. Seventy-five distinct focused checks pass. Windows passes 2,522 of 2,523 methods with one optional private skip and zero Release warnings/errors. Linux passes 368 checks with six existing missing-Arial skips and two archive SourceLink warnings. Inventory validates 445 cases across ten families. Root passes 316 focused checks and reproduces 3,167 fresh PDFs with twelve matching source Git blobs. Local 0.1.5 package smoke passes; release preparation is deferred.

The preceding bounded maintenance commits `e4fa68a6` and `e2d7798f` remain covered by this corpus/platform qualification. One synchronized font cache owns immutable borrowed arrays and LRU accounting; local and HTTP sources retain loading coalescence and existing byte limits. Concurrent loading/read/eviction regressions pass; font corruption was not demonstrated. PDF validation memoizes shared vector Forms per page while checking full invocation depth, and the builder rejects stray restores before changing content. Essential raw-string resource, encoding and nesting checks remain in production. A warm synthetic shared-resource probe reduced validation from 16.13 to .28ms and allocations from 163,792 to 46,152 bytes; ordinary scans are essentially unchanged and real public samples each validate below .1ms. These microbenchmarks are not whole-conversion speed claims. Archive/XML/image bounds, cancellation, cumulative work limits, bounded staging and atomic publication remain intact; no heap accounting or new public quotas were added.

## RV06-L177: shared numeric after-spacing and final-line note links

Frozen runtime `22efc5ce`, integrated root `2180b5b9`, reuses the shared regular-12pt wrapped frame, raw-threshold purposes and exact complete-note mapper. The existing numeric spacing qualifier bounds after-spacing to 0..8pt; before-spacing remains fallback. Annotation spacing records the final visible line's identity, excluding break-spill rows. No per-font boolean, new quota or validation framework is introduced. Source digits/raw policies, physical font sizes, Preserve, earlier nominal/uniform purposes and legacy complete-note reserve boundaries remain intact.

Seventy-two new independent controls cross nominal, uniform-met and mixed Calibri policies, both explicit page-break forms, early/final markers and one/four/eight-point after-spacing, with Arial/Times counterparts and before/range fallbacks. Sixteen retained zero-spacing/size controls and two reproducible public cases bring the new total to 90: 38 corrections and 52 exact PDF guards. All 76 corrected raster pages improve MAE, similarity and recall. Four generated owners per control resolve labels/pages and actual navigation. One changed earlier single-line after-six-point output is separately qualified against its existing Word reference: annotation bounds improve within .026pt while raster, text geometry and destinations remain identical. The two new public cases still fail MAE and changed-pixel gates on both pages at unchanged .42/.8/.8/.0055 thresholds, with required `DOCX_APPROXIMATED_FOOTNOTE` diagnostics; they remain partial.

For all 39 corrected controls, actual Word visible line sequences match across 916 body lines and 63,162 characters. Maximum absolute glyph starts/spans are .324/.348pt and body baseline residual .170pt. The unchanged Times superscript marker residual reaches .700pt; complete glyph/baseline parity is not claimed. Link bounds/destinations are within .048/.799pt. Body vertical layout, complete note/reserve geometry, source fonts/metrics, graphics and authored digits retain parent identity. Numeric after-spacing changes only final-line hit areas and generated labels where needed.

A separate confirmed gap remains: Word suppresses before-spacing after run page breaks, while the parent retains it and shifts following page-two text by the gap. The complete before/both matrix distinguishes paragraph page breaks: before-four with preceding after-zero retains the gap, while before-eight with preceding after-eight suppresses it. Ordinary Arial 10pt and Calibri 12pt controls show the same distinction. A universal suppression rule is not established; adjacent spacing contributions require separate qualification. Before/both-spacing controls remain exact fallbacks in this slice; investigate the page-start semantic rule separately before admitting that scope. Split/continued/custom/composed, rich/kept/scaled or other-size frames and automatic/line-unit/contextual spacing remain excluded.

The frozen catalogue records 2527 registered methods, 2526 passes, no failures and one optional private skip; Windows Release builds have no warnings/errors. Linux records 88 PDF and 284 targeted passes, six existing missing-Arial skips and two archive SourceLink warnings. Four new tests cover shared threshold spacing, actual final-line annotation ownership in Word/Preserve, raw-spacing/frame fallbacks and duplicate owner rejection; 79 distinct focused methods pass. Initial fixture diagnostics are retained separately: break-only paragraph modification, fake-font wrapping assumptions and a reader-rejected malformed spacing token were corrected before qualification. The older mixed-purpose one-point exclusion now tests the unchanged nine-point boundary; earlier nominal/uniform exclusions remain intact.

All 6419 frozen PDF byte/page identities pass. Of 6329 earlier outputs, 6328 retain exact bytes and one annotation-only output receives the Word qualification above; no page counts change. Root validation records 320 focused checks, 3257 fresh visual PDF byte/page identities and thirteen matching source Git blobs. Public inventory is 447 cases across ten families: 118 locked, nine text-operation locked and 320 approximate. The generator reproduces 28 ZIP payload identities. Local 0.1.5 package smoke passes; release preparation remains deferred. The completed M1/M2/M3 maintenance stays covered, with archive/XML/image bounds, cancellation, cumulative work limits, bounded staging and atomic publication retained.

## RV06-L178: semantic page-start spacing and complete note owners

Frozen runtime `a5921397`, integrated root `8807c97d`, resolves unit Word page-start spacing from existing break provenance. Run page breaks consume zero before-gap; paragraph page breaks consume only `max(0, before minus preceding after)`. Authored before/after metadata is retained; pending after-spacing is zero at the new physical frame. One plain-body check per layout avoids rescanning the document at every break. Requested geometry is passed separately from effective frame geometry, so documents without balloons use Word spacing without changing page width. The SourceThresholds/NormalizedThresholds purposes share the existing numeric 0..8pt qualifier for both sides and one exact complete-owner mapper. Earlier nominal/uniform purposes retain zero-spacing guards. First visible lines own consumed before-gap; final visible lines own after-gap. No per-font boolean or new quota is introduced.

Twenty-four ordinary/unequal-spacing controls cross Arial 10pt, Calibri 12pt and Times 14pt, run/paragraph breaks, natural flow and document start. Unequal Calibri before-eight/after-four, before-four/after-eight and before-eight/after-zero controls distinguish the paragraph-break collapse rule. Fourteen corrections improve their page-two MAE, similarity and recall; ten guards retain exact PDFs. Page-one and other unaffected rasters, authored text, physical sizes and horizontal font geometry retain identity. Thirty-six note controls plus two reproducible public cases cross nominal/uniform-met/mixed Calibri, early/final markers, before-one/four/both-eight and Arial/Times counterparts. All 76 note raster pages strictly improve; 152 body markers, 152 footer labels and 152 actual links resolve. Complete note/reserve geometry, source fonts/metrics, graphics and authored digits retain parent identity apart from the qualified body/page-start changes and generated labels.

Actual Word visible line sequences match for all 38 note/public controls across 912 body lines and 63,004 characters. Glyph starts/spans are .324/.348pt, note body baseline .195pt and ordinary baseline .252pt. The physical Times marker residual reaches .740pt; full baseline/glyph parity is not claimed. Link bounds/destinations are within .048/.799pt. Both public controls still fail MAE and changed-pixel gates on both pages at unchanged .42/.8/.8/.0055 thresholds, with required `DOCX_APPROXIMATED_FOOTNOTE` diagnostics; both remain partial. Automatic/line-unit/contextual or out-of-range spacing, rich/kept/scaled/composed, split/continued/custom/ambiguous owners retain existing fallbacks. Source digits/raw policies, physical sizes, Preserve and legacy complete-note reserve boundaries remain intact.

The initial immutable body-only diagnostic retained the effective Preserve page frame without balloons and consequently missed requested Word spacing. A rendered first-line ownership regression exposed that integration issue. Passing the requested geometry enum corrected it; the final six new tests cover consumed gaps/raw metadata, Preserve/scaled profiles, numeric/kept fallbacks, natural/document starts, shared threshold numbering and actual before-gap annotation ownership. Initial diagnostics remain separate from the final source/DLL and consumed measurements. Eighty-five distinct focused methods pass.

Windows records 2533 registered methods, 2532 passes, no failures and one optional private skip; Release builds have no warnings/errors. Linux records 88 PDF and 290 targeted passes, six existing missing-Arial skips and two archive SourceLink warnings. All 6481 frozen PDF byte/page identities pass. Of 6419 earlier outputs, 6395 retain exact bytes and 24 changes are precisely the reused before/both note controls, each fully Word-qualified; no page counts change. Integrated root records 326 focused checks, 3319 fresh visual PDF identities and thirteen matching source Git blobs. Public inventory is 449 cases across ten families: 118 locked, nine text-operation locked and 322 approximate. The generator reproduces 28 ZIP payload identities. Local 0.1.5 package smoke passes; release preparation remains deferred. Completed M1/M2/M3 maintenance and archive/XML/image bounds, cancellation, cumulative limits, bounded staging and atomic publication remain covered.

## RV06-L179: section note defaults with document restart metadata

Runtime **2fe8b191**, integrated runtime root **6ce04f2d**, corrects section-default note numbering when authored document properties request eachPage. Word uses the section's defaults and occurrence policy in these supported section shapes; the document token alone does not request rendered each-page numbering. One shared reader admission applies to footnotes and endnotes. Omitted footnote format defaults to decimal and omitted endnote format to lower Roman. Section eachPage, unknown document tokens, malformed/unsupported section properties and unsafe document shapes retain their existing fallback. Authored document settings remain unchanged.

Twenty-four fresh Word references cover both note kinds, document decimal/start-four and decimalZero/start-nine, omitted/empty/partial/full section properties, and two-section defaults or explicit restart/continuous starts. Two retained document-default controls and two earlier section-occurrence controls complete **28 controls /23 corrections /5 exact guards /112 body markers /112 note labels /112 actual links**. Every source-owned label and body/note page matches Word. Authored text, physical sizes, baselines, graphics, destinations and link heights retain parent identity. Raster comparisons contain 34 improving pages, 15 exact pages and 11 slightly worse pages; maximum MAE increase is .0453 and changed-pixel ratio increase .000230. Inherited endnote destination drift reaches 8.891pt. Semantic label correctness does not establish complete visual parity; document-end endnote spacing is the next separate qualification.

Two new semantic methods and 49 retained focused checks pass. Windows records 2535 methods, 2534 passes, zero failures and one optional private skip; strict Release builds have zero warnings/errors. Linux records 88 PDF and 292 targeted passes, six existing missing-Arial skips and two archive SourceLink warnings. Frozen source checks 6505 PDF byte/page identities, including 3343 visual outputs. Of 6481 earlier outputs, 6477 remain exact and four change; all four are independently Word-qualified without page-count changes. Integrated root records 328 focused checks, 3343 fresh exact PDFs and two matching source Git blobs. Public inventory remains 449 cases across ten families. Local 0.1.5 package smoke passes and release preparation remains deferred. Ignored evidence preserves the source/DLL pins and corrects finalizer status, DOCX count convention and package-summary schema errors before acceptance. Completed bounded maintenance and existing resource/cancellation/publication safeguards remain covered.

## RV06-L180: complete plain document-endnote boundaries

Runtime **5bd25426**, integrated runtime root **7a78ca5c**, removes the generic three-point inter-note slicer gap and collapses paired numeric paragraph spacing for complete fitting plain document-end endnotes. The flow purpose reuses the shared plain-body frame, nominal endnote admission and existing paragraph spacing resolver. Raw spacing stays intact. Automatic one-line source paragraphs qualify; fixed-height, continued, scaled, kept/contextual, multi-paragraph and composed layouts retain established fallbacks.

Independent Word qualification covers **54 controls /42 corrections /12 exact guards /216 body markers /216 note labels /216 actual links**, including Arial, Calibri and Times New Roman at 10/12/14pt and unequal numeric gaps. There are 42 improved and 66 exact raster pages, with no other pages or page-count changes. Maximum note-baseline error drops from 33.073pt to .182pt. Authored text, glyph X and physical sizes, body baselines, non-clip graphics and annotation bounds retain identity. The 126 changed note clips preserve their shape and follow the exact source-owned note translation. Labels and source/destination pages match Word; destination top residual reaches 8.965pt and is a separate follow-up. Complete visual/navigation parity is not claimed.

Of 6505 earlier outputs, 6469 remain exact and 36 change. All 36 are independently requalified against cached Word: 92 source-owned note lines resolve, 56 improve, 36 first notes stay exact and none worsen; maximum error drops from 9.056pt to .056pt. Their 36 raster pages improve and 11 remain exact. Initial broad diagnostics exposed fixed-height and composed-header regressions before shared plain-frame admission preserved those fallbacks; their pins remain unaccepted. Three new semantic tests and 36 focused endnote checks pass. Windows has 2538 registered methods, 2537 passes, zero failures and one optional private skip; strict builds have zero warnings/errors. Linux has 88 PDF and 295 targeted passes, six existing missing-Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6559 PDF byte/page identities, including 3397 visual outputs; the integrated root passes 331 focused checks, 3397 fresh exact PDFs and three matching source blobs. Public inventory remains 449 cases across ten families. Local package 0.1.5 smoke passes; release preparation remains deferred.

## RV06-L181: source-owned endnote destination boundaries

Runtime **15fb0996**, integrated runtime root **184e4b6a**, carries each admitted complete plain endnote's collapsed boundary overlap from layout into link emission. Word targets later notes below the shared previous-after/current-before overlap. The renderer consumes the placement-owned offset before geometry mapping, avoiding a second eligibility policy. First-note targets, visible L180 placement and composed/fixed/continued/scaled fallbacks keep their behavior.

Six fresh mixed-spacing Word controls join 54 retained references: **60 controls /21 corrections /39 exact guards /240 source-owned destinations /57 changed targets**. No qualified target worsens and maximum target error falls from 8.965pt to .965pt. Remaining rounding differences mean complete navigation parity is not claimed. All 120 raster pages, glyphs, graphics, layout snapshots and annotation bounds stay exact; every PDF content/resource object and link field outside the qualified target Y stays byte-identical. All labels and source/destination pages match Word. Of 6559 earlier outputs, 6544 stay exact and only 15 L180 controls change, in 45 already qualified targets. Two new semantic emitted-link/fallback tests and 38 focused endnote checks pass.

Windows records 2540 registered methods, 2539 passes, zero failures and one optional private skip; strict builds have zero warnings/errors. Linux records 88 PDF and 297 targeted passes, six existing missing-Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6565 PDF byte/page identities, including 3403 visual outputs and fresh identities for every independently measured control. Integrated root records 333 focused passes, 3403 fresh exact PDFs and four matching source blobs. Inventory remains 449 cases across ten families. Local package 0.1.5 smoke passes; release preparation remains deferred. Complete wrapped endnote paragraphs are the next separate visible-layout qualification.

## RV06-L182: complete wrapped document-endnote flow

Runtime **36a5fa01**, integrated runtime root **cd45f62b**, extends the existing complete plain endnote policy to wrapped lines owned by one source paragraph. It resolves the preceding after-gap from the paragraph's final line and reuses the shared numeric boundary resolver. No font/size-specific branch, spacing threshold or quota is added. Fixed-height, kept, continued/scaled and composed paths retain their established fallbacks and first-placement navigation.

Independent Word qualification covers **21 controls /18 corrections /3 exact guards /144 source-owned note lines /84 body markers, note labels and actual links**. Arial, Calibri and Times New Roman at 10/12/14pt cover zero and paired numeric gaps; separate fixed-line, kept and continued controls remain exact. Eighteen note raster pages improve, 26 pages stay exact and none worsen. Maximum baseline error drops from 21.121pt to .186pt and destination residual reaches .875pt. Source paragraph text, glyph X/physical sizes, body baselines, non-clip graphics and annotation bounds retain identity. Fifty-four note clips preserve their shape and follow the source-owned translation. Four inherited Word line-break differences remain unchanged; ordered source-owned baselines and full paragraph text are qualified separately from complete wrapping parity. All 6565 prior PDFs stay byte-identical. Two new complete/continued semantic methods and 40 focused endnote checks pass; the prior wrapped-only unit fallback is superseded by this independent qualification.

Windows records 2542 registered methods, 2541 passes, zero failures and one optional private skip; strict builds have zero warnings/errors. Linux records 88 PDF and 299 targeted passes, six existing missing-Arial skips and two archive SourceLink warnings. Frozen source verifies 6586 PDF byte/page identities, including 3424 visual outputs. Integrated root records 335 focused passes, 3424 fresh exact PDFs and two matching source blobs. Inventory remains 449 cases across ten families. Local package 0.1.5 smoke passes and release preparation remains deferred. Ignored evidence retains the rejected early measurement and initial compile/helper diagnostics separately from qualified source/output pins.

## RV06-L183: source-owned complete-endnote kerning

Runtime **c625f9ab**, integrated root **98f9f824**, resolves authored thresholds for complete plain endnotes through a shared resolver. Nominal advances apply to disabled pairs or an export-grid shift; native met-pair metrics remain when sizes coincide. Layout-owned admission and one bounded retry preserve note-page ownership. Raw metadata, markers, body operations, graphics and navigation remain unchanged. Baseline-marked, malformed, mixed-note, continued/scaled and composed frames retain existing fallbacks.

The 58 Word controls qualify 47 corrections and 11 exact guards: all line breaks match, 47 raster pages improve, 71 remain exact and none regress. Maximum span error falls from 3.669 to .499pt. Of 6586 prior PDFs, 6471 remain exact; all 115 changed outputs retain page counts and independently match source-owned Word text/line breaks. Maximum span error falls from 46.536 to .494pt. Seventy line spans worsen by at most .131pt; two short Arial-14pt raster pages regress slightly (MAE +.003726 /changed-pixel ratio +.000052). These residuals remain recorded.

Windows: 2545 registered /2544 passes /zero failures /one optional skip; clean strict build. Linux: 88 PDF +302 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6620 PDF byte/page identities; root verifies 338 focused passes, 3458 fresh visual identities and eight source blobs. Local package 0.1.5 smoke passes; release preparation remains deferred. Rejected diagnostic prototypes and failed helper launches remain outside accepted evidence.

## RV06-L184: automatic note-story marker ownership

Runtime **7f1a5240**, integrated root **de451590**, retains the exact automatic marker role through ordinary and complex-field reading. Adjacent text keeps its source offsets; custom text does not acquire an automatic role. The complete-endnote correction protects these runs independently of alignment and reuses shared threshold/advance rules.

Twenty-nine Word controls qualify 18 corrections and 11 exact guards, including baseline/subscript and shared-source-run markers. All line breaks and 116 marker geometries/source-owned links qualify; body operations, graphics and navigation stay unchanged. Sixteen raster pages improve, 40 stay exact and two new shared-run pages retain small residual changes (Arial met MAE +.005164 /ratio +.000069; Calibri met MAE +.000497 /ratio -.000018). Automatic-marker span residual reaches .344pt; the unchanged custom baseline fallback reaches 2.042pt. All 6620 prior PDFs remain exact.

Windows: 2547 registered /2546 passes /zero failures /one optional skip; clean strict build. Linux: 88 PDF +304 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6649 PDF byte/page identities; root verifies 340 focused passes, 3487 fresh visual identities and five source blobs. Local package 0.1.5 smoke passes; release preparation remains deferred.

## RV06-L185: source-owned custom endnote prefixes

Runtime **fbf5c74c**, integrated root **51ce0b54**, protects the first custom character through a unique typed body reference/note ID and exact source run/offset. Adjacent text retains ownership and shares the existing threshold rule. Missing, ambiguous, overlapping and mismatched owners keep the original document. Automatic/superscript, continued/scaled and composed policies remain intact.

Thirty-one Word controls qualify 26 corrections and five exact guards: all line breaks match, 26 raster pages improve, 36 remain exact and none regress. Span error falls from 3.779 to .483pt. All 124 first-glyph geometries and source-owned links qualify; body operations, graphics and navigation stay exact. Of 6649 prior PDFs, 6648 remain exact; the one custom-baseline correction joins exact ZIP parts and fresh parent/candidate PDFs to independent Word qualification. Partial Word-operation marker advances are unmeasured and are not claimed by the geometry proof.

Windows: 2549 registered /2548 passes /zero failures /one optional skip; clean strict build. Linux: 88 PDF +306 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6679 PDF byte/page identities; root verifies 342 focused passes, 3517 fresh visual identities and two source blobs. Local package 0.1.5 smoke passes; release preparation remains deferred.

## RV06-L186: continued endnote typography

Runtime **40eac530**, integrated root **0e2e7416**, shares plain paragraph/frame admission and authored threshold/advance rules across fitting and continued endnote slices. A shared source-owner mapper guards each note's first page, last page and slice count for one bounded retry. Marker ownership, body operations, graphics, navigation and vertical flow remain unchanged; fixed/kept, malformed, scaled, mixed-note and composed frames retain explicit fallbacks.

Thirteen Word controls qualify eight corrections and five exact guards. All 126 source-owned line-break differences become zero across 719 lines and 44 notes; 17 raster pages improve, 21 stay exact and none worsen. Span error falls from 6.757 to .461pt. Of 6679 prior PDFs, 6678 stay byte-identical; one continued block independently qualifies through its exact input and fresh parent/candidate PDFs. Existing vertical spacing remains a separate defect, reaching 17.06pt baseline error in the older numeric-spacing control.

Windows: 2551 registered /2550 passes /zero failures /one optional skip; clean strict build. Linux: 88 PDF +308 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6689 PDF byte/page identities; root verifies 344 focused passes, 3527 fresh visual identities and six source blobs, with the former helper path removed. Local package 0.1.5 smoke passes; release preparation remains deferred.

## RV06-L187: continued endnote paragraph spacing

Runtime **f783e253**, integrated root **2b1c30b2**, gives paragraph-after spacing to the final source-owned line without rewriting its metadata. Adjacent notes reuse the numeric boundary resolver; only first slices own navigation offsets. Fixed/kept, mixed-note, scaled/composed and insufficient-head-space policies remain explicit. Typography, bounded retry, resource limits, staging and atomic publication are preserved.

Twenty-seven controls cover 24 fully qualified corrections, one improving partial case and two exact guards across 1906 lines/100 notes. Four inherited line-to-page differences disappear; all line breaks match. Thirty-six raster pages improve, 52 stay exact and none worsen. Qualified baseline/target residuals reach .182pt/.977pt. All 100 body/marker owners, actual links and changed clips resolve to source paragraphs and placements. Of 6689 prior PDFs, 6678 remain exact; all 11 corrections independently join fresh PDFs and exact inputs to Word controls. One several-long-note case retains three inherited page-line discrepancies and one first-target-page difference; matching-page baseline error improves 33.344 to 29.344pt.

Windows: 2553 registered /2552 passes /zero failures /one optional skip; clean strict build. Linux: 88 PDF +310 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6703 PDF byte/page identities; root verifies 346 focused passes, 3541 fresh visual identities and two source blobs. Local package 0.1.5 smoke passes; release preparation remains deferred.

## RV06-L188: shared body and endnote widow heads

Runtime **b186ec9d**, integrated root **99a9171f**, shares canonical widow-control predicates between body splitting and qualified plain endnotes. Default/enabled control moves a one-line head from an occupied page; disabled control and mandatory progress on empty one-line frames remain explicit. Source-boundary spacing survives the page turn. The direct keep-metadata projection avoids allocating effective-property records. Tail behavior retains its prior rule pending independent evidence.

Sixteen Word controls qualify three corrections and 13 exact guards across 3140 lines/56 notes. All nine inherited line-to-page differences disappear, with matching first-target pages and line breaks. Baseline/target residuals reach .182pt/.844pt. Nine raster pages improve, 90 remain exact and none worsen. All 56 source owners and actual links resolve, with 259 source-owned clip mappings and exact non-note graphics. Of 6703 prior PDFs, 6702 remain byte-identical; the previous partial case independently joins its exact input and fresh PDFs to Word qualification. Page counts remain unchanged. Sixteen fresh direct-property transfers remain exact.

Windows: 2555 registered /2554 passes /zero failures /one optional skip; strict build has zero warnings/errors. Linux: 88 PDF +312 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6715 PDF byte/page identities; root verifies 348 focused passes, 3553 fresh visual identities and four source blobs. Local package 0.1.5 smoke passes; release preparation remains deferred. Bounds, cancellation, cumulative work limits, staging, atomic publication and qualified fallbacks remain intact.

The initial integrated PDF group failed an existing cancellation spill-inventory assertion. Its cause remains unverified. An isolated cleanup check and a fresh full 94-check PDF group passed. The initial failure and successful retry reports remain separate; no production cleanup change was made. A retry evidence aggregator was corrected to exclude the visual output manifest before recording the 348 accepted checks.

## RV06-L189: source-owned endnote widow tails

Runtime **97a638b2**, integrated root **04f1d139**, reserves two final source lines using the shared body widow predicate. Canonical default/on/off metadata, mandatory one-line-frame progress and existing initial-head rules remain explicit. The source-owned final after-space, cumulative slice charges, navigation offsets, typography and composed/scaled/mixed-note fallbacks are preserved. One new semantic regression and 225 focused checks pass; no new font/size branch or quota was added.

Twenty-seven Word controls qualify eight corrections and 19 exact guards across 1329 lines/108 notes. Eight line-to-page differences disappear; 16 raster pages improve, 65 remain exact and none worsen. Qualified baseline/target residuals reach .146pt/.955pt. All 108 source owners, protected markers and actual links resolve, with 258 source-owned clip mappings and exact non-note graphics. Five byte-identical inherited partial controls remain: three wrap cases retain 132 comparison-line differences, five aligned page-line differences and three first-target page differences; overall target residual reaches 28.553pt. These limits are separate from the eight qualified corrections. All 6715 earlier PDFs remain byte-identical.

Windows: 2556 registered /2555 passes /zero failures /one optional skip; strict build has zero warnings/errors. Linux: 88 PDF +313 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6742 PDF byte/page identities; root verifies 349 focused passes, 3580 fresh visual identities and two source blobs. Local package 0.1.5 smoke passes; release preparation remains deferred. Bounds, cancellation, cumulative work limits, staging and atomic publication remain intact. An initial nullable-ID build error and a test scan starting above the real capacity boundary were corrected before qualification; rejected reports remain separate.

## RV06-L190: physical endnote capacity before widow admission

Runtime **a6638ae3**, integrated root **f86fbd4c**, forces one-line progress only on empty qualified frames. Occupied frames must fit a source line before placing it. Zero-fit takes use the existing continuation/source-boundary path independently of widow metadata; legacy callers keep their prior progress rule. One new capacity/source-charge/ownership/default/on/off regression and 226 focused checks pass, including one-line-frame progress. No font/size branch or quota was added.

Twenty-seven Word controls cover 14 fully qualified corrections, one partial tradeoff and 12 exact guards across 1309 lines/108 notes. All 15 aligned page-line differences disappear and first-target pages match; qualified baseline/target residuals reach .130pt/.955pt. All 108 source owners, protected markers and actual links resolve, with 225 source-owned clip mappings and exact non-note graphics. Twenty-nine raster pages improve, 51 remain exact and one worsens. The partial Arial input still wraps into 43 candidate versus 44 Word lines (44 comparison-line differences). Its source baseline error falls 27.640 to 13.842pt, but page-3 raster MAE rises 2.138 to 3.274 and threshold-32 changed ratio .017074 to .025578; target residual remains 14.754pt. This tradeoff is separate from the 14 fully qualified corrections. Of 6742 earlier PDFs, 6739 remain exact; all three changes join exact inputs and fresh PDFs to two full corrections and this explicit partial tradeoff. Page counts remain unchanged.

Windows: 2557 registered /2556 passes /zero failures /one optional skip; strict build has zero warnings/errors. Linux: 88 PDF +314 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6766 PDF byte/page identities; root verifies 350 focused passes, 3604 fresh visual identities and two source blobs. Local package 0.1.5 smoke passes; release preparation remains deferred. Bounds, cancellation, cumulative work limits, staging, atomic publication and qualified fallbacks remain intact. The initial 26-case report and rejected zero-raster-regression assumption remain separate.

## RV06-L191: source-owned typography continuation stability

Runtime **4dc7f1f4**, integrated root **4e169e0a**, validates source ownership and first placements before admitting the existing single typography retry. Continuation ranges may change when corrected advances change wrapping; canonical range checks, qualified-slice admission and total-page stability remain explicit. Zero already disables kerning; no font/size/threshold exception was added. One new real-reflow regression and 227 focused checks pass. Initial synthetic widths did not cross their repeated-word boundary; corrected fixtures exercise actual continuation reflow with stable first owners and total pages. Rejected reports remain separate.

Twenty-six Word controls qualify four corrections and 22 exact guards across 1274 lines/104 notes. All 176 inherited comparison-line differences and three aligned page-line differences disappear. Eight raster pages improve, 70 remain exact and none worsen. Baseline/target/span residuals reach .146pt/.955pt/.416pt. All 104 source owners, protected markers and actual links resolve, with 236 source-owned clip mappings and exact non-note graphics. The three inherited Arial wrap cases, including L190's partial raster tradeoff, now match Word wrapping and page flow. Of 6766 prior PDFs, 6763 remain byte-identical; all three corrections independently join exact inputs and fresh PDFs to Word qualification. Page counts remain unchanged.

Windows: 2558 registered /2557 passes /zero failures /one optional skip; strict build has zero warnings/errors. Linux: 88 PDF +315 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6784 PDF byte/page identities; root verifies 351 focused passes, 3622 fresh visual identities and three source blobs. Local package 0.1.5 smoke passes; release preparation remains deferred. Bounds, cancellation, cumulative work limits, staging, atomic publication and qualified fallbacks remain intact. First-placement and total-page changes still keep the explicit retry fallback pending their own evidence.

## RV06-L192: qualified note-only page growth

Runtime **7596096e**, integrated root **83745467**, removes the duplicate total-page restriction from the existing qualified plain-endnote typography retry. Plain admission already excludes body fields, headers and footers; complete source ownership and stable first placements remain mandatory. Corrected wrapping may add note-only pages. Other retry paths retain their page-count checks. One regression verifies unchanged body geometry, note-only staging and exactly-once source line ownership; 228 focused note checks pass. No font/size/threshold exception, retry or quota was added.

Nine independent Word controls qualify one correction and eight byte-identical guards across 1560 lines/36 notes. The correction restores a sixth page and matches Word's line/page flow; all 44 comparison-line and seven aligned page-line differences disappear. Four raster pages improve, 47 stay exact and the missing page is restored. Baseline/target/span residuals reach .166pt/.766pt/.495pt. All 36 body owners, protected markers and actual links resolve, with 132 note clip mappings. Existing non-note graphics remain exact; the two added separator operations map to the new source-owned continuation separator. All 6784 prior PDFs remain byte-identical. An initially overstrict geometry proof was corrected to account for the source-owned added page; original diagnostic pins and reports remain separate.

Windows: 2559 registered /2558 passes /zero failures /one optional skip; strict build has zero warnings/errors. Linux: 88 PDF +316 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6793 PDF byte/page identities; root verifies 352 focused passes, 3631 fresh visual identities and three source blobs. Local package 0.1.5 smoke passes; release preparation remains deferred. Bounds, cancellation, cumulative work limits, staging, atomic publication and composed/scaled/mixed-note fallbacks remain intact. First-placement changes retain the explicit retry fallback pending separate qualification.

## RV06-L193: plain-note retry by complete source ownership

Runtime **c618c366**, integrated root **a9d46131**, replaces first-page coordinate equality with complete unique source-owner validation on the existing qualified plain-endnote retry. Corrected advances can move first placements and continuations while authored body references and marker/numbering policy remain immutable. Canonical original range metadata and qualified-slice checks remain explicit. A bounded 456-layout scan found two rejected first-placement changes; independent Word controls confirm both as wrapping/navigation defects. One regression verifies real first-page reflow, exact body geometry, immutable source markers, ordered exactly-once line ownership and substituted/malformed-owner rejection; 229 focused checks pass. No font/size/threshold branch, framework, retry or quota was added.

Twelve Word controls qualify two corrections and ten byte-identical guards across 3744 lines/48 notes. All 178 comparison-line and 17 aligned page-line differences disappear, and both shifted final-note destinations match Word. Seventeen raster pages improve, 88 stay exact and none worsen; page totals stay unchanged. Baseline/target/span residuals reach .166pt/.984pt/.495pt. All 48 body owners, protected markers and actual links resolve, with 252 source-owned clip mappings and exact non-note graphics. All 6793 prior PDFs remain byte-identical. The initial scaled diagnostic was excluded; a nullability build error and an assertion comparing translated records were corrected before qualification. Those reports remain separate.

Windows: 2560 registered /2559 passes /zero failures /one optional skip; strict build has zero warnings/errors. Linux: 88 PDF +317 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6805 PDF byte/page identities; root verifies 353 focused passes, 3643 fresh visual identities and three source blobs. Local package 0.1.5 smoke passes; release preparation remains deferred. Bounds, cancellation, cumulative work limits, staging, atomic publication and composed/scaled/mixed-note fallbacks remain intact. Fitting headers/footers need independent frame qualification and retain fallback pending the next assessment.

## RV06-L194: plain endnotes with fitting static stories

Runtime **70276108**, integrated root **d9c27610**, shares plain-body and automatic-line semantics while extending document-endnote admission to plain headers/footers that fit outside the body frame in every reachable first/even/default variant. The same static edges still own body displacement. Fixed-height, dynamic-field, rich/floating, scaled, mixed-note and overflow paths retain explicit fallbacks. Authored references/static stories and the single typography retry remain intact. One regression checks physical static geometry, immutable source ownership, alternating overflow and mixed-note boundaries; 230 focused checks pass. No font/size/threshold branch, retry or quota was added.

Fourteen independent Word controls qualify nine fitting-frame corrections and five byte-identical excluded controls across 1536 lines/36 notes. All 132 comparison-line and 57 aligned page-line differences disappear in admitted cases. Three missing pages are restored; 36 raster pages improve, 37 stay exact and none worsen. Baseline/target/span residuals reach .166pt/.766pt/.495pt. All 36 source body/marker/link owners resolve with 199 clip mappings. Existing 297 static-text operations retain exact geometry, four new operations map to added-page XML owners, and six added continuation-separator operations map to source layout. Existing non-note graphics remain exact. Three excluded overflow controls still have fewer pages than Word; exact fallback is not complete Office parity.

Source-backed static paragraph/page mapping separates repeated header/footer operations from note joins; unfiltered diagnostics remain archived. An initial fixed-height admission regression and a static-record equality assertion were corrected. Source review found possible mixed-note boundary changes outside scope; a final guard and regression preserve that path. Fourteen final PDFs transfer Word qualification exactly. The earlier prototype audit is diagnostic; the final frozen audit uses original accepted-parent expectations and proves all 6805 prior PDFs exact through the guarded source.

Windows: 2561 registered /2560 passes /zero failures /one optional skip; strict build has zero warnings/errors. Linux: 88 PDF +318 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6819 PDF byte/page checks; root verifies 354 focused passes, 3657 fresh visual identities and six source blobs. Local package 0.1.5 smoke passes; release preparation is deferred. Bounds, cancellation, cumulative work limits, page staging and atomic publication remain intact. The next bounded assessment addresses measured endnote/footer intrusion in excluded overflow controls.

## RV06-L195: plain endnotes bounded by selected static edges

Runtime **81255568**, integrated root **c6c68f4a**, fixes confirmed note/footer intrusion: the slicer previously consumed to the bottom margin while body flow used the overflowing footer's content edge. Header/footer continuation layout now shares the existing static machinery. One plain-static eligibility predicate serves frame admission and footer bounds. The selected footer edge owns separator capacity, paragraph boundaries and note takes; a usable frame replaces the blanket plain-static overflow exclusion. Consumed frames, fixed-height/dynamic/rich/floating/scaled and mixed-note paths retain explicit fallbacks. One new regression verifies selected alternating footer edges, clear clips, ordered exactly-once source lines, immutable body geometry and fitting-frame identity. It fails against the accepted parent DLL on footer intrusion and passes with the correction; 231 focused checks pass. No font/size/threshold branch, retry or quota was added.

Fourteen reused independent Word controls qualify five overflow corrections and nine byte-identical fitting controls across 2356 lines/56 notes. All 187 aligned page-line mismatches disappear and all labels/body owners/actual destinations resolve; maximum baseline/target/span residuals reach .166pt/.828pt/.495pt. Three missing pages are restored. Twenty-three raster pages improve, 53 remain exact and none worsen. All 197 clips map to source placements; existing non-note graphics and 301 static-text operations retain exact geometry. Twenty-five added static operations map to selected XML paragraph/page owners, with six source-owned continuation-separator operations. These are bounded qualifications, not complete Office parity.

The footer-floor-only diagnostic restored page totals but retained legacy 3pt spacing and an alternating paragraph-head mismatch, worsening some raster pages. Reusing the existing plain-note flow in usable static frames removes those defects. Nullable-flow metadata and an overly strict test assumption about page growth were corrected. The prior 6819-output audit changes only the five independently qualified controls; 6814 PDFs stay exact and three page totals change. Two final comment updates transfer through fresh exact frozen PDFs.

Windows: 2562 registered /2561 passes /zero failures /one optional skip; strict build has zero warnings/errors. Linux: 88 PDF +319 targeted passes, six existing Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6819 PDF byte/page checks without adding duplicate jobs for reused inputs. Root verifies 355 focused passes, 3657 fresh visual identities and six source blobs. Local package 0.1.5 smoke passes; release preparation is deferred. Bounds, cancellation, cumulative work limits, staging, atomic publication and mandatory empty-frame progress remain intact. Continued-footnote footer frames require a separate assessment.

## RV06-L196: plain footnotes share selected footer frames

Runtime **b0a48388**, integrated root **7a5aab82**, corrects confirmed footnote/footer intrusion. The selected footer edge now serves body reserves, head seating, whole/split slices, continuation separators and final drain. Plain static/note eligibility and usable-frame checks share existing semantic rules; frame admission is memoized per layout. Scratch placement uses actual document page numbers, authored page breaks bound continuation sharing, and continued tails/fresh heads share bottom-up seating. Consumed/scaled/mixed/rich/dynamic/fixed and footer-free paths retain explicit fallbacks. No font/size/threshold exception, retry or quota was added.

One new regression verifies ordered exactly-once source lines, source immutability, alternating footer glyph/rule clearance through final drain, and fitting/consumed fallbacks. It fails against accepted L195 on actual footer intrusion; 232 focused checks pass. Twelve independently rendered Word controls qualify six changed outputs and six exact controls across 1254 lines/30 note/navigation owners, 88 source clips and 44 source rules. Five corrections match page/line/body geometry within .125pt baseline error; one remains partial. All graphics map to source placements and every control clears footer text.

Seventeen raster pages improve, seventeen stay exact and four are restored. Two pages retain mixed raster metrics; no threshold was relaxed. The alternating multi-reference control retains one reference/page and one page-line difference. Fitting Calibri/Times retain inherited pagination differences and two candidate pages absent from Word's shorter output; Arial retains one 43/44-line wrapping difference. Corrected totals and exact controls do not establish complete Office parity.

Windows: 2563 registered /2562 pass /zero failures /one optional private skip; strict build has zero warnings/errors. Linux: 88 PDF +320 targeted passes, six existing missing-Arial skips and two archive SourceLink warnings. Frozen qualification verifies 6831 PDF byte/page checks: all 6819 earlier PDFs are unchanged, and twelve new comparison outputs transfer exactly. Root verifies 356 focused passes, 3669 fresh PDF identities and four source blobs. Local 0.1.5 package smoke passes; release preparation is deferred. Bounds, cancellation, cumulative work, page staging, atomic publication and bounded progress remain intact. Shared footnote reserve/deferred-head lifecycle remains a separate bounded assessment.

## RV06-L197: shared footnote reserve and deferred source heads

Runtime **f02d53d9**, integrated root **41a733c0**, corrects references and note heads being forced onto the same page in plain footer frames. One reserve owns the batch separator; rendered references hand deferred heads to the bounded drain. Seating and slicing share the body's default/on/off paragraph-fragment policy. Continued tails retain their measured rule-lap tolerance; fresh source heads leave the full ride clear. Existing frame admission/fallbacks remain intact, with no font/size/threshold exception or quota.

233 focused checks pass. New coverage exercises explicit-widow-off one-line heads, default/on whole-head deferral and ordered exactly-once final drain; fitting-batch identity uses a complete fitting batch. Accepted L196 fails the bounded widow-off control search. Twelve independent Word controls establish the five actual deferred heads and qualify three note-geometry corrections/nine exact outputs across 1254 source lines, 30 note/navigation owners, 88 clips and 44 rules. Corrected page counts, reference/split/destination pages match, with at most .133pt note baseline residual; graphics map to source placements.

Seven raster pages improve, 29 stay exact, two worsen and two have mixed metrics; largest MAE increase .080624. No threshold was relaxed. Five extra short rules on shared continuation/head pages and the first overflowing rule ride remain separate graphics gaps. Arial retains its final-note 43/44-line wrapping and destination-height difference. This does not establish complete Office parity.

Windows: 2563 passes/0 failures/one optional private skip, zero build warnings/errors. Linux: 88 PDF+321 targeted passes, six missing-Arial skips/two archive SourceLink warnings. All 6831 frozen comparisons pass: 6828 prior PDFs stay exact and three changes transfer Word-qualified evidence. Root: 357 focused checks/3669 fresh PDF identities/three matching source blobs. Local 0.1.5 package smoke passes; release preparation remains deferred.

## RV06-L198: one separator for a shared footnote compartment

Runtime **4efac6fa**, integrated root **b18502c0**, removes a redundant short head rule when a continued tail and fresh head share a plain footer-frame compartment. The continued first note supplies the full-width rule; missing/empty continuation separators and legacy frames retain their fallback. Shared reserve, eligibility, note layout, bounds and quotas remain unchanged.

233 focused note checks pass; the strengthened one-compartment regression fails against accepted L197. Twelve independent Word controls qualify three separator-only changes/nine exact PDFs. Exactly five redundant rules, five clips and five blank marks disappear. All normal-note/body/static/continuation layouts, nonblank PDF text operators, meaningful links and other graphics stay exact. 1254 source lines/30 note/navigation owners, 83 clips and 39 rules resolve; rule counts per page match Word. Five raster pages improve, 35 stay exact and none worsen. No threshold was relaxed. First overflowing rule placement and L197's fitting-body/Arial wrapping limits remain explicit.

Scoped Windows validation: 233 note+94 PDF passes, zero failures/build warnings/errors and exact catalogue identity. L197 remains the last full Windows catalogue (2563 passes/one optional private skip). Linux: 88 PDF+two targeted passes/six existing font skips/two archive SourceLink warnings. All 6831 frozen comparisons pass: 6828 prior identities stay exact; three changes transfer Word-qualified evidence. Root: 357 focused checks/3669 fresh PDFs/two source blobs. Local 0.1.5 package smoke passes; release preparation remains deferred.

## RV06-L199: one resolved height for footnote head seating and emission

Runtime **0391cbe7**, integrated root **d5947c17**, reuses the head separator height already resolved for seating through the existing emission override. The previous emission helper recomputed its legacy height. No eligibility branch, font/size exception or quota was added. 233 note checks pass; the strengthened whole/split physical-gap regression fails against accepted L198.

Twelve independent Word controls qualify six head-height-only changes/six exact PDFs. All eight incorrect gaps among twenty source heads are corrected; maximum physical gap residual .173pt. Eight source clips shrink and eight rules/blank marks move by the resolved eight-point height difference. Normal-note/body/static/continuation layout, other text/graphics and meaningful links stay exact across 1254 note lines/30 note/navigation owners/83 clips/39 rules. Eight raster pages improve, 32 stay exact and none worsen. No threshold was relaxed. L197's fitting-body raster tradeoffs and Arial wrapping/destination-height difference remain explicit.

Scoped Windows: 233 note+94 PDF passes/zero failures/build warnings/errors and exact catalogue identity. L197 remains the last full Windows catalogue (2563 passes/one optional private skip). Linux: 88 PDF+two targeted passes/six existing font skips/two archive SourceLink warnings. All 6831 frozen comparisons pass: 6825 prior identities stay exact and six changes transfer Word-qualified evidence. Root: 357 focused checks/3669 fresh PDFs/two source blobs. Local 0.1.5 package smoke passes; release preparation remains deferred.

## RV06-L200: shared source-owned note-content kerning

Runtime **c2a94e2b**, integrated root **c120e831**, reuses the note-content threshold transform through layout-owned plain footnote admission and complete canonical source-line coverage. Automatic markers/body owners remain immutable; existing endnote typography and explicit custom/malformed/legacy/incomplete fallbacks remain. Two new regressions pass. No font/size/document-case exception was added.

Twelve retained and 27 independent Word 16 controls cover three fonts/three sizes/absent, zero, met and unmet thresholds. All 57 note owners now have Word line breaks: 110 source-backed break mismatches become zero. Eighteen changed controls qualify note geometry; eighteen retain partial mixed-size geometry; three met-12pt PDFs stay exact. Body/static layout and 508 protected operations stay exact; 2416 note lines/155 clips/75 rules/57 navigation owners resolve. Arial's 43/44-line gap is corrected and destination residual is .148pt.

Raster pages: 69 improve/four exact/three tradeoffs. Calibri 10pt/met MAE improves .044884 with threshold ratio +.000100061. Arial 14pt/zero and unmet each worsen page-one MAE .272019/ratio .001067662, while page two improves MAE 4.778220/ratio .036037974. No threshold was relaxed. Mixed-size geometry remains partial: 10pt text under a 12pt marker has up to 80.737pt baseline drift; nine 14pt controls retain one continuation-line page mismatch. L197 fitting-body differences remain; full Office parity is not claimed.

Full Windows: 2565 passes/zero failures/one optional private skip and clean strict builds. Linux: 317 passes/eleven font skips/one inherited inline-image assertion failure; accepted L199 reproduces the same failure and all original note outcomes from its hash-pinned library. Both new methods pass; two archive SourceLink warnings remain. All 6858 frozen comparisons pass: 6819 prior exact identities, twelve Word-qualified changes and 27 new controls. Root: 359 focused checks/3696 fresh PDFs/five current source blobs and removed predecessor. Local 0.1.5 package smoke passes; release preparation remains deferred.

## RV06-L201: source-owned metrics for unmarked note continuation lines

Runtime **f6dfc2d7**, integrated root **6a004298**, preserves the marked first line's metric box and applies the existing content-profile/inset rules to actual wrapped lines without that source marker. Baseline resolution is shared; paragraph/marker/emission owners remain immutable. Single-line, subordinate, authored script, baseline, scaled, rich and fixed/minimum-line fallbacks remain. Four regressions pass; the content-pitch check fails against hash-pinned accepted L200 and the three retained guards pass there. Broad marker exclusions were rejected because they regressed a qualified mixed inset and the held-out one-line note.

Forty retained Word controls qualify nine corrected wrapped 10pt notes/31 exact PDFs. Maximum baseline drift falls from 80.737pt to .196pt; nine raster pages improve. The main 39 controls resolve 2416 note lines/57 note and navigation owners/155 clips/75 rules. Body/static layout and 508 protected operations stay exact; all source line breaks/page owners remain unchanged. The held-out one-line PDF is byte-identical. Nine 14pt controls retain one continuation-line page mismatch and up to 16.965pt baseline error; L200/L197 raster differences remain outside this correction. No threshold was relaxed and full Office parity is not claimed.

Full Windows: 2569 passes/zero failures/one optional private skip and clean strict builds. Linux: 321 passes/eleven font skips/one unchanged inherited inline-image assertion failure; all original L200 outcomes are retained and all four new methods pass. Two archive SourceLink warnings remain. All 6858 frozen comparisons pass: 6849 retained identities and nine Word-qualified changes. Root: 363 focused checks/3696 fresh PDFs/two source blobs. Local 0.1.5 package smoke passes; release preparation remains deferred.

## RV06-L202: shared separator flow charge for fresh footnote heads

Runtime **0bcfee39**, integrated root **6e11775f**, uses the admitted plain separator flow charge when fitting a fresh head, while retaining the separator's physical height/rule spacing. The capacity-cliff regression fails accepted L201 and passes current; all 240 note checks pass. No font/size/document-case exception or new quota was added.

Forty retained Word controls qualify nine corrected 14pt splits/31 exact PDFs. All 18 raster pages improve; maximum baseline error falls from 16.965pt to .319pt. Calibri retains .319pt and fails the unchanged .2pt geometry gate, while six Arial/Times controls pass it. Source line breaks/page ownership match Word. The main 39 controls resolve 2416 note lines/57 note and navigation owners/155 clips/75 rules; body/static layout and 508 protected operations stay exact. The held-out one-line PDF remains byte-identical. Earlier L200/L197 raster differences remain outside this correction; full Office parity is not claimed.

Full Windows: 2570 passes/zero failures/one optional private skip, clean strict builds. Linux: 322 passes/eleven font skips/one unchanged inherited assertion failure; all original L201 outcomes remain and the new method passes. Emitted-content inspection identifies that assertion's glyph-matrix ordinal assumption for a separate test-only correction. All 6858 frozen comparisons pass: 6849 retained identities and nine Word-qualified changes. Root: 364 focused checks/3696 fresh PDFs/two source blobs. Local 0.1.5 package smoke passes; release preparation remains deferred.

## RV06-L203: inline-image assertions across run and glyph emission

Test-only **b9114181**, integrated root **ca29f346**, shares line geometry across five inline-image checks and exercises default plus forced fallback glyph emission. The accepted Linux output already places following text at the 72pt image end; fixed matrix ordinals selected a preceding glyph instead. Five inherited image-group failures now pass. The separate absolute Word-coordinate calibration retains its coordinate/tolerance and explicitly requires Calibri.

Windows: 36 image/240 note checks, 273 distinct passes/zero failures/zero skips. Linux: 35 image/235 note passes, 267 distinct passes/zero failures/six font skips. Catalogue unchanged; all other note outcomes are retained. Strict builds are clean; two Linux archive SourceLink warnings remain. Seven frozen/root image PDFs are byte/page-identical. One test blob matches and all production/fixture trees remain unchanged. Full runtime/corpus/package evidence remains explicitly at L202 and was not rerun or relabeled for this test-only fix. Rendering limits, including the .319pt Calibri note residual, remain. Keep 0.1.5; release preparation deferred.

## RV06-L204: source-owned minimum note metrics and footnote flow

Runtime **679cafd3**, integrated root **5c4ab769**, shares automatic/minimum content-driven metric ownership. Actual unmarked lines use content profiles, authored minima remain bounds, and a binding minimum grows the line box above the same bottom bearing. Minimum footnote frames share existing kerning/source ownership. Automatic-only document-end layout and exact/scaled/rich/custom/subordinate/single-line metric fallbacks remain; a prototype broadening endnote frame admission was rejected. Both new regressions fail accepted L203 and pass current; all 242 note checks pass.

Eighteen independent Word controls qualify fifteen corrections/three exact guards. Nine footnote and four endnote controls pass the unchanged .2pt geometry gate; maximum corrected footnote residual is .196pt. Two Calibri minimum endnotes retain 66 source-break mismatches, and the exact Calibri guard retains 33. All 19 changed raster pages improve. Source mapping resolves 636 lines/18 note and navigation owners/44 clips/22 rules. Common body/static layout and 94 protected operations stay exact; one redundant empty Calibri minimum-18 continuation page disappears. Earlier .319pt Calibri 14pt residuals and other fidelity gaps remain; full Office parity is not claimed.

Full Windows: 2572 passes/zero failures/one optional private skip, clean strict builds. Linux: 325 passes/zero failures/eleven font skips, both new methods pass and prior L203 note outcomes remain; two archive SourceLink warnings. All 6883 frozen comparisons pass: 6858 earlier runtime PDFs and seven image guards exact/eighteen new Word controls qualified. Root: 366 focused checks/3721 fresh PDFs/four source blobs. Local 0.1.5 package smoke passes. Bounds, cancellation, cumulative work, bounded staging and atomic publication remain. Release preparation is deferred.

## RV06-L205: source-owned minimum endnote typography

Runtime **2803b923**, integrated root **3e686d0d**, separates content advances from automatic-only paragraph spacing/widow flow. Minimum endnotes retain their layout fallback and share existing threshold/nominal-advance policy. One canonical-line mapper now covers both note kinds, retaining endnote offset guards. Artificial incomplete/substituted layouts fail the parent ownership test; actual source loss was not demonstrated. Both new regressions fail accepted L204 and pass current; all 244 note checks pass.

Twenty-four retained Word controls qualify nine changes/fifteen exact guards. Six wrapped minimum endnotes pass the unchanged .2pt geometry gate (maximum .173pt), removing all 66 Calibri break mismatches. Three older one-line minimum controls change glyph spacing while placement/graphics/navigation stay exact; their 9.329..10.219pt baseline errors remain. Raster: ten improved/23 exact/one Times New Roman tradeoff (+.001557781 MAE, +.000020631 changed ratio). Source mapping resolves 660 lines/42 note and navigation owners/74 clips/28 rules. All body/static layout and 190 protected operations stay exact; page counts are unchanged. Exact-line Calibri's 33 break mismatches and the .319pt Calibri 14pt residual remain. Full Office parity is not claimed.

Full Windows: 2574 passes/zero failures/one optional private skip; strict builds clean. Linux: 327 passes/zero failures/eleven font skips, both new methods pass/prior outcomes retained; two archive SourceLink warnings. Frozen: 6883 qualified comparisons, 6874 prior exact/nine documented changes. The initial audit's three unexpected identities were separately qualified against Word and freshly reconverted before acceptance; its raw report is retained. Root: 368 focused checks/3721 fresh PDFs/five source blobs. Local 0.1.5 package smoke passes. Bounds, cancellation, cumulative work, bounded staging and atomic publication remain. Release preparation is deferred.

## RV06-L206: shared minimum note bearings and paragraph flow

Runtime **452d1c0f**, integrated root **668a3cff**, extends source-owned minimum box growth to fitting/uniform note content and shares automatic/minimum plain paragraph flow. A baseline-only prototype fixed first lines but exposed an accumulating 3pt legacy note gap. Both corrections share semantic admission. Original story scale is retained before design normalization: an expanded regression caught and rejected a scaled prototype change. Existing unmarked profiles and rich/custom/exact/source-shape fallbacks remain; no font/size exception was added.

Thirty Word controls qualify nine corrections/21 exact guards. Three retained zero-spacing and six independently rendered before/after controls meet the unchanged .2pt geometry gate; maximum baseline .179pt/destination .055pt. All nine changed raster pages improve/37 remain exact. Source mapping resolves 684 lines/66 note and navigation owners/104 clips/34 rules. Thirty body/static layouts and 286 protected operations remain exact; page counts are unchanged. Exact-24's 7.769..7.992pt baseline residuals, Calibri exact-12's 33 breaks and the .319pt Calibri 14pt residual remain. Full Office parity is not claimed.

All 246 note checks pass. Bearing and updated minimum-flow tests fail accepted L205; the expanded fallback guard passes parent/current. Full Windows: 2576 passes/zero failures/one optional private skip, clean strict builds. Linux: 329 passes/zero failures/eleven font skips, prior outcomes retained with one justified method rename; two archive SourceLink warnings. Frozen: 6889 comparisons, 6880 prior exact/three retained corrections/six new controls. Root: 370 focused checks/3727 fresh PDFs/five source blobs. Local 0.1.5 package smoke passes. Bounds, cancellation, cumulative work, bounded staging and atomic publication remain. Release preparation is deferred.

## RV06-L207: shared fixed-note baseline and paragraph flow

Runtime **50480903**, root **93f7db2a**, shares the body/table fixed-line inset and existing semantic note-shape guard. Fixed clipping remains separate from natural metrics; plain exact notes share paragraph flow and typography. Contextual/keep/malformed/rich/authored/scaled fallbacks remain. The broad baseline prototype worsened a contextual footnote raster and was rejected; both heldouts are restored exactly. No font/size branch was added.

Forty-six Word controls qualify 22 corrections/24 exact guards. The unchanged .2pt baseline gate holds (maximum .18pt); 24 raster pages improve/47 stay exact, no tradeoffs. Source mapping resolves 1034 lines/103 note/navigation owners/161 clips/52 rules. Body/static layouts and 445 protected operations stay exact; one empty continuation page is removed. XML-authored fixed boxes/pitches/story heights hold across 52 notes/468 lines. Wrapped endnote Word destination residuals up to .703pt and the earlier .319pt Calibri 14pt residual remain; full Office parity is not claimed.

All 249 note checks pass; two new regressions fail L206 and the expanded fallback guard passes both. Windows: 2579 passes/zero failures/one optional private skip, strict builds clean. Linux: 332 passes/zero failures/eleven font skips, prior outcomes retained and three new passes; two archive SourceLink warnings. Frozen: 6904 comparisons, 6882 prior exact/seven retained corrections/15 new controls. Root: 373 focused checks/3742 fresh PDFs/three source blobs. Local 0.1.5 package smoke passes. Bounds, cancellation, cumulative work, bounded staging and atomic publication remain; release preparation is deferred.

## RV06-L208: shared header/footer footnote frames

Runtime **dc733976**, root **a91064d5**, shares existing static-frame admission between plain headers and footers, including first/even/default selection. Source shape, keep, scale, overflow and canonical line ownership remain guarded. No font/size branch or unsafe API was added; the permanent safe-managed-code project flag and AGENTS.md rule remain intact.

Twelve independent Word controls qualify twelve corrections within the unchanged .2pt baseline gate (maximum .196pt); all sixteen raster pages improve. The 132 Calibri line-break mismatches disappear and two redundant empty continuation pages are removed. Source mapping resolves 424 lines/12 note/navigation owners/32 clips/16 rules. Body/static output and 64 protected operations remain exact. Navigation owners/pages hold; Word destination residuals up to .48pt and the earlier .319pt Calibri 14pt residual remain. Full Office parity is not claimed.

All 251 note checks pass; the new complete-owner regression fails the accepted parent and the rich/keep/contextual/scaled guard passes both. Windows: 2581 passes/zero failures/one optional private skip, strict builds clean. Linux: 334 passes/zero failures/eleven font skips, prior outcomes retained and two new passes; two archive warnings. Frozen: 6916 comparisons, all 6904 prior exact/twelve new controls. Root: 375 focused checks/3754 fresh PDFs/four source blobs. Local 0.1.5 package smoke passes. Bounds, cancellation, cumulative work, bounded staging and atomic publication remain; release preparation is deferred.

## RV06-L209: empty static-frame admission assessment

The isolated prototype removing static-presence admission is rejected. Nine new independent Word controls improve (99 Calibri break mismatches removed, one empty continuation page removed, maximum .196pt baseline residual and eleven improved raster pages), but the inherited audit changes 837 of 6916 PDFs. Twelve representative earlier Word controls show fifteen worsened and five improved raster pages, with maximum MAE increase 4.162835; four established focused assertions fail. Shared note retry can bypass existing body advance corrections, while physical fitting also changes reserve/fallback comparisons.

The existing guard and tests remain. No runtime change or font/size exception is integrated; broader body/note policy consolidation needs dedicated qualification. Accepted root rendering remains L208 with the permanent safe-managed constraint and all resource/cancellation/publication bounds intact. This records a confirmed failed qualification, not a font-cache corruption claim. Version 0.1.5 and deferred release preparation remain.

## RV06-L210: source-owned custom baseline footnote advances

Runtime **a7eb189f**, root **d2154ab5**, shares the existing exact custom-character owner mapper between endnotes and footnotes. Unique baseline prefixes remain protected; automatic/body owners, custom superscript, malformed/ambiguous/duplicate/missing sources and incomplete-owner fallbacks remain. No font/size branch or unsafe API is added; the rejected L209 static-presence removal stays excluded.

Nine independent Word controls remove all 99 Calibri custom-footnote break mismatches; all eleven raster pages improve. Three exact-spacing controls pass the unchanged .2pt baseline gate (.1pt maximum); six minimum controls retain inherited baseline errors up to 80.737pt and Word destination errors up to 81.047pt. Source mapping resolves 318 lines/9 note/navigation owners/22 clips/11 rules. All nine note boxes/329 placed-item coordinate owners, body/header layouts, 47 protected operations and nine prefix glyphs remain exact. Word emits one body-to-note link per control and no backlink; annotation behavior is unchanged. Office parity is not claimed.

All 253 note checks pass; the prefix regression fails the accepted parent and the ambiguity/alignment guard passes both. Windows: 2583 passes/zero failures/one optional private skip, strict builds clean. Linux: 336 passes/zero failures/eleven font skips, prior outcomes retained and two new passes; two archive warnings. Frozen: 6925 comparisons, all 6916 prior exact/nine new controls. Root: 377 focused checks/3763 fresh PDFs/two source blobs. Local 0.1.5 package smoke passes. The permanent safe-managed constraint, bounds, cancellation, cumulative work, bounded staging and atomic publication remain; release preparation is deferred.

## RV06-L211: custom note metrics from owned natural boxes

Runtime **8d06bc7d**, root **36aba86a**, shares one exact typed custom-character/unique-target mapper between layout and advances, resolved once per layout document. Only an absent dominating natural box admits a new wrapped-line metric view; existing mixed insets remain when the box is unchanged. Minimum bottom bearings, exact boxes and automatic/rich/kept/contextual/scaled/ambiguous/incomplete-owner fallbacks remain. No font/size exception or unsafe API is added.

Eighteen independent/retained Word controls qualify nine corrections/nine exact guards. The six retained inflated-box controls improve from a maximum 80.737pt baseline error to .196pt. Eight corrections pass the unchanged .2pt gate; one binding Calibri 14pt control retains .31pt, and the exact subordinate Calibri heldout retains .319pt. Fourteen raster pages improve, twelve stay exact, no tradeoffs. The broad baseline recomputation worsened two subordinate controls/four pages (maximum +.188227 MAE) and was rejected; the natural-box rule restores all six uniform/subordinate heldouts exactly. Source mapping resolves 752 lines/18 note/navigation owners/52 clips/26 rules; all eighteen body/header layouts and 98 protected operations remain exact. Navigation owners/pages hold; residuals up to .8pt remain. Full Office parity is not claimed.

All 255 note checks pass; the metric regression fails the accepted parent and the boundary guard passes both. Windows: 2585 passes/zero failures/one optional private skip, strict Release builds clean. Linux: 338 passes/zero failures/eleven font skips, prior outcomes retained and two new passes; two archive warnings. Frozen: 6934 comparisons, 6919 prior exact/six retained corrections/nine new controls. Root: 379 focused checks/3772 fresh PDFs/six source blobs. Local 0.1.5 package smoke passes. The permanent safe-managed constraint, bounds, cancellation, cumulative work, bounded staging and atomic publication remain; release preparation is deferred.

## RV06-L212: shared-run custom note qualification

Test-only qualification **8c8085bc**, root **624135f6**, adds coverage for wrapped adjacent text after a one-character custom prefix. All shipped library/CLI source trees remain identical to L211, including the permanent safe-managed constraint. Fifteen fresh Word controls across three fonts include double symbols, short/long suffixes and two/three-page continuations with 8/18pt minimum spacing. All pass the unchanged .2pt note-baseline gate (maximum .196pt), with matching line breaks and page placement; no runtime change is justified.

Exact source mapping resolves 812 lines/15 notes, 48 clips/24 rules, 153 prefix/suffix operations and 15 first-character navigation owners on 24 pages. Authored suffix sizes and split source offsets remain; Word emits no backlink. Destination residuals up to .6pt and inherited body horizontal residuals remain. Windows: 256 note passes/zero failures, strict Release clean. Linux: 251 passes/zero failures/five environmental font skips; the new regression passes and prior outcomes hold, with two archive SourceLink warnings. Root: 350 focused passes/17 fresh exact DOCX/PPTX PDFs/26 pages/one matching test Git blob. L211's wider corpus/catalogue/package evidence is inherited through identical shipped source trees, not presented as fresh execution. Version stays 0.1.5; release preparation is deferred.

## RV06-L213: admitted mixed-note typography

Runtime **65142cf8**, root **e9116b8f**, separates admitted endnote typography from legacy mixed pagination. The renderer chooses at most one admitted kind for its existing bounded retry; the shared owner scan verifies complete, unique plain complementary owners and retains their exact original source paragraphs. Equal IDs across kinds remain distinct, and losing an entire kind rejects publication. Existing static, rich, scaled and continuation guards remain; no font/size exception or unsafe API is added.

Thirteen independent/retained Word controls qualify seven corrections/six exact header fallbacks. The six empty-static controls remove 139 endnote break mismatches; the sole changed prior public mixed-note PDF improves against a fresh Word reference. All seven corrections pass the unchanged .2pt baseline gate (maximum .165pt). Seven raster pages improve/six exact/zero tradeoffs. Source mapping resolves 462 lines/26 typed note/navigation owners/52 clips/26 rules; body/static/complementary boxes and 160 protected text operations, including every complementary marker, stay exact. Header fallbacks retain 139 mismatches; destination residuals up to .8pt and unadmitted long mixed continuations remain. Full Office parity is not claimed.

All 258 focused note checks pass; the positive regression fails the accepted parent and the guard passes both. Windows: 2588 passes/zero failures/one optional private skip, strict Release clean. Linux: 341 passes/zero failures/eleven font skips; prior outcomes and L212 coverage hold, with two archive warnings. Frozen: 6962 PDF comparisons/6948 earlier exact/one qualified prior correction/thirteen Word comparisons. Root: 382 focused passes/3800 fresh PDFs/four source blobs; local 0.1.5 package smoke passes. The permanent safe-managed constraint, resource bounds, cancellation, cumulative work, bounded staging and atomic publication remain. Version stays 0.1.5; release preparation is deferred.

## RV06-L214: mixed notes in plain static frames

Runtime `808ab252`, root `7a647a00`, shares existing plain footnote source eligibility with static document-end typography. Footnote frame admission, mixed paragraph-boundary and pagination policy remain unchanged; one successful bounded retry and complete typed owners remain required. No font/size exception or footnote-presence guard was removed.

Twelve independent/retained Word controls correct 277 inherited endnote break mismatches; all twelve raster pages improve, with no tradeoffs. Endnote baselines pass the unchanged .2pt gate (maximum .165pt), and line/page ownership matches. Nine fitting controls pass full note geometry. Three overflowing-footer controls remain partial: exact inherited complementary footnote placement errors reach 65.647pt and destination residuals 64.937pt. Mapping resolves 460 lines/24 typed note/navigation owners/48 clips/24 rules; body/static/complementary boxes and 171 protected operations stay exact. Full mixed-footer parity is not claimed.

All 260 focused note checks pass; the positive regression fails the accepted parent, and the fallback guard passes both. Windows: 2590 passes/zero failures/one optional private skip, strict Release clean. Linux: 343 passes/zero failures/eleven font skips, all prior outcomes retained and two archive warnings. Frozen: 6968 PDF comparisons/6956 earlier exact/six qualified header corrections/six new footer comparisons. Root: 384 focused passes/3806 fresh PDFs/three source blobs; local 0.1.5 package smoke passes. Permanent safe-managed code, resource bounds, cancellation, cumulative work, bounded staging and atomic publication remain. Version stays 0.1.5; release preparation is deferred.

## RV06-L215: mixed-note footer clearance

Runtime `5c91aec5`, root `4bd9fe43`, shares existing plain note-source checks for physical static frames. New complementary admission also requires plain paragraph ownership; original own-footnote behavior remains. Pages carrying a footnote share the selected physical footer edge with endnotes, while empty endnote continuations retain their original fallback. Mixed footnote typography exclusion remains independent, with one successful bounded endnote retry.

Twenty-four Word controls qualify twelve corrections/twelve exact guards. Four corrections pass the unchanged .2pt full geometry gate; two short Calibri controls retain .209pt and six continuations remain partial. Sixteen paired raster pages improve, twenty stay exact, no tradeoffs; three inherited unpaired pages stay byte-exact. One Arial body-reference page and its endnote reflow independently match Word. Mapping resolves 1454 lines/48 typed note/navigation owners/116 clips/58 rules and protects 1467 text/font/advance operations; other body/static/complementary geometry stays exact. Four inherited destination-page gaps, destination residuals up to 20.983pt and larger continuation placement errors remain. Full mixed-continuation parity is not claimed.

All 262 focused note checks pass; the positive regression fails the accepted parent, the fallback guard passes both, and the continuation assertion fails the rejected physical-only prototype. Windows: 2592 passes/zero failures/one optional private skip, strict Release clean. Linux: 345 passes/zero failures/eleven font skips, prior outcomes retained, two archive warnings. Frozen: 6980 PDF comparisons/6965 earlier exact/three qualified prior corrections/twelve new comparisons. Root: 386 focused checks/3818 fresh PDFs/four source blobs; local 0.1.5 package smoke passes. Rejected all-page and physical-only models remain separate evidence. Permanent safe-managed code, bounds, cancellation, cumulative work, bounded staging and atomic publication remain. Version stays 0.1.5; release preparation is deferred.

## RV06-L216: partial endnote heads above footnotes

Runtime `dc174741`, root `9fbbe1eb`, uses existing widow rules and shared paragraph-after charges to place an admitted plain endnote head in body space above footnotes. Mixed paragraph boundaries, fitting output, typed owners and rich/scaled/static/absent-static fallbacks remain. Typography retains one bounded successful retry; no font/size exception was added.

Twenty-seven Word controls qualify seven corrected head pages/twenty exact guards and remove six inherited extra pages. All output page counts and navigation destination pages match Word. Twenty paired raster pages improve, twenty-five remain exact, no tradeoffs. Source mapping resolves 1835 lines/54 typed notes and navigation pairs/142 clips/71 rules and preserves 766 body/static/complementary text/font/advance operations. Calibri .202pt endnote/.209pt footnote errors, two page-line mismatches in each fitting header/footer continuation and existing long-footnote/destination residuals up to 8.017pt remain. No threshold relaxation or full Office-parity claim.

All 264 focused note checks pass; the positive regression fails the exact accepted parent and the fallback passes both. Windows: 2594 passes/zero failures/one optional private skip, strict Release clean. Linux: 347 passes/zero failures/eleven font skips, prior outcomes retained, two archive SourceLink warnings. Frozen: 6983 PDF checks/6976 earlier exact/four qualified prior corrections/three new comparisons. Root: 388 focused checks/3821 fresh PDFs/three source blobs; local 0.1.5 package smoke passes. Permanent safe managed code, all bounds, cancellation, cumulative work, bounded staging and atomic publication remain. Version stays 0.1.5; release preparation is deferred.

## RV06-L217: occupied mixed-note separator capacity

Runtime `e20017aa`, root `765b8273`, uses the existing nominal separator charge to exclude latent empty paragraph-after padding from the occupied head edge. Complementary clips/glyphs/rules remain exact and the edge explicitly clears the visible rule. Fitting output, mixed boundaries, typed owners and rich/scaled/static/absent-static fallbacks remain; typography retains one bounded successful retry.

Twenty-seven Word controls qualify three corrections/twenty-four exact guards. All six fitting header/footer page-line mismatches disappear; corrected endnote line/page assignments match Word. Nine paired raster pages improve/thirty-six exact/zero tradeoffs. Source mapping resolves 1835 lines/54 typed notes and navigation pairs/142 clips/71 rules and preserves 766 body/static/complementary operations. Calibri .202/.209pt, existing long-footnote residuals and destination error up to 8.017pt remain. No threshold relaxation or full Office-parity claim.

Fresh Windows qualification: 38 deduplicated affected tests pass, strict Release clean; the prior L216 full catalogue remains prior evidence. Linux: 348 PDF/note passes/zero failures/eleven font skips, new regression passes and prior outcomes retained, two archive SourceLink warnings. The exact frozen regression fails accepted L216. Frozen: 6983 PDF checks/6980 earlier exact/three retained Word corrections. Root: 389 focused checks/3821 fresh PDFs/three source blobs; local 0.1.5 package smoke passes. The interrupted diagnostic sweep contributes no acceptance passes. Permanent safe managed code, all bounds, cancellation, cumulative work, bounded staging and atomic publication remain. Version stays 0.1.5; release preparation is deferred.

## RV06-L218: source admission across deferred endnote heads

Runtime `17123bc2`, root `22e82932`, carries the qualified original mixed/static source frame across a deferred endnote head's page turn. Typography and nominal separator ownership persist on the empty page; complementary source admission shares existing plain-paragraph semantics. Mixed boundaries, footnote typography, complementary clips/rules/glyphs and rich/scaled/static/absent-static fallbacks remain, with one bounded successful retry.

Thirty Word controls qualify two custom/automatic Calibri corrections/twenty-eight exact guards. All 66 inherited endnote break mismatches disappear; note line/page assignments match Word. Two raster pages improve/fifty-two exact/zero tradeoffs. Source mapping resolves 2210 lines/60 typed notes and navigation pairs/164 clips/82 rules and preserves 1116 protected operations. Destination maximum drops from 8.017pt to .869pt. Calibri deferred endnotes retain .362pt baseline error, alongside earlier .202/.209pt and long-footnote residuals. No threshold relaxation or full Office-parity claim.

Fresh Windows qualification: 40 deduplicated affected tests pass, strict Release clean; the full catalogue remains L216 evidence. Linux: 350 PDF/note passes/zero failures/eleven font skips, new regressions pass and prior outcomes retained, two archive SourceLink warnings. The exact frozen positive fails accepted L217; fallback passes both. Frozen: 6986 PDF checks/6982 earlier exact/one qualified prior correction/three new comparisons. Root: 391 focused checks/3824 fresh PDFs/two source blobs; local 0.1.5 package smoke passes. Permanent safe managed code, all bounds, cancellation, cumulative work, bounded staging and atomic publication remain. Version stays 0.1.5; release preparation is deferred.

RV06-L219 (**d28a6104**, integrated **217ada41**) shares the actual
widow-aware paragraph take with admitted middle-footnote anchoring. Thirty-three
Word controls qualify one correction/thirty-two exact guards; the middle-page
baseline passes .2pt, one raster page improves/62 exact/no tradeoffs. Source mapping
resolves 2593 lines/66 typed notes/186 clips/93 rules; 2277 protected operations
remain exact. Three new explicit-off controls retain an inherited Arial page-line
gap/11.63pt; Calibri .362/.202/.209pt and .869pt navigation remain partial.
Fresh qualification: 41 affected Windows passes, 351 Linux passes/eleven font skips,
6989 frozen PDF checks and 392 root checks/3827 PDFs. Prior full Windows is L216;
strict Release, safe managed constraints and local 0.1.5 package smoke pass.

RV06-L220 (**a949da98**, integrated **e335204a**) charges the complete
separator inset in admitted continued footnote frames. Seventy-two Word controls
qualify 29 source corrections/43 exact guards, removing 35 page-line mismatches;
all note lines/pages match Word. Mapping resolves 7734 lines/144 typed notes/498
clips/249 rules; 4974 protected operations and canonical glyph/font/advance emission
remain exact. Forty-four raster pages improve/122 exact/fourteen new terminal pages
retain small tradeoffs (maximum MAE increase .008153/pixel-change ratio .000487).
All 33 retained controls have zero tradeoffs; no prior passing gate is relaxed.
Calibri .362/.202/.209pt and .869pt navigation remain partial. Fresh qualification:
42 affected Windows passes, 352 Linux passes/eleven font skips, 7028 frozen PDFs and
393 root checks/3866 PDFs. Full Windows remains L216; strict Release, permanent safe
managed code and local 0.1.5 package smoke pass. No widow-metadata override.

RV06-L221 (**1bec78a7**, integrated **68a1993b**) applies the shared source
kerning-threshold and nominal-advance rules to jointly admitted mixed notes in
one bounded typography retry. Both complete typed owners are required; custom
prefixes, source offsets, partially normalized pairs and existing single-kind,
rich/scaled/static fallbacks remain. The exact frozen emission regression fails
the accepted parent at the intended advance assertion.

All 72 retained Word controls preserve 7734 source lines/144 typed notes, line/page/
baseline assignments, 498 clips/249 rules, 144 navigation pairs and 4974 protected
body/static/complementary operations. All 4392 footnote terminal advances improve;
155 paired raster pages improve/25 exact/zero tradeoffs. Inherited Calibri
.362/.202/.209pt baseline and .869pt navigation residuals remain partial. No font/
size exception, widow override, threshold relaxation or full Office-parity claim.
Fresh qualification: 45 affected Windows passes, 355 Linux passes/eleven existing
font skips, 7028 frozen PDFs/6956 prior exact/72 independently qualified changes,
396 root checks/3866 fresh PDFs/three source blobs. Full Windows remains L216;
strict Release, permanent safe managed code and local 0.1.5 package smoke pass.
Version stays 0.1.5; release preparation is deferred.

RV06-L222 (**02c51486**, integrated **04ef342d**) is test-only qualification of
mixed-note positive thresholds and authored suffixes. All library, CLI and tool
source trees remain identical to L221. Twenty-four Word controls preserve 1976
source lines/48 typed owners and exact line/page/font/size ownership, including
204 threshold-dependent source-line changes. Mapping resolves 136 clips/68 rules,
48 navigation pairs and 420 body/static operations on 48 pages. Twenty controls
pass the unchanged .2pt baseline gate; four continued Calibri endnotes retain
inherited .362pt error. Terminal advance residuals up to .611pt, threshold-delta
residual .235pt and destination .8pt remain partial. No further runtime defect or
guard change is admitted.

One deterministic emission regression covers sixteen source compositions and
checks met/unmet content thresholds, actual authored suffix size and intact source
records. Fresh checks: 273 Windows note passes, 268 Linux note passes/five existing
font skips, strict Release clean, 24 frozen/root PDF identities and five focused
root methods. L221's wider corpus, PDF group, full-catalogue and package evidence
remain inherited through identical shipped source trees; they were not rerun.
Safe managed constraints and version 0.1.5/release deferral remain.


### RV06-L223: typed note separator registration

Word ignores special note formatting unless settings register the matching typed
separator ID. OoxPdf previously consumed that formatting, moving 8/16pt endnote
content by about five points. The reader now resolves the shared footnote/endnote
registration rule and supplies Word's default separator when needed. Registered
sources, normal note owners and existing layout eligibility remain intact. One
new regression and a corrected authored capacity-fixture declaration preserve
the prior expected take. Frozen **544398c7**, root **6447d723**.

Sixteen Word controls retain 1587 canonical lines/30 notes/86 clips/43 rules/30
navigation pairs and 1443 protected operations: seven raster pages improve/33
exact, with no tradeoffs or new navigation error. All 7067 frozen byte/page checks
pass: 7051 retained PDFs exact, one latent separator clip change and 15 new Word
inputs. The changed retained case preserves full text/rule/annotation emission
and exact PNG identity. Root checks all 3905 qualified PDFs and three source blobs.
Windows passes 274 note/94 PDF/one static-comment checks (368 unique methods),
strict Release is clean, Linux passes 357 checks with eleven environmental skips
and two archive SourceLink warnings, and the local packed 0.1.5 smoke passes.
The new test fails the intended assertion against the exact parent library.

Calibri's .362pt baseline and the new registered fitting control's .938pt
destination error remain partial. The origin-only draft remains unaccepted;
source-derived first-line placement continues as L224. Safe managed code, existing
bounds/cancellation/work/staging/publication limits and release deferral remain.

### RV06-L224: deferred endnote first-line origin

Admitted mixed endnotes deferred to a fresh page used a fixed subsequent-line
offset. Their head now follows the source baseline floor enlarged by the visible
mark's measured ascent. Original frame/source admissions and unavailable-metric,
rich/scaled fallbacks remain. No font/size condition is added. The new regression
covers 192 compositions, witnesses all 24 marker/size/registration/ascent combinations
and checks unavailable-metric fallbacks; it fails the intended assertion against
the exact parent library. Frozen **dd94a880**, root **8e5a94ae**.

All 119 Word controls retain 12195 source lines/238 notes, line/page assignments,
764 clips/382 rules/238 navigation pairs and 9864 protected operations. Twenty-six
heads improve/93 remain exact; 26 raster pages improve/263 exact/no tradeoffs.
Corrected rule errors stay below .115pt and Calibri glyph errors below .184pt.
Raw-inset and measured-ascent-only drafts fail the unchanged .2pt rule gate and
remain rejected. Arial/Times New Roman default marks with Calibri note content
improve from 1.066pt to .834pt glyph error and remain partial. Twenty-four note
destinations have larger differences to Word's rounded coordinates, by at most
.189pt; all targets follow their moved clips and maximum .938pt error is inherited.

Strict Release has zero warnings/errors. Windows passes 275 note/94 PDF/one static
comment checks (369 unique methods); Linux passes 358 checks with eleven existing
environment skips and two archive SourceLink warnings. All 7075 frozen byte/page
checks pass: 7048 retained identities exact, 19 independently qualified changes
and eight new Word inputs. Root verifies 3913 PDFs and two source blobs; packed
0.1.5 smoke passes. Full Windows remains L216 evidence. Shared note-gap diagnosis
continues as L225; permanent safe managed code, existing bounds/work/publication
limits and release deferral remain.

### RV06-L225: deferred-head separator-gap ownership

Admitted source-first-inset heads consumed the separator-height deficit again in
their content gap. Placement now reuses that existing term once and keeps the
mark gap as a lower bound. Existing fitting/continuation, unavailable-metric and
rich/scaled behavior remains. A Word-grounded regression covers 32 compositions
and preserves canonical owners and matching metric floors; it fails the intended
gap assertion against the exact parent. Frozen **bac866af**, root **113151ba**.

Twenty-two Word controls map 2001 source lines/44 notes/124 clips/62 rules/44
navigation pairs and 1672 protected operations. Eight endnote heads improve below
.173pt glyph error; fourteen guards remain byte-identical. Rules are unchanged,
eight raster pages improve/41 exact, and no raster or destination-coordinate
tradeoff appears. The broader guards retain 22 inherited page-line mismatches;
paired baseline deltas reach 102.532pt across mismatched flows and destination
residuals 6.393pt. Complementary footnote errors up to .834pt remain partial.
These results qualify the bounded correction, without claiming wider Office parity.

Strict Release is clean. Windows passes 276 note/94 PDF/one static-comment checks
(370 unique methods); Linux passes 359 checks/eleven existing environmental skips
and two archive SourceLink warnings. All 7095 frozen byte/page checks pass: 7073
retained outputs exact, two independently qualified changes and twenty new Word
inputs. Root checks 3933 PDFs/two source blobs; packed 0.1.5 smoke passes. Full
Windows remains L216 evidence. Shared physical-footnote gap ownership continues
as L226. Safe managed code, existing bounds/work/publication limits and release
deferral remain.

### RV06-L226: shared gap ownership for fresh physical-footnote heads

Fresh admitted physical-footnote heads repeated the separator-height deficit in
their content pushdown. They now share the accepted endnote gap helper and the
existing kind-independent nominal separator predicate. Matching metrics keep the
mark-gap floor; no-static, rich/scaled and continued heads retain their treatment.
Reserve capacity is unchanged. One regression covers sixty compositions and
fails the intended Word-gap assertion against the exact parent. Frozen **63a132c3**,
root **7bd1ef25**.

All 22 retained Word controls preserve 2001 source lines/44 notes/124 clips/62 rules,
44 navigation pairs and 1223 protected operations. Twelve first heads improve
below .121pt glyph error/ten outputs remain exact; twelve raster pages improve/
37 exact/no tradeoffs. Every rule, endnote and continued-footnote position remains
unchanged. Eight destinations differ more from Word's rounded coordinates, by at
most .145pt; each follows its corrected clip. Later footnotes retain .834pt error.
Broader guards retain 22 page-line mismatches, paired deltas up to 102.532pt across
mismatched flows and inherited 6.393pt destination error. These remain partials.

Strict Release is clean. Windows passes 277 note/94 PDF/one static-comment checks
(371 unique methods); Linux passes 360 checks/eleven existing skips and two archive
SourceLink warnings. All 7095 frozen byte/page checks pass: 7083 retained exact and
twelve independently qualified changes, with no new inputs. Root checks 3933 PDFs
and three source blobs. Two C# files differ only in LF/CRLF checkout endings;
normalized contents, Git blobs and both raw hashes are retained. Packed 0.1.5 smoke
passes; full Windows remains L216 evidence. Continuation source-gap admission
continues as L227. Safe managed code, existing bounds/work/publication limits and
release deferral remain.

### RV06-L227: shared gap admission for occupied footnote continuations

The physical-frame guard already checks complete original document sources.
Occupied continuations now use the same nominal separator gap as fresh heads,
by removing one fresh-head-only exclusion. No additional admission state or
font/size/threshold/document-shape branch is introduced. Empty-page, no-static,
rich/scaled and incomplete-metric fallbacks retain their treatment; reserve/take
capacity is unchanged. Two regressions cover 48 compositions and four rich-source
guards. The exact frozen test assembly fails against the parent at its intended
15.874pt gap assertion, after passing the no-static guard. Frozen **5b7b90b8**,
root **379bf874**.

All 22 retained Word controls preserve 2001 source lines/44 notes/124 clips/62 rules,
44 navigation pairs and 1223 protected operations. Eight occupied continuations
improve from .834pt to below .121pt glyph error/fourteen outputs remain exact.
Eight raster pages improve/41 exact/no tradeoffs. Every first head, rule,
destination and reserve/take capacity remains exact to the parent. Broader guards
retain 22 page-line mismatches, paired deltas up to 102.532pt across mismatched
flows and inherited 6.393pt destination error. These remain explicit partials.

Strict Release is clean. Windows passes 279 note/94 PDF/one static-comment checks
(373 unique methods); Linux passes 362 checks/eleven existing environmental skips
and two archive SourceLink warnings. All 7095 frozen byte/page checks pass: 7087
retained outputs exact/eight independently qualified changes/no new inputs.
Root checks 3933 PDFs/two source blobs; packed 0.1.5 smoke passes. Full Windows
remains L216 evidence. Source-owned minimum-line metric semantics continue as
L228; permanent safe managed code and bounds/work/publication policies remain.
Release preparation is deferred.

### RV06-L228: shared natural boxes for minimum note content

Source-owned content-only minimum note lines omitted the existing natural
Windows/hhea box. They now reuse the automatic related-story metric rule;
first source-prefix lines, authored minima, typographic requests, exact boxes
and providers without resolved face semantics retain their treatment. Two
regressions cover 24 source compositions/both kinds. The exact frozen test
assembly fails against the parent at its intended missing natural-box assertion.
Frozen **e2112fc3**, root **b77a3200**.

The 22 retained Word controls qualify four corrections/eighteen exact guards.
All 2001 source lines/44 notes/124 clips/62 rules/44 navigation pairs retain
ownership; canonical glyph/font/advance sequences and 425 body/static operations
remain exact. Twenty-one page-line gaps disappear; one Times continuation gap
and 11.574pt paired delta remain inherited. Nine raster pages improve/40 exact/
no tradeoffs. Corrected footnote glyphs pass .2pt (maximum .186pt). Previously
failing first-rule gaps improve but remain partial up to .729pt; every prior
passing rule gate stays passing. Endnote constant offsets remain .863pt.
Destination maximum drops 6.393 to 1.039pt, with no coordinate tradeoffs.

Strict Release is clean. Windows passes 281 note/94 PDF/one static-comment checks
(375 unique methods); Linux passes 364 checks/eleven existing environmental skips
and two archive SourceLink warnings. All 7095 frozen byte/page checks pass: 7091
retained outputs exact/four independently qualified changes/no new inputs. Root
checks 3933 PDFs/three source blobs; packed 0.1.5 smoke passes. Full Windows remains
L216 evidence. First-line natural box/bearing work continues as L229. Safe managed
code and bounds/work/publication policies remain; release preparation is deferred.

### RV06-L229: shared natural first-line box and bearing

Minimum note source-prefix lines now share the larger natural Windows/hhea box
and its matching baseline bearing with the existing content metric owner.
Authored minima, typographic requests, exact/scaled/rich/partial-metric and
unenlarged-box fallbacks retain independently tested treatment. No additional
font, size, threshold or document-shape branch is introduced. Two regressions
cover 24 source compositions/both kinds; the exact frozen test assembly fails
against the parent at its intended first natural box/bearing assertion.
Frozen **e87e560a**, root **4b42fe78**.

Twenty-two retained/four fresh Word controls qualify eight corrections/eighteen
byte-exact guards. All 2374 source lines/52 notes/154 clips/77 rules/52 navigation
pairs and 503 body/static operations retain ownership; canonical glyph/font/
advance sequences remain exact. Corrected footnote glyphs .186pt/endnote glyphs
.180pt/footnote rules .129pt pass unchanged .2pt checks. Endnote separator objects
remain exact. Eight raster pages improve/52 exact/no tradeoffs. One inherited
Times page-line gap/11.574pt paired delta remains. Four destination-coordinate
errors increase by up to .784pt; targets follow the corrected source clips on
their correct pages. Global destination maximum remains 1.039pt. These tradeoffs
are separate from visible geometry; no full Office parity claim or relaxed gate.

Strict Release is clean. Windows passes 283 note/94 PDF/one static-comment checks
(377 unique methods); Linux passes 366 checks/eleven existing environmental skips
and two archive SourceLink warnings. All 7099 frozen byte/page checks pass: 7091
prior outputs exact/four qualified retained changes/four new Word inputs. Root
checks 3937 PDFs/three source blobs; packed 0.1.5 smoke passes. Full Windows remains
L216 evidence. Mixed-note separator occupancy work continues as L230. Safe managed
code and bounds/work/publication policies remain; release preparation is deferred.

### RV06-L230: shared separator occupancy and partial-head origin

Qualified mixed-note flow now retains the canonical blank separator line's
measured ascent when excluding latent flow padding. Partial-head admission uses
the same cursor-plus-gap origin as placement; it previously reserved 6pt more.
Existing source admission, raw clips, visible rule clearance, widow policy,
incomplete metrics, empty-frame progress and bounded retry remain. No font,
size, threshold or marker-specific branch is introduced. Frozen **d04d1206**,
root **33833095**.

The ascender-only model is rejected. Although it improves the initial 34
controls, the full retained audit exposes six new page-line gaps in three
previously passing Arial automatic-marker/threshold cases. Its exact patch,
immutable runtime and independent negative comparisons remain ignored.
The paired correction retains those three accepted-parent PDFs exactly.
Two fake-metric tests guard ownership and coupled admission; both pass against
the accepted parent. The exact new test assembly fails admission against the
rejected intermediate library. Independent Word comparisons establish the
accepted-parent rendering defect.

Thirty-seven Word controls qualify six improvements/31 byte-exact guards.
All 3769 source lines/74 notes/238 clips/119 rules/74 navigation pairs and 789
body/static operations retain ownership; canonical glyph/font/advance sequences
stay exact. All seven page-line gaps disappear. Twelve raster pages improve/
81 exact/no tradeoffs. Changed line positions pass .2pt (maximum .139pt),
complementary physical footnotes/rules stay exact and the deferred fresh
separator matches Word within .006pt. Unchanged Times .211pt and fitting
mixed-font .694pt offsets remain partial. One destination page corrects;
the coordinate maximum remains 1.039pt with no new coordinate tradeoffs.

Strict Release is clean. Windows passes 285 note/94 PDF/one static-comment checks
(379 unique methods). The complete frozen Windows catalogue passes 2615 tests/
zero failures/one optional private-case skip. Linux passes 368 checks with eleven
existing environmental skips/two archive SourceLink warnings. All 7107 frozen
and 3945 root PDF byte/page identities pass: 7098 prior outputs exact/one qualified
retained change/eight new Word inputs. Two source blobs and packed 0.1.5 smoke pass.
Fitting-note separator origin/gap work continues as L231. Safe managed code and
bounds/work/publication policies remain; release preparation is deferred.


### RV06-L231: fitting-note first-bearing ownership and shared content gap

Complete fitting mixed-note heads now share the existing separator content-gap
deficit when their current origin already owns the resolved first bearing.
Deferred heads keep their source-first origin. Uncovered bearings, partial heads,
unavailable metrics and absent static frames retain explicit fallbacks. Existing
mixed source admission and shared hhea coverage are reused; no font, size,
threshold or marker-specific exception is introduced. Frozen **0258cb0d**,
root **61c39b54**. Two origin-shift models are rejected by neighboring
Word controls; exact patches, immutable runtimes and comparisons remain ignored.

Thirty-seven retained/eighteen fresh Word controls qualify six fitting mixed
improvements/49 byte-exact guards. All 4463 source lines/110 notes/310 clips/155
rules/110 navigation pairs and 951 body/static operations retain ownership;
canonical glyph/font/advance sequences stay exact. All line/page assignments,
separator origins/rules and physical footnotes remain exact. Six raster pages
improve/105 exact/no regressions. Prior passing .2pt glyph gates remain passing;
mixed residuals up to .367pt remain partial. Three destination-coordinate errors
increase by up to .713pt; targets still follow the correct typed note clips and
pages. Coordinate maximum is 1.311pt. No full Office parity claim or relaxed gate.

Two focused tests cover 32 fitting source combinations. The exact new test
assembly fails against accepted L230 at the intended excess-gap assertion;
unowned/unavailable bearing and absent-frame fallback coverage passes against
both versions. Strict Release is clean. Windows passes 287 note/94 PDF/one
static-comment checks (381 unique methods); Linux passes 370 checks with eleven
existing environmental skips/two archive SourceLink warnings. All 7125 frozen
and 3963 root PDF byte/page identities pass: 7105 prior outputs exact/two qualified
retained changes/eighteen new Word inputs. Two source blobs and packed 0.1.5 smoke
pass. Last full Windows catalogue remains L230 (2615 passes/one optional private
skip). Fitting rule-bearing investigation continues as L232. Safe managed code,
bounds/work/publication policies and release deferral remain.


### RV06-L232: source first bearing in fitting separator rule painting

Complete fitting mixed-note rules now retain the original source first bearing,
bounded by line metrics. The existing content hhea owner expresses font coverage
at the mark size for painting; flow keeps each source run's size. Deferred,
partial, rich/scaled and missing-metric/frame fallbacks remain. No additional
font, point-size, threshold or marker-specific exception is introduced.
Frozen **36a32d98**, root **fd641149**.

Fifty-five Word controls qualify six rule improvements/49 byte-exact guards.
All 4463 source lines/110 notes/310 clips/155 rules/110 navigation pairs and 951
body/static operations retain ownership. Only six rectangle Y positions change;
all glyph/font/advance/text state, flow frames, physical footnotes, clips and
navigation stay exact. Paint stays inside its canonical clip. Five raster pages
improve/106 exact/no regressions. Three changed rules pass .2pt/three remain
partial (maximum .386pt). Previous passing gates hold. Inherited glyph residuals
up to .501pt and L231's destination coordinate tradeoffs remain unchanged.
No full Office parity claim or relaxed gate.

Two regressions cover eight mixed/missing-line-bearing compositions. The exact
new test assembly fails the independent Word rule/first distance against accepted
L231; the missing-metric fallback passes against both. A draft matching fake does
not reproduce the matching Word source context; its failed calibration reports
remain ignored. Matching Word guards retain their independent byte-exact gates.
Strict Release is clean. Windows passes 288 note/94 PDF/one static-comment/one
explicit separator-fallback checks (383 unique methods); Linux passes 372 checks
with eleven existing environmental skips/two archive SourceLink warnings.
All 7125 frozen and 3963 root PDF byte/page identities pass: 7119 prior outputs
exact/six qualified changes. Two source blobs and packed 0.1.5 smoke pass.
Last full Windows catalogue remains L230 (2615 passes/one optional private skip).
Fitting content-bearing/flow-charge work continues as L233. Safe managed code,
bounds/work/publication policies and release deferral remain.

### RV06-L233: fitting-source alignment and rejected flow charge

The fitting separator regression sources now use the same minimum-8 note
spacing as their independent Word controls. This restores matching-font coverage
alongside mixed fonts and missing line metrics: twelve source compositions,
six focused methods, zero failures/skips. Root **44ae512b** changes tests only.
Strict Release has zero warnings/errors; eight fresh control PDFs remain exact
to the accepted L232 runtime.

Eight new Word controls vary separator mark/note-prefix sizes through (8,10),
(10,12), (10,14) and (12,16), using Tahoma/Verdana marks and Calibri content.
All 296 source lines/16 notes retain exact line/page assignments; inherited
baseline offsets reach .488pt. The proposed charge for content bearing beyond
the mark box predicts four worse neighbors, reaching .737pt. These are diagnostic
predictions; the model is rejected before changing production. Existing .2pt
gates and source layout remain. L232's full runtime qualification still applies;
no new full-catalogue qualification or Office parity is claimed. Source-owned
note-link spacing is the next bounded rendering investigation, L234.

### RV06-L234: resolved source gap for note-link spacing

The existing complete-owner path reads its after-gap from resolved metadata on
the first line, while the final visible line keeps ownership of the annotation
extension. Missing metadata retains the authored fallback; existing eligibility
guards remain. The shared spacing qualifier already excludes contextual sources
from this path, so the review does not demonstrate a contextual rendering defect.
Frozen **55f998a1**, root **16e336a5**.

Seventeen focused methods pass on Windows and Linux with zero failures/skips.
Strict Windows Release has zero warnings/errors; Linux retains two archive
SourceLink warnings. Twenty-four retained Word controls/48 pages preserve exact
PDF identity, including text, graphics and annotations. Their source/PDF/view
identities transfer through the accepted L232 frozen audit. One source blob and
raw source hash match. This is bounded ownership maintenance; no new full audit,
catalogue qualification, changed visual behavior or broader fidelity is claimed.
The rendering roadmap continues with public PPTX autofit/overflow probes, L235.


### RV06-L235: Office default PPTX kerning threshold

Omitted PPTX kerning now uses Office's 12pt default through the existing shared
source rule. Explicit thresholds retain their boundary/scaling; the qualified
explicit-zero no-wrap fallback remains. Twelve independent Office pairs across
four fonts, square/no-wrap and 10/12/18pt have identical omitted/explicit-12pt
text operations. The exact new frozen test assembly fails the accepted parent
at 12pt equality after passing its below-threshold composition. Frozen
**cc724b30**, root **5c8d6880**; three source blobs/raw hashes match.

Twenty-eight controls qualify eleven changed PDFs/17 exact: nine raster pages
improve/eighteen stay exact/one Calibri no-wrap control remains a partial tradeoff
(MAE +.0146815, changed-pixel ratio +.00009163). Cambria-fit MAE falls 1.359 to
.366. Source text/font/size/baseline state, paragraph/run cascades and frames
retain identity; advances, dependent X positions and seven highlight widths
change. No passing gate is lost or relaxed; complete Office parity is not claimed.

Strict Windows Release has zero warnings/errors; all 185 typography tests pass.
Four portable Linux regressions pass, with two existing archive SourceLink
warnings. All 158 retained PPTX contexts are accounted for: 157 exact/one
independently qualified change/zero unknowns. Frozen and root each reproduce 186
qualified PPTX PDFs; eight root DOCX guards remain exact. This is an affected
PPTX qualification, not a new full-catalogue or broad L232 audit. Last full
Windows remains L230. Safe managed code, resource/work/publication limits and
0.1.5 release deferral hold. Glyph-advance residuals continue as bounded L236.


### RV06-L236: glyph-position diagnosis

Seventeen independently qualified controls map 2482 visible glyph positions.
Unconditional precise-width compensation predicts worse mean Office X error in
17 of 25 neighbors; it is rejected before production. Calibri baseline/advance
residuals remain partial. Matching font metrics and retained hinting tables provide
no evidence of cache corruption or missing hinting. This is a diagnostic review;
L235 qualification remains unchanged.

### RV06-L237: automatic legendless line-chart plot reserves

Qualified untitled legendless line charts share existing measured cartesian label
reserves and use the source font's selected normal line box for top clearance.
Manual/composed/extra-axis, hidden/top-category and too-small frames retain their
prior layout fallback. A horizontal-only prototype produced two raster tradeoffs
and was rejected. Frozen **66326db3**, root **e09926db**.

Twenty-three independent Office controls qualify 19 improved raster pages/four
exact guards/no tradeoffs. Primary MAE falls 5.329 to 1.003. Source text/font/size/
advance/text state, styles, indices and path segmentation retain identity; plot,
clip, series and dependent label positions change. Top/right residuals remain
within .007/.001pt. Left/bottom residuals reach 1.423/2.722pt; default gridline
styling remains partial. No prior gate is relaxed or complete Office parity claimed.

Strict Windows Release has zero warnings/errors; 386 chart tests pass. Six portable
Linux methods pass, with two existing archive SourceLink warnings. Exact frozen
new tests fail the accepted parent twice while their fallback method passes.
All 158 retained PPTX contexts remain exact. All seven native-line inputs among
252 tracked PPTX fixtures are freshly covered: six exact/one independently
qualified change. Frozen/root each reproduce 188 PDF checks; eight root DOCX
guards remain exact. Five source blobs/raw hashes match. This is bounded affected
qualification; last full Windows remains L230. Safe managed code, resource/work/
publication limits and 0.1.5 release deferral hold. Gridline defaults continue as L238.


### RV06-L238: modern line-chart gallery gridline roles

Single line plots with verified modern gallery declarations 102/118/126 inherit
1pt major/minor gridline roles from mapped tx1 through the existing linear-light
.25/.5 white tint. Preferred style metadata preserves the selected legacy tree
and existing series recipes; authored/visibility/style-part/legacy/composed/other-
kind paths retain their fallback. A broader Cartesian prototype has eleven raster
tradeoffs, including one increased MAE, and is rejected. Frozen **77d575a9**, root
**21ba5045**.

Thirty-six Office controls qualify twelve improved raster pages/24 exact/no
tradeoffs. Span MAE falls 1.0025 to .7811. Twenty grouped gridline strokes change;
all 559 text operations/2772 path and clip geometries/74 embedded-font program
checks and axis/tick/series styles remain exact. Width-only 1.5pt declarations
still inherit 1pt; minor-grid spacing/coverage and axis-style residuals remain
partial. No existing fidelity threshold is relaxed or full Office parity claimed.

Strict Windows Release is clean: 391 chart/201 OOXML-resource methods pass. Nine
portable Linux checks pass, with two archive SourceLink warnings. Exact frozen new
tests demonstrate two parent failures/three protected passes. All 252 tracked
public PPTX inputs are freshly compared: 251 exact/one qualified span change;
all 158 retained contexts stay exact. Frozen/root each reproduce 446 PDF check
executions (36 controls+252 public+158 retained, overlapping inputs); eight root
DOCX guards stay exact. Three source blobs/raw hashes match. Last full Windows
remains L230. Permanent safe managed code, resource/work/publication limits and
0.1.5 release deferral hold. Partial line-gridline inheritance continues as L239.


### RV06-L239: partial major-gridline property ownership

Qualified single line plots retain authored major-gridline width/color/alpha and
decorations while inheriting omissions from the role through the existing reader.
Width-only keeps 1.5pt; color-only inherits 1pt. Unsupported/unresolved fills,
visibility and legacy/composed/other-kind fallbacks remain intact. Minor
normalization raises MAE/ratio by .063162/.002157 and is rejected: five candidate
minor levels differ from forty Office levels. Its geometry and existing dash
mapping remain partial. Failed composed Office export is candidate-only fallback
evidence. Base **b11614a8/root 08ec2606**; final **d40b7bc3/root 4515b4cc**.

A malformed-width probe demonstrates that initial inherited color could activate
a negative PDF width. Final source restores its prior fallback while preserving
valid zero-width hairlines. Exact final tests fail the pre-fix resolver on this
case. Four portable methods cover fourteen source compositions; no additional
parser/framework, font/size/shape exception or quota is introduced.

Twenty-four Office controls qualify five improved raster pages/19 exact/no
tradeoffs; width-only MAE falls 1.1282 to .9920. Only five grouped major-gridline
stroke states change; all 363 text operations/1966 path and clip geometries/50
embedded-font program checks and axis/tick/series ownership stay exact. Strict
Windows Release and 395 chart methods pass; thirteen Linux checks pass with two
archive SourceLink warnings. Exact final tests demonstrate four L238 failures/
five protected passes. All 252 public PPTX inputs and 158 retained contexts stay
exact. Frozen/root each reproduce 434 PDF check executions (24+252+158,
overlapping inputs); eight root DOCX guards remain exact. Three source blobs/raw
hashes match. Last full Windows remains L230; L238's 201 OOXML-resource passes are
prior evidence. Safe managed code, resource/work/publication limits and 0.1.5
release deferral hold. Minor intervals/coverage continue after bounded tick work L240, as L241.


### RV06-L240: bounded numeric axis tick expansion

Frozen **b51f06cf**, root **4872475a**, rejects a demonstrated non-progressing
axis-tick loop: min=1e16, max=1e16+4 and unit=1 cannot advance the represented
value. An isolated accepted-parent child enters the loop and is terminated after
two seconds; the fixed runtime rejects promptly. Dense generated ticks reuse
the existing 100000 chart expansion cap and cumulative chart-range budget before
allocation. Invalid/non-finite bounds and stalled additions reject through the
existing limit exception. Ordinary tick calculation and rendered output stay exact.

Five focused regressions cover dense admission, cumulative work, cap boundaries,
floating-point progress and conversion-level limit propagation. Strict Windows
Release and 400 chart/201 OOXML checks pass; fourteen portable Linux checks pass
with two archive SourceLink warnings. Exact frozen tests demonstrate one safe
expected parent failure and two protected tick passes. Frozen/root each reproduce
442 exact PDF check executions (32 unit/retained controls +252 public PPTX +158
retained contexts, overlapping inputs); eight root DOCX guards stay exact. Three
source blobs/raw hashes match. No new full catalogue, quota knob, framework or
heap accounting. Safe managed code, cancellation, existing resource/work/cache/
staging/publication bounds and 0.1.5 release deferral hold. Office source controls
continue to show minor-interval/major-level/boundary gaps; rendering work resumes
as L241 without transferring a parity claim.


### RV06-L241: line-chart minor subdivisions and source presence

Frozen **f3802d82**, root **17087483**, derives omitted minor units from the
selected major interval divided by five, with authored minor precedence. Sorted
bounded levels merge in linear work. Existing scene major-element presence owns
overlap exclusion: hidden/noFill majors reserve levels; absent majors permit
coincident minor lines. Horizontal painting has one shared owner. Other chart
kinds retain their prior path; the broad Cartesian prototype worsens a negative
column by .658452 MAE and two area cases, so it is rejected.

Forty-one Office controls qualify twelve changed PDFs: eleven improved raster
pages/29 exact/one new partial crossing tradeoff (+.002852 MAE/.000030 changed
pixel ratio). All 553 text operations/3328 non-minor graphics operations/88 font
program checks and stroke properties stay exact; only twelve grouped minor paths
change. Five portable methods include six emitted-PDF source compositions. Strict
Windows Release and 405 chart/201 OOXML checks pass; nineteen Linux checks pass
with two archive SourceLink warnings. The exact frozen emitted-geometry regression
fails the accepted parent and two ordinary tick guards pass. Public252/retained158
stay exact. Frozen/root each reproduce 451 qualified PDF check executions
(41+252+158, overlapping inputs); eight root DOCX guards and seven source blobs/raw
hashes match. No public gate/full-catalogue/parity claim is changed. L240 admission/
work bounds, safe managed code and existing cancellation/cache/staging/publication
constraints remain; release stays deferred at 0.1.5. Automatic major/label fit,
source origin/boundary/crossing and minor width/dash residuals continue as L242.

### RV06-L242: rejected automatic line-axis interval prototype

The general automatic-capacity prototype failed two independent height/font
holdouts before freezing source. A narrower prototype retained all label output
and shared the label interval only when selected font advance plus clearance fit
the plot. Among 112 Office controls, it changed 66 PDFs: 39 raster improvements,
46 exact outputs and 27 tradeoffs. Increasing grid density amplified existing
plot-geometry residuals, predominantly in Arial controls. No font-name or metric-
selection exception was introduced to evade these controls.

All 1618 text operations, 9733 non-grid graphics operations, 230 embedded-font
program checks and stroke properties remained exact; only 68 grouped grid paths
changed. Strict Release and 409 chart tests passed. Of 252 public PPTX checks,
only the intended line-span case changed; all 158 retained checks stayed exact.
The prototype was rejected before integration, with source/runtime/evidence
preserved under ignored `artifacts/plan-revision-20261005/rv06-l242/`. Production
remains at accepted L241. Category-label reserve ownership is the next
investigation; its suspected relation to the residual is not yet established.
No frozen/Linux/root acceptance, gate relaxation or parity claim is made.


### RV06-L243: authored category-label offset ownership

Frozen **cdeea3e7**, root **ec1acfca**, shares a measured offset adjustment between
automatic line plot reserve and category-label placement. Twelve independently
authored Office controls across Aptos12/18, Arial18 and Tahoma18 show reserve
increments of normal category line-height/4, with at most .00836pt residual and
stationary category baselines. Twelve mixed-axis controls separately establish
category ownership of the bottom strip and value-axis ownership of top clearance.
The broader default font-strip formula remains unverified and deferred.

All twelve affected raster pages improve with no tradeoffs. All 180 text encoding,
font, advance and X-position checks, 24 font-program checks and stroke properties
stay exact. Geometry is qualified by plot-affine movement, rigid marker/tick and
category-clip translations, and value-strip edges. All 124 default-offset controls,
252 public PPTX outputs and 158 retained outputs remain byte-identical.

Frozen/root strict Release and 408 chart/201 OOXML checks pass with zero warnings
or errors. Twenty-two Linux checks pass with two archive SourceLink warnings.
Each source reproduces 546 qualified PDF check executions (136+252+158, overlapping
inputs); eight root DOCX guards and seven source blobs/raw hashes match. Exact
frozen offset/baseline and selected-metric tests fail the parent; two ordinary
tick controls pass. Evidence is under ignored `artifacts/plan-revision-20261005/rv06-l243/`.
Safe managed code, existing bounds/work/cancellation/cache/staging/publication,
0.1.5 and release deferral remain. No full-catalogue, gate-relaxation or parity
claim is made. Source origins/endpoints continue as L244; L242 density rollout and
broader default font-strip metrics stay deferred.


### RV06-L244: linear line-axis source minima and endpoints

Frozen **ecd59d63**, root **692c225e**, shares one bounded source-minimum/interval
endpoint policy across line label measurement/emission, ticks and major/minor
grids. Supported source sequences match Office across positive, negative and
fractional minima, incomplete maxima, reversals and unit variants. Typed scene
preserves logarithmic-scale presence; unsupported log, implicit-minimum,
composed/secondary and normalized-minimum cases retain their fallback.

Twenty-four Office comparisons qualify twenty changed PDFs: seventeen raster
improvements, four exact outputs and three new partial negative-range tradeoffs.
MAE increases are .104690/.025613/.038319; threshold pixel-ratio increases are
.004235/.001576/.000208. These expose a pre-existing crossing-layout mismatch:
Office uses a lower plot reserve for interior crossings and places category
labels at the crossing/on the reversed side; candidate plot Y/category baseline
retain earlier behavior. No sign/font exception or public gate relaxation is added.

All 504 glyph outline/instruction/advance/side-bearing checks, 24 category-font
program identities, 232 text placements and 1704 stroke-state checks pass.
Numeric subsets may change with the corrected labels. Source grid/tick lattices,
label-strip width, marker/series and clip geometry are independently qualified.
All 252 public PPTX and 158 retained outputs stay exact; of 136 older controls,
only the intended offset7 case changes. A duplicated composed synthetic plot is
rejected by Office and has a candidate-only exact fallback proof.

Frozen/root strict Release and 412 chart/201 OOXML tests pass with zero warnings
or errors. Twenty-six Linux tests pass with two archive SourceLink warnings.
Each source reproduces 570 qualified PDF checks (160+252+158, overlapping inputs;
159 Office-informed control checks and one candidate-only). Eleven source blobs/
raw hashes and eight root DOCX guards match. Exact frozen emitted-source regression
fails the parent; two ordinary tick controls pass. Evidence is under ignored
`artifacts/plan-revision-20261005/rv06-l244/`. Safe managed code, existing resource/
work/cancellation/cache/staging/publication limits, 0.1.5 and release deferral remain.
No full-catalogue or parity claim is made. Crossing layout proceeds as L245.

### RV06-L245: category-axis source crossing and automatic line clearances

Revision **b07db2d9** shares category-axis crossing, label placement, clearances,
authored offsets and tick direction for the existing automatic untitled,
legendless, single-line admission. Source low/high labels stay at their edges;
nextTo labels follow the category crossing and reversed value direction.
Manual/composed/secondary, unsupported-scale, unresolved-font and small-layout
fallbacks retain their earlier behavior. Other renderer kinds are unaffected.

Thirty-nine Office comparisons qualify 36 raster improvements, two exact outputs
and one explicit partial tradeoff: MAE +.022156, threshold ratio -.000001. In that
out-of-range source, Office and candidate emit six horizontal major levels while
the parent emits five. The correct added level exposes an inherited axis-paint
difference: candidate black .75pt versus Office gray 1pt. No source-value/font
exception or relaxed public gate conceals this dependency. Axis/minor-stroke
style ownership and broader font-strip baselines remain partial.

All 78 embedded-font programs, 818 glyphs, 389 text states and 3235 paint states
remain protected; 2620 geometry changes have source-backed roles. All 252 public
PPTX and 158 retained PDFs remain exact. Eight changed earlier controls each
improve through independent Office comparisons.

Frozen/root strict Release and 417 chart/201 OOXML tests pass with zero warnings
or errors. Each source reproduces 601 exact PDF checks (191 controls, 252 public
and 158 retained, overlapping sets). Thirty-one Linux checks pass; the isolated
source-only archive build retains two SourceLink warnings. Seven source blobs/raw
identities and eight DOCX guards match. Exact frozen tests fail the parent at four
intended regressions; three protected guards pass. No new full catalogue or
complete Office parity is claimed. Version 0.1.5, release deferral, safe managed
code and existing resource/work/publication limits remain.

### RV06-L246: shared modern line-axis stroke inheritance

Commit `23292d0e` fixes ignored width-only axes and black0.75pt modern-gallery
fallbacks. Axes and major gridlines share one omission reader; the typed scene
publishes the final axis stroke. Source overrides/noFill keep precedence. The
existing single-line102/118/126 gallery admission stays bounded; classic,
composed, other-gallery and explicit-style-part fallbacks remain.

Eighteen independent Office controls plus182 affected existing checks qualify
197 lower-mean-error comparisons and3 exact guards.126 improve both metrics;
71 have threshold-only tradeoffs with lower mean error.394 axis paint states
match Office. All17448 graphics geometries,2613 complete text operations,
400 embedded font programs and4544 glyphs remain exact against the parent.
The3155 paint changes are orthogonal axis/tick strokes. Existing vertical-axis
crossing residuals and minor-stroke defaults remain separate; no geometry or
font exception was added to conceal them.

Frozen/root strict Release and422 chart/201 OOXML tests pass without warnings or
errors. Each reproduces619 exact PDF checks (209 controls,252 public,158 retained;
sets overlap). Eight DOCX guards remain exact. Thirty-six portable Linux checks
pass; the source-only archive build retains two SourceLink warnings. Four source
blobs/raw identities match; frozen tests fail the parent at four intended
regressions and retain the fallback guard. No new full-catalogue or complete
Office-parity claim; version0.1.5 and release deferral remain.

### RV06-L247: multiple source-owned footnote markers

Commit `ec21bc51` fixes per-page numbering when a paragraph contains several
generated footnote markers. The previous guard left labels105..108 where Word
restarts4/5 or09/10 per physical page. An indexed source-run/actual-segment map
and shared run qualifier use one explicit ownership scope. Numbering can own
every marker; kerning, leading and separator-reserve corrections retain their
single-reference admissions. A single numbering retry must retain all complete
marker/note pages. No font/size/threshold exception is added.

Twenty independent Word controls cover source IDs/starts, decimal/decimalZero,
break forms, Arial/Calibri/Times,14pt text and wrapped/met/mixed frames. Eighteen
corrections improve both raster metrics on all36 admitted pages. Two raw Calibri
wrapped frames retain exact parent output: their pair flags do not match source
thresholds under the existing normalized-frame policy. Those partial fallbacks
remain explicit; malformed/shared source owners, incomplete/moved notes and
unstable pages retain their existing exclusions.

All624 authored text operations and advances,160 marker operations,64 source
font-program checks,1718 shared glyphs and160 graphics operations are protected.
Every horizontal position follows exact source-hmtx marker advances on the
existing half-point superscript grid;784 projections have maximum error below
.000954pt. All vertical positions/matrices and80 distinct same-page note
destinations/vertical hit areas stay exact. Source digits and metadata remain
intact; only generated marker labels and their width-dependent positions change.

Frozen/root strict Release builds and1255 unique affected DOCX/balloon methods
pass (1253 DOCX and122 balloons, overlapping coverage). Each reproduces647 exact
PDF checks:619 inherited PPTX,20 independent DOCX and8 inherited DOCX guards.
Thirty-one portable Linux checks pass; the isolated source-only build retains
two SourceLink warnings. Seven source blobs/raw identities match. The exact
frozen rendering test fails the parent at its visible label assertion. No new
full-catalogue or complete Office parity claim. Version0.1.5, release deferral,
safe managed code and existing bounds/work/publication limits remain.

## RV06-L248: complete owners for wrapped-marker kerning

Runtime **d31c0b67** reuses the complete source-marker ownership scope for the
existing wrapped Calibri pair-normalization retry and each-page numbering.
Existing font, size, threshold, leading and reserve admissions stay in force.
Fourteen independent Word controls cover raw default/zero/unmet/met/mixed
thresholds, early/later markers, spacing and decimalZero labels. Thirteen
corrections improve both raster metrics on all26 affected pages; the met guard
stays exact. Source-authored words/order/pages,28 font programs,1064 shared glyphs,
112 graphics and56 distinct same-page note links/vertical hit areas are protected.

One qualified artifact is reused. Its prior1257-method prototype sweep and647
PDF conversions remain valid:645 outputs stay exact and the two changed inherited
dependencies have fresh Word qualification. Twenty-eight capability-specific
checks pass. Root transfer verifies398 source identities and17 identical build
assets; four CLI smoke conversions match frozen PDFs without another root build,
suite or corpus pass. Git checkout line endings explain65 raw source-hash
differences; Git blobs and canonical source bytes match. Strict Release has zero
warnings/errors. L247 provides the recent broader/Linux milestone; no new
cross-platform/package qualification is claimed here. Visual thresholds,
version0.1.5, release deferral and existing safety/work/publication limits remain.

## RV06-L249/L250: source stroke concern and upper-bound interval ownership

The L249 minor-stroke prototype is retained outside the root. It repairs a
confirmed dropped authored width, but four manual/log/missing-bound layouts
regress on both raster metrics among73 affected comparisons. Sixty-nine improve;
all prior parent measurements and converted PDFs are reused. Thirteen independent
Office controls give nine improvements/four exact guards; malformed negative
width is a candidate-only fallback. Those results do not qualify a wider rollout.

Runtime **95bbe763** resolves the shared missing-lower-bound defect independently.
One linear line-chart axis with a finite authored maximum/positive major interval
anchors its automatic minimum to that interval lattice. Maxima107/103 with
interval20 produce-13/-17. Measurement, emitted labels and subdivisions share
the resolved origin. Sixteen Office controls, including fractional and negative
ranges, qualify eleven improvements/five exact/no tradeoffs. Changed numeric
labels match Office. Source font metrics/hinting remain exact in32 checks;332
shared glyphs and22 newly used glyphs retain source outlines/advances. Different
generated inventories can change subset bytes. All64 authored text states remain
protected.261 minor intervals match Office counts/relative spacing within its
verified0.12pt interior-coordinate precision, with maximum normalized residual
.000360 and non-grid extrema recorded. Absolute parity remains incomplete.

One artifact passes25 selected Windows checks and17 Linux checks using identical
DLLs, with no Linux rebuild. Fourteen freshly identical PDFs reuse Office evidence;
twelve specific inherited/neighbors leave ten exact and two independently
qualified changes. Two further negative controls use the same artifact. Root
transfer verifies398 sources/17 identical assets and four CLI smoke PDFs without
another build or broad suite/corpus. Strict Release has zero warnings/errors.
Existing thresholds, version0.1.5, release deferral and safety/work/publication
limits remain. Manual/log geometry and the rejected minor-stroke rollout stay
explicit; no new full-catalogue or package qualification is claimed.

## RV06-L251: empty manual-layout metadata

Runtime **c7a06a01** distinguishes raw manual-layout presence from authored
coordinates. Empty nodes and target/mode metadata preserve automatic line-plot
measurement. Explicit/partial/malformed coordinates retain their existing path;
raw source metadata stays inspectable. Twelve independent Office controls qualify
eight improvements/four exact guards across metadata, font/frame widths and
source-bound/reversed axes. One malformed coordinate is candidate-only exact.
All131 text states,26 font programs/275 glyphs and1081 paint/path-topology states
are protected. Empty metadata produces exact automatic PDFs. The single changed
inherited source is canonically equivalent to the new control, with identical
independent Office page rasters; three neighboring PDFs stay exact.

One frozen artifact passes19 capability-specific checks. Root transfer verifies
398 source identities/17 assets and four CLI smoke PDFs without a rebuild or
repeated suite/corpus. Strict Release has zero warnings/errors. L250's same-DLL
Linux checks and L247's broader corpus remain the recent milestones; no new
full-catalogue/platform/package claim is made. Existing thresholds,0.1.5/release
deferral and safety/work/publication limits remain. L249 remains rejected pending
reassessment of linear-scale source paint; logarithmic geometry remains separate.

## RV06-L252: linear-scale minor-stroke ownership

Runtime **3887bd3c**, with harness correction **c3d611a0**, reuses the shared
line-property resolver for supported linear line axes. Existing typed log-scale
presence retains its unsupported geometry fallback. L250/L251's shared geometry
fixes allow the earlier source-width concern to qualify without chart-ID, font,
size or sign exceptions. All72 changed inherited comparisons improve both raster
metrics;69 comparisons and13 independent Office controls reuse evidence through
fresh exact candidate/parent PDF identities. Only three updated geometry parents
need fresh comparison. Those three also improve. Authored fields/noFill/invalid/
other-gallery/composed behavior and the log guard are protected. The source
hairline0 versus Office0.14pt residual remains explicit.

Twenty-seven targeted tests pass. The related axis/layout milestone qualifies432
chart methods and252 exact public PDFs on one library artifact. It exposed a
stale reflection invocation and flat-bundle fixture lookup: the affected test
method was corrected, its runner rebuilt without rebuilding the library,430
completed unchanged checks reused and two failed methods retested. Thirty-four
Linux checks use the same DLLs. A trailing-CR shell footer failed after all reports
and identity checks completed; those results are verified and script syntax fixed
without another run. Root transfer verifies398 sources/17 build assets and four
CLI smoke PDFs, with no root rebuild or repeated suite/corpus. Strict Release
passes; thresholds,0.1.5/release deferral and existing safety/work/publication
limits remain. No new full-catalogue or package claim. Workspace2.820GB/no errors.

## RV06-L253: authoritative section note-numbering requests

Runtime **8168e939** shares the supported restart vocabulary between reader and
renderer admission. Known document eachPage/eachSect metadata preserves an
explicit section eachPage request; document-only metadata remains separate.
Twenty-one Word controls across Arial/Calibri/Times and11/12/14pt bodies qualify
twelve corrections/24 improving pages and nine exact guards/18 pages. Generated
body/footer labels and source note pages match Word. Unknown metadata, unsupported
document shapes and stable-owner retries retain fallback; no font/size/threshold
exception or reader cascade change is introduced.

Seventy source font-metric/hinting checks,1806 shared glyphs,588 authored text
states,168 graphics and84 distinct complete note destinations/vertical hit areas
are protected. Fifty-three unique targeted methods pass;31 inherited/neighbors
stay exact. One artifact transfers398 source identities/17 build assets and
matches four root CLI smoke PDFs without another build or broad suite/corpus.
Strict Release has zero warnings/errors. L252 retains the recent broader/platform
milestone; no new full-catalogue/platform/package claim. Existing thresholds,
0.1.5/release deferral and safety/work/publication limits remain.


## RV06-L256: declared static frames and visible footer reservation

Runtime **b933fbde** uses the existing line-height profile for resolved static
exact/minimum frames. Exact pictures retain source extents and carry a clip;
footer translation and scaled-story placement keep its origin with the image.
The visible footer envelope also bounds body flow. Unresolved metrics retain
their explicit fallback. Thirty independent Word controls qualify29 primary
raster improvements and one exact guard; eight dense controls match Word body
page assignments. Forty font-metric/hinting checks,1032 decoded glyphs and1108
horizontal text/spacing states remain protected. Two earlier fixed/mixed guards
receive fresh Office comparisons and now improve. Small baseline residuals and
the independent SVG isolation/radial gaps remain open; thresholds are unchanged.

One frozen artifact passes48 unique capability-specific methods and63 exact PDF
identity checks. Matching Office inspections/rasters are reused. Root transfer
verifies399 sources,14 fixtures and17 assets, with four matching CLI smoke PDFs
and no root rebuild or repeated suite/corpus. Strict Release has zero warnings
or errors. L255 retains the recent same-DLL Linux/package milestone and L252 the
broader corpus milestone. Version0.1.5, release deferral, safe managed code and
existing bounds/cancellation/work/cache/staging/publication limits remain.


## RV06-L257: centered-autofit qualification correction

CI exposed one stale regression that required a different centered word break
and treated a preserved trailing space as visible overflow. Independent
PowerPoint output confirms the existing two lines and centered visible text.
Test correction **f3120ff8** checks visible-atom bounds and equal margins on
every line, plus preservation of the trailing-space advance. Rendering and
visual thresholds are unchanged. Five focused checks pass using the accepted
L256 library, with no library rebuild. Verified source/tree identities, seven
transferred test assets, the root catalogue and an exact CLI PDF smoke qualify
the transfer without repeating the suite or corpus. The CI delivery milestone
runs separately on Windows and Ubuntu.


## RV06-L258: DOCX source cropping and negative padding

Runtime **97448de8** retains authored source rectangles and maps source paint
inside the existing viewport across SVG, raster and preview fallback. Layout
extents stay unchanged; unusable rectangles retain the uncropped image with a
specific diagnostic. Ten independent Word controls improve both primary raster
metrics. Ten font programs,210 decoded glyphs and42 full authored text/font/
encoding/kerning/geometry/page states remain exact. Six uncropped neighbors keep
identical PDFs. Rotation/flips and VML cropping remain separate.

One frozen artifact passes29 specific methods. Root transfer verifies399 sources,
18 fixtures and17 assets, with four identical CLI smoke PDFs and no rebuild or
repeated suite/corpus. Strict Release has zero warnings/errors. Four new public
cases validate in the462-case/11-family catalogue. Three meet retained gates;
cropped exact-footer remains needs-review atSSIM0.947083/recall0.944162 against
0.985/0.99. Existing isolation/radial gaps remain open. Workspace3.703GB is below
the loose target. CI's preceding0595336c milestone passed both platforms,2674
Windows and417 portable Ubuntu methods; the new push runs CI during further
rendering work. Safe managed code,0.1.5/release deferral and existing bounds,
cancellation, cumulative work, bounded caches/staging and publication remain.


## RV06-L259: authored DOCX picture flips

Runtime **0b0a49b6** retains horizontal/vertical picture flags and applies one
viewport-centered mirror after source crop across vector/raster/preview paint.
Invalid crop fallback preserves the flip; source layout extents and static-line
clips remain unchanged. Twelve independent Word controls improve both primary
raster metrics. Twelve font programs,248 decoded glyphs and50 full authored
text/font/encoding/kerning/geometry/page states stay exact. Sixteen earlier crop
and neighboring PDFs remain identical. Rotation and VML crop/flip stay separate.

One frozen artifact passes32 specific methods. Root transfer verifies399 sources,
22 fixtures and17 assets, with four exact CLI smoke PDFs and no rebuild or
repeated suite/corpus. Strict Release has zero warnings/errors. Four new public
cases validate in the466-case/11-family catalogue. Three meet existing gates;
flipped exact-footer remains needs-review atSSIM0.961155/recall0.978416 against
0.985/0.99. Earlier residuals and all thresholds remain. The preceding crop CI
milestone0534b512 passed both platforms,2677 Windows and417 portable Ubuntu
methods plus solution/package/tool checks. Rendering continues during CI for
the new push. Version0.1.5/release deferral, safe managed code and existing
bounds/cancellation/work/cache/staging/publication limits remain.


## RV06-L260: source-frame rotation and image movement ownership

Runtime **bf8d470f** fits stored picture dimensions uniformly inside the flow
extent, top-left aligned, then rotates/crops/flips shared vector/raster/preview
paint. Paragraph line clipping includes measured image growth and consumed
final after-spacing; anchors retain full rotated viewports. Image movement now
owns both clips and line frames across placements. A short-text-slot prototype
hid pictures and was rejected despite improving primary metrics. Unresolved
source/line frames retain an explicit unrotated fallback, without Office parity.

Twenty-one independent Word controls qualify20 primary-metric improvements and
one exact missing-frame fallback. Twenty-one font programs,439 decoded glyphs
and98 full authored text/font/encoding/kerning/geometry/page states are protected.
Twenty-eight earlier crop/flip/neighbors remain exact. One frozen artifact passes
46 capability-specific methods. Root verifies399 sources,27 fixtures and17 assets
plus four matching CLI smoke PDFs, with no rebuild or repeated suite/corpus.
Strict Release has zero warnings/errors. Five new cases validate in471 cases/11
families: three meet retained gates, rotated table/exact-footer remain needs-review.
Earlier SVG gaps and all thresholds remain. The preceding3ea43bac CI milestone
passed Windows2680/portable Ubuntu417 methods, package and tool checks. Rendering
continues during CI for this push. Safe managed0.1.5/release deferral and existing
bounds/cancellation/work/cache/staging/publication limits remain. Retirement of
13 clean historical worktrees reduces the workspace from3.924GB to2.958GB while
preserving source revisions, inputs, frozen binaries and external evidence.


## RV06-L261: cell-owned rotated picture clipping

Runtime **4e209958** uses the existing cell content clip for rotated pictures,
including aligned/repeated paragraphs. The line-frame origin reuses the measured
cell baseline inset. Eight Word controls improve similarity/foreground recall;
seven also improve both primary metrics. Bottom alignment's unchanged11.46pt
position offset retains MAE+0.005318/threshold16 ratio+0.000267, while SSIM improves
0.800331->0.845339 and recall0.707421->0.856607. The remaining alignment defect is
explicit, with no font/size/bottom-specific exception. Sixteen font programs,
336 glyphs and52 complete authored text/encoding/kerning/geometry/page states
remain exact. The public rotated-table case now meets every retained gate:
MAE0.049681/SSIM0.993828/recall0.998959. Its thresholds are unchanged.

One frozen artifact passes19 specific methods and48 unchanged neighboring PDF
identities. Root verifies399 sources,27 fixtures and17 assets plus four exact CLI
smoke PDFs without rebuilding/repeating the suite or corpus. Strict Release has
zero warnings/errors. The initial paragraph-only clipping prototype is retained
as rejected evidence. Arial11, aligned-position and prior SVG residuals remain
open. The preceding2f8e730b CI milestone passed Windows2684/Ubuntu417 methods and
source Release/package/tool gates; rendering continues during CI for the push.
Safe managed0.1.5/release deferral and all existing bounds/work/publication remain.


## RV06-L262: shared cell vertical-alignment space

Runtime **3abbe52a** shares padded-cell free-space ownership between text and
drawing paths. Natural/exact slots account their baseline inset in used height;
minimum/mixed-leading frames retain separately reserved headroom. A blanket
inset-removal prototype was rejected after13 prior cases regressed. Thirty-six
Word controls qualify16 improvements/20 exact guards, preserving every earlier
public aligned-cell minimum/margin/border/grid output. Eighty-four font programs,
2384 glyphs and594 horizontal text/font/encoding/kerning/page states remain exact.
Seven Office PDF byte variants retain identical rasters/text geometry. Prior
clipped text-operation differences remain historical partial scope.

One frozen artifact passes94 capability methods and55 additional exact PDFs.
Root verifies399 sources,31 fixtures and17 assets plus four matching CLI smoke
PDFs with no rebuild or repeated suite/corpus. Strict Release has zero warnings/
errors. Four new cases validate in475 cases/11 families: two picture gates pass,
two text foreground-recall0.982529 cases remain needs-review against0.99.
Minimum-leading image placement and other small typography/SVG gaps remain open.
No thresholds or font/size/document-ID qualifiers change. Prior128d778f CI passed
Windows2686/portable Ubuntu417 methods plus source Release/package/tool gates;
rendering continues during CI for the push. Safe managed0.1.5/release deferral
and all existing bounds/cancellation/work/cache/staging/publication remain.


## RV06-L263: minimum-leading cell picture ownership

Runtime **95da5082** shares minimum-slot placement with source-affined pictures.
Picture growth occupies minimum headroom once; line displacement, following
paragraphs and alignment share that extent. Source ambiguity and established
header/rich/revision/split/row-rule fallbacks remain explicit. The initial
text-only-height prototype **ec3e32dc** was rejected after seven top/tall controls
regressed despite passing46 methods and32 exact neighbors.

Forty-five independent Word controls qualify32 improvements/13 exact/no
regressions. Ninety font programs,1899 glyphs and285 horizontal text/font/encoding/
kerning/page states remain protected. First baselines differ by at most0.062pt.
The automatic tall-picture control retains its9.262pt following-body discrepancy;
this qualification preserves that measurement fallback.

One frozen Release artifact passes46 specific methods and32 prior exact PDFs.
Root verifies399 sources/35 fixtures/17 assets and four matching CLI smoke PDFs,
with no root rebuild or repeated suites/corpora. Four public minimum-picture
cases pass the existing0.2 MAE/0.004 ratio/0.985 SSIM/0.99 recall gates using the
qualified outputs;479 cases/11 families validate. The final primary regression
fails against the pinned parent while two guard methods pass. Safe managed0.1.5,
release deferral and all existing bounds/work/publication remain. Rendering
continues while the accepted push runs CI; prior7ec86425 is green on both platforms.


## RV06-L264: shared minimum-picture row extent

Runtime **190f658d** returns full measured cell height and minimum-picture overlap;
existing ownership and complete-frame rules select the occupied height. Line
placement and measurement share the finite overlap resolver. Fixed rows, natural
floors and existing rich/header/fragment/keep/nested fallbacks remain explicit.

Thirty-two independent Word controls qualify28 improvements/four exact/no
regressions across fonts, minimum heights, picture sizes, alignment, row rules,
grids, wrapping, two paragraphs and margins. The9.262pt following-body error
becomes at most0.090pt; first baselines differ by at most0.020pt. Sixty-four font
programs,1353 glyphs and211 horizontal text/font/encoding/kerning/page states stay
protected. Twenty-seven selected prior PDFs remain exact.

One frozen Release artifact passes47 specific methods. The occupied-extent
regression fails the pinned parent while three ownership/fallback guards pass.
Root verifies399 sources/39 fixtures/17 assets and four matching CLI smoke PDFs,
without rebuild/repeated suites/corpora. Four public automatic/content/grid/wrapped
cases pass unchanged0.2 MAE/0.004 ratio/0.985 SSIM/0.99 recall gates using qualified
outputs;483 cases/11 families validate. Safe managed0.1.5 and all bounds/work/
publication remain. Prior9a908751 CI passes2691 Windows/417 Ubuntu methods plus
source Release/package/tools; rendering continues during the new push's CI.


## RV06-L265: kept-first-row admitted table frame

Runtime **17dab9cf** replaces duplicated first-row estimation with the existing
table-frame result after the kept chain. Available height, grid admission,
complete/fragment/header/rich/scaled behavior and cancellation share body layout.
Nineteen independent Word controls qualify five improved heading/picture page
assignments and14 exact/no regressions. Every changed role/page matches Word;
the104pt two-row difference remains an exact inherited fallback.

Forty-four embedded-font/metric checks,1162 glyphs and198 horizontal text/font/
encoding/kerning states stay exact. A page move can remove a duplicate font
reference; actual embedded programs and glyph maps remain identical. One frozen
Release passes67 specific keep/frame/minimum-image/page-boundary methods and14
prior exact PDFs. The primary regression fails the parent; fallback passes.
Root verifies399 sources/43 fixtures/17 assets and four matching CLI smoke PDFs
without rebuilding/repeating suites/corpora.

Three new kept-picture cases pass unchanged0.2 MAE/0.004 ratio/0.985 SSIM/0.99
recall gates. The margin case's first-page SSIM0.982089 stays needs-review against
0.985, despite corrected page ownership;487 cases/11 families validate. Safe
managed0.1.5, release deferral and all bounds/work/publication remain. Prior
3bbfc05f CI passes2692 Windows/417 Ubuntu methods plus source Release/package/
tools; rendering continues while the accepted push runs CI.


## RV06-L266: gallery2 axis fallback with style parts

Runtime **0c90245a** removes an unrelated style-part presence exclusion from
gallery2 axis fallback. Authored/no-fill strokes, existing plot/gallery guards
and major-gridline role ownership stay intact. Nine independent PowerPoint
controls qualify three improvements/six exact/no regressions. Four axis strokes
change from black0.75pt to Office gray0.537/1pt;499 path/clip/fill/decorations,
37 embedded-font/metrics,404 glyphs and145 full text-state checks remain exact.
PowerPoint ignores modified explicit axis role colors in these valid controls;
role inheritance stays deferred.

One frozen Release passes25 specific chart methods/eight exact neighbors. The
integration regression fails the parent; explicit/no-fill guard passes. Root
verifies399 sources/45 fixtures/17 assets and four matching CLI smoke PDFs, with
no rebuild/repeated suite/corpus. Two public individual composite-chart gates
pass unchanged; their broader family SSIM0.995 remains partial.489 cases/11
families validate. Invalid namespace/child-order helper outputs and an Office-
rejected line-only transformation are retained/excluded from fidelity evidence.
Safe managed0.1.5/release deferral and all bounds/work/publication remain. Prior
d73cd072 CI passes2694 Windows/417 Ubuntu methods plus source Release/package/
tools; rendering continues during the accepted push's CI.


## RV06-L268: shared default axis font size

Runtime **e2138927** replaces separate9pt/8.5pt category/value defaults with one
nominal10pt default (9.96pt after existing chart quantization). Source/style/
element sizes retain precedence. Sixteen independent PowerPoint controls across
frames, galleries, fonts, omitted/explicit sizes, bar/line and intervals qualify
ten improvements/six exact/no regressions.32 embedded-font/metric checks/368
glyphs/186 text contents remain exact. Default size and its measured layout
change together; explicit-size PDFs stay identical.

One frozen Release passes64 specific text/axis/interval/override methods/eight
exact neighbors. Primary default-size regression fails parent; source-size guard
passes. Root verifies399 sources/49 fixtures/17 assets and four matching CLI PDFs
without rebuilding/repeating suites/corpora. Source-only public inventory finds
zero previous fixtures affected, so no public corpus is replayed. Two new
individual gates pass; dense ratio0.033465 against0.025 and bar geometry/recall
stay needs-review with all thresholds unchanged.493 cases/11 families validate.
Safe managed0.1.5/release deferral and bounds/work/publication remain. Prior
3f2acc44 CI passes2696 Windows/417 Ubuntu methods plus source Release/package/
tools; rendering continues during the accepted push's CI.


The PDF inspector decodes sequential `ToUnicode` ranges containing UTF-16
surrogate pairs as distinct characters. A retained Word matrix exposed the former
decoder repeating the first mathematical letter for a/b/c/d; the correction
increments the target before decoding the complete UTF-16 string. Three focused
CMap regressions cover astral ranges, BMP ranges, multi-character strings and
explicit arrays. `OOXPDF_TEST_PDF_INSPECTOR` optionally pins a frozen inspector
assembly for those tests; ordinary test runs retain the shared tool-build path.

Six retained Word matrix PDFs verify corrected text while PDF bytes, all text
coordinates/spacing, fonts and graphics remain exact. Earlier inspection reports
may repeat the initial character of a sequential astral range. Re-inspect an
affected retained PDF when its decoded text is needed; existing raster, glyph-
outline and coordinate evidence does not require a broad corpus replay.

`ExportCompositionReferences.ps1` now selects inline mode 1 explicitly for review
and pins all six view settings, the observed ExportSettings line and export-log
hash. It preflights cached inputs/PDFs and observed settings before any Office
invocation; legacy pins require a retained log proving the requested view. A
mismatched cache is rejected for inspection, without overwriting references or
starting a global re-export. Retain historical view-agnostic comparisons by their
verified input/PDF identities; correct only references affected by the view.
The meanings of [Word's revision modes](https://learn.microsoft.com/en-us/office/vba/api/word.wdrevisionsmode)
are documented by Microsoft. Tool regression coverage uses fake inputs, with no
Office invocation or library build.

Use `BuildQualificationArtifact.ps1 -WorktreeDirectory <absolute-linked-worktree>
-Revision <full-commit> -OutputDirectory <absolute-artifact-directory>` to build
and freeze a committed batch. It rejects the main checkout, dirty/different source
and reused output directories before building, then uses absolute test/CLI project
paths regardless of the caller's current directory. Source changes or build
failures prevent freezing. Existing freeze/transfer commands retain ownership of
asset/input checks and the small transfer smoke; no tests run during build/freeze.
`TestQualificationBuild.ps1` exercises these guards using a tiny Git repository
and fake commands, including paths with spaces, without compiling the library.
