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
`RevisionsMode=1` (bespoke `ExportDocxMarkupReference.ps1` knobs; the remaining
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
recall >=.56, plus page/dimension/diagnostic and changed-pixel checks. Regenerating
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
