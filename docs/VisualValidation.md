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
