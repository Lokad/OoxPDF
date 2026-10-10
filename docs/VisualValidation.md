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
