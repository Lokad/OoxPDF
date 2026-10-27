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

The affected-test gate rejects skipped methods unless their exact names are
declared with `-ExpectedSkip`. Declare only understood platform preconditions;
reports retain skip reasons and counts. An expected skip never replaces a failed
check when joining repaired reports. Undeclared skips and failures retain their
reports and reject qualification.

For qualification, bind the case or family to the already qualified CLI:

```powershell
pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/docx-basic-paragraphs/case.json -CliPath <frozen-cli.dll> -SkipBuild -CacheOnlyReference
pwsh tools/CheckVisualFamily.ps1 -Family <name> -CliPath <frozen-cli.dll> -SkipBuild -CacheOnlyReference
```

Without `-SkipBuild`, a supplied CLI still prevents a library build; the family
builds independent validation tools once. Without `-CliPath`, ordinary development
runs build the default Debug CLI. `-SkipBuild` requires the selected runtime and
validation tools to exist. Family reports pin CLI/library hashes and reject a
runtime that changes during the batch.

PDF inspection follows the same build intent: a family prepares the inspector
once only when a selected text, graphics, chart or font gate needs it. Direct
cases prepare it once on first use; `-SkipBuild` requires the prebuilt inspector
and never invokes MSBuild. Each inspection records its actual tool hash.

For line-start gates, `maxTextLineStartDelta` bounds X. An explicit
`maxTextLineStartYDelta` also bounds matched baselines in Y; it is passed to the
helper as `BaselineTolerance`, separately from operation grouping. X-only gates
keep their existing meaning. Nonfinite coordinates or tolerances fail the gate.
Run `pwsh tools/TestTextLineStarts.ps1` for focused controls without builds or Office.

`RenderReference.ps1` and `ExportCompositionReferences.ps1` also accept
`-SkipBuild` to reuse the existing rasterizer. `RenderCachedReference.ps1` forwards
it on cache misses as well as derivative renders. Without the switch, the normal
build-freshness check remains. PDFium pin verification runs in either mode.
The reference progress log names the active operation before it starts, so a
supervised timeout identifies activation, open, export, cleanup or rasterization.

The harness converts the input, reuses a trusted Office PDF (or exports one when
explicitly requested), rasterizes both PDFs through the pinned PDFium, compares
pages and writes an assessment template. Reuse reference comparisons by exact
input/options/PDF identity. Do not duplicate frozen and root qualification.

`CompareCompositionPortfolio.ps1` reuses complete page measurements after matching
PDF and threshold-source identities. It derives failures from those measurements
and the current limits; cached `Failures` and `Thresholds` are not authoritative.
Missing fields, invalid numbers, broken page sequences and disagreement with the
candidate's `InspectedPages` count refuse reuse during batch preflight. VisualDiff's
null recall for a blank reference is inapplicable; incompatible dimensions and
missing pages retain their dimension/pagination failures. Keep incomplete attempts
as evidence and select a complete comparison or perform a fresh comparison.

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

Keep visual thresholds unchanged. Office and PDFium can differ in antialiasing, font hinting and rasterization details; record those effects alongside the failed gates. A human assessment describes practical severity and does not turn a failed threshold into a pass:

- `5`: visually indistinguishable except tiny antialiasing differences.
- `4`: good; minor spacing, kerning, or antialiasing differences only.
- `3`: usable; visible but non-blocking differences.
- `2`: poor; major layout defects or missing important elements.
- `1`: severe; page or slide mostly wrong.
- `0`: conversion failed or page missing.

Do not commit generated visual artifacts unless they are intentionally small fixtures.

`tools/ValidateVisualCases.ps1` validates manifest integrity and targeted public
fixture structure. Passing it does not establish rendering fidelity. Use the
current capability catalog rather than historical manifest totals.

## Family Parity Targets

Family and case manifests declare their visual gates and classifications.
Use `CheckVisualFamily.ps1 -Family <name> -List` to inspect the selected corpus.
A supported label requires the applicable unchanged gates and source/visual review;
manifest coverage and a successful conversion alone are insufficient.

See [Capabilities.md](Capabilities.md) for family constraints and
[PreviewRelease.md](PreviewRelease.md) for the finite preview gate. SVG focal
fills follow Office-visible centered gradients; smooth transparency follows the
Office PNG preview where its PDF export is defective. Charts, typography and
DOCX markup retain explicit approximations. Do not infer parity from a family name.

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

`RunDocxMarkupReferenceGate.ps1` and `CompareCachedDocxMarkupReference.ps1` accept
`-CliPath <frozen-cli.dll> -DocxInspectorPath <matching-inspector.dll> -SkipBuild`.
Supply both OOXML runtimes together; their library hashes must match. The shared
qualification artifact includes the Release DOCX inspector. `-PdfInspectorPath`
selects an independent prebuilt PDF inspector. Explicit paths avoid root builds;
`-SkipBuild` also reuses the other existing validation tools. Runtime hashes are
recorded and unexpected changes refuse qualification.

The three inspector wrappers accept `-InspectorPath` and `-SkipBuild`; DOCX/PPTX
inspection also accepts `-ExpectedLibrarySha256`. Default development invocations
still use the shared MSBuild gate. `TestFrozenMarkupBinding.ps1` checks explicit
binding, identity and refusal without a real build or Office export.

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

Word-compatible all-markup rendering follows Word 16's mixed view:
insertions and deletions render inline with revision styling while comments and
formatting revisions balloon. A reference exported with RevisionsMode=0 instead
balloons deletions (`Deleted:` title rows) and hides them from body lines, so
the gate compares different views: body wraps, op pairings, and page flow all
shift for non-layout reasons. Export word-compatible all-markup references with
`RevisionsMode=2`. Pin and verify revision visibility, revision/markup modes,
comments, formatting changes and insertions/deletions through
`ExportDocxMarkupReference.ps1`; record the full ExportSettings string. Word 16
accepts and reads back mixed mode 2 despite older enum documentation describing
it as unsupported. Mode 1 keeps text edits inline but suppresses formatting
balloons, so it does not match this profile when formatting changes survive.
Clean Final exports use `MarkupMode=0`, `RevisionsMode=0` and all visibility
flags disabled. Word resets `MarkupMode` to 0 when the last visible category is
hidden, so requesting all-markup mode 2 with those flags cannot pass readback.
Visibility is applied before display modes; the complete observed state is
still verified. Retain historical reference provenance and PDF identity rather
than relabeling old exports or rerunning the corpus.

Verify the view before importing: a mixed reference carries deleted
text inline at body size with no `Deleted:`/`Inserted:` title operations, while
a RevisionsMode-0 reference moves that text into balloon rows. Do not re-export
an entire corpus: retain historical view provenance and reuse comparisons when
equivalence is demonstrated for the input. Deletion-only and comment/author
controls without formatting changes have identical mode-1/mode-2 rasters.
Precedents: margin-dense-revisions and margin-landscape were re-imported
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

## Current status and history

See [preview qualification](PreviewQualification.md), [preview scope](PreviewRelease.md)
and [composition history](CompositionCheckpoint.md) for accepted outcomes and
explicit residuals. Fourteen fixed-portfolio DOCX strict gates remain failed;
manifest validity counts are separate from Office agreement. Keep default Final
and Word-compatible AllMarkup outcomes distinct.

The former per-slice validation journal remains in
[Git history](https://github.com/Lokad/OoxPDF/blob/78e7589ef3f0ddc9b1fbc38c94fa6743fd7fde16/docs/VisualValidation.md).


Local qualification uses capability-specific selections in
`tools/qualification-capabilities.json`. The reviewed single-math annotations,
display groups and review-body frame capabilities select small sets with neighboring
behavior. Their complete preceding sets remain under the corresponding `milestone-`
names; use those at broad checkpoints, not for every local edit.

`CompareCompositionPortfolio.ps1` reuses only completed `portfolio-raster/`
derivatives with matching PDF, DPI, rasterizer and complete page hashes. It leaves
legacy `raster/` evidence untouched. Stale/incomplete derivative directories refuse;
use a fresh owned run directory after investigating them. Identical PDF/threshold
comparison reports can still be reused without rasterization. Office thresholds and
residual failures stay unchanged.
