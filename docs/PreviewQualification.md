# 0.2.0-preview qualification

The bounded admitted-content/style-authority milestone is qualified.
Publication, performance optimization and feature expansion remain deferred.
Library/CLI build: `5120449747b3947d6e5064816ba40555a24ac522`.
Test-only repair: `3284b7fdb0f4b19de2eb448e31040bf8065cde22`.
Library SHA-256: `A5FF10BAC4A5EA19C1CA1AAD4FF3C63B1137CE29F54DC5E99E6A3683B97FE0A0`.
The final tool-only follow-up leaves all production source and frozen assets
unchanged. [Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40533386646)
passes at that tool-only follow-up. CI builds are separate from the frozen runtime.

| Gate | Result |
|---|---|
| Release/affected checks | One committed linked-worktree library build; 37 affected methods per host. All four new portable checks pass. One test-only rebuild corrects two fixture expectations; only those failed methods repeat. The 21 runtime assets and 2,613 public inputs stay identical. |
| Windows milestone | **3,833 pass, one inherited failure, five skips**; 3,839 unique methods. No prior pass regresses. |
| Ubuntu milestone | **1,243 pass, eight inherited failures, 85 skips**; 1,336 unique methods. The newly selected native inspector failure reproduces on the preceding artifact. No failure becomes a skip. |
| Diagnostic admission | Suppressed master features and unpainted master/layout placeholders follow the same admission rule as painting. Admitted warnings identify their owning part; inheritance and MustUnderstand remain. Three native controls retain exact PDFs; the suppressed callout now strict-exits 0, while admitted unsupported content remains strict-exit 3. |
| Style cleanup | Unused paragraph/run resolved-style records, exclusive calculations and scene-only merging removed. Structural source/cascade metadata and useful inspectors remain. Live noFill, formatting and table precedence pass; five Office-backed table PDFs remain identical. No painting defect or measured speedup is inferred from removing redundant state. |
| Optional chart summary | Fresh probes inspect all pages with an explicit prebuilt inspector. Reported pages and unverified reused coverage are visible. Independent one/two-page controls retain the 200-point page-two move and preserve single-page behavior. Tests assert live summary objects/JSON rather than terminal formatting. Both CI jobs pass this check. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact prior PDFs, diagnostics and semantics. Each outcome converts once. Reuse two strict PPTX passes/fourteen strict DOCX failures; thresholds/residuals remain. No Office export/raster replay. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. No caller migration change. |
| Package/transfer | Same package and consumer source pass Windows and Ubuntu. All 24 assets and two integration PDFs match after transfer; no root library build or suite replay. Existing public-input snapshot reused. |

Two duplicate scene-oracle methods are retired. Their meaningful table precedence
behavior remains covered by live painting and independent Office controls. New
assertions initially misstated two fixtures: a theme-free hyperlink uses black
fallback, and the master declares 556677. The correction changes only tests;
both failed attempts remain in evidence. Newly selected native-font failure is
retained separately. A later CI-only tool failure came from a formatted-output
assertion; its repair reads semantic records and changes no qualified runtime.

The preceding [qualification record](https://github.com/Lokad/OoxPDF/blob/a320bb244e260cb3d105525a890d0726162b5a4a/docs/PreviewQualification.md)
retains workbook/page-ownership results and prior runtime identities. No live
chart-color/font/size/document-shape rendering guard was replaced.

## Retained review-36 line-gate residuals

These earlier line-start screens were not replayed for this chart ownership correction.
Calibri wrapping and compact before-spacing retain their 0.25pt Y failures
(about 0.26-0.28pt). Font-table alternate wrapping retains X failures about
1.21pt/64.86pt; the boundary-invariance PPTX retains two operation-count mismatches.
No thresholds change, overall visual pass or cosmetic reclassification is claimed.
The [preceding tooling qualification](https://github.com/Lokad/OoxPDF/blob/86fbdfb5/docs/PreviewQualification.md)
retains its inputs, independent references and raw results.

## Inherited failures and limitations

**DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles** retains its local failure:
`InvalidOperationException: The second body should wrap to a second row like Office; observed Height=12.610000000000001.`
CI may skip its Office-font precondition. It is not counted as a local pass.

The three previously recorded Linux chart-link failures remain:

- `PptxChartTitleExternalHyperlinkEmitsUriAnnotation`
- `PptxChartAxisTitleExternalHyperlinkEmitsUriAnnotation`
- `PptxChartDataLabelCustomTextHyperlinkEmitsUriAnnotation`

Each observes `InvalidOperationException: Sequence contains no elements` with the
default Windows font catalog absent. The preceding qualification already reproduced
them on its parent runtime. That evidence is retained without another parent build
or method replay. They remain test portability debt; no assertion relaxation,
skip substitution or production font policy was added. Existing platform skips
remain separate from failures.

Fourteen strict DOCX visual failures still include significant typography/geometry.
SVG Review Y is 1.436035pt against 1pt; the raw Word footnote annotation discrepancy
is 46.794pt while candidate picture/link agree within .000016pt. Two field/extraction
flags and the earlier CMYK histogram .311828 versus .98 remain unresolved.
No Office parity or general color-management claim follows from this maintenance.

Qualified chart/font/size/document-shape guards, partial rendering and deferred
rollback-accounting/enum-helper concerns remain recorded in the roadmap. No live
rendering guard was replaced or broad rendering rewrite undertaken.

Out-of-range chart dates retain numeric fallback rather than exact Office tick
text. The preceding 72 DPI screen remains needs-review (MAE 1.3170 > 1.2,
recall .89705 < .90, changed ratio .02740 > .025); it was separate from the
manifest's 144 DPI gate. That completed change is not reopened here.

The newly exercised Linux method
`PptxTextFrameVerticalClipDropsLinesWithBaselineOutsideClip` fails with
`InvalidOperationException: Expected first clipped line to remain visible.`
It reproduces identically on the prior qualified artifact. Default Windows-font
glyph inspection is unavailable on this host. The failure remains recorded; it
was not changed into a skip. Other inherited native-font skips remain distinct from failures. The pinned-font
clipping regression remains covered; the new visibility regressions pass on both hosts.

The narrow clip probe still fails its 72 DPI exploratory foreground-recall limit:
.8948106592 against .90 (MAE .03772714; changed-pixel ratio .00073110). The explicit
horizontal overpaint is removed; remaining glyph/underline differences are retained
as needs-review. This screen is separate from the source manifest's 144 DPI gate.
Candidate overflow-line link annotations also differ from both Office controls;
the existing annotation policy is preserved pending independent qualification.

Three newly exercised Ubuntu inspection methods also reproduce on the preceding
qualified artifact without its native Windows font catalog:

- `PptxTextGlyphRunInspectionReportsInterGlyphAdjustmentAggregates` — `InvalidOperationException: Sequence contains no matching element`.
- `PptxNumberedAutofitFramesEmitOfficeCharacterSpacingTextState` — `InvalidOperationException: Expected 3, got 0.`
- `PptxNumberedNoAutofitFramesEmitOfficeCharacterSpacingTextState` — `InvalidOperationException: Expected 3, got 0.`

They bring the explicit Ubuntu inventory to seven failures. Three exact parent
methods were run, with no parent build; candidate failures were not replaced by
skips. Existing Arial/Cambria preconditions remain separate. Rendering through
explicit portable fonts and both package consumers still passes.

The restored-note 144 DPI screen retains foreground recall .9771410174 < .99.
MAE .01545325, changed ratio .00044409 and SSIM .99406542 pass their limits.
Both canonical control PDFs are identical to their pre-fix versions, so this is
existing typography/geometry residual, not an ID-fix regression. Keep needs-review;
no threshold change or formal manifest/agent-rating pass is claimed.

The newly selected Ubuntu `PptxSceneBuilderBuildsResolvedNodeLists` also fails
with `InvalidOperationException: Expected Hello, got .` before the changed style
assertions: the glyph inspector cannot resolve its default Windows fonts. One
exact method on the preceding qualified artifact reproduces that first failure,
without a parent build. This brings the current Ubuntu inventory to **eight**
inherited failures. Its Windows check passes, as do all new portable checks.
The failure remains recorded and is not replaced with a skip.

## Evidence and workflow

Ignored evidence: artifacts/preview-admitted-content-20261027/qualification.json.
It binds the initial and test-repaired frozen identities, exact method attempts,
Office-backed master/table controls, unchanged portfolio, summary coverage,
same-package consumers, managed/API checks, transfer, catalog and CI. The initial
tool CI failure and repaired check remain separate; broader suites are not replayed.

Package: artifacts/nuget/preview-020-admitted-content-20261027/Lokad.OoxPdf.0.2.0-preview.nupkg.
SHA-256: `0BD94E2B16F03D9B397992193670D149E534F4D9C843C2D9A45496EC76E1AB5F`.
Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity remain. Publication, performance and expansion deferred.

The last full footprint remains 6.778 GB/935,126 files before recent batches,
above the loose 5 GB target. footprint.json measures only this small milestone;
it reuses the shared public-input snapshot. No full workspace rescan, duplicate
input snapshot or blanket deletion was needed.
