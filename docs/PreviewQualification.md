# 0.2.0-preview qualification

The subsequent public-contract clarification changes comments and caller guidance
only. Compiler token comparison verifies the four C# files have unchanged executable
tokens; all 24 qualified/root runtime assets match. Eight existing stream/resource
checks pass in one focused invocation. Updated package guidance uses the identical
qualified DLL. The preceding Windows/Ubuntu, Office, corpus and CI results below
are reused; no rendering build, broader replay or new visual comparison is required.
Evidence: `artifacts/preview-public-contracts-20261026/qualification.json`;
package: `artifacts/nuget/preview-020-public-contracts-20261026/`.

The bounded section-selection/navigation milestone is qualified.
Publication, performance optimization and broader expansion remain deferred.
Library build source is `01b1c35f`; final test-only source is
`5e24a67802a48af00bc1e0bd6888334cf3f0b6f3`. The qualified library SHA-256 is
`CB3CCD69138BCF95AA6FDAD34546E2809A20B64BD39315C254F31B8669F017DC`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40321051391)
passed at the final test source. CI builds are independent of the local artifact.

| Gate | Result |
|---|---|
| Release/affected checks | One accepted corrected library artifact; 36 affected methods pass Windows, 33 pass/three declared installed-font skips on Ubuntu. Six new portable methods pass both hosts. Two inherited nullable test warnings remain. |
| Windows milestone | **3,802 pass, one inherited failure, five skips**; 3,808 unique methods, 3,810 method attempts across three invocations, including two failed harness checks repaired without replaying passes. |
| Ubuntu milestone | **1,152 pass, seven inherited failures, 72 skips**; 1,231 unique methods, 1,233 attempts. No prior pass regressed. |
| Original review controls | Six original sources reuse pinned Word Final exports. Later-section first headers and paragraph-end GoTo links now survive. Four control PDFs remain byte-identical; restored end destination residual X .248pt/Y .48602pt. |
| Independent neighbor | New four-page Word Final control confirms first header/footer selection at section start, default continuation and undefined even-story suppression. Matched text origins across all seven sources differ at most X .02401pt/Y .07996pt; font sizes agree. This is not a raster parity claim. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact prior PDFs, diagnostics and semantic results. Each outcome converted once. |
| Independent portfolio gates | Reused by identity: two strict PPTX passes/fourteen strict DOCX failures. Existing visual thresholds and note/clip recall failures remain unchanged. |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free package and consumer source pass Windows/Ubuntu. All 24 root assets, a corrected two-section DOCX and representative PPTX match without root rebuild or suite replay. |

## Conversion behavior and bounded maintenance

DOCX page records retain their owning section's first physical page. Shared
header/footer and floating-story selection, static-frame mapping and eligibility
consume that provenance; continuation templates and the existing caches retain it.
Default/even parity and page-field numbering rules are unchanged. This does not
add full continuous-section or page-number-restart support.

A bookmark after its final text item/run resolves to that source endpoint. The
same-run fallback cannot grab an earlier wrapped fragment before the requested
offset. Hidden trailing text does not own the destination. Ordinary prose keeps
its existing run-based layout rather than opting into inline-object layout for
metadata. Terminal non-text object anchors retain their existing placement rules.

The first compiled attempt fixed section headers but did not handle terminal
bookmarks on the legacy prose path; its outputs remain as an unqualified attempt.
The corrected library is built once and reused throughout accepted qualification.
A fixture-path repair and two endpoint-assertion repairs affect tests only. The
tracked report join verifies 21 unchanged runtime assets and 2,589 public inputs.
Only the two failed endpoint checks are repeated on the accepted library; broader
sets are disjoint from the affected selection. No accepted pass is replayed.
One malformed generated validation input was rejected by Word; it is retained
separately and contributes no rendering or reference-pass evidence.

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

## Evidence and workflow

Ignored evidence: `artifacts/preview-section-navigation-20261026/qualification.json`;
artifact: `candidate-qualified-final/`. Package:
`artifacts/nuget/preview-020-section-navigation-20261026/`.
Exact method sets, failed attempts, reviewed report joins, isolated temporary
folders and runtime/input hashes are retained. Original Word references and
identical portfolio comparisons are reused. The unchanged 715-case/12-family
manifest validation is reused, not counted as Office fidelity. The
[preceding qualification](https://github.com/Lokad/OoxPDF/blob/16019acf/docs/PreviewQualification.md)
retains font-failure/diagnostic controls and earlier maintenance evidence.

Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
admission/refunds, bounded caches/staging, cleanup, atomic publication and PDF
validity remain. No helper deletion, API removal, quota or framework was added.
The preserved roadmap stays deferred. The latest full workspace measurement,
before this batch, is **6.778 GB/935,126 files**, zero errors (2026-10-26), above
the loose 5 GB target. The milestone and its new shared input snapshot are measured
separately; no repeat full scan or unverified deletion was needed. Preserve useful
evidence and use infrequent stale-artifact maintenance.
