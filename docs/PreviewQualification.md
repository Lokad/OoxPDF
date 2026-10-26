# 0.2.0-preview qualification

The bounded selected-story diagnostics/XML-ownership milestone is qualified.
Publication, performance optimization and expansion remain deferred. Library and
test executables were built once at `d7185ac8ca5c846af99df2c505be6d8b77b0e5e0`.
Final expectation/input source: `a66c70c2d953382f3fc72abbc21265bde0f744b0`.
Library SHA-256: `5E1FB6CC9B3556783D629DB36E5BCECD75D28E14C458752938F7342C9C770637`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40360554766)
passes at the final source. CI builds are independent of the frozen runtime.

| Gate | Result |
|---|---|
| Release/affected checks | One committed linked-worktree build, no library/test rebuild. Initial 29 exact methods pass both hosts. Refined 35-method capability: Windows 35 pass; Ubuntu 32 pass/three declared Windows-font skips. All five new portable methods pass both hosts. |
| Windows milestone | **3,815 pass, one inherited failure, five skips**; 3,821 unique methods/3,827 attempts in three invocations. Only six failed expectation checks repeated. |
| Ubuntu milestone | **1,219 pass, seven inherited failures, 77 skips**; 1,303 unique methods/1,306 attempts. Three failed expectation checks repeated; three newly selected native-font methods skip in one extra exact-set invocation. No prior failure becomes a skip. |
| Original Office controls | Five pinned Word Final inputs retain identical candidate PDFs. The unselected uncached header now stays quiet/strict 0; a selected unsupported field still warns/strict 3. No new Office export/rasterization. |
| XML ownership | The live related-part helper and package cache share one DOM. Note marker edits detach only the note being changed; cached source XML remains unchanged across reads. Bounds, cancellation and MustUnderstand deduplication pass. No measured speedup or heap benefit claimed. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact prior PDFs, diagnostics and semantic results. Each outcome converted once. |
| Independent visual gates | Reused by identity: two strict PPTX passes/fourteen strict DOCX failures. All thresholds/residuals remain; manifest validation is not fidelity evidence. Six unresolved history-control PDFs/diagnostics are unchanged across both candidate presets. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free package and consumer source pass Windows/Ubuntu. All 24 runtime assets and two integration PDFs match after transfer; no root library build or broad suite replay. |

## Conversion behavior and bounded maintenance

Header/footer diagnostics now resolve the same typed internal source references
as the text and drawing readers. An unused relationship alone no longer makes an
uncached field a rendering fallback. Selected unsupported content and deliberate
package-level warnings remain. Notes/comments retain their existing part-level
probing; referenced-but-page-inactive variants remain outside this correction.

The related XML loader, probe and normal story reader use the existing per-package
DOM owner. The obsolete raw probe loader and its file-I/O recovery arms are removed;
malformed XML recovery remains. Style/numbering readers consume immutable source
XML. Automatic note labels are applied to a detached note element. Cached DOMs can
remain live through the reader stage; DOCX passes the resulting typed model to
layout. There is no new cache, accounting layer or performance claim. Production
source is 28 lines smaller. No layout/history guard is replaced.

Broader checks found candidate-specific gold fields expecting ignored-equation
warnings from an unused header in one mixed-source fixture. Six rows across three
files now expect zero such warnings. Office geometry, glyph/image/text values,
input hashes and all executable test assertions are unchanged. Every one of the
24 runtime asset hashes remains identical; 2,601 other public inputs are unchanged.
Only six Windows/three Ubuntu failed methods are repeated. All raw failures remain.
The original CI attempt failed these stale expectations; the final CI passes.

The tracked report join requires identical inputs, so it was not used for this
expectation-only input change. The bounded retained audit verifies all runtime
assets/inputs and permits only the six documented diagnostic field pairs. This
adds no general join framework or validation exception. Committed LF formatting
is preserved. No library or test executable was rebuilt after the data correction.

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

Ignored evidence: artifacts/preview-story-ownership-20261026/qualification.json.
It binds original/final input snapshots, exact method reports, the reviewed gold
field delta, original Office controls, XML ownership probe/audit, identical portfolio,
same-package consumers, transfer and green CI. The actual failed and qualified
artifacts remain. An unused intermediate runtime/snapshot was purged after identity
verification; its manifest and committed history remain.
The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/56871caf/docs/PreviewQualification.md)
records missing-picture recovery and reference-test cleanup.

Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, deterministic cleanup, atomic
publication and PDF validity remain. The section-history header probe is unresolved:
Word shows both current/historical text in one synthetic input. Its schema and
normalization scope is not established; this milestone preserves that behavior.

The last full workspace check remains 6.778 GB/935,126 files before recent batches
(2026-10-26), above the loose 5 GB target. This milestone is about 65.5 MB plus
47.4 MB of retained initial/final input snapshots. The unused intermediate freeze's
2,630 files were removed. No full rescan or blanket deletion was needed.
