# 0.2.0-preview qualification

The bounded optional-part ownership/workbook cleanup milestone is qualified. Publication, performance optimization and expansion remain deferred.
Library build source is `3468ededc4f456f22e59a4d357511d0ed48be548`; final
harness source is `b64617b2ca626e24f48b747fe6aa54b0c48a3a91`. Library SHA-256:
`24BB3DB079F639BAB2248302C914DF787235370B9780BE5016643BA08FA9B8B0`.
[Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/40342363065)
passed at the final harness source. CI builds are independent of the frozen runtime.

| Gate | Result |
|---|---|
| Release/affected checks | One accepted committed linked-worktree library artifact. Initial 44 methods: 44 pass Windows; 41 pass/three declared installed-font skips Ubuntu. Refined 50-method capability coverage also includes six neighboring checks exercised by broader qualification/fixture repair. All four new portable methods pass both hosts. |
| Windows milestone | **3,806 pass, one inherited failure, five skips**; 3,812 unique methods/3,817 attempts in three invocations. Five failed fixture checks were repaired; no successful method was replayed. |
| Ubuntu milestone | **1,177 pass, seven inherited failures, 73 skips**; 1,257 unique/1,259 attempts. Two fixture failures repaired; four newly selected line-spacing methods pass once. No prior pass regressed. |
| Original Office controls | Six pinned Word Final controls agree on active formatting/header content. Unreferenced styles now keep 12pt black text and unreferenced settings preserve the second-page default header. Four unaffected PDFs are byte-identical. No new Office export/rasterization. |
| Workbook cleanup | Eight internal constructor adapters, three exclusive conversion helpers and nine unused DOCX content-type constants removed. The canonical workbook constructor, static initialization, live reader/range/provenance and caches remain. Fourteen reflected test entry sites now use the live reader. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact prior PDFs, diagnostics and semantic results. Each outcome converted once. |
| Independent visual gates | Reused by identity: two strict PPTX passes/fourteen strict DOCX failures. All prior visual thresholds and residuals remain; manifest validation is not fidelity evidence. |
| API/runtime | 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free package and consumer source pass Windows/Ubuntu. All 24 runtime assets and two integration PDFs match after transfer, without a root build or broad suite replay. |

## Conversion behavior and bounded maintenance

Optional DOCX configuration and related stories follow their owning document
relationship. Missing, external, dangling or unselected payloads retain the
existing empty/default handling. Styles and numbering use the existing shared
XML loader and MustUnderstand warning deduplication. Main-document/presentation
content-type repair lookup is preserved. Caller guidance explains the correction
for handcrafted packages that relied on unreferenced parts.

Workbook cleanup removes internal test adapters and their exclusive dictionary
conversion paths. Test data passes through the live reader using the existing
workbook package fixture helper. No compatibility chain was relocated into the
library, no builder framework added, and no chart color/layout rule changed.

An initial comment edit accidentally removed a live chart-render/finally block;
the compiler rejected it before freezing. The original block is restored. That
failed build is retained separately. The accepted corrected library is built
once and reused. Broader qualification exposed five malformed test cases: two
missing owning relationships and one shared spacing fixture lacking the rels
content-type registration. Their relationships/manifest are repaired with all
assertions retained. A tracked join verifies 21 unchanged runtime assets and
2,596 public inputs. Only failed checks are repeated; newly selected Linux
spacing checks run once. The superseded CI run was cancelled to save quota.

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

Ignored evidence: `artifacts/preview-part-ownership-20261026/qualification.json`;
artifact: `candidate-qualified/`. Package:
`artifacts/nuget/preview-020-part-ownership-20261026/`.
Exact sets, raw failed reports, reviewed fixture joins, isolated temporary folders,
constructor/caller census and original input/Office pins are retained. The
unchanged 715-case/12-family manifest validation and identical portfolio comparisons
are reused. The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/fdb9d738/docs/PreviewQualification.md)
retains section/navigation and public-contract maintenance with its evidence.

Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, cleanup, atomic publication and
PDF validity remain. The roadmap stays deferred. The latest full workspace check
is 6.778 GB/935,126 files before the preceding batches (2026-10-26), above the loose
5 GB target. Measure this milestone and its shared input snapshot separately;
preserve useful evidence and avoid an unnecessary full scan or blanket deletion.
