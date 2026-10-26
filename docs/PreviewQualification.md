# 0.2.0-preview qualification

The bounded font-semantics milestone is qualified. Publication, performance
optimization and broader feature expansion remain deferred. The accepted library
was built at `f84afdf2b887286a7211725d0e7123f916939880`, SHA-256
`F4F90E77E06608231F3D496171C18849B2818F1098138A6FC51FBCA2F05E849D`.
The final harness is `137bf996aff304f54bb633991d413db4d1f4718d`; production source
and all 21 runtime assets are identical across its test-only repair.
[Windows/Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/40306968987)
at the final source. CI builds are independent of the local frozen artifact.

| Gate | Result |
|---|---|
| Release/affected checks | One accepted library artifact. All 36 affected methods pass Windows; Ubuntu has 33 pass/three explicit Calibri/Symbol precondition skips. Seven new portable tests pass both hosts. Two inherited nullable test warnings remain. |
| Windows milestone | **3,791 pass, one inherited failure, five skips**; 3,797 unique methods. |
| Ubuntu milestone | **1,118 pass, seven inherited failures, 69 skips**; 1,194 unique methods. No prior pass regressed. |
| Original font probes | Eight strict conversions, no diagnostics, all Word text retained. Four joined/split or English/localized pairs produce byte-identical PDFs. The English family name no longer replaces the ideograph with a question mark. |
| Independent geometry | The mixed-text run-boundary X shift is zero, previously 7.7109375pt. Word residuals remain X -.091003pt/Y -.900024pt. Font-slot choices match the retained Arial/MS Gothic and Arial/Courier New controls. No formal pixel-parity claim. |
| Representative documents | Nine sources/16 outcomes/32 pages have exact prior PDFs, diagnostics and semantic results. Each outcome converted once; no new Office exports or rasterizations. |
| Independent portfolio gates | Reused by identity: two strict PPTX passes/fourteen strict DOCX failures. Existing visual thresholds remain fixed; the retained note screen has its own recall failure. |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free package and consumer source pass Windows/Ubuntu. All 24 root assets and two integration PDFs match without root rebuild or suite replay. |

## Conversion behavior and bounded maintenance

Discovery indexes supported Windows family-name records as aliases for one
canonical face and owned file source. Canonical names, style ranking and catalog
face counts remain stable; aliases retain requested spelling. Portable tests cover
multilingual names, collection face indexes, shared ownership and style selection.

The existing DOCX run splitter now follows declared Ascii/HighAnsi/EastAsia and
complex-script font choices. Equivalent choices remain together; complex-script
bold/italic overrides retain their boundary. Measurement and emission consume the
same fragments, preserving UTF-16 source offsets, links, bookmarks and revisions.
Body, cell, header and footer rendering are covered without font/size-specific
exceptions or general shaping support. Two obsolete whole-text scans were removed;
the typeface resolver was renamed without a forwarding wrapper.

Two early compilation attempts failed before running tests; their logs remain.
The first broad runs then exposed three metadata tests whose record equality
compared the new alias arrays by reference. The repaired assertions compare all
scalar fields and ordered alias contents. Only those three tests reran per host,
against the exact same library. Reviewed joins verify 21 runtime assets and 2,582
public inputs; all other accepted checks are reused. Initial failed reports and
[initial CI](https://github.com/Lokad/OoxPDF/actions/runs/40305747865) remain separate.
No rendering assertion or Office threshold was relaxed.

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

Ignored evidence: `artifacts/preview-font-semantics-20261026/qualification.json`;
final artifact: `candidate-qualified/`; original production build:
`candidate-accepted/`. Package:
`artifacts/nuget/preview-020-font-semantics-20261026/`.
Eight original font probes and sixteen portfolio outcomes each converted once.
Existing Office references, unchanged comparisons and 715-case/12-family manifest
validation are reused. The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/2182bc44/docs/PreviewQualification.md)
retains earlier maintenance and unresolved gates.

Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
admission/refunds, bounded caches/staging, cleanup, atomic publication and PDF
validity remain. The preserved roadmap stays deferred. One infrequent workspace
check records **6.778 GB/935,126 files**, zero errors, above the loose 5 GB target.
Preserve inputs/useful evidence and use verified stale-artifact maintenance.
