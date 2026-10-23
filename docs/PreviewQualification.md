# 0.2.0-preview qualification

The bounded empty-label and verified-deletion milestone is qualified.
Publication, performance optimization and broader feature expansion remain deferred.
The frozen library was built at `47143072133afc83e6acbad975375a31f5327ae2`, SHA-256
`80ECB9563FC068496395125F320E3525FC2248E2F84D26F078E5A7CBAC7D45A4`.
[Windows/Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39990448543)
at the qualified source. CI builds are independent of this artifact.

| Gate | Result |
|---|---|
| Release/affected checks | One committed artifact. Windows: 68 exact methods pass. Ubuntu: 63 pass, five existing Arial-precondition skips. The new empty-label regression and redirected XML agreement checks pass on both hosts. Two inherited nullable test warnings remain. |
| Windows milestone | **3,781 pass, one inherited failure, five skips**; 3,787 unique methods in two disjoint invocations. |
| Ubuntu milestone | **1,091 pass, four inherited failures, 66 skips**; 1,161 unique methods in two disjoint invocations. No previous pass regressed. |
| Empty DOCX labels | The authored-empty label retains its list and 108pt body indent without invented bullet ink. Three ordinary/null controls retain exact PDFs. Independent Word 16 references are reused. |
| Numeric visual screen | Full-page 144 DPI MAE .00098488, changed ratio .00003146, SSIM .99930075, recall 1. All retained numeric limits pass; this is a probe screen, not a formal manifest/agent-rating claim. |
| Verified deletions | 110 production lines removed/12 added. Obsolete XML fallback/selector, unused line tally and five unreachable overflow handlers are gone; no replacement framework or public API. Independent XML tests now exercise live layout/flatten/remap. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact PDFs, diagnostics and semantic results. Each outcome converted once. |
| Independent portfolio gates | Reused by input/PDF identity: two strict PPTX passes/fourteen strict DOCX failures. Existing thresholds and prior residuals are unchanged. |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free package and consumer source pass Windows/Ubuntu. Packed guide bytes/local links match. All 24 root assets and two integration PDFs match without root rebuild or suite replay. |

## Conversion behavior and bounded maintenance

ReadNumberingLevel already supplies missing-value defaults. CreateListLabel now
preserves an explicitly empty bullet string along with its list identity, suffix,
indent and concrete-level override. Word 16 agrees: the missing marker is not
replaced with a visible bullet. The separate null-flag controls preserve their
previous behavior in both converters; no new null-flag policy was introduced.

Shape-only scene dispatch no longer retains an unreachable XML fallback. Its
selector and unused line tally were removed. Useful independent XML inspectors,
scene-fed XML frame geometry, table-cell layout, placeholder/group handling,
Mongolian remapping and bounded span memoization remain. The implementation audit
corrected the review's missed reflection references in three test helpers. Those
tests now invoke the live independent XML operations on their original fixtures;
comparison assertions and canned-font controls remain. No production-only test
wrapper or relocated cloning fallback was added.

JPEG frame/sampling bounds establish that dimension intermediates fit Int32;
chart-range Int32 operands widened before multiplication fit Int64. Five
unreachable exception translations were removed. Header/sampling/pixel checks,
plane-size rejection, checked buffer construction, real area/union/area-count
limits, capacity checks and cancellation remain. No general arithmetic sweep,
resource redesign, measured speedup or performance project is claimed.

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

## Evidence and workflow

Ignored evidence: `artifacts/preview-empty-label-cleanup-20261023/qualification.json`;
artifact: `candidate-final/`; package:
`artifacts/nuget/preview-020-empty-label-cleanup-20261023/`.
One artifact serves the exact affected set, disjoint broader remainders, four
original probes, sixteen portfolio outcomes and both package consumers. Office
references, comparisons, unchanged 715-case/12-family manifest validation and
parent failure evidence are reused. No qualified method replay, parent build,
root build, new Office export or portfolio comparison replay is needed.

The reflected-caller census correction is retained in the review erratum and
current PLAN; the original review is preserved. The old function-call-only search
was insufficient to rule out test consumers. The production dead-path conclusion
remains valid, with independent test coverage preserved through live operations.

The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/45594781/docs/PreviewQualification.md)
retains current-section/hidden-object fixes, earlier maintenance and unresolved
visual gates. Safe managed dependency-free execution, cooperative cancellation,
meaningful bounds, coarse host admission/refunds, bounded caches/staging,
deterministic cleanup, atomic publication and PDF validity remain. The roadmap
stays deferred. The last full workspace inventory is **5.943 GB/885,270 files**;
no full walk or broad purge was repeated. Only this milestone is measured.
