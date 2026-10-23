# 0.2.0-preview qualification

The bounded horizontal-clipping and report-comparison milestone is qualified.
Publication, performance optimization and broader feature expansion remain deferred.
The frozen library was built at `939bf0ec6945b263b5e5f7e5b7d08b0e7252c65f`, SHA-256
`9ED49C37F44DAE55BE8D48A40EAFECB9328197F5E05E0A0DAC35BDC2E910E45D`.
[Windows/Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39970681309)
at the qualified source revision. CI builds are independent of this artifact.

| Gate | Result |
|---|---|
| Release/affected checks | One corrected, committed artifact serves qualification. Two inherited nullable test warnings remain. Windows: all 16 affected methods pass. Ubuntu: five pass, one inherited font-dependent failure, ten explicit Arial-precondition skips. The new portable regression passes on both hosts. |
| Windows milestone | **3,777 pass, one inherited failure, five skips**; 3,783 exact unique methods in two disjoint invocations. |
| Ubuntu milestone | **1,025 pass, four inherited failures, 60 skips**; 1,089 exact unique methods in two disjoint invocations. No previous pass regressed. |
| Horizontal clipping | Three frozen conversions; narrow-frame ink beyond the source right edge falls from 643 pixels to zero, matching retained PowerPoint 16 bounds. Wide/vertical controls preserve non-clip geometry; the vertical control PDF is identical. No new Office export. |
| Report comparison | Seventeen focused controls pass locally and in both-platform CI. Actual retained perfect VisualDiff reports pass with numeric/string zeros. Required absent/null/nonnumeric/nonfinite recall is rejected; finite-drop and unrequired legacy/inapplicable behavior remain. No rerender or replacement gate framework. |
| Representative documents | Nine sources/16 outcomes/32 pages remain byte-identical, including diagnostics and semantic results. Each outcome converted once. |
| Independent portfolio gates | Reused by PDF/input identity: two strict PPTX passes/fourteen strict DOCX failures. Thresholds are unchanged; the new narrow-frame screen has its own retained recall residual. |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free package and consumer source pass Windows/Ubuntu. Packed guide bytes/local links match. All 24 root assets and two smoke PDFs match, without root build or suite replay. |

## Conversion behavior and bounded maintenance

Horizontal clipping is read through the existing text-body property path and feeds
one shared horizontal-clip decision for shape/cell frame construction. Explicit
horizontal clipping can constrain the shape while vertical overflow remains allowed.
Mandatory table-cell clipping, independently qualified vertical/preset/column rules
and transformed clips retain their existing behavior. No font/width-specific guard,
new public option or renderer framework was added.

The corrected regression covers absent/overflow/clip horizontal settings with
no-wrap/wrap in shapes and cells. Cell clipping remains mandatory. Clipped emission
now retains the same visible fragments as Office on the narrow probe; its extracted
text naturally differs from the earlier unbounded emission. Annotation eligibility
was not generalized from that comparison.

The comparator keeps finite numeric zero separate from empty/missing values.
Passed cases require finite recall when the current case requires it or when
previously measured required recall would otherwise disappear. Wholly unrequired
legacy/inapplicable recall remains compatible. No invented zero score or weakened
visual threshold is used.

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
was not changed into a skip. Ten additional affected methods explicitly skip
when Arial is absent. The new pinned-font clipping regression passes on both hosts.

The narrow clip probe still fails its 72 DPI exploratory foreground-recall limit:
.8948106592 against .90 (MAE .03772714; changed-pixel ratio .00073110). The explicit
horizontal overpaint is removed; remaining glyph/underline differences are retained
as needs-review. This screen is separate from the source manifest's 144 DPI gate.
Candidate overflow-line link annotations also differ from both Office controls;
the existing annotation policy is preserved pending independent qualification.

## Evidence and workflow

Ignored evidence: `artifacts/preview-horizontal-clip-20261023/qualification.json`;
artifact: `candidate-final/`; package:
`artifacts/nuget/preview-020-horizontal-clip-20261023/`.
One corrected artifact served focused tests, disjoint broader remainders, three
clip probes, sixteen portfolio outcomes, both consumers and verified transfer.
The initial unqualified build/CI failure is retained: the regression initially
assumed table cells allow vertical overflow. A test-only correction preserved their
existing mandatory clip before local qualification. No qualified candidate method
was replayed. A probe inspection assertion was corrected by reinspection without
reconversion. One new Linux failure required one parent method on an existing
artifact; no parent build or wider parent replay occurred.

Office references/comparisons, unchanged 715-case/12-family manifest validation
and prior Linux chart-link provenance were reused. The unused preflight compiled
files remain ignored: automatic approval review rejected their optional removal
with “blocked by policy.” Inputs, logs and failure evidence are retained.

The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/817dea56/docs/PreviewQualification.md)
retains equivalent DOCX breaks/PPTX targets, audit-comment cleanup, chart-date and
reference-tool work, BMP recovery, fonts/images and earlier Office evidence.
Safe managed dependency-free execution, cancellation, meaningful bounds, coarse
host admission/refunds, bounded caches/staging, cleanup, atomic publication and
PDF validity remain. The preserved roadmap stays deferred.

The last full workspace inventory remains **5.943 GB/885,270 files**, above the
loose 5 GB target. No full walk or broad purge was repeated; the milestone records
its own footprint and reuses shared inputs/worktree.
