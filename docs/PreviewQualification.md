# 0.2.0-preview qualification

The bounded numeric-identity/emission-inspection milestone is qualified.
Publication, performance optimization and broader feature expansion remain deferred.
The accepted library was built at `a7b329bb83ec36dfe6c2210a65081441c7f4f5c3`, SHA-256
`4674BA9A91F2560D779EB4DFDE6D9B028605C3C3A48D90F4C0B24A1D7D6D0D7D`.
[Windows/Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/40097454937)
at the accepted source. CI builds are independent of the frozen artifact.

| Gate | Result |
|---|---|
| Release/affected checks | One corrected artifact serves final qualification. All 32 affected methods pass on Windows. Ubuntu: 26 pass, three inherited native-font failures, three explicit font-precondition skips. New numeric-ID tests pass both hosts. Two inherited nullable test warnings remain. |
| Windows milestone | **3,784 pass, one inherited failure, five skips**; 3,790 unique methods in two disjoint invocations. |
| Ubuntu milestone | **1,101 pass, seven inherited failures, 69 skips**; 1,177 unique methods in two disjoint invocations. No previously qualified pass regressed. |
| DOCX identity | Padded note/numbering references now produce byte-identical PDFs to canonical references, restoring the complete note, marker and 36pt indentation. Canonical controls remain identical to pre-fix PDFs. Declaration/reference aliases, distinct/repeated notes, reserved stories, overrides and navigation are covered. |
| Emission inspection | Hidden text reports zero glyph runs while raw XML retains its frame. Visible output stays identical. Existing frame/character-spacing checks pass unchanged after preserving each node's frame identity. Actual glyph-runs.json from the existing tool bound to the qualified DLL confirms visible=2, hidden=0. |
| Representative documents | Nine sources/16 outcomes/32 pages retain exact PDFs, diagnostics and semantic results. Each accepted outcome converted once. |
| Independent portfolio gates | Reused by input/PDF identity: two strict PPTX passes/fourteen strict DOCX failures. Existing thresholds remain fixed; the note screen has its own retained recall failure. |
| API/runtime | All 247 public declarations unchanged; IL-only, no P/Invoke/imported COM, explicit AllowUnsafeBlocks=false; assembly/file version 0.2.0.0. |
| Package/transfer | Same dependency-free package and consumer source pass Windows/Ubuntu. Packed guide bytes/local links match. All 24 root assets and two integration PDFs match without root rebuild or suite replay. |

## Conversion behavior and bounded maintenance

The reader uses one decimal-value key rule for demonstrated note and numbering
identities on both declaration and reference sides. Numbering overrides and NSID
ownership remain. Styles, relationship/font identities, comments and bookmark
policy retain their existing rules. No broad identity framework or public option
was introduced. Parsed note/numbering identities are canonical decimal keys;
original package XML remains unchanged apart from existing story-marker handling.

Glyph emission inspection now uses admitted scene inputs and master visibility.
Per-node frame identities stay distinct when spans are combined, matching the
painting/coalescing boundary. Shape/table font-preflight ordering and cached
geometry remain. The independent raw XML layout view is preserved; extra source
frames are intentionally separate from emitted glyphs. No rendering/layout guard,
font/size-specific rule or resource quota was added.

The initially frozen 84276217 artifact failed three existing Windows inspection
checks and CI 40096178265: node-local frame indexes collided after admission was
shared. Its failed evidence and package remain separate. Correction a7b329bb
preserves node boundaries; the three checks join the 32-method selection with
unchanged assertions. Painting/preflight and the raw XML inspection implementation
remain unchanged. This was an inspection regression, not a rendering fix or a
reason to relax independently validated Office checks.

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

Ignored accepted evidence: `artifacts/preview-numeric-identity-final-20261024/qualification.json`;
artifact: `candidate-final/`; package:
`artifacts/nuget/preview-020-numeric-identity-final-20261024/`.
The rejected attempt remains under `artifacts/preview-numeric-identity-20261024/`.
One corrected artifact serves final tests, six original probes, sixteen portfolio
outcomes, both consumers and verified transfer. Each accepted exact method ran
once across the affected set and disjoint remainder; the changed revision was
qualified separately after the initial regression. No root or parent rebuild.

Retained Word/PowerPoint references and unchanged comparisons/715-case/12-family
manifest validation were reused. The existing PPTX inspector executable was
bound to the exact accepted DLL in an owned directory; no tool build was needed.
The [preceding qualification](https://github.com/Lokad/OoxPDF/blob/e875a130/docs/PreviewQualification.md)
retains earlier maintenance and unresolved visual gates. Safe managed dependency-
free execution, cancellation, meaningful bounds, coarse admission/refunds,
bounded caches/staging, cleanup, atomic publication and PDF validity remain.
The preserved roadmap stays deferred. The infrequent workspace check records
**6.696 GB/929,543 files**, zero errors, above the loose 5 GB target. The accepted
milestone is about 54 MB; useful failed-attempt evidence remains separate. Preserve
inputs/evidence and use verified stale-artifact maintenance rather than blanket
deletion. The loose cap remains a maintenance guideline.
