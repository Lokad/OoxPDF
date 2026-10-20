# Preview document review — 2026-10-20

All 32 candidate pages in the nine-source, sixteen-outcome preview portfolio
were viewed beside their pinned Office pages. This supplements the
[qualification record](PreviewQualification.md); no renderer was rebuilt, no
document reconverted and no Office reference regenerated. The prepared package
and frozen runtime are unchanged.

| Document | Pages reviewed across presets | Remaining difference and practical effect |
|---|---:|---|
| Report | 4 | Text leading, table rules and Review balloon/connector placement differ. Text, equations, picture and note remain readable on matching pages. |
| Letter | 2 | Similar text/table offsets; Review comment box has vertical placement differences. The letter and its note fit on one page in each mode. |
| Specification | 4 | Body baseline/leading and Review balloon placement differ. All twenty sections remain readable, with matching page splits and clear note separation. |
| Technical memo | 4 | Late table/fraction frame and Review balloon geometry differ. The table fits above the first-page footer; conclusion and note continue on page two in both engines. |
| Briefing deck | 1 | Existing strict pixel gate passes; chart title, axes, twelve columns and legend remain visible. |
| Held-out deck | 1 | Existing strict pixel gate passes; table heading and all three data rows remain visible without observed cell overflow. |
| Composed static-cell report | 4 | Table/text/static-story geometry remains approximate. Both repeated footer fractions, pictures and calculations are visible. Printed page counters agree with Word; raw cached-field flags remain unresolved in the generic content screen. |
| Composed SVG report | 4 | Small footer vector remains visible on every page. Review prose still fails the original 1-point geometry gate at 1.436 points; body and comment remain readable. |
| Linked-footnote picture report | 8 | Typography and separator geometry differ. Continuation lines 01–20, 21–51 and 52–64 occupy pages two through four in both engines; picture/fraction remain visible on page three. Page four's large preceding white space also occurs in Word. The existing Review annotation comparison remains failed separately from visible paint. |

No serious clipping, overlap or wrong page split was observed in these page-scale
views. The direct review does not establish subpixel fidelity or a universal
ink/clip guarantee. All fourteen DOCX outcomes still fail their original strict
pixel gates; only the two PPTX outcomes pass. These differences are not all
cosmetic, and neither thresholds nor raw failure labels were changed.

The separate page-bound witness found no visible non-whitespace trace rectangle
more than one point outside its page. That reporting tolerance is not a new
acceptance gate and does not detect clipping inside a page. The original
source-backed geometry and content checks remain necessary.

Paired page views, PDF/image hashes and per-outcome dispositions are retained in
ignored `artifacts/composition-c1/preview-020-completion-review/`. The review uses
existing PDF bytes and a standalone PDF inspection/rasterization tool, outside
the dependency-free shipped library.
