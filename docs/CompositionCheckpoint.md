# Composed-document reliability checkpoint

The **0.2.0** release is published on NuGet. Current source/runtime,
Windows/Linux results, package identity and CI provenance are recorded in
[PreviewQualification.md](PreviewQualification.md). See [release scope](PreviewRelease.md)
and [caller migration](PublicApiChanges.md).

The fixed portfolio has nine sources and sixteen outcomes across public Final
with preserved layout and AllMarkup with Word-compatible review margins. It
combines prose, numbering, tables, equations, images, links, headers and notes.
Supported content must survive ordinary combinations and source-preserving
transformations. The finite gate covers content omissions, pagination, serious
clipping/overlap and regressions within supported scope; it does not require
pixel-perfect Office parity.

The canonical source manifest, `tests/Lokad.OoxPdf.Tests/Cases/Composition/portfolio.json`,
includes all nine pinned sources. Default measurement and content-audit commands
use it; the older seven-source checkpoint manifest has been retired.

Independent Office references and thresholds remain fixed. Two PPTX strict gates
pass; fourteen DOCX strict gates retain typography/geometry failures. The inherited
balloon second-row assertion, SVG Review Y 1.436035pt against 1pt, raw Word
footnote annotation mismatch and two field/extraction flags remain explicit in
[current qualification](PreviewQualification.md). Correct PDF structure and
manifest validity are separate from Office fidelity.

Those are retained historical results: sixteen original Office reference PDFs
are no longer available locally. Fresh 0.3.0 references use explicit new identities;
image captures remain uncalibrated until checked against the visible application.

The 2026-11-06 source checkpoint refreshes all sixteen outcomes against fresh
Office PDF references. Page counts agree, and the content screen finds no missing
Office prose tokens or simple math. Two dynamic-field source flags remain
unresolved. Strict image results remain two PPTX passes and fourteen DOCX
failures; token presence does not prove glyph ink, clipping or placement.
This validates the development runtime, without promoting a release candidate.

The 0.2.0 maintenance checkpoint is complete. Development toward 0.3.0 now
prioritizes composed DOCX layout, generalized PPTX chart styles and effective
properties, and calibrated Office image references alongside the retained PDF
comparisons. Changes need source-backed witnesses and independent appearance
and content checks. Historical thresholds and residuals above remain visible.

The 0.3.0 table-padding/inheritance controls restore Word's text origins and
distinguish table/cell formatting from text-only or empty style declarations.
The earlier forward-parent geometry regression is corrected. All forty-two
selected Final/Review outcomes pass the original 1pt/.1pt geometry, text-color
and page checks; twenty-four also pass unchanged image gates. Eighteen image
failures and built-in style/shading approximations remain. These PDF-backed
controls await visible application-view calibration and do not replace the
original portfolio. Local evidence: `artifacts/tinh/completion.json`.

The 0.3.0 automatic-column-label checkpoint recovers eight omitted value suffixes
across five retained controls. Fifteen fresh Office PDF controls retain all 98
automatic source labels and agree on line counts. Fifty focused Windows/Linux
checks pass across .NET 10/11; two clipping regressions fail on the parent.
Unchanged image gates improve from 3 to 8 passes out of 26, with seventeen MAE
improvements and no regression. Explicit no-wrap PDFs remain byte-identical.
Manual/rich-text, stacked and other label placements retain their existing paths.
Eighteen raw image failures, a 1.046pt Times horizontal residual and an inherited
28.4pt manual-plot frame displacement remain explicit. These PDF-backed controls
are uncalibrated against the visible application view. Local evidence:
`artifacts/tbar-label-wrap-1104/qualification.json`.

The 0.3.0 factor/edge-height correction resolves three displaced or distorted
Cartesian plot rectangles, reducing 25.9â€“87.7pt errors to at most 0.003pt against
Office PDF strokes. Forty source controls retain their API/PDF audit; thirty-seven
candidate PDFs and twelve retained neighbors, including both portfolio decks,
remain byte-identical. Forty-six focused Windows/Linux checks pass across both
frameworks. Five selected image gates pass unchanged thresholds, improving from
two; the three corrected outputs also match Office's extracted word multisets.
A broader caption-admission prototype was withdrawn after a fresh accepted
control contradicted it. The 28.4pt vertical and 38.4pt horizontal admission
residuals remain open. Local evidence: `artifacts/tmanual-plot-frame-1104/qualification.json`.

The horizontal source-scale checkpoint shares one range, interval lattice and
category crossing across consistently oriented clustered bar plots. Nineteen
fresh Office controls cover signed and translated data, partial bounds,
reversal, plot order and both crossing roles; zero-bound headroom also has column
and line witnesses. Seventy focused Windows/Linux checks pass across .NET 10/11,
and all three new regression methods fail on the accepted parent. All nineteen
outputs match Office's extracted word multisets; twelve retained neighbors,
including both portfolio decks, remain byte-identical. Unchanged image gates
improve from two to fourteen passes out of twenty, with no MAE regression.
Six image failures remain, including a confirmed 28.9pt vertical plot error at
numeric-axis maximum crossing. The references remain uncalibrated against the
application view. A stale admission assertion found by CI was repaired through
harness-only builds; qualified runtime assets, images and package are reused.
Local evidence: `artifacts/thorizontal-source-range-1104/qualification.json`.

The measured automatic-column scale checkpoint improves unchanged image gates
from 12 to 36 passes across 43 source-pinned chart controls, with 33 lower errors
and no lost pass. Numeric ranges and grid intervals agree with Office in 38,
printed numeric text in 37; five range/interval errors and one further manual-label
admission difference remain. Seven image failures and the uncalibrated direct
Office-image failure remain explicit. All bars, admitted legend text/keys and
source-data ratios survive; the sixteen fixed-portfolio PDFs remain byte-identical.
All 136 focused Windows/Linux checks pass across .NET 10/11. This is an improved
approximation, not an exact Office scale algorithm. Local evidence:
`artifacts/taxis-measured-capacity-1106/qualification.json`.

The complete 9,570-line development and qualification journal remains in
[checkpoint history at the reviewed revision](https://github.com/Lokad/OoxPDF/blob/6584386f9f36bd0c32dc59c7f3043f5b39f8d062/docs/CompositionCheckpoint.md).
Detailed local evidence remains under ignored artifacts; private inputs stay
private. Historical measurements retain their original source/runtime identities.
