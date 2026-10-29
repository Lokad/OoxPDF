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

The canonical [portfolio manifest](../tests/Lokad.OoxPdf.Tests/Cases/Composition/portfolio.json)
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

The 0.2.0 maintenance checkpoint is complete. Development toward 0.3.0 now
prioritizes composed DOCX layout, generalized PPTX chart styles and effective
properties, and calibrated Office image references alongside the retained PDF
comparisons. Changes need source-backed witnesses and independent appearance
and content checks. Historical thresholds and residuals above remain visible.

The complete 9,570-line development and qualification journal remains in
[checkpoint history at the reviewed revision](https://github.com/Lokad/OoxPDF/blob/6584386f9f36bd0c32dc59c7f3043f5b39f8d062/docs/CompositionCheckpoint.md).
Detailed local evidence remains under ignored artifacts; private inputs stay
private. Historical measurements retain their original source/runtime identities.
