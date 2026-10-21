# Composed-document reliability checkpoint

The **0.2.0-preview** candidate remains unpublished. Current source/runtime,
Windows/Linux results, package identity and CI provenance are recorded in
[PreviewQualification.md](PreviewQualification.md). See [preview scope](PreviewRelease.md)
and [caller migration](PublicApiChanges.md).

The fixed portfolio has nine sources and sixteen outcomes across public Final
with preserved layout and AllMarkup with Word-compatible review margins. It
combines prose, numbering, tables, equations, images, links, headers and notes.
Supported content must survive ordinary combinations and source-preserving
transformations. The finite gate covers content omissions, pagination, serious
clipping/overlap and regressions within supported scope; it does not require
pixel-perfect Office parity.

Independent Office references and thresholds remain fixed. Two PPTX strict gates
pass; fourteen DOCX strict gates retain typography/geometry failures. The inherited
balloon second-row assertion, SVG Review Y 1.436035pt against 1pt, raw Word
footnote annotation mismatch and two field/extraction flags remain explicit in
[current qualification](PreviewQualification.md). Correct PDF structure and
manifest validity are separate from Office fidelity.

Maintenance removes retired corrections, inert state and verified inactive
helpers, retargeting useful tests/probes to live implementations. It does
not generalize chart-color, five-paragraph balloon or font/size/document-shape
rules. Future rendering corrections require a supported-scope counterexample,
source-owned measurement/placement/painting contracts and independent Office
comparison. Performance optimization and broader feature expansion are deferred.

The complete 9,570-line development and qualification journal remains in
[checkpoint history at the reviewed revision](https://github.com/Lokad/OoxPDF/blob/6584386f9f36bd0c32dc59c7f3043f5b39f8d062/docs/CompositionCheckpoint.md).
Detailed local evidence remains under ignored artifacts; private inputs stay
private. Historical measurements retain their original source/runtime identities.
