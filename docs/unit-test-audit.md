# Unit-test audit (historical)

The original September 2026 audit guided the shift toward independent Office-PDF
fidelity evidence. Its suite counts, accounting descriptions and remaining-work
list belong to that revision. They are not the current preview status.

The [complete audit at its retained revision](https://github.com/Lokad/OoxPDF/blob/21795dd9f44d0fcc913cfa15d3511726b50abe86/docs/unit-test-audit.md)
preserves the original measurements, assertions and development history.

The testing principles remain useful:

- Use unit checks for API/CLI behavior, diagnostics, deterministic bytes, input
  bounds, cancellation, resource ownership, parsers and PDF primitives.
- Check content survival, pagination, source geometry and draw order in supported
  combinations. Retarget tests to live behavior when deleting an implementation
  seam; do not discard independently grounded assertions.
- Use independent Office output and fixed visual thresholds for rendering
  correctness. Candidate-specific operator grouping or exact matrices alone are
  not an Office oracle. Preserve meaningful source-backed geometry checks.
- Select exact affected methods and their neighbors during a coherent batch,
  then run broader/platform/package qualification at a meaningful milestone.
  Reuse one frozen artifact and proven identical outputs/references.

Current evidence is in [preview qualification](PreviewQualification.md).
[Visual validation](VisualValidation.md) describes the workflow, and
[diagnostics](Diagnostics.md) and [caller migration](PublicApiChanges.md)
describe the current output totals, staged emission and coarse memory admission.
The old audit's unreported-counter and unimplemented-staging claims are superseded.
