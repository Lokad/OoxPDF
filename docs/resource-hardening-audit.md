# Resource-hardening audit (historical)

The September 2026 audit records the implementation and tests at its original
revisions. It is historical evidence, not the resource-control contract for
0.2.0-preview. Its work fuel, retained-byte ceilings and per-domain accounting
were subsequently removed; do not restore them from this audit's old verdicts.

The [complete audit at its retained revision](https://github.com/Lokad/OoxPDF/blob/21795dd9f44d0fcc913cfa15d3511726b50abe86/docs/resource-hardening-audit.md)
contains the requirement table, original commits, workload observations and
qualification results. Those results retain their original runtime and scope.

Current contracts:

- [Caller migration](PublicApiChanges.md): intentional API removals and one coarse,
  host-supplied memory-admission lease with deterministic disposal.
- [Diagnostics](Diagnostics.md): meaningful input/output bounds, cancellation,
  bounded staging and output/spill totals. Totals do not measure the managed heap.
- [Preview qualification](PreviewQualification.md): the current frozen candidate,
  platform/package evidence and unresolved rendering failures.
- [Operating envelope](OperatingEnvelope.md): measured workloads and limitations;
  historical scaling probes do not establish arbitrary-document capacity.

Keep the safe managed, dependency-free library, useful bounded caches, archive/XML/
image/font bounds, cooperative cancellation, bounded page staging, atomic file
publication and PDF structural validity. These protections have current owners;
the historical audit is not an additional implementation roadmap.
