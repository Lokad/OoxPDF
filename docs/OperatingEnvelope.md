# Operating envelope

The library offers cooperative cancellation, meaningful input/output bounds,
bounded font caches and page staging, and optional coarse host memory admission.
These protect specific operations and ownership; they do not bound the entire
managed heap or guarantee conversion of arbitrary documents. See
[limits and admission](Diagnostics.md#output-bounds-and-host-admission) and
[caller migration](PublicApiChanges.md).

File conversion publishes atomically after staging. Caller-owned streams stay
open and may retain an incomplete prefix after a failure; discard that output.
Stream conversion exceeding its resident content window requires `TemporaryStorage`;
without a provider it fails instead of creating process-temp files. File conversion
selects explicit filesystem backing. Package, staged/output I/O and owned storage
disposal and selected DOCX/PPTX font acquisition are awaited; CPU layout and
drawing remain synchronous.
`MemoryBudget` admits package buffers/parts and coarse parsed content-type/XML
cache capacity, with temporary parsing/normalization headroom. One existing package
scope holds retained estimates through its final readers; failed parsing refunds
attempted growth and cache hits reuse funding. XML admission is a coarse estimate,
not an exact graph/heap measurement. It also admits PDF content staging, private SFNT/
subset programs, retained kerning-pair arrays and decoded raster planes with PNG IDAT/scanline and baseline
JPEG component/RGB construction overlap before allocation. Image/font compression
also admits actual output capacity, growth overlap and final copying. Image payloads
remain funded through emission; font-compression storage refunds after its stream
is written. Kerning sort scratch is admitted while retained pair arrays remain
funded, then refunds after sorting; failed parsing refunds program and pair storage
together. Storage refunds after
its users finish, including failure/cancellation. Crop and generated shadow/glow
planes are admitted for their invocation; recolor reuses privately owned pixels.
Embedded-font ToUnicode streams are measured and emitted with one fixed 4 KB
formatting span, without a copied mapping list or a full text/encoded payload.
Its scopes share the caller's root; they create no private allowance. Caller-owned
buffers are not charged twice. Font leases fund acquisition and copy overlap;
explicit longer-lived cache budgets keep storage admitted through eviction or
shutdown until the final reader releases it. Governed operations without a cache
budget acquire private programs. Models, local-discovery metadata, other parsed maps and
other PDF resources remain pending, so
this is partial coverage. Null budget is ungoverned. Use
representative measurements to size incidental-storage and concurrency headroom.

## Measurement scope

The historical observations below pin older source/runtime identities; they are
not measurements of the current release. Calling-thread allocations are allocation
volume, not peak or retained memory, and exclude background threads. Diagnostic
stages overlap and must not be added. Sampled heap/process peaks may miss short
spikes. Font discovery, JIT, installed fonts, GC configuration, concurrency and
caller buffers materially affect the host footprint. No speedup or arbitrary-
document throughput claim follows from these samples.

## C281 bounded composition performance baseline, 2026-10-19

The existing probe is built separately against C280's exact frozen library, with
no library build, Office export or qualification replay. Reuse the historical
fixed-portfolio baseline under its own older runtime identity. Two additional
composed sources are measured in Final and review: each input gets a fresh child,
one cold conversion, one warmup and three warm samples. All four repeated/staged
PDFs match accepted hashes. MB below means calling-thread allocated decimal MB.

| Input | Warm Final / review MB | Median Final / review ms | Initial fonts MB | Render Final / review MB |
|---|---:|---:|---:|---:|
| Mixed box/math report | 12.295 / 12.354 | 28.560 / 29.612 | 10.630 / 10.630 | 11.497 / 11.555 |
| Ordinary picture/box footer | 2.613 / 5.711 | 14.452 / 14.478 | 1.471 / 1.475 | 2.007 / 5.026 |

Diagnostic stages overlap and are not additive; coordinated snapshots include
font preparation. The probe does not count layout invocations. Source review
confirms per-conversion parsed-font reuse across bounded coordinator retries,
alongside existing table/related-story measurement reuse. Font preparation remains
a demonstrated cost; review painting adds material allocation outside coordinated
layout. No new counter/cache/quota or speculative layout rewrite is justified by
these four outcomes. Three timing samples do not establish speedup, and sampled
peaks may miss brief spikes. ordinary-composition-allocation-evidence.json retains
runtime/input/output/environment/stage pins and the historical baseline links.

## Fixed mixed-document baseline, 2026-10-18

The existing frozen C212 library (`89fe008b`, root `a2608a43`) renders the four
fixed DOCX sources in both public presets. The probe alone is built against its
pinned library; no library rebuild or visual qualification replay.

| Source (pages) | Final warm allocation (MB) / time (ms) | AllMarkup warm allocation (MB) / time (ms) |
|---|---:|---:|
| Report (2) |26.319 /104.14 |29.634 /120.34 |
| Letter (1) |23.603 /66.74 |26.852 /77.05 |
| Specification (2) |29.065 /165.07 |32.410 /140.06 |
| Held-out memo (2) |25.119 /101.54 |28.302 /95.58 |

Windows/.NET10.0.13, default resolver (466 discovered faces), buffer output and
the default64MiB resident window. Each input/preset uses an isolated process:
one process-cold conversion followed by three warm samples; table values are
medians. MB is decimal. Allocation covers the calling thread; these are neither
live-heap ceilings nor throughput guarantees. OS caches are unspecified and
timing is noisy, so no speedup or mode ordering is claimed.

Standalone initial font-resource preparation allocates18.171-18.750MB.
Preparation diagnostics overlap; do not sum them or infer that repeated layout
is the main cost. Independent SFNT reparses are diagnostic workloads, not proof
that production repeats those parses. All eight whole/staged output identities
remain stable and equal the qualified PDFs; fourteen retained reviewed pixel
pairs and all eight failed strict gates carry. No new Office export or raster.
Evidence: `artifacts/composition-c1/compound-allocation-baseline-evidence.json`.

These are workload observations, not memory ceilings. Conversion budgets protect
specific expansion, retention and output boundaries. Host capacity also includes
runtime/JIT, resolver ownership, concurrent conversions, input/output owners and
headroom. Office and PDFium validation processes are outside library measurements.

## Historical Linux scaling observations

The 2026-10-05 forced-break workload uses one-character paragraphs with forced
page breaks and pinned DejaVu Sans. It tests staging/scaling, not mixed-document
rendering cost. Warm file-output samples reported:

| Pages | Calling-thread allocation (MiB) | Sampled managed heap peak (MiB) | Wall time (ms) |
|---:|---:|---:|---:|
| 400 | 15.96-15.99 | 13.78-16.91 | 20.9-31.6 |
| 800 | 30.69-30.72 | 14.31-18.16 | 46.8-73.4 |
| 1600 | 60.18-60.33 | 18.97-21.91 | 95.0-165.9 |

Subsequent workstation/server GC, caller-buffer, forward-only output and
concurrency experiments show that allocation volume and retained deltas can stay
similar while sampled process peaks vary substantially. These remain historical
workload observations, not current capacity ceilings. Full environment, source,
input/output identities and tables are preserved in
[the original operating record](https://github.com/Lokad/OoxPDF/blob/56d72d1aa1c23c6309129878b606502e646cbb34/docs/OperatingEnvelope.md#linux-forced-break-workload-2026-10-05).
Rendering checkpoint narratives remain in [checkpoint history](CompositionCheckpoint.md).

## Reproduction

Use `tools/NewOperatingEnvelopeFixtures.ps1` for the forced-break fixture family;
record generated hashes rather than assuming they match older archives. Build
only the allocation probe against an already qualified library with
`tools/NewFrozenAllocationRuntime.ps1`; do not rebuild or requalify the library
for a measurement-only change. Run its explicit-font self-test first.

Pin source/runtime identity, input/font hashes, OS, CPU, GC, output mode, warmup
and iteration count. Keep fixture generation and builds outside measured
processes. Mixed-document work uses the fixed portfolio, with Final/preserved and
AllMarkup/Word-compatible measured separately. Time uncontended runs before
accepting an optimization. Performance optimization is deferred for the preview.
