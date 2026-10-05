# Measured operating envelope

These are workload observations, not memory ceilings. Conversion budgets protect
specific expansion, retention and output boundaries. Host capacity also includes
runtime/JIT, resolver ownership, concurrent conversions, input/output owners and
headroom. Office and PDFium validation processes are outside library measurements.

## Linux forced-break workload, 2026-10-05

Measured renderer/tool revision: `cef195d77e4e8b116002dd8dcdb985eed4e30fd8`,
package version 0.1.5. Ubuntu 24.04.3 under WSL2 (kernel 5.15.167.4),
.NET runtime 10.0.12, SDK 10.0.112, 64-bit workstation GC, Interactive latency,
28 logical CPUs on an Intel Core i7-14700KF. This establishes the measured WSL
Linux leg; other Linux hosts and server GC require separate runs.

Each DOCX contains one `x` paragraph with `pageBreakBefore` per page and a
612 by 792 point page size. There are no images, tables, markup or charts.
The explicit single-face resolver uses DejaVu Sans (759720 bytes, SHA-256
`ae7b7855e115a5966d8b1b3f80f254ccc117ec86f9965e202ee2940453837280`).
Its regular face serves every requested family, and one lazy file source is
shared across warm conversions and each concurrent batch. This excludes
Windows font discovery and does not measure Office-compatible font matching.

The probe self-test passed before measurement, including embedded-font checks,
allocation attribution, positive sampled peaks, concurrent byte agreement and
equal DOCX/PPTX bytes across buffer, file and strict forward-only output.
For each size, three independent processes ran one cold plus eight warm file
conversions. Three equivalent series covered buffered and forward-only output at
1600 pages. Each concurrency level ran three fresh processes, with one sequential
warmup then one batch of 1, 2 or 4 conversions. Inputs and font files were on the
native Linux filesystem; report serialization ran outside conversion counters.

All ranges below include every warm sample, using MiB = 1048576 bytes.

| File-output pages | Warm calling-thread allocation | Warm managed heap peak | Post-GC retained delta | Warm wall time |
|---:|---:|---:|---:|---:|
| 400 | 15.96–15.99 MiB | 13.78–16.91 MiB | 442–454 KiB | 20.9–31.6 ms |
| 800 | 30.69–30.72 MiB | 14.31–18.16 MiB | 860–872 KiB | 46.8–73.4 ms |
| 1600 | 60.18–60.33 MiB | 18.97–21.91 MiB | 1697–1709 KiB | 95.0–165.9 ms |

Cold calling-thread allocation was 16.88 / 31.63 / 61.26 MiB respectively;
cold managed heap peaks ranged from 16.66 to 21.76 MiB. These cold costs use
the explicit font resolver and are not comparable to earlier Windows results
that included discovery of hundreds of installed faces.

| 1600-page output mode | Warm calling-thread allocation | Warm managed heap peak | Sampled process private / working set |
|---|---:|---:|---:|
| File | 60.18–60.33 MiB | 18.97–21.91 MiB | 116.33–122.79 / 92.43–102.47 MiB |
| Forward-only file-backed stream | 60.30–60.33 MiB | 19.10–21.83 MiB | 116.07–123.18 / 93.57–102.96 MiB |
| Caller memory buffer | 62.42–62.57 MiB | 19.21–21.77 MiB | 116.12–133.27 / 93.58–104.05 MiB |

Buffered output adds approximately 2.2 MiB of allocated volume on this workload.
Forward-only output rejects seek, position and length access and checks that
conversion leaves both caller streams open. It uses the stream conversion API;
file output uses the file API and its publication path.

| Parallel conversions | Batch managed heap peak | Sampled private / working set | Batch wall time |
|---:|---:|---:|---:|
| 1 | 19.62–20.93 MiB | 129.71–131.59 / 91.05–94.88 MiB | 92.4–101.9 ms |
| 2 | 24.93–27.71 MiB | 146.75–147.79 / 99.90–102.49 MiB | 115.8–135.5 ms |
| 4 | 38.80–41.50 MiB | 198.21–206.23 / 131.04–137.01 MiB | 189.3–208.6 ms |

Parallel wall times include task setup and output verification; batch peaks cover
the whole batch, including those owners. Parallelism increases throughput here
but does not reduce the latency of a single conversion. The highest observed
batch managed heap was 41.50 MiB; that is a data point, not a configured ceiling.

Every series was byte-stable, with an embedded TrueType resource. The 1600-page
output was identical across all modes/repetitions/concurrency levels: 578631
bytes, SHA-256 `e9a5b5acc87b9291933194117b026180e01c1b7ced36b247a606ebbe1e07cc41`.
Retained deltas stayed within the displayed ranges across eight warm conversions;
they are measured survivors against each pre-run baseline, not a process leak
counter or an eviction/reload test.

Original input SHA-256 values were:

- 400 pages: `469aaea0b0d04dad8e9d1016fab8b09a000f41e529b21a907f1cef3bb65d5c64`
- 800 pages: `bdcbed07ccc2906700707c2c48ce9d6effe12aa638df47102c49d1e1f73beb9c`
- 1600 pages: `3d79b42d4466aa1661839a8b1110654368200ccaf70d0fd17fd9a528673fbbf0`

## Reproduction and scope

Generate the same OOXML parts with `tools/NewOperatingEnvelopeFixtures.ps1`.
Its fixed ZIP timestamps produce a new reproducible package identity; record the
generated input hashes rather than assuming they match the original archives.
Build AllocProbe in Release and run its explicit-font self-test first:

```sh
dotnet build tools/Lokad.OoxPdf.AllocProbe -c Release
dotnet tools/Lokad.OoxPdf.AllocProbe/bin/Release/net10.0/Lokad.OoxPdf.AllocProbe.dll --self-test --font-file /usr/share/fonts/truetype/dejavu/DejaVuSans.ttf
dotnet tools/Lokad.OoxPdf.AllocProbe/bin/Release/net10.0/Lokad.OoxPdf.AllocProbe.dll --font-file /usr/share/fonts/truetype/dejavu/DejaVuSans.ttf --out artifacts/linux-pages.json --output-mode file --isolate --warmup 0 --iterations 8 artifacts/operating-envelope/inputs/breaks-400.docx artifacts/operating-envelope/inputs/breaks-800.docx artifacts/operating-envelope/inputs/breaks-1600.docx
```

Repeat three times; separately measure buffer/forward-only output and file-mode
concurrency 1/2/4. Keep builds and fixture generation outside measured processes.
Pin runtime/GC, CPU, actual font/input hashes and source revision for each report.

Calling-thread allocated bytes exclude work on background threads. Managed heap,
private memory and working set use approximate 5 ms sampling and may miss brief
spikes; their meanings differ across operating systems. Process lifetime peaks
include startup and earlier work. Neither these numbers nor image/page-content
reservations establish arbitrary-document OOM immunity. Broader resource mixes,
resolver eviction, charts/shading retention and other host/GC settings remain
separate workloads.

