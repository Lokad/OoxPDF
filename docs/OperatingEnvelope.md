# Measured operating envelope

## Related-story placement reuse, 2026-10-18

Qualified source `2d587d98` (root `bcbfabb4`) reuses immutable related-story
measurements across the coordinator's fixed-input header/footer placement retry.
Distinct body/note width keys remain. Numbering, kerning, font/geometry changes
and other conversions retain fresh owners. The dictionary ends with its owning
placement scope; no global cache, extra quota or unsafe/native code is introduced.

The same C234 inputs/options and isolated one-cold/three-warm method now allocate:

| Source | Current Final / Review (MB) | Saved Final / Review (MB) | Reduction Final / Review |
|---|---:|---:|---:|
| Fixed report | 24.489 / 27.749 | 0.062 / 0.122 | 0.25% / 0.44% |
| Single long note | 31.063 / 37.885 | 2.262 / 5.196 | 6.79% / 12.06% |
| Two plain notes | 31.206 / 38.064 | 2.264 / 5.198 | 6.76% / 12.01% |

All eight composed-note PDFs and six whole/staged benchmark identities remain
unchanged. Initial font preparation and715,743 native kerning-pair fingerprints
remain unchanged. These are calling-thread allocated volumes, not live-heap
ceilings; overlapping preparation scopes remain non-additive and timing is noisy,
so no timing speedup is claimed. No new Office export or visual raster comparison.
58Windows/42Ubuntu checks pass,16native skips; a parent-library regression counts
1,108 repeated note measurements, versus zero during the corrected placement retry.
Evidence: `artifacts/composition-c1/related-story-measurement-reuse-evidence.json`.

## Compound-note baseline, 2026-10-18

Qualified C233 source `24f3ce2d` (root `5986a1b9`) is measured through a separately
built/self-tested probe, with no library rebuild or visual qualification replay.
All six whole/staged PDFs are stable and equal the qualified outputs.

| Source (pages) | Final warm allocation (MB) / time (ms) | Review warm allocation (MB) / time (ms) |
|---|---:|---:|
| Fixed report (2) | 24.551 / 107.06 | 27.871 / 107.69 |
| Single long note (4) | 33.325 / 173.12 | 43.081 / 91.10 |
| Two plain notes (4) | 33.470 / 62.38 | 43.262 / 75.47 |

Windows/.NET10.0.13, default resolver466faces, buffer output/default64MiB resident
window. Each source/preset uses an isolated process, one cold/warmup conversion
and three warm samples; medians are shown, MB is decimal. Calling-thread allocated
volume is distinct from live heap/process memory. Timings are noisy: no speedup
or mode ordering is established. No new limit or cache is introduced.

Standalone initial font-resource preparation allocates16.656-17.530MB. Coordinated
layout plus snapshot allocates21.807-34.072MB; for the two-note report it is28.011MB
Final/34.072MB Review. These independent scopes overlap and must not be summed;
they do not count layout passes. The Review cost warrants bounded dependency/
reuse attribution, without assuming repeated layout dominates production cost.

Existing independent Office gates and explicit residuals remain unchanged.
Evidence: `artifacts/composition-c1/compound-note-allocation-evidence.json`.

## Collection-face ownership, 2026-10-18

Qualified source `94454cae` (root `06e8183a`) extracts an owned face directly from
read-only collection memory. Standalone font sources still receive defensive
copies; bounded source caches and all parsing/work limits remain. No shared
parsed-font cache or unsafe/native library code.

The same eight C213 mixed-document outcomes, options and isolated three-sample
warm methodology save exactly1,792,688bytes/conversion (decimal1.793MB), or
5.53-7.60percent. This matches one avoided1,792,664-byte `cambria.ttc` copy plus
its array header. Final medians now allocate21.811-27.272MB; AllMarkup medians
25.059-30.617MB. All eight whole/staged PDF identities and three prepared native
faces'715,743 kerning pairs/fingerprints remain unchanged. Timings are noisy
and do not establish a speedup; calling-thread allocation is not a live-heap
ceiling. Inclusive preparation diagnostics remain non-additive.

One frozen Release artifact qualifies33 Windows/31 Ubuntu checks, two native
skips. The synthetic4MiB padded-collection allocation observer fails the prior
library; concurrent ownership, extents, cancellation and rendering neighbors
qualify. Existing Office comparisons and failed DOCX thresholds carry by PDF
identity; both fixed decks retain passing gates. No new Office export or raster.
Evidence: `artifacts/composition-c1/collection-font-ownership-evidence.json`.

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

## Immutable kerning maps, 2026-10-18

Qualified library source `cefbdf99` (root `fbf76050`) keeps kerning pairs in
private immutable key/value arrays. Stable fixed-width ordering preserves signed
advances, source last-wins values and zero-record policy. Existing source-pair,
coverage and work bounds remain; all added loops observe cancellation. No lazy
shared cache, additional quota or unsafe/native code is introduced.

| Compound source | Preset | Previous warm allocation (MB) | Current warm allocation (MB) | Reduction |
|---|---|---:|---:|---:|
| fixed report (2 pages) | Final | 31.645 | 25.915 | 18.11% |
| fixed report (2 pages) | AllMarkup | 35.363 | 29.212 | 17.39% |
| skewed-cell specification (3 pages) | Final | 36.225 | 30.494 | 15.82% |
| skewed-cell specification (3 pages) | AllMarkup | 40.014 | 33.863 | 15.37% |

MB is decimal. The existing probe reuses qualified libraries without rebuilding
them: Windows/.NET10/default resolver (466 discovered faces), one isolated process
per source/preset, file output, one cold/warmup and two measured warm conversions.
Calling-thread allocation falls5.731-6.151MB per conversion. Whole/staged PDFs and
715,743 pair entries across three native faces retain exact fingerprints.
Standalone map allocation drops from4.85-5.76MB to3.19-3.80MB per face. Measured map
construction times remain comparable; these samples do not establish a timing
speedup or live-heap ceiling. Diagnostic stages overlap and are not additive.

An earlier comparison-sort prototype was rejected: lower allocation accompanied
4-6x slower map construction. Its source and measurements remain retained.
66 Windows/56 Ubuntu checks qualify with10 native skips. Two inherited Linux
Office-font assumptions reproduce on the prior library; only native-font
preconditions change, preserving the numerical assertions. Library/14 non-test
assets remain identical and only those two checks rerun. All30 selected PDFs
retain identity and existing Office comparisons/failed visual gates carry.
Evidence: `artifacts/composition-c1/kerning-storage-allocation-comparison.json`
and `kerning-storage-evidence.json`. Earlier figures below remain historical.

## Conversion-owned font preparation, 2026-10-18

Qualified source `83a86cf3` (root `9d41fdae`) reuses parsed fonts only within one
DOCX conversion's note/kerning/page-count preparation retries. Subsets and layout
are rebuilt when their inputs change. Sources retain immutable owned bytes,
existing bounded program caches, cancellation and font-parser/work limits; no
cross-conversion parsed-font cache or production quota is added.

A diagnostic resolver in the allocation probe forwards resolution/catalog access,
face indices, opaque source identities and bytes. Its `--font-load-counts` mode
records delegated program-source requests, not physical reads or allocation.
Four existing note-numbering outcomes request every used face twice before the
change and once afterward. The fixed compound report/specification already load
their main faces once; no general portfolio speedup is claimed.

| Note control | Preset | Previous warm allocation (MB) | Current warm allocation (MB) | Reduction |
|---|---|---:|---:|---:|
| spanning paragraph | Final | 28.47 | 21.02 | 26.2% |
| spanning paragraph | AllMarkup | 28.87 | 21.42 | 25.8% |
| compound equations/images | Final | 50.14 | 33.14 | 33.9% |
| compound equations/images | AllMarkup | 50.53 | 33.53 | 33.6% |

MB is decimal. Each source/preset uses an isolated cold process, one warmup and
three warm samples on Windows/.NET10/default resolver (466 discovered faces).
Calling-thread allocations fall7.45-17.00MB/conversion. Whole/staged PDFs and
prepared-font kerning-map fingerprints remain identical. Candidate timing ran
alongside qualification and is not a controlled throughput comparison. Diagnostic
stages overlap and are not additive; these are not live-heap ceilings.

155 Windows/126 Ubuntu checks pass,29 native skips. Sequential and concurrent
ownership regressions fail the preceding library. All16 portfolio PDFs retain
identity; independent Office comparisons and failed visual thresholds carry.
Evidence: `artifacts/composition-c1/checkpoint-175-note-allocation-comparison.json`
and `font-preparation-retry-evidence.json`. The earlier kerning-materialization
and mixed-document baselines below remain historical observations.

## Composed annotation report, 2026-10-16

The separately pinned [C2 report](../tests/Lokad.OoxPdf.Tests/Cases/Composition/composed-annotations-report.json)
combines tables, numbering, images, six source equations, links/bookmarks,
comments, static stories and a footnote. Both presets produce two pages.
The existing probe reuses qualified library SHA-256
`0535E097903C0F1E7B9A4457ACB47A7648351C2759C841E101268CEA31539E08`
without a library build; frozen source/input identity is `6690268c`.

| Preset | Warm allocation (MB) | Warm median time (ms) | Render allocation (MB) | Cold allocation (MB) / time (ms) |
|---|---:|---:|---:|---:|
| Final / preserved layout | 47.37 | 137.7 | 45.91 | 593.80 / 668.5 |
| AllMarkup / Word-compatible | 51.95 | 123.4 | 50.41 | 599.37 / 680.0 |

One fresh process per preset measures one cold conversion, one warmup and three
warm samples using the default resolver (466 discovered faces). This is the same
Windows/.NET/GC environment as the portfolio below. Allocations cover the calling
thread; cold allocation totals include transient allocations and are not retained
heap size. Rendering includes fonts, layout and painting. Package opening uses
0.16MB, reading 0.77MB, and writing 0.53/0.61MB. Stage samples run with warm caches.
Every whole and staged PDF matches the accepted independent-comparison output.
Three samples do not establish a speedup or explain repeated-layout costs; no
measured-fragment optimization is justified by these totals alone. Evidence:
`artifacts/composition-c1/c2-allocation-evidence.json` and the two `c2-allocation-*.json`
reports. Earlier baselines remain retained.

## Mixed-document portfolio, 2026-10-15

The six fixed public/synthetic documents in the [composition portfolio](../tests/Lokad.OoxPdf.Tests/Cases/Composition/portfolio.json) combine supported features. The four DOCX inputs are measured separately in Final/preserved layout and AllMarkup/Word-compatible modes; the two decks retain absolute-position rendering. Source and output hashes are pinned. This is a compound-document baseline, not a speedup claim.

Measured frozen library `0fa503ac` (version 0.1.5), SHA-256 `5549559DD980AE5152C4F00919F9886A4C62D60D34947055C92FE6562F5858A3`, Windows 10.0.26200, .NET 10.0.13, 64-bit workstation GC, 28 logical CPUs. One process per mode measures one cold conversion, one warmup and three warm samples per input in manifest order. Later inputs share static caches. Allocations cover the calling thread; hashing and report writing are excluded. Cold timing includes runtime/JIT effects. Three warm samples do not establish a stable performance ranking.

| Document | Final warm allocation (MB) | Review warm allocation (MB) | Final warm time (ms) | Review warm time (ms) |
|---|---:|---:|---:|---:|
| report | 47.13 | 51.68 | 111.7 | 138.9 |
| letter | 44.68 | 49.21 | 60.7 | 82.9 |
| specification | 49.76 | 54.34 | 59.7 | 76.4 |
| technical-memo | 46.03 | 50.48 | 44.2 | 66.8 |
| briefing-deck | 10.64 | â€” | 14.4 | â€” |
| held-out-deck | 28.91 | â€” | 28.7 | â€” |

MB is decimal. Every repeated output is stable and matches the accepted CLI PDF. Warm stage probes also reproduce the same complete PDF in all ten mode outcomes. For DOCX, warm rendering allocates 43.5â€“47.9 MB in Final and 47.9â€“52.4 MB in review mode; package opening uses about 0.15 MB, reading 0.60â€“0.94 MB and serialization 0.47â€“0.84 MB. The render phase includes font preparation, layout and painting. It does not identify repeated layout as the dominant cause.

The previous `0aeff058` baseline remains retained; these are updated workload observations, not an optimization comparison. Warm allocation ranges remain similar and three timing samples are too few to claim a speedup. The measurement reused the qualified library without a root rebuild.

The coordinator can repeat work for note kerning/numbering, footnote reservation and static-story displacement. Identify their actual costs and dependencies before introducing measured-fragment reuse. Existing table-cell memoization remains the model for bounded reuse when placement changes. No heap accounting framework or additional production quotas are introduced.

Use the existing allocation probe with `--docx-review` for the review preset, including stage, isolated and concurrent runs; `MeasureAllocation.ps1 -DocxReview` forwards it. The default remains Final/preserved layout. Reports record both modes and the loaded library hash. Measure against a qualified runtime, keeping tool builds separate from the library. Retained local reports and mode-propagation checks are under `artifacts/composition-c1/current-frame-*-allocations.json`.

## Mixed-document attribution, 2026-10-16

The current qualified library is `a08dcfa6`, SHA-256
`CC6630ADF1A965CB9D2CEAF3BFCE8C48CF187F2501BE4B70A6CE0360460556E2`.
The same ten portfolio outcomes remain byte-identical, and the updated allocation
probe loads this artifact without rebuilding the library. The normal cold/warm
baseline retains three warm samples; additional DOCX detail probes cover the four
mixed documents in both presets. They use prebuilt source/context and warmed
process caches. Diagnostic stages overlap and must not be added together.

| Diagnostic stage | Calling-thread allocation (decimal MB) |
|---|---:|
| Initial font-family plan | 0.058â€“0.079 |
| Initial complete font resources | 39.85â€“40.43 |
| Coordinated font/layout preparation plus inspection snapshot | 43.03â€“46.96 |
| Reparse the three prepared embedded SFNT faces | 37.28 |
| Their isolated legacy kern/GPOS pair materialization | 32.32 |

The three faces are Calibri (1,615,616 program bytes, 248,759 kerning pairs),
Calibri (1,650,632 bytes, 259,942 pairs) and Cambria Math (1,574,876 bytes, 207,042
pairs). Their isolated kern/GPOS allocation is respectively 9.67MB, 13.87MB and
8.78MB. Reparse figures include an owned SFNT copy and exclude discovery/collection
extraction; isolated kerning excludes prepared bytes/table-directory construction
and includes reflection invocation. The detail probes neither count coordinator
passes nor establish a speedup. Font preparation is the demonstrated allocation
cost; repeated layout is not established as the main cause on this portfolio.

The ensuing bounded investigation targeted kerning-pair accumulation: retain lookup values,
last-wins behavior, pair/work bounds, cancellation and owned buffers while avoiding
avoidable intermediate pair lists. Do not add a cache, class-layout framework or
production quota knob for this experiment. Existing failed visual thresholds and
independent Office comparisons remain unchanged. Evidence:
`artifacts/composition-c1/layout-attribution-evidence.json`.

## Bounded kerning materialization, 2026-10-16

Frozen `f35fa653` (root `993543d6`), library SHA-256
`0307DA26C0B96D3D30815F25F02BC1FC7D634355AB02E81D647110D68D73D6DC`,
replaces the expanded intermediate pair list with counted source batches and one
final dictionary allocation. Source contribution counts still include duplicates;
last-wins values, owned buffers, cancellation and existing pair/work bounds remain.
No cache or production quota was added.

| Document | Previous Final / review allocation (MB) | Current Final / review allocation (MB) |
|---|---:|---:|
| report | 47.14 / 51.70 | 31.20 / 34.83 |
| letter | 44.70 / 49.23 | 28.75 / 32.34 |
| specification | 49.79 / 54.35 | 33.83 / 37.36 |
| technical memo | 46.06 / 50.50 | 30.10 / 33.55 |
| briefing deck | 10.64 / â€” | 6.72 / â€” |
| held-out deck | 28.91 / â€” | 16.89 / â€” |

These calling-thread allocation measurements use the attribution baseline above
and the same ten inputs/modes, Windows/.NET/default resolver and one cold sample,
one warmup and three warm samples. DOCX allocation falls 31â€“36%, and deck
allocation 37â€“42%. All ten whole/staged PDFs retain their accepted hashes; these
outputs also supply corpus identity without another CLI conversion sweep.

A separate matched sustained-warm report control (sixteen warmups, seven samples)
compares the preceding library with the same final accumulation algorithm:
47.05MB / 45.15ms versus 31.18MB / 45.38ms. Timing ranges overlap; this is an
allocation improvement with no speedup claim. That timing control used the final
algorithm prototype; the subsequently frozen artifact supplies the ten-outcome
measurements and qualification above. Diagnostic stages overlap and are not
additive. Three actual font maps retain all 715,743 keys/values, verified again
from the final artifact by sorted-map fingerprints.

Thirty-one Windows methods pass in one invocation; seventeen Ubuntu methods pass
and four installed-Windows-font checks skip. A generated malformed class matrix
reproduces a bounds exception in the preceding library; wide offset arithmetic
now retains its valid prefix. Managed bounds already prevented unchecked memory
access, and no corruption is claimed. Transfer verifies seventeen assets/four
PDFs without a root build or repeated suite. Independent Office comparisons and
failed visual thresholds remain unchanged. Evidence:
`artifacts/composition-c1/kerning-stream-evidence.json`.

## Shared inline placement baseline, 2026-10-16

The accepted `924eb51a` library (SHA-256
`1697D44EFA8FDFEE49C2A0CD6BE12898D32243612497390571FAC9E6DE474C9C`)
is measured after the shared math, source-run and aligned-field changes. The
existing probe loads the frozen DLL in its mutable diagnostic runtime; neither
library nor probe is rebuilt. Every repeated and staged output matches all twelve
accepted fixed/complete-document PDFs. Office comparisons are not rerun.

| Document | Final / review warm allocation (MB) | Final / review median time (ms) |
|---|---:|---:|
| report | 31.36 / 35.16 | 109.5 / 110.3 |
| letter | 28.92 / 32.67 | 30.5 / 46.5 |
| specification | 34.00 / 37.64 | 45.4 / 61.9 |
| technical memo | 30.28 / 33.78 | 46.2 / 46.6 |
| briefing deck | 6.72 / â€” | 14.5 / â€” |
| held-out deck | 16.89 / â€” | 14.3 / â€” |
| complete static-cell report | 37.11 / 41.25 | 76.4 / 92.8 |

The same Windows/.NET/default-resolver protocol measures one cold conversion,
one warmup and three warm samples per input, sequentially in one process per mode.
Later inputs reuse static caches. Initial DOCX font-resource preparation allocates
23.90â€“25.67MB; inclusive coordinated layout/snapshot allocates 27.32â€“34.40MB.
These diagnostic stages overlap. They do not isolate coordinator repeats, which
remain unproven as the main cost. The first report still reflects tiered warmup;
three samples establish neither a speedup nor a timing regression. No production
optimization or quota change is justified by these totals alone. Evidence:
`artifacts/composition-c1/shared-placement-allocation-evidence.json`.

## Linux forced-break workload, 2026-10-05

Measured renderer/tool revision: `cef195d77e4e8b116002dd8dcdb985eed4e30fd8`,
package version 0.1.5. Ubuntu 24.04.3 under WSL2 (kernel 5.15.167.4),
.NET runtime 10.0.12, SDK 10.0.112, 64-bit workstation GC, Interactive latency,
28 logical CPUs on an Intel Core i7-14700KF. This establishes the measured WSL
Linux leg; default server GC is measured below, while other Linux hosts require separate runs.

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
| 400 | 15.96â€“15.99 MiB | 13.78â€“16.91 MiB | 442â€“454 KiB | 20.9â€“31.6 ms |
| 800 | 30.69â€“30.72 MiB | 14.31â€“18.16 MiB | 860â€“872 KiB | 46.8â€“73.4 ms |
| 1600 | 60.18â€“60.33 MiB | 18.97â€“21.91 MiB | 1697â€“1709 KiB | 95.0â€“165.9 ms |

Cold calling-thread allocation was 16.88 / 31.63 / 61.26 MiB respectively;
cold managed heap peaks ranged from 16.66 to 21.76 MiB. These cold costs use
the explicit font resolver and are not comparable to earlier Windows results
that included discovery of hundreds of installed faces.

| 1600-page output mode | Warm calling-thread allocation | Warm managed heap peak | Sampled process private / working set |
|---|---:|---:|---:|
| File | 60.18â€“60.33 MiB | 18.97â€“21.91 MiB | 116.33â€“122.79 / 92.43â€“102.47 MiB |
| Forward-only file-backed stream | 60.30â€“60.33 MiB | 19.10â€“21.83 MiB | 116.07â€“123.18 / 93.57â€“102.96 MiB |
| Caller memory buffer | 62.42â€“62.57 MiB | 19.21â€“21.77 MiB | 116.12â€“133.27 / 93.58â€“104.05 MiB |

Buffered output adds approximately 2.2 MiB of allocated volume on this workload.
Forward-only output rejects seek, position and length access and checks that
conversion leaves both caller streams open. It uses the stream conversion API;
file output uses the file API and its publication path.

| Parallel conversions | Batch managed heap peak | Sampled private / working set | Batch wall time |
|---:|---:|---:|---:|
| 1 | 19.62â€“20.93 MiB | 129.71â€“131.59 / 91.05â€“94.88 MiB | 92.4â€“101.9 ms |
| 2 | 24.93â€“27.71 MiB | 146.75â€“147.79 / 99.90â€“102.49 MiB | 115.8â€“135.5 ms |
| 4 | 38.80â€“41.50 MiB | 198.21â€“206.23 / 131.04â€“137.01 MiB | 189.3â€“208.6 ms |

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

## Linux GC profiles, 2026-10-06

Measured renderer/tool revision: `a0b042888607b87d2f0c74d48dfa3e7c7903b598`,
package version 0.1.5, on the same Ubuntu/WSL host and .NET 10.0.12 with 28 logical
CPUs. These fresh workstation controls include the subsequent renderer repairs;
the earlier `cef195d7` observations above retain their original provenance.

Workstation and server profiles use `DOTNET_gcServer=0` and `DOTNET_gcServer=1`,
respectively, following the [runtime GC configuration](https://learn.microsoft.com/en-us/dotnet/core/runtime-config/garbage-collector).
Every report confirms the actual GC mode, 64-bit process and Interactive latency.
There were no ambient GC overrides; heap counts and other GC settings use runtime
defaults. Both explicit-font probe self-tests passed before measurement.

The same original inputs and DejaVu Sans font were copied to the native Linux
filesystem. Each profile ran three repetitions of one process-first plus eight
warm conversions per size/output mode, and three warmed batches per concurrency
level. Profiles alternate within each repetition. Windows validation finished
before the series; native Linux load averages were zero at the start.
All 36 reports pass revision, GC, procedure, font/input and output identity checks.

Each row below includes 24 warm samples; allocation and heap values use MiB,
retained deltas use KiB, and times use milliseconds.

| GC | File-output pages | Calling-thread allocation | Managed heap peak | Post-GC retained delta | Wall time |
|---|---:|---:|---:|---:|---:|
| workstation | 400 | 16.15â€“16.15 | 13.85â€“16.58 | 434â€“450 | 20.8â€“31.6 |
| workstation | 800 | 31.01â€“31.09 | 15.16â€“18.11 | 860â€“872 | 41.2â€“84.1 |
| workstation | 1600 | 60.93â€“61.09 | 19.74â€“22.00 | 1689â€“1709 | 90.5â€“166.2 |
| server | 400 | 16.15â€“16.17 | 3.95â€“17.30 | 436â€“452 | 26.2â€“37.6 |
| server | 800 | 31.07â€“31.09 | 8.53â€“31.97 | 862â€“874 | 47.4â€“81.5 |
| server | 1600 | 60.93â€“61.07 | 14.92â€“63.28 | 1699â€“1711 | 85.7â€“181.3 |

The output-mode and concurrency tables use MiB for all memory columns.
At 1600 pages, sampled process memory also differs by GC profile:

| GC | Output mode | Calling-thread allocation | Managed heap peak | Private memory | Working set |
|---|---|---:|---:|---:|---:|
| workstation | file | 60.93â€“61.09 | 19.74â€“22.00 | 116.39â€“124.23 | 92.20â€“103.05 |
| workstation | buffer | 63.16â€“63.31 | 19.01â€“21.79 | 116.53â€“134.30 | 93.91â€“106.14 |
| workstation | forward-only | 60.93â€“61.09 | 19.38â€“21.94 | 116.46â€“124.03 | 94.11â€“106.08 |
| server | file | 60.93â€“61.07 | 14.92â€“63.28 | 161.49â€“459.64 | 84.14â€“177.17 |
| server | buffer | 63.16â€“63.31 | 12.44â€“64.76 | 161.55â€“456.75 | 83.45â€“173.16 |
| server | forward-only | 60.93â€“61.10 | 13.41â€“61.76 | 161.07â€“419.06 | 82.78â€“174.52 |

Each concurrency row covers three warmed 1600-page batches:

| GC | Parallel conversions | Batch heap peak | Private memory | Working set | Batch wall time |
|---|---:|---:|---:|---:|---:|
| workstation | 1 | 20.44â€“21.09 | 130.07â€“131.88 | 91.17â€“93.20 | 101.5â€“103.9 |
| workstation | 2 | 27.05â€“28.79 | 146.17â€“149.52 | 97.02â€“106.33 | 124.9â€“136.7 |
| workstation | 4 | 38.70â€“42.97 | 211.10â€“211.98 | 135.19â€“137.89 | 194.1â€“206.1 |
| server | 1 | 17.07â€“27.07 | 185.32â€“214.02 | 98.02â€“101.98 | 96.7â€“106.5 |
| server | 2 | 32.02â€“35.60 | 210.42â€“210.70 | 109.82â€“112.11 | 108.9â€“120.4 |
| server | 4 | 54.16â€“60.22 | 272.74â€“276.04 | 158.39â€“161.79 | 164.3â€“188.0 |

Allocation volume and retained deltas are similar across these two profiles,
while server GC has wider sampled heap ranges and higher sampled private memory.
The four-way server batches finish sooner in these samples; single-conversion
wall-time ranges overlap. This establishes these default profiles on this host
and workload. Other hosts, heap-count settings and document families require
their own measurements. The approximately 5 ms sampling can miss brief peaks.

Every size retains its previous output identity. The 1600-page output remains
578631 bytes with the SHA-256 recorded above across both GC profiles, all output
modes and every concurrency level.

To repeat either profile, prefix each measurement command below with
`DOTNET_gcServer=0` or `DOTNET_gcServer=1`; check the actual `serverGc` report field.
Repeat the same size, output-mode and concurrency series three times for each
profile with builds, self-tests and fixture generation outside the timed series.

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


## PNG-S1 scanline decoding, 2026-10-08

Frozen runtime `aabe6d3274142067cf52df6ee7c922fda5a4b5a3` streams inflated ordinary/Adam7 rows into final RGB/alpha planes, retaining exact compressed IDAT staging and two maximum-width row buffers. The reservation includes those buffers and final planes before allocation and remains owned through downstream transforms. The default 512 MiB cap and per-image pixel caps remain; configured-limit admission changes with the revised estimate. This is an estimated reservation, not a process-memory ceiling.

Twenty moderate flat/noisy inputs (256/512/1024 square, 128 by 4096 and 4096 by 128, ordinary/Adam7) retain identical RGB/alpha hashes and release their reservations. Each input receives one warm-up and three isolated-scope samples. Current-thread allocated volume excludes input fixture creation; final-plane lengths and forced-GC retained deltas are recorded separately. No process peak is measured. At 1024 square the final planes occupy 4,194,304 bytes:

| Format / pattern | Previous allocated bytes | Streaming allocated bytes | Previous reservation | Streaming reservation |
|---|---:|---:|---:|---:|
| Ordinary / flat-opaque | 8533816 | 4248360 | 8389632 | 4247487 |
| Ordinary / noise | 12905272 | 8627976 | 8389632 | 8627104 |
| Adam7 / flat-opaque | 8550912 | 4250968 | 8390528 | 4250089 |
| Adam7 / noise | 12921688 | 8628936 | 8390528 | 8627999 |

All twenty allocate less; flat-image reservations fall while noisy-image reservations can rise because compressed IDAT was previously omitted. A production allocation bound fails against full-image inflate storage and passes with row-width scratch. Thirty-four Windows and thirty-four Linux imaging checks pass, including 88 format/filter/pass-boundary inputs, split IDAT, truncation exceptions, failure cleanup and deterministic cancellation. A 36-byte fixture cap converts and 35 bytes rejects before publishing on both platforms; resource-summary peaks remain deterministic. Affected checks also pass: 201 OOXML, 36 DOCX-image, and all 21 non-SVG PPTX raster-image tests. SVG vectors do not invoke PNG decoding; their partial group log is archived and full integration retains vector coverage. Twelve public image PDFs and all 34 cached markup PDFs retain bytes; the cache keeps 545 failed gates.

The first freeze passes 34 Windows imaging checks but 33/34 on Linux because native zlib encodes the tiny budget fixture differently (IDAT 15 versus 13 bytes). Its checkout and failed evidence are archived. The correction changes only tests and their fixture: a valid stored-deflate PNG has IDAT 18 bytes and pins the portable 36/35-byte boundary. Decoder source remains identical to `a1913986`.

Exact local 0.1.5 package smoke contains tested Windows DLL SHA-256 `0386257E5E5AF9A6A7D2BED05EFE43F3270224E57AB61772ED27F1FEB456877B`. Evidence is under `artifacts/plan-revision-20261005/png-s1/`. Compressed IDAT staging remains; downstream compression/retention and full-suite integration are separate. Release preparation remains deferred.

## SVG group stream ownership, 2026-10-10

RV07-E6 uses the existing content admission and resident window for isolated SVG
Form streams as well as page streams. Each group's ASCII payload charges once
before its retained snapshot; page admission counts the remaining page stream.
Completed page/group streams enter one chunked spill store, with blanked resource
descriptors and weak source identities. A retained strong-key staging map fails
the streaming collection regression; the qualified weak map releases earlier
payload owners. Spill/resident outputs retain exact bytes, and cancellation closes
the delete-on-close spill file. All output bytes retain admission before writing.

Fourteen infrastructure regressions pass on Windows and Linux, including large
streams beyond the 64KiB scratch chunk and exact/zero limits. The earlier measured
workloads contain no SVG groups; mixed SVG peak memory remains unmeasured.
Evidence: `artifacts/plan-revision-20261005/rv07-e6/`.


## Composed-document checkpoint, 2026-10-17

C79 measures the qualified C78 library `ef2a1e999c8e2a8904c00dee41bf1bafd491f3d1`
(SHA-256 `5B74B0ABF7479B21E01920EA54F7016606D2FD3C51DC960A14BBE9995C8DB0BC`).
The seven existing public sources in
`tests/Lokad.OoxPdf.Tests/Cases/Composition/portfolio-checkpoint.json` combine
prose, numbering, tables, images, equations, navigation, static stories, notes,
comments and fields, plus two distinct PPTX decks. Five DOCX inputs run as both
Final/preserved and AllMarkup/Word-compatible. Every output matches the frozen CLI
and staged serialization byte-for-byte. No library build or test-suite replay.

The existing allocation probe was built separately with
`tools/NewFrozenAllocationRuntime.ps1` and passed its self-test. Each outcome has
one isolated-process cold conversion and one warm conversion, file output, under
.NET 10.0.13, Windows 10.0.26200, 64-bit workstation GC/Interactive, 28 logical CPUs.
The default resolver discovers 466 faces. These are calling-thread allocated
volumes in MiB, not live heap, process memory or allocation on background threads.

| Document | Pages | Cold Final / Review | Warm Final / Review |
|---|---:|---:|---:|
| Report | 2 | 551.28 / 555.72 | 30.19 / 33.71 |
| Letter | 1 | 548.77 / 553.14 | 27.65 / 31.15 |
| Specification | 2 | 553.85 / 558.31 | 32.77 / 36.32 |
| Held-out memo | 2 | 550.13 / 554.47 | 29.01 / 32.41 |
| Complete composed report | 2 | 557.59 / 562.53 | 35.52 / 39.50 |
| Briefing deck | 1 | 523.04 | 6.31 |
| Held-out deck | 1 | 535.39 | 16.02 |

Standalone initial DOCX font-resource preparation allocates 22.79â€“24.49 MiB;
coordinated layout plus its inspection snapshot allocates 26.08â€“33.08 MiB.
Those diagnostic stages overlap and must not be added; they do not count layout
invocations. For the report Final, standalone kerning preparation allocates
5,755,968 bytes for Calibri regular, 5,763,264 for Calibri bold and 4,848,568 for
Cambria Math, with 248,759 / 259,942 / 207,042 resulting pairs. This identifies
font preparation as a bounded investigation; it does not demonstrate that repeated
layout is the dominant cost or justify an unbounded parsed-font cache.

The report Final cold sample allocates 551.28 MiB while sampling 116.97 MiB managed
heap peak, 218.74 MiB private memory and a 5.52 MiB post-GC retained delta. Allocation
churn and retained/peak memory are distinct. Sampling can miss brief peaks. Office
exports ran concurrently, so reported wall times are descriptive only: no stable
benchmark, speedup claim or cross-host conclusion. Repeat uncontended samples before
accepting an optimization on timing grounds. Stage and output checks remain valid.

Evidence: `artifacts/composition-c1/checkpoint-79-allocation-final.json`,
`checkpoint-79-allocation-review.json`, `checkpoint-79-allocation-summary.json` and
`checkpoint-79-allocation-runtime/identity.json` (all under the same directory).
Keep the existing independent visual thresholds and explicit failed gates.
