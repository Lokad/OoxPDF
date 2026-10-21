# 0.2.0-preview qualification — 2026-10-21

The sixth bounded maintenance milestone is qualified with the retained failures
below. Publication, tags, releases, performance optimization and broader rendering
expansion remain deferred. See [scope](PreviewRelease.md), [migration](PublicApiChanges.md)
and the retained [document review](PreviewVisualReview.md).

One Release production/test artifact was built and frozen at
`6b863c580621ccd20f2212e3448e1b6832e6cf0b`, with 24 assets and 2,582 public inputs.
A later tool-only change at `b5b9f855845b2772585f1c5bdf97f74df33ceb8a` contains and
restores Unix runner TMPDIR; it changes no compiled production or test source.
The allocation probe alone was rebuilt/self-tested against the frozen DLL after
its forwarding call was redirected. No library rebuild or performance claim.

Library SHA-256:
`CE8F6B8EB924B35886D01DD69F23604A434BFC373A829BB17B5F44C65B223DF9`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39585004509)
at `b5b9f855`. CI compiler outputs remain independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Release build | Passed; two existing nullable warnings in unchanged DOCX test code |
| Windows milestone | 3,735 pass, **1 inherited fail**, 5 skips across 3,741 unique methods |
| Same frozen DLL on Ubuntu | 820 pass, 0 fail, 43 environmental skips across 863 unique methods |
| Exact affected capability | 22 methods: Windows 20 pass / 2 explicitly declared filesystem skips; ext4 Ubuntu all 22 pass |
| Test reuse | Each platform's affected methods excluded from its remaining milestone selection; two disjoint invocations, no unchanged-method replay |
| Filesystem regression | Three additional checks pass on Ubuntu using case-insensitive NTFS; all three fail with the preceding runtime. ext4 preserves distinct existing entries |
| Reference cache | 12 synthetic checks per platform cover completion, recovery, page names/count/bytes, DPI/tool/PDF identity and retained legacy provenance; real cached Office PDF served unchanged |
| Worker supervision | Seven fake-worker checks per platform cover spaced/quoted paths, healthy/unavailable/timeout outcomes, retained logs and unrelated process survival; no Office export |
| Adjacent comparison scripts | Text, graphics, flow and PPTX emission suites pass from spaced repository paths on both platforms |
| Affected report mechanics | Eight controls per platform: declared skips pass; unexpected skips, failures, duplicates, incomplete coverage and inconsistent counts refuse; outcomes retained and TMPDIR restored |
| CLI diagnostics | Windows 4 pass / 2 symbolic-link privilege skips; Ubuntu all 6 pass. Existing link publication and failed-publication cleanup remain covered |
| Package consumers | Windows and Ubuntu consume the same .nupkg; DLL identity, versions, caller guides, embedded fonts and admission lifetime pass |
| Representative documents | Nine sources / 16 outcomes / 32 pages; every PDF, diagnostic list and semantic result matches the preceding runtime. Each outcome converted once |
| Independent Office comparisons | Reused through proven input/mode/PDF identity: two PPTX strict passes / fourteen DOCX strict failures; no new Office exports, portfolio rasters or changed thresholds |
| Root transfer | All 24 assets verified, two integration PDFs match; no root rebuild or broad-suite replay |
| Safe managed | Explicit unsafe=false, IL-only, no P/Invoke/imported COM types; dependency-free runtime, assembly/file 0.2.0.0 |

The Windows balloon assertion remains the same inherited failure: 12.61pt single-row
body where its Office-derived expectation requires another row. The method name
and failure detail match the preceding qualification. Its assertion and rendering
policy remain unchanged. Earlier nonportable native-font inspection failures remain
historical evidence; that slow probe was not rerun. Local failures/skips are separate
from green CI.

## Changes and compatibility

File-path guards now use existing directory entries instead of inferring storage
case behavior from the OS. Case aliases on Linux NTFS cannot replace source input
or a generated PDF through diagnostics. Provably distinct case-only entries remain
allowed; two uncreated case-only outputs and case-only parent ambiguity refuse
conservatively. Choose names that differ beyond case. Captured paths, observer
exceptions, cancellation, admission lifetime and atomic publication remain covered.

The generic cache requires its existing derivative completion record and exact
ordered page identity. An interrupted/unrecorded raster recovers through fresh
staging from the trusted PDF; old partial pages remain separate as evidence.
Metadata updates replace their file atomically. Recorded corruption refuses.
The real cached Office PDF smoke preserved its exact bytes; no Office was launched.

Office supervision passes discrete arguments and retains redirected logs through
success, failure and timeout. The four neighboring comparison tests use direct
argument arrays. Affected reporting preserves all outcomes and requires exact
expected-skip declarations; it cannot silently turn an unexpected skip or failed
check into a pass. Existing report-join failure-to-skip rejection remains covered.

The four-line internal WriteBlank forwarding wrapper was deleted. Its 34 test
calls and one allocation-probe call now use WriteStaged with the same defaults.
There is no new public signature removal, rendering rule, ownership layer, quota
or validation framework. Useful inspectors, font caches, coarse host admission,
input bounds, cancellation, bounded staging and PDF validity remain.

Existing chart point-count, five-paragraph balloon and font/size/document-shape
rules remain live, documented debt. No new independent Office evidence justified
replacing them. Strict DOCX failures remain significant geometry/typography
residuals: SVG Review Y 1.436035pt exceeds 1pt; linked-footnote raw Word URI differs
by 46.794pt although candidate picture/link bounds agree within .000016pt. Two
field/extraction flags remain. None were relabeled cosmetic or passed.

## Retained candidate evidence

Ignored evidence: `artifacts/preview-alias-cache-20261021/qualification.json` and
`completion-audit.json`; frozen runtime under `candidate/`. Failed mock-runner
and environment-restoration attempts remain separately retained. The preceding
runtime's three negative-regression failures remain evidence, not current failures.

Package: `artifacts/nuget/preview-020-alias-cache-20261021/`. Its DLL, symbol PDB and
four caller guides match the qualified checkout; both consumer summaries pin the
same package. Package hashes and exact evidence locations are in
`package-identity.json`. Publication remains deferred.

Latest infrequent workspace scan: **5.361 GB**, 853,244 files, zero errors.
This counts unique file-content lengths, not allocated clusters. It remains near
the loose 5 GB target; no historical purge or protected-target deletion occurred.
