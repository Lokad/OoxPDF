# 0.2.0-preview qualification — 2026-10-21

The seventh bounded maintenance milestone is qualified with the retained failures
below. Publication, tags, releases, performance optimization and broader rendering
expansion remain deferred. See [scope](PreviewRelease.md), [migration](PublicApiChanges.md)
and the retained [document review](PreviewVisualReview.md).

One Release production/test artifact was built and frozen at
`ece0b9e67b7f47f3df9c2fa90c0a1b8405ac8094`, with 24 assets and 2,582 shared public
inputs. Tool-only repair `23e00dbf63ff0ad2dd12a5c2ba1982f734313621` clears the expected
native-refusal status after successful raster checks. The corrected script passes
in CI's caller scope against the same validator. No local library rebuild or
broad-suite replay followed this harness repair. No performance claim.

Library SHA-256:
`01E8D8588ED6F908F6A6EB10C42E11D4D98592F94D95092F0DE2F676AB7450F8`.
[Windows and Ubuntu CI passed](https://github.com/Lokad/OoxPDF/actions/runs/39593928460)
at `23e00dbf`. The first run's Ubuntu job passed; Windows ran all ten raster
assertions successfully but failed the step because the expected direct-tool
exit 2 leaked to its caller. That unsuccessful run and its log remain evidence.
CI outputs are independent of the locally qualified DLL.

| Gate | Result |
|---|---|
| Release build | Passed; two existing nullable warnings in unchanged DOCX test code |
| Windows milestone | 3,738 pass, **1 inherited fail**, 5 skips across 3,744 unique methods |
| Same frozen DLL on Ubuntu | 823 pass, 0 fail, 43 environmental skips across 866 unique methods |
| Exact affected capabilities | 41 methods: Windows 39 pass / 2 explicitly declared filesystem skips; ext4 Ubuntu all 41 pass |
| Test reuse | Each platform's affected methods excluded from its remaining milestone selection; two disjoint invocations, no unchanged-method replay |
| Directory publication | Windows 12 pass / 2 symbolic-link privilege skips; Ubuntu 14 pass. Junction/symlink ancestors, targets with linked parents, input/PDF/diagnostics refusal, uncreated outputs and allowed distinct/hardlink destinations |
| Parent regression | Retained Windows/Ubuntu probes show the preceding DLL overwrites through directory aliases. New zero-adjustment scan test fails with its retired work ceiling |
| Raster freshness | Ten real-tool checks pass, including shrinking-page reuse, direct-tool refusal, preserved existing files, fresh shorter output, spaced paths and failure/timeout; caller-scope harness repair passes against identical binaries |
| Reference cache | Twelve existing completion/recovery/provenance controls pass; earlier independent Windows/Ubuntu cache and supervision evidence remains retained |
| Existing raster identity | Two pages from the earlier validator remain byte-identical at 144 DPI after the new validator is transferred; no accepted reference was changed |
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

File-path guards now resolve directory-link ancestors, including Unix targets
whose parents still contain links. Input/output/diagnostics aliases refuse before
publication. Final output file links remain separate replaceable entries. Cyclic
or inaccessible links refuse conservatively; external concurrent directory changes
are outside this path contract. Existing case-only protection remains: provably
distinct entries are allowed; uncreated case-only outputs and case-only parent
ambiguity refuse. Captured paths, callbacks, cancellation and atomic staging remain.

The raster wrapper and direct validator refuse destinations that contain PNGs.
Use fresh output storage instead of reusing a visual run ID. Existing files are
preserved, so a shorter or failed render cannot silently consume old success
pages. Existing complete reference derivatives and fresh cache staging remain.

Kerning no longer decrements a scan-work ceiling or passes a ref fuel counter
through parser helpers. Repetitive scans observe cancellation; pair, coverage and
class storage bounds remain. Signed/last-wins, source ownership, truncated-prefix
and fallback neighbors pass. The review demonstrated a policy inconsistency on
synthetic data, not a real-document font failure or a performance improvement.

Completed prior cleanup stays complete: WriteBlank and other verified dead helpers
remain deleted; cache ownership, worker argument handling and exact-test reporting
remain shared. There is no new public signature removal, rendering rule, ownership
layer, quota or validation framework. Safe managed execution, useful caches,
coarse admission, structural bounds, bounded staging and PDF validity remain.

Existing chart point-count, five-paragraph balloon and font/size/document-shape
rules remain live, documented debt. No new independent Office evidence justified
replacing them. Strict DOCX failures remain significant geometry/typography
residuals: SVG Review Y 1.436035pt exceeds 1pt; linked-footnote raw Word URI differs
by 46.794pt although candidate picture/link bounds agree within .000016pt. Two
field/extraction flags remain. None were relabeled cosmetic or passed.

## Retained candidate evidence

Ignored evidence: `artifacts/preview-directory-raster-20261021/qualification.json`;
frozen runtime under `candidate/`. The first failed CI run, corrected caller-scope
raster evidence and an initial missing-manifest attempt (before any conversion)
remain retained. Preceding path/raster witnesses remain under
`artifacts/review-seventh-20261021/`. Earlier milestone evidence is unchanged.

Package: `artifacts/nuget/preview-020-directory-raster-20261021/`. Its DLL, symbol PDB and
four caller guides match the qualified checkout; both consumer summaries pin the
same package. Package hashes and exact evidence locations are in
`package-identity.json`. Publication remains deferred.

Latest infrequent workspace scan: **5.454 GB**, 857,242 files, zero errors.
This counts unique file-content lengths, not allocated clusters. It remains near
the loose 5 GB target; no historical purge or protected-target deletion occurred.
