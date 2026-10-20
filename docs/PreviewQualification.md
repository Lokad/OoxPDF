# 0.2.0-preview qualification — 2026-10-20

The local preview candidate is prepared; publication, tagging and release
creation remain deferred. [Scope and limitations](PreviewRelease.md) and
[caller migration](PublicApiChanges.md) accompany the package.

The canonical runtime comes from committed source `6f750f6e` and
[green Windows/Ubuntu CI](https://github.com/Lokad/OoxPDF/actions/runs/39404138180).
Its library SHA-256 is
`3291600CC47004C77CDD268E86B53FA6EA68D7213D3BDC81D535CE56F09124BC`.
Preparatory local and Ubuntu compiler outputs have different hashes and are
kept separate from this runtime's qualification.

| Gate | Result |
|---|---|
| Windows broad console suite | 3,712 passed, 0 failed, 16 skipped in CI |
| Focused CLI exit policy | Four additional exact checks pass, including strict warning exit 3; frozen CLI, no build |
| Same frozen runtime on Ubuntu | 792 passed, 0 failed, 43 skipped; one exact 835-method invocation |
| Final package consumers | Windows and Ubuntu load the same packed DLL; version metadata, font embedding and admission ownership pass |
| Representative documents | Nine sources, sixteen outcomes, 32 pages; every PDF and diagnostic list matches its accepted parent and every Office page count matches |
| Independent Office pixel gates | Two PPTX passes; fourteen DOCX failures retained, without threshold changes |
| Content screening | Fourteen clean; two static-report field flags retained and independently classified as cached-field uncertainty |
| Root transfer | All 17 runtime assets verified; two integration PDFs match; no root rebuild or broad-suite replay |
| Safe managed contract | Explicit unsafe prohibition; IL-only DLL, zero P/Invoke methods or imported COM types; no runtime package dependencies |

Final artifacts are under ignored `artifacts/nuget/preview-020/`:

- `Lokad.OoxPdf.0.2.0-preview.nupkg`, SHA-256
  `2F2DD58B54AE60A3916E858A09DD0900747B736CDCF1235D391B5E41B1E55041`.
- `Lokad.OoxPdf.0.2.0-preview.snupkg`, SHA-256
  `C5731E734D6D705284A9F795B3E886924F1D689ADB4B3B159F57420F7B3CD2E6`.

Final packaging refreshes documentation from `a47aaf78` without compiling the
library or PDB. Both match the canonical CI artifact. Package and informational
version are `0.2.0-preview`; assembly/file version is `0.2.0.0`. The final package
was consumed on both platforms after packing. Evidence/helper-only commits skip
CI to avoid replaying unchanged rendering qualification; library/test Git trees
and all transferred runtime assets still match the qualified source.

The follow-up audit removes 26 unused or test-only production methods, two
workbook wrapper records and spill-path test storage. Tests now exercise live
streaming, range metadata, text targets and pattern resources. Two shared
fixture writers no longer reserve orphan temporary files: both canonical payload
probes pass with zero orphan `.tmp` files. Existing S1–S4 API/quota/cache cleanup
remains accepted. Useful inspectors, bounded caches, coarse host admission,
cancellation, input bounds, deterministic cleanup, publication guarantees and
PDF structural validation remain.

Retained limitations include the SVG report's Review prose Y error of 1.436
points, the linked-footnote Review raw Word annotation mismatch of 46.794
points, and constrained chart-color/balloon/layout rules. The native note link
matches its physical picture within 0.000016 points. Static-report cached `1`
tokens differ from both printed outputs; actual Word/candidate counters agree,
with no Office-to-candidate token deficit. Raw failures remain failures.
No arbitrary-document ink/clipping guarantee or performance improvement is
claimed. Feature expansion and optimization are deferred.

Detailed identities, source/deletion inventories, raw residuals, platform
reports and consumer evidence are retained under
`artifacts/composition-c1/preview-020-*`, notably
`preview-020-canonical-qualification.json`. Workspace cleanup is separate from
the completed rendering/package qualification; no suites were replayed for it.

The subsequent [32-page paired review](PreviewVisualReview.md) records specific
practical effects for every portfolio document. It found no serious clipping,
overlap or wrong page split at page scale, while retaining all original failures.

After automatic review rejected agent execution, the user ran the requested
throwaway cleanup script. C365 verifies all 3,128 inventoried compiled files
absent across 184 superseded snapshots, with no errors. Their snapshot identities,
all 17 canonical and 17 root runtime assets, 2,582 canonical inputs and both final
packages retain their pinned hashes. Earlier rejected retirements were not retried.

The fresh workspace check measures 15.212 GB, down 2.703 GB from C364's 17.914 GB.
It counts unique file-content lengths by file ID, rather than allocated clusters.
The loose 5 GB target remains unmet, but the old 7.074 GB protected-content floor
is superseded by the user's completed batch. Evidence is retained in
`preview-020-completion-review/post-user-retirement-C365.json` and
`artifacts/workspace-footprint-C365.json`. Preview publication remains deferred.
