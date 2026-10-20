# 0.2.0-preview qualification — 2026-10-20

The bounded review follow-up is complete and the local preview is prepared.
Publication, tagging and release creation remain deferred. See
[scope and limitations](PreviewRelease.md) and [caller migration](PublicApiChanges.md).

One committed Release runtime is frozen from `01827c405b39c2337710c4e80ca1db49cb61436f`.
Library SHA-256:
`125016495C861D422F66A3E713E88DF6568C9A2CFBE793BBCDE21E76FC13D3BB`.
Test assembly SHA-256:
`AE585C6721E019458817D5BE0963AD95EE36F22C5322D6D96BE11416AC43779D`.
Both [Windows and Ubuntu CI jobs](https://github.com/Lokad/OoxPDF/actions/runs/39452164339)
passed their independent builds and checks. CI compiler outputs are not conflated
with this locally frozen runtime.

| Gate | Result |
|---|---|
| Affected capability | All 48 pass, selected specifically for the changed font, stream, image, text and comment paths; extracted from the one broad Windows invocation |
| Windows broad console run | 3,719 pass, **1 fail**, 11 skip; inherited balloon-wrap failure described below, not relabeled as a pass |
| Same frozen runtime on Ubuntu | 796 pass, 0 fail, 43 environmental skips; one exact 839-method invocation, no library rebuild |
| Resolver regression controls | Both new concurrency tests fail against the previous qualified DLL with dictionary-state corruption and pass against the new DLL |
| Qualification input lifecycle | Nineteen checks on Windows and nineteen on Ubuntu; interrupted copying, publication race, immutable-input ownership and corruption rejection |
| Final package consumers | Windows and Ubuntu consume the same final package; qualified DLL identity, version metadata, font embedding and admission ownership pass |
| Representative documents | Nine sources, sixteen outcomes, 32 pages; every PDF and diagnostic list equals the accepted baseline, with matching Office page counts |
| Independent visual comparisons | Sixteen comparisons carried by PDF identity; two PPTX strict passes and fourteen DOCX strict failures retained; no new Office exports, rasters or relaxed thresholds |
| Root transfer | All 17 assets verified; two integration PDFs match; no root build or suite replay |
| Safe managed contract | AllowUnsafeBlocks=false, IL-only DLL, zero P/Invoke or imported COM types, no runtime package dependencies |

The Windows balloon assertion observes a 12.61pt single-row body where its
Office-derived expectation requires a second row. A separate invocation with the
previous qualified DLL fails identically on this machine. Previous CI skipped it
because Aptos was absent. Its cause remains unresolved. The assertion, font guards
and rendering policy are unchanged; balloons remain approximate. The broad local
run therefore has a retained failure, even though all affected checks and CI pass.

## Changes and compatibility

Both built-in resolvers synchronize their existing bounded request dictionaries;
cache hits, publication and capacity eviction share one critical section. Logical
dictionary corruption is reproduced; font-program byte corruption is not shown.
Stream documentation and regressions now distinguish incomplete caller-owned
output from atomic file publication. Failed stream output must be discarded.

Six unused/test-only production methods, one record and one obsolete arithmetic
test were removed. Useful tests now call live image, comment-palette and font
resolution implementations. Useful inspectors remain. Staging now uses independent
copies per changed input tree and cleans only its invocation-owned directory;
frozen runs still share one immutable input snapshot per tree. Cross-tree alias
bookkeeping and stale accounting comments were removed. The operating envelope
retains measured workload facts and links to history instead of rendering journals.
No additional public API was removed in this follow-up; earlier breaking changes
remain covered by migration guidance.

## Local artifacts and residuals

Final packages are in `artifacts/nuget/preview-020-review-20261020/`:

- `Lokad.OoxPdf.0.2.0-preview.nupkg`, SHA-256
  `4E09172DC1ABB42C1A9EA64C341CAFADFF5638AA325748112FF3FB7BC4E7CFB0`.
- `Lokad.OoxPdf.0.2.0-preview.snupkg`, SHA-256
  `69145F1EDC3752E81AEE8EC207377A897D0C6DD2971B54760EBBCE31190B1937`.

Packaging refreshes caller documents from `bb148f8b` without compiling the tested
library/PDB. Package and informational version remain `0.2.0-preview`;
assembly/file version remains `0.2.0.0`. All four bundled caller documents and
both final-package consumers are verified. The earlier `6f750f6e` runtime/package
qualification remains historical under `artifacts/composition-c1/preview-020-*`.

Fourteen DOCX strict visual failures include geometry and typography differences,
not just cosmetic differences. SVG Review prose Y remains 1.436pt against the
original 1pt gate; raw linked-footnote Word URI difference remains 46.794pt although
native picture/link bounds agree within .000016pt. Two raw cached-field screens
retain their independent classification. Chart-color, five-paragraph balloon and
font/document-shape regimes have no new independently verified general replacement;
none was expanded. No supported-scope content or pagination regression is identified
by the fixed portfolio. This is not an arbitrary-document ink/clipping guarantee.

Full identities, deletions, reports, expected parent failures and failed harness
attempts are in `artifacts/review-follow-up-20261020/qualification.json`. The latest
workspace scan reports **4.852 GB** with zero errors, before the small final package
and transfer refresh. No historical artifacts were removed in this batch. Continue
infrequent checks against the loose 5 GB cap. Performance optimization and broader
feature expansion remain deferred.
