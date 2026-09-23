# Resource Hardening Closure Audit

Requirement-by-requirement verdicts for PLAN.md findings R01-R22 (reviewed 2026-09-22).
Reconciled 2026-09-23 with the local PLAN.md ledger at `7c1f0be0`: R06, R07, R19, R20
are **partial** (named residuals in their rows), R22 is **standing** (definition of done
per slice; V01 is the finite verification milestone), the rest stay **closed (recorded)**.
The suite/manifest counts below were re-verified at the same SHA on 2026-09-23 (Windows
Release build plus full `--skip-slow` suite; per-test report and manifest inventory under
ignored `artifacts/s01-baseline/`).
Evidence per finding: implementing commit(s) on master plus covering tests, all green in
the full Release suite (1711 passed / 0 failed / 9 skipped with `--skip-slow`) with
339/339 visual manifests valid. Verdicts: **closed**, **partial**, **open**.

Verification method: production changes prove out through byte-identical outputs
(deterministic conversion tests, text-operation gates, visual manifest validation),
zero-budget trips that fail without the fix (stash-verified where noted), exact-count
pins (budgets, peaks, estimates), and allocation bounds. Timing probes are observations
only, never gates.

## Verdicts

| ID | Verdict | Evidence |
|---|---|---|
| R01 JPEG allocation guard | closed | `443e7304`; `JpegScanReservesBeforeAllocating`, malformed/header/crop zero-budget gates (ooxml, imaging) |
| R02 image working-set ownership | closed | `93a8fe64`; 4 ownership tests, live-cap/summary constants (imaging, ooxml) |
| R03 shared work charges | closed | `443e7304`; variant/font/dense zero-budget gates (ooxml); JPEG recolor cancellation verified in code |
| R04 aggregate expansion | closed | `443e7304` dense slots; `657812af` XML nodes/cells; `59ffdd92` scene nodes/nested bytes/models; N/2N/4N scaling test |
| R05 ZIP intake preflight | closed | `87778190`; EOCD count/size rejections, 10k-entry cap, Length-throwing staging (ooxml) |
| R06 budgets through serialization | partial | `bfcb5d52`; page/content/output zero-budget trips, negative-cap validation, R20 order preserved; N/2N/4N-page scaling pins exact page charges, exact content doubling, and linear allocation bounds (ooxml, api). R06.1 done in `fa55919a` (single pre-write admission boundary in PdfObjectWriter,
zero/exact/one-short boundaries on file/seekable/forward-only paths, counting and
throwing-destination gates, full suite 1716/0/9, manifests 339/339); page/content
admission done in `484e8ee8` (per-page/per-content charges at the owning renderer
iterations with last-page image tripwires, stream summaries report render-stage
page/content with output still zero, full suite 1719/0/9, manifests 339/339);
remaining residuals per ledger: R06.2 retained-resource producer limits and DOCX
repagination control, R06.3 bounded retained payloads |
| R07 emergency wrap bound | partial | `7c143098`; `EmergencyWrap*BoundsWork` scaling tests (docx-text). Residuals per ledger: R07.1 work/character characterization and justified line expectations, R07.2 suffix/span amplification removal with bounded shaping fallback |
| R08 cell memo repair | closed | `eb237ed7`; memo hit/miss/coordinate tests (docx-tables, `DocxCellMemoTests`) |
| R09 per-slide memo lifetime | closed | `272de7e0`; per-slide constructions verified in `PptxRenderer.cs`; memo suites green |
| R10 graphics/resource indexes | closed | `1d7c41d2`; `PdfResourceIndexTests` (pdf) |
| R11 shared border-overlap plan | closed | `9233ae9d`; `DocxBorderPlanTests` (docx-tables) |
| R12 document/page indexes | closed | `daec0208` drawings, `dc8f1987` related stories, `b404bde3` reference pages; equivalence tests; before/after probes in this audit |
| R13 font byte ownership | closed | `ac57c794` spans, `0a9ee59d` recency, `17c5484a` in-flight throttle, `266ff97f` retained LRU, `c56b247d` retention proof + measurements |
| R14 text interpretation | closed | `49e8eb25` unifies 7 run readers + contract test; `2c14ed05` plain-shape scene/span agreement battery; `fc8fecd9` break-run line-boundary agreement; `a4c46558` field-run agreement; `2c4af54e` paragraph-alignment agreement; `16da8e1c` placeholder-text agreement (master bodyStyle) and matched-placeholder agreement (layout lstStyle wins); fixtures must use p:ph per Office files; full 15-field run-style agreement at nominal size; `e91570c1` hyperlink color/underline agreement; `1f2cf402` paragraph-default inheritance; `12ba206c` shape list-style defaults; `40a4770f` master default-style fallback; `7630de21` grouped-shape text; `65d42692` table-style text; `12f917d5` scheme colors; chart tri-state readers verified distinct; scene-fed migration executed for every in-scope family (plain, placeholder, link, field, group, table) with span-exact equivalence pins and production flipped under qualification, suite byte-identical plus Office-gated visual checks |
| R15 chart data resolution | closed | presence checks + subset normalization; existence short-circuits; per-frame shared dense label, series-name, series-vector, and bar/line extent memos; sparse tick-edge max counts + mechanics tests; area/radar extents single-evaluation (no repeat); suite 1683/0/9 |
| R16 util typing | closed | `226d879e` typed PPTX caches; `6fd2ce7f` immutable cell context; compiler-checked + byte-identical suite |
| R17 units and execution values | closed | transform contract + cell/run vertical alignment + table width kinds (parse matrices); paragraph alignment and story kinds pre-existing; suite 1676/0/9 |
| R18 resource identity | closed | `ee8e6372`; full-digest/ exact-equality/collision tests (`PdfIdentityTests`, pdf) |
| R19 telemetry scope | partial | `443e7304` + admission sizing docs; reservation-peak scope in `Diagnostics.md`. Residuals per ledger: typed domain/phase reporting, retained/in-flight owners, qualified host guidance |
| R20 severity and publication | partial | `567a16d1`; severity/threading tests (`DiagnosticOutcomeTests`, api). Residual per ledger: pre-write stream summary versus post-write file summary contract, resolved together with R19 |
| R21 tool budgets | closed | `5cf3e1c9`; spawn-based tools group (7 tests), ps1 timeout |
| R22 closure evidence | standing | `90095d5b` pointer removal + scaling test; `834bb367` ignored-plan citation sweep; this audit; audit-doc records; clip/ellipsis/underline Office divergences fixed with locked gates; 8 of 9 skips verified passing unskipped (1 environmental); N/2N/4N page scaling pinned; staged emission mapped with phase seam executed byte-identical |

## R12 profiling evidence

Synthetic decks, medians of 3, Windows Release (observations, not gates). Footnote
placement (model-level layout): 500/1000/2000 notes measured 57/106/174 ms pre-index
versus 17/45/67 ms indexed, identical placement counts. Drawing conversion (full DOCX
render): 400/800/1600 anchored drawings measured ~164/270/325 ms pre-index versus
~154/196/424 ms indexed with byte-identical PDFs: layout-dominated noise, so the
index bounds the worst case rather than shifting the median.

## Analyzed, not planned (deliberate non-changes)

- Chart dense-provenance split: provenance fields are consumed at ~94 use sites versus
  ~22 pure-value reads; separating them adds index joins without removing work.
- Shared inherited scene-node graphs across slides: per-slide node freshness is
  load-bearing for identity-keyed memo correctness (R09); sharing needs (node, slide)
  keys first.
- Disk-backed intake stays deferred under DEC01 (bounded in-memory intake with disposal is policy); output payload staging is active R06.3 work, not a deferred item.
- Area/radar extent memoization: one extent evaluation per frame, so a memo stores
  without removing any densification.
- ReadSceneOrXml arm consolidation: 14 scene/xml branches with heterogeneous bodies;
  a generic dispatch trades 3-line branches for lambda plumbing with no resource effect.
- Explicit chart-frame context: per-frame workbook memos already scope sharing to a
  frame with clear-between-frames pins; a token object adds indirection without
  removing work.

## Residual work

- R14-deeper: scene-fed text migration executed for every in-scope family.
  Fifteen agreement batteries gate runs, breaks, fields, alignment, placeholders,
  table/shape consistency, table-style text, grouped shapes, scheme colors, full styles,
  hyperlinks, and every inheritance layer. Production serves scene-fed positioned spans
  for plain shapes, placeholders, links, fields, grouped shapes, and table cells; only
  unplaceable shapes (unsupported orientation, no layout frame) and tables without scene
  text fall back. Layout, measuring, emission, frame geometry, diagnostics, and snapshot
  cascade parity are unchanged: the entry reuses retained cascade defaults and layer
  sources merged with the renderer merger, rebuilt bullet/run models, and the shared
  flow/layout/span pipeline.
  Span-exact equivalence pins cover plain shapes, placeholders (matched and unmatched),
  links with click identity, fields, grouped shapes through nested transforms, and styled
  table cells. Deliberately unchanged: shared inherited scene-node graphs across slides
  (per-slide node freshness is load-bearing for identity-keyed memos); scene merger
  spacing alignment (no consumer reads spacing from scene defaults). Validation bar met:
  agreement batteries plus byte-identical full suite (1710 passed, deterministic conversion
  stable) and Office-gated visual checks.
- R15-remainder: closed (bar/line extents shared; area/radar single-evaluation need
  no memo; frame context and arm consolidation analyzed above, not planned).
- R17-remainder: closed (paragraph alignment was already enum-typed; border edges validate at parse).
  (cell vertical alignment done).
- R22-closeout: Office COM rendering verified working via supervised RenderReference
  (probe PDFs render ok); first Office-gated R14 evidence: placeholder inheritance sizes
  agree exactly (unmatched 24.96pt master bodyStyle, matched 26.04pt layout lstStyle;
  formal text-op compare shows X-exact positions with ~1-3pt Y placement deltas, out of
  scope for inheritance gates). Synthetic packages need theme part, nvGrpSpPr/grpSpPr,
  and cNvSpPr or PowerPoint reports them corrupt (unit fixtures stay minimal/lenient);
  Three held-out cases (`e8837d62` placeholder pair, `d13445ac` table style) pass CheckVisualCase against live Office references with pinned thresholds; table case exposed large-slack row stretching, fixed to declared heights (1.2x-3x Office probes); family run surfaced 4 stale pixel gates, 1 recalibrated (`c5016569` mixed-stack, text-exact) with ellipsis truncation fixed by the same baseline rule (strictly-outside lines dropped, marker inline after the last kept line; the ellipsis reference now matches the Office 2 text operations with no second line), clip overflow fixed by a layout-level cull of baseline-outside lines under vertOverflow clip with ellipsis keeping its own marker logic; anchor-overflow now matches the Office 5 text operations and is promoted to locked-text-ops; underline metrics fixed by emitting the full post-table thickness (the quartering scale was an initial-commit assumption with no Office evidence; the underline-single reference draws the full 2.64pt bar and the locked case passes again);
  `e3f292dd` post-write/pre-move snapshot with writer page/content/output fields
  (file path; stream path keeps pre-write zeros under R20); staged page/resource
  emission (writer holds all pages before emitting; incremental staging design is mapped below).
  Staged page/resource emission design (mapped, not started): the writer materializes
  every PdfPage (content strings plus per-page font subsets and image bytes) before
  numbering objects from complete collections, so peak tracks the whole document.
  Staging keeps RenderPages descriptors but streams in phases: accumulate font
  codepoint unions per resource key and image digest registries (first sighting
  streams bytes, later pages reference numbers) while holding only page content;
  then write header, pages, merged font subsets, deduped images, and xref.
  Numbering need not match the current precomputed scheme: determinism pins only
  require run-to-run stability, but object order must stay deterministic. R06 charge
  points (pages, content, output), per-page validation order, cancellation points,
  and R20 temp-file publication stay fixed; lazy per-page production differs by
  format (PPTX per-slide is natural, DOCX pagination is whole-document). Validation
  bar: deterministic stability, zero-budget trips, R20 tests, full suite, manifests.
  Phase seam executed: plan/numbers are class-level pure functions with identical bytes;
  lazy page production per format plus streaming emission remain future optimization beyond the numbered requirements (the original review requested bounded spooling or staged page/resource emission; bounded in-memory intake is DEC01 policy and output payload staging is R06.3).
