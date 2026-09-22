# Resource Hardening Closure Audit

Requirement-by-requirement verdicts for PLAN.md findings R01-R22 (reviewed 2026-09-22).
Evidence per finding: implementing commit(s) on master plus covering tests, all green in
the full Release suite (1672 passed / 0 failed / 9 skipped with `--skip-slow`) with
336/336 visual manifests valid. Verdicts: **closed**, **partial**, **open**.

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
| R06 budgets through serialization | closed | `bfcb5d52`; page/content/output zero-budget trips, negative-cap validation, R20 order preserved (ooxml, api) |
| R07 emergency wrap bound | closed | `7c143098`; `EmergencyWrap*BoundsWork` scaling tests (docx-text) |
| R08 cell memo repair | closed | `eb237ed7`; memo hit/miss/coordinate tests (docx-tables, `DocxCellMemoTests`) |
| R09 per-slide memo lifetime | closed | `272de7e0`; per-slide constructions verified in `PptxRenderer.cs`; memo suites green |
| R10 graphics/resource indexes | closed | `1d7c41d2`; `PdfResourceIndexTests` (pdf) |
| R11 shared border-overlap plan | closed | `9233ae9d`; `DocxBorderPlanTests` (docx-tables) |
| R12 document/page indexes | closed | `daec0208` drawings, `dc8f1987` related stories, `b404bde3` reference pages; equivalence tests; before/after probes in this audit |
| R13 font byte ownership | closed | `ac57c794` spans, `0a9ee59d` recency, `17c5484a` in-flight throttle, `266ff97f` retained LRU, `c56b247d` retention proof + measurements |
| R14 text interpretation | partial | `49e8eb25` unifies 7 run readers + contract test; `2c14ed05` plain-shape scene/span agreement battery; `fc8fecd9` break-run line-boundary agreement; `a4c46558` field-run agreement; `2c4af54e` paragraph-alignment agreement; `16da8e1c` placeholder-text agreement (master bodyStyle) and matched-placeholder agreement (layout lstStyle wins); fixtures must use p:ph per Office files; full 15-field run-style agreement at nominal size; `e91570c1` hyperlink color/underline agreement; `1f2cf402` paragraph-default inheritance; `12ba206c` shape list-style defaults; `40a4770f` master default-style fallback; `7630de21` grouped-shape text; `65d42692` table-style text; `12f917d5` scheme colors; chart tri-state readers verified distinct |
| R15 chart data resolution | closed | presence checks + subset normalization; existence short-circuits; per-frame shared dense label, series-name, series-vector, and bar/line extent memos; sparse tick-edge max counts + mechanics tests; area/radar extents single-evaluation (no repeat); suite 1683/0/9 |
| R16 util typing | closed | `226d879e` typed PPTX caches; `6fd2ce7f` immutable cell context; compiler-checked + byte-identical suite |
| R17 units and execution values | closed | transform contract + cell/run vertical alignment + table width kinds (parse matrices); paragraph alignment and story kinds pre-existing; suite 1676/0/9 |
| R18 resource identity | closed | `ee8e6372`; full-digest/ exact-equality/collision tests (`PdfIdentityTests`, pdf) |
| R19 telemetry scope | closed | `443e7304` + admission sizing docs; reservation-peak scope in `Diagnostics.md` |
| R20 severity and publication | closed | `567a16d1`; severity/threading tests (`DiagnosticOutcomeTests`, api) |
| R21 tool budgets | closed | `5cf3e1c9`; spawn-based tools group (7 tests), ps1 timeout |
| R22 closure evidence | partial | `90095d5b` pointer removal + scaling test; `834bb367` ignored-plan citation sweep; this audit; audit-doc records |

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
- Disk spooling for intake/output: the explicit policy is bounded in-memory staging
  with disposal; spooling stays open.
- Area/radar extent memoization: one extent evaluation per frame, so a memo stores
  without removing any densification.
- ReadSceneOrXml arm consolidation: 14 scene/xml branches with heterogeneous bodies;
  a generic dispatch trades 3-line branches for lambda plumbing with no resource effect.
- Explicit chart-frame context: per-frame workbook memos already scope sharing to a
  frame with clear-between-frames pins; a token object adds indirection without
  removing work.

## Residual work

- R14-deeper: migrate layout to scene-resolved text one family at a time with
  agreement gates; share context-independent inherited nodes behind (node, slide) keys.
  Entry: ComputeTextSpansForSceneNode re-clones node.Source instead of consuming
  node.TextBody; fifteen agreement batteries gate runs, breaks, fields, alignment,
  placeholders, table/shape consistency, table-style text, grouped shapes, scheme colors, full styles, hyperlinks, and every inheritance layer (paragraph
  defRPr, shape lstStyle, layout bodies, master txStyles, master defaultTextStyle).
  Known migration prerequisites: click identity lives only in renderer models; table
  cells carry unresolved XML (no scene text model). Cell-model design (mapped, not started):
  resolve cell paragraphs via shared scene readers with shape:=cell-txBody (matches the
  renderer, which passes txBody as shape, so ph lookup yields otherStyle on both),
  inherited:=[], sources:=slide chain, plus StyleText threading through run resolution
  (scene readers lack the table-style parameter today). Migration design: build run models
  from TextBody runs reusing layout and measuring unchanged. Validation bar: agreement
  batteries plus byte-identical full suite and visual manifests before Office-gated
  variations.
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
  Two held-out placeholder-inheritance cases (`e8837d62`) pass CheckVisualCase against live Office references with pinned thresholds;
  `e3f292dd` post-write/pre-move snapshot with writer page/content/output fields
  (file path; stream path keeps pre-write zeros under R20); staged page/resource
  emission (writer holds all pages before emitting; incremental staging is a design
  slice, not started).
