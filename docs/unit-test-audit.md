# Unit Test Audit

This audit tracks the shift to Office-PDF-first fidelity work. Unit tests remain valuable, but renderer tests should not freeze candidate-specific PDF operators when Office uses a different observable PDF strategy.

## Keep As Unit Tests

- Public API behavior, CLI exit codes, diagnostics, deterministic bytes, package safety, XML hardening, image/font parsers, and low-level PDF writer primitives.
- Renderer smoke tests that assert content is not omitted: page count, media box, font/image resources, diagnostics, and broad draw-order invariants.

## Move Toward Visual Gates

- Exact renderer text coordinates such as `Tm` assertions.
- Exact shape path coordinates when the feature already has or should have a public visual case.
- Candidate-specific text operators such as `Tj` when Office uses `TJ` or different text grouping (one instance closed 2026-09-22: same-style source runs now merge into Office-compatible operations, migrating `PptxSyntheticTextBoxMergesSameStyleEmphasisSourceRuns` and the para-0 glyph-run assertions; hyperlink/math/autofit guards and hidden-advance asserts unchanged).
- Candidate-specific clipping rectangles and synthetic stroke/fill placement when Office PDF inspection shows a different composition strategy.

## Initial Evidence

`pptx-ladder-02-plain-text` Office reference from `artifacts/visual/pptx-ladder-02-plain-text/20260515-002913/reference/reference.pdf`:

- Office content stream uses marked content, an explicit slide clipping rectangle, graphics states, `/F1 24 Tf`, `1 0 0 1 72 444.58 Tm`, and `TJ` glyph positioning.
- Candidate content stream uses a simple clipping rectangle, `/F1 24 Tf`, `1 0 0 1 72 445.2 Tm`, and one `Tj`.
- The visual gate passes tightly: MAE `0.028749`, changed-pixel ratio at threshold 16 `0.000407`.

Implication: the public visual gate is the right lock for this feature. Unit tests should verify that text renders and fonts are embedded, but should avoid treating the candidate's current exact text matrix as the source of truth.

## S01 Execution Baseline (2026-09-23, `7c1f0be0`)

Fresh Release build plus full `--skip-slow` suite at the PLAN.md ledger baseline SHA,
Windows 10.0.26200 x64, .NET SDK 10.0.300-preview / runtime .NET 10.0.12, workstation GC
(default; no server-GC runtimeconfig), Windows fonts via `WindowsFontResolver`
(`WindowsFontResolverFindsInstalledFonts` passed), library version 0.1.5:

```powershell
dotnet build Lokad.OoxPdf.slnx -c Release --no-restore
dotnet run --project tests/Lokad.OoxPdf.Tests -c Release --no-build -- --skip-slow --report artifacts/s01-baseline/tests.log
pwsh -NoProfile -File tools/ValidateVisualCases.ps1
```

- Build: 0 errors (7 pre-existing warnings); suite **1711 passed / 0 failed / 9 skipped**
  (8 slow-filter skips plus 1 environmental: `PptxPrivateLayoutDiagnosticWhenRequested`
  needs private input); exit code 0.
- Visual manifests: **339/339 valid** (locked=118, locked-text-ops=9, approximate=212).
- Per-test JSON report, stdout log, exit code, build log, manifest log, `dotnet --info`,
  and HEAD pin live under ignored `artifacts/s01-baseline/` (not tracked evidence).

## Quantified Inventory (historical: refreshed 2026-09-22, Release binary 1672 passed / 0 failed / 9 skipped with --skip-slow)

- `PptxTests.cs`: 63 ` Tm`/`Tj`/`TJ` assertion hits; `DocxTests.cs`: 12 (2026-09-03 survey; recount before scheduling rewrites). These are the freeze-risk surface.
- Pilot conversion pattern (do not bulk-rewrite yet): replace `AssertContainsTextMatrixAtX(pdf, 72d)`-style exact-matrix checks with smoke assertions (page count, media box, `Tf` font resource present, `Tj`/`TJ` text drawn, diagnostics empty) and rely on the matching `visual-cases/` lock (e.g. ladder typography ports, `pptx-ladder-02-plain-text` MAE 0.028749 precedent).
- `TestCatalog.cs` name-substring classification and the hand-maintained slow list in `TestRunner.cs` (8 entries) stay as-is; the `--group` filter now scopes runs by capability family. No grouped-attribute rewrite pass is scheduled.
- The 2026-09-03 network-restore blocker no longer applies: Release builds and the full suite verify with cached packages (`dotnet build ... --no-restore`; the SourceLink build dependency resolves from cache). The rewrite list below is unblocked and proceeds family by family behind visual gates.

## First Rewrite Candidates

- PPTX text layout tests with exact `Tm` expectations: body insets, line breaks, tabs, explicit tab stops, large-text baseline, mixed-run centering/wrapping, list-style defaults, empty paragraphs, vertical anchoring, placeholder bounds.
- DOCX layout tests with exact `Tm` or rectangle expectations: exact line height, row heights, table geometry, paragraph styling.
- Shape/table tests with exact path coordinates should remain only until equivalent public visual cases are created or tightened.

## Resource-Budget Test Coverage (PLAN R01-R06, R13, R21)

- Budget boundaries close on observable evidence, not counters alone: zero-budget
  conversions trip end to end without publishing (images, fonts, dense slots, pages,
  content, output, scene nodes, nested packages, workbook models, XML nodes, cells),
  and re-access after eviction re-reads byte-identical programs (image pixels, font
  programs, pack downloads, snapshot files).
- Aggregate accounting scales linearly by assertion (N/2N/4N XML documents charge
  5N objects), and reserved-versus-retained lifetimes are pinned (live image bytes
  held across owner lifetime, released on dispose and on decode failure).
- Allocation observations (Windows, .NET 10, Release, 64-bit, workstation GC, not
  gates): synthetic directory discovery (1 TTC with trailing dead space + 1 TTF,
  3 faces) allocates ~40 KB; one warm minimal-DOCX conversion allocates ~4.7 MB
  with no growth across repeats (4,705,856 then 4,681,448 then 4,680,112 bytes) and
  byte-stable output.
- Repeated conversions reuse snapshot-retained font bytes: overwriting the font
  file after the first conversion leaves subsequent outputs byte-identical (the
  custom font is proven embedded, so the check is not vacuous).
- Layout/render index profiling (Windows, Release, medians of 3, not gates):
  footnote placement over synthetic note decks measured 57/106/174 ms pre-index
  versus 17/45/67 ms indexed (500/1000/2000 notes, identical placement counts);
  full-DOCX conversion over drawing decks measured ~164/270/325 ms pre-index
  versus ~154/196/424 ms indexed (400/800/1600 anchored drawings, byte-identical
  PDFs), i.e. layout-dominated noise with the index bounding the worst case rather
  than shifting the median. The reference-to-rendered-page index then removed the
  remaining per-check all-pages scans.
- The 9 skips are 8 slow-filter skips plus 1 environmental precondition
  (`PptxPrivateLayoutDiagnosticWhenRequested` needs private input), unchanged by this work.

## Remaining Gaps

- Office-provenance visual runs still need Office/COM renders for new held-out
  geometry/font/page variations; manifest validation here verifies inventory only.
- The page/content/output budget counters enforce limits but do not yet report in
  CONVERSION_RESOURCE_SUMMARY; staged page/resource emission stays open.
- Exact-matrix/path freeze-risk surface above is unchanged by the budget work.
