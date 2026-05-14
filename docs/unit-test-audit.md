# Unit Test Audit

This audit tracks the shift to Office-PDF-first fidelity work. Unit tests remain valuable, but renderer tests should not freeze candidate-specific PDF operators when Office uses a different observable PDF strategy.

## Keep As Unit Tests

- Public API behavior, CLI exit codes, diagnostics, deterministic bytes, package safety, XML hardening, image/font parsers, and low-level PDF writer primitives.
- Renderer smoke tests that assert content is not omitted: page count, media box, font/image resources, diagnostics, and broad draw-order invariants.

## Move Toward Visual Gates

- Exact renderer text coordinates such as `Tm` assertions.
- Exact shape path coordinates when the feature already has or should have a public visual case.
- Candidate-specific text operators such as `Tj` when Office uses `TJ` or different text grouping.
- Candidate-specific clipping rectangles and synthetic stroke/fill placement when Office PDF inspection shows a different composition strategy.

## Initial Evidence

`pptx-ladder-02-plain-text` Office reference from `artifacts/visual/pptx-ladder-02-plain-text/20260515-002913/reference/reference.pdf`:

- Office content stream uses marked content, an explicit slide clipping rectangle, graphics states, `/F1 24 Tf`, `1 0 0 1 72 444.58 Tm`, and `TJ` glyph positioning.
- Candidate content stream uses a simple clipping rectangle, `/F1 24 Tf`, `1 0 0 1 72 445.2 Tm`, and one `Tj`.
- The visual gate passes tightly: MAE `0.028749`, changed-pixel ratio at threshold 16 `0.000407`.

Implication: the public visual gate is the right lock for this feature. Unit tests should verify that text renders and fonts are embedded, but should avoid treating the candidate's current exact text matrix as the source of truth.

## First Rewrite Candidates

- PPTX text layout tests with exact `Tm` expectations: body insets, line breaks, tabs, explicit tab stops, large-text baseline, mixed-run centering/wrapping, list-style defaults, empty paragraphs, vertical anchoring, placeholder bounds.
- DOCX layout tests with exact `Tm` or rectangle expectations: exact line height, row heights, table geometry, paragraph styling.
- Shape/table tests with exact path coordinates should remain only until equivalent public visual cases are created or tightened.
