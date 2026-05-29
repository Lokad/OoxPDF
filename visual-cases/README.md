# Visual Cases

Visual cases are organized like `pptx-renderer`: small Office-oracle cases grouped by capability family.

- `cases/<id>/case.json`: one public synthetic visual gate.
- `families/<id>.json`: capability-family manifest using case-name patterns.

Useful commands:

```powershell
powershell -ExecutionPolicy Bypass -File tools/CheckVisualFamily.ps1 -Family pptx-typography -List
powershell -ExecutionPolicy Bypass -File tools/CheckVisualFamily.ps1 -Family pptx-shapes -Limit 3
powershell -ExecutionPolicy Bypass -File tools/CheckVisualFamily.ps1 -Family pptx-images -OnlyUnsupported -UpdateCatalog
powershell -ExecutionPolicy Bypass -File tools/CompareVisualReports.ps1 -Baseline artifacts/visual/reports/pptx-images-old.json -Current artifacts/visual/reports/pptx-images.json
powershell -ExecutionPolicy Bypass -File tools/ValidateVisualCases.ps1
powershell -ExecutionPolicy Bypass -File tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-ladder-04-font-size-port/case.json
powershell -ExecutionPolicy Bypass -File tools/ComparePdfTextOperations.ps1 -Reference artifacts/.../reference-inspect/text-operations.json -Candidate artifacts/.../candidate-inspect/text-operations.json -MatchByPosition
powershell -ExecutionPolicy Bypass -File tools/ComparePdfTextLineStarts.ps1 -Reference artifacts/.../reference-inspect/text-operations.json -Candidate artifacts/.../candidate-inspect/text-operations.json -MatchByPosition
```

Family runs write ignored latest and timestamped summaries under `artifacts/visual/reports/`.
Use `-UpdateCatalog` to update the ignored local support catalog at `artifacts/visual/support-catalog.json`.
Family manifests define strict support thresholds. A case can pass its manifest gates while still being
cataloged as `needs-review` until it reaches the stricter SSIM/color-histogram support bar.

PPTX families:

- `pptx-typography`: text layout, fonts, runs, bullets, spacing, highlights.
- `pptx-smoke`: broad legacy smoke cases for basic slides, text, shapes, tables, images, and corporate themes.
- `pptx-shapes`: preset geometry, connectors, arrowheads, joins, adjustments.
- `pptx-images`: pictures, crops, alpha, rotation, flips, picture fills.
- `pptx-composition`: groups, z-order, placeholders, inheritance, mixed slides.
- `pptx-tables`: grid layout, merges, borders, fills, text, margins, styles.
- `pptx-charts`: native chart rendering and parity targets.
- `pptx-effects`: transparency, advanced fills, shadows, diagnostics.

Keep private-document findings out of this tree. Convert them into minimal public synthetic cases first.
