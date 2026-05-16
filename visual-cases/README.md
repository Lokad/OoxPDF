# Visual Cases

Visual cases are organized like `pptx-renderer`: small Office-oracle cases grouped by capability family.

- `cases/<id>/case.json`: one public synthetic visual gate.
- `families/<id>.json`: capability-family manifest using case-name patterns.

Useful commands:

```powershell
powershell -ExecutionPolicy Bypass -File tools/CheckVisualFamily.ps1 -Family pptx-typography -List
powershell -ExecutionPolicy Bypass -File tools/CheckVisualFamily.ps1 -Family pptx-shapes -Limit 3
powershell -ExecutionPolicy Bypass -File tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-ladder-04-font-size-port/case.json
```

PPTX families:

- `pptx-typography`: text layout, fonts, runs, bullets, spacing, highlights.
- `pptx-shapes`: preset geometry, connectors, arrowheads, joins, adjustments.
- `pptx-images`: pictures, crops, alpha, rotation, flips, picture fills.
- `pptx-composition`: groups, z-order, placeholders, inheritance, mixed slides.
- `pptx-tables`: grid layout, merges, borders, fills, text, margins, styles.
- `pptx-charts`: chart fallback and future chart rendering.
- `pptx-effects`: transparency, advanced fills, shadows, diagnostics.

Keep private-document findings out of this tree. Convert them into minimal public synthetic cases first.
