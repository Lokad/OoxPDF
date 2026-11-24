# Visual Cases

Public DOCX and PPTX visual cases use Office reference exports and are grouped by capability.

- `cases/<id>/case.json`: one public synthetic visual gate.
- `families/<id>.json`: capability-family manifest using case-name patterns.

Run the tools with PowerShell 7.

```powershell
pwsh -File tools/CheckVisualFamily.ps1 -Family pptx-typography -List
pwsh -File tools/CheckVisualFamily.ps1 -Family pptx-typography -CasePattern "*spacing*" -Tag needs-review -List
pwsh -File tools/CheckVisualFamily.ps1 -Family pptx-shapes -Limit 3
pwsh -File tools/CheckVisualFamily.ps1 -Family pptx-images -OnlyUnsupported -UpdateCatalog
pwsh -File tools/CompareVisualReports.ps1 -Baseline artifacts/visual/reports/pptx-images-old.json -Current artifacts/visual/reports/pptx-images.json
pwsh -File tools/ValidateVisualCases.ps1
pwsh -File tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-ladder-04-font-size-port/case.json
pwsh -File tools/ComparePdfTextOperations.ps1 -Reference artifacts/.../reference-inspect/text-operations.json -Candidate artifacts/.../candidate-inspect/text-operations.json -MatchByPosition
pwsh -File tools/ComparePdfTextLineStarts.ps1 -Reference artifacts/.../reference-inspect/text-operations.json -Candidate artifacts/.../candidate-inspect/text-operations.json -MatchByPosition
```

Family runs write ignored latest and timestamped summaries under `artifacts/visual/reports/`.
Use `-UpdateCatalog` to update the ignored local support catalog at `artifacts/visual/support-catalog.json`.
With `-OnlyUnsupported`, a past supported verdict skips a case only when the
case and family manifests still match. For qualification of a new runtime or
reference, run the selected cases without this historical filter.
Family manifests define strict support thresholds. A case can pass its manifest gates while still being
cataloged as `needs-review` until it reaches the stricter SSIM/color-histogram support bar.

Compare reports from matching source files, case manifests, reference PDFs,
render settings, visual tools and support thresholds. Candidate library versions
can differ. Comparisons between two older unversioned reports are historical.
Regenerate an older report to compare it with a versioned report.

### DOCX families

- `docx-layout`, `docx-markup`, `docx-math`, and `docx-svg`.

### PPTX families

- `pptx-typography`: text layout, fonts, runs, bullets, spacing, highlights.
- `pptx-smoke`: broad legacy smoke cases for basic slides, text, shapes, tables, images, and corporate themes.
- `pptx-shapes`: preset geometry, connectors, arrowheads, joins, adjustments.
- `pptx-images`: pictures, crops, alpha, rotation, flips, picture fills.
- `pptx-composition`: groups, z-order, placeholders, inheritance, mixed slides.
- `pptx-tables`: grid layout, merges, borders, fills, text, margins, styles.
- `pptx-charts`: native chart rendering and parity targets.
- `pptx-effects`: transparency, advanced fills, shadows, diagnostics.

Keep private-document findings out of this tree. Convert them into minimal public synthetic cases first.
