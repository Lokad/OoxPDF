# Visual Cases

Visual cases describe OOXML inputs, expected page counts, and comparison settings for the validation harness.

Run a case from the repository root:

    pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/pptx-blank/case.json

The script writes timestamped outputs under `artifacts/visual/<case-id>/<run-id>/`, including Office reference PNGs, the candidate PDF, candidate PNGs, `comparison/metrics.json`, `comparison/index.html`, and `assessment.md`.

PowerPoint and Word reference rendering require Microsoft Office COM automation on Windows. Candidate PDF rasterization requires `tools/vendor/pdfium/win-x64/bin/pdfium.dll`.
