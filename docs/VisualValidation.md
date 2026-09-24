# Visual Validation

Visual validation exists because a valid PDF structure does not prove visual fidelity. The harness renders an Office-produced reference, renders the candidate PDF with PDFium, compares PNG pages, and creates files that an agent or contributor can inspect.

## Requirements

- Windows with Microsoft Office installed.
- PowerShell.
- .NET SDK.
- `tools/vendor/pdfium/win-x64/bin/pdfium.dll`.
- .NET restore needs network access to nuget.org (or a local feed such as the ignored `artifacts/localfeed/` workaround); without it only prebuilt binaries under `bin/` can run.
- Reference rendering drives Office through COM: PowerPoint closes other presentations and Word runs headless (`DisplayAlerts = 0`); large decks export silently for many minutes with no progress output, and a missing `artifacts/reference-cache/` entry fails cache-only private runs instead of invoking Office.

The repository vendors the exact rasterizer binary (`tools/vendor/pdfium/win-x64/bin/pdfium.dll`, PDFium 150.0.7834.0 per `VERSION`); checkouts must use it as-is. Do not replace it from `releases/latest`: a different build silently changes raster output, and `tools/RasterizePdf.ps1` verifies the binary against the pinned SHA256 in `tools/vendor/pdfium/win-x64/pdfium.sha256` before every run. To upgrade, fetch the versioned release matching `VERSION`, replace the DLL, and update the pin file in the same commit.

## Running A Case

```powershell
pwsh tools/CheckVisualCase.ps1 -Case visual-cases/cases/docx-basic-paragraphs/case.json
```

The script:

1. Builds the CLI and converts the test input to `candidate/output.pdf`.
2. Uses PowerPoint or Word COM automation to render reference output.
3. Rasterizes the candidate PDF through the local PDFium rasterizer.
4. Runs `Lokad.OoxPdf.VisualDiff`.
5. Writes an `assessment.md` template.

Outputs are timestamped under:

```text
artifacts/visual/<case-id>/<run-id>/
```

Important files:

- `reference/page-001.png`: Office reference page.
- `candidate/output.pdf`: generated candidate PDF.
- `candidate/page-001.png`: PDFium rasterization of the candidate.
- `candidate/diagnostics.json`: emitted renderer diagnostics.
- `comparison/metrics.json`: dimensions and pixel-difference metrics.
- `comparison/index.html`: side-by-side browser review.
- `assessment.md`: agent review notes and rating.

## Assessment

Pixel metrics are advisory. Office and PDFium can differ in antialiasing, font hinting, and rasterization details even when the PDF is acceptable. The assessment rating is the authoritative result:

- `5`: visually indistinguishable except tiny antialiasing differences.
- `4`: good; minor spacing, kerning, or antialiasing differences only.
- `3`: usable; visible but non-blocking differences.
- `2`: poor; major layout defects or missing important elements.
- `1`: severe; page or slide mostly wrong.
- `0`: conversion failed or page missing.

Do not commit generated visual artifacts unless they are intentionally small fixtures.

## Family Parity Targets (Q07 triage, 2026-09-22)

Each family has an explicit target classification so needs-review triage has a
destination. Counts below track the manifest state as of 2026-09-24 (triage rounds 1-14 baselined 2026-09-22)
(118 locked, 9 locked-text-ops, 212 approximate, 0 needs-review across 339
cases; family rows overlap on shared patterns so their totals exceed 339).

| Family | Cases | Target | Rationale |
|---|---|---|---|
| pptx-images | 13 / 0 / 1 / 0 | Locked except known-approx effects | Pixel-faithful embedding; effectively there |
| pptx-charts | 28 / 0 / 32 / 0 | Tier-1 ports locked; rest approximate with recorded gaps | Matches the Tier-1/Tier-2 split in Capabilities.md; triage complete 2026-09-22 |
| pptx-typography | 15 / 9 / 71 / 0 | Approximate by default; exact ports lock | Font-metric approximations are structural; text-ops locks pin emission; run-merging fix aligns op grouping with Office |
| pptx-tables | 9 / 0 / 7 / 0 | Mixed; first-pass built-ins lock, rich styles approximate | Per Capabilities table-style scope; triage complete 2026-09-22 |
| pptx-shapes | 21 / 0 / 12 / 0 | Approximate; small preset geometry only | Preset-geometry scope in Capabilities.md; triage complete 2026-09-22 |
| pptx-composition | 10 / 0 / 3 / 0 | Mixed; master/layout inheritance locks case by case | Triage complete 2026-09-22 |
| pptx-effects | 5 / 0 / 4 / 0 | Mixed; raster shadows/glows and gradient approximations render, other effects approximate | Per Capabilities effects scope; triage complete 2026-09-22 |
| pptx-smoke | 6 / 0 / 0 / 0 | Locked | Blank/size discovery must stay exact |
| docx-layout | 11 / 0 / 82 / 0 | Mixed; greedy-wrap approximations stay approximate | Latin greedy wrapping scope in Capabilities.md; triage complete 2026-09-22 |
| docx-markup | 0 / 0 / 33 / 0 (gated via reference cache) | Approximate with margin modes tracked | Reference-cache workflow below |

Columns are locked / locked-text-ops / approximate / needs-review.

### Approximate-family limitations

Named limitation, intended next improvement, and gate for each approximate family. Counts are the 2026-09-24 validated state.

- pptx-images: SVG vectors stay approximate (strip tessellation with diagnosed fallbacks); next is SVG stroke rendering and reference-pixel gradient and transform cases (need Office references). Gate: family 14/14, pptx-images group, structural color assertions.
- pptx-charts: Tier-2 behavior stays approximate (trendlines, secondary axes, leader-line labels, percent stacking) with calibrated unstyled-series shading; next is other chart families and styles with held-out Office calibration. Gate: family 60/60, pptx-charts group, structure, color, and label checks plus pixels.
- pptx-typography: font metrics stay structural approximations with autofit and overflow handling approximate; next is held-out Office calibration for the remaining approximate ports. Gate: family 95/95, pptx-typography group, text operations and line starts plus pixels.
- pptx-tables: rich table styles stay approximate while mixed-run word wrap is fixed; next is held-out Office calibration for rich-style cells. Gate: pptx-tables group, text operations and line starts plus pixels.
- pptx-shapes: small preset geometry only with scene-only custom geometry; next is held-out Office calibration for further geometry kinds. Gate: family 33/33, pptx-shapes group, path-operation assertions plus pixels.
- pptx-composition: master and layout inheritance locks case by case; next is extending locked inheritance coverage. Gate: family gates plus pixels.
- pptx-effects: raster shadows and glows approximate, other effects unsupported with diagnostics; next is held-out Office calibration for further effects. Gate: family 9/9 plus pixels.
- docx-layout: Latin greedy wrapping with body-path inline images and approximate columns, notes, and floating wrap; next is table and related-story atom paths with Office calibration. Gate: family 60/60, docx-text groups, words and line starts plus pixels.
- docx-markup: margin modes tracked with diagnosed fallbacks including word-compatible text; next is word-compatible Office calibration (references unavailable here). Gate: cached Office gates plus layout snapshots.

### Lock policy

- New locks require a reviewed agent rating of 4 or 5, tight pixel gates at
  roughly 2-3x observed headroom (MAE, changed-pixel ratio, foreground recall),
  empty diagnostics where the manifest demands it, and passing family support
  gates (SSIM/histogram correlation where the family defines them).
- The 2026-09-21 corpus predates this bar (uniform rating 3, recall thresholds
  often absent); those locks are grandfathered because their pixel gates do the
  real work. A 2026-09-22 audit confirmed all 59 then-locked cases reference
  committed inputs.
- Promotions in this round: `pptx-ladder-11-chart-bar-clustered-port`,
  `pptx-ladder-11-chart-pie-5-categories-port`,
  `pptx-ladder-11-chart-column-clustered-port`, and
  `pptx-ladder-11-chart-bar-stacked-port` to locked (indistinguishable renders,
  all metrics tight including family gates, empty diagnostics);
  `pptx-ladder-11-chart-line-markers-port` to approximate (marker-fill shade
  gap from the documented 0.975x unstyled-series approximation; fails family
  histogram correlation at 0.77); `pptx-ladder-11-chart-line-3series-port` to
  approximate (same unstyled-series shade family, no markers; histcorr 0.79).

## Cached DOCX Markup References

DOCX markup parity work uses cache-only reference comparisons so autonomous runs do not launch Word or COM. Generate a reference request for the public DOCX markup cases:

```powershell
pwsh tools/NewDocxMarkupReferenceRequest.ps1 -OutputDirectory artifacts/docx-markup-reference-requests/public
```

Use `-MissingOnly` when refreshing the handoff after some cached references already exist; the request summary records ready and missing counts without launching Office.

After trusted Office PDFs have been produced on a controlled setup, import them with `tools/ImportDocxMarkupReferenceCache.ps1`. Then audit or run the cached gate:

```powershell
pwsh tools/RunDocxMarkupReferenceGate.ps1 -CacheStatusOnly
pwsh tools/RunDocxMarkupReferenceGate.ps1 -FailOnDeltas
```

The cache-status output includes `missing-import-commands.ps1`, which lists only the references still absent from `artifacts/reference-cache/`.
