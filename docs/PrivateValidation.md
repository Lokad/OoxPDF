# Private Validation

Use `tools/CheckPrivateCase.ps1` for documents that must not be versioned or made public.

## Layout

Keep private manifests and inputs under the ignored `private-cases/` directory:

```text
private-cases/
  inputs/
    board-deck.pptx
  board-deck.json
```

Example manifest:

```json
{
  "id": "private-board-deck",
  "kind": "pptx",
  "input": "./inputs/board-deck.pptx",
  "dpi": 144,
  "tags": ["private", "pptx"],
  "expected": {
    "minAgentRating": 3,
    "pageCountMustMatch": true,
    "dimensionsMustMatch": true
  },
  "allowedUnsupportedFeatures": []
}
```

## Guard Check

Before running a conversion, validate that the manifest and input are under `private-cases/` and are not git-tracked:

```powershell
pwsh tools/CheckPrivateCase.ps1 -Case private-cases/board-deck.json -ValidateOnly
```

## Run

```powershell
pwsh tools/CheckPrivateCase.ps1 -Case private-cases/board-deck.json
```

Artifacts are written under ignored paths:

```text
artifacts/private-visual/<case-id>/<run-id>/
```

Do not commit private inputs, private manifests, generated reference images, generated candidate images, diagnostics, comparison HTML, or assessments. When recording follow-up work publicly, write only anonymized findings such as feature gaps and page/slide numbers without private text or screenshots.

For DOCX visual comparisons that target a specific review view, private manifests may set `docxMarkup` and `docxMarkupGeometry` using the same values as the public visual-case manifests and CLI. When either value is present, `CheckPrivateCase.ps1` forwards the options to candidate conversion and reads the Office reference from the matching cached variant in cache-only mode. A missing cached reference fails the run instead of invoking Office or COM.

```json
{
  "id": "private-review-doc-all-markup",
  "kind": "docx",
  "input": "./inputs/review-doc.docx",
  "docxMarkup": "all",
  "docxMarkupGeometry": "word-compatible"
}
```

## DOCX Markup Reference Handoff

Do not render Word/Office references on a machine where Word export is not proven headless. For DOCX markup parity work, first create a private-safe reference request:

```powershell
pwsh tools/NewDocxMarkupReferenceRequest.ps1 -IncludePrivate -OutputDirectory artifacts/docx-markup-reference-requests/current
```

The request records case ids, input hashes, markup modes, cache variants, expected cache keys, cache readiness fields, and import command templates. Private entries use placeholders such as `<private-case-manifest>` and do not record private filenames or document text. Add `-MissingOnly` when refreshing the request after some references have already been imported.

On a controlled Office setup, export the requested Word PDFs for `final`, `original`, `simple`, and `all` markup modes. Back on the development machine, import each trusted PDF into the ignored cache:

```powershell
pwsh tools/ImportDocxMarkupReferenceCache.ps1 -Case private-cases/review-doc.json -ReferencePdf artifacts/private-refs/review-doc-all.pdf -DocxMarkup all -DocxMarkupGeometry word-compatible -Dpi 144 -Force
```

The import stores `reference.pdf`, rasterized pages, and `reference-metadata.json` under `artifacts/reference-cache/`. Committed files must contain only private-safe hashes, counts, and anonymized feature classes.

To audit cache readiness without rendering or copying references:

```powershell
pwsh tools/RunDocxMarkupReferenceGate.ps1 -CacheStatusOnly -IncludePrivate
```

The cache-status output includes `missing-import-commands.ps1`. Private rows use placeholders and markup overrides rather than private manifest paths.

After the required cache entry exists, run the cached comparison:

```powershell
pwsh tools/CompareCachedDocxMarkupReference.ps1 -Case private-cases/review-doc.json -DocxMarkup all -DocxMarkupGeometry word-compatible -CaseId private-docx-markup-all -PrivateSafeSummary
```

## DOCX Markup Candidate Check

For markup-heavy DOCX files with comments or tracked changes, use the candidate-only checker until Office reference rendering is configured to run headlessly:

```powershell
pwsh tools/CheckPrivateDocxMarkup.ps1 -Case private-cases/review-doc.json -ValidateOnly
pwsh tools/CheckPrivateDocxMarkup.ps1 -Case private-cases/review-doc.json -MarkupMode simple
pwsh tools/CheckPrivateDocxMarkup.ps1 -Case private-cases/review-doc.json -MarkupMode all
```

The DOCX markup checker builds the CLI, converts only the candidate PDF, writes diagnostics, runs DOCX/PDF inspection, and records revision/comment feature counts. It does not invoke Office, COM, printer drivers, reference rendering, rasterization, or visual comparison.

Record only private-safe metrics in public notes: page count, diagnostic ids, revision element counts, comment marker/body counts, inspection summary counts, and anonymized feature gaps. Do not copy private text, filenames, author names, screenshots, or generated PDFs into public issues or docs.
