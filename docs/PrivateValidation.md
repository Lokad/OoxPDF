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
