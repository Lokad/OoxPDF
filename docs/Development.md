# Development and validation

Install the SDK pinned in `global.json` (.NET 11 RC3) and the .NET 10 SDK/runtime
for compatibility checks. Projects target `net10.0;net11.0`. Use net11 for extensive
development tests and net10 explicitly for compatibility.

```powershell
dotnet build Lokad.OoxPdf.slnx -c Release
dotnet run --no-build --project tests/Lokad.OoxPdf.Tests -c Release -f net11.0 -- --skip-slow
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj -c Release --no-build
```

Release packages/symbols/checksums go directly to ignored `artifacts/nuget/`.
CI covers Windows and Ubuntu. Native validation tools are never library dependencies. Restore PDFium from the
tracked `tools/vendor/pdfium/win-x64/` version and checksum pins; do not substitute
an unpinned latest binary.
See [AGENTS.md](../AGENTS.md) for the script map and workspace conventions.

## Test ownership

| Contract | Primary owner |
|---|---|
| ZIP/XML, relationships, source precedence and bounds | OOXML and format-reader/model tests |
| List formats, counters and effective formatting | Reader/numbering tests with independent literal expectations |
| Advances, wrapping, clipping, page and story ownership | DOCX layout/inline/table/note tests; PPTX scene/text tests |
| Font cache concurrency, leases and shutdown | Font cache/source/resolver tests |
| PDF resource/encoding/nesting and serialization | PDF builder/writer tests |
| Cancellation, refusal, ownership and atomic publication | Public API/resource guarantee tests |
| CLI arguments, diagnostics publication and strict exits | CLI tests and frozen publication script |
| Installed API and package contents | Package consumer and inspection scripts |
| Supported features composed together | Public composition portfolio and portable rendering tests |
| Office appearance | Independent references plus fixed visual/geometry thresholds |

Prefer one semantic owner per rule. Add integration checks for distinct interactions,
not another copy of the same private-helper assertion. Inspections are useful when
they reveal source/layout ownership; helper call counts alone are not correctness.

The console runner accepts `--group`, `--test`, or `--test-set <file.json>` (an array
of exact method names). `tools/RunDocxAffectedTests.ps1` selects capabilities from
`tools/qualification-capabilities.json` and runs each selected method once.
Use `=MethodName` in the manifest or `-TestFilter` to select one exact,
case-sensitive name. Bare selectors match case-insensitive substrings.
`-FastGroup` unions portable groups; declare understood environment skips explicitly.
Keep runner TMP/TEMP in its ignored artifact directory. The qualification command
owns a fresh physical temporary leaf and removes it after complete passing results;
failed runs and cleanup refusals retain it. Keep useful inputs/references outside
that disposable leaf. Direct console runs need caller-owned temporary storage.

## Frozen qualification

Build an explicit committed linked worktree once with
`tools/BuildQualificationArtifact.ps1`. Reuse its CLI, inspector and test assembly
for a coherent batch. Harness-only repairs can reuse unchanged production bytes;
added public fixtures enter a new snapshot, while modified/deleted existing inputs
are refused. Report joins require identical complete input identities and
affected reports that identify the declared frozen library and test assembly.

Use focused affected methods/neighbors during changes. Batch broad Windows/Ubuntu,
framework and package checks at meaningful milestones. Transfer qualified assets
with identity checks and a small integration smoke, without rebuilding or replaying
unchanged suites. Each frozen artifact pins one framework. Do not poll closed CI.

When changing the shared build helper, use `tools/TestBuildIncrementality.ps1`
to check imported MSBuild properties, output provenance and stale DLL refusal.

`TestPackageSmoke.ps1 -PackagePath <nupkg> -Framework <tfm>
-ExpectedLibrarySha256 <hash>` consumes existing package bytes without rebuilding.
Use the same package on both hosts/frameworks. `-InspectOnly` checks metadata and
packaged documentation without restoring or executing a consumer. In PowerShell,
pass matching framework and hash arrays in the same order to inspect the archive
once and run both consumers.

## Visible content and visual comparisons

Public source cases live under `tests/Lokad.OoxPdf.Tests/Cases/`; visual manifests
live in `visual-cases/`. Composition manifests select source/mode outcomes with
explicit content contracts. A validated manifest is not a visual pass.

Check missing words, formulas, labels and pictures independently of page scores:
source ownership/order, actual font-program glyph ink, clips, picture colors,
page targets and consequential overlap. Keep unknown paint traces explicit.
Extraction or a high image score cannot waive content loss. Native tools have
bounded coverage, particularly masks/blends, oblique/CFF/Type1 text and occlusion.

`CheckVisualCase.ps1 -CliPath <qualified.dll> -SkipBuild` reuses a frozen converter;
`CheckVisualFamily.ps1` does the same for capability families. Keep independent
Office references, producer/input/font identities and raw thresholds. Reuse
comparisons only with verified matching identities. PDF/image exports do not
calibrate Office application view; record view/zoom/markup settings separately.

`MeasureCompositionPortfolio.ps1` measures frozen outcomes.
`CompareCompositionPortfolio.ps1` compares pinned Office references. Use
`-InspectorPath <qualified.dll>` to count PDF pages without rebuilding or
rasterizing reused comparisons.
`AuditCompositionContent.py` and `ScreenPdfInk.py` provide bounded independent
content/ink checks. Portfolio joins check the selected modes and content
expectations against the pinned text operations and verify the diagnostic
sidecars. Reinspect existing PDFs with `-ReuseCandidateDirectory` when a join
reports changed expectations or a missing text-inspection hash. Composition limits live in
`tools/CompositionDocxThresholds.json` and `tools/CompositionPptxThresholds.json`.
Reports record their identities and
reapply current limits to reused pixel metrics. Null recall on a matched page
requires a measured zero reference foreground count. Reports distinguish omissions,
pagination/overlap, significant
geometry, cosmetic differences and unknowns. [Capabilities](Capabilities.md) lists
current limitations; development evidence and experiment history stay in ignored
artifacts or Git, rather than public reference pages.

## Private documents

Place private manifests and inputs under ignored `private-cases/`. Run
`pwsh tools/CheckPrivateCase.ps1 -Case private-cases/<case>.json`; outputs stay in
ignored `artifacts/private-visual/`. The script rejects tracked inputs, paths
outside private-cases and unsafe IDs. Public notes record anonymized feature gaps,
never private text, screenshots, filenames or document contents.

Office reference operations require an available authorized environment. Respect
rejected actions. Missing references block their conclusions, not unrelated work.
Preserve source inputs, useful evidence, current qualified artifacts and published
packages when retiring stale generated output. Check the loose 5 GB workspace cap
infrequently with `pwsh tools/MeasureWorkspaceFootprint.ps1`.
