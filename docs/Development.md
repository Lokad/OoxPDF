# Development and validation

Install the SDK pinned in `global.json` (.NET 11 RC3) and the .NET 10 SDK/runtime
for compatibility checks. Projects target `net10.0;net11.0`. Use net11 for extensive
development tests and net10 explicitly for compatibility.

```powershell
dotnet build Lokad.OoxPdf.slnx -c Release
$testTemp = (New-Item -ItemType Directory -Force artifacts/test-tmp).FullName
$env:TMP = $env:TEMP = $env:TMPDIR = $testTemp
dotnet run --no-build --project tests/Lokad.OoxPdf.Tests -c Release -f net11.0 -- --skip-slow
dotnet pack src/Lokad.OoxPdf/Lokad.OoxPdf.csproj -c Release --no-build -o artifacts/development-pack/
```

The example writes development packages to `artifacts/development-pack/`.
Release packages, symbols and checksums go directly to ignored `artifacts/nuget/`.
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
`-AllDocx` includes the DOCX groups, including `docx-composition-portable`
and `docx-math-portable`, plus shared DOCX and balloon tests.
`-FastGroup` adds fast methods from the named groups. Skips from methods selected
only by a fast group appear under `GroupOnlySkips`. Use `-ExpectedSkip` with exact
method names for the primary selection, including methods also in a fast group.
Qualification requires at least one passing primary method.
Bind direct console runs to a caller-owned temporary parent under `artifacts/`.
The console creates and cleans its own child directory. The qualification command
supplies a fresh parent and removes it when empty after a passing gate. Failed
runs and cleanup refusals retain scratch. Keep inputs and references outside it.

## Frozen qualification

Build an explicit committed linked worktree once with
`tools/BuildQualificationArtifact.ps1`. Reuse its CLI, inspector and test assembly
for a coherent batch. Harness-only repairs can reuse unchanged production bytes;
added public fixtures enter a new snapshot, while modified/deleted existing inputs
are refused. Report joins require identical complete input identities and
affected reports that identify the declared frozen library and test assembly.
Each retained skip must be allowed by its producing selection. A check that
failed in any attempt still needs a later pass.
The freeze records the SDK source revision embedded in each primary assembly.
The test assembly must match the checkout revision. Older production binaries
require matching production sources and build inputs at their embedded revisions.

Use focused affected methods/neighbors during changes. Batch broad Windows/Ubuntu,
framework and package checks at meaningful milestones. Transfer qualified assets
with identity checks and a small integration smoke, without rebuilding or replaying
unchanged suites. Each frozen artifact pins one framework. Do not poll closed CI.

Each hosted candidate archive contains both CLI runtimes and their build
identities. The Windows archive also contains the canonical NuGet package,
symbols and Windows consumer summaries. The Ubuntu package job consumes those
same bytes and publishes its summaries in `package-smoke-ubuntu-latest`.
The console-test artifacts carry the hosted rendering reports. Use a local frozen
build for the test runner and inspectors.

After the hosted jobs pass, bind the package and verify its binaries from the
repository root. Use a stage directory under `artifacts/` containing
`bound-inputs.json`, `hosted-trigger.json` and `hosted-ci.json`. The binding
records the two library and symbol identities and all four consumer receipts.

```powershell
python -B tools/BindHostedPackage.py $stage
pwsh -NoProfile -File tools/VerifyHostedPackageBinaries.ps1 -StageDirectory $stage
```

Use `--offline` with retained artifact metadata and archives to replay the
binding. The stage checks write separate receipts. The package and symbols
are stored under their revision-qualified names in `artifacts/nuget/`.

Run `python -B tools/CheckHostedPackage.py transfer <stage>` on Windows and
Ubuntu to compare the hosted CLI outputs with the qualified local receipts
for both frameworks. After staging, run the same command with `readback`
in place of `transfer` to verify the canonical assets, consumer receipts,
and exact transfer outputs. Readback reuses retained evidence.

Console tests launch prebuilt tools from the matching local build configuration
or an explicitly bound assembly. For frozen tests, bind the tools your selection
uses with `OOXPDF_TEST_CLI_ASSEMBLY`, `OOXPDF_TEST_PDF_INSPECTOR`,
`OOXPDF_TEST_VISUAL_DIFF` or `OOXPDF_TEST_PDF_RASTERIZER`.

When changing the shared build helper, use `tools/TestBuildIncrementality.ps1`
to check imported MSBuild properties, output provenance and stale DLL refusal.

`TestPackageSmoke.ps1 -PackagePath <nupkg> -Framework <tfm>
-ExpectedLibrarySha256 <hash>` requires an existing package and its qualified
library hash. Pack separately with the command above.
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
For cached DOCX markup comparisons, supply selected authored phrases in
`expected.markupTextPhrases`. The report checks their decoded text for the
case's declared markup view and shows whether that check ran. Verify ink separately.
For a frozen raster comparison, pass `-RasterizerPath` and `-VisualDiffPath`
alongside the pinned CLI and inspectors. The cached reference derivative uses
the supplied rasterizer's identity, and one comparison produces both page and
region metrics.
With `-SkipRasterDiff`, cached markup comparisons use the verified reference
PDF and its metadata directly. Cached comparisons require v2 metadata bound to
the input and PDF digests. Flat legacy entries and entries previously adopted
from them require a trusted re-import for the exact input and Office view.

`CheckVisualCase.ps1` and `CheckVisualFamily.ps1` accept `-CliPath`,
`-RasterizerPath`, `-VisualDiffPath` and `-PdfInspectorPath` for prebuilt tools.
Use `-SkipBuild -CacheOnlyReference` to reuse qualified binaries and cached
Office references. The selected rasterizer supplies both page sets and uses
its pinned PDFium dependency in the expected relative location.
Keep independent Office references, producer/input/font identities and raw
thresholds. Reuse
comparisons only with verified matching identities. PDF/image exports do not
calibrate Office application view; record view/zoom/markup settings separately.
The visual-family comparison uses separate reference and candidate page counts.
An inherited threshold failure still needs complete comparison measurements.
Older bound family reports need regeneration before comparison.

`MeasureCompositionPortfolio.ps1` measures frozen outcomes.
`CompareCompositionPortfolio.ps1` compares pinned Office references. Use
`-InspectorPath <qualified.dll>` to count PDF pages without rebuilding or
rasterizing reused comparisons.
`AuditCompositionContent.py` and `ScreenPdfInk.py` provide bounded independent
content and ink checks. Use `AuditCompositionContent.py --screen-pictures` with
PyMuPDF at portfolio milestones. It checks authored PNG colors and the named
SVG fill in rendered regions against pinned Office PDFs. Unsupported witnesses
stay unverified. Use `ScreenPdfInk.py --pdf <pdf> --expected-pdf-sha256 <hash>
--text-witness 1 "Authored heading" --output <report.json>` for selected
phrases in an embedded TrueType paint span. Repeat `--text-witness` for more
page and phrase pairs. The report retains missing paint spans, clipped
glyphs and unresolved states. An unmet witness gives a nonzero exit status.
Use this alongside decoded-text checks and raster review for occlusion,
complex paint or text spread across multiple spans.
Decoded phrase and token checks require a page assignment. The inspector
retains unassigned text, including text in Forms whose invocation it cannot
resolve. Check those phrases with independent page-level evidence.
`compareDecodedTextOperations` requires resolved text on both sides. The
inspector marks unmapped hex glyph codes with a replacement character and
retains the original tokens for diagnosis. Reinspect the PDFs when an older
report may contain guessed text.

Measurement records declared formula and picture expectations. A positive
expectation stays unverified until an audit supplies a witness. Pass the
`AuditCompositionContent.py` report via `-ContentAuditPath` on measurement,
join or comparison. A join retains the report's hash. Comparison checks the
bound Office PDF and can reuse pixel metrics while reporting content misses
separately. Simple formula witnesses check decoded expression text. Native
picture witnesses check the curated paint regions. Exact placement, glyph ink
and unsupported source scopes still need their own review.
For selected field results, declare contextual phrases and page numbers in
`ExpectedFieldResults`. The audit checks each phrase in both PDFs and reports
missing results separately from unresolved field behavior.

Portfolio joins check the selected modes and content
expectations against the pinned text operations and verify the diagnostic
sidecars. Font and graphics summaries stay with the source inspection in
`JoinedFrom.RunDirectory`. Reinspect existing PDFs with `-ReuseCandidateDirectory`
when a join reports changed expectations or a text-inspection hash is missing.
Measurements pin the inspector's logical page count and log. Comparison
recounts older PDFs when their measured count lacks a pin. Reinspect existing
PDFs before a content audit when the measurement lacks a page-inspection pin.
Numerical visual limits live in
`tools/CompositionDocxThresholds.json` and `tools/CompositionPptxThresholds.json`.
Comparisons check page counts and raster dimensions at the selected resolution. Reports record the limit
file identities and apply current limits to reused pixel metrics. Null recall on
a matched page requires a measured zero reference foreground count. Review the
conversion diagnostics in the measured outcomes alongside visual and content
findings. [Capabilities](Capabilities.md) lists
current limitations; development evidence and experiment history stay in ignored
artifacts or Git, rather than public reference pages.

## Private documents

Place private manifests and inputs under ignored `private-cases/`. Run
`pwsh tools/CheckPrivateCase.ps1 -Case private-cases/<case>.json`; outputs stay in
ignored `artifacts/private-visual/`. The script rejects tracked inputs, paths
outside private-cases and unsafe IDs. Public notes record anonymized feature gaps,
never private text, screenshots, filenames or document contents.
Both private case commands accept `-CliPath <qualified.dll>` to reuse a built
CLI. Reference rendering and optional inspections have their own tool builds.

Office reference operations require an available authorized environment. Respect
rejected actions. Missing references block their conclusions, not unrelated work.
Preserve source inputs, useful evidence, current qualified artifacts and published
packages when retiring stale generated output. Check the loose 5 GB workspace cap
infrequently with `pwsh tools/MeasureWorkspaceFootprint.ps1`.
