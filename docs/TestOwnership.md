# Test ownership

The dependency-free console runner owns library contracts. Run its selected TFM
explicitly; a PowerShell process may use an older runtime than the library.
The table is the starting map for 0.3.0; rendering slices refine their affected
selections and remove duplicate assertions as they change.

| Contract | Primary owner | What it proves |
|---|---|---|
| Package relationships, MCE and input bounds | `OoxmlTests`, `OoxMarkupCompatibilityTests`, `OoxmlBoundsTests` | Source interpretation and recovery, without a visual oracle |
| Effective format properties | Format reader/model tests, including the PPTX scene tests | Resolved values and precedence from source declarations |
| Measurement and placement | DOCX inline/layout tests; PPTX scene/text agreement tests | Advances, wrapping, geometry, clipping and placement ownership |
| PDF resources, encoding and nesting | `PdfWriterTests`, `PdfResourceIndexTests`, `PdfTransparencyGroupTests` | Structural PDF invariants independent of Office operation grouping |
| Public conversion and lifetime | `PublicApiTests`, `OoxMemoryAdmissionTests`, `OoxResourceGuaranteeTests`, `FontProgramCacheTests` | Publication, cancellation, refusal, cleanup and bounded cache ownership |
| CLI diagnostics publication | `CliTests` and `TestCliDiagnosticsPublication.ps1` | CLI argument/exit behavior and frozen output publication through filesystem aliases |
| Content and appearance together | Fixed composition portfolio and capability visual cases | Supported content survives composition; independent Office appearance and source/PDF semantic evidence remain separate |

`PublicApiTests.PublicationLinks.cs` owns eleven direct library link scenarios
through four console methods. These replace the library calls previously loaded
inside `TestCliDiagnosticsPublication.ps1`; the script retains its twelve CLI
scenarios. Windows junctions and hard links run without symlink privilege. Only
the file-symlink methods declare that environmental precondition. Linux runs all
eleven library scenarios.

`PptxChartPaletteTests` compares painted output against 481 independent Office bar
colors across counts, themes and orientations. It replaces eighteen tests of
private fitted helpers. The `pptx-automatic-palette` selection includes negative
inversion, explicit paint, neighboring axes/labels and line-stroke controls.
The 205-point witness retains one point-specific, one-byte rounding residual;
whole-slide thresholds are unchanged and remain separate from color assertions.

`DocxCommentsTests.PlainParagraphs.cs` pins eight independent Office controls for
plain comment paragraphs, including neighboring counts and a wrapped middle.
The mixed-balloon matrix checks measurement/painting and unsupported fallbacks;
the reader test owns paragraph-mark retention. Select their union once through
`docx-plain-balloon-paragraphs`. The exhaustive mixed-body matrix remains a broader
milestone check; it is distinct from the compact paragraph-boundary matrix.

`pptx-fallback-hyperlinks` checks embedded and standard-font paths together:
chart titles/axes/labels, shape/cell runs, internal destinations, boundaries,
rotation, clipping and diagnostics. `MissingTestFontResolver` supplies neither a
usable primary face nor discovery fonts, so host fonts cannot conceal collapsed
fallback advances. Ubuntu CI runs its union with the existing picture/font
capabilities in one invocation.

Chart and axis titles are interpreted by the scene builder. Rendering and reserved
space consume that model; removed XML-only title branches no longer need private
branch-equivalence tests. `pptx-chart-titles` retains authored text/style, inferred
titles, manual/overlay layouts, axis reserves, links and unsupported-position
diagnostics. Synthetic style checks supply an explicit font; bold-face selection
observes requested styles instead of counting host-dependent embedded font files.

PPTX inspection accepts the same explicit font owner for model measurement and
glyph preparation. The live inspector exposes `--font-directory` (PowerShell:
`-FontDirectory`); retain/pin those font inputs for comparisons. Five inherited
Linux checks now supply pinned faces, preserving their assertions and avoiding
environmental skips. The portable paint/inspection control checks the same source
and face on both paths. `pptx-font-identity` includes the five checks for Ubuntu CI.

Use `framework-compatibility` in `tools/qualification-capabilities.json` for the
small .NET 10 lane: public conversion, streams, cancellation, atomic publication,
fonts, CLI and three composed-document controls. Extensive suites use .NET 11.
Do not run all portable rendering tests again merely to check framework transfer.
Reuse one frozen artifact per TFM and retain each report's actual runtime identity.

Independent Office comparisons remain necessary for fidelity changes. Parser or
layout assertions alone cannot establish visual correctness; image comparisons
alone cannot establish searchable text, annotations or links. Historical failures
and unavailable checks retain their explicit classifications.
