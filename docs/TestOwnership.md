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

Use `framework-compatibility` in `tools/qualification-capabilities.json` for the
small .NET 10 lane: public conversion, streams, cancellation, atomic publication,
fonts, CLI and three composed-document controls. Extensive suites use .NET 11.
Do not run all portable rendering tests again merely to check framework transfer.
Reuse one frozen artifact per TFM and retain each report's actual runtime identity.

Independent Office comparisons remain necessary for fidelity changes. Parser or
layout assertions alone cannot establish visual correctness; image comparisons
alone cannot establish searchable text, annotations or links. Historical failures
and unavailable checks retain their explicit classifications.
