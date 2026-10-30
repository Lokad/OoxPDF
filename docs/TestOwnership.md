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
fonts, CLI and composed-document controls. The forty-method checkpoint also
covers inferred chart titles, table padding, note-list order and related-story
drawing diagnostics. Extensive suites use .NET 11.
Do not run all portable rendering tests again merely to check framework transfer.
Reuse one frozen artifact per TFM and retain each report's actual runtime identity.

Independent Office comparisons remain necessary for fidelity changes. Parser or
layout assertions alone cannot establish visual correctness; image comparisons
alone cannot establish searchable text, annotations or links. Historical failures
and unavailable checks retain their explicit classifications.

`DocxNormalTableStyleTests` checks sixteen source-pinned Word controls for the
intrinsic style's name/custom metadata, padding and paint, including direct and
descendant overrides. `docx-table-style-margins` includes the existing built-in
padding controls; its union with `docx-conditional-table-styles` covers inheritance
and banding once per method. Full-page gates retain their original thresholds.

`docx-note-list-order` checks lists in footnote/endnote tables against pinned Word
labels, including unused stories and reordered XML. Portable painting retains
both fractions, the picture and the visible list sequence in Final and Review;
source order and cached XML remain unchanged. Numbering/restart/ID and existing
note-table checks share the same focused capability selection.

`docx-table-style-margins` owns resolved padding and grid origins across body,
header, inheritance, explicit indentation and direct overrides. It retains the
one-point independent Word origin gate, bare/TableNormal/incomplete-chain
fallbacks and the existing style-color oracle. Thirteen `TableStyleInheritance`
controls distinguish empty/text-only forward inheritance from table/cell formatting,
including individual property groups, a single conditional region, inherited
color/borders, an ordinary descendant and TableNormal.
The former forward-only origin exception is removed; all margin controls use
the same independent Word origin assertion.
For inheritance changes, select this capability together with
`docx-conditional-table-styles`; the runner coalesces their exact method sets.
`TableBandSizes` pins Word fills for absent/positive/zero sizes, both axes,
look flags, explicit membership and inheritance; band availability is separate
from the existing edge/region precedence assertions.

`docx-conditional-table-styles` covers the shared table-style cascade, direct
formatting, inheritance, false flags, margins and borders. Four pinned public
controls and nine edge/span controls retain 201 independent Word cell-fill expectations.
These include compatibility mode 15, coincident edges and trailing-only fallback.
Their font-independent checks run in the Ubuntu capability
union; full-page comparisons remain separate from resolved-color correctness.

`docx-page-print-origins` qualifies body origins, annotation targets, inspection,
static stories and following-note boundaries together. Eight source-pinned public
controls retain sixteen independent Word PDF measurements. Portable tests change
only the first page's font and require stable continuation paint and targets;
Ubuntu CI selects their union with the existing rendering capabilities. The
separate Windows-font oracles retain their declared environmental preconditions.
Full-page image failures remain separate from the one-point geometry checks.

Portfolio measurement and the source-text gate share decoded operation assembly:
preserve emitted spaces and operation order within each page; separate pages with
a newline. `TestCompositionMeasurement.ps1` checks punctuation/run splits, missing
spaces/words, page boundaries and decoded Unicode without builds or conversion.
This is a content-presence screen; geometry and visual gates remain independent.

`PptxTableMergeTests` pins six Office table controls and checks actual painted
content, source grid ownership, styles, links and row-height recovery. The three
portable methods join neighboring table checks through `pptx-table-merges-portable`;
the separate Office-origin oracle requires the declared installed Aptos faces.
