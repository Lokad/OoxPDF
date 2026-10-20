# 0.2.0-preview candidate

The candidate is prepared for .NET 10. Publication is deferred. This is a
preview of the supported scope below, with known rendering approximations; it
does not promise arbitrary-document fidelity or pixel-perfect Office parity.
[Caller migration](PublicApiChanges.md) lists the intentional API removals and
requires recompiling consumers. Package/informational version is
`0.2.0-preview`; assembly/file version is `0.2.0.0`.

## Rendering scope

| Status | Scope |
|---|---|
| Supported within qualified constraints | DOCX text and formatting, ordinary numbering, fixed-grid tables, raster pictures and bounded SVG paint, simple equations composed with text/bookmarks/pictures, qualified body/cell/header/footer placement and notes. Source-owned links use the resulting physical geometry. PPTX retains its distinct absolute-position layout, common shapes, tables, pictures, cached-data charts and hyperlinks. |
| Approximate | DOCX pagination and typography for complex documents; review balloons, ranges and revision geometry; some note and nested-frame combinations; chart automatic layout, unstyled colors and point-count color regimes; alpha/color nuances. Correct conversion does not imply a passing Office pixel comparison. |
| Unsupported or staged fallback | Arbitrary floating/nested/fragmented DOCX stories, general Word field evaluation, unrestricted equation/shaping coverage, advanced charts and SmartArt, arbitrary custom picture geometry, and SVG picture-link composition beyond qualified raster links. Unsupported content may receive a diagnostic and be simplified or omitted; inspect diagnostics before relying on a conversion. |

These are scope summaries, not blanket guarantees for every combination. The
[capability inventory](Capabilities.md) and [markup modes](DocxMarkupModes.md)
describe constraints and fallbacks. Public DOCX defaults remain Final with
preserved document layout. Qualification also measures AllMarkup with the
Word-compatible review margin separately.

## Finite candidate gate

The gate covers demonstrated content loss, serious clipping/overlap, wrong
pagination and regressions within supported scope. It uses focused semantic and
resource-lifetime checks, a Windows broad suite, a portable Ubuntu suite,
package-only consumers, and nine representative synthetic/public documents in
sixteen outcomes. The portfolio combines reports, letters, specifications,
decks, tables, numbering, links, equations, pictures, headers and notes.

Independent Office references and existing visual thresholds stay fixed. Reports
separate content/page defects, significant geometry/typography differences and
cosmetic residuals. Manifest validity counts do not measure Office fidelity.
Unknown content or severe geometry failures must be resolved or explicitly
excluded from supported scope before accepting the candidate.

The frozen runtime is qualified from source commit `6f750f6e`: 3,712 Windows
tests passed (16 skips), and the exact same runtime passed 792 Ubuntu tests
(43 skips). Both platforms consume the same packed DLL. The nine-document,
sixteen-outcome portfolio retains all accepted PDF bytes and matching Office
page counts. Its two PPTX outcomes pass the strict pixel gate; all fourteen DOCX
outcomes retain their existing strict visual failures. No thresholds changed.
[Candidate CI](https://github.com/Lokad/OoxPDF/actions/runs/39404138180) also
passed the independent Windows/Ubuntu build and package checks.

The source, runtime, package and report identities are retained in the local
candidate record. Preparatory builds with different compiler outputs are kept
separate. Root transfer verifies all 17 assets and two PDF integration smokes;
it does not rebuild or repeat the broad suite. Packaging documentation and
symbols reuse the qualified library and PDB. No performance speedup is claimed.
Broader features and optimization are deferred.

## Remaining document differences

- Strict DOCX failures include typography and geometry differences; they are
  not all classified as cosmetic. The composed SVG report's Review prose
  baseline differs by 1.436 points and still fails the original 1-point gate.
- The linked-footnote Review case retains a 46.794-point mismatch with Word's
  PDF link annotation rectangle. The candidate annotation follows its physical
  picture within 0.000016 points. The raw Office comparison remains failed.
- Two static-report content screens retain raw field flags. Source cached
  PAGE/NUMPAGES values of `1` differ from both printed outputs; independently
  inspected Word and candidate page counters agree. Neither has a token deficit
  against the other. This does not establish general Word-field evaluation.
- Point-count chart-color regimes and exact-five-paragraph balloon rules remain
  constrained behavior. Independent Office evidence does not justify a general
  replacement. They were preserved without extending their special cases.

No new content omission, pagination error or serious overlap is demonstrated by
the fixed portfolio. PDF identity and source/geometry checks are not exhaustive
ink or clipping proofs for arbitrary inputs. Keep diagnostics and independently
validate documents outside the qualified scope.

## Preserved runtime contract

The library remains safe managed code with no runtime dependencies or native
interop. Cooperative cancellation, meaningful archive/XML/image/font bounds,
bounded caches and page staging, optional coarse host memory admission,
deterministic cleanup, atomic file publication and PDF structural checks remain.
Host admission is not comprehensive heap accounting. Separate Office/PDFium
validation tools are outside the shipped library.
