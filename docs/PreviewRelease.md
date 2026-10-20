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

The review follow-up freezes source `01827c40`. All 48 affected checks pass
within one Windows broad invocation: 3,719 pass, 11 skip, and one inherited
balloon-wrapping assertion fails identically against the previous qualified DLL
on the same machine. The assertion and rendering policy remain unchanged. The
exact frozen runtime passes 796 Ubuntu checks with 43 environmental skips.
Windows and Ubuntu package consumers use the same packed DLL. Independent
platform build/package checks are recorded in
[candidate CI](https://github.com/Lokad/OoxPDF/actions/runs/39452164339).

The second review also repairs diagnostic-observer propagation and cancellation
from the final resource-summary callback. Caller failures keep their original
exception identity, type and stack; PPTX malformed-content recovery no longer
swallows them. The unused image-backed luminosity-mask subsystem and test-only
decoder wrappers are removed; live PNG alpha masks and vector luminosity groups
remain. Frozen visual family binding and byte-stable fixture checkout support the
one-artifact workflow. These repairs add no public API removal.

All sixteen portfolio PDFs and diagnostic lists remain byte-identical to the
accepted baseline, with matching Office page counts. Two PPTX strict pixel gates
pass; fourteen DOCX strict gates remain failed. No thresholds changed. Shared
resolver request caches now synchronize hits, publication and bounded eviction.
Stream-failure documentation explicitly requires discarding incomplete output;
file publication remains atomic. No additional public API was removed by this
follow-up.

The source, runtime, package and report identities are retained in the local
candidate record. Preparatory builds with different compiler outputs are kept
separate. Root transfer verifies all 17 assets and two PDF integration smokes;
it does not rebuild or repeat the broad suite. Packaging documentation and
symbols reuse the qualified library and PDB. No performance speedup is claimed.
Broader features and optimization are deferred.

## Remaining document differences

- The local balloon-wrap assertion observes a 12.61pt single-row body where its
  Office-derived expectation requires a second row. The old and new DLLs agree.
  Previous CI skipped this check because Aptos was absent. Its underlying cause
  remains unresolved; no new font/document-shape guard or relaxed assertion was
  introduced. Review balloons remain approximate.

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
