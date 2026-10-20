# 0.2.0-preview candidate

The candidate is being prepared for .NET 10. Publication is deferred. This is a
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

Qualification is pending. The candidate record will pin the committed source,
frozen runtime, package identity, platform/test outcomes and residual failures.
No performance speedup is claimed. Broader features and optimization are deferred.

## Preserved runtime contract

The library remains safe managed code with no runtime dependencies or native
interop. Cooperative cancellation, meaningful archive/XML/image/font bounds,
bounded caches and page staging, optional coarse host memory admission,
deterministic cleanup, atomic file publication and PDF structural checks remain.
Host admission is not comprehensive heap accounting. Separate Office/PDFium
validation tools are outside the shipped library.
