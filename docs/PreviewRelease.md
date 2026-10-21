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

Exact source/runtime identities, Windows and Ubuntu results, package consumers
and CI provenance are recorded in the
[current qualification record](https://github.com/Lokad/OoxPDF/blob/master/docs/PreviewQualification.md).
The broad local suite retains an inherited balloon-wrapping assertion failure
that also occurs with the preceding qualified DLL on the same machine. The
assertion and rendering policy remain unchanged. Local failed checks and
platform skips are reported separately from green CI.

Maintenance repairs preserve diagnostic-observer exception identity/type/stack and
observe final-summary cancellation before publication. CLI diagnostics validate
literal path separation and publish atomically, so existing hard/symbolic links
cannot overwrite source or old destination files. Unused luminosity-mask/decoder paths, inactive overloads, inert PPTX emission
and legend state, and unread argument plumbing are removed. Tests target live
implementations; useful inspectors, PNG alpha masks, vector luminosity groups
and independently validated rendering rules remain. No additional public API
is removed by this cleanup.

File conversion captures absolute input/output paths before host callbacks, so
working-directory changes cannot redirect input lookup or final publication.
File-path guards resolve directory-link ancestors and use existing entries to
protect case aliases on either platform. Provably distinct names remain allowed;
unresolved links or case-only paths refuse conservatively. Raster validation preserves
spaced process arguments and its external timeout; transfer checks include the
frozen DOCX inspector's source and parent build settings. These are contract and
tool repairs, with no additional public API removal or rendering rule.

Qualification checks root build settings and binds converter/inspectors to one
frozen runtime. Local capability selections keep explicit broad milestone sets.
Complete raster derivatives bind PDF, DPI, tool and page identities; stale or
incomplete derivatives cannot satisfy a hit. Missing completion records recover
through fresh staging from the retained PDF; old partial pages stay separate.
Office worker arguments preserve spaces and logs through success/failure/timeout.
Affected qualification declares expected platform skips and retains all outcomes;
unexpected skips still reject its gate. CI retains all
frozen runtime assets. These repairs add no public API removal or rendering rule.
Raster wrappers, cached-reference serving and the direct rasterizer refuse destinations containing PNGs;
use a fresh output directory rather than reusing a visual run ID. Existing files
are preserved. Kerning scans use cancellation rather than work-only fuel while
retaining pair/coverage/class expansion bounds.

All sixteen portfolio PDFs and diagnostic lists remain byte-identical to the
accepted baseline, with matching Office page counts. Two PPTX strict pixel gates
pass; fourteen DOCX strict gates remain failed. No thresholds changed. Shared
resolver request caches now synchronize hits, publication and bounded eviction.
Stream-failure documentation explicitly requires discarding incomplete output;
file publication remains atomic. No additional public API was removed by this
follow-up.

The source, runtime, package and report identities are retained in the local
candidate record. Preparatory builds with different compiler outputs are kept
separate. Root transfer verifies every frozen asset and two PDF integration smokes;
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

Recent maintenance protects final input-link referents from output and diagnostics
publication and removes wrapping fuel, unused counters and inactive helpers.
Retired zero-valued markup corrections are removed while authored spacing,
revision colors and nonzero balloon geometry remain. Tests and allocation tooling
use live implementations. There is no additional public API removal.
