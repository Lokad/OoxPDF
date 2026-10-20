# Native DOCX equations

Supported OMML is rendered through managed math boxes, shared font preparation
and ordered inline measurement. This is approximate equation rendering, with
explicit fallbacks for unsupported structures; it is not general Word equation
layout. See [preview scope](PreviewRelease.md) and [capabilities](Capabilities.md).

## Supported combinations

Within the qualified subset, bookmarks and unchanged-format run splitting preserve
mathematical content and source ownership. Fitting multiple equations, prose and
inline pictures share measurement and placement in body paragraphs, simple table
cells and direct headers/footers. Numbering owns the first authored row; formulas
beside prose use inline sizing while a later math-only row owns display sizing.
Links and annotations follow the resulting physical geometry.

Literal text, fractions, scripts, roots, delimiters, n-ary operators, functions,
matrices and cases have qualified bounded forms. Their presence does not imply
unrestricted nesting or line breaking. A sole inline equation uses standalone
display behavior in qualified body/cell paragraphs without altering its source
flag. Wider or unfit mixed objects retain diagnosed fallback. Painted math rows
are atomic; arbitrary row fragmentation remains unqualified.

Body-table placement reuses prepared boxes, cell measurements and clips. Resolved
automatic, minimum and exact row heights have independent Office controls. A zero
exact height uses natural height; small exact rows keep their clips. Supported
vertical-merge restart owners retain content. Unknown/continuation merge owners,
vertical text, fit-text and richer nested cells keep their existing fallback.

Direct static stories use shared measurement, source frames and physical painting
for the qualified plain/mixed-content subset. Header/body reservation and footer
bottom anchoring use the same owning frame. Complex floating, nested, fragmented,
field-heavy and incomplete static stories remain constrained. Supported note
composition uses source-owned admission; continued, complex and page-dependent
note combinations retain explicit fallbacks.

## Revisions and empty arguments

Supported inserted/deleted mathematical runs and bounded structural wrappers retain
their actual source owners. Final filters deleted literal text while preserving
qualified structural frames. Word-compatible Review applies actual author ink and
deletion strikes to revised literals; structural symbols retain qualified authored
ink. Unowned metadata, move ranges, unknown/deeper/multiple/nested revision wrappers
retain fallbacks. This does not establish arbitrary revised-equation support.

Qualified empty fraction/script/root/matrix/function/operator arguments use
metric-only placeholder geometry, without dummy glyphs or font subsets. Remaining
glyphs and structural paint stay visible, including applicable Final deletions.
Known placeholder-visibility profiles share the independently observed printed
behavior. Unsupported empty/deeper structures keep their diagnosed fallback.

## Evidence and remaining differences

Independent Office comparisons cover content, source ink, group origins and
body/cell/static-story placement. Public API defaults are Final with preserved
document layout; AllMarkup with Word-compatible geometry is measured separately.
A successful conversion, matching page count or manifest validation does not mean
a strict visual pass. Review diagnostics for ignored/approximate content.

Remaining differences include operator-cell X offsets of 1.163pt, larger structural
child placement up to 1.96pt, table-border offsets, revision bands and root/body
clearance. The fixed portfolio's fourteen DOCX strict gates remain failed. No
threshold is relaxed; these geometry/typography differences are not all cosmetic.
General operator breaking, broad OMML nesting and arbitrary floating/nested stories
remain deferred. Do not extend font/size/document-shape exceptions without
independent Office evidence for a semantic rule.

[Preview qualification](PreviewQualification.md) records the current frozen runtime,
portfolio and residuals. [Composition history](CompositionCheckpoint.md) keeps
milestone evidence. The superseded equation-development journal remains in
[Git history](https://github.com/Lokad/OoxPDF/blob/78e7589ef3f0ddc9b1fbc38c94fa6743fd7fde16/docs/docx-equations.md).
It includes obsolete exclusions such as bookmarked equations disappearing; those
are historical defects already fixed by shared inline composition.

## Qualification workflow

Use `RunDocxAffectedTests.ps1 -Capability <changed-capability>` against one committed
frozen Release runner, including neighboring behavior. Keep TMP/TEMP scoped to its
ignored run output. Reuse independent Office references and accepted comparisons
only after checking input/options/PDF identity. Transfer verified assets with a
small integration smoke; do not rebuild the root library or replay the same suite.
Broader Windows/Linux, package and portfolio checks close coherent milestones.
