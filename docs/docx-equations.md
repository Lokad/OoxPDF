# Native DOCX equations

Supported OMML is rendered through managed math boxes, shared font preparation
and ordered inline measurement. This is approximate equation rendering, with
explicit fallbacks for unsupported structures; it is not general Word equation
layout. See [release scope](PreviewRelease.md) and [capabilities](Capabilities.md).

## Supported combinations

Within the qualified subset, bookmarks and unchanged-format run splitting preserve
mathematical content and source ownership. Fitting multiple equations, prose and
inline pictures share measurement and placement in body paragraphs, simple table
cells and direct headers/footers. Numbering owns the first authored row; formulas
beside prose use inline sizing while a later math-only row owns display sizing.
Links and annotations follow the resulting physical geometry.
In qualified ordinary textboxes, omitting unsupported mathematics retains its
diagnostic while supported captions and pictures keep their shared wrapping.
Prepared math keeps its measured frame; entirely unsupported paragraphs and
ordinary fixed-line pictures keep their existing fallback. Exact Office geometry
for new combinations remains unqualified.
Hyperlinks at an equation's root merge their fragments once per placed line and
use its occupied frame and resolved after-gap. Links inside mathematical arguments
retain local glyph/rule bounds. Targets survive wrapping and qualified revisions;
surrounding paragraph links keep their own geometry. Other Office rectangle and
clipping differences remain approximate.
Bookmark boundaries inside supported mathematics retain source positions on
measured rows and arguments. Their internal-link targets follow the same box
translations, including fraction/script placement and inline wrapping.
Comment ranges may start or end inside supported equations without removing
the formula or its comment body. Their geometry uses the containing inline
equation's occupied frame; highlighting only part of a formula remains approximate.
Rich-text content controls and custom-XML metadata preserve supported mathematical
rows and literals, including qualified revisions and source annotations. Controls
read their stored `sdtContent`; metadata properties do not supply printed content. Unknown structures, orphan
argument roles and nested revisions retain their existing fallbacks.
Empty proofing and permission boundaries inside mathematical rows and arguments
preserve supported formula content. Permission boundaries can retain sub-point
Office geometry differences. Bookmark/comment navigation and revision ownership
remain separate from these editing markers.
Adjacent display containers combine in source order; the later container supplies
its justification, including the default when omitted. Bare formulas add content
without replacing the container's properties. The first effective display row
owns subsequent display-row placement in the text paragraph. Authored properties
remain available; separate paragraphs stay independent. Unknown justification
values retain their diagnosed fallback.
`centerGroup` centers the maximum occupied row width and preserves each equation
box inside that group. Explicit `center` centers rows individually; prose ends a
group. An omitted local choice inherits `mathPr/defJc`, defaulting to `centerGroup`.
Raw local and document values remain distinct. Unknown document defaults keep
the prior row-centered approximation; unknown local values keep their diagnosed
fallback. Geometry and ink thresholds still apply.
A math-only row can promote an inline source formula to that display owner;
an inline formula beside prose does not.
Ordinary source characters, including spaces and hidden text, keep math inline.
Links and pictures also determine the shared row mode. Hidden text stays unpainted;
its presence still affects equation layout. Hidden manual-break combinations
remain approximate.

Authored body page breaks can separate supported formulas without discarding
prose, pictures or notes. Spanning links follow content on each page; a spanning
bookmark resolves to its first placed source content. Portable regression checks
cover this in Final/preserved and AllMarkup/Word-compatible modes with pinned
fonts. These checks establish content and navigation preservation, not agreement
with Office page geometry or unrestricted multi-column flow.

Literal text, fractions, scripts, roots, delimiters, n-ary operators, functions,
matrices and cases have qualified bounded forms. Their presence does not imply
unrestricted nesting or line breaking. A sole inline equation uses standalone
display behavior in qualified body/cell paragraphs without altering its source
flag. Wider or unfit mixed objects retain diagnosed fallback. Painted math rows
are atomic; arbitrary row fragmentation remains unqualified.

Automatic paragraph spacing includes resolved style/default values. Prepared
display equations use their nominal metric owner for ordinary expansion while
child glyphs retain their own sizes and ink bounds. Mixed inline rows add ordinary
expansion after their occupied math/picture frame, once per completed line.
Supported picture-only paragraphs share that frame across body, cells and static
stories. Minimum, exact and compressed spacing retain their qualified contracts;
these rules do not establish full-page Office fidelity.

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

A unique parsed paragraph-property change retains supported equations in its
resolved body, cell or direct header/footer paragraph. Duplicate or unowned
changes and unresolved joined property owners retain their fallback. Formatting
balloons and complete-page fidelity remain approximate.

Empty bookmarks, proofing and permission boundaries inside these revision rows
preserve supported formulas. Included bookmarks keep their measured source targets;
excluded revisions do not publish their bookmarks or revive deleted text.

Content controls and transparent inline metadata preserve the same qualified
deleted-equation frames in Final. Their deleted text, links and other excluded
content remain filtered; the metadata does not create another revision owner.

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

[Release qualification](PreviewQualification.md) records the current frozen runtime,
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
