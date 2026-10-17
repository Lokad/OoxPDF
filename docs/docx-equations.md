# Native DOCX equations

Supported equations retain actual inserted/deleted source ownership across body,
simple cells and direct headers/footers. Word-compatible painting shares source
author colors and deletion strike metrics; unowned metadata, move ranges and
unsupported structural revisions retain explicit fallbacks. Adjacent revised text,
bookmarks and unchanged-format run splits do not exclude supported mathematics.
Independent references pin their Office view settings with each source/PDF.

Whole-equation insertion/deletion wrappers inside `oMathPara` retain supported
root, delimiter and n-ary content, alongside prose, numbering, images and other
formulas. Final preserves known literal deletion frames and structural paint;
Review keeps structural symbols in authored ink while literal text receives the
actual author's color and deletion strike. Structural-only insertions retain
actual control ownership without invented text. Independent group gates pass
within 0.683pt X / 0.584pt Y; strict raster gates remain failed. Empty structural
insertions still have an extra Review margin bar. Unknown/deeper/nested wrappers
retain explicit fallback.

In Final, deleting a supported bar fraction's text preserves its empty arguments
and structural rule. Authored empty bar fractions share the same metric-only
placeholder measurement as hidden matrix cells, without a dummy glyph or PDF font
subset. Controls at 9/11/18pt retain the same independent Word rasters for deleted
text and authored empty structure. A partial empty bar-fraction argument also retains its placeholder geometry.
Empty stacks and unsupported richer revision structures retain fallback. Approximation diagnostics
and failed visual thresholds remain separate from content/geometry checks.

Nonempty literal run insertions/deletions inside a supported formula preserve
actual run ownership and affect only their glyph ink in Word-compatible Review.
Final filters deleted text while retaining supported fraction structure and other
arguments. Independent numerator controls cover body/cell/header/footer placement;
bookmarks and splitting the mathematical run preserve geometry. Empty or
structurally different inner revision wrappers retain explicit fallback.

Empty subscript/superscript arguments and partial empty matrix cells now retain
metric-only placeholder geometry and the remaining source glyphs. Independent
Word profiles with omitted, disabled or enabled matrix placeholder visibility
produce identical printed rasters. The same known profiles now share one PDF
layout rule. Empty literal script bases and visible n-ary limits also retain metric-only
geometry and remaining source content, including Final deletion. Five independent
Word roles qualify deleted/authored-empty identity and sub-point group starts.
Unpainted empty function/grid forms, unknown properties and richer structures retain
fallback. Revision-role body controls cover literals, scripts, roots and matrices. Additional
mixed controls qualify empty script/matrix roles with prose, images, numbering and
links across body/cells/direct headers/footers in both presets. Geometry remains
approximate; strict visual gates stay failed. Empty bases/operator limits also preserve
content and source ink in these mixed containers. Two operator-cell outcomes retain
a 1.163pt horizontal group-start residual; the one-point geometry gate stays failed.

Empty literal function arguments and n-ary bodies now share the same metric-only
argument measurement. Function names, operators and remaining limits stay visible,
including Final deletion, with independently validated source ink and group starts.
An empty literal function name retains its visible argument. Supported root,
delimiter and n-ary controls retain structural paint even when no literal source
runs remain, using shared font preparation and authored control/paragraph metrics.
No dummy source runs or glyphs are introduced. Independent Word controls qualify
content, source ink and group starts; unpainted empty function/grid forms and
unsupported deeper structures retain explicit fallback. Supported structural-only
controls also preserve content and source ink in mixed numbered/image/link paragraphs
across body/cells/direct headers/footers. One operator-cell Review outcome retains
the known 1.163pt X residual; strict gates remain failed.

Signed degree overlap now clamps the complete root origin. Narrow degrees can begin
to its right without negative group origins; wider degrees retain leading space.
Independent original/revision controls retain shared root/body/bar translation and
horizontal group starts within 0.264pt. Vertical-clearance residuals remain separate.

Empty visible root degrees retain shared script-level placeholder metrics and the
remaining root/body glyphs, including Final deletion. Independent inline/display
controls retain source ink and transformations. Larger structural frames still have
inherited child-placement residuals up to 1.96pt; those geometry gates remain failed.
Stored nonempty hidden degrees and unknown/deeper children retain explicit fallback.

Numbered paragraphs retain supported explicit math zones alongside prose and
inline pictures. The list label owns only the first authored row: its formulas
use inline sizing, while a later math-only row after a hard break owns display
sizing and alignment. Source math flags, bookmarks and unchanged-format run
splitting retain their identities across body/cell/header/footer placement.
Independent controls retain sub-point geometry residuals and failed visual gates;
unreferenced header-part diagnostics remain separate from printed content.

Standalone display equations also paint in body-table rows with resolved
automatic, exact or minimum heights. Cell
measurement, line placement, clipping and keep estimation reuse the prepared
math boxes and conversion-owned fonts. Five independent Word controls cover
literals, fractions, scripts, adjacent cells and surrounding prose: all 18 math
outlines match, with at most 0.10 pt horizontal and 0.34 pt vertical differences.
Declared frames add 13 independent Word controls: all 40 math outlines match,
within 0.05 pt horizontally and 0.18 pt vertically. Small exact rows use the
existing cell clip. Zero exact height uses natural content height, fixing plain
cells as well. The visual thresholds remain unchanged and unmet. Unknown,
negative or unresolved heights, vertical merges, nested/nonparagraph cell
content, fit/vertical text and revisions retain fallback. Fitting inline cell
equations with plain prose share body wrapping. Multiple fitting inline equations
and inline pictures now use shared ordered-fragment measurement and placement in
body, cell and direct header/footer paragraphs. Explicit formulas beside prose
use the same inline source-row sizing while math-only rows retain display style.
Source flags stay unchanged, including across bookmarks and run splitting.
Wider display combinations and
unfit objects retain fallback. Painted math-bearing rows are atomic, with wider fragmentation still
unqualified.
Ordinary default/left tabs now share measured text/math/image advances across
these containers. Fitting right/center tab fields also measure their complete
text/math advance. Decimal tabs, leaders and wrapping/out-of-frame aligned fields
retain explicit fallback. Closed comment ranges can coexist with supported inline math
in body paragraphs and cells: source boundaries survive bookmarks and unchanged-
format run splitting. Highlights and connectors consume placed source extents,
including formula bounds; selection height belongs to the line. Final output
retains the equations while hiding comments. Continuous automatic footnote references with owned display runs also compose
with equations in body paragraphs and cells; superscript offsets share the
container scale and source marker roles delimit note links. Custom/unowned marks,
page/section-dependent numbering, equation/endnote combinations, fields, revisions
and wider related/floating-story combinations remain staged. Independent Word
controls preserve the original visual thresholds; significant typography and
annotation-height residuals remain.

Simple and complex NUMPAGES fields can accompany supported equations in body
paragraphs and simple cells when one bounded retry preserves the finalized page
count. The retry recomputes note-dependent layout and preserves canonical source
fields. If restoring the equations changes the total, the original diagnostic
fallback remains. Independent two-page controls cover both public modes, bookmarks
and run splitting; visual field-font/cell-origin residuals remain. See the
[composition checkpoint](CompositionCheckpoint.md) for evidence and limitations.

The earlier body qualifications below retain their scope and residuals.

Native OMML rendering is in progress. A first native slice paints standalone
literal display rows in the DOCX body, with left, center or right alignment,
and fitting inline literal rows alongside prose. It also paints one display
fraction with literal numerator and denominator, using the font's OpenType MATH
axis, shifts, minimum gaps, rule thickness and leading. Default display scripts
with literal base/subscript/superscript rows now use separate measured boxes,
font script-style glyph variants and owned subset mappings. Those script boxes
also participate in fitting inline prose and following-word wrapping. One
inline fraction with literal numerator and denominator uses the same ownership
and placement contract.
Authored no-bar fractions (`m:type m:val="noBar"`) now use the same source roles and
font resources as bar fractions. Parent MATH stack shifts and minimum ink gaps
place the two children, with no painted rule. Six independent Word controls
improve across display/inline placement, 12/20 pt mathematics and normal Calibri
children; two paired bar controls remain byte-identical. The 12 pt display
controls differ by at most 0.11 pt horizontally and 0.07 pt vertically. Larger
and inline controls retain residuals up to 1.31 pt, and the existing visual
thresholds remain unchanged. Other fraction types retain the whole-equation fallback.

A script base may now contain one bar fraction or no-bar stack with literal
children. The existing measured boxes compose their glyphs, rules and extents;
fraction children inherit display/inline style, while the outer script argument
retains its own script style. Eight independent Word controls improve across
subscript/superscript and display/inline placement. Display position differences
are below 0.15 pt horizontally and 0.14 pt vertically; inline differences reach
0.65 pt vertically. Sixteen direct neighbor PDFs remain byte-identical. This
first nesting grammar retains deeper structures on explicit fallback and keeps
the existing visual thresholds unchanged.
One bar fraction or no-bar stack can also sit alongside literal mathematical
prefix/suffix atoms, such as `y = fraction + z`. The literal formatter's shared
atom rules receive neighboring box context, so binary operators keep their
spacing at fragment boundaries. Measured glyphs and rules move together, and
logical advance remains separate from occupied ink. Eight independent Word
controls improve across display/inline placement and both fraction types, with
position differences below 0.14 pt horizontally and 0.07 pt vertically.
Twenty-one direct neighbor PDFs remain byte-identical; visual thresholds remain
unchanged. Deeper compound rows retain fallback.
The same sequence path now accepts one subscript/superscript structure alongside
literal atoms. Eight uniform-size Word controls improve with position differences
below 0.12 pt horizontally and 0.05 pt vertically. Two additional 20 pt script/
12 pt literal contrasts establish that the surrounding equation owns the gap
before the structure; the child's authored size retains its glyph geometry.
These contrasts also improve over the missing-row fallback, while internal
script placement and line-height residuals reach 1.81 pt horizontally and 1.07 pt
vertically. No size-specific correction is applied. Twenty-four direct neighbor
PDFs remain byte-identical, including the accepted fraction rows and font styles.
Supported literal, fraction, stack and script atoms can now form a sequence in
one equation. Composition iterates in source order under the existing work and
cancellation limits, with independent child rules and a single parent-leading
contribution. Eight Word display/inline controls cover fraction/script order,
two fractions and stack/script mixtures; all improve, with position differences
below 0.11 pt horizontally and 0.07 pt vertically. Thirty-four prior capability
PDFs remain byte-identical. The same frozen DLLs pass 32 selected portable math
checks on Linux without rebuilding. Existing visual gates and unsupported deeper/
rich structures remain explicit; this does not qualify multiple separate equation
elements in a paragraph or additional stories.

Literal square roots now use the same measured boxes and body/sequence placement.
The reader retains radicand, degree and degree-visibility roles; an empty hidden
degree admits the first root slice. The parent owns the root size, operator glue
and leading, while radicand glyphs keep their authored sizes. Ready-made vertical
variants and a stored overbar use the font's MATH gap/rule metrics. Assemblies,
richer nesting and unsupported metric/font owners retain fallback.
Eight Word controls across 8/12/20 pt radicands improve; all candidate root outlines
match Word/source variants, and 25 prior capability PDFs remain byte-identical.
Larger-child errors drop from about 4.1 pt to below 0.7 pt after parent ownership
is preserved. Small-root placement still differs by up to 1.34 pt; full visual
gates remain under review with unchanged thresholds. The variant reader's 851
queries match independent fontTools values.

Implicit root mappings publish after authored glyph mappings are known. A focused
synthetic test reproduces a prototype Unicode collision and verifies that either
source order preserves the authored character, retaining root fallback when an
encoding owner would conflict. This is a confirmed subset-mapping defect in the
rejected prototype; shared font-cache memory corruption remains unverified.
Root control properties now retain their inherited source formatting. The
configured math parent supplies the implicit glyph and rule independently of a
normal-text radicand's physical font. Eight Word controls cover math/normal text,
Calibri/Arial and authored child colors: six improve and two remain exact.
Root glyphs and bars keep inherited black while children keep their own color
and face. Thirty-three prior PDFs remain byte-identical; position differences
reach 0.12 pt horizontally and 0.46 pt vertically, with unchanged visual gates.
Neutral and supported authored radical control properties now retain typed
size/color/font ownership through the existing run-property reader. Each root
resolves its own implicit font resource, including sibling roots with different
controls. Shared formatting rules keep unknown, revision-bearing, decorative and
bold control formatting on explicit fallback. Ten independent Word controls
across display/inline placement, 8/12/20 pt controls, green color and italic flags
improve against the previous ignored output. Mathematical root glyphs remain
upright under italic control flags, matching Word. All ten candidate outlines
match Word's selected source variants; 22 prior PDFs remain byte-identical.
Position differences reach 0.27 pt horizontally and 1.93 pt vertically for
enlarged controls. Existing visual thresholds remain unchanged.

Visible literal root degrees now use the shared argument-size rule with a
script-script starting level. Parent MATH kerns retain degree/root overlap and
the raise percentage positions degree ink against the selected root. Radicand,
root and overbar move together; degree ink contributes to occupied extents.
Eight independent Word controls improve. Six standard controls differ by at
most 0.16 pt horizontally and 0.14 pt vertically; enlarged controls retain up
to 1.54 pt horizontal and 1.93 pt vertical differences and different optical
degree glyphs. All eight root outlines and ten standard degree outlines match
Word/source fingerprints. Twenty-nine focused methods and 22 byte-identical
neighbors qualify one artifact, transferred with four matching CLI PDFs.
Stored hidden degrees retain source and fallback: Word moves their text into
the radicand in the measured controls, requiring separate qualification.
Four public cases retain the existing visual thresholds.

A single fraction or no-bar stack can also serve as a square/nth-root radicand.
Shared bounded argument discovery supplies the same display/inline child styles
as a fraction/stack script base. The root reuses the measured child box, moving
its glyphs and rules together, while the root owns its overbar and parent
leading. Normal-text children keep independent font resources and rule colors.
Eight Word controls improve with differences below 0.17 pt horizontally and
0.53 pt vertically. All eight root outlines and four visible degree outlines
match Word/source. Forty-five focused methods, 24 unchanged neighbor PDFs and
four exact transfer smoke PDFs qualify one artifact. Four public cases retain
the existing thresholds. Assembly-only heights, richer or arbitrary recursive
radicands and unsupported degree/control formatting retain fallback.

Supported subscript/superscript boxes can serve as radicands. Their arguments
retain existing glyph selection and kerning; the root supplies cramped context
for the font's superscript shift. A script box separately retains its nominal
subscript descent for enclosing operators, preserving the ink frame used by
ordinary paragraph flow. This inferred Office sizing profile is supported by
twelve independent controls, including descenders, paired scripts and 20 pt
sources. The initial ink-only prototype had two visual regressions and is
retained as rejected evidence. Cramped spacing and enclosure descent make all
twelve improve without font/point/document guards. Horizontal differences stay
below 0.12 pt; vertical differences reach 2.74 pt in the enlarged inline case.
All 46 compared glyph outlines match Word/source. Forty-two focused methods,
27 unchanged neighbor PDFs and four exact CLI transfer smoke PDFs qualify one
artifact. Four public cases retain their original visual thresholds; arbitrary
recursive/rich radicands and assembly-only heights remain explicit fallbacks.

The reader retains equation position,
display alignment, resolved run formatting, revision visibility and mathematical
structure for literals, scripts and fractions. Unsupported or incomplete
constructs retain an explicit source representation. Equations are separate from
prose runs, preserving existing field, bookmark, note and picture indices.

Painted rows report `DOCX_APPROXIMATED_EQUATION`. Inline formulas share the prose
baseline and move following prose by their measured width. Following prose can
wrap at its word boundary while the complete formula stays at its fitting source
line position. A formula that needs operator breaking or relocation retains
the whole-row fallback. Deeper nested scripts or fractions,
table/story placements, reference/revision decorations and display rows with
compressed automatic or exact/minimum line spacing retain
`DOCX_UNSUPPORTED_EQUATION` with part provenance
and `Ignored` fallback. Missing fonts or glyphs also keep the whole-row fallback.
These diagnostics describe actual painted content, not just parsed structure.

Literal-row component measurement and painting now use the same bounded glyph
box model, with mathematical alphabet selection, operator spacing and explicit
fallbacks. The isolated inline reference passes the existing DOCX visual gates.
The display reference retains a positioning residual and remains under review;
substituting Word positions is a diagnostic only. These checks do not qualify
DOCX paragraph placement, wrapping or pagination by themselves. Ten additional
body controls qualify the initial display-row placement across font sizes,
alignment and a page boundary. Equation page assignments match Word; most
math-row placement residuals are below 0.13 pt horizontally and 0.08 pt
vertically. Default preceding prose retains an inherited spacing residual, and
full visual gates still need review. Nine additional inline controls cover
surrounding prose, alignment, text and math sizes, and an overfull fallback.
Eight fitting controls improve against Word; the overfull control stays exact.
Baseline residuals are at most 0.14 pt, while existing prose metrics leave
horizontal residuals up to 1.13 pt. Nine further controls qualify following-prose
wrapping across text/math sizes, alignment and long tails. All improve against
Word, with equation baseline residuals at most 0.14 pt. Horizontal residuals reach
2.63 pt, and following-line spacing still needs work. Full visual gates still
need review. A whole-formula relocation prototype was rejected because Word
breaks those equations at mathematical operators.
Nine display-fraction controls cover sizes, alignment, descenders, explicit bar
type and a page boundary. All improve against Word, with matching equation page
assignments. Eight controls have vertical residuals below 0.18 pt; preceding
default prose retains its separate 2.37 pt spacing residual. The numerator,
denominator and rule share stored measured geometry and font resources. These
results remain approximate at the existing visual gates. Missing or invalid MATH
constants, mixed fraction metric owners, nested structures and unsupported
control formatting retain the complete fallback.

Ten display-script controls cover 8/12/16/20/24 pt sizes, individual and paired
roles, descenders and a body boundary. All improve against Word, with matching
equation page assignments and emitted font sizes. Font MATH scaling, downward
half-point nominal rounding and the PDF export grid are resolved separately;
the `ssty` glyphs match independent font inspection. Glyph-to-Unicode ownership
keeps alternate subsets distinct while retaining canonical subset identities.
Before corner kerning, horizontal residuals reached 0.53 pt; descender/subscript vertical residuals
previously reached 1.14 pt. The shared distinction between ordinary glyphs and
composed or font-declared extended bases reduces that descender control to
0.40 pt; seven descender controls improve and ten protected script/composed
controls stay exact. Default preceding prose retains a separate spacing residual.
Shared corner kerning now reads signed font corrections at two heights. Word
contrasts favor the greater corner sum, while the OpenType overview describes
the minimum. This deliberate Office behavior reduces the ordinary script
controls' maximum horizontal residual to 0.20 pt; composed boxes retain their
previous behavior. Six new lower/upper contrasts and 624 independent font
queries support the rule. Small raster trades remain recorded for 8 pt, default
prose and q controls; all visual thresholds are unchanged.
Twelve inline-script controls cover surrounding prose, sizes, descenders,
following-word wrapping, two body boundaries and an overwide fallback. Eleven
gain painted math and improve against Word; one fallback stays exact. Placement
residuals reach 1.09 pt horizontally and 1.52 pt vertically. At the 120 pt body
boundary, the previous and current renderer both move following prose to page 2
while Word keeps it on page 1; the inherited prose-spacing issue remains open.
The 140 pt boundary keeps following prose within the admitted body. All original
visual and page-count gates remain unchanged.
Thirteen further controls qualify inline fractions and multi-atom script
arguments. Twelve improve against Word and one overwide fallback remains exact.
Numerator/denominator glyphs use script style; parent source metrics own the
axis, rule and shifts even when child sizes differ. Script-style binary and
relation operators omit ordinary-row glue. Larger-child placement improves from
6.24 pt to 0.30 pt after preserving that ownership; smaller math retains a
1.07 pt vertical residual. Full visual gates still need review.
Empty argument properties and explicit zero argument size now share the default
row eligibility rule. A cached Word script control improves from SSIM 0.7820 to
0.9238 and foreground recall 0.4879 to 0.8237; emitted script sizes match Word,
with position differences below 0.05 pt horizontally and 0.40 pt vertically.
Forty-four focused methods cover typed source properties, native painting and
existing math invariants. Nineteen selected capability neighbors remain byte
identical; four root CLI smoke outputs reuse the same qualified artifact.
Unknown, duplicate or malformed controls retain the whole-equation fallback.
Layout does not reparse the retained raw XML. An absent argument-size value
also uses the documented neutral default.

Signed argument-size increments now adjust supported subscript/superscript
roles within the font's text, script and script-script styles. Parent metric
ownership remains separate. Stored glyph identities drive corner kerning,
keeping it consistent with the selected and embedded argument glyphs. Fourteen
independent Word controls cover signed increments, paired/individual roles,
12/20 pt sizes, operator spacing and neutral defaults: thirteen improve and one
stays exact. Emitted sizes match Word; horizontal residuals are below 0.11 pt,
while enlarged paired scripts retain vertical residuals up to 1.30 pt.
Forty-six focused methods qualify one Release artifact, with 24 exact selected
neighbors and four matching root CLI smoke outputs. No root rebuild or corpus
replay is required. Three new public cases keep their original review gates.
Nonzero controls on other row roles, unsupported formatting, nested structures
and wider owners remain staged. These results remain approximate.

Standalone display rows now admit explicit single or expanded automatic line
spacing. Expansion adds an ordinary math-font line's extra advance after the
occupied equation frame, preserving its baseline and tall ink. The same stored
height drives placement and keep measurement. Eighteen Word controls cover
literals, fractions, scripts, preceding prose, child/parent sizes and four fixed
or compressed fallbacks. Twelve improve raster comparisons; two script controls
gain foreground recall while SSIM falls by about 0.0039; four fallbacks stay
identical. Existing preceding-prose placement differences reach 2.38 pt.
Display fractions now share parent-owned axis, rule and shift metrics with inline
fractions. The enlarged-child denominator difference falls from 9.41 pt to
0.91 pt; horizontal spacing differences up to 1.86 pt remain. Fifty focused
methods qualify one Release artifact; 33 selected neighbors and four root smoke
outputs retain their pinned bytes. Three new public cases keep their review
gates. Compressed spacing needs its separate baseline contract.

Fraction operator spacing now uses the same parent math size as the axis, rule
and shifts, while glyph sizes, advances and italic corrections retain their
own font ownership. Word's embedded outlines confirm ordinary source glyphs
for both default and enlarged arguments. Six controls cover 8/12/16/18/20 pt
children under a 12 pt parent: five improve and one stays exact, without raster
trades. The enlarged argument's horizontal residual falls from 1.86 pt to
0.53 pt; vertical differences up to 1.16 pt remain. Fifty-two focused methods
and 48 exact selected neighbors qualify one artifact; four root smoke outputs
match without rebuilding. Two new public cases preserve the review thresholds.

Explicit math run styles now select upright, italic, bold or bold-italic Latin
letters through one alphabet mapper. Bold digits retain their upright bold
mathematical glyphs in either bold style; operators keep their ordinary symbols.
Existing font/subset ownership also serves styled script arguments. Ten Word
controls cover display/inline rows, 12/20 pt sizes, scripts and the italic h
exception. Eight improve, two stay exact and 54 embedded glyph outlines match
the installed source font. Position differences remain below 0.19 pt horizontally
and 0.28 pt vertically; 20 pt literal emission retains a 0.04 pt size difference.
Fifty-six focused methods, 54 exact existing math PDFs and four matching root
smoke outputs qualify one artifact. Three public cases keep their review gates.
Other styled alphabets and inheritance from ordinary Word bold/italic run
properties remain on explicit fallback, alongside unsupported decorations.

Basic Greek upper/lowercase letters now share the same bold/bold-italic mapper.
Eight Word controls qualify six improvements and two exact upright/italic
controls; 36 embedded glyph outlines match the source font. Position differences
stay below 0.13 pt horizontally and 0.28 pt vertically. Thirteen specifically
affected alphabet, encoding, fallback, cancellation and budget methods qualify
one artifact; ten portable math-box tests also pass on Ubuntu using those exact
assembly bytes. Sixteen direct neighbors and four root smoke PDFs stay exact.
Three public cases retain their review gates. Variant symbols and other styled
alphabets retain fallback until their semantic mappings are qualified.

Common theta, phi, epsilon, kappa, rho and pi variants, partial differential and
nabla now retain their own mathematical style glyphs. A shared symbol table
serves italic, bold and bold-italic without replacing variant letters with
ordinary Greek forms. Eight Word controls qualify six improvements/two exact
plain rows; all 56 embedded outlines match source glyphs. Maximum position
differences are 0.13 pt horizontally and 0.02 pt vertically. Fifteen affected
methods, 24 exact direct neighbors and four matching root smoke PDFs qualify one
artifact. Two public variant cases keep their review gates; unsupported alphabets
and Word-property inheritance remain staged.

Normal-text math runs (`m:nor`) now preserve a typed absent/true/false flag.
True retains ordinary text glyphs and authored faces while suppressing math
italics, operator classification, glue and italic correction. Explicit false
uses mathematical glyphs and the math face; existing absent-flag face behavior
stays intact. Unknown, duplicate or malformed normal-text controls retain the
complete fallback. Ten Word controls cover default/Calibri/Arial faces,
display/inline placement, false flags and operators inside normal text. All
improve, with position differences below 0.19 pt horizontally and 0.23 pt
vertically. Twenty-one source/style/encoding/fallback methods and 23 exact direct
neighbors qualify one artifact; four root smoke PDFs match without rebuilding.
Three public cases keep their review gates. Rich normal-text formatting and
script/fraction metric owners still need separate qualification.

Normal-text subscript/superscript arguments now take their scale from the parent
math face while retaining child glyphs, ordinary-text variants and authored
faces. The base still owns valid MATH placement constants; unsupported base or
mixed mathematical metric owners retain fallback. Eight Word controls cover
math/Calibri/Arial children, both script roles and signed size increments. All
improve; maximum position differences are 0.18 pt horizontally and 0.11 pt
vertically. Twenty-eight focused methods and 39 exact style/scale/kerning/boundary
neighbors qualify one artifact, with four matching root smoke PDFs. Three
public cases retain their review gates. Normal-text fractions and richer
formatting remain separate capability work.

Normal-text and mixed fraction children now retain independent math/Calibri/Arial
faces while a parent metric owner supplies MATH scale, axis, rule, shifts and
leading. The existing conversion cache supplies that parent; a metric-only face
does not publish an unused PDF font. A typed font/size owner carries leading
through the composed box, avoiding the roughly 1.8 pt baseline differences of
the rejected glyph-only prototype. Eight Word controls qualify seven
improvements and one exact math-face display fraction; maximum differences are
0.11 pt horizontally and 0.06 pt vertically. Twenty-four specifically affected
methods, 33 exact direct neighbors and four matching root smoke PDFs qualify one
artifact. Three public cases retain their review gates. Rich formatting,
unsupported parent/mixed mathematical metrics and nesting remain staged.

Normal-text bold/italic now requests the authored text face instead of only its
family. Available Calibri/Arial faces retain their own weight and slant. The
single math face retains ordinary outlines and uses Word's measured one-third
local italic shear; glyph measurement includes the sheared right ink edge and
painting balances its saved graphics state. The initial missing-style omission
prototype was rejected because it lost visible italic text. Nine Word controls
all improve, with maximum differences below 0.17 pt horizontally and 0.23 pt
vertically. Twenty focused methods, 28 exact direct neighbors and four matching
root smoke PDFs qualify one artifact. Three public cases keep their review
gates. Other unavailable non-math styles and unsupported decorations retain
explicit fallback; the global PDF italic-emission rule is unchanged.

Composed scripts and display centering now distinguish logical advance from
occupied slanted ink. The box retains both widths, preserving fit/clipping
extents while Word-compatible centering uses text advance. Eight composition
controls qualify two sheared-script improvements and six exact real-face,
fraction and tight-inline controls. Horizontal differences fall from about
0.75–0.78 pt to below 0.17 pt on the corrected scripts. Tight inline examples
retain their separate inherited vertical differences up to 2.56 pt. Twenty-two
focused methods, 37 exact style/metric neighbors and four matching root smoke
PDFs qualify one artifact; the previous default-face style controls stay exact.
Two public script cases retain their review gates.

The six `docx-omml-*` test inputs were authored with Lokad.DocxEdit 0.2.0 at
revision `67e303117f3e51465f3da3d4971c357c091d407c`. They contain these expressions
in both inline and display positions:

| Fixture kind | DocxEdit LaTeX input |
| --- | --- |
| literal | `\alpha+\beta=\gamma` |
| scripts | `x_i^2` |
| fraction | `\frac{a+b}{c}` |

The authoring tool and Word reference renderer are validation tools only.
Neither is a dependency of the entirely safe managed OoxPdf library.

Delimiter source now retains argument lists, control formatting and authored
begin/end/separator characters, growth and shape. Omitted parentheses and
separator defaults remain distinct from explicitly empty characters, following
the [OOXML begin-character definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.beginchar?view=openxml-3.0.1).
Shared control-property validation also serves roots. Malformed/unknown values,
missing roles and nested content remain explicit; iterative parsing and
cancellation have focused coverage. Seventeen source/fallback methods qualify
one source artifact. Eight independent Word controls yield fourteen matching
source glyph outlines; 32 focused vertical-variant queries match fontTools.

Centered growing delimiters now paint one literal argument in standalone
display or fitting inline body equations. Parenthesis, bracket and one-sided
brace controls use ready-made variants centered on the parent MATH axis, while
child glyphs retain their own font and color. Roots and delimiters share one
per-node implicit font owner and deferred glyph publication. Alias tests in both
source orders preserve authored Unicode; an unresolved alias or missing closing
CID retains complete fallback. Six Word controls improve with differences below
0.15 pt horizontally and 0.07 pt vertically; all 28 compared glyph outlines
match Word/source. Thirty-four focused methods, 26 unchanged neighbor PDFs and
four exact transfer smoke PDFs qualify one artifact. Four public cases retain
their original visual thresholds. Fraction enclosures retain their measured
source/fallback until height selection is independently qualified, alongside
other shape/growth/separator/recursive/assembly paths.

Literal delimiter fitting now selects the nearest ready-made variant, preferring
the taller shape on an exact tie. The same bounded reader retains the original
adequate-height selection for roots. Both reject malformed records and targets
beyond every ready-made variant, preserving assembly fallback and cumulative work
limits. Explicit parent metrics also own the ordinary baseline floor; independently
sized child ink remains in the occupied box instead of adding the child's separate
typographic line gap. Six 12/16/20 pt and descender controls improve, with vertical
differences below 0.06 pt and horizontal differences up to 0.40 pt. One prior small
root control also improves, while 37 other PDFs remain exact. Forty-two focused
methods, 62 independent font queries, 34 Word/source glyph outlines and four exact
transfer smoke PDFs qualify one artifact. Three public cases retain thresholds.
Measured fraction enclosures need a distinct sizing profile and remain staged;
smaller authored fraction text also shows an unqualified optical-glyph difference.

Authored non-growing delimiters use canonical base glyphs at the control's size,
while children retain their own size and occupied ink. These controls register
only the glyphs they use, avoiding unnecessary stretch work and fallback on an
unused variant alias. Growing controls still prepare their own variants in either
source order; a fixed control also works when a MATH face has no variant table.
A rejected prototype's unnecessary omission is reproduced in a focused alias
test; authored Unicode remains preserved. Six independent Word controls improve
across display/inline placement and 12/20 pt child/control sizes. Differences
reach 0.46 pt horizontally and 0.07 pt vertically; all 30 glyph outlines match
Word/source. Thirty-nine focused methods, 30 exact neighboring PDFs and four
matching transfer smoke PDFs qualify one artifact. Three public cases retain
the existing visual thresholds. Fraction/shape/separator/recursive profiles remain
staged, and font-cache memory corruption remains unverified.

Delimiter argument lists now compose independent literal rows in source order.
Outer characters and supported separators share the largest argument frame and
the authored growth flag; explicitly empty separators emit no synthetic glyph
or advance. Child sizes/colors remain independent, and a missing separator CID
keeps complete fallback. Six two-argument Word controls improve across default/
empty separators and mixed sizes. Position differences stay below 0.65 pt
horizontally and 0.06 pt vertically; all 52 visible glyph outlines match
Word/source. Word's invisible space for an empty separator is recorded separately
in geometry evidence. Forty-one focused methods, 36 unchanged neighbor PDFs and
four matching transfer smoke PDFs qualify one artifact. Three public cases retain
the existing thresholds. Fractions, other shapes/separator glyph profiles and
arbitrary recursive arguments remain staged.

N-ary source now retains integral/sum/product characters, bodies, lower/upper
arguments, growth, hidden limits, limit placement and control formatting. Shared
scalar/control validation preserves unknown or malformed values and distinguishes
omitted placement from the authored default for a present `limLoc`, following
the [OOXML limit-location definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.limitlocation?view=openxml-3.0.1).
Iterative parsing and cancellation have focused coverage. Twenty-one source and
fallback methods qualify one artifact; 23 source/control PDFs remain unchanged.
Eight Word operator outlines match source variants and six targeted selections
match the bounded reader. The display integral uses a taller variant than the
font's display minimum alone selects, so that sizing profile remains unresolved.
The source component alone does not establish native operator placement.


Display sums and products now paint literal bodies and explicit stacked limits
using ready-made MATH variants, existing script rows and the shared control-font
owner. Child sizes, fonts and colors stay independent. Four Word controls improve,
including independent control/body size contrasts; glyph-position differences
reach 0.62 pt horizontally and 0.32 pt vertically. All four operator outlines and
20 of 24 total glyph outlines match Word/source. A 20 pt control with 12 pt limits
uses different optical child glyphs in Word; that residual remains recorded,
without adding a font/size correction branch. Fifteen focused methods, 21 exact
neighbor PDFs and four transfer smoke PDFs qualify one artifact. Two public cases
retain existing visual thresholds and explicit needs-review status. Integral sizing, omitted document defaults, hidden limits, richer arguments and
assemblies remain staged with complete fallback.


Sum/product side limits share argument sizing, constant validation and baseline
shifts with ordinary script boxes. Inline equations use side limits even when
the source requests undOvr, matching the retained Word controls. Display versus
inline trailing body spacing is an inferred Office profile supported by six
independent controls, including larger-control and larger-body contrasts.
All six comparisons improve; position differences reach 0.64 pt horizontally
and 0.24 pt vertically. Larger-control optical child glyph differences remain
explicit, with no added font/size rule. Twenty-four focused methods, 29 exact
neighbor PDFs (including ordinary scripts, fraction bases, cramped roots and
stacked operators) and four exact transfer smoke PDFs qualify one artifact.
Three public cases retain existing visual thresholds. Hidden limits, omitted
document defaults, integral sizing, rich/deeper structures and assemblies
remain staged with complete fallback.


Hidden sum/product limits are omitted only when their supported literal argument
is truly empty. The remaining body/limits retain shared measurement and font
ownership. Six Word comparisons improve across lower, upper and both-hidden
display/inline cases. Position differences reach 0.12 pt horizontally and 0.47 pt
vertically; existing visual thresholds remain unchanged. Eighteen focused methods,
22 unchanged neighbor PDFs, six Word comparisons and four exact transfer smoke
PDFs qualify one artifact. Three public cases retain explicit visual residuals.

Six additional Word controls show that stored nonempty hidden arguments are
reassigned into visible equation roles rather than discarded. Those arguments,
including whitespace or deeper structures, retain source and complete fallback.
This follows the same preservation principle as stored hidden root degrees.
Document-default limit settings, integral sizing, richer structures and assemblies
remain staged.


Document limit defaults now have one immutable source owner shared by equations.
An authored node location takes precedence; otherwise sums/products consult
the document's naryLim setting, whose omitted/present-without-value default is
undOvr. Inline equations retain the independently established side-limit rule.
Integral settings are retained separately; display integral rendering remains staged.
Malformed or duplicate defaults preserve source and complete fallback when used;
an explicit node does not depend on an unused invalid document setting. The
existing scalar-property validator is shared with node parsing. These defaults
follow the [n-ary setting definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.narylimitlocation?view=openxml-3.0.1)
and [integral setting definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.integrallimitlocation?view=openxml-3.0.1).

Eight independent Word controls qualify precedence and omitted/default values:
six gain native output and two explicit placements remain byte-for-byte unchanged.
Position differences stay below 0.16 pt horizontally and 0.24 pt vertically;
all 48 glyph outlines match Word/source. Twenty-six focused methods, 30 exact
neighbor PDFs and four exact transfer smoke PDFs qualify one artifact. Three
public cases retain existing visual thresholds. Default operator characters,
integral sizing, stored hidden content, richer structures and assemblies remain
staged with explicit fallback.


Sum/product bodies now reuse the bounded fraction/stack/script argument grammar
and measurement shared with roots. Child source roles, script levels, inner
fraction rules and colors retain their existing owners; the n-ary control owns
its operator and outer limits. Literal child checks bound composition before
measurement recurs, and missing body encoding retains complete fallback.
Six Word controls improve across display/inline fraction, stack and paired-script
bodies. All six operator variants retain the existing selection profile, and all
44 glyph outlines match Word/source. Position differences reach 0.11 pt
horizontally and 0.49 pt vertically. Twenty-nine focused methods, 24 exact
neighbors (including fraction/script roots) and four matching transfer smoke
PDFs qualify one artifact. Three public cases retain existing visual thresholds.
Arbitrary deeper/rich bodies, integral sizing, default operator characters and
assemblies remain staged; full visual parity is not established.


Inline integrals now reuse the n-ary boxes with font-owned extended-shape and
italic metrics. Extended operators use the shared occupied-base script shifts;
italic correction lets the lower limit precede the upper limit. Integral document
defaults remain separate from sum/product defaults, with explicit node precedence.
Eight independent Word controls improve across 8/12/16/20 pt sizes, explicit
stacked requests, document defaults and fraction/script bodies. All 35 glyph
outlines match Word/source; position differences reach 0.12 pt horizontally and
0.37 pt vertically. Twenty-four focused methods, 29 unchanged neighbor PDFs and
four matching transfer smoke PDFs qualify one artifact. Three public cases keep
the existing visual thresholds. Display integral sizing remains unresolved and
retains complete fallback; no font/point correction is used to enable it.


Centered literal delimiter boxes now participate in bounded mixed rows, sharing
one eligibility rule with standalone delimiters. Surrounding literals, multiple
arguments and fixed/growing controls retain source order, child sizes/colors,
per-node controls and the existing parent spacing/line frame. Missing encoding
in a later control keeps the entire row in fallback. Six Word comparisons
improve across display/inline single, multiple-argument and mixed-growth rows.
Position differences reach 0.67 pt horizontally and 0.06 pt vertically; all
74 glyph outlines match Word/source.
Twenty-three focused methods, 26 unchanged neighboring PDFs and four matching
transfer smoke PDFs qualify one artifact. Three public cases retain existing
visual thresholds. Nested delimiter arguments, unqualified shape/growth profiles
and arbitrary deeper structures remain staged with complete fallback.


N-ary growth now distinguishes an omitted/off property from explicit on. Unlike
delimiters, omitted n-ary growth defaults to off, following the
[OOXML growth definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.growoperators?view=openxml-3.0.1).
Off keeps the normal display minimum or inline base selection. Explicit on fits
the body's occupied height through the bounded ready-made variant reader, in
display and inline equations. It does not use a delimiter's symmetric extension
around the math axis; independent contrasts reject that model. Heights beyond
available variants and missing selected encoding retain complete fallback.

Thirteen qualified Word controls cover source flags and 24/36/48 pt taller-body
contrasts. Six previously omitted controls gain output, five wrong-height cases
are repaired and two adequate operators stay byte-for-byte unchanged. All 90
glyph outlines and 11 independent font queries match. Position differences reach
0.91 pt horizontally and 0.24 pt vertically. Twenty-six focused methods, 33 exact
neighbor PDFs and four matching transfer smoke PDFs qualify one artifact. Three
public cases retain existing visual thresholds. One separately measured display
integral growth-off case retains its unresolved sizing fallback. No font/point
correction, glyph assembly or new quota/cache framework is introduced.


Supported n-ary boxes now participate in bounded mixed rows through the same
contextual eligibility used for standalone equations. Each node retains its
document defaults, source/control ownership and complete fallback. A semantic
large-operator edge owns incoming thin math spacing; the body's final atom owns
outgoing spacing. The existing row composer handles both literal boundaries and
adjacent boxes. This replaces the prototype's ordinary-atom assumption, which
left ordinary-prefix operators too close, without font/point correction branches.

Nine independent Word comparisons improve across binary/ordinary prefixes and
adjacent sum/product/inline-integral controls. Position differences stay below
0.14 pt horizontally and 0.24 pt vertically; all 94 glyph outlines match
Word/source. Five binary-prefix PDFs remain exact
after the edge change and reuse their inspections/rasters. Twenty-five focused
methods, 22 unchanged neighboring PDFs and four matching transfer smoke PDFs
qualify one artifact. Three public cases retain existing visual thresholds;
unqualified display integrals and unsupported deeper/rich structures keep fallback.


Function application source now retains typed function-name and argument roles,
control formatting/family and nested supported/unknown content. Function-name
roles require their owning function; malformed or duplicate roles preserve
traversed source without repeated raw subtree serialization. The parser reuses
shared control/property validation and handles 2,048 nested functions iteratively
with cancellation. This follows the [OOXML function-name definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.functionname?view=openxml-3.0.1).

Six independent Word references cover literal, delimiter, scripted and fraction
arguments, a larger control and an independently declared normal-text name font.
Twenty-one focused source methods and 20 unchanged source/native neighboring PDFs
qualify one artifact, with four exact transfer smoke PDFs. This source component
qualified the source component before native painting. The subsequent native
qualification below retains those source safeguards and visual thresholds.


Native function application now paints a literal name followed by a bounded
literal, delimiter, scripted or fraction argument. The invisible control owns
application spacing and the typographic frame; independently declared name and
argument fonts/colors keep their own paint. Metric-only controls reuse the
conversion font owner without preparing an unused PDF subset or resource. An
opening delimiter suppresses application glue, using the same semantic edge that
also improves ordinary-prefix delimiter rows. Unsupported or deeper names and
arguments retain complete fallback. The mixed-row qualification follows below.

Nine independent Word controls improve in both similarity and foreground recall,
including two ordinary-prefix delimiter contrasts and an all-normal Arial
function. All 50 embedded glyph outlines match Word and their source fonts.
Text-operation differences stay below 0.16 pt horizontally and 0.063 pt vertically.
Twenty-one focused methods and 16 exact neighboring PDFs qualify one strict
Release artifact; four exact smoke PDFs verify its transfer without rebuilding or
repeating the suite. Three public cases preserve existing thresholds and explicit
visual residuals. The all-normal function emits only the painted Arial equation
font, while its metric control remains separate from PDF resource ownership.


Functions now share standalone and mixed-row eligibility through the existing row
composer. Ordinary/binary prefixes, outgoing relations and adjacent controls keep
semantic atom spacing and source order. In a mixed row the surrounding parent
size owns application spacing and the invisible control frame; standalone
functions retain their authored control size. Independent paint sizes, fonts and
colors remain intact. The rejected standalone-frame prototype displaced three
20 pt control contrasts by 7.62 pt; the parent-context rule corrects those cases
in both adjacent-control orders without a font or size whitelist.

Twelve independent Word controls improve in similarity and foreground recall;
all 116 glyph outlines match Word/source. Nine 12 pt surrounding-row controls
remain within 0.11 pt horizontally and 0.057 pt vertically. Three 16/20 pt parent
contrasts validate the context rule but retain an existing surrounding-prose
horizontal offset up to 1.37 pt and vertical differences up to 0.095 pt. These
residuals remain visible and do not trigger a size-specific correction.
Eighteen focused methods, 17 exact standalone/mixed-row neighboring PDFs and four
exact transfer smoke PDFs qualify one strict Release artifact. Three public cases
retain the existing visual thresholds. Structured names and deeper application
arguments still use complete fallback, with cancellation and work limits intact.


Delimiters now admit one bounded fraction or no-bar stack in each argument,
alongside literal arguments. Discovery reuses the existing fraction roles and
script levels, and measurement delegates to the shared argument box. The child
box supplies its complete enclosure height; literal arguments retain their
symmetric axis frame. The enclosing owner takes the maximum requirement across
arguments. Per-node control paint, translated child bars, separators, fixed
growth, source order and complete fallback remain intact. Wider nested fraction
children and unsupported properties/shapes still use fallback; no font or size
eligibility branch is added.

Eleven independent Word comparisons improve in similarity and foreground recall,
covering both modes, parentheses/brackets, stacks, child/parent size contrasts
and a function with a fenced fractional argument. The earlier literal-axis
prototype chose an oversized display parenthesis; the complete-box profile
corrects that case. All 57 matching outlines out of 69 and the 12 mismatches are
recorded individually. Three cross-size cases retain delimiter-variant or optical
glyph differences, with baseline residuals up to 4.47 pt. Before the semantic
spacing qualification below, the composed function also retained about 2 pt of
application-spacing difference. These are explicit
approximate results, not full Office parity. The minimum delimited-expression
constant in the [OpenType MATH definition](https://learn.microsoft.com/en-us/typography/opentype/spec/math)
does not by itself establish Word's remaining cross-size selection rule; no
unverified minimum or size-based correction is applied.

Twenty-one focused methods and 17 exact neighboring PDFs qualify one strict
Release artifact; four exact smoke PDFs verify its transfer. Four public cases
retain existing thresholds, including the low-fidelity larger-parent contrast.
An older literal-only regression was updated to keep testing unsupported deeper
nesting after fractions became eligible. Its failed report remains separate from
the final qualification. Safe managed code, cancellation/work bounds, bounded
staging and atomic publication remain unchanged.


Painted enclosing marks now use the font's `DelimitedSubFormulaMinHeight` to
classify a delimited subformula as an inner mathematical atom. That shared edge
supplies thin spacing around ordinary neighbors and after function names while
retaining binary/relation priority. The threshold classifies spacing; it does not
force a larger stretched glyph. The rule is inferred from independent Word
controls alongside the [MATH constant's documented role](https://learn.microsoft.com/en-us/typography/opentype/spec/math).

Nine Word controls qualify six improvements and three exact guards. Literal
16 pt and short-stack controls retain opening semantics; larger painted marks
around a 20 pt literal or a fraction receive the additional thin space. A tall
argument with fixed small parentheses stays unchanged. That contrast rejects the
prototype's child-height classification and supports painted-mark ownership,
without a font/size/function/fraction whitelist. All 77 glyph outlines match
Word/source. The retained function/fenced-fraction X residual falls from 2.09 pt
to 0.09 pt; the set retains differences up to 0.43 pt X and 0.072 pt Y.

Twenty affected methods and 18 exact neighboring PDFs qualify one strict Release
artifact; four exact integration PDFs verify transfer. Three public cases retain
existing thresholds. The earlier cross-size stretch and optical-glyph residuals
remain unchanged; this spacing correction does not claim to resolve them.


Matrix source now retains typed rows and cell arguments through the iterative
OMML parser, following the [matrix object definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.matrix?view=openxml-3.0.1).
Matrices have no implicit delimiters. Empty cells, unequal row widths, nested
known/unknown content, source argument properties and authored controls retain
their boundaries. An absent control remains distinct from an authored one;
cell formatting stays attached to its own run. Unknown matrix properties,
malformed/orphan row roles and unqualified shapes remain explicit, without
repeated raw subtree serialization. Unqualified spacing properties remain raw
source; the typed column and base-justification qualifications follow below.

Six Word references retain ordered cell meanings and geometry for display/inline
2x2 grids, a 3x2 grid, a fractional cell, a larger authored control and an Arial
normal-text cell. All 31 embedded glyph outlines match their source fonts. The
separate inspector's corrected UTF-16 range decoding supplies these text reports;
the Office PDFs and geometry remain unchanged. Sixteen source-capability methods,
14 exact source/native PDFs and four matching transfer PDFs qualify one strict
Release artifact. At this source checkpoint native matrix grid painting retained complete fallback;
the subsequent literal-grid qualification follows below.


Native standalone/fitting-inline matrices now paint nonempty literal cells in
centered columns. The [default column gap is one em](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.columngap?view=openxml-3.0.1).
The existing argument boxes retain logical widths, fonts and colors; the control
font owns ordinary row frames, occupied cell ink can expand them, and the whole
grid centers on the mathematical axis. `MathLeading` is applied once to the
composed formula. Metric-only controls reuse the existing conversion font owner
without preparing an unused PDF resource. An omitted control uses the surrounding
equation size while each painted cell retains its authored size. The rejected
first-visible-character/frame prototype displaced a mixed-size contrast by
10.61 pt; shared control ownership corrects it without a font/size whitelist.

Seven independent Word controls improve in similarity and foreground recall,
including 2x2/3x2 grids, explicit larger controls, an independently larger first
cell and mixed/all-Arial paint. All 34 glyph outlines match Word/source. Text
operations remain within 0.37 pt X and 0.18 pt Y; full visual results retain
explicit residuals. The all-normal grid emits only its painted Arial equation
resource. Twelve affected methods, ten exact neighboring PDFs and four exact
transfer smoke PDFs qualify one strict Release artifact. Three public cases keep
existing visual thresholds. Structured cells, empty placeholders, mixed rows,
explicit spacing properties and deeper grids remain in complete fallback. The
column and base-justification qualifications follow below.


Explicit matrix column declarations now retain typed count/justification groups.
An [omitted or val-less count defaults to one](https://learn.microsoft.com/en-za/dotnet/api/documentformat.openxml.math.matrixcolumncount?view=openxml-2.10.0);
[omitted justification defaults to center](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.matrixcolumnjustification?view=openxml-3.0.1).
Known groups apply left, center or right alignment within the same logical slots
used by implicit matrices. Declarations must cover the actual source columns in
this bounded capability. Malformed, incomplete or overdeclared groups retain
complete fallback. Schema validation retains the count as a scalar; one shared
resolver checks it against the actual remaining columns before filling an array
sized only from the source grid. Checked spans and cancellation preserve bounded
work without a count-driven allocation or another quota knob.

Six independent Word controls improve across explicit center/left/right,
three-column grouped alignment, omitted/val-less defaults and a larger control.
All 47 glyph outlines match Word/source. Text-operation residuals remain within
0.46 pt X and 0.18 pt Y. Fifteen focused methods and 12 exact neighboring PDFs
qualify one strict Release artifact; four exact integration PDFs verify transfer.
Three public cases retain the existing thresholds and visible residuals.
Earlier implicit matrices, metric-only font ownership and unsupported spacing,
structured/empty-cell and deeper-grid fallbacks remain intact.

Matrix base justification now accepts the documented
[`top`, `center` and `bot` values](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.basejustification.val?view=openxml-3.0.1).
Omitted and val-less declarations retain the centered default. Independent Word
controls support anchoring the first row baseline for top alignment and the last
row baseline for bottom alignment. Center alignment retains the existing grid
frame centered on the mathematical axis. The same row geometry and column slots
serve all three origins, without font, size or document-specific corrections.
Malformed, duplicate and unknown declarations preserve source and full fallback.

Eight Word controls improve across display/inline top and bottom alignment,
explicit/val-less center defaults, a larger control and a three-row grid. All 34
glyph outlines match Word/source; text positions remain within 0.37 pt X and
0.18 pt Y. Fifteen matrix-specific methods qualify one strict Release artifact,
with 12 exact neighboring PDFs and four exact integration PDFs at transfer.
Three public cases retain the existing visual thresholds and residuals. Spacing,
empty cells, mixed rows and deeper grids remain staged. Structured cells have the
separate bounded qualification below.

Matrices now reuse the existing argument-box grammar for one fraction/stack or
supported script cell. Ordered source discovery and one metric-only matrix owner
remain shared. Each cell's glyphs and fraction rules move together into the same
logical column slot; child ink can expand a row. Fraction constants use the
matrix control size while painted children retain their own authored sizes and
display/script roles. Literal cells retain their earlier geometry exactly.
Unsupported nesting, unknown cell roles and late missing glyphs retain complete
fallback without publishing a partial grid or isolated rules.

Seven independent Word controls improve, including the previously retained
fractional-cell reference; that reference's Office PDF, corrected inspection and
raster evidence are reused by exact PDF identity. Twenty-two matrix/shared-child
methods and 12 exact neighboring PDFs qualify one strict Release artifact;
four exact integration PDFs verify transfer. Three public cases keep the existing
thresholds. The outline audit matches 35 of 37 glyphs; two optical fraction-child
glyphs in the larger-control contrast remain different. Ordinary-size text positions remain within 0.57 pt X and
0.83 pt Y. A larger-control fraction aligns its children within 0.08 pt Y but
retains a 4.28 pt following-row residual. The rejected independent child-metric
prototype displaced the fraction children by 4.67 pt. These measurements qualify
bounded approximate support and retain the row-spacing concern for investigation;
they do not establish full matrix parity. No font/size/document-specific correction
branch is added.

Explicit default matrix spacing declarations now share the omitted-property
profile. The parser accepts val-less or zero `cGpRule`, `cGp`, `cSp`, `rSpRule`
and `rSp` through the same scalar validation. This follows the documented
[column-gap default](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.columngap?view=openxml-3.0.1),
[zero minimum column width](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.columnspacing?view=openxml-3.0.1)
and [default row-spacing rule](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.rowspacingrule?view=openxml-3.0.1).
At this checkpoint nondefault values retained fallback; the column-spacing
qualification below extends that profile. Malformed leaves, duplicates and
foreign declarations retain raw source and complete fallback. No quota is added.

Six independent Word controls have identical math text operations to their
omitted-property equivalents. Candidate PDFs reproduce the previously qualified
literal, grouped-column, base-aligned and structured grids byte for byte.
Thirteen source/native methods and eight exact neighbors qualify one artifact;
four exact integration PDFs verify transfer. Three public cases retain existing
thresholds. The 34/36 outline result and larger-control row displacement are
inherited unchanged from the structured-cell qualification. A separate six-control
assessment confirms row-clearance differences across row positions and matched
paint/control scales; its exact clearance metric remains unresolved, so no fitted
constant or conditional correction is introduced.

Matrices with explicitly hidden empty cells now retain the
[`plcHide` boolean and its val-less-on default](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.hideplaceholder?view=openxml-3.0.1).
Word comparisons show that hidden placeholders still reserve the control font's
U+2B1A dotted-square metrics. Empty slots reuse that advance and extent without
painting a glyph or preparing a PDF font resource. Empty rows retain the shared
control frame; empty columns retain their slots and gaps. Each empty slot charges
the existing cumulative work budget. Missing placeholder metrics, unknown cell
properties and wholly empty grids retain complete fallback. C90 extends the
known placeholder-visibility profiles after independent Word counterprobes. No additional cache, quota or native code is introduced.

Six Word controls improve across diagonal gaps, empty rows/columns, a larger
control and all-Arial paint. Text-operation residuals remain within 0.34 pt X and
0.12 pt Y; all 17 painted glyph outlines match Word/source. The rejected zero-width,
extent-free prototype displaced text by up to 12.89 pt X and 7.20 pt Y. Twenty-four
focused methods and 12 exact neighbors, including two visible-placeholder fallback
guards, qualify one strict Release artifact. Four exact integration PDFs verify
transfer; three public cases retain existing visual thresholds. The all-Arial
grid continues to publish only its painted font. Earlier structured-row and optical
residuals remain unchanged.

Authored matrix column spacing now uses the documented
[`cGpRule` units](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.columngaprule?view=openxml-3.0.1):
single/1.5/double em gaps, exact twips, or multiples of half an em. The gap amount
is ignored under the three standard rules, as specified. Numeric `cSp` minimum
widths retain twips independently of control font size. One resolver applies
these values to the existing logical column slots. The formula's logical extent
ends at its last argument, including a hidden placeholder box, while unused
trailing slot padding affects neither centering nor following-prose placement.
The rejected complete-slot extent displaced a centered minimum-width probe by
4.15 pt. No size or document-specific correction is needed.

Seven Word controls improve across minimum widths, every gap form, zero gap,
grouped justification and a larger control. All 36 glyph outlines match; text
positions remain within 0.37 pt X and 0.18 pt Y. Twenty-seven focused methods,
12 exact neighbors and four exact integration PDFs qualify one strict Release
artifact. Three public cases retain existing thresholds. Scalar source values
never determine buffer lengths; invalid rules/amounts, unsupported unit strings,
nondefault row spacing and earlier unqualified structures retain full fallback.

Matrices now participate in the existing bounded mixed-row composer. One
node-level admission guard serves standalone and mixed use, retaining the same
cell, spacing, placeholder and nesting rules. Ordinary, binary and relation
neighbors use the existing atom spacing. Adjacent grids retain independent source
controls and painted fonts. Independent Word contrasts show that the surrounding
row owns the effective control size for matrices and functions. That context now
passes through the shared measurement entry point, removing the composer's
function-specific dispatch. Standalone authored controls keep their earlier
geometry. The rejected standalone-frame-in-row prototype displaced an adjacent
matrix contrast by 10.43 pt Y; shared row context corrects both control orders.

Eight Word controls improve across display/inline ordinary and operator rows,
adjacent control orders, grouped minimum-width columns and hidden cells. All 70
glyph outlines match Word/source. Residuals remain within 1.68 pt X and 0.35 pt Y;
full visual gates retain explicit differences. Twenty-four matrix/shared-row
methods, 12 exact neighboring PDFs and four exact integration PDFs qualify one
strict Release artifact. Three public cases retain existing thresholds. Unknown
late matrices reject the complete row; cancellation, cumulative work and the
existing fit/page/publication bounds remain shared.

Bounded matrices can now be enclosed by existing ready-made delimiters. Ordered
control owners preserve both the outer painted marks and the inner metric-only
matrix without duplicating source runs. The existing conversion-owned font loader,
bounded caches and deferred operator-glyph publication remain shared. Literal and
fraction arguments keep their previous behavior; arbitrary deeper delimiter
nesting remains in complete fallback. Matrices expose their complete occupied
height through the same enclosure property already used by fractions.

Seven independent Word controls improve across round/square fences, three rows,
hidden cells, Arial paint and inner/outer control-size contrasts. The outline audit
matches all 28 cell glyphs and four of 14 enclosing marks (32/42 total). Complete
enclosure height improves four mark outlines and reduces the largest Y residual
from 6.50 pt to 3.23 pt; ten optical fence choices remain different. Text positions
stay within 0.69 pt X, with explicit vertical residuals. No fitted font/size rule is
added. Thirty-eight control/encoding methods, 14 exact neighboring PDFs and four
exact integration PDFs qualify one strict Release artifact. Three public cases
retain existing thresholds. An all-Arial grid publishes its painted cell and fence
fonts while the inner metric control adds no extra PDF resource.

Function arguments now admit the already-supported bounded matrix through the
same discovery and argument measurement. Name and matrix controls retain separate
metric owners; names, cell fonts and hidden slots keep their existing paint
semantics. Surrounding function/row context supplies the effective matrix control
size. Unknown or deeper function arguments remain in complete fallback.

Eight Word controls qualify seven improvements and one byte-identical existing
fenced-function guard, covering normal Arial paint, grouped columns, hidden cells
and inner/outer/mixed-row control contrasts. The outline audit matches 63/65
glyphs; the two inherited fence variants remain different. Ordinary controls stay
within 0.40 pt X and 0.35 pt Y. The larger function control retains a 3.15 pt Y
residual, and a mixed-row contrast retains 2.74 pt X. Twenty-three focused methods,
12 exact neighbors and four exact transfer PDFs qualify one strict Release
artifact. Three public cases keep existing thresholds. An all-normal application
publishes only its painted Arial resource despite two mathematical metric owners.

Complete function applications now own their outgoing operator spacing, rather
than inheriting the last atom of their argument. Independent Word controls show
thin spacing before ordinary or opening neighbors, unchanged binary/relation
priority, and adjacency before closing marks or punctuation. The existing shared
operator/punctuation pair supplies the latter rule. Names, argument boxes and
fraction rules retain their independently measured internal geometry. No font,
size, document or argument-shape qualifier is added.

Twelve new Word controls and one retained matrix-function holdout qualify six
improvements and seven byte-identical guards. Literal/normal/fraction ordinary
boundaries improve from roughly 1–2 pt X to below 0.07 pt; the existing mixed
matrix-function holdout improves from 2.74 pt to 0.74 pt. The audit matches 81/85
outlines; four inherited enclosing-mark variants remain different. Seventeen
application/shared-row methods, 12 exact neighboring PDFs and four exact transfer
PDFs qualify one strict Release artifact. Three public cases retain thresholds
and explicit visual residuals. A separate uniform cell italic-correction prototype
was rejected after larger-control and mixed-row regressions; its evidence is
retained and accepted cell widths are unchanged.

One supported inline equation in a plain body-table cell now reuses the body
wrapper and prepared box. The source run boundary owns its position; following
prose can wrap while the equation remains complete. Cell measurement and placement
share occupied width and vertical expansion, including existing declared-row
alignment and clipping. Unfit boxes retain ordinary prose pagination; only painted
math boxes keep their row atomic. Multiple equations, rich cell owners and wider
story/fragment support remain staged.

Six independent Word controls cover literal, fraction, script and root content,
following-prose wrapping and a centered exact-height row. All six improve and all
22 math glyph outlines match; positions remain within 0.064 pt X / 0.352 pt Y.
Visual SSIM ranges from 0.898 to 0.981, so existing strict thresholds remain under
review. Twenty-three focused methods, 23 exact neighboring/fallback PDFs and four
exact transfer PDFs qualify one strict Release artifact. All three added test
methods appear explicitly in the selected-method audit. No root rebuild, repeated
suite or Office/raster replay is needed to transfer the artifact.

A single inline equation that is the paragraph's sole visible content follows
Word's standalone display behavior in the body and supported table cells. Its
layout uses centered display geometry and display-style children even when the
paragraph has left/right alignment; the authored inline flag and equation object
remain intact. The temporary style selection participates in the same font/glyph
preparation and composed box measurement. Mixed prose keeps inline ownership;
unfit, multiple and unsupported source owners keep explicit fallback.

Eight independent Word controls cover literal/fraction/script/root content, three
body paragraph alignments and two cell alignments. All improve and all 30 math
outlines match, within 0.155 pt X / 0.340 pt Y. Twenty-nine affected methods, 22 exact
neighbor PDFs and four transfer PDFs qualify one Release artifact. Three cases
retain unchanged visual thresholds and explicit residuals. The prior inline-cell
artifact also passes nine portable methods on its exact DLLs; eight additional
cell alignment/indent/prose-owner contrasts match 32/32 outlines without another
renderer guard.

Standalone equations now paint in plain direct header/footer paragraphs through
the same prepared boxes and shared paint path. Static layout reserves their
occupied height, moves following prose, updates body frame limits and translates
footer content as one bottom-anchored block. Repeated first/default references
share one parsed equation. Mixed-inline, multiple, table and richer static owners
retain explicit fallback pending their own placement qualification.

Source part counts and original body-element references come from the existing
per-read story cache. Complete painting clears that part's ignored warning and
adds an approximate warning with its own part name. Partial painting retains
both warnings. A reader without a diagnostic sink still retains header/footer-only
source warnings from the parsed cache, without enabling unrelated feature probes.

Eight independent Word controls cover literal/fraction/script/root content with
following static and body prose. All improve, and all 28 math outlines match.
Math differences stay within 0.155 pt X / 0.430 pt Y; following header/body prose
is within 0.19 pt and footer prose within 0.11 pt vertically. Forty affected
methods, 33 exact neighboring PDFs (including SVG headers/footers and mixed-source
fallback) and four exact integration PDFs qualify one Release artifact. Three
public cases retain unchanged visual thresholds and explicit remaining gates.

Complete natural static stories now reuse the existing layout at source scale,
then map the whole text/math block through one print coordinate transform. The
same source frame owns header/body reservation and footer bottom anchoring. Body
placement restores a header displacement before selecting its print origin, so
that displacement is applied once. This shared rule fixes scaled plain footers
as well as mathematical ones; it adds no font/size-specific correction.

Fourteen Word controls qualify six improvements and eight byte-identical guards,
with 64/64 matching math outlines. Footer prose moves from below the page to
138.571 pt versus Word's 138.79 pt. Scaled math residuals fall below 0.20 pt and
header/body text below 0.075 pt. Forty-eight capability checks, 23 exact neighbor
PDFs and four transfer PDFs qualify one Release artifact. Unscaled static stories,
actual 120 pt controls and scaled SVG fallbacks preserve PDF identities. Declared,
field/image/table/rich stories, incomplete metrics/math and table-first or multiple
column body ownership retain existing paths. Visual thresholds remain unchanged;
remaining comment typography and paint differences stay explicit.


Natural static text now shares its content eligibility independently of left,
center or right alignment. The body/plain-paragraph policy retains its existing
left-alignment requirement; justified and richer static owners retain fallback.
Four of eleven independent Word controls improve, seven retain exact PDFs, and
24/24 mathematical outlines match. One Release artifact passes 59 focused tests,
35 frozen output identity checks and four root transfer smoke checks. This closes
the aligned static-frame slice; it does not establish broad mixed-content support.
Adding a bookmark around a supported equation still suppresses it in both default
and Word-compatible modes. Composition and document-level outcomes take priority
over further small static typography corrections.

Numbered paragraphs keep the first text run's authored formatting separately from
its inline revision decoration. Unchanged labels and their separators inherit
that source formatting, while numbering-level overrides and placement-time font
metrics retain precedence. Paragraph/mark revision ownership stays on its existing
path. This does not add support for deletion balloons or revised equations.
