# Native DOCX equations

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
