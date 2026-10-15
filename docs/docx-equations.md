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
unchanged. Multiple structured children and deeper compound rows retain fallback.
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
