# Native DOCX equations

Native OMML rendering is in progress. A first native slice paints standalone
literal display rows in the DOCX body, with left, center or right alignment,
and fitting inline literal rows alongside prose. It also paints one display
fraction with literal numerator and denominator, using the font's OpenType MATH
axis, shifts, minimum gaps, rule thickness and leading. Default display scripts
with literal base/subscript/superscript rows now use separate measured boxes,
font script-style glyph variants and owned subset mappings.
The reader retains equation position,
display alignment, resolved run formatting, revision visibility and mathematical
structure for literals, scripts and fractions. Unsupported or incomplete
constructs retain an explicit source representation. Equations are separate from
prose runs, preserving existing field, bookmark, note and picture indices.

Painted rows report `DOCX_APPROXIMATED_EQUATION`. Inline formulas share the prose
baseline and move following prose by their measured width. Following prose can
wrap at its word boundary while the complete formula stays at its fitting source
line position. A formula that needs operator breaking or relocation retains
the whole-row fallback. Inline or nested scripts and fractions,
table/story placements, reference/revision decorations and unsupported explicit
line rules retain `DOCX_UNSUPPORTED_EQUATION` with part provenance
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
These results remain approximate. Argument controls, inline/nested structures and wider owners remain staged,
with their source structure preserved.

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
