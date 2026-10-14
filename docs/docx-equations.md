# Native DOCX equations

Native OMML rendering is in progress. A first native slice paints standalone
literal display rows in the DOCX body, with left, center or right alignment,
and fitting inline literal rows alongside prose.
The reader retains equation position,
display alignment, resolved run formatting, revision visibility and mathematical
structure for literals, scripts and fractions. Unsupported or incomplete
constructs retain an explicit source representation. Equations are separate from
prose runs, preserving existing field, bookmark, note and picture indices.

Painted rows report `DOCX_APPROXIMATED_EQUATION`. Inline formulas share the prose
baseline and move following prose by their measured width. A formula that would
require new wrapping retains the whole-row fallback. Scripts, fractions,
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
horizontal residuals up to 1.13 pt. Full visual gates still need review.
Scripts and fractions remain outside the
first rendering slice, with their source structure preserved.

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
