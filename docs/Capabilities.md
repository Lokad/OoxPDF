# Capabilities

This is the **0.3.0 development scope**, including approximations and
explicit fallbacks. A successful conversion or absence of diagnostics does not
establish Office fidelity. Features outside this inventory remain unsupported.
See [release scope](PreviewRelease.md), [measured qualification](PreviewQualification.md)
and [diagnostics](Diagnostics.md) before relying on a conversion.

## Shared input and compatibility behavior

- Transitional OOXML DOCX/PPTX packages and common Strict DOCX vocabulary: prose,
  styles, fixed tables, images, links, supported math and related stories. Common
  point/inch/metric and percentage forms use the existing layout units. Strict
  dialect warnings remain; Strict PPTX, arbitrary Strict-only property forms and
  general bidirectional layout remain unsupported. Input bounds remain enforced.
- Exactly one AlternateContent representation: the first understood Choice,
  otherwise Fallback, otherwise no representation. Supported content inside
  ignorable ProcessContent wrappers is exposed to the ordinary format reader.
  Inherited namespace-prefix/QName declarations resolve at their original scope;
  nested wrappers and namespace wildcards are supported.
- MustUnderstand prefix lists produce a warning for unsupported or unresolved
  requirements in active content. The preview continues with partial handling;
  it does not implement complete MCE validation, preservation or schema processing.
  Other extension content keeps its existing reader fallback.
- Shared font resolution, TrueType embedding, glyph coverage fallback, managed
  image decoding and PDF primitives. Complex-script shaping and CFF embedding
  remain unsupported. Custom font identities must distinguish different programs.
- The bounded PNG decoder supports 1/2/4/8-bit grayscale and indexed images,
  plus 8-bit RGB, gray/alpha and RGBA. Palette transparency, transparent-color
  keys and explicit alpha use the shared image path. Sixteen-bit PNG inputs
  retain the unsupported-image fallback; complete PNG conformance is not promised.
- Cooperative cancellation, bounded caches/page staging, optional coarse host
  memory admission and atomic file publication. These are specific safeguards,
  not whole-heap accounting; see the [operating envelope](OperatingEnvelope.md).

## PPTX

Supported within the following constraints:

- Slide order/size, hidden-slide exclusion and master-artwork suppression by
  both slide and layout. Solid backgrounds select one slide/layout/master owner,
  including authored opacity. Common inherited shapes, theme colors and Latin theme fonts
  resolved from each slide's master. Missing master themes retain the existing
  presentation/first-master fallback; theme-override expansion remains unsupported.
- Authored DrawingML tint/shade transforms share linear-light color resolution
  across supported solid shapes, chart point fills and legend keys. Supported
  saturation modifiers retain source transform order.
- Common solid shapes, paths, straight lines, rotation/flips and nested group
  translation/scaling. Advanced/custom geometry and effects remain partial.
  Preset arcs pin finite adjustment angles to the preset range; non-finite
  numeric guides use defaults in both stroked and arrow-outline paths.
- JPEG/PNG pictures and bounded native SVG paint. Crop/fill clipping, source
  transforms and supported rectangle/ellipse/rounded/polygon picture masks share
  paint geometry. Unsupported masks retain diagnosed rectangular frames; bounded
  supported shadows/blur retain their qualified fallbacks.
- Fixed-grid tables with fills, explicit borders, merged continuations/row spans,
  text and calibrated built-in styles. Cells following horizontal/rectangular
  merges retain their source columns, text, styles and links. Vertical anchoring
  and typography remain approximate; arbitrary automatic table layout is not promised.
- External URI and internal-slide links on supported shapes, pictures, groups,
  frames, tables, text runs and chart text, using their resulting transformed bounds.
  Unresolved/action-only links are diagnosed and omitted.
- Native cached-data bar/column, line, area, pie/doughnut, scatter, bubble and radar
  charts, with supported titles, legends, axes/ticks, labels and source formatting.
  Bar/column series and point negative-inversion flags control fills, fallback
  outlines and label swatches; explicit point outlines remain authored.
  Axes/data labels share number formatting, including dates, numeric sections,
  scaling, fractions and supported accounting/currency forms. Unsupported syntax,
  formula-only data without cached values and unsupported chart kinds are diagnosed.
  Optional workbook styles and shared strings follow that workbook's own
  relationships; unreferenced parts do not affect cached chart rendering.
  Defined label containers use Office-observed component defaults while retaining
  explicit false values, deletion, custom text and number formatting. Series and
  their point labels own component defaults independently of inherited text style.
  Scatter/bubble category labels include the formatted X value; automatic labels
  retain their own text clipping region outside the data plot.
  Area labels share category-point text, links and formatting with line labels,
  using the center of each standard, stacked or percentage layer. Sparse source
  values retain labels; stacked areas preserve and label zero slots through
  missing values, including an explicitly requested Gap, as observed in Office.
  Standard areas using Zero also label missing slots as zero; Gap retains holes.
  Line point markers retain source symbols, sizes, fills and outlines across
  series and sparse data. Formatting-only overrides keep series geometry;
  explicit symbols own marker defaults. Legend/key ownership follows the series
  or point appearance, with marker keys bounded independently of text size.
  Automatic line labels default to Right; area labels default to Center.
  Measured line boxes determine baseline; occupied frames keep three-point
  clearance from authored marker extents independently of font size. Text
  backgrounds enclose padded text, with keys alongside.
  Plain line labels share chart-relative word/text-element wrapping and
  authored no-wrap paragraphs. Factor-sized bodies center text and use ellipsis
  when height is exhausted. Their stroke-key reference and marker capacity follow
  chart-root typography independently of the label's own font. Rich, rotated,
  edge-sized and unknown-property labels retain the existing approximate path;
  key and absolute text geometry still have measured residuals.
  Line/area label layouts use the outer chart frame for factor offsets, edge
  origins and factor-sized extents. Sized labels center text, retain measured
  baselines and fit before offsets; zero or missing extents leave no label.
  Invalid sized edge origins retain automatic placement. Mixed edge extents,
  typography, key spacing and multiline fitting remain approximate.
  Defined point labels own component defaults; a defined series owns its point
  declarations. Fully disabled Cartesian plots suppress plot-point requests,
  while series and pie/doughnut points retain independent label scope.
  Label-set deletion hides default labels while preserving explicit child
  requests. Defined series and points own their deletion default; absent plot
  containers keep series settings local to their source series.
  Point labels without local formatting retain their container's text style,
  including geometry and component overrides. Explicit point shape/text markup
  starts a local scope from chart defaults; series formatting keeps its own
  scope. Font provenance and raw source presence remain inspectable.
- Failed slide nodes rewind partial paint/annotations and report a diagnostic;
  neighboring nodes may still render. Inspect diagnostic severity and fallback.

Approximate or restricted:

- Textboxes support paragraphs/runs, Latin greedy wrapping, font styling and
  alignment. Font metrics, kerning, autofit, distributed text, overflow, rotation,
  columns and table/bullet combinations retain qualified approximations. Nominal
  advances and explicit-zero kerning apply only in the qualified unrotated,
  horizontal, single-column no-wrap/noAutofit frames; no blanket parity is claimed.
- Automatic numbering covers the implemented Arabic, alphabetic and Roman
  period/parenthesis forms with supported levels/starts and sequence resets.
  Empty/plain paragraphs and changed settings affect continuation; richer formats
  and table combinations remain restricted.
- Bullet size/font/color follow the nearest authored choice, including explicit
  follow-text resets; inherited styling survives a direct marker. Typography and
  layout remain subject to the approximations above.
- Clustered/stacked bar/column, line-with-markers and plain pie are closer-parity
  targets. Area, scatter/bubble, secondary axes, percent-stacked and leader-label
  layouts remain approximate; trendlines are diagnosed and omitted.
  Area label layout remains approximate. Label baselines, legend-key shapes and
  spacing, polar label formatting and plot-level point-label fallback retain
  known differences from Office; restored text can still fail image thresholds.
  Manual point-label coordinates and multiline label wrapping remain approximate.
- Automatic chart layout, label strips/baselines, dense intervals, source crossings,
  reversed axes and manual-layout combinations retain their semantic guards.
  Qualified automatic single-line plots share axis/label placement, including
  captions beside source crossings with titles and bottom legends. Authored
  coordinates, composed/secondary plots, unsupported scales and unresolved fonts
  keep their fallbacks. Empty manual metadata retains automatic measurement.
- Column, line and area category labels share measured rotation, spacing and
  wrapping. Automatic outer captions share height with the plot and may truncate
  within their footer. Authored plot geometry remains fixed. Automatic vertical
  value labels retain endpoints and omit overlapping ink; fixed intervals remain
  unchanged. A measured automatic column scale also retains its allocated captions
  through later authored inner-plot placement. Vertical numeric strips measure the same automatic interval policy
  used to emit their labels. Resolved Cartesian spans below one retain proportional
  coordinates
  and tick intervals. Negative-only data shares positive-axis bound rounding and
  preserves authored limits. Column captions and measured line layouts follow
  numeric low/high positions and value direction; axis position does not change
  column data lengths. Remaining caption placement, dense allocation, automatic
  axis ranges and font geometry remain approximate.
  Numeric strips can reserve space for an authored range endpoint even when no
  major tick is painted there; measurement and tick emission have distinct roles.
  Standalone Cartesian plots with one supported value axis share a retained
  authored linear minimum across labels, tick marks and grids, including
  horizontal bars; only aligned upper
  endpoints become ticks. Single column/line/area plots with one supported value
  axis also anchor an automatic lower bound to an authored maximum and major
  unit. An authored major unit governs their automatic bounds and headroom,
  including signed, zero and fractional data. Translated lower bounds retain the
  same interval count. Labels, tick marks and grids share those source intervals;
  automatic minor grids subdivide a major interval into five unless overridden.
  Automatic plots admit oblique category-caption frames horizontally using the
  same measured geometry as painting. Generated ellipsis markers remain visible
  outside the occupied vertical content band; horizontal bounds include their
  full width. Authored ellipses remain ordinary text. Manual geometry, clipping
  and horizontal wrapping retain their existing policies. Text fitting and
  short-frame caption geometry remain approximate.
  Built-in chart fitting retains its default reference when overridden role or
  global fonts change; effective fonts still determine glyph geometry.
  Category/value axis sizes in those charts inherit from chart text and local
  axis formatting. External role sizes do not fill a missing size; other role
  formatting and custom-style fallbacks retain their existing owners.
  Automatic single clustered-column plots with one supported linear value axis
  select bounds and intervals from measured title, legend and numeric-font capacity.
  Numeric strips, labels, grids and data coordinates consume that same scale.
  Authored inner data rectangles retain this automatic allocation while keeping
  their source placement; the measured legend footprint is reused.
  Automatic limits and intervals remain approximate, including large-font and
  range-boundary cases; authored outer frames, manual furniture, composed and
  secondary axes retain their fallbacks.
  Authored bounds and units remain authoritative within the supported
  single-axis scope above. Remaining horizontal frame geometry is approximate;
  Unqualified composed/secondary interval ownership retains its fallback.
  Bar/line and additional-column plots sharing the primary value-axis identity
  emit that axis's numeric labels once. Distinct secondary source axes retain
  their independent labels and visibility; composed scale/frame geometry remains
  approximate.
  Clustered-column and standard-line plots sharing one reciprocal category/value
  pair use a common source-data range and declared linear intervals. The resolved
  scale feeds numeric measurement, bar/line placement, labels and grids, including
  partial bounds and empty sibling plots. Their automatic column frames share
  the measured numeric header and category band across value-axis reversal,
  crossing and label sides. Manual furniture, composed stacking, unpaired and
  logarithmic scales retain their existing fallbacks. Fully automatic capacity,
  font geometry, clustering and some image comparisons remain approximate.
  Consistently oriented clustered horizontal plots also share the source scale
  and category crossing, including signed values, authored bounds and reversal.
  Standalone clustered horizontal bars use the same source-interval owner;
  stacked horizontal bounds keep their existing fallback. Automatic clustered
  horizontal frames reserve the numeric-caption band on its source side,
  independently of the numeric stroke. Low/high/next-to placement follows the
  paired category direction, including reversal, with measured titles and
  top/bottom legends. Authored furniture, axis titles, side legends and interior
  ordinal crossings retain their existing placement fallback. Clustered
  horizontal categories mirror bar rows and internal series offsets together,
  preserving source values and paint across single/composed plots. Captions,
  data-label/key frames and linked runs follow that direction. Remaining
  typography and unqualified layout combinations stay approximate.
  A single reciprocal category/value pair shares horizontal crossing geometry
  across plot admission, strokes, ticks and labels; low/high label bands remain
  separate from the stroke. Clustered bars and label swatches classify inversion
  relative to that origin. Unpaired/composed/logarithmic placement keeps its
  fallback. Nonzero stacked origins and remaining label/stroke geometry stay
  approximate.
  Automatic centered horizontal labels measure their text, key reservation and
  font line before centering the occupied frame on the painted bar. Manual/custom
  labels, key-only content and other positions retain their existing fallback;
  key-only painting can differ from Office's omitted label.
  Scatter/bubble axes with a retained authored minimum share that interval origin
  across numeric labels, major ticks and grids; only aligned upper endpoints are
  ticks. Numeric pairs share direction across connecting paths, markers/bubbles,
  labels and grids. Crossings use the perpendicular scale; explicit numeric
  crossings, range edges and low/high/next-to labels retain distinct placement.
  An explicit `between` declaration shifts a numeric crossing by half a value
  unit; midpoint/omitted declarations and minimum/maximum edges keep their own
  meaning. Fixed-bound scatter controls match Office axis coordinates and reversal
  movements within0.05pt. Automatic ranges/steps, bubble plot padding, tick paint
  and remaining text geometry are approximate. Logarithmic axes retain their
  existing fallback.
  Scatter connections follow series line visibility across omitted, marker-only,
  line and smooth styles. Missing series smoothing follows PowerPoint's true
  default; explicit false keeps straight segments and series line noFill hides
  connections. Unknown/rejected styles retain their fallback. Curve fitting and
  automatic interval fitting remain approximate. Scatter axes with authored
  positive major intervals or inferred vertical intervals measure curve
  overshoot and share resolved bounds
  across measurement, labels, marks/grids and placement. Independent controls
  cover partial bounds, fractional/negative ranges, constants and point order;
  inferred vertical units also preserve source-relative intervals when an upper
  bound is authored, across tested frame heights and label sizes. Existing fully
  automatic clustered axes retain their interval density. Crossing coordinates
  agree within0.005pt. This does not establish exact curve
  shaping or application-view parity. Missing scaling and rejected or logarithmic
  declarations retain their existing fallback.
  Area-sized bubbles with authored horizontal bounds and a positive major
  interval share measured outer and inner containers, circle sizes, axis ranges,
  labels, clipping and circular legend keys. Source `bubbleScale` percentages,
  including zero, and subunit sizes are preserved in this scope. Independent
  font/frame/scale/title/series controls match native radii within0.005pt;
  caption, center and legend placement remain approximate. Automatic horizontal
  intervals, authored container geometry, composed plots, width-sized and negative
  bubbles retain the existing fallback.
  Automatic scatter X intervals use the available authored inner factor width
  and axis font size; numeric captions, marks and legend measurement share the
  resolved scale. Inner manual plots also resolve factor-position offsets together
  with right-edge widths; equivalent declarations preserve their geometry and
  captions. Three independent edge controls match complete axes within0.04pt,
  correcting displacement up to80pt and extra axis captions. Automatic/outer X
  plot widths, edge-height generalization and bubble automatic
  intervals retain their approximation. Required integer-zero placeholders retain
  padding across the shared number formatter's existing sections, literal masks,
  grouping and scaling. Source-owned inner factor plots admit measured endpoint
  and side captions inside the chart frame, preserving width when possible and
  shrinking when required. Hidden series lines do not change caption admission;
  numeric captions stay centered on ticks even when wider than their nominal slot.
  Independent controls match Office X-axis coordinates within0.04pt. Vertical
  admission, other plot layouts, source axis paint and remaining typography retain
  their approximation; this does not establish complete chart parity.
  Inner column/line/area plots with one linear numeric axis and horizontal,
  single-line ordinal captions admit numeric side bands and axis endcaps using
  their measured fonts. Category and numeric constraints activate independently;
  geometry shared by plots, labels and links follows the resulting rectangle.
  Independent Office PDFs correct displacements of28.4pt vertically and38.38pt
  horizontally; two Arial baseline residuals remain 1.18–1.52 pt. Outer layouts,
  authored title/legend geometry, explicit numeric crossings, secondary/log axes
  and rotated/wrapped categories retain their existing paths. Native-font
  geometry checks remain separate from portable content checks at font-dependent
  admission boundaries. Numeric-only inner overflow retains the fitted source
  interval and its opposite edge before restoring complete caption spans;
  category-endpoint overflow retains whole-interval admission. References remain
  uncalibrated against the application view.
  Explicit overflow in inner rotated captions without numeric readmission keeps
  wrapped source lines; authored `wrap="none"` keeps the full source on one line.
  Measured line stacks share the physical chart clip. Right-angle wrapped bodies
  retain their fitted source-reference edge on either side and rotation direction.
  Long-caption reference extents remain approximate: right-angle font/size
  residuals below one point and larger oblique displacements remain unresolved;
  matching the data rectangle alone does not establish caption-placement parity.
  Complete factor-sized outer column/line/area declarations also center their
  source area before allocating numeric and category-caption bands inside it.
  Shortened endpoint captions release their measured space; a present category
  axis retains a nominal minimum band when labels/ticks are hidden. Data axes,
  plot fill and linked label geometry share the admitted inner rectangle.
  Numeric-caption admission precedes automatic category angle/skip fitting.
  Vertical numeric captions keep at least their measured text width during
  painting and clipping, including formatted values and either label side.
  Side gaps follow measured normal-line geometry; admission, painting and
  chained strips share that gap. Unavailable metrics retain the nominal fallback.
  This remains a measured approximation: Calibri spacing and composed absolute
  placement retain independent residuals.
  Rotated single-line category fragments allocate logical text advance and
  normal-line spans against the authored width before numeric carving. Source
  pitch and ordinal reversal own reference movements; measured painting and
  clipping retain their anchors. Mirrored outer captions share their measured gap;
  authored label offsets add normal-line spacing rather than scaling the full
  baseline reserve. Numeric-side bands, baseline typography and empty-caption
  spacing remain approximate, as do incomplete declarations and
  the exclusions above. Some Office references paint category text as vector
  outlines; their validation uses graphics/raster output as well as extraction.
  Automatic oblique captions retain logical source-reference bands and clip
  painted overhang at the physical chart frame, avoiding a second full-ink inset.
  Line/area captions, data points and line data labels share category direction
  while retaining source indices and point formatting. Ordinary clustered
  columns with one paired ordinal/linear numeric axis share that direction,
  including intra-cluster series offsets, sparse slots, labels and links.
  Plain column labels use their painted bar centers. Stacked, composed and
  secondary-axis column reversal retain their existing fallbacks.
- Automatic bottom legends with filled keys admit naturally fitting names on one row and
  align wrapped rows in uniform columns, including the incomplete final row.
  Composed bar charts collect all series entries before measuring the legend.
  Whole-row height admission retains a source-ordered prefix and diagnoses
  clipped overflow. Manual frames retain their authored bounds. Spacing and
  absolute placement remain approximate.
- Single line/area plots share measured numeric strips, title clearance, bottom
  legend packing and source caption crossings. Overlays retain plot space;
  authored/composed frames and unsupported scales retain their fallbacks. Range
  endpoint advance is reserved even without a painted endpoint tick. Automatic
  font sizing, unusual axis declarations and remaining geometry are approximate.
  Paired categorical axes use the numeric crossing and category orientation for
  the numeric axis side; explicit low/high labels follow the paired category
  direction, including reversal. Measurement, painting and secondary-axis slots
  share that rule. Unpaired,
  date/numeric and interior ordinal crossings retain their placement fallback.
  Automatic no-title line/area plots with a right non-overlay legend admit
  horizontal, unwrapped category captions and numeric labels together, keeping
  the measured legend anchor. Area and supported bubble plots share side-legend
  allocation from actual text advances and the selected font's normal line box;
  area key size and row spacing use that line box too. Painted key positions,
  text baselines and remaining gap residuals stay approximate. Rotated/wrapped captions retain their prior
  layout, as do authored/composed frames and unsupported scales.
  Column category strokes and captions use the category crossing independently
  of the numeric axis crossing. Measured caption bands share line/area placement
  rules. Clustered columns use the category origin for lengths and inverted fill;
  nonzero stacked and horizontal origins retain their existing approximation.
  Authored maximum crossings select the high caption/tick side; a numeric crossing
  at the same bound keeps its numeric placement. Explicit low/high captions retain
  precedence over next-to placement.
  Numeric-crossing ticks retain the authored physical axis side; implicit
  crossings mirror that side with value direction. Next-to captions at or beyond
  a reversed numeric upper bound reuse the measured low-side band. Office's
  displaced ticks for that combination remain approximate; the numerical axis
  stroke and fill retain their separate, independently validated coordinates.
  Standard area fills and the first stacked layer close on the resolved category
  crossing; later stack layers retain their cumulative lower values and source
  series positions.
- Standalone line and area plots follow the source category-slot mode for data
  and captions. Line markers and data labels share those coordinates; category
  ticks delimit slots or coincide with points as specified. Sparse points retain
  their source positions. Composed line plots retain their separate placement
  fallback; category direction, automatic spacing and label fitting remain
  approximate. A lone `midCat` point and caption use the source endpoint, moving
  to the opposite end under reversal; between-slot categories remain centered. Line
  markers retain their full horizontal extent at category edges; numeric-axis
  clipping remains active.
- Standalone line plots with resolved theme defaults inherit curve, stroke and
  automatic-marker paint independently of explicit smoothing. Theme line widths
  and color transforms feed the plot and legend; authored line width/color and
  no-fill overrides remain distinct from automatic marker paint. Missing themes
  retain the existing approximate fallback; composed/scatter defaults stay separate.
- Qualified gallery defaults, authored stroke/fill precedence, line smoothing,
  markers and automatic six-accent bar palettes are implemented. Classic/composed
  gallery roles, minor-stroke defaults, automatic tick fitting and arbitrary
  color-style methods remain incomplete. Category bar legends share point paint
  and wrap within the chart frame; crowded frames diagnose clipped entries.
  Automatic bar theme colors follow declared series indices and the shared
  declared index span, including sparse indices and empty siblings. Bar paint,
  legend entries and label keys share this recipe. Lone point-styled category
  keys retain point fills; series keys preserve negative inversion without
  adopting authored point fills. Source-index handling for custom and single-
  accent recipes has not been generalized. Plain automatic outside-end labels
  above clustered columns use measured word/text-element lines or authored
  no-wrap paragraphs and admit their fitted text/key frame inside the chart.
  Mandatory paragraph breaks remain. Whole text elements wider
  than the fitting width retain their full glyphs and normal paragraph flow;
  their fitted source width owns alignment, with residual native font geometry.
  Position-only manual labels share that measured frame: factor offsets move its
  admitted automatic position, edge coordinates use the full chart, and text,
  shape and key retain one placement. Factor-sized manual labels retain their
  requested text extents and a separate key reference. Wrapped text is centered
  vertically and uses word-aware ellipsis when the height cannot hold all lines;
  no-wrap keeps complete paragraphs. Missing dimensions, explicit zero extents
  and rejected declarations follow independently observed Office behavior.
  Empty point text-property bodies, edge-sized labels, explicit overflow/clip modes,
  rich-text,
  negative-end, stacked and other placements retain their
  existing paths and can still clip long text; no general collision avoidance
  is claimed.
  Authored inner Cartesian plots resolve factor positions and edge-sized
  dimensions together on both axes. Caption admission near chart edges remains
  approximate, with unresolved differences from Office. Distinct outer layouts
  retain their existing paths.
  Existing independent Office thresholds still apply.
- A single-series line plot with point variation shares source-indexed paint
  across incoming straight/curved segments, markers, category-name legend entries
  and data-label keys. Explicit point/series strokes and series marker appearance
  retain precedence. Omitted variation follows the qualified enabled default;
  explicit false and multiple series retain series legend entries. Automatic point
  markers repeat nine symbols; a series marker of `none` suppresses them. Plot
  marker flags alone do not suppress automatic line markers. Authored/composed
  styles, dense layout, label placement and font geometry remain approximate.
- Legends without an overlay flag follow PowerPoint's overlay default. Specify
  `<c:overlay val="0"/>` to reserve plot space. Legend geometry remains approximate.
  Bottom line legends share measured wrapping with horizontal bar legends;
  empty layout metadata preserves wrapping; authored manual frames constrain overflow.
- Authored chart titles survive automatic-title deletion flags. Automatic-title
  inference remains limited to an explicitly enabled, unambiguous series name;
  inferred titles resolve built-in/default sizes and honor explicit chart text
  sizes before measured title-line reservation. Authored-title defaults, wrapping
  and composed/manual title geometry remain approximate. Empty-title Office
  placeholders are not synthesized.
- Horizontal shape text renders an ellipsis marker when vertical overflow hides
  later lines. Non-horizontal text and chart-title contexts retain diagnosed
  clipping fallbacks. Complex effects and typography may be simplified or omitted
  with diagnostics.

Unsupported: stock/surface/3D charts, SmartArt, live OLE, macros/ActiveX,
audio/video, animation/transitions, unrestricted custom effects, bidirectional
layout and general OpenType shaping. Static fallback diagnostics are listed below.

## DOCX

Bookmarks before inline equations retain their destinations across empty or
hidden runs; intervening visible text, pictures and text boxes keep their boundaries.

Inline page/column tokens inside cells leave supported paragraph content together,
including equations, pictures and navigation. Actual paragraph boundaries remain
present. Standalone page-break paragraphs keep their existing cell flow handling;
neighboring equations still prepare. General body column flow, standalone cell
break geometry and revision pagination keep their stated limits.

Final/Simple joins across deleted paragraph marks retain supported inline
equations, pictures, links and bookmarks in source order. Consumed marks keep
their revision provenance. Original/AllMarkup retain the source paragraphs;
unconsumed mark revisions keep their existing equation eligibility fallback.
General paragraph-format changes and unrestricted revision layout remain partial.

- Optional styles, settings, numbering, fonts/themes and related stories follow
  the document's owning relationships. Unreferenced package parts retain the
  corresponding empty/default behavior rather than configuring active content.
- Header/footer rendering diagnostics follow the same typed internal story
  references as their readers. A relationship alone does not select a story;
  selected unsupported payloads and package-level warnings remain diagnosed.
Supported within the following constraints:

- Body paragraphs/runs, common run/paragraph/document defaults and styles,
  margins/page sizes, paragraph spacing/line-height inheritance and basic page
  breaking. Qualified section/page/column breaks retain their explicit limits.
  Authored point-based spacing at explicit page starts follows local paragraph/run
  break ownership in unscaled layouts.
- Latin font resolution through explicit families, themes and font-table alternates.
  Undeclared prose/list families use the documented Times New Roman fallback when
  available, after an explicit resolver document default. Missing faces/glyphs are
  diagnosed; no general shaping or font-independent Office parity is promised.
- Simple decimal/bullet numbering and twip-based indents/tabs. Numbering font
  selection follows numbering-level values and the paragraph mark/style/defaults,
  independently of first-run styling. Supported inserted marks preserve adjacent
  text/equations; old/new review renumbering and marker revision decoration remain
  incomplete. Deleted/unnumbered marks keep their existing fallbacks.
- Fixed-grid tables with fills, common collapsed border styles/conflicts,
  nil/none suppression, repeated headers, split rows and cell content. Actual page
  occupancy controls following-prose admission. An empty terminal paragraph after
  a nested table keeps its source anchors without adding occupied height;
  its bookmarks use the placed insertion frame. Ordinary text and references
  keep their frames. Scaled Review tables share
  printed body capacity for authored and nested rows. Oversized kept rows retain
  atomic overflow; arbitrary rich-row fragmentation remains approximate.
  Conditional bands, columns,
  rows and corners share one style cascade; direct formatting remains above styles.
  Table-look bitmasks apply when no explicit flags are present; an explicit
  declaration enables omitted flags, including band suppression. Absent/empty
  declarations keep Word's default look; inspection retains raw source tokens.
  Bands require a positive resolved size; zero or an omitted size disables the
  axis, following [Word's band-size default](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/77519f5b-08b3-4547-9f5d-0c4d87854220).
  Custom-style padding stays inside the grid; explicit indentation and TableNormal
  keep their own anchors. Table/cell formatting inherits earlier parent
  declarations; empty and text-only styles retain forward references. The intrinsic
  non-custom Normal Table name supplies built-in padding and ignores authored
  children; custom styles and descendant/direct overrides retain their formatting.
  Other built-in style defaults remain approximate.
  Adjacent tables with the same resolved style share logical row membership in
  body, cell and static stories; paragraph/style boundaries, different fixed/autofit
  modes and positioned tables keep them separate. Fragment row looks, margins and borders retain their owners;
  joined grids map spans and leading omitted columns while the first table owns
  the frame and preferred width. Unconstrained automatic tables use measured
  content and cell insets rather than the authored grid as a fixed width.
  Supported intersecting spans share measured content columns and project their
  subdivisions onto one placement grid; merged cells can enlarge ordinary columns.
  Explicit preferred widths and fixed layout retain their sizing rules. Leading
  gaps, unmeasured nested blocks and crossing spans without an independent content
  partition retain existing width fallbacks. Arbitrary spanning-cell autofit and
  positioned-table placement remain approximate.
  Complex implicit borders,
  vertical merges and nested/fragmented layout retain qualified restrictions.
- Inline raster/SVG pictures in qualified body/cell/header/footer containers,
  including supported crop/flips/rotation and source frames. Picture-only body
  paragraphs use occupied height and collapsed spacing. Qualified automatic
  picture-only paragraphs share their measured row frame across body, cells and
  static stories. Expanded ordinary spacing follows the occupied math/picture
  row once; prepared display equations retain their own spacing owner.
  Inline pictures follow
  effective run hiding, including inherited styles and explicit false overrides;
  drawing-level `docPr hidden` alone retains Word's print behavior. Floating pictures/textboxes
  have bounded positioning, clipping, z-order and first-pass wrap exclusions;
  arbitrary Word reflow around floating content is unsupported.
- Custom footnote/endnote marks support a nonwhitespace BMP character immediately
  after the reference in the same XML run. Each retains its own link alongside
  other note references and supported inline mathematics. Note layout remains approximate.
- Default headers/footers, simple page fields, supported notes and source-owned
  links/bookmarks. Omitted header/footer distances use Word's half-inch default,
  independently of body margins; explicit distances retain their values.
  Static-story placement shares physical frames with painting
  and annotations. Without printed balloons, unit-scale Review pages preserve
  the Final text, equation and picture origins on note-only continuation pages.
  Complete unscaled endnote rows that fit a fresh note frame move intact to
  the next page; source-affined picture-only paragraphs share this admission when
  every picture survives its cell, ancestor and explicit viewport clips.
  Measured paragraphs and rows retain source order. In mixed
  paragraph/table notes, oversized measured rows also defer from occupied frames
  before bounded clipping; empty frames make progress. Height-only continuation,
  scaled, already-clipped and unsupported compositions retain their fallbacks.
  General fragmentation of note rows remains unsupported. Automatic complete rows
  with measured text/math lines or complete unrotated source-cell pictures
  extending below the note frame report `DOCX_CLIPPED_NOTE_TABLE_ROW`, including
  partial clipping. Rotated pictures retain their diagnostic fallback.
  Adjacent ordinary unscaled endnotes share paragraph before/after spacing even
  after an earlier table; first-note/page-turn geometry remains approximate.
  An already-started final next-page section keeps its document-end notes even
  without body text. In known unscaled frames, its default empty paragraph uses
  ordinary paragraph flow; remaining font metrics still affect the note origin.
  Document-endnote boundaries include the trailing ordinary paragraph's measured
  automatic leading. In known unscaled frames, a complete measured first note
  owns its start spacing even when later notes require continuation pages.
  This does not admit the whole group to fitting-flow or counter corrections.
  Scaled and unsupported source/static frames retain their existing policy.
  Empty continuous/parity sections retain existing fallbacks.
  Complex/continued/page-dependent note compositions and arbitrary
  nested stories retain explicit fallbacks and documented geometry residuals.
  Ordinary endnote placement follows the document-wide position, matching Word;
  section numbering and authored position metadata remain independent.
  Registered endnote separators retain paragraph spacing. Before-space moves the
  rule and content; after-space advances content below the rule, including empty
  continuation separators. Absent and unregistered separators share one reader-normalized default.
  Exact line-height, rich separators and unqualified note frames retain their
  existing fallbacks; default spacing normalization is applied through admitted
  paragraph flow.
- Simple/complex fields with stored results, including supported nested cached
  fields and bounded dynamic placeholders. Locked PAGE/NUMPAGES fields preserve
  their stored results, including empty results; nested fields own their locks.
  Complete empty stored results also preserve neighboring supported equations.
  Supported picture-only and equation-only stored results retain their content
  and neighboring equations in body, cell and default header/footer paragraphs. They share the
  existing inline layout; stored fields are not evaluated.
  Complete repeating table-header rows retain PAGE from the original admitted source row;
  ordinary body cells and page headers/footers keep their own page inputs.
  This does not evaluate general Word
  fields, live cross-document data or arbitrary aligned/wrapping field layouts.
- Final, Original, SimpleMarkup and AllMarkup views, with preserved-layout and
  Word-compatible review geometry kept distinct. Public defaults are Final with
  preserved document layout. See [markup modes](DocxMarkupModes.md) for revision,
  comment, field and margin behavior.

Approximate or restricted:

- Supported native equations share ordered inline measurement with text/pictures.
  Fitting multiple formulas, bookmarks and unchanged-format run splitting preserve
  content in qualified body paragraphs, simple cells and direct headers/footers.
  Resolved paragraph-formatting history and a unique owned paragraph-mark format
  change retain neighboring equations and source navigation, with or without a list
  label. Duplicate or unowned property changes keep their existing fallback;
  bold/italic descriptions share run and paragraph-mark resolution. Word-compatible
  Review anchors owned paragraph-mark captions to the measured text strut; footer
  captions retain their printed frame. For measured paragraph-mark formatting
  ranges, revision bars cover the occupied row and its resolved trailing gap.
  Owned inline pictures preserve separate changed paragraph bands. Unmeasured
  rows and floating/table compositions retain their existing bar fallbacks; bar
  edges and complete-page appearance remain approximate. Other changes
  retain generic summaries; balloon placement and complete-page fidelity remain
  approximate.
  Tables inside mapped inline header/footer boxes retain their rows, equations,
  links and bookmark targets in the painted frame, including nested tables.
  Word-compatible Review preserves first-row prose baselines alongside equations
  in qualified body/header/footer boxes. Explicit cell spacing, vertical merges
  and unsupported alignment retain their existing approximations.
  Horizontal cell padding and complete-page appearance remain approximate.
  Adjacent display containers combine using the later container's justification;
  subsequent display rows share the first effective display row's owner.
  `centerGroup` and the omitted default center one maximum-width group across
  consecutive display rows; explicit `center` centers each row. Prose ends a
  group. Local declarations override the retained document math default.
  Ordinary source text, including spaces and hidden text, links and pictures
  retain inline math without discarding neighboring formulas. Hidden text is
  unpainted; hidden manual-break geometry remains approximate.
  Authored body page breaks retain supported formulas, prose, pictures and note
  references in source order. Links can span the resulting pages; bookmarks target
  the first placed source content. This content-preservation coverage does not
  establish Office pagination or visual parity for arbitrary combinations.
  Bookmarks on ordered inline pictures and boxes use the occupied line top.
  Mid-line objects supply their horizontal extent; line-start objects retain the
  containing frame's viewport origin. Legacy picture-only paragraphs keep their
  existing navigation fallback.
  On a line with an active text/equation hyperlink, an adjacent picture in the
  same source range extends its clickable bounds using the painted frame.
  Ordered final lines retain the resolved paragraph after-gap; mixed body, cell,
  header and footer controls match all four Word-PDF rectangle edges within 1 pt.
  Root-row math links merge fragments per placed line and share its occupied
  frame; links inside mathematical arguments retain local measured bounds.
  Drawing URI targets stay separate. Image-only ranges, rotated/clipped pictures
  and unresolved placement owners retain their fallbacks. Other vertical click
  bounds remain approximate; application-view calibration is still open.
  Prepared display rows reuse math boxes and source frames. Mathematical layout
  emits DOCX_APPROXIMATED_EQUATION; omitted unsupported math is diagnosed separately.
- Bounded literal/fraction/script/root/delimiter/n-ary/function/matrix/cases forms,
  supported revisions, empty arguments and ready-made matrix fences retain
  individual metric/paint ownership. Unknown/deeper structures, continuation or
  unresolved merge owners, unfit objects and arbitrary fragmentation keep fallbacks.
  See [native equations](docx-equations.md) for the authoritative construct and
  placement limits; support for one construct does not imply unrestricted nesting.
- Pagination, line leading, kerning, numbering, table/story geometry and complex
  composition remain approximate. Explicit-page-start contextual, automatic, line-based
  and print-scaled spacing retain existing fallbacks. Font/size/document-shape guards protect independently
  qualified behavior; they are not promises for every combination.
  Bookmarks and unchanged-format hyperlinks preserve qualified plain-body minimum
  line placement; their navigation targets follow the placed content.
  Ordinary minimum-spaced body text uses each measured line's natural baseline,
  including wrapping, horizontal alignment and page continuation in Final and unscaled Word-compatible
  Review. Paragraph shading leaves the first line's extra leading unpainted and
  fills continuation slots. Note-link fallbacks remain.
  Ordinary text binds its horizontal origin after column/page admission; measured
  minimum baselines follow the admitted lines across column and page transitions.
  Plain text and supported prose with inline equations reflow when explicit column
  widths change. Inline pictures skip continuation columns too narrow to fit them.
  Plain-text widow control can move the final pair while preserving a nonempty head.
  Complex mixed-object flow retains its qualified approximations. Review-mode
  vertical placement on custom paper remains approximate.
  Explicit before/after point spacing preserves these baselines; connected
  paragraph backgrounds retain their measured continuity, including different colors.
  Note-numbering, kerning and reduced-print alignment/spacing retain their separate limits.
  Comment ranges and nonpainting comment references preserve the same text-frame
  classification. Actual note markers retain their own minimum-slot rules;
  printed Review balloons still own their scale and placement.
  Hyperlink metadata also preserves qualified paragraph spacing, minimum cell
  slots and complete endnote-block boundaries; the measured run formatting remains.
  Supported numeric field results share that spacing rule. NUMPAGES in normal
  notes is measured with the finalized count through the bounded count retry.
  Splitting a complex field instruction across unchanged-format runs preserves
  its complete opcode and cached result, including PAGEREF. This does not add
  evaluation of cached-only field kinds.
  Supported note equations coexist with PAGE and NUMPAGES in paragraphs and cells.
  Fresh footnotes measure counters with their rendered reference page before
  slicing; unresolved references retain their existing fallback.
  Endnote PAGE is measured with its rendered reference page for existing plain
  paragraph flow, including section-end notes and continuation, and complete mixed
  groups at section or document end that
  pass the source/frame/fit proof above the actual selected footer. Counter
  ownership is independent of optional static-story typography corrections;
  unresolved references and unfit note owners retain their fallback.
  Complete fitting mixed groups also retain their equations
  and counters alongside complete page-contained automatic header/footer frames,
  including complete simple tables, fully prepared equations, supported inline pictures
  and PAGE/NUMPAGES static fields with actual page inputs;
  with actual selected edges constraining available note space at unit scale.
  Unsupported, clipped or unmeasured static stories keep placement fallback.
  Unrendered references and other unqualified
  owners retain their fallback. Mixed/unfit table rows and other
  unqualified note PAGE values remain approximate; full field parity is not promised.
  Complete fitting section-end groups share separator and paragraph-boundary
  spacing with document-end groups; continuation and unqualified frames keep their fallback.
  Highlighting and run shading retain these text metrics while painting their
  backgrounds; font-changing and unsupported layout combinations keep fallbacks.
  Automatic text and equation colors adapt to solid paragraph/cell/run shading;
  explicit foregrounds remain authored and highlights do not change the automatic
  choice. Paragraph backgrounds use placed line, equation and picture frames in
  body flow, cells and headers/footers, including wrapping, indents and page
  fragments. Empty and picture-only paragraphs retain their backgrounds. Pattern,
  theme-fill, paragraph-border and complex review-frame interactions remain
  approximate; application-view calibration and remaining geometry gates are open.
  Simple single-line body minimum spacing uses measured font height across sizes,
  bold/mixed faces, kept paragraphs and neighboring tables/wrapped paragraphs.
  Qualified minimum-line backgrounds follow natural line height while the
  occupied minimum slot and joined source spacing remain unchanged.
  In scaled Review output, qualified plain minimum paragraphs share printed
  advances and source baseline bearings with kept/wrapped flow. Independent Word
  PDF controls cover pagination and paragraph origins. Plain scaled backgrounds
  also use printed natural frames and actual connected leading/spacing, including
  suppressed spacing and different-color boundaries. Strict image differences
  remain. Unresolved static frames, unsupported spacing/contextual styles and
  complex note-marker formatting retain restrictions.
- Review balloons, ranges, revision bands, comment collision/layout and note-marker
  revisions remain partial. Remaining typography/geometry and extraction failures
  are recorded in [document review](PreviewVisualReview.md); successful conversion
  does not turn a failed visual threshold into a pass.
  Measured ordinary-body ranges separate authored selection height from natural
  glyph brackets; balloons use measured font boxes where available. Object,
  point and unmeasured marker geometry retain their approximations.
  Review print fitting never magnifies an already fitting page; its gray lane
  follows the placed balloons at unit scale too. Qualified footer formatting
  balloons share occupied frames and review UI fonts; font-fallback, unmeasured
  and custom-page placement retain approximations.
  Plain comment paragraphs share measured rows without a five-paragraph limit;
  independent Office controls cover two through seven paragraphs and middle-row
  wrapping. Supported fitting/wrapped terminal font pairs now use the same
  count-independent leading rows; names resolving to one face retain paragraph
  boundaries. Existing glyph, style, width and spacing fallbacks remain;
  whole-page parity is unresolved.
- SVG gradients follow Office-visible centered focal geometry and smooth preview
  transparency. Broader SVG composition, transformed picture-link combinations
  and some alpha/color effects retain restrictions and explicit residuals.

Unsupported: full Word balloon/collision pagination, floating charts, SmartArt,
live OLE/macros, unrestricted section/column balancing, arbitrary nested/floating
story reflow, complex scripts/bidirectional shaping and unlisted equation constructs.
Detected unsupported content may be simplified or omitted; diagnostics are not
a complete inventory of every feature in an arbitrary document.

## PDF Output

- PPTX run and chart hyperlinks survive standard PDF font fallback. Missing-font
  typography remains an approximation; diagnostics still identify substitution.

- Deterministic resource naming/object order for stable input, fonts and options;
  optional fixed creation-date metadata. There is no Deterministic option/mode.
- Subset embedded TrueType/CID fonts with ToUnicode, plus standard Type1/WinAnsi
  fallbacks. Subsets retain required compound glyphs; assembled structural symbols
  may use ActualText. CFF/OpenType-CFF outlines are substituted and diagnosed.
- Supported path/fill/stroke/transform/clipping/text operators, raster images,
  bounded gradients/patterns, transparency Form XObjects and soft masks.
- External/internal Link annotations for supported DOCX/PPTX source links.
- Minimal metadata; no interactive AcroForms, other annotation kinds, outlines,
  embedded files, tagged PDF, audio/video or JavaScript. No PDF/A or PDF/UA claim.
  See the source-backed [PDF profile](PdfProfile.md) for emission and validation scope.

Detailed historical calibration and qualification narratives remain in
[the prior capability record](https://github.com/Lokad/OoxPDF/blob/b4ae23938be7df60bb6b1be09985e73fc2e2a523/docs/Capabilities.md).
Condensing this inventory does not remove those rendering rules, evidence or
thresholds, and does not expand supported combinations.


## Static-content fallback matrix

Each row links an unsupported family to its detection diagnostic, fallback, and covering tests (synthetic in-code fixtures). Severities and dedup scope are defined in Diagnostics.

| Family | Detection | Fallback | Covering tests |
|---|---|---|---|
| PPTX stock, surface, and 3D charts | `PPTX_UNSUPPORTED_CHART` Warning | Ignored | `PptxStockChartKindWarnsUnsupportedChart` (pptx-model) |
| PPTX SmartArt | `PPTX_UNSUPPORTED_SMARTART` Warning | Ignored | `PptxUnsupportedFeaturesEmitDiagnostics` (pptx-core) |
| PPTX OLE objects | `PPTX_UNSUPPORTED_OLE_OBJECT` Warning | Ignored | `PptxUnsupportedFeaturesEmitDiagnostics` (pptx-core) |
| PPTX audio, video, animation, transitions | matching `PPTX_UNSUPPORTED_*` Warning | Ignored | `PptxUnsupportedFeaturesEmitDiagnostics` (pptx-core) |
| PPTX missing image parts | `IMAGE_MISSING_PART` Error | Ignored, conversion continues | `PptxMissingImagePartDiagnosesErrorAndContinues` (pptx-images) |
| DOCX missing picture assets | `IMAGE_MISSING_PART` Error | Retain supported picture frame, omit pixels | DocxImageRecoveryTests (rendering-portable) |
| Undecodable images | `IMAGE_UNSUPPORTED_FORMAT` Error | Ignored | DocxImagesTests, ImagingTests, PptxImagesTests |
| Cropped-image decode | `IMAGE_CROP_UNSUPPORTED_FORMAT` Warning | PDF clipping | PptxImagesTests |
| SVG pictures | `SVG_UNSUPPORTED_CONTENT` Error/Warning | Ignored / Partial / RasterImage | PPTX SVG diagnostics; DOCX native/preview fallback and viewport tests |
| DOCX floating charts | `DOCX_UNSUPPORTED_CHART` Warning | Ignored | `DocxUnsupportedChartInHeaderIsDiagnosed` (docx-core) |
| DOCX SmartArt | `DOCX_UNSUPPORTED_SMARTART` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| Supported DOCX body equations | `DOCX_APPROXIMATED_EQUATION` Warning | Approximated | `DocxNativeStructuredSequencesRetainSourceOrderAndIndependentRules` (docx-core) |
| Unsupported DOCX equation constructs/placements | `DOCX_UNSUPPORTED_EQUATION` Warning | Ignored | `DocxUnsupportedEquationInFooterIsDiagnosed` (docx-core) |
| DOCX OLE objects | `DOCX_UNSUPPORTED_OLE_OBJECT` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| DOCX external images | `DOCX_UNSUPPORTED_EXTERNAL_IMAGE` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| DOCX macros | `DOCX_UNSUPPORTED_MACRO` Warning | Ignored | `DocxUnsupportedFeaturesEmitDiagnostics` (docx-core) |
| Strict OOXML | `OOXML_STRICT_DIALECT` Warning | Partial common DOCX support; Strict PPTX unsupported | OoxmlStrictDocxTests and dialect probes (ooxml) |
| Font packs | `FONT_PACK_*` throw codes | Conversion fails with DiagnosticId | FontTests pack suites (fonts) |
