# 0.3.0 development qualification

Updated 2026-11-12. **0.3.0 is not published and is not a completed release
candidate.** The published [0.2.0 qualification](PreviewQualification.md) remains
separate. Development targets .NET 10 and .NET 11.

Manual pie factor offsets now reference the same automatic frame in both axes,
including measured title and legend space. Six independent cached controls reduce
Y error from 30.186pt to 0.003pt at the unchanged 1pt gate; six formerly failing
image comparisons now pass. Both side cases still fail, with Y error reduced from
43.416pt to 15.12pt. All 48 legend names and 536 measured glyphs remain complete;
X, radius, text and other paint stay unchanged. Four compound-label controls retain
832 uncut glyphs without measured label collisions. Two final paints were viewed.
Windows and Ubuntu each pass 67 focused checks, joining two test-only repairs
without replaying the other 65 checks. The parent fails the measured-Y regression;
all original failures are retained. Seven neighboring PDFs, including four portfolio
outcomes, remain exact. An extreme out-of-frame Y still omits six legend names with
an existing overflow warning in both parent and candidate; this is an explicit
unfit fallback, not content-completeness evidence. No new Office action or view
calibration occurred. The broader gate for this Y correction is pending.

Automatic polar labels also place complete measured text/key/background/link units
in separate rows when needed, retaining authored/unfit fallbacks. The preceding
admission source passed Windows/Linux CI, both frameworks and four consumers of
the same package. The selected 6E96 development package belongs to that source,
before the Y correction. Publication and broader acceptance remain deferred.
A fixed .19-width wrapping prediction matched only 10 of 18 cached Office label
groups and was rejected as a universal rule; wrapping remains approximate.

Manual 2D-pie radii now follow the smaller authored viewport dimension through
shared fitting. Eight pinned cached Office controls improve from up to 16.633pt
radius error to at most 0.004pt at the unchanged 1pt gate. All 48 legend names and
536 measured glyphs remain complete and unclipped; centers, text and colors stay
exact. Both hosts pass 23 focused checks, joining one explicit-label fixture repair
without replaying the other 22. The parent fails the real radius assertion. Seven
neighbor PDFs, including four portfolio outcomes, remain exact. Before the factor-offset correction above, all eight image
gates still failed: six MAEs improved and two worsened with unresolved centers.
Manual height/label/explosion Office geometry, edge-Y/side/legend anchors and application
view remain unqualified. The prior factor-X correction and bounded-offset repair
remain covered. Broader framework/package results still belong to their exact
artifacts; the radius batch passed Windows/Linux CI, while the selected package
precedes both recent local slices. Publication is deferred.

DOCX keep-with-next groups now admit once per source group (`7c94a3470`). Two
oversized plain/mixed chains and one exact-fitting keep-lines control previously
created an unnecessary third page; all three now use two pages and match their
independent free-flow control PDFs exactly. All 120 source labels, 272 measured
glyphs and eight picture paints remain present, with no measured cuts or unknown
text traces. Independent keep-lines and explicit page/column/section barriers stay
covered. Windows passes 22 focused checks; Linux passes 19 with three declared
Windows-font skips. Two test-only repairs preserve their original failures and
reuse the runtime; only repaired methods reran. All fourteen DOCX portfolio PDFs
and diagnostics remain exact, preserving their existing visual failures. Existing
Word-backed spacing/page controls pass; the new inputs have no calibrated Word
view/geometry reference. Framework and package results belong to each immutable
artifact's qualification record; local rendering checks alone do not close those
gates. Publication and broader acceptance remain open.

A separate paint-order review of the sixteen retained portfolio outcomes binds
56,513 glyphs to their actual font programs. No later paint among the 218 opaque
rectangles and 22 axis-aligned opaque pictures examined intersects those glyph ink
bounds. Eleven other
fills also have bounds disjoint from earlier glyph ink. Five controlled PDFs check
erasure, paint order, transparency, clipping and opaque image masks, including
independent final-pixel checks. This bounded
screen adds no Office-view or arbitrary occlusion claim: strokes, shading, text
over text and annotation appearances remain unqualified. Existing visual failures
and thresholds are unchanged.

Ordinary SVG labels now paint through the common DOCX/PPTX image path
(`06bbe1379`, font-owner integration `fa2184024`). The two synthetic inputs that
previously painted only a blue rectangle now retain all 13 non-space glyph paints
per label, inside their diagram viewports, without SVG warnings. Both strict CLI
conversions succeed and match their default PDFs. Independent Arial source-program
glyph bounds/advances match all 26 path paints within 0.000382pt; the labels were
also visually reviewed. These are outlined picture labels, not embedded PDF text:
source GIDs are known, but no PDF font/GID metadata or pointwise contour proof is
claimed. No new Office geometry or application-view qualification is implied.

Harmless class/id attributes and descriptive metadata also preserve the labels
(`b50b8eaab`). An initially over-broad guard dropped text when a class or title was
added; three omissions now match the exact visible control PDF. Actual stylesheets
remain distinct and explicitly unsupported. Metadata never becomes visible text
or shapes. The original parent fails the missing-label assertion, and the preceding
SVG runtime fails the metadata regression.

Windows and Ubuntu each pass 29 exact checks with no skips in one invocation per
host on the final production runtime. Earlier assertion failures and repairs remain
recorded separately: platform line endings, numeric matrix isolation and an incorrect
expected glyph-Y sign. Compound-list Office ordering stays unqualified.
A composed report in Final/Review and a held-out table deck retain 39 new label
glyphs while all pre-existing glyph traces, positions, page counts and diagram frames
match their parent outputs. Source-font bounds match within 0.000615pt and final-paint
crops show complete labels. All sixteen original portfolio PDFs and diagnostics stay
exact, retaining independent evidence and two PPTX passes/fourteen DOCX image failures.
Transfer verifies 24 assets and two exact smokes without a root build or suite replay.
Preceding CI43235373217/sourcec43b0a2c0 is closed green; its intermediate payload was
superseded by the metadata repair. Current framework/package closure belongs
to the selected artifact's qualification record. Publication and broader acceptance stay open.

Automatic note references and pictures now retain source order when their runs
are joined (`df7000b47`). Four previously failing Final/Review outputs match their
unchanged separate-run controls; the parent fails the glyph-geometry regression.
Windows passes 21 focused checks; Ubuntu passes 18 with three declared font/platform
skips. Native programs bind 14,016 glyphs without measured cuts or unknowns.
CI43195810829/source2127c65f6 is closed green: Windows4,825/19skips,
Ubuntu2,317/133skips,zero failures; each host also passes69 net10 checks.

Distinct note references with equal labels also retain their own links within one
run (`9312cddfe`). Two outputs lose the erroneous overlapping links while retaining
exact page paint streams, native font programs and pictures; six controls stay
byte-identical. All four joined/separate pairs match. The parent fails the link
regression. Windows and Ubuntu each pass 20 focused checks with no skips after a
test-only fixture repair; only the repaired fixture and three neighbors reran.
Native programs bind 13,802 glyphs without measured cuts/unknowns. CI43201992628/
source98b8de174 is closed green: Windows 4,826/19 skips, Ubuntu 2,324/133 skips, zero
failures and 69 net10 checks per host. The preceding failed CI remains recorded;
its fixture now authors formatting before parsing instead of invalidating parsed
note owners. The original geometry thresholds remain unchanged.

Cell-local padding now keeps the containing table grid instead of creating an
inherited text anchor (`505ac9c89`). Body and nested tables share the existing grid
metadata; explicit zero/nonzero indentation, table/row/style margins and legacy
controls retain their earlier rules. Eight focused horizontal failures become
zero at the unchanged 1pt Word threshold (maximum residual 0.264pt). The strengthened
test covers all 32 independent container outcomes, including previously excluded
body/nested origins; the parent fails its actual absolute-origin assertion.
Windows passes 26 focused checks; Ubuntu passes 15 with 11 reviewed Windows-font
skips. The initial incomplete skip declaration remains recorded without rerunning
tests. Ten outputs retain every native character/GID and picture; 1,060 bound
glyphs have no measured cuts or unknowns. Four protected control PDFs and all
sixteen original portfolio PDFs/diagnostics stay exact. Two PPTX passes/fourteen
DOCX image failures remain. A shared-scale before/after/Word crop was reviewed;
new Office actions remain prohibited and application-view calibration unfinished.
Historical conflicting Word anchors remain unresolved. Combined CI43206178920/
source358d5c907
is now closed green: Windows 4,826/19 skips, Ubuntu 2,325/133 skips, zero failures.
Each host passes 69 net10 checks plus 5 Windows/4 Ubuntu exact supplements. Both
Linux frameworks consume the same Windows package. Publication, release acceptance
and the broader roadmap stay open.

Ordered body paragraphs share widow/orphan fragment selection with ordinary text
and continued notes (`3abcf3afc`, bounded preview `5c3decdf7`). A small inline
picture had left the first line of a three-line paragraph on the preceding page
despite default widow control. The shared rule retains default/on/off behavior,
single-line and longer tails, text/picture/math content, PAGE rebinding and bookmark
destinations. The actual B7 parent fails the new row-page assertion.

Admission previews the fitting fragment and at most two following lines. This
removes an introduced per-frame rewrap of the whole remaining paragraph. One cold
257-page check falls from 4.69s to 1.77s, including startup; all three workload PDFs
remain exact. This is a bounded regression check, not a general throughput claim.
An initial timing overlapped tests and is retained separately from the serial run.

Windows passes 21 focused methods; Ubuntu passes 16 with five declared Windows Arial
skips. The matrix covers three/four/five-line heads and tails. An earlier new
assertion assumed whitespace separators between PDF text operations; its failed
attempts and single-method repairs remain recorded. The changed production runtime
runs the complete exact set once per host. Forty PDFs stay byte-identical to the
qualified widow outcomes, including all sixteen portfolio PDFs. Native/source-page
checks, reviewed crops and independent Office comparisons are reused by identity;
two PPTX passes/fourteen DOCX image failures remain. Transfer verifies 24 assets and
two exact smoke PDFs without a root build or suite replay.

The preceding widow gate CI43211571221/source58bb78292 is closed green: Windows
4,828/19 skips, Ubuntu 2,333/133 skips, zero failures and 69 net10 checks per host.
Its intermediate package was not downloaded while this repair was active. Current
runtime/framework/package closure is recorded separately beside the selected
payload. New source/Word pagination and application-view calibration remain
unqualified; publication and release acceptance stay open. No Office action.

Computed PAGE/NUMPAGES results now remain atomic beside pictures in the same source
run (`4da834039`). Four failing Final/Review outputs become byte-identical to their
unchanged separate-run controls. Windows passes 30 focused checks; Ubuntu passes
26 with four reviewed font/platform skips. The parent fails the new picture-placement
regression. CI43179331533 is closed green on both hosts; its Windows main suite
passes 4,822 checks with 19 skips.

Fitting Review endnotes now bind counters independently of print scale
(`b7c26a7fc`, harness `1656a5db0`). Two existing fractions reappear in each of two
composed outputs; the other ten footer-cell/endnote/continued-footnote outputs
remain exact. All six joined/separate pairs are identical. Actual native programs
bind 30,887 glyphs with no measured cuts or unknowns. Windows passes 23 focused
checks; Ubuntu passes 22 with one declared platform skip. One repaired assertion
reran on Windows, reusing 21 runtime assets and 3,438 public inputs. The parent
fails on the missing fractions. New mixed-reference PAGE-value and Word-geometry
parity remain unqualified. All sixteen original portfolio PDFs/diagnostics stay
exact, preserving two PPTX passes/fourteen DOCX visual failures. Transfer verifies
24 assets and two exact smoke PDFs without a root build or suite replay. The
coherent follow-up gate and package identity are recorded separately; application
view, release acceptance and the broader roadmap remain open.

Computed simple fields now keep a distinct source position (`4a93dce6d`, harness
`68ee9a07e`). PAGE/NUMPAGES-only hyperlinks retain their click area, and following
text remains outside the link. Both new regressions fail on the parent. Windows
passes 42 focused checks; Ubuntu passes 41 with one reviewed Cambria Math
availability skip. The initial strict skip gate is retained without replay.
Four causal outcomes restore two annotations. The original composed report now
covers its page digit as the pinned Word PDF does in both modes. Native glyph
programs/IDs, text positions, pictures and pagination remain unchanged; 21,531
freshly bound glyph instances have no measured cuts or unknowns. Fourteen portfolio
PDFs are exact; the two changed PDFs retain identical visual metrics and thresholds
(two PPTX passes/fourteen DOCX failures overall). One transformed Review connector
moves 1.051 points to match the literal-result control; its Word geometry remains
unqualified. Transfer verifies 24 assets and two exact smoke PDFs without a root
build. CI43160881670 on `61b437941` is green: Windows 4,820 checks/19 skips,
Ubuntu 2,308/130, 69 .NET 10 checks per host and four CI package consumers. Both
new methods also pass on each host with the exact CI .NET 10 assembly and two
owned public inputs. Two CI-runtime smoke PDFs match. Both Linux frameworks
consume the same Windows package; the prepared payload below contains this fix.
No new Office action was taken. Release acceptance remains open.

Read-only triage of three existing body/nested-table roots retains both labels
and the complete fraction in all six Final/Review outputs. Actual font programs
bind 695 glyphs with no measured cuts or unknowns; inspected cell crops remain
readable. The horizontal placement failures, including the 7.480-point residual
and conflicting root-table Word anchors, remain unresolved. No placement rule,
threshold or production code changed for this triage.

The terminal-bookmark repair (`a180e12f7`, reviewed harness `0a62f09f3`) restores
two internal annotations lost when an explicit page break leaves the bookmark
after unpainted trailing whitespace. Destination resolution keeps existing points
first, then uses the preceding painted text endpoint. It preserves the source
boundary before the break; transformed-source Word viewport geometry is unqualified.
Windows and Ubuntu each pass 27 focused checks without skips. One new whitespace
assertion was corrected after failing on both parent and candidate; original
exact endpoint checks and Office thresholds remain. Only that method reran on
Windows, reusing 21 runtime assets/3,438 inputs. The parent still fails the composed
regression. Four causal outputs retain identical native font-program/GID paints,
text geometry, paths, picture pixels/alpha and existing links; 14,370 bound glyphs
have no measured cuts or unknowns. All sixteen portfolio PDFs/diagnostics are
exact, retaining two PPTX passes/fourteen DOCX image failures. Transfer verifies
24 assets/two exact PDFs without rebuilding or replaying a suite. Frozen .NET 11
library: `E1F4DA0726E7E527EA8C4C7BEB0788F81FF63A99E31079099616E10DAB67C0F1`.
The coherent gate CI43153375365 on `9304d7aa2` is green: Windows4,818 checks/
19 skips, Ubuntu2,301/130, 69 .NET10 checks per host and four CI consumers.
Four additional affected .NET10 checks pass per host on the exact CI assembly.
Both CI-runtime smokes match, and both Linux frameworks consume the same Windows
package. The prepared package below contains this repair. Calibration, release
acceptance and the broader roadmap remain open; publication remains deferred.

The retained-source ownership repair (`c17da72f8`, runtime from `b3fa46305`) keeps
an unsupported revised equation from suppressing an independently supported
neighbor. Four public body/cell/header/footnote roots recover ten fraction paints
and 56 native mathematical glyphs. The added accent remains diagnosed and omitted;
copied owners, missing source positions and empty-only literals retain fallback.
Source ownership and control discovery share an iterative retained-source walk,
while paint eligibility remains separate.

Windows and Ubuntu each pass 29 unique focused checks without skips. A repaired
test distinguishes authored revision bars from unchanged supported paint; its
single-method rerun joins 28 earlier passes using 21 identical runtime assets and
3,438 public inputs. The discarded bar-policy trial and failed attempts remain
recorded; the existing rendering policy is preserved. Native checks bind 3,056
glyph instances across 24 outcomes with no measured cuts or unknowns. Twenty-one
control PDFs are exact; three Review outputs retain existing revision bands and,
in the footnote case, first-seen author-palette changes. Actual font-program/GID
sequences, text geometry and image paints match controls. The parent fails on
missing mathematical text. The related-story extension reuses the frozen runtime.
All sixteen portfolio PDFs/diagnostics remain exact, preserving the two PPTX
passes/fourteen DOCX image failures. Transfer verifies 24 assets/two exact smoke
PDFs without rebuilding or replaying a suite. The coherent gate CI43136688558 on
`81eae757b` passes: Windows 4,816 checks/19 skips, Ubuntu 2,296/130, 69 .NET 10
checks per host and four package consumers. Four additional ownership/control/
cancellation checks pass on each host using the exact CI .NET 10 assembly. Both
CI-runtime smoke PDFs match, and both Linux frameworks consume the same Windows
package successfully. The prepared package below contains this repair; Office
appearance, calibration and release acceptance remain open.

A recolored existing wash overlaps eight denominator font boxes. A selected exact
fill knockout recovers no dark pixels at 300 DPI, and the inspected words remain
readable. This does not demonstrate a visible defect or establish arbitrary ink
occlusion safety; no painting order, geometry or threshold was changed.

The coherent math-run repair (`f8bde8a1d`, including `4ddb7e1da`) shares readable
run and nonprinting-marker rules between revision admission and leaf parsing.
A stored pagination marker previously removed a supported formula; an empty run
beside actual revised text could separately remove its fraction. Four public
roots now recover six formula paints and eight additional fraction paints
(26 mathematical glyphs). All 26 control/negative outcomes retain exact PDFs and
diagnostics, including four empty-only revision owners that remain unsupported.
Native checks bind 3,652 glyph instances with zero measured clip cuts or unknowns.
Content-bearing, foreign or misplaced markers and unknown empty-run structures
retain their fallbacks. The marker's stored-pagination role and mathematical-run
parent are documented in the
[Open XML reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.lastrenderedpagebreak?view=openxml-3.0.1).

Windows and Ubuntu each pass 21 focused checks without skips in one invocation
on the final frozen .NET 11 library
`9978749BC9396F8E2A7F11EE775A456AFFF34B6A86DC9045B17C559485F428E5`.
The parent fails the actual mathematical-content assertion. An initial proposal
incorrectly admitted empty-only revision owners; the existing ownership guard
caught this, and its failed attempts remain recorded. The corrected shared rule
requires actual revision content while permitting readable empty siblings.
Transfer verifies 24 assets/two exact PDFs, including the repaired input and
six-name legend, without a root build or suite replay. All sixteen portfolio
PDFs/diagnostics remain exact; two PPTX passes/fourteen DOCX visual failures remain.
The coherent milestone passed CI43130020058 on `55c5b06e3`: Windows 4,814 passes/
19 skips, Ubuntu 2,295/130, 69 .NET 10 checks per host and four package consumers.
Terminal logs are retained once. Four additional metadata checks pass per host on
the exact CI .NET 10 assembly; both CI-runtime smoke PDFs remain exact. The same
Windows package also passes separate Linux consumers under both frameworks.
The prepared package below contains both repairs. Transformed-source Word
appearance, application-view calibration and the broader roadmap remain open.

The revised-inline-box owner repair (`88ca67f52`) restores fifteen formula paints
across four public roots/24 Final/Review outcomes in body, cells and headers.
Visible drawing insertions/deletions now use the existing revision ownership set;
invalid owners and frames retain their fallbacks. Native checks bind 2,027 glyphs
without measured clip cuts, unknown traces or picture/ink intersections. Eight
original PDFs and four unmasked Final raster pairs remain exact; eight Review
differences are confined to existing revision-margin bands. Transformed-source
Word geometry and application view remain unqualified.

The adjacent diagnostic repair (`3b4683d12`) uses the existing selected revision
and row rules in equation probes and expected-paint counts. Intentionally hidden
equations no longer produce unsupported-equation warnings, including when an
entire note body is excluded before layout. Final's supported deleted-math
structure and warnings for visible malformed or unsupported payloads remain.
Two hidden Final cases now pass strict CLI conversion with no diagnostics.

Twelve additional revision-setting/view outcomes pass strict mode and match the
existing Final PDFs exactly. Six source-preserving footer placement outcomes
retain all five visible formulas and the one intentional omission; actual native
font-program/GID sequences match controls, and 408 measured glyphs clear clips.
Separate screens find no cross-baseline or same-baseline cross-paint ink-box
intersections among the portfolio's 56,513 bound glyphs. Within-paint same-baseline
collisions, rotated text and arbitrary occlusion remain outside those screens;
new-source Office appearance remains unqualified.

Windows passes 64 focused checks; Ubuntu passes 56 with eight declared font skips.
The final Windows selection runs one causal check and 63 neighbors without
duplicating methods; Ubuntu runs one selection. The parent fails the restored
hidden-warning assertion. Earlier production attempts exposed retained deleted
math and absent note stories; their failures remain recorded. Two intermediate
harness repairs reused the runtime. Final frozen .NET 11 library:
`28314ACA5325238004F9FE3E002311CC5A96D77D9B85091BE35178E9E4E338C4`.
Transfer verifies 24 assets/two exact smoke PDFs without a root build or suite
replay. All 24 causal PDFs remain byte-identical to the content repair, with only
two false warnings removed. All sixteen portfolio PDFs/diagnostics remain exact;
independent paint evidence and two PPTX passes/fourteen DOCX image failures remain.
The coherent milestone passed CI43118101912 on `a228b65dd`: Windows 4,810 passes/
19 skips, Ubuntu 2,293/130, 69 .NET 10 checks per host and four package consumers.
Terminal logs are retained once. Four additional .NET 10 revision/equation checks
pass on each host using the exact CI assembly; two CI-runtime smoke PDFs match.
The prepared package below contains the earlier box-ownership and selected-view
diagnostic repairs.
Publication, application-view calibration and the broader roadmap remain open.

The independent comment-body follow-up finds no new defect in the report, letter
and specification. Their three complete Review sentences appear on the same
pinned Word pages, while Final intentionally omits them. All 264 selected glyphs
across both producers match actual embedded cmaps and clear known clips; six crop
pairs are reviewed. This covers simple literal one-paragraph bodies with shared
source ancestry, without a general comment or application-view claim.

The nested-inline-box fallback repair (`0bf14111c`) keeps a valid box's own frame
independent of its unsupported children. Textbox stories measure supported source
items through the shared ordered path, skipping nested frames at their original
indices. Previously, a nested box could erase formulas beside the outer box or
inside it in body, cell and header content. Four public sources/control transforms
now retain ten previously missing equation paints across sixteen Final/Review
outcomes. Actual font-program/GID checks measure 1,414 glyphs without clip cuts,
unknown traces or picture/ink intersections. All eight original control PDFs and
eight transformed/control page-raster pairs are exact; this also removes a 2.547pt
Review child-baseline difference. Nested content remains unsupported and receives
the precise `DOCX_UNSUPPORTED_NESTED_INLINE_TEXTBOX` warning.

Windows passes 46 focused checks; Ubuntu passes 39 with seven declared font skips.
Test-only repairs add the bookmark needed to activate lazy source-index capture
and update an older staged-path assertion to retain the valid outer frame while
ignoring nested paint. Explicit-slot/revision guards remain; empty paragraph
struts remain valid. Only repaired methods were rerun per host; six previously
omitted neighboring methods join the capability. Reports preserve original failed
attempts and verify 21 unchanged runtime assets/3,438 inputs. An earlier
harness repair bound math to its pinned mathematical font; the parent regression
then fails on the missing first formula. Intermediate production results and the
initial test-compile failure remain recorded. Final frozen .NET 11 library:
`AE5B7A3FD66AA41D329FF4EF7257A620C54EBED5F38753F98492456EB5306A31`.
Transfer verifies 24 assets and two exact smoke PDFs without a root build or suite
replay. All sixteen portfolio PDFs/diagnostics are unchanged, retaining the two
PPTX passes/fourteen DOCX image failures. Additional split-run/bookmark/body-link
checks in cells and headers retain all eight formulas/links, 584 bound glyphs and
twelve exact raster pairs. Eight two-page outcomes retain 704 glyphs/twelve
formulas: the moved cell formula and link target follow page two, while header
math repeats on both pages. All eight page-raster pairs are exact. These extensions
need no new library build or rendering fix. CI43099685261's earlier Windows
staged-path assertion failure remains recorded. Replacement CI43101429109 on
`f5f0a7999` is green: Windows 4,805 passes/19 skips, Ubuntu 2,292/130, 69 .NET 10
checks per host, four package consumers and ten inspections per host. Terminal
logs are retained once. Four additional .NET 10 box fallback/ownership/source-index
checks pass on each host using the CI assembly. The exact CI package below now
contains this repair. New-source Word geometry/application view remain
unqualified. No Office action, publication or roadmap completion is claimed.

The fitting-note frame repair (`5551aeff6`) applies the existing printed-story map
to complete fitting normal-story table subtrees and their drawing URI targets.
Review previously painted a footnote table below the following note and let an
endnote calculation overlap following note text. Four causal sources/eight
outcomes now retain source order, uniform prose/math/picture scale and existing
internal destinations. Independent checks bind 2,992 glyphs to their actual font
programs/GIDs: no measured clip cuts, unknown traces or picture/ink intersections.
All four Final PDFs remain exact. Exact/unfit/fragmented/merged/unresolved table
owners retain their staged fallbacks; transformed-source Word geometry remains
unqualified. No Office action or application-view claim.

The new portable regression fails on the parent. Windows passes 38 focused checks;
Ubuntu passes 31 with seven font skips (six predeclared and one indirect guard
reviewed after the original strict skip failure). Four fitting/unfit long-note
boundary PDFs/diagnostics and all sixteen portfolio outputs remain exact; the two
PPTX passes/fourteen DOCX image failures are unchanged. Corrected source XML order
reproduces all prototype parent outputs; retained attempts document the repair.
Two local production builds include an unqualified trial before the physical-page
fit guard. Source/native evidence uses the final frozen .NET 11 library:
`D227C261BDD418F9EAA24DE8A965921B18B818E1567A705A7140416A173FCC5F`.
Transfer verifies 24 assets and two exact smoke PDFs without a root build or suite
replay. CI43081400120 on `8a4d30380` is green: Windows 4,802 passes/19 skips,
Ubuntu 2,291/130, 69 .NET 10 checks per host, four package consumers and ten package
inspections per host. Two additional .NET 10 note/footer regressions pass on each
host using the CI assembly. Their first attempts stopped at missing public inputs;
five hash-pinned fixtures repaired that binding without an assembly rebuild.
Original setup failures remain recorded. The prepared package contains this
repair; publication and the broader roadmap remain open.

Further fixed-runtime composition checks retain numbered note captions, resolved
NUMPAGES fields, both formulas and complete opaque pictures. Sixteen list/field/
keep outcomes measure 5,800 bound glyphs without cuts, unknown traces or picture
intersections. Eight literal/field page-raster pairs are exact. Four additional
body-link outcomes target the correct split/complete note-math page, retain the
sixteen original links and reproduce twelve source-page rasters exactly. No new
renderer defect or fix is claimed; exact Word pagination remains unqualified.

Seven existing footnote admission/continuation methods now share the pagination
partial. Their bodies, shared helpers, assertions and 4,821-method catalogue are
unchanged. Nine affected/neighbor checks pass on each host under .NET 11. One
harness-only build reuses all 21 runtime assets/3,438 public inputs; transfer
verifies 24 assets and two exact smoke PDFs. Production, the portfolio, the prepared
package and the closed CI qualification remain unchanged; no broader replay.

The aligned-cell follow-up (`f5509b8ab`) keeps the shared printed frame for complete,
measured centered/bottom-aligned cells. Four Review cases previously let pictures
cover six caption letters; native picture-knockout checks reveal 512–514 hidden
dark pixels per case. All eight Final/Review outcomes now retain complete glyphs
and pictures without measured intersections. The regression fails on the parent;
91 Windows checks pass, and Ubuntu passes 74 with 17 declared environment skips.
All sixteen portfolio PDFs/diagnostics and four existing PNG/SVG alignment controls
remain exact. Nested and legacy-only alignment keep their qualified fallbacks.
New-source Word geometry/application view remain unqualified; no Office action.
The frozen .NET 11 library SHA-256 is
`303EC4FF9B7E47E0EACBA376152B0C7AEC27E3C7DC65ED8001F54A0BFDA576F2`.
CI43062947587 on `e7d591999` is green: Windows 4,800 passes/19 skips, Ubuntu
2,289/130, 69 .NET 10 checks per host, four package consumers and ten package
inspections per host. That package is retained as parent evidence.

Six footer-cell outcomes retain 21,483 bound glyphs, all 24 source PNG paints,
repeated captions and the correct PAGE/NUMPAGES values without measured cuts or
picture/ink intersections. Body/note links remain unchanged; these sources do not
author footer hyperlinks. The new portable check passes both hosts and the parent,
so no additional footer defect is claimed. One harness-only build reuses the same
library; two test API compile mistakes and their failed attempts are retained.
Joined focused coverage is 92 Windows passes and 75 Ubuntu passes/17 declared
skips, without replaying the preceding 91 checks. Those focused checks followed
CI43062947587; the current broader CI above also includes the footer regression.

The preceding rendering slice (`044032aaf`) publishes the existing validated Review
table frame for complete rows on every page. A long table previously let pictures
cover caption glyph bottoms although text-count and clipping checks passed. Native
ink/raster checks now retain 216 body-row paints, twelve mixed headers and all
17,992 measured glyphs; eight causal outcomes have no picture/text intersections.
The new portable regression fails on the parent. Focused .NET 11 checks pass 95 on
Windows; Ubuntu passes 82 with 13 font skips (two indirect guards reviewed and
confirmed after the initial strict skip gate). Sixteen portfolio PDFs/diagnostics
and the authored six-name legend remain byte-identical; their two PPTX passes and
fourteen DOCX image failures are retained. No new Office/view parity is claimed.
The frozen library SHA-256 is
`DE83321FD2123FCB232056F935194EF84832D24C9C60443083510447D13867D2`.
Broader CI43041456284 on `03ef23fcc` is green: Windows 4,799 passes/19 skips,
Ubuntu 2,288/130, 69 .NET 10 checks per host, four package consumers and ten package
preflight checks per host. Production/test source matches the frozen slice;
CI builds have their own revision metadata. This does not complete the roadmap.
The package refresh below uses the current CI production payloads; the previous
prepared archive is retained as parent evidence. Publication stays deferred.

Additional kept-cell checks move complete prose/formula/picture rows to a fresh
page in both modes, with the following paragraph intact. Eight outcomes retain
29,110 glyphs without clip cuts/unknown traces and complete opaque pictures;
the portable regression passes on both hosts and fails on the parent. This extends
the existing fix's evidence without another library change or Office export.

The previous local rendering baseline pins production
`c1b59358b13c7c6b6492f8ab7cd20a7881bc1d2a`; harness
`7196e18eb13656c6b6f233cfd98db470b0818786` repairs only a test fixture.
The frozen .NET 11 library SHA-256 is
`840C42E925641050D35B2F697F0450E0DFA460577B94339F9A373F1312056C08`.
Documentation-only updates reuse that frozen library.

## Verified outcomes

| Gate | Result and scope |
|---|---|
| Selected comment-view diagnostics | Final/Original and document settings that hide comments no longer report intentional omission as unsupported. Visible SimpleMarkup/AllMarkup comments retain their approximation warning and related-part provenance. The CLI still publishes PDFs and returns 3 for genuine warnings. Eight causal PDFs and all sixteen portfolio PDFs are byte-identical; four causal/seven portfolio legacy warnings are removed. Both hosts pass 60 of 61 exact checks with one declared case-insensitive-filesystem skip. A single font-dependent CLI fixture failed on Linux; a test-only repair isolates strict exits with a textless anchor, while reader/rendering tests retain prose/body coverage. Only that method was rerun per host; prior reports remain retained and joined after runtime/input identity checks. Twenty-four transferred assets and two smoke PDFs match. No root build, rerasterization or Office action. Broader CI43015189979 is green; release scope remains open. |
| Words after fixed-line picture overhangs (preceding batch) | Four oversized-picture outcomes hid four/five letters of After. Captured ordinary content now retains legacy width-fitted picture paint inside the shared source cursor, allowing the following word to wrap. All eight causal outcomes retain 40 glyphs without measured cuts or unknowns; picture bounds, visible fractions and source colors remain unchanged. Four terminal-overhang PDFs remain exact. Both static paths share the original baseline clip, retaining the 22pt-line/17.6pt-clip assertion. Portable coverage includes two fonts, body/cells, both presets, source-equivalent split runs, whitespace, bookmarks and cancellation. Exact new-source Word geometry is unqualified. |
| Contained fixed-line pictures (preceding batch) | Four fitting exact16 picture/prose outcomes hid four letters of After despite complete PDF operations. Captured ordinary rows now reuse shared contained-picture wrapping with the container's physical baseline, preserving fixed advance. All four words and authored picture dimensions are restored; six required outcomes retain 30 glyphs without measured cuts or unknown traces. Both first-frame and oversized fallback PDFs remain identical. The rendered regression fails against the parent. Existing independent exact-frame Word gates passed unchanged; new-source Word geometry remains unqualified. |
| Minimum-spaced pictures (preceding batch) | Changing only automatic spacing to atLeast16/32 left a partial A and hid the other four letters of After in four cell/preset outcomes. The reader now captures ordinary picture source items through the existing publication helper; captured ordinary content shares automatic/minimum measurement. All four words and authored 240x18pt pictures are restored, with all source color bands and 30 glyphs clear of measured clips. Automatic and exact-height control PDFs remain identical. Portable coverage checks two fonts, body/cells, both presets, split runs, declared minimum frames and cancellation. Reconstructed minimum frames keep their qualified fallback; math/box/exact eligibility remains distinct. Two regressions fail against the parent. Exact new-source Word geometry is unqualified. |
| Leading pictures and following words | Two narrow-first-line cell controls clipped four letters of After despite complete source operations. Owned automatic-flow pictures and prose now share the existing ordered cursor; the word and full-size picture move into the continuation frame. All six causal outcomes retain the word, complete picture bounds and source colors; 30 actual-program glyphs have no clip cuts or unknowns, and the unindented PDF stays exact. Picture paint height stays separate from occupied descent. Three pinned Word controls preserve the original 0.1pt dimension/0.75pt baseline gates. Shared width-roundoff handling also removes an unnecessary cell line without changing the original hhea/image-baseline checks. At that milestone, declared, unowned and unfit rows retained their fallbacks; the fitting minimum/fixed rows are qualified separately above. No exact new-source Word/view or arbitrary overpaint claim. |
| Autofit paragraph insets | A 72pt first-line indent hid every glyph of a short label while text operations remained present. Natural cell widths now fund the same paragraph offsets used by placement, including positive right insets and measured inline content. Both missing labels recover the controls' 270 raster ink pixels; four control PDFs are unchanged and 30 actual-program glyphs have no measured cuts/unknown traces. Portable checks cover two fonts/presets, split runs, grid hints, left/first/right insets and cancellation; the regression fails on the parent. No all-content, compressed-width or exact new-source Word/view claim. |
| Terminal pictures | Fitting terminal/trailing-space pictures escaped width admission: body retained 69.98% of bounds and cells 39.78%. Both wrapping paths now share source-boundary admission, including final rows. Adjacent drawings use their combined advance at a shared offset, correcting a trial that retained only 8.49% of the second cell picture. Eight pictures in six native outputs retain complete bounds and all source colors; 636 actual-program glyphs have no measured clip/unknown traces. Portable regressions cover both presets, two pinned fonts, merged/split runs and adjacent pictures; parent/trial failures are retained. Unfit groups keep their existing fallback; arbitrary group splitting and exact new-source Word/view geometry remain unqualified. |
| Pictures after soft breaks | Two/three supported soft breaks caused an opaque inline picture to cover part of the following word, although the whole word remained extractable. Picture attachment now reuses wrapped source starts and per-line source offsets, removing duplicate painted-length reconstruction. Four native controls retain all source words/pages with no following-word overpaint; the unchanged no-break PDF is identical. The screen binds 44 glyphs without clip cuts or unknown traces; paired crops confirm restored visible words. Portable checks cover pinned Sans/Math, both presets and separate/merged runs; the regression fails against the parent. No new-source Word/application-view geometry claim. |
| Pictures inside text runs | Merging equally formatted text/picture/text runs moved PNG/SVG pictures about 35.6pt left and shifted preceding prose behind them by the picture advance. The reader now retains text child boundaries, and shared wrapping/placement uses recorded source run numbers and offsets. Both corrected merged-run PDFs exactly match their original controls. Two portable regressions cover pinned Sans/Math fonts, Final/Review, multiple drawings in one run and cancellation; both fail against the parent. The affected native screen measures 22 glyphs without cuts or unknown traces. No ordinary-layout migration or new layout layer. |
| Ordered picture rotation | Shared placement publishes the measured line frame, restoring six authored quarter-turns lost beside mathematics in body/header/footer content. Eight controls bind 546 glyphs without cuts or unknown text traces. The global transformed-origin change is reverted: five existing source-pinned Word PDF comparisons exposed ordinary PNG/SVG regressions, including a previously passing anchored picture. All five corrected PDFs are byte-identical to their earlier qualified outputs, restoring one pass and retaining four failures. A portable regression checks the Word anchor within the original 0.001pt gate with two pinned fonts and both modes. Source-fit dimensions, source-centered rotation, clips and invalid-frame fallback remain. |
| Actual legend ink (prior milestone) | Two caller-supplied pinned Math-font labels retain their full arrow glyphs beyond nominal font metrics. Clipping uses resolved glyph bounds at the existing PDF emission font size; known empty whitespace no longer aborts ink measurement. Text origins, programs, paths and colors stay exact. The two public-API outputs bind 126 filled glyphs with no legend cuts; one unchanged synthetic-bold title stroke per PDF remains unqualified. Bubble legend alignment survives inserting a nonpainting space. No exact new-source Office or application-view claim. |
| Side-legend ink | Four automatic left/right pie legends emitted all six names but cut 28 glyphs in their bottom labels. Their clips now include measured font ascent/descent within the chart frame; explicit authored viewports retain clipping. Twelve pie controls bind 804 glyphs with no cuts or unknown traces. Four pinned PowerPoint PDFs independently retain the complete labels. Text/key/plot origins, font programs, paths, colors and links remain unchanged. Whole-page pie comparisons remain two passes/ten failures; the four affected MAEs worsen slightly despite restored ink, and those results remain retained. |
| Picture bookmarks | Four ordinary body/cell Final/Review targets now reveal their placed pictures instead of following prose 33pt to the right. Source destinations and URI links share the existing physical picture traversal; ordered measured targets keep their path and URI decode/resource/clip guards remain. Page commands, actual font programs, image streams, native glyph positions and paths stay exact; the original 0.001pt gate remains. Portable target/gap checks cover body/cell/header/footer, wrapped/unwrapped content and nonpainting runs with/without math; the target regression fails on the parent. The prior sixteen static-picture targets remain protected. Exact combined-source Office geometry and unresolved container ownership remain unqualified. |
| Direct header/footer pictures | Ordinary mixed pictures bypassed shared wrapping even when their source could be measured. Four Final controls retained only 54.14% of the picture bounds. Dispatch now uses the shared inline path; all sixteen direct/boxed controls retain full pictures and clear source text. The screen binds 1,880 glyphs to actual programs/resources/GIDs with no paper/rectangle-clip cuts or unknown traces. Eight boxed PDFs are identical; all source/page/diagnostic identities remain. Independent static Word anchors retain their gates; new combined-source Word geometry is unqualified. |
| Fitting pictures in declared rows (prior milestone) | Fixed spacing bypassed wrapping and left a fitting picture's green center missing. Shared measurement now admits pictures contained above the declared baseline and grows minimum rows from their measured content. Twenty-four floating/inline automatic, fixed and minimum controls retain complete pictures. Eight inline outcomes keep every source text bound clear of the picture. The 28-outcome screen measures 4,046 actual-program-bound glyphs without paper/rectangle-clip cuts or unknown traces. Four oversized exact-slot fallbacks remain clipped and unqualified; authored floating `wrapNone` body-overlap residuals remain. New exact combined-source Word geometry is unqualified. |
| Nested picture/header equation | The preceding milestone restored native header math from 56.5% to 100% black-mark retention by correcting object-only origins. Its sixteen Final/Review probes retain all pictures and 55,880 actual-program-bound glyphs without paper/rectangle-clip cuts or unknown traces. Object-only frames preserve container displacement; text/math keeps independently tested Word anchors. |
| Leading blank/equation pages | The preceding milestone restored five complete footer labels (175 hidden glyphs) across three Review probes, including later note-only continuation pages. Static stories use their own measured page frame; standalone equations retain their measured body origin. Its six Final/Review outcomes contain 30,527 glyphs with no paper/rectangle-clip cuts or unknown traces. Portable regressions remain in the broader CI suite. |
| Note continuation composition | A leading blank page shifts the illustrated long-note report from four to five pages. All original text origins, sizes and bounds, picture/mask streams and link geometry remain unchanged; destinations move exactly one page. All 64 source continuation lines and both pictures survive in each mode. These are source-preserving transformation checks, not a new Word export. |
| Content-preserving transformations | Three held-out composed sources retain all native paint across six Final/Review outcomes after adding 149 paired paragraph bookmarks in body/cell/static/note stories (33,653 glyphs). A separate split of 179 equivalent text runs preserves actual-program/GID sequences, pictures, links, pages, paths/clips and diagnostics in six outcomes. Raw PDF/operation identity fails in all six split-run outcomes and remains recorded; maximum glyph-origin/bounds difference is 0.009308pt. A fresh actual-program ink screen measures 28,741 glyphs without cuts or unknowns. No new omission or pagination defect is demonstrated; no sub-point tuning or threshold change. No builds or Office actions. |
| Fitting mixed notes | Four sources/eight outcomes retain exact PDFs and all 64 continuation lines in their original order/pages. Twenty pictures retain both source colors and complete clips. Independent head/link/picture geometry keeps the original one-point gates. Native PDF/raster files are absent from the selected legacy folders, so no fresh whole-page comparison is claimed. |
| Earlier composition fixes | Table-first static frames, body-column printed-height admission, textbox picture/math neighbors and shared row/cell/box translation retain their bounded controls. Column balancing and general note-row fragmentation remain unsupported and diagnosed. |
| Static ownership repair | Broader CI exposed an unwanted ordered row for legacy images without source positions. The shared path now requires source ownership and keeps their block fallback. The previously failing ownership control and its wrapping neighbor pass; all 45 document outputs remain byte-identical, reusing their content and reference evidence. |
| Fixed-picture fallback | Broader CI caught changed placement of an oversized picture in an unmeasured exact-height row. The shared terminal-admission rule now retains that fallback; fitting declared pictures keep ordered measurement. All four fallback PDFs exactly match the preceding accepted outputs and the existing glyph/separation regression passes unchanged. Raw picture bounds alone had suggested NORTH-label overpaint, but native child clips preserve that label; no newly fixed text loss is claimed. The unfit picture remains approximate. |
| Preceding rendering checks | One exact 349-method selection per host: 347 Windows passes/two declared Aptos skips; 276 Ubuntu passes/73 predeclared font skips, with no failures or unexpected skips. Fifty-one neighbor PDFs remain identical; the sole changed SVG-footer PDF has identical glyph/path origins, pixel-identical 96-DPI output and identical fresh metrics against the pinned Word PDF. Original portfolio results remain two PPTX passes/fourteen DOCX failures; native Word pictures one/four. Twenty-four transferred assets/two smoke PDFs match. One accepted frozen artifact/shared fixture view; no root build or local full-DOCX replay. The broad parent test first fails on an inherited 0.072pt body-boundary overlap; the controlling content-loss proof is the pinned native parent screen (four clipped-word cases before, none afterward). |

The preceding autofit batch passed CI42976650266: 4,790 Windows/2,261 Ubuntu
checks, 19/129 skips, 69 .NET 10 checks per host and four package consumers.
The first leading-picture batch failed CI42986382781 on an existing page-start
assertion: the newly ordered body rows omitted before-spacing inspection metadata.
Its page-count/placement checks passed. The repair restores that metadata from the
shared spacing profile and retains the original assertions. All 36 rendering PDFs
are byte-identical; their content and independent Office evidence are reused.

The preceding leading-picture source passed CI42988198130:
**4,792 Windows/2,276 Ubuntu passes**, 19/130 skips, 69 .NET 10 checks per host
and four package consumers. Both terminal host logs were retained once. The earlier
failed run remains recorded. CI builds platform binaries separately; source coverage
does not imply binary identity across platforms.

The minimum-picture source passed CI42999234040 after a documentation-only
repair: **4,793 Windows/2,283 Ubuntu passes**, 19/130 skips, 69 .NET 10 checks
per host and four package consumers. The preceding failed CI42997335304 remains
recorded: its package preflight rejected a shipped relative CI hyperlink, which
resolved to an absent package file. The link was removed; the package check stayed
unchanged. Both terminal host logs were retained once. No local build or suite replay
was needed for the repair.

The contained fixed-line picture source passed CI43002444011:
**4,795 Windows/2,285 Ubuntu passes**, 19/130 skips, 69 .NET 10 checks per host
and all four package consumers. Both terminal host logs were retained once.
Platform binaries qualify source separately from the accepted local frozen artifact.
Status-only documentation updates reuse that qualification without local builds,
conversions or suite replay. The broader roadmap remains open.

The following-word/overhang source passed CI43008622492:
**4,796 Windows/2,286 Ubuntu passes**, 19/130 skips, 69 .NET 10 checks per host
and all four package consumers. Terminal host logs were retained once. The
broader roadmap and publication remain open.

The selected-comment-view source passed CI43015189979:
**4,798 Windows/2,287 Ubuntu passes**, 19/130 skips, 69 .NET 10
checks per host and all four package consumers. Terminal logs were retained once.
The subsequent composition checks reuse the same frozen production runtime;
documentation-only status updates preserve its qualification. The broader
roadmap and publication remain open.

Six additional page-admission probes reuse the frozen artifact: a fitting kept
cell stays on page one, an unfitting kept cell moves intact to page two, and a
body picture stays on page one while its following word resumes once on page two.
All 72 glyphs are clear of measured clips. No new-source Word geometry is claimed.

## Finite document checkpoint

The fixed portfolio contains nine sources and sixteen outcomes: both default
Final and Word-compatible Review for each DOCX, and one outcome per PPTX. Current
source, PDF, reference, mode and threshold identities have been reconciled. All
thirty DOCX Word/candidate page pairs have explicit pins. Four Final pages now
use pinned PDFium boards in place of less structured legacy thumbnail provenance;
the cached Word pages and existing candidate PDFs were reused. Shared template
ancestry limits the diversity of this cohort.

| Document | Final / Review pages | Reviewed content and pagination |
|---|---|---|
| Operations report | 2 / 2 | Heading, equations, illustration, table, items 1–3, sections 1–12, conclusion, note and footer retained. Sections 9–12 stay on page two. |
| Letter | 1 / 1 | Equations, illustration, acceptance items, sections 1–4, conclusion and note retained. |
| Specification | 2 / 2 | Sections 1–8 stay on page one; 9–20, conclusion and note on page two. |
| Held-out technical memo | 2 / 2 | Sections 1–8 and the equation table remain on page one; conclusion and note on page two. |
| Held-out static-cell report | 2 / 2 | Nested-cell/static-story equations, illustration, prose and transitions retained. |
| Held-out static-SVG report | 2 / 2 | Body illustration, header math and footer SVG retained; sections 9–12 stay on page two. |
| Held-out linked-footnote report | 4 / 4 | All 64 continuation lines retain order/pages; pictures, math and links remain. |
| Briefing and held-out decks | One page each | Pinned chart/table/text outcomes retain their two visual passes. |

This checkpoint reuses 56,513 actual-program glyph checks with zero measured cuts
or unknown traces. The separate authored-legend smoke remains identical to the
six-name/54-glyph audit. Its six origins still fail the one-point gate
(11.169/31.387pt), and MAE rose from 0.621337 to 0.713692 when content was restored.
The earlier whole-page pass had concealed missing names. Both portfolio PPTX
image gates pass; all fourteen DOCX gates still
fail, with typography, equation/table geometry and Review placement differences.
These are not blanket cosmetic classifications or threshold waivers. Two legacy
PPTX references have empty optional export logs; their source/mode/PDF identities
remain pinned. Application-view calibration remains unfinished.

Moving the existing mixed picture/equation paragraph into an ordinary table cell
in three portfolio sources adds six Final/Review controls. Native program/GID and
character sequences, picture pixels, owning pages, equations and diagnostics are
preserved; 30,327 glyphs have no measured cuts or unknowns. All eight picture frames
fit their rectangle clips. The generic screen retains eight alpha-mask abstentions;
a separate byte-level proof establishes that these unchanged Gray8 masks are fully
opaque, and all eight paired picture crops were reviewed. This qualifies these
source pictures without changing general mask handling or visual thresholds.

No new serious omission, clipping or wrong page was demonstrated in this bounded
review. The broader release gate remains open; candidate preparation is recorded
below and publication remains deferred. That evidence reconciliation reused the
qualified production runtime and closed CI evidence without a library build, full
suite, package qualification or Office action.

Candidate packaging must keep both framework assets and their portable symbols
bound to one committed build. Transferring a single-framework frozen runtime does
not refresh the other framework's root outputs. Reuse coherent qualified CI assets
or freeze both targets in an explicit worktree before packing. Verify every packed
library/PDB identity, current caller documentation and its local links. Cross-platform
package checks should consume the same selected package with each framework's
expected library hash, preserving the original qualified runtime.

## Prepared development package payload

Development packages, symbols, checksums and qualification records live directly
in `artifacts/nuget/`. Each coherent CI payload retains its compiled revision,
framework DLL/PDB identities and consumer results. Documentation-only derivatives
also identify changed entries; original results stay bound to their original ZIP.
The preceding CI43160881670 / `61b437941` payload and its caller-guide derivative
remain preserved as historical qualifications, including the constructor binary
break and recompilation warning.

Exact selected package, symbol and document hashes belong to
`SHA256SUMS.0.3.0.txt` and `Lokad.OoxPdf.0.3.0.qualification.json` beside the
artifacts; this guide does not embed its containing package's hash. Framework
entries must match their verified DLL/PDB bytes, with coherent 0.3.0.0 assembly/file
versions and committed informational versions. Packaged local links are checked.
Previous assets remain under source-qualified names; published 0.2.0 is protected.
Later qualification notes alone do not require another repack.

Source-runtime native/portfolio evidence stays separate from compiled CI identity.
Source/build equivalence, exact smoke PDFs and focused compatibility checks bridge
those artifacts without relabeling portfolio PDFs as package-produced. Reports
name the actual runtime, selected tests, inputs and original visual failures.
Transfer reuses the frozen artifact; it does not replay the full portfolio or suite.

`TestPackageSmoke.ps1 -InspectOnly` checks all framework build identities and caller
links without packing, restoring or loading assemblies. Consumer checks cover
conversion, cancellation, admission refusal, ownership and output publication.
The selected framework still requires its qualified hash: matching versions alone
do not qualify unverified bytes. Preparation does not complete the remaining
rendering/calibration/resource roadmap or authorize publication.

## Independent content checks and limits

Ink checks bind each native paint span to its **actual embedded font program,
resource and glyph IDs**. Normalized font names can alias regular/bold subsets;
the earlier name-only 112,702-glyph portfolio claim is withdrawn. No corruption of
library font data was demonstrated. Corrected evidence checks 56,513 unchanged
portfolio glyphs and 66,612 fitting-note glyphs through verified PDF identities.

Selected paint controls cover two composed roots in Final and Review. Six PNG
placements retain exactly the isolated source foreground, including the picture
on the continued footnote page. Four footer SVG instances retain both authored
shapes; maximum channel difference from isolated shape controls is two units.
All 64 continuation-note markers remain on Word's pages in each mode. These
checks reuse frozen PDFs, preserve the failed image gates and do not establish
arbitrary SVG/effect occlusion or application-view parity.

The complete fixed-portfolio PNG screen binds all 22 image paints to exact source
pixels, reusing six prior controls and adding sixteen final-paint controls.
Twenty retain the complete isolated-source foreground. Two page-two footer/cell
instances retain three/two pixels above the unchanged 16-channel-unit threshold
at 288 DPI; both source bindings and page bounds pass. These residuals remain
open. No missing picture was demonstrated; this does not prove arbitrary image
occlusion or replace the fourteen failed DOCX image comparisons.

Painted glyph/picture checks supplement source-text presence, extraction and image
scores. These bounded checks do not establish arbitrary nonrectangular clipping,
image occlusion, reading order or application-view parity. Oblique/mask clips,
stroked text and unavailable/unsupported font programs remain unqualified. Native
Python validation runs on Windows; Linux library checks are separate.

The current `DocxWordCompatibleBalloonBodyWrapsLikeOfficeTitles` check passes
without skips on both the current and preceding qualified 0.3.0 runtimes. Its
explicit-Aptos fixture reaches 20.999pt above the unchanged 16pt Office threshold,
with twelve primary font resolutions and no fallback. The original undeclared-
font input still uses twelve fallback runs and a 12.262pt second balloon. The
published 0.2.0 failure remains historical evidence; differing source font
contracts do not prove an original-input rendering fix or Word-default parity.
No renderer, assertion or font policy changed during this residual check.

A custom resolver mapping all text to pinned DejaVu Math exposes a small clipped
initial footnote A: 0.066pt in Final/0.050pt in Review. The same bearing overhang
exists in one-column controls. All six probes paint the complete source label;
the original 0.01pt ink gate still fails. This cosmetic residual is deferred.

## Remaining release gaps

- Fragmented/merged or unplaced picture-source targets retain unresolved ownership;
  exact combined-source Office destination geometry remains unqualified.
  Two Review run-split
  controls differ by 0.003479pt and fail the retained 0.001pt origin audit; picture
  and source-content checks pass. This small residual is deferred without tuning.
- The earlier claim that removing a synthetic 4.5pt/25% rotated-picture cut fixed
  a proven defect is withdrawn. The source-pinned Word inline-PNG PDF also clips
  25% of its transformed image; both anchored-PNG PDFs retain the full image.
  The exact composed-source clip remains unqualified. These supplemental Word
  PDFs use default Final and do not qualify the public AllMarkup gates or
  application view. Four original image-threshold failures remain unchanged.
- Pinned body/nested-table origins now pass all 32 strengthened container outcomes
  at the original 1pt threshold. Conflicting historical Word anchors and unmeasured
  combinations remain open. General note-row fragmentation/overflow, floating
  table placement, custom Review canvas, numbering, annotations and typography
  retain their stated limits. Unfit fixed-line picture slots keep their clipped fallback.
- The inherited unspaced picture/word witness belongs to pinned DejaVu Math,
  correcting its earlier Sans attribution. Its negative A bearing overlaps the
  later image by 0.072pt in body and cell controls, failing the original 0.01pt
  screen. Actual embedded outlines match the pinned source font. Image knockout
  affects ten dark pixels at 1200 DPI and none at 300 DPI; the complete word stays
  readable. Both spaced Math and Sans controls pass. This small overpaint residual
  is deferred; no geometry or threshold changed. Source/font/PDF identities and
  raster controls are retained. Exact Word appearance and arbitrary occlusion
  remain unqualified; rejected descent/paint trials remain unconsumed.
- Words following captured ordinary fixed-picture overhangs now wrap without loss.
  Terminal, reconstructed and prepared overhang frames keep their approximate
  fallbacks. Intentional picture clipping and exact new-source Word geometry
  remain unresolved; this does not establish arbitrary full-picture visibility.
- The native SVG fixed-footer PDF changes only its rotation clip from 30pt
  (including paragraph after-spacing) to its measured 22pt slot. Pixels and fresh
  Word-PDF metrics are identical; the original structural-similarity/recall
  failure remains. No colored SVG pixels appear in either candidate or the
  pinned Word PDF. Application view is unavailable, so no stronger source-picture
  completeness or newly demonstrated library-omission conclusion is claimed.
- An unfit 300-label note row retains 279 complete labels, six partial and fifteen
  off-paper labels. The pinned Office PDFs also retain 279 complete/six partial,
  but omit the other fifteen paint operations. Both fitting 240-label controls
  remain complete. General fragmentation stays unresolved and warned.
- PPTX manual/side furniture, automatic ranges, dense labels and font/plot/palette
  geometry remain approximate. No new count-specific or document-shape patch is
  claimed. Original visual thresholds and failures remain unchanged.
- **Office application-view calibration is unfinished.** Office PDFs remain
  bootstrap references. Unavailable or rejected captures block dependent
  conclusions only; no new Office export was needed for these fixes.
- Coarse host-memory coverage remains partial. Safe managed execution, cancellation,
  input bounds, bounded caches/staging, deterministic ownership/refunds, atomic
  publication and PDF validity remain; no comprehensive heap ceiling is claimed.

The broader roadmap remains open; publication/tagging stay deferred. See
[capabilities](Capabilities.md), [diagnostics](Diagnostics.md) and
[caller migration](PublicApiChanges.md).
