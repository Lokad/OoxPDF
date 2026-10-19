# Composed-document reliability checkpoint

## C268 authored signed table indentation (2026-10-19)

Negative authored table indents were clamped to zero, causing36pt/41.4pt text
placement errors. The shared grid now preserves signed placement; modern text
and legacy outer-border anchors keep their existing roles. Hanging placement
does not enlarge the container basis for percentage widths. Seven source-exact
controls/14Word views cover fixed/auto/percentage, style, border, legacy and
small negative indents. All prose/math are retained; glyph origins meet1pt,
maxX0.036pt/Y0.371pt. Word object-model cell widths match within0.1pt. Bookmarks
and unchanged-format run splitting preserve content/placement. Four rasters
reviewed; all14strict visual failures remain under unchanged thresholds.

One committed frozen Release artifact:15Windows/8Ubuntu/7native font skips,
one exact invocation/platform, no harness repairs. Both new observers fail the
parent. Existing28table-control/10fixed comparisons reuse actual source/parsed
reachability proof; no repeated conversions/Office export for those outcomes.
17assets/four smokes, no root build/suite/package replay. Evidence:
artifacts/composition-c1/table-signed-indents-evidence.json.

A separate nine-source Final-only starting-border review found varying implicit
Word indentation; removing just the border exclusion leaves an unexplained
offset. No implicit-border correction or border-width family accepted. Preserve
the5.736pt residual. table-start-border-ownership-review.json. C267CI39244587426
is green on Windows/Ubuntu, including broader/package checks.

## C267 shared modern table-grid origin (2026-10-19)

Authored starting margins without a visible starting border now belong once to
the shared grid, including explicit-zero styles and unrelated border edges.
Legacy, positive indent/style and visible starting-border anchors remain.
Fourteen source-exact controls/28 independent Word mode views:26 admitted views
meet1pt/.1pt geometry, max X0.024pt/Y0.514pt; the two visible-start-border PDFs
retain identity and their5.736pt X residual. All28 strict visual failures remain.
The unchanged held-out compound document retains three formulas/12math glyphs,
one image and one internal link. Both modes now meet1pt/.1pt absolute geometry:
math X0.129pt/Y0.165pt, image0.153pt, link0.046pt, target0.270pt. No prose/math
omissions; two held-out strict visual failures remain. Four rasters reviewed.

One production artifact:35Windows checks and18portable Ubuntu passes. Eight
historical Windows-Arial assertions fail identically on unchanged C266; native
prerequisites are now explicit. Original8failures/9skips remain. The tracked
joiner rejected failure-to-skip replacement; no passing Ubuntu35aggregate is
claimed. Only those eight prerequisite checks reran/platform;14runtime assets
and all2,331inputs are identical. Both new observers fail C266 meaningfully.
Eight new edge/style Word views;22 prior views reused. Fixed10outcomes reuse
parsed/source reachability proof, with zero fixed conversions.17assets/four
root smokes, no root build/suite/package replay. Broader/package CI follows.
Evidence:artifacts/composition-c1/table-grid-origins-evidence.json.

## C266 coarse host memory admission and bounded rule review (2026-10-19)

Removed12 cumulative-work/retained-image/retained-font/live-image limit members
and their counters. Six existing output/page/staging limits remain. Archive/XML/
image/font bounds, cancellation, useful caches and PDF structural validation stay.
Optional IOoxPdfMemoryAdmission supplies one host-sized shared-memory lease per
conversion, covering significant retained storage and construction peaks. Refusal
precedes package construction/output; refunds follow staged ownership and precede
atomic file publication. Existing shared bounded caches remain a host baseline.
This is coarse host admission, with no library heap-accounting promise, object
ledger, graph tracking, finalizer, replacement fuel or extra quota knob.
Decoded/staged owners now drop their references on disposal; crop/recolor poll
existing cancellation tokens. Public migration:docs/PublicApiChanges.md.

One committed frozen Release artifact:491Windows/484Ubuntu/7native skips. All
refusal, cancellation, concurrent/nested admission, lifetime and publication
checks pass. Two prequalification harness build failures are retained. One
observer incorrectly required fonts in a picture-only deck; only that corrected
check reran/platform,14runtime assets/all2,315inputs unchanged. Tracked joins
preserve original failures. No complete-suite repair replay. CLI tests reuse the
same immutable CLI; package smoke loads the exact455F6D38library and checks the
public admission API without rebuilding it.17assets/four root smokes, no root
build/suite replay.10fixed PDFs/Office comparisons and five picture decks/26pages
retain identity. Both portfolio decks and25picture pages pass;8DOCXstrict gates
and one prior cosmetic picture failure remain. Content-audit screens are zero.
No Office exports/rasters or performance claim.

S4 is a separate read-only assessment:15 chart-color regimes6..90 and the96+
fraction curve still use count bands; five-paragraph marker/font/row guards are
narrowly qualified. The source notes same-frame confounding; model combinations
do not establish a general Office rule. Keep existing behavior/fallbacks and
explicit residuals. No policy change, new special-case family or renderer rewrite.
Resume the saved C264 shared table/grid-origin review with its independent Word
controls. Evidence:artifacts/composition-c1/memory-admission-evidence.json and
simplification-rendering-generalization-review.json.

Workspace scan:15.300GB/822,062files/zero errors; the loose5GBtarget remains unmet.
Automatic approval review rejected removal of four completed runner scratch
directories/660files/1.536MBlogical as "blocked by policy" before process creation.
Zero removed/reclaimed; preserve those four targets without alternate retry.
Evidence:memory-admission-completed-scratch-retirement-rejected.json.

## C265 bounded simplification S1/S2 (2026-10-19)

Removed13 verified unreferenced methods (111lines): PlaceRelatedStory,
MeasureDrawableTextSpans, RenderPositionedTextSpans, FlattenTextLayout,
ReadWorkbookData, ReadSceneSecondaryRightValueAxis, GetPlotBoxRatios,
IsDarkChartPattern, ReadOptionalOoxmlBooleanElement, ReadChartElementDouble,
ReadChartElementInt, ReadChartAxisScalingValue and ReadChartAxisUnitValue.
Workbook tests call live ReadWorkbookDataCore with cancellation. Positioned-font
preparation used by the text inspector remains. Global memo totals/reset are
removed; the cache and local counters remain. Its revised observation compares
cached and fresh cell-placement geometry using one local memo.

Removed no-op public Deterministic and Strict options; migration is documented
in docs/PublicApiChanges.md. Output remains deterministic; CLI --strict still
publishes the same PDF and returns3 for warnings versus ordinary0.
One frozen Release artifact:25Windows/23Ubuntu passes. Two native chart checks
remain unqualified on Ubuntu and reproduce identical failures on unchanged C263;
raw reports stay failed, with no harness repair or pass relabeling. All10fixed
portfolio PDFs equal retained qualified outputs;10Office comparisons reuse exact
PDF/reference/threshold identities. Both decks pass;8DOCXstrict failures remain.
Content audit has zero deficit/unresolved screens.17assets/four transfer smokes,
zero root builds/suite replays/Office exports. S3 ownership/admission and S4
rendering-rule assessment remain active; no protection behavior changed here.
Evidence:artifacts/composition-c1/library-simplification-evidence.json.

## C263 cell-terminal paragraph-mark floors (2026-10-19)

Without a separate non-spanning cell, a merged source now gets its minimum row
strut from the parsed terminal paragraph mark. Wrapped content/visible run size
remain a complete-span requirement. Reuse existing line-height/spacing rules;
unknown marker metrics preserve the prior estimate. No font/size-specific branch.
Ten separate controls/both presets distinguish source/continuation marks, visible
run size/face, marker face, document defaults and multiple paragraphs. Twenty
outcomes pass existing1pt Word geometry, maximum0.401pt. The former20.65pt wrapped
source overcharge closes; source/Office/simple-math deficits zero. Both new
observers fail C262. Twenty paired pins/strict visual failures remain, alongside
inherited cell-left5.376pt.
One production artifact:98Windows/79Ubuntu/19native skips, one exact invocation
per platform. Ubuntu exposed an older terminal-space observer's installed-font
assumption; it also fails on C262. The observer now uses the existing pinned font;
only that method reran/platform, with14runtime assets/all2,315inputs unchanged.
Original failed attempt/parent diagnostic retained. Prior18source outcomes/fixed8
reuse source proof.17assets/four smokes include two new wraps and two byte-identical
24pt math spans; no root build/suite or repeated Office export. Root8432e284;
artifacts/composition-c1/merge-cell-mark-floors-evidence.json. C262CI39229269664 is
green on Windows/Ubuntu. Next:C264shared table/grid horizontal-origin review.

## C261/C262 complete source-span row floors (2026-10-19)

A held-out combination retained three equations, an image and navigation but
exposed20.2pt later-owner and15.962pt following-flow errors. Eight separate Word
controls establish the rule: a non-spanning cell owns its row floor; wrapped
prose/images and multiple paragraphs belong to the complete admitted span.
The shared span path now handles plain sources as well as math, using the existing
cell-property eligibility rule. Where an independent floor is unknown, keep the
prior estimate. Fixed/partial/dynamic/forced/header/note-dependent sources retain
staging; missing inline bounds retain continuation fallback.
Sixteen source outcomes pass existing1pt Word flow anchors, max0.050pt. The unchanged
held-out input then passes all12math glyphs, prose, image and link checks: max glyph
Y0.165pt, following prose0.060pt. No held-out rule fitting. All12prior math PDFs
retain their bytes; fixed8reuse source proof.18new/held-out paired pins and18strict
visual failures remain; source/simple-math/Office-word deficits zero. Inherited
cell-left5.376pt remains.
One production artifact:91Windows/73Ubuntu/18native skips, one exact invocation
per platform; only one corrected model rerun/platform after preserving its
line-only measurer fallback.14runtime assets/all2,303inputs unchanged; original
failure retained. Both new observers fail C258.17assets/four smokes; no root build,
suite replay or repeated Office export. Rootaecf7cd1; ignored
artifacts/composition-c1/merge-source-row-floors-evidence.json. C258/C260CI are green
on both platforms. WorkspaceC262is15.158GB/816,701files, zero errors; no reclaim
claimed and the loose cap remains unmet.

## C260 pinned inspection reconciliation (2026-10-19)

The routine portfolio command now accepts -ReuseCandidateDirectory to reinspect
source/PDF/runtime-pinned candidates with a different qualified inspector.
Preflight rejects changed runtime/input/PDF, missing assets and duplicate outcomes
before publishing output. Original PDF bytes, cold CLI timings and outcome/report
provenance remain pinned; the tracked self-test proves zero converter calls.
C259/C258's30PDFs and30strict comparisons are unchanged. The existing immutable
ActualText inspector84EA672B corrects older root-inspector y-to-x extraction:
current source/simple-math/Office-word and per-page deficits are zero. All30strict
visual failures and original30raw token screens remain in evidence; this is no
new visual pass. No conversion, library build, rendering test replay, Office export
or raster work. Rootccee67e7; artifacts/composition-c1/portfolio-reinspection-evidence.json.

## C258 complete merged-span admission (2026-10-19)

Static measured mathematics pays its expansion once across a complete admitted
span. Existing row floors remain; a forward prefix pass satisfies overlapping
owners. The admitted frame is reused through placement, and its full source is
consumed once. Fixed rows, partial/oversized/dynamic frames, forced breaks,
headers and note-dependent capacity retain staging. This is a bounded body-table
contract, not general fragmented-span layout.
Six source-pinned probes/both presets now match independent following-flow and
page anchors within0.308pt. Side row2stays on Word'sp1; its old page flags close.
Exact-row PDFs retain their prior bytes. The native flow test fails on C259.
One production artifact:77Windows/60Ubuntu checks,17native skips. Initial exact77
ran once/platform; two model checks reran after a1e-9 arithmetic floor tolerance,
with14runtime assets/all2,293inputs unchanged and failed attempts retained.
Twelve outcomes/fourteen paired pins; two comparisons reused by exact identities,
12strict visual failures remain. Four fixed DOCX inputs plus a static neighbor
retain source proof; fixed8outcomes reused.17assets/four smokes, no root build,
suite replay or Office export. Twelve raw x-token screens from the older root
inspector remain; frozen ActualText-inspector reconciliation is the next bounded
qualification task. Independent glyphs/ink show x+2/y. Inherited cell-left5.376pt
remains. Rootbfb2aa70; artifacts/composition-c1/merged-span-admission-evidence.json.
C259CI39223543480 is green on Windows and Ubuntu.

## C259 authored fraction-control frames (2026-10-19)

Filled bar/stack fractions now use their resolved authored control size for rule,
gaps and leading while argument runs keep their source/script sizes. Neutral
controls retain enclosing inheritance; unknown styles do not invent an owner.
Nine deterministic source-pinned controls cover smaller/neutral/larger sizes,
body/cell/merged placement and direct/character-style declarations. Both presets
retain four independently matched8.04pt glyphs; maximum math/prose baseline error
is0.384pt under the existing1pt/.1pt contracts. Ordinary following flow passes;
merged flow remains up to20.322pt off and inherited cell-left5.376pt remains.
One frozen Release artifact:92Windows/67Ubuntu checks,25native skips, one exact
invocation/platform. The native Word-anchor regression fails on C257. Eighteen
outcomes/eighteen paired pixel pins;18strict failures and18raw Office x-token
screens remain recorded. All source prose/simple math is present; viewed native
and Office ink plus mapped glyph checks supplement, rather than erase, extraction
flags. Fixed8DOCX outcomes reuse source proof.17assets/four smokes; no root build,
suite replay or duplicate Office exports. Initial assertion-message build failure
is retained; no tests ran on it. Rootc55ee32a; ignored
artifacts/composition-c1/fraction-control-frames-evidence.json. Next:C258 shared
merged-span height/admission; preserve exact/minimum floors and partial fallback.

C257 admitted merged-source consumption (2026-10-19): complete static top-aligned inline frames already admitted in the restart row/page are consumed once. Continuations retain visual source/shading/border ownership; dynamic, unmeasured, shifted, fixed-incomplete and oversized frames keep fallback. Reuses the existing finite containment predicate and consolidates page capacity in the context. Three-row, overlapping-owner and page-boundary controls now retain each original paragraph/fraction once; Word glyph pages/baselines keep1pt gates with bookmarks/run splitting. One frozen production artifact passes41Windows/31Ubuntu methods,10native skips. One initial41-method invocation/platform, then only one model rerun/platform after adding oversized-owner proof;14runtime assets/all2,278inputs unchanged and tracked joins retain provenance. Both new native/model observers fail C255 meaningfully. Original4PDFs and six untouched span-probe PDFs remain byte-identical; fixed8outputs reused by source proof. Six new outcomes/eight paired pins; native three-row/p2and portablep2reviewed. Six strict failures/two page-membership audit flags retained: Side row2remains nativep2vsWordp1; no global source/Office-word/simple-math deficit.17assets/four root smokes, no root build/suite/repeated Word exports. Root300c6d0d; artifacts/composition-c1/merged-source-consumption-evidence.json. C255CI39164253097 had3553Windows passes/two obsolete blanket-restart assertions/16skips, Ubuntu green. Tests now distinguish explicit source restart from continuation/unknown roles and preserve placement guards; missed nested/note neighbors join the capability. Failed CI logs remain, C257CI tracked after push.

C256 bounded merged-span review (read-only, 2026-10-19): six source-exact controls exercise row count, minimum/exact floors, page boundary, overlapping owners and a taller fraction control on C255. Twelve outcomes/twelve Word views once,14paired pins, all12strict failures retained. Raster review confirms repeated original prose/equations in three-row, overlapping-owner and page-transition cases. Minimum/ordinary spans retain about7.17pt excess following flow; exact rows match following flow within0.02pt. Near-boundary Side row2moves to nativep2rather than Wordp1. The24pt fraction-control probe retains matching8.04pt child glyph sizes; cell-prose baseline differs10.223pt, with numerator/denominator spacing discrepancies. The initially inferred child-size mismatch was corrected against recorded glyph data before qualification; do not substitute child fonts/sizes. Cause/measurement and shared span admission remain separate C259/C258work. A distinct unqualified consumption prototype motivated C257, with no row-height/typography change. Evidence artifacts/composition-c1/merged-span-review-evidence.json.

C255 owning vertically merged-cell mathematics (2026-10-18): preparation follows the explicit restart cell's original visible paragraph; restart ownership is shared with existing merge layout. Placement/typography guards remain unchanged, and continuation/unknown/fit/vertical owners retain fallback. Both presets restore all4fraction glyphs and surrounding prose; relative math anchors meet1pt/0.1pt size gates against Word, including bookmarks/run splitting. One frozen Release artifact passes31Windows/24Ubuntu methods,7native skips, one exact invocation/platform. The new native observer fails C252 on missing mathematics. Both unmerged PDFs remain byte-identical; two comparisons reused/two fresh, four paired pins and native/portable Review pages inspected. Four strict failures remain: existing absolute cell-left mismatch5.376pt, merged following-paragraph difference7.166pt. The latter follows full math height being charged to the first row while the complete merged span already has space; no fitted offset or absolute-fidelity claim. Four candidate outcomes only; fixed sources have no merged math owners, reusing eight accepted outputs by source proof. Content audit has zero unresolved screens.17assets/four root smokes, no root build/suite/repeated Office exports. A derived-Tables test compile error was fixed before freezing/testing; failed build retained. Rootb344d10b; artifacts/composition-c1/owning-merged-math-evidence.json. C252CI39163065672 is green on both platforms.

C254 merged-cell composition review (read-only, 2026-10-18): the supported fraction/prose cell receives an empty second row as control; declaring restart/continue ownership is the only merged variant change, with XML reversal verified. C252retains4math glyphs in the unmerged control and zero in the merged cell, both presets; Word retains all4in both. Source prose remains, so the missing equation is a confirmed composition omission. Four Office views once/four paired pins, merged Review pair inspected; strict failures retained. No library build/suite in the review. A separate unqualified source-preparation prototype restored content and exposed the row-span spacing residual; no cache-corruption claim. Evidence artifacts/composition-c1/merged-math-cell-review-evidence.json.

C253 combined held-out note (read-only, 2026-10-18): the accepted C252 artifact handles two fractions, a hyperlink/bookmark, the picture and40pt minimum together without another rule. Both modes match Word4pages/20-30-14 memberships. All64numbered baselines,16mathematical glyph anchors, note heads/targets, picture bounds and hyperlink edges/destination meet existing1pt gates; source64/two pictures/50math glyphs survive once/order, zero ignored equations. All four linked source glyph origins are clickable; neighboring unlinked ink can touch padded bounds in Word too. Twelve unaffected page PNGs retain bytes against the minimum-spacing control across both sides/modes. Eight paired pins retained, Reviewp3pair viewed, both strict failures remain. Two candidate outcomes/two Word exports once; no library build/suite or growing-family replay. Keep the source held out. Evidence artifacts/composition-c1/combined-linked-note-evidence.json.

C252 measured hyperlink ownership in note placement (2026-10-18): valid source ranges now annotate their existing measured content without disqualifying the paragraph placement frame. Original runs and typography rules remain; incomplete/unmeasured ranges and fixed/fragmented/complex fallbacks stay. Both modes match Word4pages/20-31-13 note memberships, with all64paragraphs/two pictures once/order. All64numbered baselines, heads/targets, image/math origins, mathematical link edges and bookmark destination meet existing1pt gates; all4linked mathematical origins are clickable. One production Release artifact passes45Windows/34Ubuntu methods,11native skips. Initial45-method invocation/platform plus only2Windows/1Ubuntu corrected checks: a new width assertion aggregated opposite edge errors; repaired to the established edge contract and strengthened ink coverage.14runtime assets/all2,269inputs unchanged, failed report retained, tracked join command preserves provenance. Native observer fails C250; a parent model attempt cannot bind the new internal helper and is excluded from regression claims. All26earlier PDFs unchanged;26comparisons reused/two fresh,108paired pins.28strict failures/four raw math-word screens retained, zero source/simple-math deficits. Fixed sources have no related hyperlinks, reusing eight accepted outputs by source proof. New Reviewp3pair/portablep4 reviewed;17assets/four smokes, no root build/suite/repeated Office exports. Root1cd73f19; artifacts/composition-c1/linked-note-placement-evidence.json. C250CI39162326418 is green on both platforms.

C251 linked mixed-note holdout (read-only, 2026-10-18): wrapping the existing note fraction beside its picture in an internal hyperlink changes only footnotes.xml; XML reversal and existing Calculation target verified. C250produces5vsWord4pages in both presets, delaying note2fromp3top4. All64paragraphs/two pictures/math and a working link remain; confirmed placement defect with no observed omission. Two Word views once, eight paired/two unmatched raster pins, Reviewp3pair reviewed; both strict gates retain pagination failures. No library build/suite. Evidence artifacts/composition-c1/linked-mixed-note-review-evidence.json.

C250 measured minimum note batches (2026-10-18): minimum-spaced paragraphs now share the content-driven note path when complete, source-owned inline frames meet the authored height floor. Measurement and typography are unchanged; fixed/fragmented/complex and legacy image-only fallbacks remain. Both modes match Word's four pages: low-minimum note memberships20/31/13, larger minimum20/30/14. All64 numbered baselines, note heads/targets, pictures and mathematical anchors meet existing1pt gates; bookmarks/run splitting preserve geometry. One frozen Release artifact passes125Windows and99Ubuntu checks,26native skips, one exact invocation/platform; both new native/model observers fail C246. All22earlier PDFs and both neutral-minimum PDFs retain bytes; eight fixed portfolio outputs reused by source proof.22comparisons reused/four fresh,100paired raster pins; new large-minimum Reviewp3pair and portablep4 reviewed.26strict visual failures/four raw math-word screens retained, zero source/simple-math deficits.17assets/four root smokes; no root build/suite or repeated Word exports. Rootba2f74f0; artifacts/composition-c1/measured-note-minimums-evidence.json. C248CI39160011043 is green on Windows/Ubuntu, including broader/package checks.

C249 authored-minimum mixed-note holdouts (read-only, 2026-10-18): adding only a1pt or40pt minimum to the existing picture/fraction paragraph produces5native pages in both modes versus Word4, delaying note2 fromp3 top4. All64paragraphs/two pictures remain once/order; mathematical counts are51(22body+5perpage+4note), zero ignored equations. Word's1pt output retains all8automatic-control PNGbytes. Existing source-owned inline frames already measure the larger minimum correctly; the automatic-only batch guard is the confirmed composition defect. A bounded content-driven/frame-floor prototype matches Word4pages with20-31-13or20-30-14 memberships and subpoint heads/pictures; its neutral PDFs retain accepted automatic bytes. Fixed/fragmented/legacy-unmeasured image frames keep fallback. No typography rewrite or fitted offset. Four Office views once;16paired/four unmatched candidate pins retained, all4strict gates remain failed. Review has no library build/suite; the separate prototype is unqualified pending C250. Evidence artifacts/composition-c1/minimum-mixed-note-review-evidence.json.

C248 reusable qualification report joins (2026-10-18): JoinQualificationTestReports.ps1 consumes one manifest and checks committed production source, every protected runtime asset and all public input hashes across declared frozen artifacts before atomically joining reviewed test-only repairs. Only test DLL/executable/PDB may differ; dependencies/configuration remain protected. Exact selections cannot drop methods, undeclared repairs are rejected and an initial failure must end passed; raw failed attempts/source identities remain. A small local fixture checks positive provenance plus six pre-publication rejections and restores its original pinned inputs afterward. Git helper recursion/scalar-index mistakes were caught by self-checks and corrected before integration; failed logs retained. Applied to C246's122Windows methods/124attempts with14runtime assets/2,263inputs unchanged, reproducing all prior outcomes without builds, tests or conversions. Four tracked tool/config blobs match transfer; CI adds the self-check on both platforms. Root87b86ab4; artifacts/composition-c1/qualification-report-join-evidence.json. C246CI39158486662 is green on Windows/Ubuntu, including broader/package checks.

C247 second-fraction/bookmark note holdout (read-only, 2026-10-18): the accepted C246 frozen artifact handles a second atomic fraction inside bookmark metadata beside the note picture with no new rule. Both modes match Word4pages/20-31-13 memberships; all64paragraphs/two pictures once/order and50math glyphs, zero ignored equations. All16note-math glyph anchors, note heads/targets and picture bounds meet existing1pt gates. Every page outside the added mathematics is pixel-identical to the single-fraction control in candidate and Word:12PNGidentities across both modes. Eight paired raster pins retained, Reviewp3pair viewed; both strict visual failures remain. Two Word views once, no library build/suite. Keep this held-out source outside immediate public development fixtures. Evidence artifacts/composition-c1/two-formula-picture-note-evidence.json.

C246 measured atomic math/picture note batches (2026-10-18): complete source equation ownership and existing finite inline frames admit mixed automatic paragraph batches, without promoting typography. Missing/duplicated/foreign/unmeasured owners, fragmented formulas, fixed/minimum and complex sources retain fallback. Both modes match Word4pages/20-31-13 numbered note memberships; all64paragraphs/two images once/order. Heads, targets, image bounds and four added mathematical glyph anchors keep independent1pt gates; bookmarks/run splitting stable.122Windows/97Ubuntu checks pass,25native skips: one initial121-method Windows run plus only2-method/1-method harness repairs, one final122-method Linux invocation.14production runtime assets/all2,263inputs unchanged throughout repairs; failures and separate trace diagnostics retained, production source restored exactly. Initial model mutation invalidated its source view; A+B then legitimately fragmented, corrected to an atomic ownership control. The new native observer fails C244. All20earlier note PDFs unchanged;20comparisons reused/two fresh,84paired pins, newReviewp3pair and portablep4reviewed.22strict failures/four raw math-word screens retained; zero source/simple-math deficits. Four fixed DOCX sources have no note equations, reusing C244eight outputs without replay.17assets/four root smokes, no root build/suite/repeated Word exports. C244CI39156101487 had3542Windows passes/one missed synthetic margin assertion, Ubuntu green. Its exact native source PDFs are byte-identical before/after: observer now separates trailing after-space from occupied line bounds and keeps25take; exact check added to capability selections. Inherited unformatted Word/library font-default and page-membership differences remain recorded. Rootcdc4306a and observer7c233a2e; artifacts/composition-c1/mixed-note-batches-evidence.json. Workspace14.714GB/794,117files, zero errors; loose5GBtarget remains unmet and all rejected retirements protected.

C245 mixed equation/picture note holdout (read-only, 2026-10-18): cloning one already-supported body fraction beside the existing note2 picture changes only footnotes.xml, with source reversal checked. C244produces5pages in both modes versus Word4 and delays the second note/image/formula fromp3 top4. All64paragraphs/two pictures remain once/order;51mathematical glyphs match22body+5perpage+4note with no ignored equation. The defect is layout composition, with no observed source omission. Two Word views exported once, eight paired/two unmatched candidate raster pins retained; Reviewp3pair reviewed. No library build/suite. Evidence artifacts/composition-c1/mixed-math-picture-note-review-evidence.json.

C244 admitted note seating and mapped-header boundary (2026-10-18): fresh physical-frame heads consume design after-space fit rules in both print transforms; their retained occupied height cannot raise seating above the admitted top. Mapped headers expose a physical boundary and use the same design conversion as footers, preserving the legacy unmapped path. No font/size/document-shape qualification added. Three public wrapping/later-picture controls match Word4pages and all numbered memberships in both modes;64paragraphs and2/3images once/order, head/target/image gates<1pt with bookmark/run splitting.118Windows/94Ubuntu checks pass,24native skips, one exact invocation/platform; the new independent geometry observer fails C242. Thirteen older note PDFs and all eight fixed DOCX milestone PDFs remain byte-identical. Original-footer Final moves first-page baselines0.003pt at the clamp, retaining all Word gates and unchanged remaining pages; fresh comparison kept.13comparisons reused/seven fresh,76paired pixel pins, new Reviewp3pair and portablep4reviewed. All20strict failures/four raw math-word screens retained; zero source/simple-math deficits, no repeated Word exports.17assets/four root smokes, no root build/suite replay. Rootf9fc81e5; artifacts/composition-c1/note-batch-frames-evidence.json. C242CI39154681802 green on both platforms; C244push/CI tracked while the next mixed equation/picture note holdout proceeds.

C243 wrapped/multiple-picture note holdouts (read-only, 2026-10-18): the accepted C242 frozen artifact retains all64paragraphs and all2/3pictures once/order. Word and candidate wrap the expanded intro into the same six lines, yet Final's head is5.833pt early and Review's10.573pt late; Review moves one extra paragraph top4. Wrapping-only and later-picture-only controls reproduce independent geometry errors. Eight paired PNGpins retained and Reviewp3pair reviewed. Six independent Word views exported once across three sources; no library build/suite for this review. Source owners and image growth remain consistent. A bounded prototype identifies fit-space versus occupied-height seating and a mapped-header/legacy-coordinate mismatch; adding an explicit physical header boundary with the existing footer conversion gives the three controls' Word page membership and subpoint heads/pictures. Prototype-only results are not qualified; C244 is qualifying the coherent batch, retaining all controls/thresholds and fallback branches. Evidence artifacts/composition-c1/wrapped-note-pictures-review-evidence.json. C242CI39154681802 is green on Windows/Ubuntu.

C242 source-owned note pictures (2026-10-18): measured image source-line owners now follow text slicing; image growth is accounted separately from paragraph after-space. Existing source-affinity/property rules admit automatic picture-note batches while preserving typography, fixed/minimum and complex-story fallbacks. Both modes match Word4pages/20-31-13 numbered paragraphs; all64once/order, two image operations once, heads/targets/image rectangles within1pt. Bookmark/run-split transformations retain picture geometry.100Windows/80Ubuntu checks pass with20native skips, once/platform; all three new observers fail C240. Twelve earlier PDFs remain byte-identical;12Office comparisons reused/two new,52paired PNGpins, nativeReviewp3pair and portablep3/p4reviewed. All14strict failures/four raw math-word flags retained; zero source/simple-math deficits. No new Office exports. Four fixed DOCX sources contain no related-story pictures, retaining C238eight outputs by source proof.17assets/four root smokes, no root build/suite. Root96b22707; artifacts/composition-c1/note-picture-ownership-evidence.json. C240CI39150776688 is green on both platforms. Next: a held-out wrapped/multiple-picture note combination against the accepted frozen artifact, before choosing another implementation slice.

C241 long-note inline-picture holdout (read-only, 2026-10-18): append a clone of the existing report image to note2intro, changing only footnotes.xml and adding its local relationship part; body, references,64numbered paragraphs and original picture bytes stay unchanged, source reversal checked. Word retains4pages/two image placements and20/31/13 numbered note paragraphs. Native Final has4pages/two operations but20/12/32 membership; Review has5pages/three image operations and20/12/30/2 membership. Both delay the second note/picture fromWordp3 top4. All64paragraphs surviveonce/order and both source targets remain. The extrap5image operation is clipped out in the reviewed raster: redundant emission, not demonstrated visible duplication or corruption. Eight paired/one unmatched candidate PNGhashes retained; nativeReviewp5/WordReviewp4 reviewed. Two Word exports once, both strict failures, zero library builds/suites. Image-bearing paragraphs leave the plain batch guard; canonical image ownership through narrowed/sliced stories is the next bounded investigation. Evidence artifacts/composition-c1/long-note-inline-picture-observations.json.

C240 implicit-rule physical frames (2026-10-18): separator-free plain notes use the existing occupied body/table boundary, natural source-font rule height and paragraph takes; continuation width follows that same admitted frame, with conversion cancellation propagated. The accepted reader/default representation, mathematical continuation and authored fallbacks stay intact. Both modes match Word4pages/22-10-18-14 numbered paragraphs, all64once inreference order, head/target/rule gates<1pt and clear table/math/footer ink. Both new observers fail C238. Ten earlier PDFs retain bytes; all8new native page images equal the explicit-separator control.10Office comparisons reused/2fresh,44paired pins, Reviewp2pair and pinned-font p2viewed. All12strict failures/four raw math-word screens retained; source/simple-math deficits zero, no new Office exports. Final87Windows/70Ubuntu checks,17native skips: one87-method invocation/platform and only2Windows/1Ubuntu observer repairs to exclude table borders, with14runtime assets/all2,253inputs unchanged and no library rebuild. The initial reader-normalization approach and a shared-owned-height prototype changed independently validated math-note takes, were rejected and stay off master; two other initial failures exposed full-path footnote/endnote classification in tests, corrected to filenames. An intermediate short-continuation rule was caught by pixel review before publication; stronger Word/default-width observers retain that evidence. Four fixed DOCX sources explicitly declare separators, so retain C238eight outputs without replay.17assets/four root smokes, no root build/suite. Rootde3e0462; artifacts/composition-c1/implicit-note-frames-evidence.json. C238CI39147409207 is green on Windows/Ubuntu. Automatic approval review rejected removal of164completed test scratch directories/27,827files as blocked by policy: zero files/bytes removed, no alternate retirement/retry; preserve the inventory and existing protected artifacts. The loose5GB workspace target remains unmet.

C239 absent table-note separator holdout (read-only, 2026-10-18): removing only the typed separator story preserves all64raw native paragraphs once/order and4pages, but note heads/rules overprint table/body in both presets; Reviewp1paired raster confirms unreadable text. Heads are26.099ptFinal/50.925ptReview early; subsequent head72.012/118.896pt early. All8independent Word rasters are byte-identical to the C238explicit-separator control: Word supplies the visible default. Two exports once, two strict failures/eight paired PNGhashes retained, no library build/suite. C240 will investigate existing default-story parsing; the reader intentionally retains legacy absent-footnote behavior while synthesizing absent endnote separators, so this requires source-default and neighboring Office controls rather than a fitted table adjustment. Evidence artifacts/composition-c1/table-note-without-separator-observations.json.

C238 table-note admission and occupied body boundary (2026-10-18): remove the duplicated complete-story table reserve/map/factory; live in-flight source owners admit the table. Finite inline image bounds and closed row frames share the existing physical body-to-note boundary independently of footer typography, retaining paragraph-spacing/complex/column fallbacks. New public table-note-reference-report retains every source node and reverses note reference order2/1. Both modes now match Word4pages,22/10/18/14 numbered paragraphs; all64once/order33-64/1-32, heads<0.245pt and targets<0.839pt, disjoint table/math/footer ink. Final69Windows/54Ubuntu checks,15native skips, once/platform; both new observers fail C236. The initial35/28-pass artifact still overlapped the table: independent Reviewp1raster rejected it; stronger actual-table/math ink observers and a real production correction required the final freeze. Failed artifact/evidence retained. Eight existing note PDFs and eight fixed DOCX portfolio PDFs remain byte-identical. Eight Office comparisons/28paired pixel identities reused;2new comparisons/8pairs checked, both p1pairs viewed plus pinned-font p1. All10strict failures/four raw math-word screens retained; source prose/simple math deficits zero. Two Word views exported once, no repeats;17assets/four root smokes, no root build/suite. Root675ad6e5; artifacts/composition-c1/table-note-body-boundaries-evidence.json. C236CI39145256938 is green on Windows/Ubuntu. Workspace scan14.261GB/771,789files remains above the loose5GB target; zero scan errors, rejected retirements preserved.

C237 table-reference holdout (read-only, 2026-10-18): moving only note2's existing label/reference into the report's last ordinary table cell preserves all other package parts and reverses reference order to2/1. Frozen C236 emits6pages in each preset versus Word4: the table header is isolated onp2 and its reference/head arrives onp3 rather thanp1. All64numbered note paragraphs survive once in reference order33-64/1-32, and each automatic link reaches its own head. Both strict visual gates remain failed;8paired/4unmatched candidate rasters pinned and Finalp2pair reviewed. Two Office exports once; no library build or suite. Whole-story table pre-reservation is a suspected cause pending a prototype, not yet a proven fix. Evidence artifacts/composition-c1/table-note-reference-observations.json.

## C236: reuse measured stories within fixed-input placement retry (2026-10-18)

One coordinator-owned width dictionary now spans the two header/footer placement
calls with identical document, font measurers, math plans, geometry and scale.
The355.135/468pt contexts remain separate. Other font/numbering/kerning/page-count
preparations and conversions retain fresh owners; no global cache, new quota or
unsafe/native code. The existing Create ABI remains intact. A source/provenance
and changed-placement count regression measures zero repeated note text calls;
it fails accepted C233 with1,108 calls, rather than missing-member failure.

One frozen2d587d98 artifact qualifies58Windows/42Ubuntu checks,16native skips,
once per platform. Selection covers normal/owned notes plus static/comment/field,
inline-image and concurrent shared-resolver neighbors; no harness repair. All
2,249 input pins remain unchanged. Eight composed-note PDFs and six staged/whole
benchmark outputs retain bytes. Initial font preparation and715,743 kerning-pair
fingerprints remain unchanged. Same C234 isolated1cold/3warm method saves long-note
2.262-2.264MB Final/5.196-5.198MB Review(6.76-6.79%/12.01-12.06%). Calling-thread
allocation is not live heap; diagnostic stages overlap, timing is noisy and no
timing speedup is claimed. The probe is rebuilt separately; library not rebuilt.

All eight Office comparisons,28paired PNG hashes/review chains, failed strict
gates and four raw math-word audit flags carry by whole-PDF/input/reference/
threshold identity. No new export, raster comparison, visual review or audit
replay.17assets/four root smoke PDFs match; no root build/suite or ordinary corpus
replay. Rootbcbfabb4, library2E4AB2C38587DFF3D213E27AA0C35FD844A4EBA7FC7F35034E087CD23F25FC01.
Evidence: related-story-measurement-reuse-evidence.json; docs/OperatingEnvelope.md.
C233CI39142860977 is green on both platforms, including broader/package checks.
C237 returns to the realistic table/cell reference combination.

## C235: related-story measurement ownership across placement retries (2026-10-18)

Read-only C233 attribution confirms a concrete repeated pure operation. The same
source document and prepared resources require a header-placement retry: initial
header bottoms717.676/718.237pt lie below the720pt body top. Each engine.Create
builds a fresh related-story width dictionary, recomputing immutable measurements.
Final uses468pt; Review needs355.135pt body and468pt note widths. The two normal
notes measure130lines at the narrow width and66 at full width, so these distinct
contexts must remain. Standalone factory medians are2.282MB for full width and
2.956MB for the narrow Review width; reflection/PowerShell dispatch is included.
These diagnostic figures are not additive production savings or a speedup claim.
SegoeUI is added during painting, so it does not explain coordinated snapshots.
No library builds, PDF conversions, tests or Office exports replayed.
C236 will share measured related stories only across the two placement calls,
with identical source/font/math/geometry and rounded width keys; other coordinator
retries/conversions retain separate ownership. Evidence:
related-story-measurement-attribution.json and four frozen attribution snapshots.

## C234: compound-note allocation baseline (2026-10-18)

The existing probe alone is built/self-tested against frozen C233; library bytes
remain pinned. Fixed report, single long note and two-note report are measured in
both presets: one isolated process per input/preset, one cold/warmup and three
warm conversions, buffer output/default64MiB resident window. All six whole/
staged PDFs are stable and equal qualified output hashes; no Office export,
visual qualification or library build is replayed. Calling-thread allocation
is not live heap/process memory. Two-note warm medians33.470MB Final/43.262MB
Review; timing is noisy, with no speedup/mode-order claim. Initial font resources
allocate16.656-17.530MB; coordinated layout/snapshot21.807-34.072MB. Those diagnostic
scopes overlap and do not count layout passes. C235 will attribute coordinated
Review preparation before choosing any optimization; no new quota/cache.
Evidence: compound-note-allocation-evidence.json and docs/OperatingEnvelope.md.

## C233: ordered plain-note batches and uncapped source admission (2026-10-18)

The C232 overlap/order defect closes: deferred heads remain owned by the drain,
plain paragraph notes share a bounded continuation batch and one separator, and
source progress/placement remain separate from typography admission. Existing
plain paragraph semantics and non-rendering metadata qualify the batch; keeps,
contextual/unknown spacing, tables, floating/rich content retain explicit existing
fallbacks. Cancellation remains checked. Both presets match Word20/31/13, all64
paragraphs once/in order, without overlapping note containers. The second head
agrees−0.011/+0.241pt; four automatic note links target the correct pages within1pt.

One bounded two-outcome prototype preceded freezing. The first frozen50-method
qualification exposed a real legacy capacity defect: page-capped painting height
could admit all43remaining lines rather than the independently validated17.
Whole-note admission now uses uncapped measured source height in both separator
routes; continued seating retains full after-spacing. A five-method mutable probe
confirmed both legacy take controls, the owner model and native/portable controls.
This production correction required final frozen qualification, not an assertion
change. Initial49Windows/36Ubuntu passes,13native skips and one failure each remain
retained. Final frozen24f3ce2d passes50Windows/37Ubuntu,13native skips, once per
final platform. Three strengthened observers fail accepted C231.

2,246 public pins unchanged/two source manifest/Office JSONs updated/one public
source added, with byte-identical deterministic fixture regeneration. Eight old
Word views reused, no new exports. Six earlier note PDFs retain bytes; six
comparisons reuse PDF/threshold identity and two fresh comparisons keep all eight
strict visual failures. Source prose/simple math deficits are zero; four raw
rate/time Word prose screens remain.28paired PNG hashes/chains retained; Final
page3/Reviewpage2pairs and portablepage4of5 reviewed. FixedC231portfolio outputs
have no pending note tails, leaving the new drain branch unreachable; their
whole source heights already fit, so no fixed portfolio replay is required.

17assets/four root smoke PDFs match, no root build/suite. Library/CLI Release
builds have zero warnings/errors; tests retain two inherited nullable warnings.
Root5986a1b9, library5B3FA19D981EA6976B37655079F1CEFA6D9293FBF2690C045AED3171D3C46CC5.
Evidence: plain-note-batches-capacity-evidence.json. C231CI39140627108 is green on
both platforms, including broader/package checks. C234 measures compound-note costs.

## C232: two plain-note owners reveal overlap and reordered tails (2026-10-18)

A source-only report variant splits64 numbered paragraphs evenly across two
normal note owners, adding a second body reference and one public intro. Mutation
reversal restores every original body/note node; other parts remain unchanged.
Frozen C231 produces both presets without builds/test replay; two new Word views
are exported once. Word retains four pages and numbered membership20/31/13.
Candidate also has four pages/all64 paragraphs once, but paints both note heads
at identical baselines, overlapping20 paragraph pairs and reordering the first
tail after the second note. Review page2 visually confirms unreadable overlap.
Both strict gates fail; eight paired PNG hashes retained, one pair reviewed.
Drain ownership of unplaced heads is still tied to typography admission, while
shared dedicated draining is limited to inline-owned stories. C233 will separate
those contracts for bounded paragraph notes and retain complex-source fallbacks.
Evidence: two-plain-notes-observations.json. C231CI39140627108 is still running,
Ubuntu green; no full-CI success claim yet.

## C231: ordinary note frames and explicit reservation/source ownership (2026-10-18)

Placement accepts a usable measured ordinary margin independently of footer
intrusion and typography. Explicit already-placed source progress prevents a
narrowed remainder at paragraph-local line zero being treated as a fresh head.
The original-footer sibling now matches Word21/32/11 in both presets; rich-footer
controls retain17/31/16. All64 paragraphs survive once/in order and first/last
note/footer origins retain1pt bounds. No font/size/document-shape exception.

C229 CI39139224596 caught three real complete-admission regressions, independently
reproduced on first C231. Placement now receives existing separator reservation
ownership, preventing a second charge. Final frozen8a5edf7a qualifies49 Windows/
36 Ubuntu passes,13 native-font skips, once per final platform. Exact selection
adds four complete-batch/minimum-bookmark neighbors. These independent Word page/
baseline gates are preserved. Production correction required a final affected
qualification; earlier failed artifacts/reports are retained.

The first45-method qualification had44 Windows/33 Ubuntu passes,11 native skips
and one assertion failure each: it compared an unusable-footer fallback with the
corrected ordinary baseline. Read-only prior/current numerical traces prove the
actual unusable fallback unchanged. The assertion now compares two explicit
unusable controls and verifies complete source ownership/typography fallback.
14 runtime assets/all2,248 input pins remain identical through that test-only
repair; only the corrected method reran per platform. That interim result was
superseded by the separate production reservation fix and final49-method result.

2,245 older pins unchanged/two source manifest/Office JSONs updated/one public
source added; deterministic fixture regeneration retains every source byte.
Six existing Word views reused. All six final PDFs match first C231, so all six
comparisons reuse PDF/threshold identity; no new exports or raster comparisons.
Six strict visual failures remain. Source prose/simple math deficits are zero,
with four raw rate/time Word prose screens retained. Twenty paired PNG hashes/
chains retained; two ordinary-footer first-note pairs and one of five portable
pages reviewed. Both strengthened observers fail accepted C229.

Eight fixed DOCX PDFs were measured once at this broader placement milestone;
all retain accepted bytes. The final pure cap-ownership bypass is nonbinding on
those unchanged outputs, requiring no portfolio replay. PPTX remains unchanged.
17 assets/four root smoke PDFs match, no root build/suite. Library/CLI Release
builds have zero warnings/errors; tests retain two inherited nullable warnings.
Root6430a1e0, libraryF68C774CA58682BD3F1B758DA5EB4C43813F76CF9EFB06C5D5E45BDF93C5CC8C.
Evidence: note-reservation-ownership-evidence.json. C232 checks two plain-note owners.

## C230: ordinary-footer sibling exposes placement dependency (2026-10-18)

Removing only the added mixed footer paragraph restores the original report
footer while preserving all64 note paragraphs and all original body/static
content. Frozen C229 renders two modes without builds/suite replay; two new Word
views are exported once. Review notes again overlap body ink, visually confirmed
on page2. Word21/32/11 source membership differs from Final22/33/9 and Review
24/32/8; all64 paragraphs remain once/in order, with four pages in every output.
The shared placement resolver rejects an ordinary bottom-margin frame unless
typography qualifies it, so physical placement still depends on footer intrusion.
Two strict visual gates remain failed; eight paired PNG hashes retained, one
pair reviewed. C231 will admit ordinary measured placement independently of
intrusion/typography and retain focused legacy separator/admission controls.
Evidence: long-note-footer-sibling-observations.json. No font/line-height fit.

## C229: shared physical body boundary seats initial note heads (2026-10-18)

C228's body overlap and initial head-capacity errors close in the admitted
physical frame. Body-to-note coordinate conversion is extracted from inline-
note eligibility and reused for placement; the head reserves its separator once
from placed body bounds. Typography/kerning admission and old inline-note paths
remain. No font/size/document-shape exception or fitted offset is added.
Both presets now match Word17/31/16 source membership and retain all64 paragraphs
once/in order. First-head errors19.425/58.699pt become+0.002/+0.271pt; all note
page first/last origins retain1pt gates (max0.432/0.432pt).

One frozen c0bba51c artifact:37 Windows/29 Ubuntu passes, eight native skips,
once per platform; both strengthened observers fail accepted C227. Two bounded
prototype conversions preceded qualification; no root build/suite replay.
2,246 input pins unchanged; one Office-facts JSON adds independently retained
head origins and page membership. Four existing Word views reused, no exports.
Separator-less PDFs and all equation operations remain identical. Two comparisons
reused/two fresh; all4 strict visual failures remain. Source prose/simple math
deficits are zero, per-page prose membership deficits clear, and four raw Word
rate/time math-word screens remain. Twelve paired PNG hashes/chains retained;
two first-note page pairs and one of six portable pages reviewed.

Existing fixed-source header math retains static typography fallback, and C227's
measured footer frames do not intrude; the new placement-cap branch is unreachable
there. Source SHA proof suffices without portfolio/deck preparations or conversions.
17 assets/four root smoke PDFs match. Library/CLI Release builds have zero
warnings/errors; tests retain two inherited nullable warnings. Root a80f0fe2,
library 08B2EC5BFCE39BC3F2F623A9FFE22C4D57BCA6D60C2E03BEDFC7F1BB1B789230.
Evidence: note-body-boundaries-evidence.json. C227 CI39138298794 is green on both
platforms, including broader/package checks. C230 checks the footer-free sibling.

## C228: first-note body collision diagnosis (2026-10-18)

Retained C227/Word PDFs plus two read-only prepared layouts expose a confirmed
Review body/note overlap. The first note baseline lies28.386pt before the final
body baseline; Word places it28.850pt below. Body baselines themselves differ
only1.463pt. Final head remains one extra paragraph high. These are placement/
capacity defects, not missing content or evidence for a font/line-pitch fit.
The existing body-to-note conversion is gated by owned-inline note eligibility;
plain multi-paragraph notes keep a scaled body cursor in note design space.
Missing separator charge is an additional hypothesis requiring a bounded probe.
No production builds, conversions, exports, tests or rasters replayed.
C229 will share physical boundary ownership while retaining typography admission.
Evidence: continued-footer-head-review.json and two head-frame snapshots.

## C227: continued-note occupancy and fallback endnote clearance (2026-10-18)

Physical footer availability now reaches the normal unqualified continuation
drain, separately from typography/kerning admission. The inherited continuation
collision closes: all64 source paragraphs survive exactly once/in order, and
last-note/footer origins agree with Word within0.433pt on every note page.
Middle-page capacity now matches Word31 paragraphs. First-page membership still
differs: Word17/31/16 vs Final18/31/15 and Review21/31/12. Initial head errors
19.425/58.699pt remain; no font/size/document-shape exception or fitted line pitch.

C225 CI39136390437 caught a real neighboring regression: raising a footnote
exposed apparent fallback space underneath it for an unqualified mixed endnote.
Fallback endnotes now clear complete occupied footnote boxes. Partial-head and
typography qualifications retain their existing contracts. The existing CI guard
fails C225 and passes C227; both new native/portable continuation observers fail
C225 ink clearance. One parent-runtime missing-fixture failure is retained; only
that guard reran after adding the sixth explicit input. No production repair of
an assertion, harness rebuild, threshold change or qualification replay.

One frozen eba3c1eb artifact:37 Windows/29 Ubuntu passes, eight native-font skips,
once per platform. Library/CLI Release builds have zero warnings/errors; tests
retain two inherited nullable warnings. All2,243 old public pins unchanged/four
added, with deterministic byte-identical fixture regeneration. Four C226 Word
exports reused; two unchanged separator-less PDFs reuse comparisons, two fresh
comparisons retain all four strict failures. Equation operations remain identical.
Source prose/simple math deficits are zero; four raw Word prose screens retain
rate/time math-word and page-membership flags. Twelve paired PNG hashes/chains,
two new continuation pairs reviewed; one of six portable pages reviewed.

Four fixed DOCX sources have no endnotes or intruding footer boundaries in either
preset, proving the new branches ineffective there; four Final layout preparations
and prior Review frames suffice, without portfolio/deck conversion replay.
17 assets/four root smoke PDFs match; no root build/suite. Root fa792a9c, library
B4074E815678F61D206B6C687832EFE1B074220C664560E8BC7887070DF49C47.
Evidence: continued-footer-notes-evidence.json. C228 traces initial head capacity.

## C226: separator-less and continued-note footer holdouts (2026-10-18)

Four source-only outcomes reuse frozen C225, with four new Word exports once;
no production build or test replay. Removing the separator retains note baselines
within0.433pt of Word. The long note retains all64 numbered paragraphs once/in
order and four pages in both modes, but continuation pages3/4 visibly overlap the
footer. Accepted C223 reproduces that collision, establishing inherited debt.
Word numbered-paragraph membership17/31/16 differs from Final18/33/13 and Review
21/32/11. Existing line pitch is close; no font/line-height fit is supported.
Initial head errors−19.425/−58.699pt and capacity causes remain unverified.
All four strict gates remain failed. Twelve paired PNG hashes retained; four
page pairs reviewed. Evidence: occupied-footer-continuation-evidence.json.

C225 CI39136390437 has Ubuntu green and one Windows mixed-endnote fallback
assertion failure; its cause is pending diagnosis, not dismissed as stale.
C227 prioritizes measured continuation occupancy and adjacent mixed-note guards.

## C225: measured footer occupancy seats notes and separators (2026-10-18)

C224's confirmed overlap is closed. Placement uses an intruding measured footer
boundary when it leaves a usable frame, independently of note typography/reflow
admission. Existing eligibility flags remain, including sliced continuations;
rich complementary/header guards still pass. The separator consumes the same
physical placement height. No font/size/document-shape exception or quota added.
Footer-note baseline errors26.466/20.528pt become−0.016/+0.432pt; independent
separator Y errors+0.127/−0.331pt. All four note/footer/rule1pt gates retain bounds.

Final frozen41b72c27:31 Windows/24 Ubuntu checks, seven native skips, once per
platform. Two final observers fail accepted C223 ink clearance. Visual review of
the first build exposed a separator-height error despite its31/24 test passes;
that production correction added a Word rectangle observer and required final
affected-set qualification. The new observer fails the first build's rule.
Earlier evidence retained; no root/corpus replay. A pre-execution selector fix
changed only the tracked capability manifest, with17 runtime assets/all2,243 pins
unchanged and no rebuild. Final2,239 old pins unchanged/four public inputs added.

All four C224 Word views reused, no exports. The note-story sibling PDFs retain
bytes; all equation operations/96 added glyph anchors remain unchanged. Two
comparisons reused/two fresh, all four strict failures retained. Eight paired PNG
hashes/reuse chains retained: two corrected page2pairs viewed, two review identities
carried, first pages retained without a review claim. Both portable review pages
viewed; note/footer ink is disjoint. Content audit has zero source-prose/simple-
math deficits; four raw Office Arial rate/time word flags remain. Font/nonblank
operation/source/reference identity carries the audit; only whitespace separator
anchor matrices change. Four measured fixed-portfolio footer boundaries exclude
the new intrusion path, avoiding unrelated corpus/deck conversions or suites.

17 assets/four smokes match on transfer, no root build/suite. Release library/CLI
are warning-free; two inherited nullable test warnings remain. Root f4294258;
occupied-footer-notes-evidence.json and occupied-footer-notes-transfer/. C223
CI39133625732 is green on Windows/Ubuntu. Next C226 checks two bounded omitted-
separator/continued-plain-note variants against the same frozen artifact and
independent Word views before choosing any further correction.

## C224: linked-inline footer/note holdouts expose footer reservation gap (2026-10-18)

Two source-only report transforms append a short linked equation/image paragraph
to the existing referenced footer or normal footnote, retaining all original
content and other package parts. Four frozen C223 outputs/Word views keep two
pages, expected prose and every equation: footer64math glyphs, note48.96 added
math glyph anchors agree within0.215ptX/0.417ptY/0.007pt size. Added story links
reach the existing body bookmark within0.276pt. Raw token floors retain a shared
note/note1 automatic-reference adjacency flag in Word/candidate; not relabeled.

Confirmed defect: expanded footer painting agrees with Word, but footnote text
ignores its occupied boundary. It sits26.466ptFinal/20.528ptReview too low and
visibly overlaps footer prose; font boxes intersect1084/624pt2 versus zero Word.
The automatic note destination follows that misplaced note. The note-story
sibling stays−0.028ptFinal/+0.407ptReview from Word, preserving its image/math.

All four strict gates remain failed. Eight pixel pairs hash-pinned; four complete
second-page pairs reviewed, first pages retained without a review claim. Four
new Word exports retained once; no production build/suite replay. Evidence:
linked-inline-story-holdouts-evidence.json and its source/PDF/note-footer geometry
pins. Next C225 traces shared measured footer space into note reservation and
placement, retaining validated ordinary/static/minimum/fixed frames and limits.

## C223: printed font ownership in hyperlink line tiling (2026-10-18)

Confirmed C222 review-cell annotation gap is closed. The next-line legacy top
used source font size beside placed baselines and printed current glyph size.
Shared tiling now applies the existing print scale once to that next-line size;
no font/point/container exception. All380 previously excluded source character
baseline origins are covered. Review-cell Word vertical rectangle edge delta
falls to0.052pt; horizontal trailing-space bounds remain unresolved. Body review
keeps its0.874pt vertical residual and byte-identical PDF. Both Final PDFs also
remain identical; the fourth PDF changes only annotation rectangle arrays.

Frozen production d0e2e673, final harness df36b777.20 Windows/15 Ubuntu checks,
five native skips. Initial Windows result19pass/1new portable harness failure:
it incorrectly required Word's line count despite pinned-font wrapping. Exact
native Word counts/gates stay unchanged; only portable count was corrected.
14 production/runtime assets and2,239 public pins unchanged; only that corrected
Windows entry rerun. Ubuntu ran the complete exact set once. Two observers fail
C221 using five independently copied inputs.2,235 old pins unchanged/four inputs
added. Release builds pass; two inherited nullable test warnings, library/CLI clean.

Four cached C222 Office views reused, no exports. All text/font/graphics JSON
identities remain exact, carrying64 Word math anchors and source/image checks.
Three comparisons reused/one fresh; all four candidate/Word PNG pairs equal their
reviewed C222 hashes. All strict failures remain. One portable review-cell page
viewed; its sixteen lines differ from Word's twelve. The fixed four DOCX source
links all take unchanged qualified body boxes, proved by four layout preparations;
no ordinary portfolio replay.17 assets/four smokes match after transfer, without
root build/suite replay. Root bcfc8487; printed-link-frames-evidence.json and
printed-link-frames-transfer/. Next: two bounded related/static-story linked-inline
holdouts on this artifact, retaining independent Office geometry and all limits.

## C222: justified linked-picture body/cell holdouts (2026-10-18)

Two source-only heldouts combine wrapped justified prose, an unchanged image,
paired fractions and an internal source hyperlink/bookmark. All four C221 frozen
outcomes retain16 math glyphs/one illustration/full main-document source prose,
with the same one-page pagination as Word.64 independent glyph origins pass the
existing1pt/.1pt gates (maxX0.891pt/Y0.170pt/size0.011pt); image edges differ by
at most0.749pt. All four complete native/Word pixel pairs reviewed and hash-pinned.

Both PDFs retain6body/12cell links and page0 destinations. Review-cell clickable
rectangles end up to0.154pt above380of456 source character baseline origins.
This demonstrates an annotation-frame inconsistency, not missing ink or proof
that those glyphs cannot be clicked. Left-aligned C219/C221 controls have identical
PDFs and repeat the gap: inherited behavior. Horizontal link edges also retain
up to2.710pt differences; trailing wrap-space ownership remains unresolved.

All four strict gates remain failed. A first prose screen accidentally joined
paragraphs into souththe/southobserved in both outputs; raw evidence retained.
The corrected paragraph-aware source screen has zero deficits. Four new Word
exports retained; four heldout conversions and four small parent/current left
controls used existing artifacts. No production builds or suite replays.
Evidence: justified-composition-holdouts-evidence.json and its pinned math,
link-origin-coverage and left-parent-identity reports. C221 CI39132266264 is green
on Windows/Ubuntu, including broader/package checks. Next C223 traces shared
hyperlink next-line font coordinates, retaining body/final contracts and all
cached Word views before choosing a bounded correction.

## C221: shared ordered justification preserves mixed formulas (2026-10-18)

Confirmed C220 omissions are closed: four justified report outcomes restore two
formulas/nine glyphs each (36 total), including wrapped prose. Shared placement
stretches retained ASCII spaces after line breaking; intrinsic equation/image
advances, source ranges, display alignment and tab positions retain ownership.
The existing plain-text span emitter is shared rather than duplicated. Final
paragraph lines remain natural; legacy single-equation justification keeps its
qualified placement path. No font/size/document-shape exception was added.

One committed Release artifact54a83c09:16 Windows checks/10 Ubuntu checks,
six native skips, once/platform, no repairs. Three new observers fail C219.
All2,230 old input pins remain unchanged; five public inputs added. Six qualified
PDFs equal the prototype; left controls retain C219 bytes.54 independent Word
target glyph anchors retain existing1pt/.1pt gates: maxX0.187pt/Y0.843pt/size0.04pt.
All six C220 Word exports reused; no new exports or unrelated corpus/deck replay.
Source inspection excludes newly supported justification from the fixed four
DOCX. The extracted ordinary-text emitter and preserved legacy path retain
focused text/image/line-start/display/static/fraction neighbors.

All six strict gates remain failed. Content audit has zero source-prose/simple-
math deficits; four raw Office-word flags retain Word's source/sourc/e PDF run
split. They are not relabeled as passes.12 native pixel pairs retain both PNG
hashes/reuse chains; six pairs reviewed now, plus both pages of a portable-font
review PDF. Font-dependent portable wrapping and existing comment/image/table
geometry differences remain.17 assets/four smokes match on transfer, without a
root build/suite replay. Root b2ddbe91; inline-justification-evidence.json,
inline-justification-office-observations.json and inline-justification-transfer/.

Next: bounded source-only justified image/link/container holdouts on this frozen
artifact; characterize content, pagination and independent geometry before
choosing another correction. No broader support claim follows from these probes.

C220 confirms that paragraph justification still breaks supported composition.
A single-line or wrapped report paragraph containing two supported formulas
omits both in each public preset: nine glyphs/outcome,36 missing overall.
The wrapped left-aligned sibling retains all32 document math glyphs and prose;
all six independent Word outputs retain32 glyphs/two pages. The converter
succeeds with an unsupported/ignored diagnostic for each justified outcome.

Three source transformations preserve all formulas/fonts/story owners, with
ordinary public prose added only for wrapped controls. Six Word exports are
cached once. Frozen C219 supplies all measurements without production build or
test replay; all six strict gates remain failed. C221 should reuse the existing
justification contract in ordered placement, including continuation/last-line
ownership and source geometry. No new single-line/font/size/source-shape guard.
Evidence: justified-math-composition-review.json. C219 CI39130228459 is green
on both platforms, including broader/package checks.

C219 restores64 missing glyphs caused by unused tab-stop metadata. Shared
resolution retains the source leader; ordered composition validates the stops
actually consumed, including a line retry. Bars and unknown kinds keep the
paragraph-owner fallback because they may draw independently. Active decimal/
leader stops, overlapping aligned fields and out-of-frame cases stay staged.
No font/size/document guard, new rendering framework, quota or unsafe/native code.

One frozen Release artifact qualifies10 Windows/6 Ubuntu checks, four native
skips, one exact invocation/platform and no repair. Three observers fail C216.
Unused forward, behind-cursor and cleared stops retain complete measured geometry
and prepared owners at two scales. Existing wrap, active fallback, cancellation,
right/center fields, source cursors and link neighbors qualify. Independent Word
field offsets retain their0.2pt gates; pinned portable content runs on both hosts.
All2,226 old input pins stay unchanged and four public inputs are added.

All eight frozen PDFs equal their original supported controls. Four comparisons
reuse identity; four new paired pages are reviewed and all eight pairs retain
both PNG hashes. Four C218 Word exports are reused without further export. All
eight strict visual gates remain failed; the Word cell-frame shift0.84/about
0.65pt remains unverified and receives no fitted correction. Fixed ordinary
sources contain no inline tabs, excluding an unrelated corpus/deck replay.

Root eda06270; consumed-tab-ownership-evidence.json. Transfer verifies17 assets/
four smoke PDFs without root build or suite replay. C216 CI39127757053 is green
on both platforms. Push and continue C220's bounded justified mixed-content
review against the frozen artifact while CI runs. Alignment remains an
architectural concern until a source transformation demonstrates content loss.

C218 confirms another supported-feature composition omission on frozen C216.
An unused720pt decimal stop in a body paragraph or dotted-leader stop in a cell
removes both supported formulas in both presets:64 missing glyphs. Original
left-tab controls retain16 glyphs/preset; source fields/content are unchanged.
Office retains the formulas. Body raster hashes are identical; cell inline
field offsets/Y are unchanged while its whole table shifts0.84pt Final/about
0.65pt Review against the older cache. That frame difference remains unverified
and receives no fitted correction. Four new Word exports are cached once.

A C219 prototype validates consumed positioning stops and preserves their source
leader metadata. Bars/unknown kinds keep the owner fallback; active decimal or
leader tabs remain unsupported. All four transformed PDFs recover16 glyphs and
are byte-identical to their original supported control. Final focused frozen
qualification remains pending. Evidence: unused-tab-math-composition-review.json.
C216 CI39127757053 is green on both platforms, including broader/package checks.

C217 reuses frozen C216 for two held-out exact-slot combinations: a report
paragraph containing math and an inline picture, and a wrapped footnote. Both
presets retain32 math glyphs, all550/61 source-prose tokens and Word page counts
(two/one). Four new Word exports are cached once. No production/test change,
build, suite replay or whole-corpus qualification.

Note math groups remain withinX0.171pt/Y0.399pt; note headY0.305pt and following
bodyY0.128pt. The picture stays visible with colored-pixel edges within1pt.
One neighboring table denominator retainsY1.014pt, narrowly failing the unchanged
one-point gate; all four strict visual gates remain failed. Four paired target
pages are reviewed and all six raster-pair hashes retained. Token counts do not
establish reading order or clipping. No fitted correction is accepted.
Evidence: exact-slot-holdouts-evidence.json. Continue C218's bounded unused-tab
metadata review; treat content-loss risk as unverified until demonstrated.

C216 restores104 missing math glyphs when paragraphs use exact line spacing.
Finite positive fixed slots now share ordered inline composition, their authored
physical advance and the existing exact text baseline. Measured ink may overhang
without enlarging the slot. Body, cell and static containers keep their placement
owners; automatic/minimum, zero/invalid and unsupported source fallbacks remain.
No font/size/document-shape correction, extra quota or unsafe/native library code.

Five minimal public transformations retain source equations/fonts. Eight body/
cell/header outcomes restore96 glyphs; both composed-report presets restore four
glyphs each and keep two pages, images, tables, header/footer, notes and markup.
All104 independent Word origins pass existing one-point gates: maximumX0.546pt/
Y0.1973pt. Physical font-size gates remain unchanged. Compact12pt slots retain
Word's overlapping ink;72pt slots fit. Ten C215 exports stay cached with their
source/view/PDF/log pins; no further Word export.

One frozen Release production artifact qualifies22 Windows/13 Ubuntu checks,
nine native skips, one initial exact invocation/platform. Three new observers
fail C214. The new native report observer first assumed an unsplit prefix, then
compared raw Tf sizes with scaled Word sizes. Two test-only corrections read the
complete rendered line and apply the existing print map. Only that native entry
reruns twice;14 production/runtime assets and2,226 public pins stay unchanged.
All initial/failing reports remain. Its portable path and native-platform skip
are unchanged, so no suite replay. Builds pass with two inherited test warnings.

All ten frozen PDFs equal the reviewed prototype. Twelve reviewed Office pairs
reuse both PNG hashes; four portable pages retain supported content. All ten
strict visual gates remain failed; eight complex-math audit screens remain
unresolved. Source-prose/simple-math/Office-word deficits are zero in all ten.
The fixed ordinary portfolio has no exact math paragraphs or exact style defaults,
so source eligibility excludes another ordinary-corpus/PPTX replay. Transfer
verifies17 assets/four smoke PDFs without root build/suite replay.

Root d51d7a19/observer4d5bfdfb; exact-line-math-evidence.json. C214
CI39126058981 is green on both platforms. Push this coherent batch and continue
C217 read-only exact-slot picture/note composition holdouts against its frozen
artifact while CI runs. Change production only after a concrete defect appears.

C215 confirms a visible composition defect using frozen C214. Changing only
three paragraph-spacing attributes from atLeast to exact makes all three
supported equations disappear in body, cell and header controls. Both presets
succeed with an unsupported/ignored diagnostic and retain prose. Eight outcomes
lose12 math glyphs each. A composed report with one48pt exact paragraph loses
four of32 glyphs in both presets. Fonts, sizes and source formulas are unchanged.

Ten new Word exports are cached once with source/view/PDF/log pins. Word keeps
all formulas:72pt slots fit, while12pt slots allow tall ink to overlap neighbors.
This is paragraph line ownership, distinct from previously qualified exact-row
clipping. A provisional shared fixed-slot prototype restores104 missing glyphs;
eight ordinary controls have maximum independent mathX0.546pt/Y0.096pt. No
font/size/document-shape fitting or threshold relaxation. C216 must qualify the
final artifact and preserve the held-out report's scaled composition.
Evidence: exact-line-math-composition-review.json. C214 CI39126058981 is green
on both platforms, including broader/package checks.

C214 removes the confirmed full-TTC copy before selected-face extraction.
The shared source stays read-only; extraction creates the privately owned SFNT
directly. Standalone fonts retain defensive copies and all parsers/subsetters
retain independent buffers. Shared directory bounds, cache retention, fallbacks
and cancellation remain; table copies now observe cancellation. No parsed-font
cache, additional quota, dependency or unsafe/native library code.

One frozen Release artifact qualifies33 Windows/31 Ubuntu checks, two native
skips, one exact invocation/platform and no repair. All2,218 inputs retain pins.
Three new portable observers cover sliced sources, concurrent independent
ownership, byte-exact faces, unused collection padding, malformed collection
extents and cancellation. The allocation observer fails C212 at4,204,672bytes.
The original Ubuntu invocation used an unsupported JSON-report option: retain
its observed complete31pass/0fail/2skip stdout result without suite replay.
Builds pass; the unchanged test file retains two existing nullable warnings.

Every one of eight fixed DOCX warm medians saves1,792,688bytes/conversion:
5.53-7.60percent, matching the avoided1,792,664-byte native TTC plus array
header. All eight whole/staged PDFs and715,743 pair entries/fingerprints across
three native faces retain identity. Timing stays observational; no speedup or
live-heap ceiling is claimed and preparation diagnostics remain non-additive.
Both fixed PPTX PDFs and passing visual gates retain identity. Fourteen reviewed
DOCX pairs/two deck pairs carry with both PNG hashes and comparison reuse chains;
all eight DOCX strict failures remain. No new Office export or raster.

Root06e8183a; collection-font-ownership-evidence.json. Transfer verifies17 assets
and four smoke PDFs without root build, suite or full corpus replay. C212
CI39124225988 is green. Push the bounded ownership batch and resume rendering
composition work while CI runs; no additional speculative font cache.

C213 records the four fixed mixed DOCX sources in Final/preserved and
AllMarkup/Word-compatible without rebuilding the library. One process-cold
conversion and three warm samples per input/preset use the existing isolated
allocation probe, Windows/.NET10.0.13, default466-face resolver and buffer output.
The default64MiB resident window stays unchanged. All eight whole/staged outputs
are stable and equal the previously qualified PDFs; fourteen reviewed Office/
candidate pairs and all eight failed strict gates carry by identity. No new
Office export, raster, suite replay or library build.

| Source (pages) | Final warm MB / ms | AllMarkup warm MB / ms |
|---|---:|---:|
| Report (2) |26.319 /104.14 |29.634 /120.34 |
| Letter (1) |23.603 /66.74 |26.852 /77.05 |
| Specification (2) |29.065 /165.07 |32.410 /140.06 |
| Held-out memo (2) |25.119 /101.54 |28.302 /95.58 |

MB is decimal; values are medians of three warm samples and measure allocation
on the calling thread, not retained heap. Timing is noisy and establishes no
speedup or mode ordering. Standalone initial font resources allocate18.171-
18.750MB; inclusive preparation diagnostics overlap and must not be added.
Independent SFNT reparse probes do not establish repeated production parsing.
Font loading copies a whole TTC before extracting its independently owned face:
C214 will test direct read-only extraction while preserving defensive ownership,
bounds/cancellation and existing bounded source caches. No shared parsed-font
cache or new quota. Evidence: compound-allocation-baseline-evidence.json.
C212 CI39124225988 is green on both platforms, including broader/package checks.

C212 restores full default continuation rules when the package omits the
continuation story. Source progress distinguishes a continued note from a fresh
or still-unplaced head. Reuse the existing authored-continuation design width
and one print map for both sliced and complete tails; retain short head rules,
explicit separator behavior and unqualified fallbacks. No font/size/source-shape
qualification, layout adjustment or new quota.

The old144/109.272pt rules replace Word's468.07/354.98pt continuation widths.
Existing cached references now pass46 independent one-point rule gates:
maxX0.036pt/Y0.263pt/width0.155pt/height0.01304pt. Ten rectangles strengthen an
existing gold contract;2,217 other input pins and all source/reference documents
remain unchanged. No new Office export. One frozen Release artifact qualifies
18 focused Windows/14 Ubuntu checks, four native skips, one exact invocation/
platform and no repair. Both strengthened observers fail C211.

All twenty focused outcomes retain identical complete decoded text operations.
Their graphics lists differ only by24 rule-width extensions, including identical
clips. Two fitting PDFs/comparisons and both raster hashes reuse identity.
The fixed DOCX portfolio's four sources author their separators explicitly, so
none can enter the default-only branch; avoid another ordinary-corpus/PPTX replay.
Twelve paired native rule bands include fresh second-note heads, and six portable
pages cover automatic/minimum/image/footer combinations. All twenty strict gates
remain failed; some foreground-recall/changed-pixel components pass. Existing
extraction and cosmetic residuals remain, with no threshold change.

Transfer verifies17 assets/four smoke PDFs without root build or suite replay.
Root a2608a43; default-note-rules-evidence.json. C211 CI39122616601 is green on
both platforms, including broader/package checks. Continue C213's bounded mixed-
document allocation baseline against this frozen library while C212 CI runs;
inclusive stage diagnostics must not be added together or treated as a measured
speedup. Choose optimization only after the portfolio measurements.

C211 makes exact body comment geometry consume the already measured source
box. Background expands or compresses to that box; connector bottom and balloon
top have distinct placement owners. Glyph brackets retain their text attachment.
Complete previous-source-run ending ownership now covers body paragraphs as well
as cells. Finite/source frame checks preserve content-driven inline and unmeasured
static/cell fallbacks. No font, point-size, source-shape or fitted-offset rule.

Three existing Word groups cover a tall exact paragraph, compact multi-line body
text and a table neighbor. Their former11.625pt selection heights now follow the
60.707/8.717pt source boxes; maximum independent selection-edge error0.124pt and
balloon-title error0.557pt. No new document or Office export;2,217 prior pins stay
unchanged and one gold contract is added. Four legacy standard-sheet reference
identities are imported from verified source/PDF/export-log pins; original PDFs,
logs and rasters stay unchanged.

Raster review rejected the first production prototype because authored spacing
also stretched its glyph brackets. A new80-to120pt transformation proves those
brackets retain their text geometry while background changes; it fails the rejected
artifact. The final corrected frozen Release artifact qualifies32 Windows/25
Ubuntu checks, seven native skips, one exact invocation/platform and no final
harness repair. Two source-frame observers fail C210. An initial native whitespace
failure reproduces on C210 under Linux: declare that entry's native-font platform
precondition, keeping its Windows assertions and pinned both-platform coverage.
Its test-only repair preserves14 production/runtime assets and2,218 public pins.
Earlier rejected qualification and visual evidence remains intact.

Eight fixed DOCX portfolio PDFs retain identity. Thirteen comparisons/fourteen
original reviewed pairs reuse PDF/threshold and both PNG hashes. Eight control
content screens reuse complete identical decoded text-operation bytes; source-
prose/simple-math deficits remain zero, with four rate/time word-extraction flags.
Three changed native bands/two native Final page pairs/four portable pages are
reviewed. All eight control strict gates remain failed; both compact Review
foreground-recall gates pass. X/width/bracket style and approximate compact note/
separator origins remain. Transfer verifies17 assets/four smoke PDFs, no root
build/suite replay. Root6f7b280c/2411b860; exact-comment-final-evidence.json.
C210 CI39120537050 is green on both platforms, including broader/package checks.
Continue C212 default continuation-rule semantics while C211 CI runs.

C210 closes the inherited Review body/note page ownership failure. Paragraph
admission rejected a line against the printed boundary, but the note-drain
predicate and retry loop rechecked the less restrictive legacy body boundary.
The progress predicate now owns pending-note progress only; a retry receives the
same live container boundary, including final after-space. Generic image/break
capacities, cancellation and no-progress guards remain. No point/font/document
qualification or fitted offset is added.

Using the same four cached Word exports, the formerly unresolved body moves to
page3 (within0.083pt) and the complete note tail stays onpage2 (within0.375pt).
All four body/note/math/footer origin groups now pass the unchanged one-point
gates. Reference PDF pins and coordinates are unchanged; both formerly unresolved
anchors and their page2 row count become mandatory. The strengthened observer
fails C209. One frozen Release artifact qualifies39 Windows/28 Ubuntu checks,
11 native skips, one exact invocation/platform, no harness correction.2,216 input
pins stay unchanged; only that gold contract is strengthened, with no new exports.

Twenty-two relevant footnote neighbor PDFs and three controls retain identity;
25 comparisons/36 original reviewed pairs are reused with raster hashes verified.
Unrelated PPTX/plain-static inputs are not reconverted. The portable controlPDF
also retains identity and its three reviewed pages. Two changed native bands are
reviewed. Source-prose/simple-math deficits are zero; four raw rate/time extraction
flags remain. The fresh strict pixel comparison still fails. Continued rules are
too short, and an exact-paragraph comment range/balloon still uses the wrong
vertical extent. Those remain explicit follow-ups, not accepted geometry.

Transfer verifies17 assets/four smoke PDFs without root build or suite replay.
Root ac9e5429; owned-note-admission-evidence.json. Periodic scan reports13.527GB/
726,781files, zero errors; loose5GB cap remains unmet and protected retirement
targets are untouched. C209 CI39119744730 is green on both platforms, including broader/package checks.
Push C210 and continue source-backed comment geometry review while its CI runs.

C209 makes intermediate owned-note pages use the same design/footer frame,
paragraph slicer and bottom seating as heads and final tails. Normalize that
frame once, including separator-less dispatch. Synthetic head admission consumes
a measured natural font line instead of an authored minimum or mathematical box;
it constrains the actual body bottom. Finite metrics, source ownership and single
columns govern the shared path. Exact/rich/multicolumn/unknown-provider fallbacks,
bounds, cancellation and no-progress checks remain. No fitted font/size/shift rule.

C208's unchanged long/multiple-note holdouts retain all content and Word's three/
four pages, but exposed intermediate overlap and a 52.790 pt high tail. C209
matches their page2 row counts:41/41 for the long note and40/40 for multiple notes,
both presets; first/last origins differ by less than0.4 pt. Two minimal controls
are promoted byte-for-byte with four independent Word exports once. Three complete
body/tail origin groups pass unchanged one-point gates. The fourth Review group
passes head/footer/math origins but retains an inherited body-admission defect:
its exact-height body enters page2 where Word reserves that page for the note;
the candidate's last three note rows then move to page3. C207 also fails this
control. Keep its independent unresolved anchors and raw failed comparison.

One frozen Release production artifact qualifies39 Windows/28 Ubuntu checks,
with11 native-font skips and one exact capability invocation/platform. The first
Windows run passed38 and failed the new row observer because it counted three
whitespace-only footer shows. Only that assertion is corrected and rerun. All14
production/runtime assets and2,217 input pins remain unchanged; the three test
assets are replaced. Both new observers fail C207. The2,213 prior input pins stay
unchanged/four added. No production rebuild or root suite replay.

Twenty-four neighbor/compound PDFs and their24 comparisons/42 reviewed raster
pairs retain identity, including both original PNG hashes. Twelve new native
paired bands and a seven-page portable overview confirm clearance; glyph-bounds
assertions cover both modes. Four source-prose/simple-math screens have no deficits
or unresolved features. Four raw Office-word screens retain rate/time extraction
flags, while the full32-glyph mathematical stream is present. All eight new/control
and held-out strict pixel gates remain failed. Continuation rules remain too short;
Review body/note admission is next. Transfer verifies17 assets/four identical smoke
PDFs. Root dd026ebd/813b4d3a; owned-note-containers-evidence.json. C207 CI39116871142
is green on both platforms. C209 CI runs after the scoped push; continue work.

C207 closes the retained portable note/footer overlap through a shared metric
family rule. The resolved face requests typographic metrics, but minimum static
baselines subtracted its much larger Windows descender. Use the existing provider
preference and complete finite typographic height/inset to derive descent, keeping
the Windows fallback. Text and image planning share it. The measured occupied
minimum box now also supplies footer admission and maps once for painting.
Zero/exact/unknown/incomplete-metric and unmapped-frame fallbacks remain. No font
name, numeric fitted offset or document-specific qualification.

Two exact public controls change only the declared face to an Office-installed
face requesting typographic metrics. Four Word exports run once. Their former
13-18 pt errors fall within 0.458 pt for static baselines/0.14 pt for body origins;
all four unchanged one-point groups pass. The portable note/footer baseline gap
increases from 5.134 to 17.655 pt. A glyph-bounds regression checks actual prose
ink separation in both presets and body/empty-tail cases without rewriting authored
minima; four complete portable pages confirm clearance. Earlier failed portable
evidence remains. All four new source/Office-word screens pass without deficits
or unresolved features. All four new foreground/SSIM gates remain failed.

One frozen Release artifact qualifies 39 Windows/25 Ubuntu checks with 14 native
font skips, one exact invocation/platform and no harness repair. Three new or
strengthened observers fail C206: Office baselines, glyph clearance and retained
occupied-frame ownership. All 2,209 prior input pins remain/four added. Twenty
neighboring PDFs/four compound outputs retain bytes, carrying 24 comparisons and
42 reviewed pairs with both raster hashes. Four native static bands and four
portable pages reviewed. Seventeen transferred assets/four smokes match; no root
build or suite replay. Root792b53c3; typographic-static-minimums-evidence.json.
C206 CI39115065520 is green on Windows/Ubuntu. Continue long/multiple owned-note
holdouts on frozen C207 while CI runs.

C206 makes owned footnote tails use the same bottom-anchored container as body
pages. Removing only the following body paragraph previously raised the dedicated
note tail 528.115 pt in Final/400.32 pt in review and hid all three review footer
rows below paper. Static continuation stories now reuse an existing document body
origin. Notes also recover their design footer boundary from the existing page
map: body edge 90.212 pt versus note edge 96 pt predicts 4.392 pt of the earlier
review displacement. One cached footer layout supplies both boundaries. Existing
source/typography/fixed/multicol guards, cancellation and progress fallback remain;
no fitted offset or font/size/document-specific rule.

Four new Word exports run once/four compound exports are reused. Two exact public
promotions preserve all 2,205 old pins/four added. All four body/empty-tail Word
anchor groups pass (math first glyph X0.171/Y0.389 pt; prose/footer Y0.422 pt).
All 32 math glyphs and Word's two pages remain per control; the image/math footer
holdout closes its former note offset. Four raw rate/time token screens retain
flags despite complete mathematical streams and zero prose/simple-math deficits.
All four new strict visual gates and all four compound gates remain failed.

One frozen Release artifact qualifies 61 Windows/43 Ubuntu checks with 18 native
font skips, one exact invocation/platform. Both new observers fail C205. A missing
diagnostics import prevented initial test compilation; corrected before freezing
or running any checks, with the failed build log retained. No post-freeze repair.
Eighteen neighboring PDFs/three compound PDFs retain bytes; 21 comparisons and
38 prior reviewed pairs carry with both raster hashes. Ten native note/footer
bands were reviewed. Four complete portable pages retain a confirmed overlap of
note prose and the first footer line under the pinned math font; C205 reproduces
worse clearance. Presence tests do not establish clearance. Keep this failure
explicit and prioritize consistent natural/static metric families and occupied
bounds in C207. Evidence note-footer-frame-portable-residual.json.

Seventeen transferred assets/four smokes match; no root build/suite replay.
Rootef9c24ee/3ef02e91; note-footer-frames-evidence.json and its Word observations.
C205 CI39113410819 is green on Windows/Ubuntu. Continue C207 while C206 CI runs.

C205 restores visible minimum-height static prose through the existing shared
printed frame. Plain atLeast paragraphs were excluded from that frame, unlike
supported minimum-height mathematical paragraphs. One eligibility predicate now
uses content-driven line flow; exact/unknown/malformed rules, source ownership,
alignment/image guards and the existing transform remain. No fitted shift or
font/size/document-specific branch. Separate reproducible plain header/footer
controls establish the correction; compound image/math note inputs stay unchanged.

Four independent Word exports run once; four existing compound exports are reused.
All four standalone one-point static/body anchor groups pass (maximum static
0.436 pt, body 0.14 pt). Three standalone footer rows and all six repeated footer
rows in the review compound are restored onto the page. Source-prose/Office-word
screens pass all four new outcomes with no unresolved features. All four new
strict visual gates and all four compound gates remain failed. The compound note
is still 4.753/4.737 pt too low at first/last baselines; footer baselines now agree
within 0.436 pt. That note-versus-footer frame concern remains separate and open.

One committed frozen Release artifact qualifies 28 Windows/18 Ubuntu checks with
ten native-font skips, one exact invocation/platform. All three new observers
fail C203: portable visibility, native body admission and complete measured-frame
ownership. No harness repair. All 2,201 prior input pins remain/four added.
Eighteen neighboring PDFs and three compound PDFs retain bytes; 21 comparisons
and 38 reviewed pairs carry with both raster hashes/reuse chains pinned. Six
changed native static/note bands and two complete portable pages were reviewed.
Seventeen transferred assets/four smokes match, with no root build or suite replay.
Root0871b780; minimum-static-frames-evidence.json. C203 CI39112468677 is green on
Windows and Ubuntu. Continue bounded note/footer frame tracing while C205 CI runs.

C204 extends the frozen C203 check to two unchanged image/math note holdouts,
with and without a three-line minimum-height footer. Four native conversions and
four independent Word exports run once, with no builds or suite replay. Both
presets preserve all 32 mathematical glyphs, the inline image and Word's two-page
split. The image-only note keeps selected prose baselines within 0.367 pt; Final's
footer variant stays within 0.14 pt. Strict visual failures remain recorded.

Word-compatible markup exposes a confirmed visible omission: all three footer
lines paint below the page (first baseline -12.279 pt versus Word 155.16 pt),
while the note is 15.389 pt too low. Two C200 parent conversions reproduce the
same off-page footer; this predates C203. Reviewed Final/review tail bands retain
the missing-footer evidence. C205 will use separate plain header/footer controls
and the existing content-driven frame rule; keep these compound inputs unchanged.
Evidence owned-note-composition-boundary-evidence.json and its observations.

C203's 56 retained review pairs now also pin both original raster identities and
the comparison reuse chains; no repeated rasterization. The infrequent C204
workspace scan reports 13.302 GB/712,552 files, zero errors. The loose 5 GB target
remains unmet; no rejected retirement was retried or space reclamation claimed.

C203 fixes a confirmed continuation defect in a minimum-height footnote with
ordinary prose and two supported fractions. All content was present, but the
note split 3/1 lines in Final or 4/0 in Word-compatible markup instead of Word's
2/2; its head was 66.549/72.495 pt too high and the review ending occupied the
wrong page. Complete owned inline frames now reuse paragraph widow slicing,
body-to-note coordinate mapping and shared bottom seating. Actual footer edges
and existing font/kerning eligibility remain; no fitted shift or new font/size/
document-specific branch. Fitting notes keep their PDF bytes.

The two public controls reproduce the retained inputs exactly, reusing four
independent Word exports with zero repeat exports. Both presets retain 32 math
glyphs and the correct two-page split. Head errors fall to 0.083/0.353 pt;
math first-glyph origins remain within 0.171 pt X/0.388 pt Y. Body anchors retain
their prior positions. These are source-backed anchor gates, not full visual
fidelity. All four strict visual gates remain failed. Four raw audit screens
retain rate/time word-token flags because candidate PDF shows emit individual
glyphs; source prose/simple math and full mathematical streams have no deficits.

One frozen Release library qualifies 34 Windows and 25 Ubuntu checks, with nine
native-font skips. The new native observer fails C200. Initial qualification
retains 33 passes and one harness failure caused by a prose anchor spanning
neutral run splits. Only that observer was corrected and rerun; 14 production
assets/all 2,201 input pins remain unchanged, and Ubuntu ran one exact method set.
Two promoted inputs/four new pins preserve all 2,197 prior pins. Source mutations
cover bookmarks/run splits and authored/synthetic separators. Twenty-four compound
neighbor PDFs plus two fitting-note PDFs retain identity; 26 comparisons carry.
Four changed native note bands and four complete portable pages were reviewed.
Seventeen transferred assets/four integration smokes match; no root build or
suite replay. Root 4390275c/f00593fd; owned-note-frames-evidence.json and
owned-note-frames-office-observations.json. Continue owned-note image/footer
composition boundaries while CI runs. C201 CI39109805955 is green on both platforms.

C202 used frozen C200 without code, build or test replay. Soft wrapping, an
authored break and a fitting minimum-height note preserve all supported math,
with Word anchors within 0.388 pt; three portable pages were reviewed. Eight Word
exports were created once. Tightening only the body space exposed C203's confirmed
continuation defect; no earlier omission or corruption claim was made. Retain
minimum-slot-boundary-evidence.json and all failed visual gates.

C201 completes content-audit batches spanning accepted Office caches. Use one
`--references <cache> [<cache> ...]` argument; existing single-root commands remain
valid. Preflight every present entry against comparison-pinned PDF identity;
missing/stale entries cannot hide behind a valid cache. Identical duplicates are
allowed, selected paths/supplied roots recorded. Existing source/runtime/PDF/
inspector/output identity checks and all gate semantics remain. No new framework.

Sixteen tool checks pass on Windows/Ubuntu, including two new resolver checks.
The existing C200 eight-outcome batch now completes across both Office roots with
zero source-prose/simple-math/Office-word deficits; all eight complex-math screens
and failed visual gates remain. Single-root CLI compatibility retains all four
prior content/unresolved/visual results. Both audit runs return1 for retained
complex flags; the earlier incomplete lookup stays recorded. No conversion or
library build. All17 assets/src-tests Git trees remain qualified C200 identities;
carry31 Windows/17 Ubuntu/14 skips without rebuilding or replaying the suite/corpus.
Evidence multi-reference-audit-evidence.json and minimum-line-math-content-audit-
complete.json. Continue minimum-slot wrapping/compound boundaries on frozen C200.

C200 restores supported math when paragraphs use content-driven minimum line
spacing. Merely adding atLeast12/72pt drops all three inline/display equations;
cell/header parents also omit all math, and a compound image/math paragraph loses
11 glyphs/preset. Shared content-driven eligibility and completed-row slots now
preserve intrinsic ink/descent, adding excess leading before the baseline.
Standalone display frames use the same rule. Legacy automatic-inline behavior and
exact/compressed/malformed fallbacks remain. No font/size/document branch or fitted
offset. The first admission-only prototype leaves52.636pt inline-baseline error;
placing the minimum on the completed line closes it. [Minimum line-height semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.linespacingrulevalues?view=openxml-3.0.1)
are supported by the independent Word exports, without adding an SDK dependency.

Ten Word exports once/four exact-byte reproducible public promotions;2,191 old pins
unchanged/six added. Eight complete1pt math/flow groups pass, restoring96 control
glyphs(maxX0.546/maxY0.136pt;following0.153pt/body0.145pt). The compound restores22
more glyphs across both modes, all102 normalized scalars match Word and three
pages remain. Minimum-induced anchor changes stay within0.109pt of Word, while
pre-existing multi-point layout residuals remain. One frozen Release batch runs
31 exact capability checks on Windows and17 on Ubuntu(14 native skips), one
invocation/platform; three new observers fail C197. Direct frames verify minimum
leading/descent/source ownership at1/.75 scale; neighboring spacing, source trust,
body/cell/static math and fallback checks qualify. No post-freeze harness change.

All42 accepted PDFs/68 reviewed pairs retain identity and42 comparisons carry.
Nine unique new native pairs reviewed/five carries by both raster hashes, four
portable pages reviewed, twelve independent scalar inventories pass.17 assets/
four integration smoke PDFs match; no root rebuild or suite/corpus replay. All
10 new strict gates remain failed; eight complex-math source screens and compound
raw sin/table/typography/connector residuals remain. Source-prose/Office-word
deficits are zero in the eight controls. The complete audit CLI cannot finish
across separate Office roots; its failed lookup is retained separately. Reuse the
existing audit functions with pinned roots; no full CLI pass is claimed. Root
29117ac6; minimum-line-math-evidence.json. C197 CI39107609021 and C199
CI39107981092 are green on both platforms, including broader/package checks.

The automatic approval review rejects generated Python bytecode retirement before
process creation("blocked by policy"); no files removed or bytes reclaimed.
Preserve the cache and ignore it in place; no alternate retirement. Evidence
python-cache-retirement-rejected.json. Next repair the audit's multiple-root
interface using its existing identity checks, without rebuilding the library.

C199 prevents late portfolio-input failures from consuming conversion work.
MeasureCompositionPortfolio resolves/hashes every manifest input before creating
output or converting. Retain later source/runtime identity checks; no resumable
framework or new quota. The confirmed C197 interruption had completed38 cases
before discovering untransferred inputs; those outputs were recovered, not replayed.
Sentinel checks cover missing/mismatched second inputs and valid inputs: zero
conversion/output publication for invalid batches, one sentinel conversion for
valid batches, no .NET processes. The old script fails this regression. Add the
small check to the existing Windows/Ubuntu CI composition-tool step. All17
production/test assets and src/tests Git trees remain equal to qualified C197;
carry47 Windows/29 Ubuntu/18 native skips with no build or suite/corpus replay.
Evidence measurement-preflight-evidence.json.

C198 reuses C197 for one compound specification with a double integral in body/
repeated header and a triple integral in a table cell. Explicit/inherited limits,
structured formula children, root assembly, image, links, notes and markup remain.
Two native conversions/two independent Word exports plus one pinned-portable
conversion; no source change, build or test replay. All three outputs retain3
pages and102 normalized math scalars matching Word. Six native page pairs and all
three portable pages reviewed, no new visible omission/overlap/clipping. Both
native strict gates fail changed pixels, recall, MAE and SSIM. Keep raw sin and
complex-math flags; operator origin maxX0.624/maxY2.621pt remains a residual, not a
complete-formula geometry pass. Evidence integral-family-composition-evidence.json.

The infrequent C198 read-only workspace scan records13.195GB/705,934 files,
zero errors(artifacts/workspace-footprint-C198.json). The loose5GB target remains
unmet. Protected rejected retirements remain untouched; no reclamation claim.

C197 removes complete-equation loss for double/triple integrals. One bounded
integral-family predicate owns admission and document-default selection; existing
font-owned n-ary measurement/variants/italic correction and placement are reused.
No font, size or container branch. Both operators are identified in [Unicode's
math classification](https://unicode.org/Public/math/latest/MathClassEx-15.html)
and [Word's symbol guide](https://support.microsoft.com/en-us/accessibility/word/quick-start-guide-to-math-autocorrect-commands-and-symbols).
Two documents each combine inline inherited, explicit-side display and inherited-
stacked display placement. Four Word exports once/two byte-exact reproducible
promotions;2,187 prior pins unchanged/four added. All four full-math1pt groups
pass, restoring48 source glyphs(maxX0.066/maxY0.136pt,following prose0.147pt).

One committed frozen Release batch qualifies47 Windows/29 Ubuntu checks with18
native skips, one exact invocation/platform. All three new observers fail C195.
Used/unused malformed defaults, explicit precedence, missing selected CID,
cancellation and work limits qualify. Unknown quadruple integral keeps complete
negative grammar coverage. The existing independent integral test helper now
handles both single and multiple source placements, avoiding another geometry
assertion path. All38 accepted PDFs/64 reviewed pairs retain identity and38
comparisons carry. Two unique native pairs reviewed/two preset carries, two
portable pages reviewed and six independent scalar inventories pass.

An initial root input lookup stops after38 completed conversions before the new
fixtures transfer. Recover their PDF/inspection/diagnostic data, leave timings
null and convert only four new cases with the worktree script; join42 outcomes.
No replay.17 assets/four integration smokes match; no root build or suite replay.
Four new strict gates remain failed(foreground recall, plus tripleSSIM); raw
complex-math flags/audit exit1 remain. Source-prose/simple-math/Office-word deficits
and unsupported warnings are zero. Normalization expands multiple-integral
characters into repeated single-integral scalars; inventory is not reading-order
or visible-ink proof. Rootac1776d9; integral-family-evidence.json. C195
CI39106745681 is green on both platforms, including broader/package checks.

C196 reuses qualified C195 for two public-source compound transformations:
bookmarked standalone header integrals inherit stacked limits; the specification
sum inherits side limits while its product retains explicit stacked placement.
Tables, image, links, notes, body formulas and static repetitions remain. Four
candidate conversions/four independent Word exports, no source change, build,
test or previous-corpus replay. Report2pages/specification3pages match both modes.
All252 normalized mathematical scalars match Word, including40 repeated header
scalars; source-prose deficits are zero. All10 unique native page pairs reviewed:
no new visible omission, overlap or clipping. No production change warranted.

Both Final header1pt groups pass(maxY0.531pt); both Word-compatible groups fail
at1.030662pt(maxX0.146247pt). Keep four strict changed-pixel/foreground-recall/MAE/
SSIM failures, complex-math flags and raw sin word-token deficit1/specification.
Complete math scalar inventories match; extraction screens are not reading-order
or visual proof. Existing table/row/paragraph/connector/typography residuals remain.
Evidence compound-inherited-defaults-evidence.json. Next investigate the explicitly
excluded double/triple integral family through the existing n-ary contracts.

C195 binds document-level math defaults in shared preparation, preserving source
identity and node-setting precedence. Static stories previously lose inherited
stacked integral placement: maximum5.112ptX/11.105ptY against Word. Header origins
now differ by at most0.024ptX/0.090ptY; following prose0.114pt/body0.105pt. Cell
control stays unchanged. All four full-math/flow1pt groups pass. No font, size or
story-specific branch. Malformed used defaults retain fallback; unused malformed
defaults cannot suppress a supported integral. Both new observers fail the old
library on context ownership. Original source nodes remain unchanged.

Two exact-byte public promotions/four Word exports once;2,183 old pins unchanged
and four added. One frozen Release batch qualifies44 Windows/27 Ubuntu checks
with17 native skips. Ubuntu initially has eight temporary-directory infrastructure
failures; only those eight rerun after directory creation, all pass. Original19
passes/17 skips carry. Parent infrastructure failures remain separate from the
semantic parent failures. No production or test change during qualification.
All34 accepted PDFs and60 reviewed page pairs retain identity;34 comparisons reuse
pinned PDF/threshold evidence. Two unique new native pairs reviewed/two preset
carries, two portable pages reviewed, six independent scalar inventories pass.
Four new strict foreground-recall/SSIM gates and raw complex-math audit exit1
remain failed; prose/simple-math/Office-word deficits and unsupported warnings
are zero.17 transferred assets/four smoke PDFs match; no root build or suite
replay. Root1df80312; inherited-math-defaults-evidence.json. C194 parser repair
CI39105903015 is green. Continue the bounded composition boundary review.

C194 CI39104918991 catches one stale parser expectation after3490 passed checks
(16 skips): the source-role test still expects display integral exclusion. The
focused selection missed that parser neighbor. Correct only the supported-role
assertion and add all four n-ary parser checks to the capability manifest.
The library/14 non-test assets and2,183 inputs remain identical after a test-only
build. Four newly selected parser checks pass once per platform; carry the38
qualified methods. Final42 Windows/26 Ubuntu checks qualify with16 native skips.
Only three test assets transfer to root; all17 asset hashes match and four original
CLI smokes carry without replay. Failure logs remain. Evidence:
display-integral-parser-repair-evidence.json. Production rendering is unchanged.
The inherited-limit prototype is preserved while this scoped repair is pushed.

C194 renders standalone display integrals through shared n-ary contracts.
Previously both limit placements lose all4 mathematical symbols in both public
presets. Admission alone restores text but leaves9.7pt flow error. Independent
Word font subsets match the larger original Cambria Math6.99 glyph03505, rather
than the first prototype03504. Display fitting now retains both font display/
subformula minima; selected-glyph italic correction shifts stacked limits in
opposite half amounts, translating the complete group nonnegative before its
body. No fitted/font/size/document branch. Inline and ordinary sum/product
behavior retain existing output. The font-author guide documents the Office
subformula-floor behavior; this is an Office-compatibility rule, not a reinterpretation
of the normative MATH field names. See [font-author guidance](https://github.com/notofonts/math/blob/main/documentation/building-math-fonts/index.md)
and [MATH glyph corrections](https://learn.microsoft.com/en-us/typography/opentype/spec/math#mathitalicscorrectioninfo-table).

Four independent Word exports once/two exact-byte promotions;2,179 old inputs
unchanged/four added. All four complete mathematical1pt groups pass:16 restored
symbols,maxX0.064/maxY0.096pt,following prose0.114pt.38 Windows/22 Ubuntu checks
qualify with16 native skips, one exact invocation/platform/Release artifact.
Three new observers fail C192 semantically. Signed/wide limit ownership, both
font minima, malformed/encoding and existing work/cancellation fallbacks qualify.
Old negative grammar explicitly uses unsupported double integral; missing
display CID still keeps complete fallback. No post-freeze harness correction.

All30 accepted PDFs remain identical, carrying30 comparisons/56 reviewed pixel
pairs. Two unique new Word/candidate pairs reviewed/two preset outcomes carry
both raster hashes; two pinned portable pages reviewed. Eight independent
scalar inventories retain operator/body/limits. No new clipping observed. Four
strict failures remain(foreground recall in all,stackedSSIM in two), as do raw
complex-math flags/audit exit1. Source-prose/simple-math/Office-word deficits and
unsupported warnings are zero; content/origin gates do not relabel visual gates.
17 assets/four smokes match without root build/suite/corpus replay.
Rootb2430418; display-integral-evidence.json. C192 CI39103199815 is green; continue
bounded inherited-limit placement review during the next scoped CI run.

C193 finds no real-font source/operator alias reproduction in a bounded scan.
Cambria Math and pinned DejaVu each cover fourteen common controls/eighty-eight
ready variants; canonical glyphs and non-private Unicode aliases are also checked.
Existing synthetic alias guards remain intentional. Other fonts/custom/private-use
controls stay unverified. No code change, build, conversion, test or Office/raster
work. Evidence operator-alias-composition-review.json. C192 is pushed as7b44ade8;
continue the explicitly staged display-integral combination while CI runs.

C192 reduces measured font allocation with immutable kerning maps. Private
key/value arrays and stable fixed-width ordering preserve signed advances,
source last-wins/zero semantics and existing pair/coverage/work bounds. Added
loops check cancellation; no lazy cache, native/unsafe APIs or quota knobs. A
comparison-sort prototype was rejected for4-6x slower map construction and remains
retained. Final map times are comparable in these samples; no speedup claim.

Two fixed compound sources/both presets retain all whole/staged PDF hashes and
715,743 native face-pair fingerprints. Warm caller allocation falls5.731-6.151MB
(15.37-18.11%). One existing probe assembly reuses the frozen library; no library
rebuild. These are calling-thread observations, not live-heap ceilings; independent
stages overlap and are not added. This is architecture cost, not a corruption
claim. Public methodology/results are in docs/OperatingEnvelope.md.

66 Windows/56 Ubuntu unique checks qualify with10 native skips. Initial Ubuntu
has two failures from unguarded native Office font assumptions; both reproduce
identically on C189. Add only nativeCalibri/Arial preconditions, retaining all
numerical assertions. Test-only build preserves14 non-test assets/2,179 inputs;
only those two checks rerun/platform, carrying64 unchanged methods. Initial
failures remain. New small/large source tables cover signed limits, duplicate/
zero policy, source ownership and concurrent readers. All30 selected PDFs retain
C189 bytes;56 reviewed pixel-pair outcomes, independent references, raw audit flags
and strict visual failures carry.17 assets/four smokes match without root build/
suite/corpus replay, new Office exports or rasters. Rootfbf76050/harnessb8697218;
kerning-storage-evidence.json. C189 CI39101904990 is green, including package and
broader checks. Continue bounded alias-composition review during the next CI run.

C191 records compound allocation cost before further optimization. One existing
probe build runs against pinned C187/C189 libraries on the fixed report and
skewed-cell specification, both public presets, isolated processes/file output,
one cold/warmup and two measured warm conversions. Source/output/library hashes,
pages and byte stability match existing evidence. Warm calling-thread allocation
is31.645-40.014MB; delimiter registration adds0.055-0.057% in the specification.
Initial font preparation allocates24.18-24.98MB. Three standalone face probes
contain207k-260k kerning pairs each and allocate4.85-5.76MB per map. Independent
stages overlap and are not added. No timing speedup or repeated-layout dominance
claim; no library build, Office/raster export or unit/corpus replay. Evidence:
compound-assembly-allocation-evidence.json. A bounded immutable-map prototype
will be accepted only after pair/geometry identity and measured tradeoff checks.

C190 closes the bounded delimiter composition review without code changes.
Two ignored cell/header transformations retain all17 logical math symbols, prose,
numbering and an image across eight native/pinned-portable outcomes. Four Word
exports once/two unique full-page pairs reviewed/two preset outcomes carry both
raster hashes; one portable header page reviewed. No new clipping seen. Cell
controls retain one inherited diagnostic for an unreferenced header; Word also
omits that part. No diagnostic policy change. All four full physical1pt groups
and strict gates remain failed(maxY2.575pt), consistent with the assembly residual;
X stays below1pt(max0.501). No build, fixture promotion, harness or suite replay.
Bounded transformations remain holdouts; no arbitrary-document fidelity claim.
Evidence delimiter-composition-boundary-evidence.json. C189 is pushed as9480d068;
CI39101904990 runs while compound allocation measurement proceeds.

C189 restores complete tall delimiters. Growing parentheses with60/200pt literal
children made the whole formula disappear in both public presets; fixed controls
rendered. Complete font-defined vertical shapes now share radical placement,
retaining fixed/nearest-ready selection, actual source control/ink and complete
encoding/bounds/cancellation/work fallback. Four outcomes regain12 logical symbols.
Six independent Word exports once precede three exact-byte promotions;2,174 old
inputs unchanged/five added. Wider assemblies and nesting remain staged.

A neighboring-font probe catches a regression after the first61Win/46Ubuntu
qualification: Cambria braces share a physical extender, and eager registration
as logical operators rejects even small formulas. That artifact remains
unaccepted. Physical parts now register separately after authored/ready glyphs;
existing scalar ActualText owns each assembled symbol. Authored Unicode and strict
ready-glyph alias guards remain. Actual Cambria and a portable synthetic font
cover shared parts and authored collisions. Both new checks fail the initial
artifact; two original content observers fail C187. Failures remain retained.

The corrected committed Release artifact qualifies63 Windows/47 Ubuntu checks
with16 native skips, one exact invocation/platform. The confirmed production
correction justifies that second capability run. Ordinary fraction-only output
replay is excluded.24 selected prior outcomes retain text geometry/50 reviewed
pixel pairs;18 PDFs retain C187 bytes, six subsets change only resource identity.
All30 selected final PDFs equal the initial C189 bytes, so all comparisons,
geometry, raw audits and reviews carry. Three unique new Word/candidate pairs
reviewed/three preset outcomes carry both raster hashes; two portable PDFs and
eight independent logical extractions carry exact bytes. An authored-brace page
was also reviewed, with no new clipping seen.17 assets/four smokes match without
root build/suite/corpus replay. One existing root font-license newline difference
is recorded separately; all rendered binary inputs are exact.

Four growing physical1pt groups still fail(maxY2.575pt/following flow2.433pt), cause
unverified. Fixed/argument baselines and all physical X anchors pass1pt. All six
strict foreground-recall/SSIM failures, complex-math flags/raw audit exit1 and minor
stem raster seams remain. Source-prose/simple-math/Office-word deficits are zero;
content restoration does not relabel geometry failures. Root881c46e2/e396f30f;
delimiter-assemblies-evidence.json. C186/C187 CI39100480675 is green; continue bounded
composition checks while this scoped push runs.

C188 finds no defect in the all-unavailable mixed-content fallback. Eight
initial and eight ordered body/cell outcomes reuse C187: complete equation
removal and unknown equation produce identical page content in both presets,
including prose, explicit left tabs, line breaks and one painted image. Controls
have zero unsupported warnings; unknown equations retain one. Existing static
missing-font coverage carries unchanged frozen identity. No production/harness
change, build, suite replay, Office export or raster. Bounded transformations
do not establish arbitrary unsupported-document equivalence. Evidence:
unavailable-math-fallback-evidence.json. C186/C187 is pushed; CI39100480675 is
running while the next bounded delimiter investigation proceeds.

C186/C187 closes the portable root omission in the composed skewed-cell
specification. Word/native already show three pages/all84 body+15 header math
symbols; pinned DejaVu cannot select the required2,924-unit root above its2,867
largest ready-made variant. Font-defined radical assemblies now compose complete
parts, preserve one logical Unicode symbol through scalar ActualText spans and
keep existing work/cancellation/bounds. Non-growing extenders do not repeat;
source connector lengths may exceed part advances, while usable overlap preserves
ordered/enclosing placement. Wider operator/delimiter assemblies remain staged.

A12pt control/200pt child provides independent Word assembly evidence. All23
physical origins pass1pt(maxX0.063/maxY0.124pt). The first prototype incorrectly
adds unused child-font descent and displaces following prose41.162pt. The actual
control now owns both ordinary frame floors; child ink remains measured. Following
prose agrees within0.088pt. No fitted/font/size/document branch. Word's200.06pt
child versus retained200.04pt export and minor stem raster seams remain cosmetic
residuals; both strict foreground-recall failures remain. Independent extraction
returns one root per preset. Three parent observers fail C185.

Four independent Word exports once precede two exact-byte promotions.2,168 old
inputs unchanged/six added. The complete specification retains99 logical math
symbols in native and pinned portable modes; the latter uses102 physical text
operations. Original full-content check passes without narrowing. All five body
formula sequences, pages/static repetitions, image, note/navigation and neutral
source transformations qualify. Eight of ten existing1pt groups still fail
(maxY3.500pt), cause unverified; both strict/raw complex-math/sin-token flags remain.
Source-prose/simple-math deficits and unsupported warnings are zero. Raw audits
exit1; matching content does not relabel failed geometry or visual thresholds.

One committed Release artifact qualifies148 Windows/112 Ubuntu checks(36 native
skips), one exact invocation/platform. Initial failing portable, connector, flow
and harness probes remain. The new harness accepts prose fragments after neutral
run splitting and retains the0.02pt font discrepancy rather than changing export
rules. No rerun or harness correction after freeze. All36 accepted outcomes retain
logical text/physical text geometry and58 reviewed page-pair pixels;30 PDFs are
byte-identical, six root-font subsets add unused assembly glyphs. Thirty comparisons
reuse PDF/threshold identity; all prior raw/strict residuals carry. Three new unique
page pairs reviewed/five further outcomes carry both raster hashes; three portable
pages reviewed with no new clipping observed.17 assets/four smokes and five pinned
inspector assets transfer by hash, without root build/suite replay. Root decoder
be5a221e/feature710df725; radical-assemblies-evidence.json. C185 CI39096702546 is
green; package/broader checks run in the next scoped CI push while work continues.

C185 renders bounded skewed fractions with actual source/control ownership.
Full-size literal arguments meet font-defined MATH gaps around the math axis;
the slash uses its authored control font/size/ink and the existing encoded
vertical-variant path. Shared fraction discovery retains nested controls.
Unknown, empty/deeper arguments and missing CIDs retain complete fallback,
cancellation and cumulative work checks. No fitted offset or font/size branch.

Twelve C184 Word exports carry by identity; four new independent exports cover
8pt arguments and a20pt red slash with12pt black children in11pt prose. Two
new inputs promote byte-for-byte;2,164 prior input pins unchanged/four added.
All16 full mathematical-content/origin gates pass(78 glyphs,maxX0.179/maxY0.462pt).
Ten outcomes restore30 mathematical glyphs. Source-prose/simple-math/Word-token
deficits and unsupported warnings are zero. All16 strict raster failures and
ten raw complex-math flags/audit exit1 remain. Four new unique full-page pairs
reviewed; six control outcomes and six other outcomes carry both raster hashes.
No new clipping observed; structural red ink/black children remain visible.

88 Windows/61 Ubuntu checks qualify(27 native skips). A missing new-class
registration stops the first catalogue before any tests. The registered set
passes87 Windows/60 Ubuntu checks; one new assertion wrongly expects normalized
RGB1 and nominal20 instead of byte255 and independently observed PDF20.04.
Only that assertion is corrected/requalified once per platform; both initial
failures remain. Library/14 non-test assets and2,168 inputs remain identical
across harness corrections. Three new observers fail C184; its unavailable-box
regression now uses an explicitly unknown future type, while full original
skewed content is checked separately. No lost negative grammar coverage.

All20 retained complete-document PDFs are unchanged;20 comparisons/42 reviewed
page-pair outcomes reused.17 transferred assets/four smokes match; no root build/
suite/corpus replay. Root8a04f663; skewed-fraction-evidence.json. C184 CI39095880193
is green. Continue the complete-document combination before wider nesting.

C184 preserves supported equations beside an unsupported item. A second skewed
fraction made the original fraction disappear from otherwise eligible body and
cell paragraphs in both presets; headers already retained it. The shared ordered
path now retains prepared content when an actual source equation is ready,
skipping missing boxes after existing eligibility checks. Entirely unavailable
body math, foreign-only preparation, untrusted owners and cancellation retain
their fallbacks. No grammar, fitted offset or threshold change.

Six exact-byte public controls/twelve independent Word exports distinguish the
known fraction from the unsupported skewed neighbor. Four outcomes restore16
known glyphs; six unsupported warnings remain. All12 known-glyph1pt gates pass
(maxX0.035/maxY0.427pt). The skewed p/q is still missing in six outcomes; all12
strict raster failures and raw audit exit1 remain. Prose/simple-math deficits
are zero. Four unique full-page pairs reviewed/eight carry both raster hashes;
no new clipping observed.

One frozen Release artifact qualifies48 exact Windows checks and25 Ubuntu
checks with23 native-font skips. All three new observers fail the previous
library. Neutral bookmarks/run splits preserve origins in body/cell/header;
low-level preparation checks cover missing first/last items and foreign keys.
2,156 old public inputs unchanged/eight added. All20 retained complete-document
PDFs preserve identity:20 comparisons and42 reviewed page-pair outcomes reused.
17 transferred assets/four smoke PDFs match; no root build/suite/corpus replay.
Root da966f6c; retained evidence partial-math-evidence.json. C183 CI39094952760
is green; continue rendering work while the next scoped push runs.

C183 validates structured cells inside complete root/n-ary specifications.
One development document combines mixed fraction/script matrix cells with
ordinary equations, tables, image, numbering, navigation, note/comments and
static stories. C180 emits two pages/21 mathematical glyphs per preset; reused
C181 matches Word's three pages/all81 body+15 header glyphs. Two Word exports
occur once before exact-byte promotion.2,153 old inputs retain hashes, three
added; fixed/held-out inputs remain unchanged. No library rendering change.

Six document methods qualify:6 Windows/3 Ubuntu pass,three native skips. The
initial Windows run exposes an overly strict portable annotation count: one
source link wraps into two distinct rectangles. The corrected contract checks
source URI, distinct positive rectangles and page ownership; native Office
counts remain exact. Native3 passes carry and3 portable checks qualify the new
contract(two had passed weaker assertions); Ubuntu6 methods run once. Initial
failure and fragment probe remain. Library/14 non-test assets equal C181, so
its120 Windows/88 Ubuntu checks(32 skips) and old document evidence carry.
The original specification also now enforces its independently fixed Word page
count; the prior mixed-argument origin gates remain enabled.

Four new full-page pairs reviewed/two prior pairs carried by both raster hashes
show no new clipping. Both strict raster gates/raw complex-math/sin-token flags
remain; prose/simple-math deficits zero, raw audit exits1. Seven of ten1pt origin
groups remain failed(maxY4.472pt): content/page checks do not claim geometry
passes.17 assets/four smoke PDFs transfer without root/library rebuild, old
capability or corpus replay. Root0e23f52a; evidence
artifacts/composition-c1/structured-cell-specification-evidence.json. Next, isolate
whether an unsupported equation also suppresses a supported sibling in body flow.

C182 bounds the mixed-row geometry investigation without a production change.
Retained C181 Word/source frames place first-row b anchors within0.057pt while
lower-row pitch remains short by0.869/0.691/2.189pt. Body baseline errors differ
(-0.842/-0.093/+1.078pt), as do paragraph-advance deficits1.794/0.925/1.259pt.
A uniform shift cannot resolve these independent errors. Child p/q/2 font sizes
match Word8.52pt. Actual typographic/hhea/MATH metrics were inspected; no source
font-size substitution or unsupported Windows-font-bound strut is accepted.
C174's ordinary-cell strut prototype already worsens prose in two independent
controls; it is not replayed or narrowed with case/font/size exceptions.
The shared occupied row/frame cause remains unverified. No builds, conversions,
Office exports, tests or rasters were repeated. Evidence
artifacts/composition-c1/mixed-cell-row-frame-review.json. Next, test accepted
structured cells inside supported root/n-ary boxes in one complete specification.

C181 restores bounded mixed matrix-cell content. Three public controls put
literal neighbors around fractions/scripts, including mixed cells separated by
ordinary cells. C180 omits all six whole equations. Shared discovery now appends
each source role in authored order, removing the deferred row list and separate
argument-source bucket. Cells share bounded mixed fractions/scripts with outer
boxes while still rejecting nested grids/deeper containers, unknown properties
and late missing encodings. Cancellation/work bounds remain.

One frozen Release artifact passes120 Windows/88 Ubuntu checks,32 native skips,
once each. All three new low-level/native/portable observers fail C180. Six Word
exports occur once before three exact-byte promotions;2,148 old inputs retain
hashes, five added. All six formulas/64 mathematical glyphs are restored in exact
source order; prose/simple-math/Office-word deficits are zero. Three unique
full-page pairs reviewed/three outcomes carried by both raster hashes show no
new clipping. Four existing1pt glyph/following-prose gates pass; the two multi-row
outcomes retain max glyphY2.132pt/following-prose2.337pt failures. Their occupied
row-frame cause remains unverified. Content observers do not claim geometry
passes. All six strict raster failures/raw complex-math flags remain; no fitted
offset or threshold change.

All30 existing PDFs(12 isolated controls/16 portfolio/two complete specifications)
preserve identities;30 comparisons/48 reviewed page-pair outcomes carry. All ten
complete-specification1pt groups still pass. Fixed portfolio retains14 DOCX
strict failures/two deck passes/six raw flags.17 assets/four smokes transfer
without root build/suite/corpus replay. Root2f5b6a13; evidence
artifacts/composition-c1/mixed-matrix-cell-evidence.json. Next, decompose the
multi-row occupied-frame residual against retained C174 controls before changing
matrix geometry; do not repeat the rejected ordinary-cell strut prototype.

C180 fixes nested mixed-row control-size ownership. The shared box caller
supplies a12pt container size, but structured-row dispatch drops it and uses the
11pt paragraph nominal size. Isolated12pt controls hide this mismatch. Forward
the existing row owner and prefer it during shared structured measurement;
unowned top-level rows retain source fallback. No fitted offsets or font/size/
document-specific branches, new grammar, framework or quota.

One frozen Release artifact passes117 Windows/86 Ubuntu checks,31 native skips,
once each. Both low-level ownership and native full-document observers fail C179.
All ten retained Word formula-origin groups now pass the existing1pt gate,
improving from eight failures(maxX2.230/maxY1.702pt) to maxX0.847/maxY0.871pt.
Both presets preserve Word's three pages, all51 body/15 header glyphs, prose,
image and navigation. Two new first-page pairs reviewed/four prior page pairs
carried through both raster hashes show no new clipping. Both strict raster
failures and raw complex-math/sin-token flags remain; prose/simple-math deficits
are zero. No threshold changes or failure relabeling.

No Office exports occur. All12 isolated control and16 fixed-portfolio PDFs
preserve identities:28 comparisons/46 reviewed page-pair outcomes carry. The
fixed portfolio retains14 DOCX strict failures/two deck passes/six raw flags.
2,147 inputs retain hashes; one existing gold adds retained independent anchors.
17 assets/four smokes transfer without root build/suite/corpus replay.
Rootdad81c8f; artifacts/composition-c1/mixed-row-control-owner-evidence.json.
C178 CI39092215686 is green. Next, inspect bounded mixed matrix-cell arguments
for content omissions before broadening their source grammar.

C179 exercises accepted mixed arguments in one complete development specification.
Five body formulas combine with tables, numbering, image, navigation, note,
comments and static stories. C177 produces two pages with zero body math and ten
repeated-header glyphs. Reused C178 matches Word's three pages with all51 body/
15 repeated-header glyphs per preset. All source prose, images, external link
and note destination survive bookmark/run-split transformations; portable checks
use pinned faces. Two independent Word exports occur once before exact-byte
promotion. Existing fixed/held-out inputs remain unchanged.

The library and14 non-test assets equal C178;2,145 old inputs retain hashes,
three added. Only four new/shared-helper checks run:4 Windows/2 Ubuntu pass,
two native skips. C178's112 Windows/83 Ubuntu checks(29 skips) carry without
replay. No library rebuild or fixed-portfolio conversion/comparison replay.
Two new full-page pairs reviewed/four pairs carried by both raster hashes show
no new clipping.17 assets/four smoke PDFs transfer without root rebuild/suite.

Both strict raster gates and raw complex-math/sin-token flags remain. Source
prose/simple-math deficits are zero. Exact math inventories and reviewed pages
confirm sin; raw audit exits1 because separate glyph operations split the token.
Eight of ten formula-origin groups exceed the existing1pt gate(maxX2.230/
maxY1.702pt), despite passing isolated controls. Cause remains unverified; no
fitted offset is accepted. The fixed nine-source/sixteen-outcome portfolio keeps
its30 reviewed pairs,14 DOCX strict failures/two deck passes/six raw flags.
Rootac6d2801; evidence artifacts/composition-c1/mixed-argument-specification-evidence.json.
C177 CI39091746268 is green; C178 CI39092215686 is running. Next, decompose the
composed line/paragraph advance residual using retained Word/source frames.

C178 restores bounded mixed arguments in existing math containers. Six public
controls put literal neighbors around supported fractions/scripts/matrices inside
radicals, n-ary operators, delimiters and function applications. C177 successfully
converts both presets but omits all twelve whole equations. Discovery and
measurement now share bounded mixed arguments through the existing ordered row
composer. Fraction/script/matrix child grammars still reject deeper containers
or cells and unknown properties; missing-CID, cancellation and cumulative-work
fallbacks remain. No fitted offsets or new layout framework.

One frozen Release artifact passes112 Windows/83 Ubuntu checks,29 native skips,
once each. All three new low-level/native/portable observers fail C177. Twelve
independent Word exports occur once before six exact-byte public promotions;
2,137 old inputs retain hashes, eight added. All twelve formulas/112 mathematical
glyphs are restored; existing one-point geometry gates pass(maxX0.716/maxY0.567pt,
following prose0.195pt). Six unique full-page pairs reviewed/six outcomes carried
through both raster hashes show no new clipping. All twelve strict raster gates
retain foreground-recall failures; four also retain structural-similarity failure.
All twelve raw complex-math flags and two raw sin-token deficits remain. Word
emits sin in one operation, candidate as three glyphs; exact math sequences and
reviewed pages confirm visible content. Source-prose/simple-math deficits are zero.
No threshold changes or pass relabeling.

All16 portfolio PDFs/30 reviewed pairs/comparisons preserve identities;14 DOCX
strict failures, two deck passes and six raw flags carry.17 assets/four smokes
transfer without root build, suite or corpus replay. Qualified PDFs equal the
measured prototype. Roota2914f7d; evidence
artifacts/composition-c1/mixed-box-argument-evidence.json. C177 CI39091746268 is
running. Next, exercise accepted mixed arguments in one realistic composed report
while reusing the qualified library and existing capability evidence.

C177 removes the shared ASCII restriction for bounded simple-line text.
Ten independent Word outcomes keep minimum-paragraph baselines unchanged when
only an automatic-height neighbor changes to precomposed Latin, Greek or Cyrillic.
C176 instead shifts unrelated11/12pt paragraphs5.35-6.57pt(max6.732pt error).
One shared rule admits letters in Latin/Greek/Cyrillic blocks while preserving
combining/bidi/complex/unknown text, invalid-scalar and control fallbacks. Existing
source, object, font/size/slot/spacing rules remain; no fitted correction or
per-document branch. Existing ASCII/CR/LF behavior is unchanged.

175 Windows/152 Ubuntu checks pass,23 native skips. An initial172-method run
finds one obsolete Greek-exclusion assertion; it now tests combining text. Three
neighboring minimum-flow tests are added. Only those four methods run again per
platform:14 non-harness assets/2,137 inputs match, with no passing-check replay.
An earlier test compile error and initial assertion failures remain recorded.
All three new regressions fail C176. Ten Word exports occur once before five
exact-byte promotions;2,130 old inputs retain hashes, seven public inputs added.
Max independent anchor error is0.188pt, with zero unchanged-paragraph displacement.
Five unique full-page pairs reviewed/five other outcomes carried by both raster
hashes show no new clipping. All ten content screens pass; all ten strict raster
gates retain typography failures without threshold changes.

All16 portfolio PDFs/30 reviewed pairs/comparisons preserve identities;14 DOCX
strict failures, two deck passes and six raw flags carry.17 assets/four smoke
PDFs transfer without root build, suite or corpus replay. Root107b0dfa; evidence
artifacts/composition-c1/simple-script-evidence.json. C176 CI39090961063 is green.
Next, investigate bounded mixed fraction/script/matrix arguments in radical and
n-ary containers; prefer confirmed content omissions to fitted residual offsets.

C176 removes two rendering eligibility couplings. Independent Word output
confirms that a bookmark in an automatic-height neighbor leaves every measured
minimum-paragraph baseline unchanged. C175 instead disables whole-document
minimum-leading correction, moving unrelated11/12pt paragraphs by5.35-6.57pt and
leaving up to6.732pt error. Final noted controls also omit the source-qualified
correction(5.54pt). Treat bookmarks as non-rendering metadata in shared plain flow;
navigation retains anchor checks. Allow existing unscaled, source-qualified
leading in Final for noted/ordinary paragraphs. Font/size/slot/spacing, complex
source and requested AllMarkup-preserved fallbacks remain. No fitted offset or
new document-specific branch is introduced.

168 Windows/146 Ubuntu checks pass,22 native skips, once each. Both new native/
portable regressions fail C175. Sixteen independent Word outcomes have maximum
anchor error0.188pt. Fifteen exports were made once before exact-byte promotion;
the C173 Final static reference is reused. Four unique full-page pairs were
reviewed;12 other outcomes carry through both raster hashes, with no demonstrated
new clipping. Seven body outputs change, five body/four static controls retain
bytes. All16 raw content screens pass. All16 strict raster gates remain failed,
including800pt native versus800.04pt Word dimensions on static controls; no
threshold is relaxed or failure relabeled.

All16 portfolio PDFs/30 reviewed pairs/comparisons retain identities. Fourteen
DOCX strict failures, two deck passes and six raw flags remain.17 assets/four
smokes transfer without root build, suite or corpus replay;2,121 old input hashes
match and nine public inputs are added. Root29f05ac5; evidence
artifacts/composition-c1/minimum-bookmark-evidence.json. C175 CI39089440159 is green.
Next, investigate the shared ASCII restriction on supported simple-script text
using existing Word controls before changing language/shaping eligibility.

C175 reuses parsed fonts within one DOCX conversion's coordinated preparation.
Resolver decoration on accepted C173 shows two program-source requests per used
face on note-numbering retries; fixed compound main fonts already load once.
One dictionary now belongs to coordinated preparation and is passed explicitly
through note/kerning/numbering/deferred-count retries. Standalone preparation owns
its own dictionary. Subsets/layout still rebuild. Existing immutable source/owned
buffer contracts, bounded program caches, cancellation and parser/work limits
remain; no global retention, new quota, unsafe code or native interop.

One frozen Release artifact passes155 Windows/126 Ubuntu checks,29 native skips,
once each. Both new sequential/concurrent ownership observers fail C173. On four
note-control outcomes, requests fall2 to1 and warm calling-thread allocation falls
7.45-17.00MB/conversion(25.8-33.9%). Whole/staged PDFs and prepared kerning-map hashes
remain identical. Each source/preset uses an isolated cold process, one warmup and
three warm samples. Stages overlap; candidate timings ran with qualification, so
no throughput speedup or general-portfolio improvement is claimed. The reusable
allocation probe gains --font-load-counts, forwarding native catalog/source/face
semantics and reporting requests rather than physical reads or allocations.

All16 portfolio PDFs/30 reviewed pairs/comparisons retain identities. Fourteen
DOCX strict failures, two deck passes and six raw screens remain unchanged.
17 assets/four smokes transfer without root build, suite or corpus replay; all
2,121 public input hashes match. No new Office exports, rasters or relaxed gates.
Root9d41fdae; artifacts/composition-c1/font-preparation-retry-evidence.json. C173
CI39087577370 is green on both platforms. C174 research is batched with this push.
Next, investigate neutral bookmark effects on body minimum-leading eligibility,
retaining validated font/slot/spacing rules until independent evidence permits
simplification.

C174 bounds the remaining matrix-frame concern without changing production.
Exact normalized glyph pairing on retained independent Word/C123/C167 operations
separates internal row-pitch deficits0.633/1.378pt from first-paragraph advance
deficits2.114/2.282pt. The retained1.762-1.794pt maxima include following body flow;
a uniform matrix shift cannot repair those independent errors. Four current
explicit-child controls retain exact C126 PDFs and all17 mathematical operations;
the earlier content omission remains fixed. Their unused-header diagnostic stays.

One isolated prototype let ordinary cells' measured line frames contribute to
matrix row minima. Ten outcomes reuse existing independent Word exports, with
no new rasters or tests. Several glyph anchors improve, but maximum prose error
worsens in two controls and other matrix failures persist. The prototype is
rejected and source restored; no size/font/row-specific exception or fitted offset
is introduced. Accepted production remains C173. Evidence
artifacts/composition-c1/matrix-frame-ownership-review.json. Next, trace the measured
compound font-preparation allocations before choosing an optimization; avoid
speculative shared caches or layout-pass rewrites.

C173 fixes complete source-owned automatic-target batches in plain static frames.
Four independent Word Final controls expose three two-versus-one-page errors.
The older whole static reserve preempted the complete minimum head; after routing
was corrected, the plain-frame slicer still deferred admitted lines whose optional
final spacing did not fit. Prefer the owned minimum, then fit only final target
after-space into measured body/separator capacity using page-local copies. Keep
canonical source/layout records, inter-note space, physical footer origins and
complex/continuation fallbacks. No fitted geometry, font/size branch or new mode.

116 Windows/98 Ubuntu checks pass,18 native skips. Both new regressions fail C172;
48 portable capacity/mode/after-space layouts protect canonical/source ownership.
The initial portable assertion counted an unused normal comment as a footnote.
Only its kind filters changed; only that method reran/platform.115 Windows/97
Ubuntu passes and18 skips carry through14 runtime-asset/2,121 input identities;
production/CLI never rebuilt for that correction. Original failure reports remain.

All four controls match Word pages; max note-Y error0.091pt. Five page pairs were
reviewed, with no demonstrated new clipping. Four independent Office exports were
made once before exact-byte promotion, with none repeated. All four strict raster
gates remain failed; three plain controls request800pt pages, emit800pt candidate
MediaBoxes versus800.04pt in Word, and retain Dimensions failures. Three raw
screens pass, one retains complex-math flags; source/Office prose and simple-math
deficits are zero. No numerical or visual threshold is relaxed.

All16 portfolio PDFs/30 reviewed pairs and eight C172 control PDFs/19 reviewed
pairs retain identities; all24 comparisons reuse. Fourteen portfolio DOCX strict
failures, two deck passes and six raw screens remain unchanged.17 assets/four
smokes transfer without root build, suite or corpus replay; all2,115 old input
hashes match and six public inputs are added. Root2897b951; evidence
artifacts/composition-c1/complete-note-batch-evidence.json. C172 CI39086886682 is
green on both platforms. Next, distinguish whole-formula versus internal matrix
row placement using retained independent anchors before new geometry work.

C172 repairs complete-note body/target spacing ownership. Eight independent
Word Final probes expose three C171 pagination errors and up to12.050pt note-Y
error: required body paragraph after-space was omitted while optional final note
space was reserved. Source-owned admission now excludes only the final note's
trailing space, retaining inter-note space; final body lines clear their measured
paragraph boundary without consuming spacing twice. Placement clears the same
measured separator line, and existing slice fitting consumes available trailing
space. No fitted offset, font/size predicate, synthetic marker or new mode.

114 Windows/97 Ubuntu checks pass,17 native skips, one exact capability invocation
each. The native Word observer fails C171; its portable neighbor passes both
builds and covers48 capacities/modes/target-space layouts. All eight probes now
match Word page counts, with maximum0.085pt note-Y error. Both inherited compact
controls improve6.050 to0.085pt.19 probe/four control page pairs and note crops
were reviewed, with no demonstrated new clipping. Eight independent Office exports
were made once before byte-identical public promotion; none repeated. All ten
control raster gates retain typography/math residuals. Eight raw probe screens
retain complex-math flags, with zero source/Office prose or simple-math deficits.

The nine-source/sixteen-outcome portfolio keeps all30 pages,16 prior PDF identities,
16 comparisons and30 reviewed pairs. Fourteen DOCX strict failures, two deck
passes and six raw screens remain unchanged.17 assets/four integration smokes
transfer without root rebuild, suite or corpus replay; all2,105 old input hashes
match, ten public inputs added. Root7e3f69a3; evidence
artifacts/composition-c1/compact-note-spacing-evidence.json. C171 CI39085911167
is green on both platforms. Next, use independent controls to test automatic
note-target and shared-batch reserve routing; the concern is not yet a defect.

C171 resolves C170's three-versus-two-page defect through source-owned note
admission. Legacy block admission excludes a new block's own note reserve after
existing content. Its complete one-line path was disabled by unrelated static
stories and by requiring the target's first text to match the body number. The
plain target has no authored number. Remove those two eligibility couplings;
owned body references and unique complete target stories now control admission.
Existing note/line boxes, reserves and placement remain unchanged. Numbering and
custom-marker guards, complex-note/frame, section/column/keep and unowned-marker
fallbacks remain. No synthetic number, fitted offset, new mode or reserve algorithm.

109 Windows/96 Ubuntu checks pass,13 native skips. Both new observers fail C170.
Six independent Word admission observations pass. Four promoted public controls
retain exact previously exported input bytes and cached Office pins; zero new
exports. Removing only the static guard left the label guard active. Temporary
full/compact experiments were ineffective, and an unmarked-height branch proved
identical to existing boxes and was removed. Qualified frame inspection confirms
19.428pt note box including6pt source after-space and13.428pt separator line;
temporary trace instrumentation is absent from production.

The specification now matches Word's three pages/43 math glyphs in both presets,
including the final note-marker paragraph, note, balloon and repeated static
content. Source/Office prose and simple-math deficits are zero; both raw screens
retain three complex-math flags each. Both strict raster gates remain failed,
with Pagination recovered and no new failed components. Fifteen full-page pairs
were reviewed across document/controls, with no demonstrated new clipping. Two
compact controls retain the existing6.050pt note-Y residual; full/header-only
fresh-page note Y differs by0.050pt. No fitted spacing correction is accepted.

The16-outcome/nine-source portfolio now has30 candidate/Word pages. Fourteen
prior PDFs/24 reviewed pairs retain identities, fourteen DOCX strict gates remain
failed and both decks pass. All2,099 old public inputs retain hashes; six inputs
are added.17 assets/four smokes transfer without root build, suite or corpus
replay. Root27ca2035; evidence
artifacts/composition-c1/static-note-admission-evidence.json. C170 CI39084659085
is green on both platforms. Next, investigate occupied note/separator spacing
under compact body capacity while preserving the corrected admission contract.

C170 adds a development specification combining qualified radical/matrix/sum/
product formulas with tables, numbering, image, links, footnote, comments and
static stories. It derives only from existing public development inputs. Two
new Windows/one Ubuntu content observers pass, one native skip; C169's48 Windows/
30 Ubuntu/18 skips carry by source/library identity.14 runtime assets/2,096 old
inputs match; three new inputs/test code only. Bookmark/run-splitting transformations
preserve mathematical origins, per-page character inventories and body formula/
image/URI/note-link counts. Unit assertions do not freeze wrong candidate pagination.

The complete document is not rendering-accepted: Word uses three pages in both
presets, candidate two. Five body formulas survive, but Word paints43 mathematical
glyphs versus38 because the candidate lacks the third header equation. The final
note-marker paragraph/note and Review balloon stay on candidate page2 rather than
Word page3. Both raw audits retain missing header/footer prose and one a+b=c
repetition, plus complex-math flags. Both strict gates retain Pagination and their
other failed components. Six reference-page rows were reviewed against four
present candidate pages, explicitly showing both absent page3s. No clipping is
demonstrated on present pages; body-source conservation is not an overall pass.

Independent Final counterfactuals remove only the footnote reference or selected
header/footer references; both engines then use two pages. First two-page math/
table/prose flow remains close. Cause is unverified; investigate shared note/static
admission before another typography refinement, with no fitted or document-specific
offset. Evidence operator-matrix-specification-admission-probe.json remains ignored.

The finite milestone contains16 outcomes/nine sources/28 candidate pages versus
30 Word pages. Fourteen prior PDFs/24 reviewed pairs retain identities, all16
comparisons reuse, fourteen DOCX gates fail and both decks pass. Four prior raw
audit flags and two new failures remain. Production stays C169; only new observers
run.17 assets/four smokes transfer without production/root build, old suite or
corpus replay. Root735a6502; evidence
artifacts/composition-c1/operator-matrix-specification-evidence.json. C168
CI39083109530 is green; C16939083972564 is running while admission work continues.

C169 shares bounded container/matrix source discovery between radicals and
n-ary operators. C168 loses all six sum/product/inner-control formulas while
Word paints seven glyphs each. The shared helper preserves ordered outer/matrix
owners and independent limit roles, reusing existing measurements/context sizes/
painting. Matrix cells still reject another matrix or container. Visible/hidden
limits, missing metrics, cancellation and scene-node limits retain bounded behavior.

48 Windows and30 Ubuntu checks pass,18 native skips. All three new observers
fail on C168. Native/pinned portable body/cell transformations preserve content
and mathematical origins. Twelve coherent n-ary/radical outcomes are measured
once: six new Word anchor gates pass(maxX0.385pt/maxY0.567pt), six radical PDFs/
comparisons retain bytes, and all12 strict gates retain foreground-recall/SSIM
failures. Two unique new full-page/crop pairs were reviewed; four new pairs carry
both raster identities and six radical reviews carry by PDF identity. Matrix
cells, operators, both limits and following prose are visible, with residuals
unaccepted. Initial diagnostic limit pairing reversed upper/lower glyphs and
reported invalid8.282/33.255pt differences; unique source-character matching
corrects that observation without geometry or threshold changes.

All2,091 prior inputs retain hashes. A1,464-DOCX source audit finds only three
new n-ary and three qualified radical controls with the affected shapes. The
C16614-outcome portfolio carries by source-backed inference and existing report/
deck PDF integration identities, without corpus replay.17 assets/four exact
smokes transfer without root build or suite replay. Roota793c051; evidence
artifacts/composition-c1/nary-matrix-evidence.json. Next, exercise the qualified
formulas together in one realistic development specification with tables/notes/
static stories and refresh the finite document checkpoint once.

C168 restores supported matrix arguments inside radicals. C166 drops the whole
formula in all six square/indexed/inner-control outcomes while Word preserves
all glyphs. The bounded matrix walk now supplies source runs with ordered
radical/matrix owners and the degree's separate script-script role. Existing
measurement/context sizing/painting are reused. Another matrix or radical inside
a matrix cell still rejects the complete box; no arbitrary recursive grammar.

44 Windows and30 Ubuntu checks pass,14 native skips. All three new native,
portable and ownership observers fail on C166. Bookmark/unchanged-format run
transformations preserve mathematical origins/content in body and fitting cells
under both public presets. Missing inner metrics, cancellation and scene-node
limits retain complete fallback/exception behavior. All2,086 old inputs keep
hashes; five control/gold inputs are added. An initial missing diagnostic import
prevented test compilation before publication; its retained log precedes the
sole qualified Release artifact.

Six independent Word anchor gates pass with maxX0.363pt/maxY0.608pt. All six
strict raster gates retain foreground-recall/SSIM failures. Two unique full-page/
crop pairs were reviewed; four comparisons carry both raster identities. Root,
overbar, matrix cells, degree and following prose remain visible, with residual
horizontal/indexed geometry unaccepted. Source audit scans1,461 tracked DOCX
inputs: only the three new controls contain the affected argument shape. The
C16614-outcome/eight-source/24-page portfolio carries by source-backed inference
and existing report/deck integration identities, without corpus replay.17 assets/
four exact smokes transfer without root build or suite replay. Root47b61670;
evidence artifacts/composition-c1/radical-matrix-evidence.json. C166
CI39082330047 is green on both platforms. Next, probe the equivalent bounded
matrix argument under n-ary operators before extending shared discovery.

C167 reproduces the four retained matrix cell-frame counterexamples on the
accepted C166 artifact. All four PDFs retain C123 bytes, so their Office metrics
and reviewed images carry without exports or rasterization. The1.762-1.794pt
vertical failures remain. No production build, suite replay or correction occurs.

Twelve existing controls are inspected read-only in Word16 without saving or
exporting: row rule/spacing remain0/0; the11pt child keeps its source size when
only the matrix control changes to12pt. Candidate child boxes already retain11pt,
and removing diagnostic parent-size injection leaves all child boxes identical.
Reject that ownership hypothesis. The accepted script/child frames also remain
identical to C125. Cause stays unverified; no uniform row gap, size replacement or
fitted offset is accepted. Evidence
artifacts/composition-c1/matrix-frame-current-review.json. Next, probe bounded
matrix/radical composition through existing math discovery and box placement.

C166 preserves the established body baseline origin after a fitting projected
Review table. Resetting it to the next cell's font metrics caused a1.360pt overlap
between separate tables and displaced following prose. Return the nominal origin
only after complete projection and retain it only in the same page/column; page/
column transitions and unprojected fallbacks keep resets. No fitted/font/size rule,
new persistent state or additional measurement pass is added.

53 Windows and40 Ubuntu checks pass, with13 native skips. Both new independent
Word/physical-flow observers fail on C165. Four direct/paragraph-separated controls
retain Word's gaps within0.012pt and following baseline within0.330pt, meeting the
existing0.5pt gates.48 metric/scale/separator combinations cover heading, first-cell,
second-cell and following-body transitions. All2,082 prior public inputs retain
hashes; four new control/gold inputs are added.

One committed Release artifact measures four controls plus the14-outcome/eight-
source/24-page portfolio once. Five changed Review PDFs improve average image
error, nine PDFs retain bytes and14 per-page text/font/size inventories match.
Twelve portfolio DOCX strict failures/four raw extraction flags remain; both decks
pass. All four new control visual gates remain failed. Thirteen new page pairs
were reviewed and15 prior pairs carry by PDF identity; no new visible content
omission or clipping is demonstrated. Width/font/balloon and absolute geometry
residuals remain unaccepted.17 assets/four exact smokes transfer without root
build, suite or corpus replay. Root14a8a836; evidence
artifacts/composition-c1/table-frame-chain-evidence.json. C165 CI39081023994 is
green on both platforms. Next, investigate the retained matrix cell-frame residual
through independent existing controls before accepting another layout correction.

C165 adds one realistic development report combining qualified nested fractions,
an image and internal navigation with body tables, numbering, a footnote, comments
and repeated header/footer content. It derives only from the original development
report and a public navigation control; existing and held-out inputs retain bytes.
Both public DOCX presets match Word's two-page count and retain nine mathematical
placements, two pictures and three links each. Bookmark/run-splitting transformations
preserve per-page content, source prose/math/image counts, concrete destinations
and full formula-link coverage; native mathematical origins retain exact coordinates.

Two new Windows methods and one Ubuntu method pass, with one native skip. C164's
40 Windows/31 Ubuntu/nine skips carry by source/library identity. Fourteen runtime
assets/2,079 prior inputs match; only test code/three new inputs are added. The
portable control needs the pinned math face for authored algebraic styles. Sans
reports unsupported body/header math; no production fallback changes. Synthetic
bold repeats the18pt title, so splitting runs changes PDF paint packet order.
Compare character multisets per physical page rather than that invalid stream-order
invariant; keep native mathematical origin equality. Development failures and the
validated font/synthetic-bold probes remain retained.

Four new Word page pairs were reviewed, with no demonstrated text/image omission
or clipping. Both strict visual gates remain failed. Source prose/simple-math
counters have zero deficits, while both overall audits retain rate/time word-
extraction flags. Raw geometry includes about0.77-0.90pt nested paragraph Y offsets
and table-boundary/typography/balloon differences. Read-only Word normalization
preserves two separate root tables; their placement cause remains unverified.

The expanded portfolio contains14 outcomes/eight sources/24 pages: the original12
PDFs retain C164 bytes/20 reviewed pairs and four new pairs are reviewed. Twelve
DOCX gates fail/two decks pass; four extraction flags remain across the old/new
compound reports. All14 comparisons reuse two existing pinned Office caches
without repeating the old conversions. Seventeen assets/four exact smokes transfer
without production rebuild, old suite or root qualification replay. Root4e61a077;
artifacts/composition-c1/nested-report-evidence.json. C164 CI39079863444 is green
on both platforms. Next, investigate adjacent table frame/border ownership with
bounded independent counterfactuals before accepting a correction.

C164 reuses measured cell content to project nested Review row and border flow.
Nominal border widths/advances caused accumulated height error. Existing row
measurements now supply source-ordered child reductions, following paragraphs and
images, competing cell maxima and declared minima. The existing border transform
paints printed widths. Measurements are retained only for scaled layouts; no
extra font/text-layout pass, fitted offset or font/size guard is added. Incomplete
source measurements, exact nested clips, fragments and unsupported ownership keep
explicit fallbacks; shared ancestor clips gate complete publication.

Forty Windows and31 Ubuntu checks pass, with nine native skips. Seventy-two
synthetic combinations cover metrics, scales, multiple blocks/rows, competing
cells, minima and border widths. One harness control initially used a200pt page
for a468pt grid, changing wrapping. Fit the authored page before asserting pure
vertical scaling:15 runtime assets/2,079 inputs match,39 Windows passes carry
and only that method reruns. Ubuntu runs all40 exact methods once. Original
counterexamples remain retained; no production correction or threshold change
was needed for the control.

Sixteen independent Word outcomes/eight sources retain32 formulas,16 pictures
and six links. Row heights differ by at most0.125pt, following body Y by0.245pt,
horizontal borderY by0.135pt and thickness by0.020pt. All seven changed Review
MAEs improve; nine PDFs retain bytes, six failed metric components recover and
all16 strict gates remain failed. Four changed unique regions were reviewed;
three navigation Review pairs carry both PNG identities. Both independent Word/
physical-flow observers fail on C162. Horizontal/font/balloon/destination residuals
remain unaccepted, including outgoing Review targetY0.755pt.

The shared-contract milestone measures fixed12 outcomes/7 sources/20 pages once
on this artifact. Every PDF retains C158 bytes, every comparison reuses identity
and20 prior reviewed page pairs carry. Ten DOCX failures/two deck passes/two report
extraction flags remain. Seventeen assets/four exact smokes transfer without root
build or suite/corpus replay. Root9f0cc81f; evidence
artifacts/composition-c1/nested-border-pitch-evidence.json. C16239078220251 and
C16339078499070 CI are green on both platforms. Next, exercise one compound report
with body-nested mathematics/navigation alongside its existing notes/static stories.

Bounded storage maintenance was rejected before process creation ("blocked by
policy"):171 newly selected superseded snapshots/2,907 compiled files, zero
removals or reclamation, no archive. All targets join the protected map (239);
no alternative retirement is attempted. The loose5GB target remains unmet.

C163 removes batch-specific Office reference copying from the comparison path.
The existing command accepts multiple cache directories, checks source/mode/PDF
pins and pinned export-log hashes, and rejects conflicting exports before any
rendering stage. It records the chosen reference directory and provenance hashes;
identical duplicate exports remain reusable.

Twelve orchestration checks pass for missing/stale pins, source/mode/PDF/log
mismatches, conflicting exports, identical duplicates and retained failed gates.
Ten actual C162 outcomes reuse two original Office caches with exact PDF hashes,
metrics, failures and thresholds. No library build, Office export, conversion,
rasterization or production-suite replay occurs; source/tests retain C162
identity. Root2a016618; artifacts/composition-c1/reference-cache-reuse-evidence.json.
C162 CI39078220251 has green Ubuntu and pending Windows. Next, use retained
Word/source geometry to investigate nested border pitch without fitted offsets.

C162 projects one complete guarded nested body table through the existing
Review origin. Ordinary text converts its nominal inset, while measured
mathematical/picture frames translate with their container exactly once at each
depth. All descendant lines must survive shared ancestor clips before any
replacement is published. Nested border pitch, exact clips and fragments retain
explicit fallbacks; qualified flat-table geometry remains.

Thirty-eight Windows and30 Ubuntu checks pass, with eight native skips. Sixteen
independent Word outcomes/eight sources/16 pages retain32 formulas,16 pictures
and six internal links. Mathematical Y differs by at most0.218pt and named
surrounding prose by0.452pt, correcting the inherited Review2.654pt displacement.
All seven changed Review MAEs improve; nine PDFs retain accepted bytes, all16
strict visual gates remain failed and no failed component is added. Four changed
unique region pairs were reviewed; three navigation Review pairs carry both PNG
identities. Both Word/occupied-slot observers fail on C161. Horizontal mathematics
up to0.916pt, absolute destination coordinates (ReviewY0.750pt), borders,
typography and connector residuals remain unaccepted.

C161 CI39077405467 passes Ubuntu; Windows3421pass/1fail/16skip exposes an old
blanket table-note exclusion previously masked by invalid occupied metadata.
A real1pt source clip now tests exclusion, and a complete-cell positive observer
keeps numbering separate from typography eligibility. One production build runs
36 exact methods; harness-only corrections preserve14 final runtime assets and
2,078 inputs. Thirty-seven Windows passes carry and only the corrected new
start-four observer reruns; Ubuntu runs only the two added/corrected checks.
Original CI and development failures remain retained.

A source audit of1,455 tracked DOCX inputs finds seven matching nested-body
sources, all included plus the direct control. C158's fixed12/7/20 carries source
scope and report/deck smokes. Sixteen pinned Word references/PNGs reuse without
exports/reference raster replay. Seventeen assets/four exact smokes transfer
without root build or suite/corpus replay. Rooteee789fb; evidence
artifacts/composition-c1/review-nested-placement-evidence.json. The infrequent
workspace scan records12.075GB/631,514 files, zero errors; the loose5GB target
remains unmet and no rejected retirement is retried. Next: simplify reusable
reference lookup, then measure nested border pitch before changing height rules.

C161 keeps ordinary scaled cell occupied bounds in physical coordinates. The
old baseline inset could exceed its measured height; independent Word controls
also demonstrate Review note rectangles2.44-3.09pt too high. Scaling that inset
corrects metadata and invisible rectangles while preserving painting and purpose
guards. Six independent Word outcomes/three sources retain all six note links;
all rectangle edges are within0.484pt, pure-font controls within0.058pt.

Thirty-three Windows and23 Ubuntu checks pass, with ten native skips. One frozen
artifact qualifies30 exact methods then only three legacy neighbors found during
source audit; the final capability includes all33 for one future invocation.
Thirty-two metric/scale/depth combinations retain valid bounds. Ten layout
contexts preserve26 placed lines and all row/cell geometry and offsets; eight
invalid occupied bounds are corrected. Three Final PDFs retain bytes; three
Review PDFs change invisible rectangles only. All six rasters and metric sets
retain exact parent values; all six strict visual gates remain failed. Six page
pairs were reviewed in three contacts. Both independent parent observers fail.

A read-only audit of1,455 tracked DOCX sources classifies19 body-cell note
combinations:16 use ordered math frames, two are new ordinary controls and one is
an older ordinary control. All three affected sources are independently checked.
Two legacy Word outputs were refreshed because old unpinned mode0 exports do not
establish current mode2 provenance; originals remain. C158's fixed12/7/20 carries
source scope and report/deck smokes. Seventeen assets/four exact smokes transfer
without root build or suite/corpus replay. Root5f8155a9; evidence
artifacts/composition-c1/scaled-cell-struts-evidence.json. C160 CI39075501306 is
green. The visible nested-math Review2.654pt offset remains unaccepted; next,
investigate shared ordinary/nested placement rather than adding fitted offsets.

C160 keeps ordinary body-cell bookmark destinations bound to their placement
owner, retaining the existing three-point viewport context. The old walk checked
nullable local story metadata after dropping the actual owner. Mathematical
target handling and painting remain unchanged. Six independent Word outcomes
exercise links within, into and out of nested content; all 12 formulas, six
pictures and six internal links survive, including neutral source transformations.

Twenty-four Windows and 19 Ubuntu checks pass, with five native skips. The
synthetic observer now uses emitted cell origin; its earlier margin assumptions
were contradicted by Word. Fourteen runtime assets/2,073 inputs match, 23 Windows
passes carry and only that observer reruns. Initial failures remain retained.
Four PDFs retain parent bytes; only two outgoing destinations change. All six
candidate rasters and visual metric sets retain exact values; all six strict
gates remain failed. Three unique region pairs were reviewed, three others carry
by both PNG identities. Final outgoing X differs by 0.48pt and both source-relative
context gates pass within 0.5pt. Raw formula-target and Review coordinate residuals
remain unaccepted. The independent context observer fails on C159.

A read-only audit of 1,453 tracked DOCX sources finds only the new outgoing source
in the changed body-cell target combination. C158's fixed 12 outcomes/7 sources/
20 pages carry scope and exact report/deck smokes; ten DOCX failures/two deck
passes/two report extraction flags remain. Seventeen assets/four exact smokes
transfer without root build or suite/corpus replay. Root9b6d06e6; evidence
artifacts/composition-c1/nested-navigation-evidence.json. C159 CI39074500899 is green
on both platforms. Next: establish the cause of Review nested-table frame residuals.

C159 admits complete fitting nested endnote blocks through shared source and
clip ownership. Every measured line must survive all ancestor clips, so a visible
first line cannot hide a clipped continuation. Original paragraph identities,
source uniqueness, row/revision/property guards and conservative whole-block fit
remain; typography and continuation policies keep their previous scopes.

Fifty Windows and 45 Ubuntu checks pass, with five native skips. A harness-only
assertion proves the partial-clip control retains both visible and excluded lines:
14 runtime assets/2,068 inputs match, 49 passing methods carry and only that method
reruns per platform. The final transferred artifact contains the stronger check.
Six retained independent Word outcomes require no new export or reference raster.
Both changed document-average MAEs improve; four control PDFs retain bytes and all
six strict gates remain failed. Note/body Y differs by at most 0.065pt, mathematical
X/Y by 0.057/0.063pt, correcting the inherited 8.063pt offset and later story gaps.
Two Final page pairs were reviewed; two Review pairs carry by both PNG identities.
The unchanged half-point boundary observer fails on C158.

Seventeen assets/four exact smokes transfer without root build or suite/corpus
replay. C158's fixed 12 outcomes/7 sources/20 pages carry source scope and exact
report/deck smokes; ten DOCX failures/two deck passes/two report extraction flags
remain. Rootdd74b523; artifacts/composition-c1/nested-endnote-boundaries-evidence.json.
Next: probe navigation alongside restored nested content before changing guards.
The infrequent workspace scan records 11.95GB/624,124 files, zero errors. The loose
5GB target remains unmet; no previously rejected retirement was retried.

C158 restores ordinary nested cell mathematics through shared source ownership,
measurement and painting. One prepared map follows guarded source blocks across
two body depths, headers/footers and notes. Containing mathematical rows remain
atomic; whole-cell note fitting and typography guards retain their prior scopes.
The static probe also demonstrates an inherited table-only footer below the page.
The complete measured block now bottom-anchors at the footer distance, and nested
rows join the existing physical static-frame map. Independent Word origins are
within 0.5pt in both modes, correcting prior errors of about 54pt and 221pt.

One final Release artifact passes 56 Windows and 39 Ubuntu checks, with 17 native
font skips. Ten independent Word outcomes/five sources/14 pages retain 24 formula
placements, pictures and all note destinations. All ten whole-document mean MAEs
improve without new failed components; all ten strict visual gates remain failed.
Fourteen page/region pairs were reviewed. Final body math Y differs by at most
0.029pt and footnote Y by 0.082pt. Review body Y2.654pt and nested endnote Y8.063pt,
plus border/typography differences, remain unaccepted. Content and footer-origin
observers fail independently on their respective parent artifacts.

This shared-frame milestone measures the fixed 12 outcomes/7 sources/20 pages
once on the final artifact: every PDF equals C151, so its Office comparisons and
20 reviewed pairs carry by identity. Ten DOCX failures/two deck passes/two report
extraction flags remain. Source inventories identify no older tracked input with
the expanded combination. Seventeen assets/four exact smokes transfer without a
root build or suite/corpus replay. Root03fc252e; evidence
artifacts/composition-c1/inside-nested-static-evidence.json. C157 CI39072934179 is
green on both platforms. Next: complete nested endnote boundary ownership using
retained Word evidence, preserving fitting/clip/revision/continuation fallbacks.

C157 preserves supported direct mathematical paragraphs beside ordinary nested
tables. Shared cell properties admit the direct owner; authoritative BodyElements
prevent stale paragraph preparation. Picture and nested placement reuse one
occupied-frame map. Whole-cell fitting/placement guards remain, and mathematics
inside the nested table is still an explicit content gap for the next slice.

One Release artifact qualifies34 Windows/25 Ubuntu passes/nine native-font skips.
Eight independent Word outcomes retain two formulas; four affected comparisons
improve, four PDFs retain parent bytes, no new failure components, all eight strict
gates remain failed. Eight top-page regions were reviewed. Final math Y differs
by at most0.029pt; absolute X0.917pt and Review Y2.510pt remain unaccepted. The
inside-nested control retains both missing formulas and its previous PDF bytes.
Parent content observer fails expected2/got0. Initial development failures and
raw residuals remain retained; existing visual thresholds are unchanged.

A read-only inventory covers1,446 tracked DOCX files/zero skips; only the two new
adjacent-table fixtures match the changed source combination. Fixed12 outcomes/
7 sources/20 pages carry scope and exact report/deck smokes; their ten DOCX visual
failures/two deck passes/two report extraction flags remain.17 assets/four exact
smokes transfer without root build or suite/corpus replay. Root118b9188; evidence
artifacts/composition-c1/nested-cell-neighbors-evidence.json. Next: bounded source
and measured-frame ownership for ordinary inside-nested mathematics.

C156 fixes declared run-default ownership across all paragraph containers. Word
uses nominal10pt when rPrDefault is present without a size, even for an empty
container; the parent used12pt. The shared cascade now supplies the owned fallback
beneath table/style/direct properties. Unowned12pt, explicit default11pt and direct
15pt remain; paragraph13pt precedence is covered. No family or leading correction.

One Release artifact passes30 Windows/29 Ubuntu checks/one native skip.24 independent
Word outcomes cover two families/both presets/body/cell/header/footer/footnote,
with144 visible role-size observations and max0.04pt PDF size quantization. Eight
comparisons improve/16 parent PDFs remain exact/no new failed components; all24
strict raster gates remain failed. Twelve Final labeled-region pairs reviewed;
twelve Review pairs carry by both PNG identities. Labels and the custom note
source/target survive; no full word/order audit. Parent visible-size observer fails.

Read-only bounded inventory of1,442 tracked DOCX inputs has zero skips and only
four new owner cases affected. No older fixture/fixed source has unresolved owned
size; all five fixed DOCX sources explicitly declare11pt. Report/deck smokes match,
so prior C15112 outcomes/seven sources/twenty pages remain separately qualified,
with ten DOCX failures/two deck passes/two report extraction flags; no corpus replay.
The retained independent Aptos default-owner trigger now takes27 lines like Word27,
first baseline within0.034pt/rule0.082pt. This is an additional frozen-layout
measurement, not another visual pass. Seventeen assets/four integration smokes
transfer without root build/suite replay. Rootfbff99f0;
`artifacts/composition-c1/run-default-owner-size-evidence.json`. C155 CI is green.
Next: probe supported mathematics adjacent to ordinary nested tables before
changing source preparation or nested-frame eligibility.

C155 makes complete fitting endnote blocks own source paragraph boundaries.
Independent plain/image-removal/paragraph-joining controls isolate8 latent
separator points and3 legacy points per subsequent story. Placement now uses
exact first/last owners, measured fit, existing simple-cell/occupied-clip rules
and unique paragraph identities. Typography and continuation admission remain
separate. Compound overflow, scaled/static/current-footnote frames, unsupported/
nested/clipped cells, keeps and unknown spacing retain fallback. No fitted offset
or font/size branch is added.

Final Release artifact passes127 Windows/95 Ubuntu checks/32 native skips. An
initial126/1 report exposed a repeated paragraph object; the source-ambiguity
correction preserves every old fallback assertion. Seven independent Word
outcomes now have note Y max0.063pt/math X0.136/Y0.063pt. Eleven actual PDFs yield
six improved comparisons/five parent identities/no new failed components; all
strict raster gates remain failed. Ten Final page pairs reviewed/four affected
Review pairs carry by both PNG identities. Eight footnote-control pairs carry
from C154. All11 final PDFs match development outputs, so comparisons/reviews
carry without raster replay. Content/pictures/four targets survive; no full
word/order audit. Parent independent boundary observer fails on C154.

Seventeen assets/four integration smokes transfer without root build/suite replay.
Fixed notes are footnotes or the static complete report, so the new fitting
purpose does not enter; report/deck smokes match. Prior C15112 outcomes/seven
sources/twenty pages remain separately qualified, with ten DOCX failures/two deck
passes/two report extraction flags. Rootc44cc806;
`artifacts/composition-c1/fitting-endnote-boundaries-evidence.json`. C154 CI is green.
Next: investigate declared run-default ownership without a size. The retained
independent styles counterfactual uses Word9.96pt ordinary/15pt explicit text,
versus native12/15; no leading correction is justified by that observation.

C154 restores two formulas omitted when supported mixed note content enters a
simple footnote/endnote table. Normal note tables now join preparation through
exact source owners and existing simple-cell rules; related-table frames receive
the prepared map for shared measurement/placement/painting. Nested/rich-cell,
continuation, page-input, typography and counter fallbacks remain unchanged.

One Release artifact passes56 Windows/33 Ubuntu checks/23 native skips. Source
splitting/bookmarks retain content and targets; source/prepared cell identities,
unsupported stories/cells and cancellation are covered. Eight independent Word
outcomes: four table comparisons improve/four direct PDFs retain parent bytes;
no new failed components, all strict gates still fail. Footnote math origins max
0.057pt X/0.097pt Y. Endnote origins retain8.063pt Y displacement already visible
in unchanged direct controls; later note gaps also remain approximate.

Eight Final page pairs reviewed; eight Review pairs carry by both PNG identities.
Two formulas/one image/four distinct destinations and note prose survive each
outcome. Alphabetic inventories equal, not a complete word/order audit. Parent
regression fails expected2/got0. Seventeen assets/four integration smokes transfer
without root build/suite replay. Fixed note parts have no tables; the complete
report's two direct endnote equations retain their path/report smoke PDF. Prior
C15112 fixed outcomes/seven sources/twenty pages remain separately qualified;
ten DOCX failures/two deck passes/two extraction flags remain, no corpus replay.
Root5a9fe26a; `artifacts/composition-c1/note-table-composition-evidence.json`.

C153 CI39068586597 has both platform jobs completed successfully. Workspace scan
at21:00UTC reports11.763GB/608,332 files/zero errors; the loose5GB target remains
unmet. Rejected retirement targets remain protected; no reclamation is claimed.
Next: isolate trailing-body/separator/story-boundary ownership behind the endnote
displacement before extending any qualified flow policy.

C153 closes two retained Aptos fixture failures without a production metric change.
The original style-less sources resolved to Times New Roman before the supplied
Aptos measurer fallback; synthetic separator marks also own document defaults.
Both controls now share their exact original public source with an explicit Aptos
document-default test resolver and a complete font-owner assertion. Fresh Word
retains24 head lines/last21/separator682.54/first664.18. Qualified native output is
24/21/682.520/664.119, with every original numerical assertion unchanged.

Only the two corrected methods run: Windows2 passes/Ubuntu2 native skips. Fourteen
runtime assets/2,029 prior inputs are exact, and the106 prior qualified methods'
helpers are unchanged. C152106 Windows/93 Ubuntu passes/13 skips carry; combined
108/93/15. No library/CLI build, passing suite or rendering replay. Seventeen
assets/two integration smokes transfer without root build. Rootdf981703;
`artifacts/composition-c1/aptos-head-fixture-evidence.json`. C152 CI39067792661 is green.

A separate styles/defaults counterfactual changes Word paragraph flow to27 head
lines. It is retained, not used to qualify the original24-line control. Existing
visual failures, portfolio scopes and extraction flags stay explicit. Next:
probe equations/images/bookmarks inside note tables before changing admission.

C152 restores capacity when body notes and a cell note coexist. Four independent
Word outcomes confirm correct parent page counts/counters but63.247-73.488pt of
unused first-page capacity. Measured body-line registration now coexists with
complete table notes' existing block reservation. Exact paragraph/run ownership
survives multiple cell paragraphs; pure-table/nested/keep/complex-note/frame
fallbacks remain. No font/size rule or fitted offset is added.

C151 CI39066861742 caught14 missed checks. Ordinary occupancy had reused metadata
that also selected qualified body hyperlink padding and minimum-leading policy.
C152 publishes explicit measured occupancy instead, preserving existing painting
and all old assertions. Its106 Windows/93 Ubuntu checks/13 native skips include
all14 failures unchanged. Final qualification uses one frozen artifact/one exact
set per platform; development qualification and failure logs remain retained.

All four Word comparisons improve without new failed components; all strict
visual gates remain failed. Markers max0.187pt Y, mathematical origins max0.165pt X/
0.051pt Y. Five Final page pairs reviewed; five Review pairs carry by both PNG
identities. Eight fractions/four pictures/four same-page destinations and authored
105/108 survive. Alphabetic inventories equal, not a full word/order audit. Parent
half-point geometry observer fails. Twelve pure-table PDFs retain accepted bytes.

Seventeen assets/four integration smokes transfer without root build/suite replay.
Prior C151 fixed twelve outcomes/seven sources/twenty pages remain separately
qualified: fixed inputs have no cell-note references, report/deck smokes match.
Ten DOCX gate failures/two deck passes/two extraction flags and two Aptos baseline
failures remain; no new full corpus/content audit. Root2abe15c2;
`artifacts/composition-c1/mixed-container-admission-evidence.json`.
Next: investigate the pinned Aptos continuation-head residuals before changing
metrics or thresholds. Release preparation stays deferred at0.1.5.

C151 completes the ordinary-cell measurement contract. Four parent outcomes had
visible note glyphs but no occupied line bounds, leaving105-108 instead of Word09/10.
Legacy cell lines now publish actual height/baseline inset from their existing
placement measurements. Painting and container placement stay unchanged; shared
C150 source/clip ownership admits the complete slots. No font/size rule or model
flag is added. Missing and clipped measurements retain their existing fallbacks.

87 Windows/75 Ubuntu checks/12 native skips include slot containment, larger first
runs and existing note hit-area/before/after-spacing contracts. Two new observers
initially invoked an equation-only bookmark helper. The corrected helper preserves
all passing math arguments:14 runtime assets/2,025 inputs exact, only failed
ordinary observers rerun. No library/CLI rebuild, passing suite or conversion
replay after this harness correction.

Four independent Word outcomes restore source/story09/10 labels and four distinct
same-page targets, marker Y within0.187pt. All four comparisons improve/no new
failed components; strict typography/spacing gates still fail. Four Final page
pairs reviewed/four Review pairs carry by both PNG identities. Parent label
regression fails. Seventeen assets/four smokes transfer without root build/suite
replay. Final twelve portfolio outcomes/seven sources/twenty pages and eight
mixed-cell PDFs retain accepted bytes. Office comparisons and reviewed ink carry;
ten DOCX failures/two deck passes/two extraction flags and existing Aptos baseline
failures remain. No new full content audit. Root68ca1278;
`artifacts/composition-c1/plain-cell-line-frames-evidence.json`. C150 CI is green.
Next: investigate complete-note admission when a long body paragraph and cell note
coexist, before changing reservation or any validated container fallback.

C150 fixes another composition boundary: moving supported mixed paragraphs into
cells left provisional105-108 instead of page counters09/10. Numbering now shares
body/cell source order and complete measured line/paint-clip ownership. One
paragraph replacement is published through ordered items, references and both
cell views; derived cell text updates, authored runs and unchanged manual-mark
containers retain their identities. Typography keeps its narrower map. Missing
frames, clipped/ambiguous/merged owners, structural revisions and unstable pages
retain existing fallbacks; legacy ordinary-cell measurements remain staged.

One frozen artifact passes70 Windows/59 Ubuntu checks/11 native skips. Eight
independent Word outcomes restore automatic/custom counters with eight equations/
four image draws/four destinations and authored105/108 intact.64 math operations
stay exact to parent. All eight comparisons improve/no new failed components;
strict typography/spacing gates remain failed. Eight Final page pairs reviewed;
eight Review pairs carry by both PNG identities. Parent label regression fails.
Missing/duplicate/clipped/unmeasured/merged owners reject atomically; source/view
identity, unchanged manual targets, no-op and cancellation regressions pass.

Seventeen assets/four smokes transfer without root build/suite replay. The fixed
DOCX portfolio has no section eachPage request, so C148's12 outcomes/seven sources/
twenty reviewed pages carry by source scope/report/deck smokes; no full replay or
content audit. Ten DOCX failures/two deck passes/two extraction flags and two Aptos
baseline failures remain. Rootf7f5a9b1;
`artifacts/composition-c1/cell-page-numbering-evidence.json`. C149 CI is green.
Next: measure ordinary cell-line ownership against Word before publishing more
frame metadata or changing note annotation slots.

C149 challenges the accepted library with eight fractions/four images/four
bookmarks/four automatic notes inside one naturally paginated paragraph. Four
independent Word outcomes match Calibri's three pages/Arial's two pages and every
note's labels/page membership. Math origins differ by at most0.165pt X/0.051pt Y;
marker Y differs by at most0.187pt. Five Final page pairs reviewed; five Review
pairs carry by both raster identities. All four strict visual gates still fail
spacing/typography thresholds. No visible content loss, overlap or clipping was
demonstrated on these pages; arbitrary-document parity is not established.

No production change is needed. Only the harness builds:14 runtime assets/2,012
prior inputs stay exact, four inputs added. Seven new/helper-neighbor Windows
methods pass; Ubuntu passes three/skips four native checks. Neutral bookmarks and
unchanged-format run splitting retain text, page counts, note rectangles/targets
and mathematical glyph placement. Combined103 unique methods carry101 Windows
passes/two unchanged Aptos failures and92 Ubuntu passes/11 native skips. No library/
CLI build,99-check replay or conversion replay. An initial local-name compile error
was corrected before freezing or tests; no failed artifact was published.

Seventeen assets/four integration PDFs transfer without root build/suite replay.
C148's fixed12 outcomes/seven sources/twenty reviewed pages and eight earlier
compound controls carry by unchanged runtime/source scope/report/deck smokes.
Ten DOCX failures/two deck passes/two extraction flags and two Aptos baseline
failures remain; no new full content audit. Roota6299c85;
`artifacts/composition-c1/spanning-note-math-evidence.json`. C148 CI is green on both
platforms. Next: independently probe mixed paragraph note ownership in table cells.

The infrequent workspace check measures11.539GB/592,016 files, zero errors. The
loose5GB target remains unmet. Automatic approval review rejected retiring two
superseded development compiled snapshots before process creation ("blocked by
policy"); zero files removed. Inputs/useful evidence and both snapshots remain,
with68 snapshot names protected; no alternate retirement attempt is made.

C148 removes an extra Calibri page and large unused body regions in natural
paragraphs with multiple notes. Block-wide registration reserved future markers
and repeated separators (161.1875pt on page one). Complete one-line body notes
now register at their measured source fragments and share one separator. Completed
active records retire per page; the coordinator skips its redundant reserve retry.
No font/size/count rule is added. Split/complex stories, section/column changes,
keep chains, static frames and unowned source markers retain existing paths.
Preflight skips documents without footnotes and checks cancellation during scans.

All four independent Word controls now agree on two pages, note membership and
09/10 labels. Source/target marker geometry is within0.147pt X/0.283pt Y. Four
comparisons improve; all strict typography/spacing gates still fail unchanged
thresholds. Four Final page pairs reviewed/four Review pairs carry by raster
identity. Four destinations/authored digits survive; alphabetic inventory does
not establish complete word extraction or reading order. Parent geometry fails.

Final qualification:97 Windows passes/two unchanged Aptos failures;90 Ubuntu
passes/nine native skips. Both Aptos continuation-head failures reproduce the
same details on C147; they remain failures. An unexpectedly slow direct Windows
run retains40 completed passes and resumes only59 unreported methods with the
usual isolated temporary directory. No completed final passing method is replayed.
The late cancellation/fast-path source changes justify final-artifact qualification.

Final twelve portfolio outcomes/seven sources/twenty pages and eight compound
math/image/bookmark/custom-note PDFs retain accepted bytes. Office comparisons
and reviewed pages carry by identity: ten DOCX failures/two deck passes/two report
extraction flags remain. No new full content audit. Seventeen assets/four smokes
transfer without root build/suite replay. Root0ba654ef;
`artifacts/composition-c1/complete-note-admission-evidence.json`. C147 CI is green.
Next: mixed formulas/images/bookmarks in one naturally paginated note paragraph;
prefer harness-only qualification if the accepted library already composes them.

C147 fixes numbering for one paragraph spanning physical pages. Four parent
outcomes kept105-108 although every marker already shared its complete note's
page. Numbering now resolves each measured glyph/note owner independently of
other paragraph lines. Typography retains the same-paragraph-page guard; source/
target completeness, cancellation and one bounded stable-page retry remain.

51 Windows/45 Ubuntu checks/six native skips. Initial49 passes/two harness failures
are retained; two corrected observers rerun with14 unchanged runtime assets and
2,012 exact inputs. No library rebuild or passing Windows test replay after the
harness correction. Four independent Word comparisons improve/no new components;
all strict gates still fail. Five Final candidate pages/four Word page pairs plus
the extra Calibri page reviewed; Review carries by Office/candidate PNG identity.
Four destinations/authored runs survive. Alphabetic inventory matches Office;
word splitting/reading order remain extraction limitations.

Arial's two-page note membership matches Word. Calibri still produces three pages
versus two, and both fonts leave large empty body regions before page breaks.
These confirmed pagination/capacity defects remain C148's priority; this numbering
fix does not establish document parity. Exact parent regression fails. Seventeen
assets/four smokes transfer without root build/suite replay. Root29ae2905;
`artifacts/composition-c1/spanning-page-numbering-evidence.json`. C146 CI is green
on both platforms. Next: establish body admission/reserve ownership before changing
natural-paragraph continuation geometry.

C146 fixes automatic page numbering beside a custom mark. Four parent outcomes
kept105-107 instead of09/09/10. Complete measured custom source/target glyph pages
now participate in ownership, while only automatic references consume counters.
Authored marker identities and target paragraphs stay unchanged. Numbering has
its own stable-page check; typography retains its stricter map and baseline-only
font-target admission. Cancellation reaches the coordinator and page mapper.

One artifact passes127 Windows/117 Ubuntu checks/ten native skips. Same-paragraph
and separate-paragraph references work; invisible/mismatched targets reject
without partial changes, and correct labels avoid another retry. Four independent
Word outcomes agree on custom/automatic body+story labels. Eight formulas/four
image draws/authored digits/four destinations survive; math operations stay exact.
Four comparisons improve/no new failed components; all strict gates remain failed.
Four Final page pairs reviewed/four Review pairs carry by raster identity. Parent
fails the independent label regression.

Seventeen assets/four identical smokes transfer without root build/suite replay.
The fixed portfolio carries by no-custom source scope/report/deck smokes; no new
twelve-outcome replay or full content audit. Rootdab4c48b;
`artifacts/composition-c1/mixed-custom-page-numbering-evidence.json`. C145 CI is
green on both platforms. Next: investigate paragraph-spanning marker ownership
before changing the retained same-page paragraph guard.

C145 fixes another confirmed composition omission: custom footnote/endnote marks
suppressed adjacent one/two fractions in all eight parent outcomes. Exact source
run/offset/first-character indexing is shared with preparation; qualified note
references use ordered placement. Unique owners are indexed once, cancellation
stays checked, and the existing plain font-target admission remains unchanged.
Revised, missing, duplicate or invalid marker sources retain fallback.

One artifact passes123 Windows/114 Ubuntu checks/nine native skips. Eight independent
Word outcomes restore math, both custom glyphs, bookmarks, note prose and four
distinct destinations. Source splitting and *tail preserve first-character link
width. All eight comparisons improve/no new failed components. Math origins are
within0.136pt X/0.020pt Y; strict foreground-recall/SSIM gates still fail. Eight
Final page pairs reviewed; eight Review pairs carry by raster identity. Parent
fails the independent content regression.

Seventeen assets/four matching smokes transfer without root build/suite replay.
The fixed DOCX portfolio has no custom markers; prior milestone results carry by
source scope/report/deck smokes, no twelve-outcome replay or full content audit.
Rootfb70f24a; `artifacts/composition-c1/custom-note-math-evidence.json`.
C142-C144 CI is green on both platforms. Next: reproduce whether a custom mark
can improperly disable automatic eachPage counters in neighboring notes.

C144 fixes a metadata-only composition omission. Explicitly declaring matching
document-level footnote page restarts suppressed all eight equations in C143's
compound paragraphs. Automatic footnote page dependencies now retain actual source
run ownership; numbering owns measured label resolution and stable-page acceptance.
Endnote page restarts, unknown/custom markers and unstable owners retain fallback.

One artifact passes110 Windows/102 Ubuntu checks/eight native skips. Four new Word
outcomes restore eight equations/four image draws/correct body+story labels each.
All four candidate PDFs equal the equivalent accepted C143 compounds; fresh Word
rasters also equal C143. Eight prior visual page pairs carry by identity, without
review replay. All four parent comparisons improve/no new failed components;
all strict gates remain failed. The parent fails the independent content regression.

Seventeen assets/four smokes transfer without root build/suite replay. Fixed
portfolio evidence carries by source scope and unchanged report/deck smokes;
no new twelve-outcome replay or full content audit. Root4b525e8c;
`artifacts/composition-c1/explicit-page-note-source-evidence.json`. An infrequent
workspace scan measures11.376GB/581,812 files, zero errors; the loose5GB target
remains unmet and protected retirements remain untouched. No reclaimed space is
claimed. Next: custom note-marker source ownership alongside supported equations.

C143 challenges the accepted numbering library with compound source content:
each paragraph has two fractions, one inline PNG and an enclosing bookmark.
Both font controls/both modes retain eight formulas/four image draws, correct
body/story labels, authored digits and four distinct note destinations. Four
independent Word outcomes preserve all content; four Final page pairs reviewed,
four Review pairs carry by PNG identity. All four strict gates remain failed
with typography/wrapping residuals. No production eligibility change is needed.

Only the test harness builds: three test assets change, fourteen runtime assets
and all1,990 prior inputs stay exact. Five affected Windows checks pass; Ubuntu
passes three/skips two native checks. The other100 Windows/96 Ubuntu/four skipped
checks carry from C142 by unchanged library, source scope and input identities.
Combined105 Windows/99 Ubuntu/six native skips. No library/CLI build,103-check
replay or conversion replay. Seventeen assets/four smokes transfer without root
build/suite replay. Rootaf0789da;
`artifacts/composition-c1/compound-page-notes-evidence.json`. The fixed portfolio
and retained flags carry unchanged; no new full audit. Next, reproduce explicit
document-level page-restart source metadata before changing its fallback.

C142 fixes numbering when supported equations accompany section-level page
restarts. Math was already painted, but the plain-body gate left provisional
labels105-108. Measured numbering now owns complete marker/note page maps
independently of typography. Shared settings/static/frame/source guards remain;
one bounded retry requires stable pages. Changed marker runs retain one identity
across paragraph runs, references and ordered source items. Partial source slices
and unsupported explicit restart metadata keep fallback. Typography admission
is unchanged; no new font/size condition is added.

One artifact passes103 Windows/98 Ubuntu checks/five native skips. A tool-only
selector correction precedes selection; no rebuild or passing check replay.
Four independent Word outcomes agree on all body/story09/10 labels; four formulas,
authored105/108 digits and four distinct note destinations survive. Math operations
are exact to the parent. All four comparisons improve/no new failed components;
all strict gates still fail. Four original no-math PDFs are byte-identical. Four
Final page pairs reviewed/four Review pairs carry by raster identity. The parent
fails the independent label regression.

Seventeen assets/four smokes transfer, no root build/suite replay. The fixed
portfolio carries by absence of page-restart metadata and unchanged report/deck
smokes; no new twelve-outcome replay or content audit. Root12a7dbb1;
`artifacts/composition-c1/page-note-owned-content-evidence.json`. C140/C141 CI is
green on both platforms. Next: compound numbering/source-ownership transformations.

C141 shares the existing terminal-paragraph boundary with note flow. Independent
Word counterfactuals agree for implicit, explicit empty/default/zero and short
trailing paragraphs. The already represented TerminalParagraphMark owns zero
after-spacing; earlier positioned table rows expose page-space bounds. No source
run is invented. Known body/static/scale/column/source/slicing fallbacks remain.

Seventy-one Windows/62 Ubuntu checks pass/nine native skips on one artifact.
Eight outcomes reuse two Word references/add six exports. All eight comparisons
improve without new failed components; all strict gates still fail. Original
table note Y error17.059pt becomes0.059pt across all four shapes. Mathematical
operations stay exact, eight fractions/four distinct note destinations survive.
Eight Final page pairs reviewed; eight Review pairs carry by PNG identity. The
parent fails the independent position observer.

Both composite-report PDFs are actually reconverted and stay exact to C140.
The remaining portfolio carries by source scope and the held-out deck smoke;
no new twelve-outcome replay or content audit. Seventeen assets/four smokes
transfer without root build/suite replay. Rootdcda71df;
`artifacts/composition-c1/terminal-table-note-flow-evidence.json`. C140 CI remains
in progress; work continues while it runs.

C140 separates note static-frame validation from unrelated body contents while
retaining shared final-section/mirror/floating guards and existing static, scale,
source and slicing fallbacks. Complete endnote paragraphs lose the legacy3pt gap.
The body/note boundary charges only measured occupied depth beyond ordinary text
descent. Missing/malformed/exact frames keep their existing charge. This closes
the mixed-body note displacement without a font, size or document-shape branch.

Sixty-nine Windows/61 Ubuntu checks pass/eight native skips on the final artifact.
The first development admission widened header fallbacks and fails three unchanged
observers on both platforms; that artifact never transfers. Static validation is
restored before the measured-depth correction. Every production revision reruns
its selected tests; no passing suite is replayed on root. The final body note Y
error falls6.284pt to0.051pt, paragraph pitch agrees within0.009pt, and mathematical
operations stay exact. Two visual comparisons improve/six PDFs stay exact, no new
failed components. All eight strict gates still fail; terminal-table note error
17.059pt remains. Changed page pairs reviewed, unchanged peers carry by identity.

The final milestone converts the fixed twelve outcomes/seven sources/twenty pages
once with the accepted artifact. Every PDF stays exact to C136; twelve comparisons
and twenty earlier page reviews carry. Ten DOCX failures/two deck passes and two
complete-report extraction flags remain. Seventeen assets/four smokes transfer,
no root build or suite replay. Rootada3b6f7;
`artifacts/composition-c1/endnote-owned-frame-evidence.json`. C138/C139 CI is green
on both platforms. Next: independent terminal-table boundary counterfactuals.

C139 fixes confirmed equation loss beside section-restarted footnotes/endnotes.
The reader already resolves section counters; shared marker ownership now admits
those labels. Page-dependent and unknown restarts retain fallback. No font, size
or document-shape condition is added. Four two-section body/cell controls cover
decimal/Roman starts, run splitting and four distinct note destinations in both
DOCX modes. Thirty Windows and24 Ubuntu checks pass/six native skips, once each.
A missing namespace import caused an initial harness build failure; no artifact
was frozen. An Ubuntu invocation first supplied a report instead of a method
array; no tests ran until the corrected invocation. These records remain retained.

Eight independent Word outcomes restore all eight fractions each; all eight
visual comparisons improve without new failed components. Math origins are
within0.465pt X/0.025pt Y, markers0.069pt Y, footnote prose0.087pt Y. All eight
strict gates still fail. Endnote prose retains6.284pt body/17.059pt cell Y errors;
this separate placement defect is the next bounded investigation. Eight Final
page pairs reviewed; eight Review pairs carry by reference/candidate PNG identity.
The parent fails the content regression with zero of eight fractions.

One committed artifact transfers17 assets/four identical smokes with no root
build/suite replay. The portfolio's five DOCX sources have no restart metadata;
prior milestone outcomes carry by source scope, plus unchanged report/deck smokes.
No new full portfolio pass is claimed. Roote0689b40;
`artifacts/composition-c1/section-note-math-evidence.json`.

C138 adds manifest-driven outcome joining for one frozen runtime. It validates
input/runtime/PDF identities and missing/conflicting outcomes before publishing,
copies independent PDFs and inspection summaries, and retains source provenance
and raw classifications. Raster and comparison evidence stays at its original
path. Twelve actual parent outcomes join with twelve comparisons reused and an
identical geometry report: no build, conversion, Office export or metric replay.
Focused orchestration checks pass; CI now exercises them on both platforms.
Root1a8c7700; `artifacts/composition-c1/checkpoint-138-evidence.json`.

The bounded inventory measures10.927GB in artifacts, including7.434GB composition,
2.011GB earlier plan work and0.145GB shared input snapshots. Compiled outputs
remain independently owned across frozen artifacts. No cache redesign or protected
retirement is attempted and no space reclaimed is claimed. Avoiding about4.8MB of
prospective report copies does not meet the loose5GB workspace target. All66 known
protected frozen names remain preserved. C137 CI39053583094 is green on Windows
and Ubuntu. Next, investigate restarted-note/equation composition with a reproducer
before classifying any suspected content loss.

C137 admits exact-row projection only when complete occupied/visible inline
frames fit the prospective cell after border advance. Tight source clips,
incomplete/overflowing frames and complex/fragmented rows retain fallback. Every
row validates before mutation; cancellation stays checked. Shared pure clip and
inline-extent calculations serve placement and painting; existing border advance
arithmetic is consolidated without changing its value.

Twelve independent controls reuse two references/add ten fit/clip/border/flow
exports. Eight comparisons improve, four deliberately clipped PDFs stay exact,
no new failed components. Fitting math errors fall3.333pt to1.441pt; balloon tops
are within0.778pt and following prose within0.186pt. All strict/full one-point
gates still fail. Twelve visual pairs reviewed, plus clipped-margin detail.
Raw clipped text operations are not visible-content proof. One default Final
clipping peer is byte-identical; no new twelve-document portfolio replay.

The first capability batch passes81 Windows/70 Ubuntu checks but exposes the
fixture's wrong expected image count for the wholly clipped24pt row. Word/parent/
candidate all omit that image. Only metadata changes:17 runtime assets/1,973 other
inputs stay exact, no library/test rebuild or conversion replay. Only two affected
observers repeat; final83 Windows/71 Ubuntu/twelve native skips. The parent fails
the independent frame regression. Seventeen assets/four smokes transfer without
root build/suite replay. Rootbc835e3d;
`artifacts/composition-c1/exact-row-frames-evidence.json`. C135/C136 CI39050721245
is green on both platforms.

Storage review confirms existing freezing already shares1,973 unchanged files/
17.6MB by file identity; only66,838B metadata is new. An input-cache redesign is
unnecessary. Retained report copies warrant a bounded inventory; prospective
copy savings are not reclaimed disk space. Protected retirements remain untouched.

C136 refreshes the fixed twelve outcomes with the accepted C135 artifact:
seven source documents/twenty pages retain accepted PDF and reference identities.
Twelve comparisons and twenty previously reviewed page pairs carry, without Office
exports, raster/metrics replay, library rebuild or suite replay. Ten DOCX gates
still fail; both decks pass. Ten earlier content screens and the complete report's
two unresolved dynamic-field/math-tokenization flags carry; no new audit pass.
Routine expected-string/ignored-math screens show no new flags.

Manifest integrity passes for715 cases/twelve families; this is not715 Office
matches. The extended reference tool reproduces twelve earlier goldens byte-for-
byte without exports/builds. Existing CI supplies broader Windows corpus, Ubuntu
portable, package and tool checks; local suites are not duplicated. C124 compound
allocation evidence stays unchanged; no new performance claim. Workspace remains
11.134GB/556,727 files with zero scan errors, above the loose5GB target. No protected
retirement is attempted. Evidence `artifacts/composition-c1/checkpoint-136-evidence.json`.
Exact-height occupied-frame fit and clips are addressed by C137 above.

C135 coalesces actual structural cell owners before packing: consecutive rows
share a range only for the same table/column span, revision kind and known author.
Dates stay source metadata without splitting the range. Same-row/diagonal spans,
gaps, authors and kinds remain separate. Geometry cannot join these distinct
source ranges. Complete-grid, unfragmented, nonmerged/nonnested owners qualify;
existing fallbacks and other modes remain. Raw source owners and candidate/revision
totals are retained, even when two cells print one label.

Eighteen independent outcomes agree on range counts and labels; four references
reuse, fourteen new exports establish boundaries. Ten comparisons improve, eight
PDFs remain exact, no new failed components. All eighteen body text/math geometry
reports remain exact to the parent and visible child content survives. Twelve
visual pairs reviewed, six carry by raster identity. All strict/full one-point
gates still fail; count agreement does not establish Office parity. Margin-bar
segmentation, dash/width/end attachment, matrix and exact-height residuals remain.
The earlier narrow-cell paint-order extraction flags stay; this slice has none.

One Release artifact passes33 Windows/22 Ubuntu checks/eleven native skips, once
per platform, with no observer corrections or repeated tests. Both new regressions
fail on the parent. Existing reference tooling pins actual Office balloon frames/
text from a declared region. Seventeen assets/four smokes transfer without root
build/suite replay. Rootfa711fc4; `artifacts/composition-c1/cell-ranges-evidence.json`.
C134 CI39049120946 is green on both platforms. Next is one finite portfolio refresh
on this accepted artifact, reusing comparisons/reviews and CI milestone checks.

C134 attaches structural cell connectors to the first printed paragraph and
its shared occupied inline frame, while cell-top placement stays independent.
The same extent helper serves existing comment ranges. Narrow/same-author/
different-author controls establish type-owned insertion/deletion stroke colors;
mixed owners retain the review fallback. No fitted offset is introduced.

Eighteen comparisons improve with no new failed components; all strict/full
one-point gates remain failed. Eight first-source controls are within0.403pt X/
0.893pt Y and match type color. All eighteen inspected text-operation reports
are byte-identical to the parent; no prose, math or label placement change.
Fourteen visual pairs reviewed, four carry by raster identity. Twelve references
reuse and six new exports suffice. Two same-author cases retain extra labels:
Word coalesces their adjacent source cells into one range, candidate emits two.
Dash/width/end attachment, exact-height and matrix residuals remain explicit.

A30-method capability batch passes28 Windows/19 Ubuntu checks initially, with
new content observers failing on wrapped prose. A whitespace-only correction
also fails because mathematical paint operations intervene between wrapped words.
Only those observers repeat after screening prose separately from independently
checked math text/count/images. Fourteen other assets/all1,947 inputs retain
identities; no library rebuild or conversion replay. Final30 Windows/20 Ubuntu/
ten native skips. Both new regressions fail on the parent. Two raw extraction
flags stay classified, not relabeled as passes; reviewed content is present.
Seventeen assets/four smokes transfer without root build/suite replay. Root5d87a1f3;
`artifacts/composition-c1/cell-connectors-evidence.json`. C133 CI39047972023 is
green on both platforms. Next is bounded source-range coalescing with adjacency,
author, type and separation controls; first-source qualification does not claim it.

C133 fixes two shared table-frame defects: declared minima stayed in source
points while exact heights already used the review scale, and baseline projection
rejected automatic/minimum declarations despite measured physical child frames.
Recognized minima now consume the existing scale once and share the automatic
projection. Exact/unknown heights and complex/fragmented cell fallbacks remain;
all rows validate before mutation and cancellation stays checked. No fitted offset
or extra font/size/document-shape guard is introduced.

Twelve independent review outcomes reuse six references and add six neighboring
height-rule exports. Eight comparisons improve; four PDFs remain exact, including
both exact-height controls. All content/ink survives. Ten automatic/minimum labels
are within0.271pt of Word; changed math errors fall to1.212pt and prose to0.539pt.
Unchanged exact-height controls retain3.061pt math/1.665pt label errors. All twelve
strict/full one-point gates remain failed; no new failed components or relaxed
thresholds. Six distinct visual pairs reviewed, six carry by raster identity.
One default Final control also remains byte-identical.

One frozen Release artifact passes78 Windows/69 Ubuntu checks/nine native skips,
once per platform. Both new regressions fail on the parent. Seventeen assets/four
smokes transfer without root rebuild/suite replay. The seven portfolio sources
contain no declared table heights; C131 carries by explicit source inference and
original report Review/deck smokes, without twelve conversions. Root18a4d363;
`artifacts/composition-c1/minimum-row-placement-evidence.json`. C132 CI39047234888
is green on both platforms. Next is source-backed connector placement: centered
body-crossing connectors, type colors and balloon frame height remain visible.

C132 assigns structural cell labels to their printed cell top, independently
of connector anchors. Six independent row-height/preceding-row controls confirm
that Word retains this owner. One preferred-top rule also preserves existing
footer-story formatting placement: two actual formatting peer PDFs stay exact.
No fitted offset or extra font/size/document-shape guard is introduced.

All six MAEs improve without new failed components; all strict/full one-point
gates remain failed. Original label error falls25.545pt to0.228pt. Tall/following
rows retain1.665pt/7.424pt label errors; declared-height review-table projection,
connector center/colors and balloon frame height remain approximate. All child
math/text/ink survives. Six distinct visual pairs reviewed. Six references were
exported, including two original-source references repeated in the probe batch;
no later exports repeat.

One 23-method capability batch passes on Windows. Ubuntu passes17 and initially
exposes an inherited installed-font absolute packing assertion, reproduced exactly
by the parent. Its final observer retains the Windows tolerance and explicitly
skips native coordinates elsewhere; portable layout/overflow/grouping remain.
A proposed relative-gap observer also fails fallback-font anchors and is removed;
all failed reports stay. Only that observer reruns. Fourteen other assets and all
1,929 inputs remain identical, so22 Windows/17 Ubuntu passing observers carry.
Final counts23 Windows/17 Ubuntu/six skips. Seventeen assets/four smoke PDFs
transfer without root rebuild or suite replay. Root73794b52;
`artifacts/composition-c1/cell-structure-placement-evidence.json`. C130/C131 CI
are green on both platforms. C131 portfolio comparisons carry by explicit source
scope and formatting peer identities, plus original report Review/deck smokes.
Next is source-backed declared-height table placement; no parity claim.

C131 refreshes all twelve fixed-portfolio outcomes with the accepted C130
artifact, without a library rebuild or test replay. Seven source documents and
twenty inspected pages retain their accepted PDF bytes and reference identities.
All twelve comparisons and twenty previously reviewed page pairs carry by identity;
no Office exports, rasterization or metrics repeat. Ten DOCX visual gates still
fail; both decks pass. Ten earlier prose/math screens and the complete report's
two unresolved dynamic-field/math-tokenization flags carry; this is no new content
audit pass. No new ignored equations or missing expected strings are observed.
Compound allocation evidence remains C124; no performance claim. Evidence:
`artifacts/composition-c1/checkpoint-131-evidence.json`. Next bounded work is
source-backed structural-label placement, retaining the25.545pt residual.

C130 represents cell insertion/deletion as distinct structure owners read from
actual cell properties. They preserve child content and ink. Main-story markers
share the printed review lane, bar and balloon selection; Word omits these markers
from headers/footers, while qualified formatting keeps its wider story scope.
One story policy governs all three decisions. Cell balloons retain actual cell
ownership instead of duplicating their row rollup.

Six sources/twelve Word outcomes retain all math. Two body review comparisons
improve; ten control PDFs remain byte-identical. Main math errors fall from91.878pt
X/76.325pt Y to below0.517pt X/1.169pt Y. Labels and marker presence are restored.
Structural labels still sit25.545pt low, with connector color/frame differences;
these significant residuals stay explicit. All strict/full one-point gates remain
failed, with no new components or threshold relaxation. Five distinct visual pairs
reviewed; seven carry by raster identity.

Four exact-source references are reused and eight new exports retained. Initial
PDF screens expose the static-story policy difference before tests. Only four
affected static review PDFs repeat; eight earlier screens carry by explicit source/
story inference. Nullable story access is corrected before artifact publication;
no prototype tests run. One final Release artifact passes120 Windows/92 Ubuntu
checks/28 native skips, once each. Parent checks fail on absent bars/labels.
Seventeen assets/four smokes transfer without root build/suite replay. Root7c5a8e64;
`artifacts/composition-c1/cell-structure-story-evidence.json`. C129 CI39043825422
is green on both platforms. Next milestone reuses this artifact for one finite
portfolio refresh; structural-label placement remains a source-frame task.

C129 selects tracked row property markers through existing markup visibility before
conditional styling, child reads and layout. Cancellation remains checked for rows
and property iteration. Actually visible deleted rows keep their independent math
owners; moving/unknown owners and deleting cell wrappers stay staged.

Four sources/eight independent Word outcomes remove deleted prose/images from all
four Final views and restore four equations in each review view. A middle deleted
row preserves both neighboring rows; header/footer controls preserve body text.
All eight MAEs improve, with no new failed components. Two strict gates pass. The
sole-row Final image is pixel-identical to Word; its raw ForegroundRecall gate
remains failed because the blank-image metric is undefined. Mixed survivor-row
typography and all four review matrix/rule gates remain failed; review math max
0.572pt X/1.778pt Y, prose2.172pt Y. Seven distinct visual pairs reviewed; one more
carries by raster identity. No numerical threshold change.

One frozen Release artifact passes108 Windows and82 Ubuntu checks/26 native skips,
once each. Reader tests cover all four API modes; Office comparisons cover the two
public presets. Parent checks fail on retained images/prose. Two exact-source Word
exports and two earlier parent screens are reused; six new exports/parent outcomes
and eight final candidate conversions complete the batch. Source-backed parent
carry stays explicit. Reusable manifests now express per-mode expectations and an
intentionally empty math view. Raw hidden-row/unused-header diagnostics stay.

Seventeen assets/four smokes transfer without root build/suite replay. The finite
portfolio carries by source inference and original report Review/deck identities.
Rootae220f68; `artifacts/composition-c1/row-visibility-evidence.json`. C128
CI39042102010 green on both platforms. Latest workspace scan10.892GB/541,429 files;
automatic review rejected retirement of two obsolete unqualified compiled snapshots
before process creation, deleting nothing. Their protected map/evidence remain.

C128 replaces blanket row/cell revision exclusions with non-deleting source-owner
rules. Supported insertions and current property geometry keep the actual child
paragraph/equation ownership. Deleting, moving and unknown kinds stay staged.
Static table formatting reuses existing revision traversal for the same printed
review lane and fitted anchor as other qualified owners.

Nine sources/eighteen Word outcomes restore all four equations in fourteen outcomes
(parent zero math operations versus21 glyphs). Four unmodeled cell-marker control
PDFs retain accepted C127 bytes. Prequalification exposes static-format review
errors of91.518pt X/136.540pt Y; the shared lane fix reduces them below0.155pt X/
1.763pt Y. No fitted offsets. All fourteen MAEs improve without new failure
components. All strict raster/full one-point gates remain failed. Changed outcomes
retain matrix/rule/placement residuals up to1.778pt math Y/2.172pt prose Y. Unmodeled
cell insertion/deletion review markers retain91.878pt X/76.325pt Y and missing bar/
lane state. Fifteen distinct crop pairs reviewed; three carry by raster identity.

Only two affected review PDFs are converted after the lane fix; sixteen prior
screens carry by explicit branch/source inference. No prototype tests run. The
final Release artifact passes100 Windows and76 Ubuntu checks/24 native skips, once
each. Four unchanged controls remain in the complete raw comparisons; new content/
ink assertions cover the fourteen changed outcomes. Parent checks fail21 versus0.
Seventeen assets/four smokes transfer without root build/suite replay. Seven exact
portfolio inputs contain no affected table revision owners, carrying C124's twelve
outcomes by source inference; original report Review/deck PDFs remain smoked.
Root6d75f2cb; `artifacts/composition-c1/table-math-lane-evidence.json`. C126/C127
CI39039390276/39039923804 green on both platforms. The independent deleted-row Final
probe confirms retained prose that Word hides; row visibility is the next fix.

C127 verifies the same authored empty-script/painted-fraction controls in cells,
headers and footers. Six sources/twelve independent Word outcomes retain all math
content and ink. Matrix row/baseline, fraction-rule and paragraph-placement gates
remain failed: math max0.574pt X/1.794pt Y, prose max2.026pt Y. Six crop pairs are
reviewed; six more carry by identical Final/Review rasters. No geometry parity claim.

Production and CLI remain exactly C126. Fourteen non-test assets and all1,891 prior
inputs match. The175 Windows/109 Ubuntu passes/66 skips carry without replay; only
three new Windows checks and two Ubuntu checks/one native skip run, once each.
An inapplicable unpainted-script rule expectation is corrected before qualification,
without repeating exports/conversions. Seventeen assets/four smoke PDFs transfer
without a root build/suite replay. Root847c3bd4; evidence
`artifacts/composition-c1/explicit-placement-evidence.json`. The original report
Review and held-out deck PDF identities remain smoked; fixed-portfolio carry stays
source-backed. Next bounded composition probe: equations in revised table owners,
confirming visible omissions before replacing any row/cell eligibility guard.

C126 makes fractions and scripts reuse the existing validated control-property owner
and actual run/font metadata. Scripts retain typed invalid-property state instead
of excluding every authored property. Unknown/malformed properties and deeper
structures keep explicit fallbacks; raw XML remains. No geometry offsets or native
code are introduced.

Eight sources/sixteen independent Word outcomes restore every omitted equation.
Four earlier exact-source exports are reused; twelve new exports complete empty/
painted controls at both sizes and both public presets. All sixteen MAEs improve,
with no new failure components. All strict raster and full one-point gates remain
failed: matrix row/baseline and rule residuals reach2.821pt Y. Eight crop pairs are
reviewed; the other eight carry by identical Final/Review rasters. Content and ink
are restored without claiming geometry parity.

One frozen Release artifact passes175 Windows and109 Ubuntu checks/66 native skips,
once each. Parent checks fail expected17 versus14 glyphs. Seventeen assets/four
smoke PDFs transfer without root build/suite replay. Seven exact portfolio sources
contain no affected property owners; C124's twelve outcomes carry by source-backed
inference, with original composed-report Review/deck PDF identities smoked. Existing
visual failures and two extraction flags stay. Root1f8f7714; evidence
`artifacts/composition-c1/explicit-control-evidence.json`. C123/C124 CI39037534189/
39038260535 are green on both platforms.

C125 inspects actual matrix/child frames without changing production. Ten independent
Word Final counterfactuals vary matrix/child control sizes and row placement. The
residual depends on both; a uniform row gap or font-size substitution is not
supported. Four C123 matrix gates remain failed, with their cause unverified.

The same probes confirm a higher-priority content defect: four ordinary explicit
fraction/script child controls lose the complete equation, while Word preserves
all three matrix glyphs. Existing parser property restrictions cause the fallback.
Next bounded change shares existing control-property validation and actual source
ownership, with independent empty/painted controls and neighboring composition checks.
Evidence: `artifacts/composition-c1/matrix-row-frame-review.json`. No rebuild, test
replay or speculative geometry correction.

C124 refreshes the finite portfolio using the accepted C123 artifact: twelve actual
outcomes, seven sources and twenty inspected pages retain every accepted PDF byte.
All reference and threshold identities match, so twelve comparisons and twenty
reviewed page pairs carry without repeated Office exports, rasters or metrics.
Ten DOCX strict gates fail and both decks pass. Ten content screens and two unresolved
complete-report dynamic-field/math-tokenization flags carry; no fresh audit pass.

The same artifact supplies four isolated compound allocation probes, one warmup
and three warm samples each, retaining all four accepted PDF identities. Warm
calling-thread totals remain35.9-40.5MiB; changes versus C117 range from minus3,120
to plus15,488 bytes. Initial font preparation remains24.5-24.6MiB with466 faces.
Diagnostic stages overlap and cannot be summed; no speedup or repeated-layout
bottleneck is claimed. No library rebuild or suite replay. Evidence:
`artifacts/composition-c1/checkpoint-124-evidence.json`. Matrix row/baseline residuals
remain the next bounded geometry investigation.

C123 consolidates actual child metric ownership across admitted math containers.
An iterative source-order walk skips hidden operator roles. Metric-only root
classification reuses discovered runs; child metrics share the conversion font cache.
All-metric preparation requires every control to be metric-only, preserving painted
parent glyphs. Roots, operators, script bases and matrix cells forward actual metric
owners. Valid occupied structured cells no longer require glyphs. No grammar
widening, invented sources/subsets, quota knobs or native code.

Ten sources/twenty independent Word outcomes restore fourteen omissions and preserve
six primitive-fraction controls byte-for-byte. Fourteen MAEs improve without new
failure components. Sixteen one-point gates pass; four matrix cases retain1.762-
1.794pt Y residuals. All twenty strict raster gates remain failed. Ten crop pairs
reviewed; identical Final/Review rasters carry the other ten reviews.

The initial eighteen-outcome PDF screen restores ten omissions and leaves the two
matrix-script outcomes. Their missing forwarding/glyph-required guard is corrected
before qualification; an empty-fraction cell adds only two independent outcomes.
Original manifest SHA, failed prototype outputs and separate parent reports stay.
One final Release artifact passes195 Windows and120 Ubuntu checks/75 native skips,
once each. No tests repeat. Parent checks fail expected16 versus14 glyphs.

Seventeen assets/four smoke PDFs transfer without root build/suite replay, including
two original-portfolio identities. Roote8f9e3c2; evidence
`artifacts/composition-c1/child-metric-evidence.json`. C122 CI39035669276 is green
on Windows/Ubuntu. The next milestone reuses this artifact for one finite fixed-
portfolio refresh and the existing compound allocation probe. Matrix occupied-frame
residuals remain a separate source-backed geometry task, without fitted offsets.

C122 makes delimiter discovery, function discovery and argument measurement share
existing bounded fraction/script/matrix box kinds. Supported subscript, superscript,
combined-script and fraction-based script boxes now survive delimiters. Unknown or
nested delimiters and deeper source-role fallbacks remain explicit.

Four sources/eight independent Word outcomes restore four to eight omitted glyphs.
Six one-point gates pass; two superscript cases retain1.093pt Y residuals. Prose max
0.834pt X /0.722pt Y. All eight MAEs improve without new failure components; all
strict raster gates stay failed. Four crop pairs reviewed, carrying four more by
identical Final/Review rasters. No geometry threshold is relaxed.

One frozen Release artifact passes53 exact Windows delimiter/script/function/encoding
checks and37 Ubuntu checks/16 native skips, once each. Parent checks fail expected18
versus14 glyphs. Seven exact portfolio sources contain no scripts inside delimiters,
carrying C121's twelve actual accepted comparisons by source inference. Seventeen
assets/four smoke PDFs transfer without root build/suite/corpus replay. Rootc2f90a4a;
evidence `artifacts/composition-c1/delimiter-script-evidence.json`.
C120/C121 CI39034795913 is green on Windows/Ubuntu. Next bounded review: supported
child boxes whose metric owner is missing from parent preparation, using independent
Word probes before consolidating control discovery/forwarding.

C120 retains admitted bar fractions whose script and literal arguments have only
metrics. Actual controls own fonts and nominal sizes; existing placeholder boxes
compose structural ink without invented source glyphs or subsets. Empty literal
runs cannot choose an unrelated font/size or create unused resources. Unsupported
empty stacks and deeper structures retain explicit fallbacks.

Four sources/eight Word outcomes restore missing bars/frames although both old and
reference text totals are14 glyphs. All math/prose anchors pass: math max0.107pt X /
0.730pt Y, prose0.024pt /0.102pt. Four full rule gates pass; four display width gates
retain0.110pt against0.1pt. All eight MAEs improve, strict raster gates remain failed
without new components. Four crop pairs reviewed; the other four carry by identical
Final/Review rasters. Empty-run variants paint identically and publish no math subset.

One frozen Release artifact passes118 Windows checks and74 Ubuntu checks/44 native
skips, once each. Parent checks fail two rules versus one. Seventeen assets/four
integration PDFs transfer without root build/suite replay. Root9aa208ca; evidence
`artifacts/composition-c1/glyphless-fraction-evidence.json`.

C121 reuses the same artifact for twelve actual fixed-portfolio outcomes and twenty
inspected pages. All accepted candidate/reference PDF identities match. Twelve
comparisons and twenty previously reviewed page pairs carry without new Office
exports, rasters or metrics. Ten DOCX gates fail and both decks pass unchanged
thresholds. Ten prior content screens and two unresolved complete-report dynamic-
field/math-tokenization flags carry by identity; no fresh audit pass is claimed.
No tests/library rebuild or allocation rerun; descriptive cold timings imply no
speedup. Evidence `artifacts/composition-c1/checkpoint-121-evidence.json`.
C119 CI39033900340 is green on Windows/Ubuntu. Next bounded omission probe: supported
script boxes inside delimiters, aligning discovery with existing argument measurement.

C119 retains an entirely empty script argument beside a visible fraction sibling.
Shared script-argument recognition publishes the actual fraction/child owners in
order within the admitted bounded grammar. Existing metric-only preparation then
resolves the child's font without recursive discovery, dummy glyphs or extra subsets.

Four sources/eight Word outcomes restore one or three omitted sibling glyphs each.
All eight math/prose anchor gates pass (math max0.107pt X /0.911pt Y). Four full
rule/anchor gates pass; four raw width failures retain0.108-0.110pt residuals against
the unchanged0.1pt threshold. All eight MAEs improve without new components; all
strict raster gates stay failed. Four crop pairs reviewed, with the other four
carried by identical Final/Review rasters. Structural bars and visible siblings are
present; no arbitrary-document fidelity is claimed.

Qualification111 Windows passes and70 Ubuntu passes/41 native skips. The initial
build catches an omitted test-constructor preset before any immutable output/test
run. The isolated publication assertion then needs the existing mathematical-
Unicode NFKC normalization. Test-only correction carries109 Windows/69 Ubuntu
passes and reruns only two/one methods; two invocations each. Production/CLI/all
1,857 inputs remain exact. Failed build/reports retained; no geometry threshold or
skip-guard change. Both parent checks fail expected15 versus14 glyphs.

The exact C118 seven-source audit carries C111's twelve actual accepted portfolio
comparisons after current source SHA checks. Seventeen assets/four integration
PDFs transfer without root build/suite replay. Rootd16f1aca; evidence
`artifacts/composition-c1/nested-metric-evidence.json`. C118 CI39032927411 is green
on Windows/Ubuntu. Workspace12:42Z10.629GB/522,968 files, no errors; loose5GB target
remains unmet and rejected retirements stay untouched. Next: fully glyphless
fraction arguments, using independent structural-rule/frame evidence.

C118 composes a supported literal script in a fraction numerator or denominator,
inline or display. Shared discovery and nominal math styles publish text/script/
script-script variants before measurement. Existing boxes consume those roles;
denominators carry cramped style. Uniform source nominal sizes remain required,
while child glyph sizes follow semantic roles. Deeper nesting and unverified whole-
deletion retention keep explicit whole-box fallbacks; no font/size/document guard.

Four sources/eight independent Word outcomes restore four or six omitted glyphs
each. All eight MAEs improve without new failure components. Four full display
rule/anchor gates pass (math max0.107pt X /0.719pt Y). Four inline gates remain
failed: up to1.057pt baseline or0.108pt rule width against unchanged thresholds.
All eight strict raster gates fail. Four crop pairs reviewed; all Final/Review
rasters are identical per source/reference, carrying the other four visual reviews.
Reviewed content is visible without demonstrated clipping/overlap; residual geometry
is explicit and does not establish arbitrary fidelity.

One production Release artifact qualifies107 Windows and68 Ubuntu passes/39 skips.
Eight PDF/Office outcomes precede the affected batch, establishing its actual scope.
Two new assertions initially assume a filename implies effective source-row style
and use a synthetic font missing script-script scale. A test-only correction uses
the shared row policy and a complete scale while retaining missing-scale fallback.
Carry105 Windows/66 Ubuntu passes and rerun only those two methods per platform;
production/CLI/all1,851 inputs stay exact. Two invocations each; failed reports stay.
Both parent checks fail expected18 versus14 glyphs. Cancellation, cumulative scene
work and deeper fallback checks pass.

Seven exact portfolio sources contain no script inside fraction arguments, carrying
C111's twelve actual accepted comparisons by source inference. Seventeen assets/
four smoke PDFs transfer without root build/suite replay. Root0f4360d6; evidence
`artifacts/composition-c1/fraction-script-evidence.json`. C115 CI39031466120 and
C116/C117 CI39032000763 are green on Windows/Ubuntu. Next bounded probe: nested
metric-only script ownership beside a visible fraction sibling.

C116 reviews six retained revision-band Word edge gates using actual frozen mapped
frames, which remain exact C104 identities. Occupied-frame, text-slot and scaled-
padding alternatives fail existing gates and trade errors across stories. No
production replacement is accepted; the3.660pt maximum residual remains. Evidence
`artifacts/composition-c1/revision-band-frame-contract-review.json` distinguishes
the legacy quarter-line/padding concern from a demonstrated new coordinate defect.
No exports, conversions, builds or tests repeat for this analytical review.

C117 measures the composed report and revision-composition report in both presets
against the C115 library. Four isolated cold processes, one warmup and three warm
samples per outcome retain all accepted PDF hashes; staged outputs also match.
Calling-thread allocation is558.0-563.6MiB cold and35.9-40.5MiB warm. Standalone
initial font-resource preparation is24.5-24.6MiB with466 discovered faces. These
stages overlap and are not additive; repeated pagination is not established as
the dominant cost. Timing is descriptive, with no speedup claim. The probe alone
builds/self-tests; library, rendering suites and Office exports do not repeat.
Evidence `artifacts/composition-c1/checkpoint-117-evidence.json`. Preserve bounded
cache ownership and investigate specific font costs before new cache/framework work.

C114 CI39031072723 is green on Windows and Ubuntu. C1154938873f is pushed; status
check pending. Next bounded content probe: scripts in fraction arguments, using
shared math-style contracts and explicit deeper fallbacks.

C115 verifies the C114 mixed-control contract in supported cells and direct
headers/footers. The visible script and plus sign coexist with a revised empty
fraction in both orders and both public DOCX presets. No production change.

Six sources/twelve new strict-view Word outcomes pass the existing rule/anchor
gates: math max0.522pt X /0.591pt Y, prose0.610pt /0.174pt. Six crop pairs reviewed;
all twelve strict raster gates remain failed. Actual structural bars, revision
neutrality, content, image presence and bookmark/run-split invariance are checked.

Three new Windows methods and two Ubuntu methods/one native skip run once each.
C114's84 Windows passes and58 Ubuntu passes/26 skips carry by production tree,
non-test asset, all1,837 old input and previous test-source identities. A test-only
build reuses the qualified library/CLI; no old suite or corpus conversions repeat.
C111's twelve actual accepted portfolio comparisons carry unchanged production.
Seventeen assets/four smoke PDFs transfer without root build. Root5589a2ed; evidence
`artifacts/composition-c1/mixed-control-placement-evidence.json`. C114 CI39031072723
has Ubuntu green/Windows running; C113 CI39030203632 is green on both platforms.
Next: investigate retained revision-band extents through actual physical frames,
without adding fitted container/font/size offsets.

C114 retains an empty fraction beside a visible script and plus sign. Shared
preparation resolves its actual metric owner through the existing conversion cache
without an unused subset. Revision ownership follows actual controls independently
of painted literals; ordinary neighbors cannot create review ink. The existing
parent-metric preference and deeper/unknown fallbacks remain intact.

Six sources/twelve independent Word outcomes expose ten omissions of four visible
glyphs and a bar. All ten affected rule/anchor gates pass: math max0.106pt X /0.423pt
Y, prose0.130pt /0.010pt. Ten MAEs improve; two visible Review PDFs retain exact
bytes. Their0.108pt width residual and differing strike segmentation remain failed
in the raw rule report. All twelve strict raster gates fail without new components.
Six crop pairs reviewed; bookmarks/run splitting preserve content and placement.

Qualification is84 Windows passes and58 Ubuntu passes/26 native skips on one
production Release artifact. Two new observers initially equate Word's three
overlapping run/operator strikes with four generated glyph strikes. The harness
now counts black structural bars and margin bars exactly and checks revision-strike
presence plus actual glyph author ink. Raw gates and numerical thresholds remain
unchanged. A test-only rebuild proves library/CLI/all1,837 input identities; carry
80 Windows/56 Ubuntu passes and rerun only four/two observers. Two invocations each;
failed reports retained. Both parent checks fail expected18 versus14 glyphs.

Seven exact portfolio sources have no revisions or empty fraction/script/matrix
controls, including inside painted equations. C111's twelve actual accepted PDF
comparisons carry by source inference. Seventeen assets/four integration PDFs
transfer without root build/suite replay. The inherited unused-header diagnostic
remains source-classified; no diagnostic policy change. Rootef9f052e; evidence
`artifacts/composition-c1/mixed-controls-evidence.json`. C113 CI39030203632 is green
on Windows and Ubuntu. Next: the same mixed controls in cells/direct static stories.

C113 publishes ordered actual math controls independently of literal run presence.
Shared metric-only preparation resolves each control font; occupied geometry drives
placement and edges. Supported flat Final deletion retains actual control frames
without recursive discovery, dummy glyphs or math subsets. Deeper/unknown and
matrix/function deletion fallbacks remain explicit.

Six sources/twelve Word outcomes expose ten missing bar/frame results. All ten
affected rule/anchor gates pass: prose max 0.089pt X / 0.010pt Y; neighboring math
0.106pt X / 0.351pt Y. Ten MAEs improve; two visible Review PDFs retain exact bytes.
Their inherited 0.108pt rule-width residual exceeds the preserved 0.1pt gate and
remains failed in the raw report. All twelve strict raster gates remain failed,
without new components. Six crop pairs reviewed; actual rule/strike ink, revision
neutrality and bookmark/run-split transformations are checked.

One production Release artifact qualifies 81 Windows methods and 56 Ubuntu passes/
25 skips. A new native assertion initially includes the unchanged Review controls.
A harness-only selection correction checks ten affected geometry outcomes and two
unchanged content/ink outcomes. Preserve the failed report and all raw gates; carry
80 passes and rerun only that native check. Production/CLI/all 1,829 inputs remain
identical. Two Windows invocations; Ubuntu once with unchanged branches/skip guard.
Two parent checks fail on actual expected two rules versus one.

C111's twelve actual accepted portfolio comparisons carry by seven-source audit:
no revisions/glyphless equations, and painted-source measurement stays unchanged.
Seventeen assets/four smoke PDFs transfer without root build/suite replay. Root
75a3b9be; evidence `artifacts/composition-c1/glyphless-compound-controls-evidence.json`.
C112 CI39028685192 is green on both platforms. Workspace at 11:41Z is 10.450GB /
512,756 files, no errors; loose 5GB target remains unmet. Rejected retirements stay
untouched and no reclamation is claimed.

C112 retains empty subscript/superscript/combined slots through actual implicit
controls and existing MATH placeholder measurement. Zero-run boxes require an
explicitly resolved metric owner; painted scripts keep their original first-source
font/size path. Metric-only preparation publishes no dummy glyph or math subset.
Unsupported script properties retain their existing fallback. Cancellation and
cumulative scene limits remain intact.

Five sources/ten Word outcomes cover authored blanks, literal deletion and empty
insertion. All one-point math/prose gates pass: max 0.107pt X / 0.766pt Y. Parent
advance errors reached 17.892pt and baseline errors 2.4pt. Nine MAEs improve, while
the visible deleted-script Review PDF remains byte-identical. All ten strict raster
gates remain failed without new components. Six crop pairs reviewed; bookmarks and
unchanged-format run splitting preserve content, ink and placement.

One frozen Release artifact passes 77 Windows methods and 54 Ubuntu checks/23
native skips, once each. Two parent checks fail on actual missing frame geometry/
main-part fallback, without new production APIs. Seven exact original inputs have
no empty scripts or revisions; C111's twelve actual accepted portfolio comparisons
carry by source audit, with painted-script measurements unchanged. Seventeen assets/
four smoke PDFs transfer without root build/suite replay. Root cc589b76; evidence
`artifacts/composition-c1/empty-script-frames-evidence.json`.
C110/C111 CI39027225044 is green on both platforms.

C111 refreshes the fixed seven-source portfolio using C110's already qualified
artifact: twelve actual conversions/twenty inspected pages retain accepted PDF
bytes, expected text and zero ignored equations. All twelve independent Office
comparisons reuse exact candidate/reference/threshold identities. Ten DOCX gates
remain failed; both decks pass. Ten prior prose/math screens and two unresolved
complete-report screens carry by source/PDF/reference identity; no new audit pass
is claimed. Twenty earlier reviewed page pairs carry by PDF identity. No extra
build, broad suite, Office export, rasterization or frozen artifact is needed.
Cold process times are recorded without a speedup claim. Evidence:
`artifacts/composition-c1/checkpoint-111-evidence.json`.

C110 restores six missing glyphless fraction bars through the existing control
ownership contract. Whole/fragment insertion and deletion retain actual owners,
metric-only preparation and neutral black structural ink, with no dummy glyph or
math subset. Eight independent Word rule/anchor gates pass; rule residuals max
0.021pt X / 0.015pt Y, width 0.068pt and thickness 0.006pt. Word and candidate have
zero margin bars. Six MAEs improve/two unaffected Final deletion PDFs remain exact.
Six equal-scale crop pairs reviewed; all eight strict raster gates remain failed.

One production artifact qualifies 53 Windows methods and 40 Ubuntu passes/13 skips.
A native observer initially compares local rectangles to page coordinates. Its
harness-only correction accounts for emitted transforms; actual production/CLI/
all 1,814 inputs remain identical. Carry 52 passes and rerun only that native check:
two Windows invocations, Ubuntu once with unchanged executed branches/skip guard.
The failed report remains retained. Two parent checks fail on actual expected two
fraction rules versus one. Seventeen assets/four smoke PDFs transfer without root
build/suite replay. Root 71124339; evidence
`artifacts/composition-c1/glyphless-fraction-revisions-evidence.json`.
C109 CI39026052784 is green on both platforms. Version remains 0.1.5; release
preparation stays deferred.

C109 fixes wrapper-only content loss in supported compound script/fraction rows.
Insertion syntax now reuses the ordinary literal/control grammar and actual source
owners. Both orders retain all eight formerly omitted glyphs in both presets.
Unknown, mixed-unknown and nested wrappers retain explicit fallback; unsupported
Final compound deletion preserves traversed source rather than dropping structure.

Four sources/eight strict Word outcomes pass all one-point anchors: math max
0.136pt X / 0.459pt Y, prose max 0.169pt X / 0.046pt Y. Both inserted Final PDFs
are byte-identical to unwrapped controls. Four MAEs improve; four ordinary PDFs
remain exact. All eight strict raster gates stay failed without new components.
Four crop pairs reviewed; bookmark/run splitting preserves content, ink and placement.

One frozen Release artifact passes 32 Windows methods and 23 Ubuntu checks/nine
native skips, one invocation each. Two parent checks fail on actual 22-versus-14
glyph counts. A test-only compile error is corrected before qualification; retained
validation-tool/path setup failures require no conversion or test replay. Seven
original sources contain no revisions, so twelve accepted portfolio comparisons
carry by source audit. Seventeen assets/four smoke PDFs transfer without root build
or suite replay. Root c8238f8f; evidence
`artifacts/composition-c1/compound-revision-rows-evidence.json`.
C108 CI39025340929 is green on both platforms. Workspace at 10:39Z is 10.360GB /
505,752 files, no errors; the loose 5GB target remains unmet. Rejected retirements
remain untouched; no reclamation is claimed.

C108 qualifies literal matrix-cell deletions across body, table cell, direct header
and footer paragraphs through unchanged C107 production. Four actual literal owners
retain an unrevised metric control. Final keeps blank slots; Review keeps source
ink/strikes. Bookmarks and unchanged-format run splitting preserve placement.

Four sources/eight strict Word outcomes pass all one-point anchor gates: math max
0.501pt X / 0.401pt Y, prose max 0.567pt X / 0.398pt Y. Eight crop pairs reviewed;
all eight strict raster gates remain failed. Three new Windows checks and two
Ubuntu checks/one native skip run once each. Exact production/CLI/1,796 existing
input identities carry C107's 65 Windows and 46 Ubuntu/19 skips without replay.
Only tests rebuild. Twelve accepted portfolio comparisons carry through the same
identity/source audit; seventeen assets/four smoke PDFs transfer without root build.

Separate whole-control probes do not justify a shared deletion-frame rule: empty
and font-bearing control properties differ in Word. Their cause remains unverified,
raw observations retained, and no narrowly qualified runtime branch is added.
A reusable summary command records existing inspected math/prose gates, including
failures, without conversion or PDF parsing. Root af06b6cc; evidence
`artifacts/composition-c1/matrix-cell-revisions-evidence.json`.
C107 CI39024753073 is green on both platforms.

C107 retains authored-empty matrix slots through existing hidden-placeholder
measurement and shared metric-only font preparation. Actual control/paragraph
metrics own advance and occupied height; no dummy run, placeholder glyph or math
PDF subset is published. Source-free measurement still requires a metric owner.
Shared standalone-control discovery retains actual empty-grid revision ownership.

Five sources/ten independent Word outcomes cover explicit 11/18pt controls,
implicit metrics and whole-deletion counterfactuals. Six affected authored-empty
gates pass: maximum prose residual 0.087pt X / 0.318pt Y, and neighboring math
0.106pt X / 0.454pt Y. Parent errors reached 51.162pt advance and 15.120pt baseline.
All six affected MAEs improve; four deletion PDFs retain exact bytes. Six crop
pairs reviewed. All ten strict raster outcomes remain failed, with no new components.
The inherited 18pt Review deletion baseline misses by 2.424pt; its raw geometry gate
stays failed. Word distinguishes tested empty grids and whole deletions; no blanket
deletion-frame rule is added, and the implicit-deletion discrepancy remains separate.

One production Release artifact qualifies 65 Windows methods and 46 Ubuntu passes/
19 native skips. The first native observer assumes an unsplit text run. Its correction
exposes the inherited deletion residual. Both failures stay retained. Harness-only
corrections prove production/CLI/runtime/all 1,796 inputs identical; 64 Windows
passes carry and only the corrected native check reruns. Accepted qualification uses
two Windows invocations plus one rejected intermediate observer invocation; Ubuntu
runs once, with executed branches/skip guard unchanged. Two parent checks fail on
actual missing frame geometry/source admission, without new production APIs.

Seven original sources contain no all-empty matrices; twelve accepted portfolio
comparisons carry by source audit. Seventeen assets/four smoke PDFs transfer with
no root build/suite replay. Root 205429a7; evidence
`artifacts/composition-c1/empty-matrices-evidence.json`.
C105 CI39023185874 and C106 CI39023621260 are green on both platforms.

C106 fixes confirmed script/function/matrix omissions inside insertion revisions.
Revision wrappers now reuse the existing single-control grammar instead of a
second, narrower allowlist. Actual literal owners/properties survive; role/orphan
rows and unknown/multiple/nested compositions keep explicit fallbacks. Richer
Final deletion frames remain bounded rather than silently losing structural paint.

Three mixed sources/six independent Word outcomes retain all expected math and
revision ink. Bookmark and unchanged-format run splitting preserve placement.
Both native and pinned parent checks fail on actual 17-versus-14 missing glyphs.
All six one-point group gates pass (0.106pt X / 0.620pt Y maximum); all six MAEs
improve. Six strict raster gates still fail, without new components. Six same-scale
crop pairs reviewed, including the image and neighboring formulas.

One frozen Release artifact passes 23 exact Windows checks and 17 Ubuntu checks
with six native skips, one invocation each. Seven original sources contain no
affected revisions; twelve accepted portfolio comparisons carry by source audit,
with two actual original portfolio transfer smokes. Seventeen assets/four smoke
PDFs transfer without root build or suite replay. A reusable command pins math-run
anchors from existing source/PDF/export-log/view identities, without exporting or
rebuilding. Root 88d6f14c; evidence
`artifacts/composition-c1/control-revisions-evidence.json`.

C105 fixes a confirmed unit mismatch in Review band painting. Mapped note frames
already carry physical heights, while their legacy line-height field retains
unscaled design units. Painting now consumes the physical frame, using the existing
fallback for lines without one. No layout, text or font/size-specific rule changes.

One frozen Release artifact passes 28 Windows checks and 23 Ubuntu checks with
five native skips, once per platform. The parent fails the independent one-point
note-bottom gate. Word bottom residual improves from 1.832pt to 0.320pt; the largest
remaining band-edge residual falls from 6.759pt to 3.660pt. All six full vertical
gates still fail. Both modes retain exact text operations and Final PDF bytes;
Review page two MAE improves slightly. Both strict raster gates remain failed,
with no new components. One note crop pair reviewed; no fitted baseline offsets.

Seven original portfolio sources contain no affected revisions. Accepted twelve
PDF comparisons carry by source inference, with two actual portfolio transfer
smokes. Seventeen assets/four smoke PDFs transfer without root build or suite
replay. Root 58656cd7; evidence
`artifacts/composition-c1/mapped-revision-band-height-evidence.json`.
C104 CI39022239555 is green on both platforms.

C104 checks a complete report with five whole-equation/fragment revisions across
body, static cells and an endnote. Shared mapped-line traversal retains placement
container ownership; Review now has six separate story bars across two pages,
matching Word counts rather than joining each page through unchanged content.
Body adjacency uses its own source flow; existing table/column/floating fallbacks
remain. Both modes preserve byte-identical text operations and Final PDF bytes.
The two Review page MAEs improve; strict raster outcomes remain failed. Horizontal
bar gates pass; all six vertical edge gates fail, up to 6.759pt. No vertical pass
or full-document Office-fidelity claim. Four whole-page pairs reviewed.

The portable check exposed inherited, real note-content loss: a font-derived first
inset of -7.238pt caused ordered layout to reject otherwise prepared/owned math.
Shared finite/nonnegative hint admission now uses the existing measured fallback.
Valid Office-measured hints stay intact. Pinned math-face rendering restores four
Final/eight Review note glyphs (total 52/72), including wrapping and actual revision
ink. Two pinned before/after crop pairs reviewed.

The first built candidate passed native report checks but failed portable content:
Windows 94 pass/1 fail, Ubuntu 78 pass/1 fail/16 skips. After production correction
the expanded 108-method selection passes once on Windows; Ubuntu 92 pass/16 native
skips once. Initial failures and a preceding compile-error log remain retained;
no passing checks carry across changed libraries. All twelve existing portfolio
PDFs retain accepted bytes and independent comparisons reuse exact identities.
Two content screens still flag extraction/source ambiguities; raw failures remain
even though source-backed Word math/ink counts and reviewed pages retain content.

The new report is a development regression fixture; its initial held-out label is
corrected without changing binaries, actual DOCX or golden data. A metadata-only
freeze reuses qualification by asset identity, with no build/test/conversion replay.
Existing held-out originals remain unchanged. Seventeen assets/four smoke PDFs
transfer without root build or suite replay. Root b8ba9044; evidence
`artifacts/composition-c1/revision-story-bands-evidence.json`. C103 CI39020057269 is
green on both platforms. Workspace 09:42Z: 10.229GB/497,024 files, no errors; loose
5GB target remains unmet and rejected retirements stay untouched.

C103 closes another confirmed content omission: insertion/deletion around a single
root, delimiter, n-ary or bar-fraction node inside `oMath`. Revision syntax now
publishes the actual control rather than a synthetic row; literal/control owners
and properties survive. Final shares existing literal deletion-frame preservation.
Iterative traversal retains cancellation and explicit unknown/deeper/multiple/
nested-wrapper fallbacks, without a public option or container-specific rule.

Eight sources/sixteen independent Office outcomes improve MAE; all one-point group
gates pass (0.683pt X / 0.768pt Y maximum). Strict raster gates remain failed with
no new components. Five same-scale crop pairs reviewed, including deleted fraction
structure. Two parent checks retain the actual 18-versus-14 missing glyphs.

The first production prototype retained a synthetic row and failed content/source
ownership: Windows 213 pass/3 fail, Ubuntu 149 pass/2 fail/65 skips. After production
correction all 216 selected methods rerun. One observer then incorrectly assumed
unique paragraph revision entries; a harness-only correction asserts every actual
reference and reruns only that method per platform. Accepted production qualification
is 216 Windows passes, Ubuntu 151 passes/65 native skips, two invocations each with
unchanged checks carried; raw failures retained. Initial prototype adds one rejected
invocation per platform. No checks carry across changed production libraries.

Seven fixed portfolio sources contain no affected structural revision wrappers.
C100 accepted twelve PDF identities/comparisons carry by source inference with two
actual portfolio transfer smokes; the initial rejected artifact also converted
twelve unchanged PDFs, which are not final-artifact qualification. Seventeen
assets/four smoke PDFs transfer without root build or suite/corpus replay. Root 08ba96fb;
evidence `artifacts/composition-c1/structural-fragments-evidence.json`.
C101 CI39018292234 is green on both platforms; C102 pushed as dd380df8.

C102 qualifies whole root insertion/operator deletion in mixed numbered paragraphs
inside cells and direct headers/footers through unchanged C101 production code.
Six sources/twelve strict Office outcomes preserve math, image, prose and source
ink; bookmarks and unchanged-format run splits preserve placement. Eleven group
gates pass; the operator-cell Review case retains the known 1.163pt X failure.
Maximum Y residual is 0.384pt. All twelve strict raster gates remain failed; four
same-scale crop pairs reviewed. No container guard or threshold change.

A harness-only build adds two checks: Windows 2 pass once; Ubuntu 1 pass/1 native
skip once. Production source/library/CLI and existing inputs are identical; C101
73 Windows and 62 Ubuntu/11 skips, plus C100 twelve fixed portfolio identities and
comparisons, carry without replay. Seventeen assets/four smoke PDFs transfer
without root build or suite replay. Root 43c5b904; evidence
`artifacts/composition-c1/whole-math-placement-evidence.json`.
C100 CI39017791481 is green on both platforms; C101 pushed as f385403c.

C101 removes extra Review bars from structural-only equation revisions through
the shared non-rendering source policy. Actual parsed control owners remain valid;
visible literal revisions, unowned/value-equal metadata and unknown positions
retain their existing behavior. Revision-band painting reuses cached paragraph
classification and threads cancellation through source discovery.

Three new deletion sources/six strict Word references confirm the insertion
findings. Seventeen actual outcomes yield six MAE improvements/eleven exact PDFs;
all seventeen group gates pass and strict raster gates remain failed without new
components. Six neutral Review outcomes have zero Word/candidate bars; four crop
pairs reviewed. One Release artifact passes 73 Windows checks and 62 Ubuntu checks/
11 native skips, one invocation per platform. Two parent checks retain the actual
extra bar. Fixed portfolio C100 results carry by source inference rather than a
corpus replay; two actual portfolio transfer smokes preserve complete Review/deck
bytes. Seventeen assets/four smoke PDFs transfer without root rebuild or suite
replay. Root b2676a1b; evidence `artifacts/composition-c1/structural-neutrality-evidence.json`.

Workspace scan at 08:36Z remains above the loose cap: 10.071GB/487,263 files, no
measurement errors. Rejected retirements remain untouched; no reclamation claimed.

C100 fixes confirmed whole-equation loss inside insertion/deletion wrappers in
`oMathPara`. One actual source owner survives iterative traversal. Known literal
root, delimiter and n-ary deletion frames retain structural paint in Final; their
Review symbols keep authored ink while literal text receives author color/strike.
Empty structural controls retain actual ownership without invented source runs.
Unknown/deeper/nested compositions keep explicit fallbacks.

Eleven mixed sources yield twenty-two independent Office outcomes: all MAEs improve,
all one-point group gates pass (0.683pt X / 0.584pt Y maximum), and all strict raster
gates remain failed without new components. Six crop pairs reviewed. Four deleted
roles retain one strike per actual revised literal glyph, with structural symbols
unstruck. Three empty structural insertions still draw an extra Review margin bar
absent in Word; this confirmed residual is the next bounded semantic task. The
unused related-header warning remains distinct from printed content.

One production Release artifact qualifies 142 Windows methods and 89 Ubuntu passes/
53 native skips. Three synthetic equality assertions initially included distinct
source-owner identity in geometry; a harness-only correction compares painted
fields and asserts actual ownership separately. Only four changed checks rerun;
138 Windows and 85 Ubuntu passes carry by unchanged production/helpers/inputs.
Both platforms use two invocations; raw failures retained. Two parent regressions
fail with actual 14-versus-18 missing math glyphs. All twelve fixed portfolio PDFs
retain accepted bytes; independent comparisons reuse exact identities. Seventeen
assets/four smoke PDFs transfer without root rebuild or suite replay. Root 7f7f2b8d;
evidence: `artifacts/composition-c1/whole-structural-revisions-evidence.json`.
Previous C98 CI39015102889 is green on both platforms.

C99 retains a bounded vertical-root investigation with no production correction.
Twelve independent inline/display controls vary structural/body sizes and descender
ink. Body baseline residual reaches 3.658pt; a uniform size offset is not supported
by the countercontrols. In the inspected descender case, embedded Word root/body
and degree glyph bounds match the installed Cambria Math outlines and units.
Font mismatch is not established there; the frame/clearance cause remains unverified.

The [OpenType MATH specification](https://learn.microsoft.com/en-us/typography/opentype/spec/math)
provides placement parameters rather than a complete layout algorithm. Keep
independent Office output and raw numerical failures; no fitted font/size branch.
References/parent output: `artifacts/composition-c1/root-frame-office` and
`root-frame-parent`; findings: `root-frame-evidence.json`. Continue with whole-equation
revision composition, where the reader's fraction-only filtering may omit visible
structural content. No new qualification or geometry pass is claimed here.

C98 clamps the complete radical origin after signed degree overlap. A narrow
degree can begin to its right while the complete group stays nonnegative; wider
degrees retain leading space. Root, radicand and bar retain shared translation.
Six unrevised public controls independently confirm the Word relationship; no new
DOCX fixtures or font/size/context-specific rule.

Across those six controls and sixteen revision outcomes, four MAEs improve and
eighteen PDFs retain exact bytes. All twenty-two horizontal group-start gates pass,
within 0.264pt. Ten vertical gates still fail at up to 1.96pt, and all twenty-two
strict raster gates remain failed with no new failure components. Four crop pairs
reviewed; the vertical body/prose discrepancy remains separate and visible.

One production Release artifact passes a 33-method root-specific selection. Parent
Word reproduces the 1.325pt defect; the initial synthetic strict-inequality observer
was too weak against floating-point cancellation noise. A harness-only margin fix
runs only that check on each platform through unchanged production/CLI/inputs:
Windows 33 pass and Ubuntu 22 pass/11 native skips, two invocations per platform with
unchanged methods carried. Both corrected parent regressions now fail on actual
placement; original weak pass retained without relabeling.

Seven fixed portfolio inputs contain no radicals; broader checkpoint carries by
source inference with two actual portfolio smoke PDF identities. Seventeen assets/
four smoke PDFs transfer without root rebuild or corpus/suite replay. Root 15d7d915;
evidence: `artifacts/composition-c1/root-degree-kerns-evidence.json`.
C96/C97 CI39013761358 is green on both platforms.

C97 retains root/body content beside empty visible degrees, using the existing
script-level argument sizing and metric-only placeholders. Eight inline/display
controls include independently sized structural frames. Four deleted/authored-empty
Final pairs have identical Word rasters. Unknown/deeper, stored nonempty hidden
degrees and missing-metric/encoding fallbacks remain.

Sixteen Word outcomes restore content/ink: twelve MAEs improve; four nonempty Review
PDFs stay parent-identical. Five crop pairs reviewed. Eight group-start outcomes
meet one point; eight larger-control outcomes still fail, reaching 1.54pt X/1.96pt Y.
The unchanged Review PDFs expose the same inherited error. Its cause remains
unverified. All sixteen strict raster gates also remain failed; no thresholds relaxed.
Content/ink observers explicitly do not claim those geometry gates passed.

One production Release artifact supplies qualification. The initial native observer
correctly exposed that residual: 30 pass/1 fail retained. A content-only observer
replaces its qualification role; a harness-only build preserves production/CLI/input
identity, carries thirty unchanged methods and runs only that observer. Final
selection has 31 passing methods through two invocations; Ubuntu 21 pass/10 native skips
once. Two parent checks fail on actual missing root/body glyphs.

Seven pinned portfolio inputs contain zero radicals, so the C95 broad checkpoint
carries by source inference rather than a new twelve-PDF identity claim. Four smoke
PDFs include actual complete-report Review/deck identities. Seventeen assets transfer
without root rebuild/corpus/suite replay. The existing generator retains its earlier
eight source fixtures/manifest byte for byte. Root 2e6e641f; evidence:
`artifacts/composition-c1/empty-root-degrees-evidence.json`.
C95 CI39012795079 is green on Windows and Ubuntu.

C96 checks C95 structural-only roots, delimiters and summations beside prose,
pictures, links, numbering and other formulas across body/cells/direct headers and
footers. Twelve sources/twenty-four strict Word outcomes preserve exact content,
source ink and bookmark/applicable run-split invariants. Four crop pairs reviewed.
Twenty-three outcomes meet the one-point group-start gate. Operator-cell Review
retains C93's 1.163pt X residual; maximum Y residual is 0.390pt. All twenty-four
strict raster gates remain failed. No threshold relaxation or container correction.

Two new Windows checks pass once; Ubuntu one pass/one native skip once. Production
source/library/CLI remain C95-identical: its 121 checks and twelve portfolio PDF
identities/comparisons/audit flags carry without replay. Harness-only build;
seventeen assets/four smoke PDFs transfer without root rebuild. The existing
placement generator now handles both families; the earlier eight fixtures and
manifest retain exact bytes. Root a2133569; evidence:
`artifacts/composition-c1/structural-placement-evidence.json`.
C95 CI39012795079 is green on Windows and Ubuntu.

C95 separates structural mathematical paint from printable literal source runs.
Supported roots, delimiters and n-ary operators keep their actual control owners
when all children are empty. One existing control-font preparation path serves
both cases; shared control font/size resolution preserves authored precedence.
Empty function names use shared metric-only measurement and retain actual argument
and nested-control owners. No dummy source runs or arbitrary recursive expansion.
Unpainted empty function/grid forms, unknown/deeper structures and missing metrics
or encoding retain explicit fallbacks.

Eight sources/sixteen strict Word outcomes restore remaining arguments and symbols.
Four deleted/authored-empty Final pairs have identical Word rasters. Twelve MAEs
improve; four PDFs stay exact. Group-start residuals stay within 0.155pt X/0.149pt Y.
Five crop pairs reviewed; all sixteen strict gates remain failed unchanged, with no
new failure components or relaxed thresholds. Source-specific review ink remains.

One frozen Release artifact passes 121 exact Windows checks once; Ubuntu passes
76 and skips 45 native checks once. New checks cover actual source/control identity,
unknown children, cancellation and cumulative work limits with page enumeration.
Two parent checks fail on actual omitted glyphs. Twelve fresh portfolio PDFs retain
accepted bytes, comparisons and content-audit flags. Seventeen assets/four smoke
PDFs transfer without root rebuild/suite replay. Root 81110ef7; evidence:
`artifacts/composition-c1/structural-math-paint-evidence.json`.
C94 CI39012308087 is green on both platforms.

C94 restores function names and n-ary operators/remaining limits when a literal
argument/body is empty. Three Word roles show identical deleted/authored-empty
Final rasters. Shared box-argument measurement reuses metric-only literals and
forwards existing neighbor/operator spacing; no font/size/container-specific guard.
Unknown, deeper, missing-metric/encoding and control-only fallbacks remain.

Six sources/twelve strict Word outcomes retain exact content/ink and bookmark/run
split invariants. Nine MAEs improve; three PDFs stay exact. Group-start residuals
stay within 0.111pt X/0.235pt Y. Four crop pairs reviewed; all twelve strict gates
remain failed unchanged with no new failure components or relaxed thresholds.

One frozen Release artifact passes 117 exact Windows checks once. Ubuntu passes
73 and explicitly skips 44 native checks once. Selection includes font and native
neighbors for shared function/delimiter/radical/n-ary/matrix argument behavior.
Two parent checks fail on actual missing function glyphs. Twelve fresh portfolio
PDFs retain accepted bytes; comparisons and unresolved content flags carry by
identity. Seventeen assets/four smoke PDFs transfer without root rebuild or suite
replay. Root a49f1333; evidence:
`artifacts/composition-c1/empty-math-bodies-evidence.json`.
C93 CI39011779448 is green on both platforms.

C93 checks the unchanged C92 binary with empty script bases and operator limits
beside prose, images, links, numbering and other formulas across body/cells/direct
headers/footers. Sixteen independent Word outcomes preserve exact content and ink;
bookmarks and unchanged-format run splits preserve geometry. Four same-scale crop
pairs show the remaining content without demonstrated omission/overlap/clipping.

Fourteen outcomes meet the existing one-point group-start gate. Both operator-cell
modes still fail it at up to 1.163pt horizontally; body/operator and cell-origin
residuals accumulate. Maximum Y residual is 0.626pt. All sixteen strict raster gates
also remain failed. No threshold relaxation or container-specific correction.
Content/ink regression checks explicitly do not claim those geometry gates passed.

Two new checks pass once on Windows; Ubuntu one pass/one native skip once.
Production source/library/CLI remain C92-identical, so its 77 checks and twelve
portfolio PDFs/comparisons/audit flags carry without replay. Harness-only build;
seventeen assets/four smoke PDFs transfer without root rebuild. Root e5e9dc69;
evidence: `artifacts/composition-c1/remainder-placement-evidence.json`.

C92 preserves remaining script glyphs and n-ary operator/body content when a
literal base or visible limit is empty, including Final deletion. Existing shared
argument measurement supplies metric-only geometry; an empty base is one atom,
classified by its font glyph. No font/size/container-specific correction added.
Five roles have independently identical deleted/authored-empty Word rasters.

Thirteen sources/twenty-six strict Word outcomes restore content: fifteen MAEs
improve and eleven PDFs stay exact. Group-start X/Y residuals stay within
0.122/0.340pt. Six same-scale pairs reviewed; all twenty-six strict gates remain
failed with no new failure components or changed thresholds. Native and pinned
checks preserve source ink, bookmarks and unchanged-format mathematical run splits.
Unknown structures, mixed metrics and wholly empty/control-only fallbacks remain.

One final frozen Release artifact passes 77 Windows checks once; Ubuntu passes
48 and skips 29 native checks once. Two parent checks fail on actual missing glyphs.
The first prototype's 1.267pt subscript failure is retained; after its semantic
atom correction all checks run on the new DLL. Twelve fresh fixed-portfolio PDFs
retain accepted bytes, reusing comparisons and the unresolved complete-report
content flags. Seventeen assets/four smoke PDFs transfer without root rebuild or
suite replay. Root daa1476f; evidence:
`artifacts/composition-c1/math-remainders-evidence.json`.
C92 CI39011336511 is green on Windows and Ubuntu.

C91 qualifies empty script/matrix roles beside prose, links, numbering and pictures
across body, simple cells and direct headers/footers. Sixteen strict Word outcomes
retain every expected mathematical glyph and source ink in both presets. Group-start
X/Y residuals stay within 0.521/0.620pt. Bookmarks and prose/math-run splitting retain
geometry; four same-scale crop pairs reviewed. All sixteen strict visual gates
remain failed unchanged, with typography/review differences explicit.

The C90 production binary handles these placements unchanged. Two new methods pass
once on Windows; Ubuntu passes the portable method and explicitly skips native
geometry once. No production build, speculative container correction or repeated
corpus/suite. C90's 36 checks and twelve portfolio outcomes carry by unchanged
production/CLI and relevant test/helper identities.

C90 CI39009124734 failed one older font-group matrix test on both platforms; raw
logs remain retained. Its visibility-exclusion expectation conflicts with the strict
Word counterprobes. The corrected test verifies identical known-profile slots and
retains unknown/wholly-empty/missing-metric fallbacks, cancellation and work limits.
It passes once on each platform. MathFontMatrix joins future affected selections.
Only this corrected check is rerun through the unchanged library; the placement
checks carry through unchanged source/helpers. Seventeen assets/four smoke PDFs
transfer without root build/suite replay. Root bbccd55a; font fix c4d4c458; evidence:
`artifacts/composition-c1/empty-role-placement-evidence.json`.
C91 CI39010092649 repaired CI on Windows and Ubuntu; its earlier failure stays retained.

C90 shares literal-argument measurement across fractions and scripts, retaining
metric-only empty sub/sup slots. Matrix discovery uses the same semantic emptiness
rule. Independent Word counterprobes show deleted/authored-empty script and matrix
arguments have identical rasters; known matrix placeholder visibility profiles also
have identical rasters in both presets. Their previous exclusion is removed. Empty
script bases, wholly empty grids and unknown properties retain explicit fallback.
Fraction/script operator spacing retains each role's previously qualified behavior.

Fourteen sources/twenty-eight strict Word outcomes generalize revisions across
literal, script, root and matrix roles. Eleven affected MAEs improve; seventeen
PDFs retain bytes. All twenty-eight gates still fail unchanged thresholds without
new components. Maximum group-start X/Y residuals are 0.455/0.700pt. Exact content
and ink cover every glyph; grouped Word strings use their actual starts, with no
estimated interior anchors. Six crop pairs reviewed (four changed, two reused by
PDF identity). Typography/strict visual failures remain. No raw prose-screen pass
is claimed for manifests without expected strings; focused tests assert the prose.

One frozen artifact passes 36 focused Windows checks once and eleven Ubuntu checks/
25 explicit skips once. Two actual parent failures show missing script remainders;
frozen CLI outputs independently establish the matrix omission. A selector-only
correction before tests keeps the same production/test source and artifact; no
extra build/freeze. All twelve actual portfolio PDFs retain accepted comparisons,
including ten failed DOCX gates/two deck passes. Seventeen assets/four smoke PDFs
transfer without root build or suite replay. Root d78c970c; evidence:
`artifacts/composition-c1/empty-math-arguments-evidence.json`.
C89 CI 39007590057 is green on both platforms; C90 CI follows the scoped push.

C89 composes tracked literal runs inside supported mathematics. Recognized
nonempty insertion/deletion wrappers use the existing iterative parser and declared
view filtering. Eligibility and author order require actual measured run identities
and source positions; a discovery hint alone cannot qualify metadata. Glyphs retain
their source revision, so review color/strikes affect only changed numerator text.
The denominator and structural rule retain their own ink. Shared placeholders now
also measure a partial empty bar-fraction argument after Final deletion. Empty or
structurally different inner revision wrappers and empty stacks retain fallback.

Sixteen strict Word outcomes cover all four containers and both presets. All sixteen
MAEs improve and all sixteen raw gates still fail unchanged thresholds, without new
components. Inserted/Review mathematics retains eighteen glyphs; Final deletion
retains its denominator and fourteen unchanged glyphs. Maximum X/Y residuals are
0.591/0.768pt. Five same-scale crop pairs reviewed; typography/review annotations
remain approximate. Bookmarks and splitting prose or the revised mathematical run
retain geometry/ink. The older whole-equation Final deletion now also uses the
one-point geometry assertion without its previous retained-scope exclusion.

One artifact passes all 52 Windows checks once and 26 Ubuntu checks/26 explicit
skips once. Two parent failures demonstrate actual missing glyphs. All twelve
freshly converted portfolio PDFs retain accepted bytes/comparisons; ten DOCX gates
fail/two decks pass. Seventeen assets/four smoke PDFs transfer without root build
or suite replay. Root fab2c593; evidence:
`artifacts/composition-c1/math-run-revisions-evidence.json`.
C88 CI 39006770470 is green. The 05:45Z footprint is 9.700GB/460,567 files,
zero errors; the loose cap remains unmet. No blocked retirement was retried.

C88 preserves empty bar-fraction structure after deleting its mathematical text
in Final. Independent Word rasters at 9/11/18pt equal authored empty fractions;
removing the equation equals an empty row and differs. Word keeps the structural
bar as well as layout, correcting the earlier hypothesis of hidden metrics only.
Matrix hidden cells and empty fraction arguments share the existing dotted-square
metric box. Normal/empty fractions share composition and paragraph preparation.
Metric-only font ownership adds no dummy glyph or PDF subset. Unsupported empty
stacks, richer deletion structures and unknown properties retain fallback.

Eleven independent Final frames supplement the neighboring sixteen equation
outcomes. All nine affected MAEs improve; fourteen neighboring PDFs retain bytes.
All 23 raw gates still fail unchanged thresholds without new components. Maximum
first/following baseline residuals are 0.174/0.810pt; structural rule X/Y/width/height
residuals stay within 0.501/0.153/0.090/0.031pt. Five same-scale crop pairs reviewed.
The two new 18pt outputs retain raw full-phrase extraction flags from wrapped math
paint order; separate prose/all-fourteen-math-glyph checks pass, without clearing
those flags. Typography and visual failures remain explicit.

Windows qualifies 47 focused methods (45 unchanged checks carried through identical
runtime/input identities and two corrected prose observers). Ubuntu passes 22 with
25 explicit skips in one invocation. Two parent failures establish actual baseline
and missing-rule defects. Initial compile/observer failures remain retained. All
twelve freshly converted portfolio PDFs retain accepted bytes/comparisons; ten
DOCX gates fail and both decks pass. Seventeen assets/four smoke PDFs transfer
without root build/suite replay. Root 1b5b99f7; evidence:
`artifacts/composition-c1/empty-fractions-evidence.json`.
C88 CI 39006770470 is green on both platforms.

C87 renders supported source-owned inserted/deleted equations through the shared
inline path. Actual wrapper/run revision identities and source positions establish
ownership; aggregate/value-equal metadata, wrong positions and unsupported kinds
retain fallback. Equation authors share source order with prose/pictures, including
an equation tied to its following prose index. Math glyphs reuse the existing author
palette and font strike metrics. Word uses no insertion underline here; fraction
rules retain their authored ink. No font/size/document-shape rule was added.

Sixteen strict Word outcomes cover insertion/deletion across body/cell/header/footer
and both public presets. The twelve affected MAEs improve; four Final deletion PDFs
retain parent bytes. All sixteen gates still fail unchanged thresholds without new
components. Four omitted fraction glyphs return alongside the other equations,
prose and picture. The affected maximum X/Y residuals are 0.551/0.768 points. Five
same-scale crop pairs reviewed. The unchanged Final deletion body/cell/header flow
still differs by up to 6.766 points; this is an explicit next defect, outside the
affected geometry assertion. Its raw failures remain; no tolerance was widened.

One production artifact qualifies 28 Windows methods (26 unchanged checks carried
through identical DLL/CLI/input identities, two corrected grayscale observers) and
nineteen Ubuntu checks/nine skips in one invocation. Raw observer failures remain.
Two parent regressions fail actual ownership/content checks. Neutral bookmarks/run
splits preserve all math geometry. Seven-source portfolio comparisons carry by
absence of revised equation wrappers, not by a corpus replay claim. Seventeen
assets/four smoke PDFs transfer without root build or suite replay. Root 142e320f;
evidence: `artifacts/composition-c1/equation-revisions-evidence.json`.
C87 CI 39005833382 is green on both platforms.

C86 shares actual non-rendering revision eligibility with Word change-bar
classification, memoized within each painting conversion. Ordinary unscaled Final
now uses the existing minimum-leading correction. Existing font/size/shape guards,
related-story placement and classic markup fallbacks remain. Twelve independent
Word outcomes establish identical rasters across control/empty/formatted-empty
variants and both presets within each source size. Ten MAEs improve; two Review
control PDFs retain bytes. All twelve raw gates still fail unchanged thresholds.
Two neighboring empty-owner MAEs improve; three PDFs retain bytes and all five
neighbor gates still fail. No new failure components. Two same-scale crop pairs
confirm placement and removal of the phantom bar; typography remains approximate.

One final artifact passes 57 Windows checks and 48 Ubuntu checks/nine explicit
skips, one invocation each. Two actual parent failures establish eligibility and
Final baseline defects. A placement-only prototype's two output-neutrality failures
remain retained; its results were not carried across the changed library. After
production was committed, adding empty insertions to all 21 held-out memo body/cell
paragraphs retains its accepted candidate PDF and both independent Word raster
pages byte-for-byte, including 32 math text operations. No fitting followed that
probe. Its existing four failed gate components and C79 visual review carry by
identity. All twelve freshly converted portfolio PDFs also retain accepted bytes;
ten DOCX gates fail and both decks pass. Seventeen assets/four smoke PDFs transfer
without root rebuild or suite replay. Root a716c42c; evidence:
`artifacts/composition-c1/revision-neutrality-evidence.json`.
C85 CI 39003038812 is green. The 05:01Z scan reports 9.598 GB/462,083 files,
zero errors; the loose cap remains unmet and no blocked retirement was retried.

C85 retains actual empty source-run revision ownership as small metadata records.
Their source indices/revision references participate in author ordering and shared
paragraph eligibility; they carry no visible text or font/layout purpose. Aggregate
or value-equal revision metadata still cannot establish ownership. Other source
fallbacks and cancellation remain intact.

This closes the retained empty-insertion color mismatch (all thirteen prior Word
owner-color facts now match) and a confirmed composition defect: adding an empty
insertion previously omitted a supported fraction in both Final and Review output.
The numerator/denominator now remain visible. Current Final PDFs with/without the
empty insertion are byte-identical; independent Word Final rasters are identical
too. Four new Review references and two Final references supplement the retained
empty-owner reference. Four of seven MAEs improve and three PDFs retain bytes;
all seven raw gates still fail unchanged thresholds, without new failure components.
Four same-scale pairs were reviewed. Typography/review annotations remain approximate.

Windows passes 31 focused checks; Ubuntu passes eighteen with thirteen explicit
skips, each in one invocation. An actual C84 parent regression fails its author
color check; parent CLI outputs independently establish the fraction omission in
both modes. The existing unreferenced-header equation diagnostic remains recorded.
Seventeen assets/four smoke PDFs transfer without root build or suite replay.
Seven-source corpus reports carry by verified absence of relevant source revisions;
this is source inference, not a full-corpus reconversion claim. Complete Final/deck
transfer smoke additionally retains bytes. Root aedf0585; evidence:
`artifacts/composition-c1/empty-revision-owners-evidence.json`.
C85 CI 39003038812 is green on both platforms.

C84 shares print-scaled image dimensions between body/cell wrapping, occupied
size estimates and placement. Relative source-baseline geometry keeps the first
print origin independent of scaled flow; leading picture-only blocks own their
origin. Rendering and continued-note keep-out reuse the same origin selection.
Retained picture-font descent contributes below-baseline space. Source-positioned
cell images use authored spacing; unpositioned model objects keep their existing
block fallback. No font-name, point-size or whole-document eligibility rule.

Twenty Review controls cover shared runs, picture-only paragraphs, three image
sizes, fonts, authors and all four containers. Six new strict Review references
and two Final references supplement fourteen retained references. All sixteen
affected Review MAEs improve, four accepted static PDFs retain bytes, and no new
failure components appear. All twenty Review gates still fail unchanged thresholds.
Final body retains bytes; the image-only Final cell's MAE improves 0.16618 to
0.01075 and recall 0.64423 to 0.98951. Both Final foreground-recall gates still fail.
Five Review and two Final same-scale pairs were reviewed for visible content.

Windows passes all 68 exact capability/neighbor checks; Ubuntu passes 55 with
thirteen explicit skips, each in one invocation. New independent geometry checks
reject both C83's size defect and the first dimension-only prototype's 3.294-point
text displacement. That prototype passed 53 Windows checks, demonstrating why its
failed visual comparisons mattered. Later failed origin/legacy-fallback runs also
remain retained; no unsuccessful candidate was transferred to root.

One final frozen artifact produces all accepted comparisons. The twelve fixed
portfolio PDFs retain accepted bytes and reuse reference/threshold comparisons;
ten DOCX gates fail and both decks pass. Earlier raw content-audit flags carry by
PDF identity. Seventeen assets/four smoke PDFs transfer without root build or
suite replay. Root 32e9a9be; evidence:
`artifacts/composition-c1/body-cell-picture-composition-evidence.json`.
C83 CI 39000959707 is green on both platforms; C84 CI follows the scoped push.

C83 admits actual source-owned inserted pictures to the existing semantic owner
check and shared natural static-frame mapping. Aggregate/value-equal metadata,
missing styles, wrong source positions, unsupported owners and declared/floating
frames retain their guarded paths. Wrapping, image rectangles, clips and text now
receive one physical header/footer mapping. No new font, size or document-shape
qualification was added.

Four independent Word header/footer controls restore visible picture/text content
and print scaling. The shared-run footer baseline moves from -35.242 to 129.716
points (Word 130.03). All four MAEs improve; their remaining gate failure is SSIM.
Ten neighboring picture PDFs retain bytes and comparisons. All fourteen raw gates
still fail unchanged thresholds. Four same-scale pairs were reviewed; body/cell
picture scaling and review annotations remain separate residuals.

Windows qualifies 40 exact methods: 39 retained through an identical DLL and one
corrected ownership observer (plain paragraphs create ordered items lazily). Ubuntu
passes 24 with sixteen explicit skips in one invocation. Two regressions fail on
C82 through actual ownership/geometry checks. Failed evidence stays retained.

One milestone conversion of the twelve fixed portfolio outcomes retains every
accepted PDF byte; all comparisons reuse candidate/reference/threshold identities.
Ten DOCX gates fail and two decks pass as before. Seventeen assets/four smoke PDFs
transfer without root build or suite replay. Root 2a0b3abc; evidence:
`artifacts/composition-c1/static-picture-frames-evidence.json`. The 03:53Z workspace
scan reports 9.441 GB/449,573 files, zero errors; the loose target remains unmet.
No blocked retirement was retried and no reclamation is claimed.

C82 retains inserted pictures' actual run styles and paints Word-compatible
review indicators from the existing text measurer, author registry and final image
geometry. No separate font cache or unused picture-only face is embedded. Shared
inline planning also retains preceding picture advance: adjacent source pictures
no longer occupy the same X in body, cell, header and footer painting.

Fourteen strict Word controls include ten new placement/style/shared-run references
and four retained references. Native relative indicator metrics, pinned portable
fonts, fallback ink, image ownership, ordering and following text pass. Windows
qualifies 42 methods (40 carried through an identical DLL plus two observer
corrections); Ubuntu passes 33 with nine explicit skips, one invocation. Both new
regressions fail on C81 through actual rendering. Raw failed runs remain retained.

All fourteen visual gates still fail unchanged thresholds: three MAEs improve,
nine worsen and two stay equal, with no new failure components. Every text-operation
set retains geometry and ink after resolving one harmless font-resource alias.
Ten single-picture PDFs retain bytes through the overlap repair. Same-scale review
confirms restored second-picture visibility in body/cell/header and the Arial14
indicator. It also establishes two inherited residuals: review picture frames are
36 points versus Word's 27.32, and footer content is below the viewport (baseline
-35.242 versus Word's 130.03). Passing emitted-content checks do not qualify that
footer as visible. C83 prioritizes its shared static-frame ownership repair.

Seventeen assets/four smoke PDFs transfer without root build or suite replay.
Complete-report Final and the held-out deck retain bytes; seven-source portfolio
comparisons carry by explicit source scope, not a new full-corpus identity claim.
Root b36146b2; evidence: `artifacts/composition-c1/picture-revision-evidence.json`.
The repaired C81 CI run 38996975092 is green on both platforms.

C81 shares the retained source-author registry between comment slots and final
Word-compatible insertion/deletion glyph and decoration painting. Classic colors,
unknown/move-owner fallbacks and derived numbering-change colors keep their
existing paths. No geometry, source-object or layout-eligibility changes.

Six independent Word glyph/underline/strike facts match, reusing the previous
references and adding one deletion-after-image control. All fourteen outcomes
retain glyph and path geometry; four MAEs improve, ten Review PDFs retain bytes,
and all fourteen visual gates remain failed without new failure components. Two
page pairs were reviewed. Classic AllMarkup additionally retains bytes. The
inserted picture remains visible but lacks Word's review indicator; balloon ink,
empty insertion ownership and incomplete-format fallbacks remain separate work.

Windows qualifies 40 unique methods: 34 unchanged results carried through an
identical production DLL, four affected ink/decoration checks rerun, and two
CI-selected content checks added. The first run's 36 passes/two failures remain
recorded: the fallback observer assumed literal strings instead of positioned byte
glyphs, and an older assertion required a fixed color across multiple authors.
The new independent deletion reference validates its source-ink correction.
Ubuntu qualifies 31 passes/nine skips, carrying the original 38-method invocation
and adding only the two content checks. Two new regressions fail on C80.

C80 CI (38995777285) caught two corrupted Greek assertions on both platforms. A
Windows-encoding test edit caused that damage; original Git UTF-8 bytes restore
exactly those two lines. The modified-file audit found no other such damage.
Both content checks pass against C80's unchanged library and the current artifact,
and on Ubuntu. Root 5319542d records the assertion repair; this was not a renderer
content omission. The next CI run will cover the repaired source and C81 batch.

Seventeen assets/four smoke PDFs transfer without root build or suite replay;
complete-report Final and held-out deck retain bytes. Corpus comparisons carry
by the unchanged C80 source-scope proof, not a new full-corpus conversion claim.
Root 6d4e1aec contains the rendering batch. Evidence:
`artifacts/composition-c1/revision-ink-evidence.json`. The latest completed workspace
scan is 9.312 GB/445,148 files with zero errors. The loose 5 GB target remains unmet;
no blocked retirement was retried and no reclamation is claimed.

C80 registers review comment authors from actual source owners, including revised
paragraph marks, run formatting and inserted images. Paragraph marks register
before inline owners; inline order follows recorded source indices. Every actual
revised mark is retained independently of numbering/comments. Aggregate paragraph
revision metadata cannot invent an author slot. Layout eligibility stays unchanged.

Thirteen fresh, strict Word controls distinguish absent/empty revisions, same
authors, ordinary text, marks, formatting and images, including their order beside
text revisions. Twelve now match actual comment fill/stroke paints; the empty
insertion remains a confirmed mismatch because parsing discards its owner. Seven
MAEs improve, six Review PDFs retain bytes, and all thirteen visual gates remain
failed with no new failure components. Three Final controls retain bytes. Every
source-body text operation retains ink and geometry. Three pairs were reviewed;
image/text multi-author revision ink still differs, and the incomplete-format
control keeps its generic balloon fallback. These residuals are not color passes.

The first artifact passes 21/24 Windows and 16/18 executed Ubuntu checks: it reveals
that ordinary revised marks were not retained. The corrected production artifact
passes all 31 specific mark/author/format/comment checks on Windows and 25 on Ubuntu
with six explicit skips, each once in one invocation. Two new regressions fail on
C78. No failed/previous-DLL qualification is carried into the corrected artifact.

Seventeen assets/four smoke PDFs transfer without root rebuild or suite replay.
The seven fixed/complete sources contain no changed active revision owners, so C79
comparisons carry by source inference rather than new corpus PDF identity claims;
the complete-report Final and held-out deck additionally retain transfer-smoke
bytes. C79 CI is green on both platforms (38994875743). Root 1996ef81 contains the
production batch. Evidence: `artifacts/composition-c1/nontext-authors-evidence.json`.

C79's font investigation found the measured kerning expansion already uses a
pre-sized bounded map. Further storage/cache changes exceed the bounded maintenance
slice; no optimization or speedup is claimed. Keep its baseline for later choices.

C79 reuses the qualified C78 library and test assembly for a composed-document
checkpoint. The stable supplemental `portfolio-checkpoint.json` includes the six
fixed inputs and the existing complete report; none of the seven source packages
changed. Twelve new CLI outputs are byte-identical to the accepted corpus. All
twenty Office/candidate page pairs were reviewed at matching scale: no visible
content loss, overlap or clipping was demonstrated. Typography, cumulative line
spacing, table/math geometry and review range/connector placement remain different.
Ten DOCX numerical gates still fail; both decks pass unchanged thresholds.

Strict reference preflight rejected the early cache's missing identity pins and
incomplete/different declared Word modes. Ten fresh Word exports satisfy all eight
readbacks; two fresh PowerPoint exports retain their earlier raster bytes. DOCX
reference pixels changed, so their comparisons were recomputed. These numerical
changes reflect refreshed references, not renderer changes. Raw prior reports stay
retained; this checkpoint does not retroactively qualify their export settings.

The content screen passes ten outcomes and leaves the complete report's two modes
unresolved. Four rate/time fractions are present as separate glyph paints; its
word-token counter fragments them. Cached field text also differs from displayed
page-dependent values in both Word and candidate. The reviewed pages visibly retain
those fractions, images and page fields. Raw field/token flags and audit exit 1
remain recorded, with no screen or visual gate relabeled as passing.

`NewFrozenAllocationRuntime.ps1` builds only the existing probe against the exact
frozen DLL, verifies both copies and runs its self-test. One isolated cold plus one
warm sample per outcome establishes a compound-document allocation baseline;
`OperatingEnvelope.md` records its limits. Initial font preparation dominates the
measured warm allocation; repeated layout is not established as the main cost.
No speedup or stable timing claim follows from samples taken during Office exports.

No library/root rebuild or local suite replay was performed. C78 source identity
matches the green Windows/Ubuntu CI run 38992440492, including the full Windows
skip-slow suite, portable rendering, package and tool checks. Its binaries are not
claimed identical to the qualified local artifact. Evidence:
`artifacts/composition-c1/checkpoint-79-evidence.json` and linked raw reports.
The latest workspace scan is 9.190 GB; the loose 5 GB target remains unmet. Prior
retirement rejections preserve originals and verified archives; no space reclaimed.

C78 replaces the balloon formatter's run-only ownership check with the existing
shared paragraph ownership check. An actual inserted numbering mark can coexist
with run formatting without forcing a generic inserted-text summary. Source
revisions, inline marker ink and counters remain unchanged; missing/unnumbered/
misplaced/copied mark metadata and incomplete formatting keep their fallbacks.

Four retained Word Review references now match "Formatted: Font: Bold" across
body/cell/header/footer controls. Eight mode outcomes retain source-story text,
ink, font tags, matrices, sizes, spacing and advances; four Final PDFs retain bytes.
Horizontal title/body deltas are .185/.253pt. Vertical residuals remain explicit:
body/footer .559/.541pt, cell/header 1.160/1.152pt. Four Review MAEs/recalls improve;
body/footer clear the recall component, but all eight visual gates still fail
unchanged thresholds. Header/footer pairs were reviewed. No new Word exports.

One frozen Release artifact passes 46 Windows methods and 38 Ubuntu methods/eight
explicit skips, each in one invocation; two new checks fail on C77. Seventeen
assets/four smoke PDFs transfer without a root rebuild or suite replay. All eight
parent comparisons are reused through identical PDFs; four Final comparisons
carry to the candidate. Five fixed/complete input identities match the earlier
zero-inserted-mark inventory, so ten corpus comparisons carry by scope inference
without new corpus conversions or PDF-identity claims. C77 CI is green on both
platforms (38991450463). Evidence:
`artifacts/composition-c1/mark-formatting-balloon-evidence.json` and linked reports.

C77 resolves numbering definitions before advancing shared source timelines.
Concrete IDs sharing a definition continue its sequence; explicit starts apply on
first use in each current/original timeline. Matching valid NSIDs identify the
first XML definition, including its formatting, independently of paragraph-reference
order. Zero/case, distinct definitions/NSIDs, conflicting level text, returned IDs,
instance overrides and marked/unmarked controls establish those semantic rules.
Word SaveAs preserves the concrete IDs in the original failing control; ignored
normalized copies/source XML retain that independent preparation evidence.

Twenty-eight fresh Word exports and four C76 references qualify 32 outcomes: all
116 numeric paints match text/color and the unchanged one-point geometry tolerance
(maximum X/Y deltas .113/.619pt). Both previously unqualified continuation outcomes
now match numeric content in Final and Review. Twenty-six MAEs/recalls improve,
six remain identical with unchanged PDF bytes, and no new gate component fails.
All 32 visual gates still fail unchanged thresholds. Two page pairs were reviewed;
this does not qualify whole-document typography, geometry or richer multilevel
restart interactions beyond existing neighboring coverage.

One frozen Release artifact passes 165 Windows checks and 125 Ubuntu checks/40
explicit skips, each in one invocation; two new rendering checks fail on C76.
Seventeen assets/four smoke PDFs transfer without a root rebuild or suite replay.
A conservative package-wide audit finds no shared used instances or duplicate
NSIDs in the five fixed/complete DOCX inputs. Prior ten corpus comparisons carry
by source-scope inference, without new corpus conversions/PDF-identity claims.
C76 CI is green on both platforms (38990293304). Evidence:
`artifacts/composition-c1/numbering-continuation-evidence.json` and linked reports.

Workspace retirement of 68 archived superseded compiled files was rejected by
automatic approval review before process creation as "blocked by policy". All
originals and the verified archive remain; no bytes are claimed reclaimed. The
four exact snapshot names are added to the blocked-target map and must not be
retried. Evidence: `c76-superseded-compiled-retirement-rejected.json`.

C76 keeps separate current/original numbering timelines after actual inserted
numbering marks. Original state is allocated lazily, shared through source-story
tables and invalidated for unsupported original ownership. Derived counters retain
a real preceding mark as their trigger without inventing source revisions or
balloons. Shared prefix measurement/placement handles old strike/current underline,
suffixes and overflow to the next configured tab; painting owns automatic blue
review ink. Both prefix runs enter the existing font plan. Final paths retain bytes.

Eighteen new Word exports and 22 retained references cover moved/two insertions,
digit transitions, suffixes, custom tabs, level decoration, explicit restart and
body/cells/static math. Of forty retained outcomes, 38 qualify inserted-mark
semantics: all 148 numeric paints match text/color within the existing one-point
geometry tolerance; 34 old-number paints are restored. Twenty actual-mark prefixes
recover underline extent within .456pt. All twenty Review outcomes improve in MAE
and foreground recall, all twenty Final PDFs stay byte-identical, and no new gate
component fails. All forty visual gates still fail unchanged thresholds. Three
page pairs were reviewed; mixed mark/formatting balloons remain generic.

The independent-concrete-ID control exposes a separate inherited defect: Word
Final paints 1/2/3 and the candidate 1/1/2. Both raw modes, metrics and Word facts
remain explicitly unqualified. The explicit-start-override control establishes a
real independent restart. C77 must establish continuation/restart semantics with
independent Office controls before changing counter identity; do not infer a
universal shared-abstract-numbering rule from this single input.

The first production qualification passed 159 neighbors but failed its three new
checks: previous-counter runs were not registered in font preparation. The shared
font-plan correction is requalified on one frozen Release artifact: 162 Windows
checks and 123 Ubuntu checks/39 explicit skips, each once; two rendering checks
fail causally on C75. Seventeen assets/four integration PDFs transfer, without a
root rebuild or suite replay. An earlier snapshot consistency correction occurred
before tests. Initial failed outputs/reports remain; no results from the changed
DLL are carried. Prior ten corpus comparisons carry by C74/C75 source-scope
exclusion, with no new corpus-PDF identity claim. C75 CI is green on both platforms
(38988927008). Evidence:
`artifacts/composition-c1/numbering-counter-evidence.json` and linked reports.

C75 gives an inserted numbering mark its own decoration/revision source. The
shared prefix factory keeps placement-time font/spacing choices while restoring
source ink and actual revision identity. Final inherits paragraph color and
level-only underline; Review uses single revision underline over level styling.
First-run authored formatting stays on its source. Whole-paragraph and unsupported
mark ownership retain their prior paths; no font/size/container guards are added.

Six new Word exports and sixteen retained references qualify first-marker color
and underline in 22 outcomes across body/cells/headers/footers, including explicit
level color/none/double overrides. All nonwhitespace text content, matrices, sizes,
spacing and advances remain exactly equal to the parent. First-marker geometry
retains the one-point tolerance. One frozen Release artifact passes 159 Windows
checks and 121 Ubuntu checks/38 explicit skips, each in one invocation; two new
checks fail causally on the parent. Seventeen assets/four smoke PDFs transfer
without a root build or suite replay.

Eighteen MAEs improve/four remain equal. Recall improves in six, stays equal in ten
and decreases in six by at most .001094, with no new failing gate components.
All 22 visual gates still fail unchanged thresholds. Two page pairs were reviewed.
Word's following old/new counters, their review ink, the underline across the
whole numbering-tab gap and mixed mark/formatting balloons remain explicit gaps.
The fixed ten corpus comparisons carry through C74's source-scope exclusion;
no corpus qualification conversions or new corpus-PDF identity claim. C74 CI is
green on both platforms (38988044843). Latest workspace scan is 9.038GB, so the
loose 5GB target remains unmet. Evidence:
`artifacts/composition-c1/numbering-mark-decoration-evidence.json` and linked reports.

C74 retains the actual inserted numbering mark through the existing paragraph-
mark run. Shared revision ownership admits that source independently of visible
text; copied/misplaced metadata, deleted marks, unnumbered marks and unrelated
revisions retain their fallbacks. Eight static outcomes recover their first
number, and eight mixed-math outcomes recover the omitted first formula, from
14 to 18 mathematical text operations. Body/cell/header/footer controls cover
both presets; all sixteen first numbers meet the existing one-point Word frame
tolerance. Fourteen fresh Word exports and two retained references qualify that
content/geometry result.

Qualification passes 156 Windows methods and 119 Ubuntu methods/37 explicit skips.
The two new rendering checks initially counted shared source stories repeatedly;
only those checks rerun after reference-identity enumeration, carrying 154 Windows
results on the same DLL. Ubuntu runs once; two checks fail on the parent. Twelve
MAEs improve and four remain identical. Recall improves in eleven, stays identical
in four and decreases by .006 in footer Review; no new gate component fails. All
sixteen gates still fail unchanged thresholds. Three page pairs were reviewed.
Marker decoration and Word's old/new review counters remain explicit defects;
recovering the first number does not qualify those behaviors.

The five fixed/complete DOCX sources contain no inserted paragraph marks or block
insertions. Their prior ten comparisons are reused by source-backed scope
exclusion, without new corpus conversions or a new PDF-identity claim. Routine
portfolio inspection now writes summaries by default, with `-IncludeRawStreams`
for focused audits. A math/image PDF retains identical text/graphics/font reports
while avoiding eighteen raw files. Evidence:
`artifacts/composition-c1/numbering-mark-evidence.json` and linked reports.

C73 gives unchanged, text-bearing numbering markers the resolved paragraph color
and numbering-level underline. First-run authored/review decoration stays on its
source; existing marker-revision fallbacks and other decoration policies remain.
Fourteen Word outcomes pin color/underline facts; twelve unchanged-marker outcomes
match all 36 emitted number colors and underline counts. MAE improves in all
twelve, recall stays identical, and no new gate component fails. All fourteen
visual gates remain failed at unchanged thresholds. Two accepted pairs and the
known mark-revision pair were reviewed.

Qualification passes 141 Windows methods and 105 Ubuntu methods/36 explicit skips.
Two older ink checks incorrectly gave the unchanged number its first run's blue
color. Two fresh Word formatting-revision exports establish black numbers and a
blue Summary; only those two checks rerun on the unchanged DLL, carrying 139
Windows results. Ubuntu runs once; two new checks fail on the parent. The external
PDF inspector now reports device fill through its existing graphics-state stack;
default/gray/RGB/CMYK, nested q/Q and literal/comment checks agree across platforms.
Production self-validation is unchanged.

All fourteen outcomes retain their non-whitespace text's parent font/matrix/
spacing fields; newly undecorated separator spaces can omit paint operations.
All ten fixed/complete corpus PDFs retain bytes and comparisons. The two separate
mark-only insertion controls retain parent PDFs and confirmed defects: the first
number disappears in both modes, and Review emits two numbers where Word paints
five, including old/new counter pairs. This is a high-priority content/revision
gap, not a qualified match. Evidence:
`artifacts/composition-c1/list-decoration-evidence.json` and linked reports.

C72 adds a test-only, hash-pinned DejaVu Math TeX Gyre face with retained upstream
2.37 notices. TrueType outlines, MATH constants and the five mathematical italic
letters missing from the existing Sans face are checked explicitly. Forty
ordinary-math composition matrices cover body/cells/headers/footers, bookmarks,
run splitting, revisions and images in both public presets, without rewriting
authored math to normal text. Eight composed pages also pass serialization.
The old missing-glyph and explicit-normal controls remain.

Twelve selected Windows methods and eleven Ubuntu methods/one explicit native
skip pass, each in one invocation. A test-only build leaves the C71 library
byte-identical; its independent Office and corpus comparisons are reused. This
adds portable content/resource coverage, not an Office-typography claim or a
runtime dependency. Evidence:
`artifacts/composition-c1/portable-default-math-evidence.json`.

C71 corrects undeclared Latin prose fallback across body, cells and static stories.
After an explicit resolver document default, the fallback prefers Word's Times
New Roman before the existing modern-face alternatives. Explicit family, theme
and alternate resolution retains its precedence; no font-size or shape guard is
added. Four fresh Word exports at 11/14pt plus two retained controls establish
38 matching emitted anchors. Qualification passes 150 Windows and 116 Ubuntu
methods/34 explicit skips, each in one invocation. Two checks fail on the parent;
its portable custom-default control still passes.

All six focused visual gates still fail unchanged thresholds, but MAE and recall
improve in every outcome. At 14pt, MAE falls from 2.704 to 0.221 and foreground
recall rises from 0.276 to 0.958. Two page pairs were reviewed; ten fixed/complete
corpus PDFs retain their bytes and comparisons. Evidence:
`artifacts/composition-c1/default-typeface-evidence.json` and linked reports.

Source inspection also refines the earlier list-family header warning: its
related header has no section reference. Both Word and candidate omit that unused
story. The package contains an equation, but this warning does not establish a
visible omission. Raw diagnostics, source counts and earlier severity remain
alongside `artifacts/composition-c1/list-family-unused-header-classification.json`;
neither diagnostic policy nor visual gates are relaxed.

C70 bounds new qualification-asset duplication within each fresh frozen output.
Identical assets share files there; writable builds, other artifacts and transfer
copies remain independent. Seventeen tool checks exercise those ownership rules.
The real artifact loads its catalogue, verifies all C69 asset/input hashes and
avoids 11,294,694 duplicate bytes across five aliases. Rendering qualification is
reused without a build, suite, conversion or root transfer. This prevents growth;
the existing workspace still exceeds its loose 5GB target. Evidence:
`artifacts/composition-c1/asset-sharing-evidence.json`.

C69 extends the same semantic numbering owner to font families and script slots.
The level overrides the resolved paragraph mark/styles/defaults; an undeclared
family uses Word's Times New Roman label fallback. Shared label construction
keeps absent slots from borrowing first-run formatting. Size, color, underline
and revision-decoration policies remain independently qualified.

Six source transformations in clean Final and visible Review produce twelve Word
references and 36 matching emitted label families. Qualification passes 139
Windows methods in one invocation and 108 Ubuntu methods/31 explicit skips in
one invocation. Both new checks fail on the parent. All twelve visual gates still
fail unchanged thresholds. Against six parent Final controls, MAE improves in
five and stays identical in the explicit-level control; recall decreases in the
unstyled fallback case, with no new failure components. Two page pairs were
reviewed. Unthemed non-list prose retains a different fallback from Word there.

All ten fixed/complete corpus PDFs retain their bytes, reusing comparisons.
List-style manifests now count the inherited unsupported header equation rather
than declaring zero; its diagnostics and raw measured reports remain. A metadata-
only repin verifies all seventeen assets unchanged without repeating tests or
conversions. Evidence: `artifacts/composition-c1/list-font-family-evidence.json`
and its linked source-count correction, face, visual and preservation reports.

C68 resolves list-label weight and slant from their numbering/paragraph owners.
Explicit level values take precedence; missing values inherit the already-resolved
paragraph mark, including styles and defaults. The first text run keeps its own
formatting. Independently validated label-size precedence and existing color,
underline and revision-decoration policies remain. The parser reuses its resolved
mark instead of adding font-, size- or document-shape qualifications.

Ten independent Word exports and eighteen retained references cover level, mark,
Normal, ordinary/inserted bold text and body/cell/header/footer combinations. All
84 emitted label faces match Word across 28 outcomes, verified through embedded
TrueType weight/italic tables and Word font dictionaries. Qualification passes
126 Windows and 99 Ubuntu methods/27 explicit skips. A test reader omitted
synthetic italic matrices; an explicit reader option corrects it. Only the two
affected style methods rerun on the same library, carrying 124 Windows results;
Ubuntu runs once. Two new checks fail on the parent. The old authored-bold
assertion is corrected using the independent ordinary/inserted controls.

All 28 visual gates still fail unchanged thresholds. Among eighteen retained
comparisons, MAE improves in seventeen and worsens by 0.000074 in the second-
paragraph footer Review control; recall decreases in eight. No new failure
component appears. Two page pairs were reviewed. All ten fixed/complete Final/
Review PDFs retain parent bytes, reusing their comparisons without raster replay.
The font-face result is not a claim of complete Office fidelity. Evidence:
`artifacts/composition-c1/list-font-style-evidence.json` and linked reports.

C67 keeps supported static content in its frame when the supplied math font lacks
glyphs. The pinned face has math constants but lacks five mathematical italic
letters used by these controls. Their equations correctly retain unsupported
diagnostics; rejecting the entire natural frame also displaced surrounding prose
and pictures below the page. Static fallback now measures the remaining ordered
items under the same source-owner rules. Body/cell callers retain their strict
default; available math boxes, bounds and cancellation remain unchanged.

The portable regression covers eight source controls in both public presets,
preserving three labels, prose rows, pictures and source runs while retaining the
unsupported diagnostics. Qualification passes 90 Windows and 74 Ubuntu methods/16
explicit skips. Expanding the harness to both presets reruns only three affected
checks on the same production DLL, carrying 87 unchanged results. One new check
fails on the parent. Independent probes establish the same glyph/frame failure
on C65 and C66. At footer distance 18, the pinned PDF summary returns from Y −16.61
to 150.84; both before/after conversions retain the same two diagnostics. Four
pinned before/after PDFs are inspected and one full page pair is reviewed.

Eight native control PDFs remain byte-identical, reusing their independent Office
comparisons and eight failed visual gates without inspection/raster/diff replay.
No new Office exports or corpus conversions. The five retained native portfolio
outcomes have zero ignored-equation diagnostics, leaving the new missing-math
branch inactive; this is diagnostic/source-based reuse, not fresh qualification.
Missing-font equations remain unsupported; list-label style and review geometry
are separate residuals. Evidence:
`artifacts/composition-c1/static-missing-math-evidence.json` and linked reports.

C66 restores text-only static list labels through the existing ordered inline
measurement and placement path. A bounded static-list entry uses shared semantic
source ownership; other callers keep their existing default. Ordinary paragraphs
without a parser-side mixed-content view derive items from their original source
runs, without rewriting the parsed model. Pictures, exact/minimum spacing and
unsupported owners/tabs retain their explicit legacy fallbacks.

Four footer/header controls cover first/second formatting changes in both public
presets. All 24 expected labels render; independent Word coordinate differences
stay below 0.33pt. Six clean Final/mixed references are exported with view read-back
and two footer Review references are reused. Qualification passes 83 Windows and
68 Ubuntu methods/15 explicit skips, once per platform; two new checks fail on the
parent. The same production/CLI bytes qualify the source-pinned measurements and
the final harness, without repeating eight conversions. Five fixed/complete inputs
have no static or inherited-style numbering requests, so retain their corpus
evidence. The tested header Final body text operations match the parent exactly.

All eight visual gates still fail unchanged thresholds. Two full page pairs were
reviewed; list-label weight in revised paragraphs, typography and review connector/
bar geometry remain approximate. The compound pinned-font math frame fallback is
separate from this text-only migration. Evidence:
`artifacts/composition-c1/static-list-labels-evidence.json` and linked reports.

C65 restores text-only footer prose that fell below the page when removing math
and pictures left bookmarks, links, numbering and formatting metadata. Natural
static frames now share the existing inline source-owner rules; math eligibility
keeps its equation requirement and legacy note guards remain unchanged. Footer
formatting placement consumes the measured physical story top, independently of
the body keep-out edge and the changed line's connector. No fitted displacement
or font/size-specific branch was added.

Eight source-identified Word controls cover footer distances, first/second changed
paragraphs, text-only transformations and the original header/footer combination.
The two plain footer baselines return from below the page to within 0.33pt of Word;
their balloon baselines differ by 0.007pt. Four complex footer balloons differ by
at most 0.57pt. Header paragraph translation and the source connector's horizontal
path remain unchanged. The original header balloon retains its 1.152pt residual.

Qualification covers 81 Windows passes and 67 Ubuntu passes/14 explicit skips.
Harness-only corrections retain the production DLL: two frame checks rerun after
accounting for faux-bold paint; two comment checks rerun after making their font
requirements explicit. Two new checks fail causally on the parent; two old Linux
font assumptions also fail on the parent. The portable frame check covers the two
text-only controls; compound pinned-font math keeps its unsupported-frame fallback.
Eight visual gates still fail unchanged thresholds, and twelve diagnostics remain,
including six ignored header-equation diagnostics. Two full page pairs were
reviewed. Four fixed Review PDFs and the complete composed Review report retain
their prior bytes/comparisons; no Final corpus replay or new Office exports.
Plain static list labels are still missing and are the next content priority.
Evidence: `artifacts/composition-c1/static-source-frames-evidence.json` and linked
reports; the portfolio preservation report records a corrected manifest equation
count without repeating conversions.

C64 skips the font-resource/layout retry when complete measured footnote owners
already have the requested labels. On a public one-page control, three warm
calling-thread allocation samples fall from a 19.54 MB median to 10.46 MB, saving
9.08 MB (46%). Its PDF stays byte-identical to both C62 and C63. The two-page
correction still retries and retains C63 PDF bytes; allocation changes by 72 bytes.
Stage measurements overlap and timings are noisy, so no timing speedup is claimed.

Ownership validation is independent of replacement: minimum-leading and font
adjustments now request the complete source-marker/page map without constructing
a numbered document. The initial no-op exit disabled those callers and CI caught
nine geometry/link regressions. The correction restores all nine, retaining the
single-line eligibility and matching note-source prefix checks. C63's stale
mixed-threshold expectation was also corrected; neighboring selections now include
mixed thresholds, spacing, minimum leading and contextual note links.

Final qualification covers 73 Windows passes and 71 Ubuntu passes/two native-font
skips. One new harness check initially used a wrapped fixture outside the legacy
single-line scope; only that corrected method reruns on the same library, reusing
72 unchanged results. Both allocation medians and PDF hashes remain identical to
the initial C64 measurements. Four typography PDFs retain their independent Office
comparisons, including all four failed gates and approximation warnings. Seventeen
assets/four affected Final/Review/deck smoke PDFs transfer without root rebuild or
corpus replay. The tracked generator reproduces the no-op fixture exactly. Evidence:
`artifacts/composition-c1/note-owner-validation-evidence.json` and
`artifacts/composition-c1/note-owner-validation-transfer/evidence.json`; the initial
qualification and CI failures remain alongside them.

C63 fixes default each-page footnote numbering through complete measured owners.
The old default path retained document labels 105–108 while Word and the review
path printed section labels 09/10 on each page. The existing source-marker/actual-
segment/same-page-note map now drives numbering independently of font adjustment.
Legacy adjustment, leading and reserve guards remain; known settings/full fonts,
same-page wrapped paragraphs, complete one-line notes and one stable-page retry
remain required. No font/size-specific numbering branch was added.

Four independent Word controls cover Calibri 12 zero/met thresholds, Arial 11 and
Times 14. Both presets retain correct labels, authored 105/108, original source runs
and 32 distinct same-page destinations. 48 Windows methods and 46 Ubuntu methods/two
native-font skips pass once; two new checks fail on the parent. The existing
correct Review PDF and four transfer smoke PDFs stay identical; 17 assets transfer
without root rebuild/suite replay. Direct body tables exclude the five fixed/
complete DOCX inputs, so their corpus qualification is reused. All four default
visual gates still fail unchanged thresholds and their footnote approximation
warnings remain. First-page Calibri 12-zero and Times 14 pairs were reviewed;
typography/geometry remain approximate. The default view now uses the existing
bounded retry; no speedup or unchanged-runtime-cost claim. Evidence:
`artifacts/composition-c1/measured-footnote-owners-evidence.json` and linked reports.

C62 closes the supported-polygon picture blur fallback from the geometry handoff.
Picture masks, vector shadows and blurred shadows now share the existing fixed
preset vertices. A safe span-based distance query handles convex/concave polygons
and maps image rows to PDF Y; the same bounded raster planes, image budgets and
per-row cancellation remain. Arbitrary custom geometry keeps its diagnosed
rectangular frame. The shipped library remains safe managed and dependency-free.

Frozen d9f598c1 qualifies 43 Windows and 43 Ubuntu methods once each, with zero
skips and one causal parent failure. Five public decks cover 26 page pairs: 25 gates
pass and the prior alpha-color case remains failed. Triangle SSIM improves from
0.87693 to 0.99813; rotated right-triangle, star and grouped-chevron controls pass
unchanged 0.995 SSIM/0.98 color thresholds. The four changed/new pairs were reviewed
including full transformed/clipped footprints and following paint. Nineteen old
picture pages retain PDF identity; all compared prototype PDFs match the frozen
artifact, so raster/diff evidence is reused. Four transfer smoke PDFs match after
17-asset transfer, without root rebuild or DOCX corpus replay. Evidence:
`artifacts/pptx-picture-geometry/polygon-blur-evidence.json` and linked reports.

The image-only portfolio exposed a tool assumption: the PDF inspector omits an
empty text-operation report. Measurement now treats that as no text and counts
page objects independently. The failed first measurement is retained; the same
production DLL then measures all five decks / 26 pages without rebuilding.

C61 matches white formatting-balloon fill in seven independently inspected Word
paths, retaining comment-author colors, mixed summaries and other profiles.
Thirty-five Windows checks and four pinned-font Ubuntu checks pass once each;
two new color checks fail on the parent. Eight Review conversions retain identical
text operations, path geometry and strokes. Seven fills change; the void PDF and
four transfer smoke PDFs stay identical. No Office exports or fixed corpus replay.
All eight visual gates still fail unchanged thresholds: MAE improves but foreground
recall decreases in seven nonvoid cases, with new cell/footer component failures.
The native fill correction is accepted with these metric residuals explicit.
Evidence: `formatting-white-fill-evidence.json` and linked raw reports.

Six independent header/footer transformations distinguish footer story anchoring
from header paragraph anchoring. Exact occupied-frame placement remains unresolved;
no fixed per-story offset was accepted. These findings guide subsequent placement
work and do not qualify new geometry.

C60 corrects mixed-view run-formatting descriptions through resolved before/after
bold/italic state. Seven independent Word controls establish "Formatted: Font:"
with Bold, Not Bold, Italic or Bold/Italic descriptions across body, cell, header
and footer. Unsupported/incomplete, malformed, character-style and unowned changes
retain generic summaries. Other markup profiles retain their prior labels. The
same source-owned metadata drives measurement and painting; source runs stay intact.
Independent title/description X positions retain one-point checks.

The unchanged production DLL qualifies 27 Windows methods in two invocations,
including two old tests corrected to use deterministic embedded fonts. Ubuntu:
23 passes/four installed-font skips. Two new content checks fail on the parent;
the two old font failures also reproduce there. Sixteen frozen control outcomes
retain 18 mathematical operations; all eight Final PDFs and the void Review PDF
retain parent bytes. Fourteen existing header equation diagnostics remain, with
no aggregate content-audit pass claimed. Fixed corpus evidence is reused and four
integration PDFs match after checked transfer of 17 assets. All sixteen visual
gates still fail unchanged thresholds. Body/footer visual pairs retain corrected
text and visible content; footer balloon Y, fill, typography and connector/bar
geometry remain. Evidence: `formatting-description-evidence.json` and linked reports.

Latest bounded batches, C56–C57: unchanged equations survive adjacent source-owned
run-formatting revisions in body, cell, header and footer paragraphs. Five public
controls cover a bold change and an empty change; all ten Final/Review outcomes
retain 18 mathematical operations instead of dropping the first fraction in eight
outcomes. Parsing, mathematical preparation and static frame mapping share actual
run ownership. Unowned, paragraph/mark, move and equation revisions keep fallbacks.

Independent Word 16 mixed-mode references establish that direct static formatting
balloons fit the page. Shared frame mapping reduces intermediate errors up to
163pt to maximum X/Y errors of 0.877/0.768pt, retaining the one-point check. C56
qualifies 23 Windows methods once, 18 Ubuntu checks/five native-font skips and two
causal parent failures. Twelve existing document PDFs retain their bytes.

C57 separates formatting metadata from text-revision ink in the shared glyph and
decoration paths. Eight independent Word text operations confirm black glyph ink;
authored color/underline transformations remain covered. Formatting balloons,
change bars, insertion/deletion/move behavior and source geometry remain intact.
Frozen `44d45883` qualifies 18 Windows methods once, 14 Ubuntu checks/four
native-font skips and two causal parent failures. All five Review text/graphics
inspections retain C56 bytes. Fixed documents contain no property changes, so
prior corpus qualification is retained. Both batches transfer 17 assets/four
identical smoke PDFs without root builds or repeated suites.

All ten C56 and five C57 numerical gates still fail unchanged thresholds.
Header/footer and body Review page pairs retain visible equations and images;
typography, bar extent and balloon content/geometry still differ from Word.
Formatting-only text and labels now retain authored ink. No aggregate prose audit
or complete fidelity claim is made for these controls. Evidence: ignored
`artifacts/composition-c1/math-run-format-evidence.json` and
`formatting-review-ink-evidence.json`.

C58 consolidates bounded kerning-token parsing while retaining qualified frame
eligibility. C59 gives ordinary source baseline runs one advance purpose shared by
measurement and emission, without rewriting source runs or legacy note/numbering
state. Omitted or zero thresholds disable pairs, positive thresholds use authored
size, and glyph advances use nominal size while export glyph sizes stay rounded.
Generated, shifted, malformed and automatic-marker runs retain existing purposes.
The [Open XML kerning definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.kern?view=openxml-3.0.1)
establishes threshold/inheritance semantics; independent Word exports establish
the observed nominal-advance behavior. Eight controls span three font families,
11/12/14pt and omitted/zero/met/unmet thresholds. Both renderer presets retain exact
line text, 0.15pt relative line-gap and 0.1pt first-line endpoint checks.

One unchanged production library qualifies 1,641 Windows milestone methods. Two
harness-only corrections use a 1e-9pt roundoff tolerance and deterministic tracking
font; three runner invocations total, with no full suite replay. Ubuntu qualifies
32 methods with six installed-font skips; two regressions fail on the parent.
Seventeen assets/four smoke PDFs transfer without root builds. All fourteen fixed
DOCX page comparisons and four complete-report comparisons improve MAE and
foreground recall. Twelve comparisons are reused by candidate/reference/threshold
identity; eight new control comparisons retain their failures. All eight fixed
DOCX, two complete-report and eight control gates still fail unchanged thresholds;
the two unchanged decks pass. Ten fixed content audits have zero deficit/unresolved
screens. Both complete-report audits retain field extraction uncertainty and
prose/math-token differences; expected simple math is present.

Single warm samples on report/memo Final show allocation rising 1.1–1.2%, while
explicit advance compensation adds 22–29KB per PDF (28–35%). Stage and whole PDFs
match. Timings are noisy and stages overlap; no speedup is claimed. Evidence:
`source-advances-evidence.json` and linked reports under the ignored evidence root.

The fixed [portfolio](../tests/Lokad.OoxPdf.Tests/Cases/Composition/portfolio.json)
contains four synthetic DOCX documents and two existing composite decks. The
report, letter and specification exercise mixed content; the technical memo and
second deck are held out from fitting. The DOCX documents have different lengths,
and the memo moves the table after its supporting discussion. This is a bounded
reliability checkpoint; version 0.1.5 release preparation remains deferred.

Each DOCX source has six equations across its body, table and header, an inline
image, numbering, bookmark, external link, footer, footnote and comment. Both
default Final/preserved layout and AllMarkup/Word-compatible output are measured.
Expected prose and source counts are pinned before layout corrections. Source
counts do not establish that every item painted or remained visible.

The initial baseline retains all expected prose but ignores equations in all
eight DOCX mode outcomes. All eight fail existing DOCX numerical visual gates;
the report and held-out memo also differ from Word in AllMarkup page count. Both
decks pass their inherited numerical gates. Passing a raster gate alone does not
establish complete fidelity. Missing supported content and significant pagination,
overlap or clipping are checkpoint blockers. Typography and cosmetic residuals
remain explicit and thresholds are unchanged.

The first correction removes bookmarks from math exclusion policy: source
annotations have no measured or painted extent. Transformation regressions cover
standalone and inline body equations, a cell and direct header/footer paragraphs,
in both modes, including unchanged-format prose run splitting. Nineteen capability and neighbor methods pass in one invocation; the new
regressions fail on the previous library. Seventeen frozen assets transfer with
four identical integration smoke PDFs and no root rebuild. Independent bookmarked
Word controls still fail foreground/SSIM gates; no complete visual pass is claimed.
Equation-only bookmark navigation remains unresolved; this correction does not
complete the checkpoint.

Reusable commands:

- `NewCompositionPortfolio.ps1` authors fixtures; do not regenerate pinned inputs
  during qualification. ZIP metadata may differ on regeneration.
- `MeasureCompositionPortfolio.ps1` converts using an explicit existing CLI and
  inspector, verifies identities and reports content findings separately by mode.
- `CompareCompositionPortfolio.ps1` uses independent Office PDFs and existing
  numerical thresholds. Cache references for unchanged source/mode identities.
- `RunDocxAffectedTests.ps1 -Capability docx-inline-annotations` runs a tracked
  capability plus neighbors as one exact method set. `--test-set file.json` also
  accepts an explicit JSON array directly in the console test runner.

Cold CLI wall time includes process startup. Calling-thread allocation and stage
measurements belong to the separate allocation probe; no speedup is claimed.
Retained local baseline evidence is under `artifacts/composition-c1/`.


The next component retains ordered text, image, equation, tab and break items for
math-bearing source paragraphs. Inline equation offsets and visible-prose policy
consume this view. Original object identities and image positions within a run
are preserved; zero-width annotations stay metadata. Fields, references, textboxes
and revisions retain the legacy source path. Replacing a legacy content collection
invalidates its borrowed view; placement-only changes preserve it. This source foundation preserves the baseline output identities.

Twenty-seven selected methods pass in one invocation; six portable source/bookmark
methods also pass on the exact artifact in Ubuntu. Seventeen assets and four
integration PDFs transfer without root rebuild. All ten portfolio PDFs
retain their baseline bytes, so their Office comparisons are reused. Mixed-document
warm allocation medians rise by approximately 11–25KB for the DOCX sources; deck
allocation medians are unchanged. Three-iteration timings are noisy and establish
no speedup. These measurements precede any repeated-layout optimization.


The bounded mixed-content path prepares each equation by its original identity.
Ordered fragments retain measured text segments, advances, baseline insets and
occupied heights. Body, cell and direct header/footer containers use the same
wrapping and placement functions, including inline pictures and multiple fitting
inline equations. Placement translates measured geometry instead of measuring
text again. Original text-run affinity survives hidden-run filtering; character
spacing and pair kerning survive tokenization. Long words reuse the existing
bounded Unicode wrapping rules. Keep-next reserves the complete mixed frame.

Eight separate probes place two equations and one picture in the body, a cell,
a header or a footer, with and without review scaling. They supplement the fixed
portfolio and do not replace its held-out documents. Source-to-page mapping moves
text, equations and pictures together; one existing vertical-frame owner applies
to all three. Independent Office comparisons retain their numerical failures.
Content preservation does not establish Office geometry or unclipped painting.

Display combinations, tabs, fields, revisions, list labels and rich source owners
retain explicit fallbacks. The older single-equation marker path remains until a
separate measured migration preserves its Office comparisons. This is a bounded
migration, not completion of the composed-document checkpoint.


The mixed-content Release artifact passes 59 capability and neighboring methods
in one invocation with no skips. Eight DOCX portfolio outcomes now retain all
expected prose and have no ignored-equation diagnostic. The development report's
native checks verify all six equations and an actual image drawing in both modes;
sixteen placement/mode checks verify two equations and one image drawing each.
All 22 placement PDFs match their independently compared prototypes exactly, so
the Office comparisons are reused. Both deck PDFs also retain their baseline bytes.
All eight DOCX portfolio visual comparisons still fail; the report and memo retain
their AllMarkup page-count mismatch. These residuals block checkpoint completion.


The `rendering-portable` console group adds three small actual-rendering checks:
DOCX mixed content in both modes and a PPTX composite chart. All use the same
pinned DejaVu Sans font, license and provenance under `Cases/PortableFonts`.
The dedicated DOCX fixture authors normal mathematical text because this font
lacks the portfolio's default mathematical italic alphabet. Ubuntu CI runs this
group alongside its existing portable groups. Windows retains the full suite.


The next paragraph-frame correction compares already scaled advances with the
printed body capacity for single-column review output. Headers, footers and
other document features no longer select a larger paragraph capacity. Three
independent exact-height controls preserve two Word pages in both modes with no
static story, a header or a footer. The composed report also keeps Word's break
between review sections 8 and 9. Default report output remains byte-identical.
Visual gates remain unmet; matching pagination alone does not establish fidelity.

Table fragment capacity and its existing keep/widow fallbacks remain separately
qualified. A broader prototype exposed disagreement between measured membership
and painted continuation lines when combining printed capacity with the older
fragment baseline path. That prototype is retained as rejected development
evidence; table baseline/membership ownership needs its own bounded correction.
Following body paragraphs still use paragraph capacity after a table fallback.


The paragraph-frame artifact passes 46 selected methods in one invocation and
three portable rendering methods on Ubuntu. All ten portfolio mode outcomes have
matching page totals, expected prose and no ignored-equation diagnostic. The
held-out memo recovers its second AllMarkup page. Four default DOCX PDFs and both
decks remain byte-identical. Four integration PDFs match after transferring the
17 qualified assets without a root build or repeated suite. All eight DOCX raster
gates still fail. Fourteen of sixteen comparisons reuse exact candidate/reference
PDF and threshold identities; only two need new raster comparisons.


A subsequent content-preservation regression reproduces missing table words on
the delivered artifact: explicit widow control paints 117 of 120 synthetic words.
Independent 200-word public controls paint 198 in review output on the parent.
Automatic rows now publish their printed first-baseline geometry before fragment
membership, using the same natural-cell rule for admission and painting. The
later table translation does not scale that inset again. Keep/widow page-capacity
fallbacks, authored heights and rich-cell fallbacks remain explicit. Focused
coverage checks all words and baselines inside fragment bounds across lengths and
widow settings; emitted text alone does not establish unclipped visual parity.
The independent controls retain numerical visual and pagination failures.


The table-fragment artifact passes 43 exact methods in one invocation; the three
portable renderers and the word-preservation regression also pass in Ubuntu.
All ten fixed portfolio PDFs retain their accepted bytes. Four independent table
control PDFs match their compared prototype identities. Seventeen assets and four
root integration PDFs transfer without rebuild or repeated suite. Native controls
retain pagination/visual failures and three unresolved Office clip matches per
review source; candidate clips are classified. This fixes content loss within the
bounded natural-cell path and leaves broader frame/fidelity work explicit.


Natural automatic text-only body-table rows now receive the same printed capacity as body
paragraphs. The containing frame supplies this height once; fragment projection
scales baseline geometry separately. Headers, footers and unrelated document
stories no longer choose a larger table frame. Authored row heights, repeated
headers, explicit row/paragraph keeps and rich cells retain staged fallbacks.

Widow-protected splitting can choose one earlier whole-line boundary and recheck
every cell, retaining at least two continuation lines. This bounded retry honors
cancellation; invalid boundaries retain the whole-row fallback. The shared line
measurements determine both admission and painting. Two-cell checks ensure an
unprotected first cell cannot bypass the next cell's widow rule or lose content.
Pinned long-table controls also cover plain headers and footers in both modes.


The held-out memo rejected an initial broader frame candidate: it moved a
mathematical row to the next page, whereas Word keeps that row with its heading.
Text-metric eligibility now explicitly excludes equations as well as pictures
and textboxes; mixed-object rows retain their established frame and painting.
This semantic ownership boundary applies to every source and is not keyed to the
memo, a font, a size or a fitted threshold. Wider object-height ownership remains
unqualified. Held-out comparisons remain acceptance checks, not fitting inputs.


A transformation regression also retains all prose when unsupported mathematics
falls back. Font-baseline eligibility stays separate from frame-capacity
eligibility: an ignored object cannot switch off the published prose baseline.
The regression catches 117/120 words in an intermediate candidate and passes after
restoring the shared line-metric owner. Neither rejected candidate is delivered.


The accepted natural text-table frame passes 54 selected methods in one invocation
and five portable checks under Ubuntu. All eight long-table mode outcomes preserve
200 words and match Word's four-page totals; other numerical gates remain unmet.
All ten fixed portfolio PDFs retain accepted bytes, including the held-out memo.
Eighteen independent comparisons reuse PDF/reference/threshold identities. Four
root smoke PDFs match after transferring 17 assets without rebuild or suite replay.
The two rejected intermediate revisions remain retained as development evidence.


Equation-only bookmark navigation now consumes measured paragraph-frame origins
and tops across body, cell, header and footer stories. Those paragraphs do not
manufacture text segments for navigation. The same relative frame origin survives
alignment, translation and static-story scaling; existing text destinations keep
their qualified viewport rule. Mixed-inline/multiple-equation source annotation
binding remains a separate migration.

A compound public fixture has four links and four equation-only story targets,
with an explicit review comment. Word exports all four destinations in both modes;
the parent drops them. Candidate text and graphics inspection reports match the
parent byte-for-byte in both modes. Default destination differences are below
one point; review-mode vertical differences reach about 3.4pt, reflecting retained
layout residuals. No fitted destination correction or visual-threshold change is
introduced. A normal-math variant exercises all four targets with the pinned
portable font, independently of installed Windows fonts.


Navigation qualification passes 33 selected methods in one invocation and all
four actual portable renderers under Ubuntu. Each compound mode now exports four
links, matching Word's target count. The corresponding PNGs, text operations and
graphics operations retain parent bytes. All ten fixed portfolio PDFs are exact;
17 assets and four integration PDFs transfer without rebuild or suite replay.
Coordinate residuals and failed visual gates remain explicit; mixed-inline source
annotation binding is still open.


Standalone body equations now appear as `Equation` in layout inspection, with
source block and column ownership, structural box bounds, explicit baseline and
occupied line height. Previously they appeared as `Unknown` with zero geometry.
These measurements describe layout frames, not glyph ink. The inspection-only
change passes 14 selected methods in one invocation and the pinned-font regression
on Ubuntu, where both DOCX modes are exercised. Four integration PDFs retain their
previous bytes after transferring 17 qualified assets without a root rebuild or
suite replay. No visual-fidelity improvement is claimed for this telemetry fix.


Bookmarks at formulas in the ordered mixed-content path now retain an inline
source boundary. Navigation waits for that formula's actual placed line, including
lines with no prose segments. The shared paragraph frame supplies the origin in
body, cell and static stories; existing run-relative fallback remains available.
A new control combines two formulas in each of those four stories, an inline image
and an explicit body line break. Word exports eight destinations in both modes;
the parent exports none. The candidate restores all eight without changing the
control's raster, text operations or graphics operations.

This is navigation recovery, not equation-geometry parity. Word centers adjacent
formulas as a display group with larger fraction typography; the candidate still
uses left-aligned inline math. Target differences reach 218.5pt
horizontally in Final and 165.9pt in review mode; vertical differences reach about
8pt. Both numerical visual gates still fail. Investigate the common math-only
paragraph layout before subpoint refinements; do not fit navigation offsets to
Word while painting formulas elsewhere.

The accepted batch passes 60 selected methods in one invocation and six portable
rendering/inspection checks on Ubuntu. All ten fixed portfolio PDFs are exact,
so their independent comparisons reuse verified identities. Seventeen assets and
four integration PDFs transfer without a root rebuild or suite replay. A new
metadata test initially omitted the separate static-story source enumerator; its
correction changes no production code. C1 remains open.


Fitting math-only rows with multiple adjacent formulas now share display-style
measurement and center the complete group in body, cell and static containers.
The same alignment helper consumes the measured group width in all three paths;
source equations remain immutable. Mixed prose, pictures and authored hard-break
rows retain their existing inline path in this component. Navigation retains
container context at line start and uses the placed position of a later formula.

Independent controls place two, three and two formulas in left-, center- and
right-aligned paragraphs. Word centers each math-only group. Final foreground
recall improves from 0.276 to 0.755; review recall improves from 0.988 to 0.994.
Other numerical visual gates still fail. In the compound navigation control,
maximum horizontal destination difference falls from 218.5pt to 0.54pt; remaining
vertical differences reach about 8pt. The isolated formula after an image/break
still uses inline typography and left alignment; its shared row ownership is the
next rendering problem, without fitting a paragraph-specific correction.

Qualification passes 71 selected methods in one invocation and seven portable
rendering/inspection checks on Ubuntu. All ten portfolio PDFs remain exact, with
14 independent comparisons reused by source/reference/PDF/threshold identity.
Four smoke PDFs match after transferring 17 assets without root rebuild or suite
replay. C1 remains incomplete and all eight DOCX portfolio gates remain failed.


Authored breaks now delimit shared source rows for mathematical style. A formula
on a row without another visible inline owner uses display measurement, including
when another row of the paragraph contains an image. Prepared boxes retain that
style, so wrapping and container placement consume the same decision. Soft
wrapping stays unchanged. Invisible prose cannot change row ownership; its
metamorphic geometry regression fails on the parent and passes after aligning the
source visibility rule. Paragraphs without prepared math bypass row classification.

Independent image/break variants cover body, cell, header and footer, separately
in both modes. All eight destinations survive in each outcome. Final target
differences are at most 0.52pt horizontally and 0.93pt vertically; review residuals
reach 0.54pt and 3.39pt. All eight numerical visual gates still fail. Foreground
recall improves in every control, but review cell MAE increases about 0.004;
retain that nonmonotonic residual rather than claiming universal metric gains.

The frozen batch passes 73 selected methods in one invocation and eight portable
rendering/inspection checks on Ubuntu, including both modes in all four containers.
All ten fixed portfolio PDFs remain exact. Eighteen independent comparisons reuse
verified identities. Seventeen assets and four integration PDFs transfer without
root rebuild or suite replay. C1 remains open; legacy single-inline marker paths
and broader composition/geometry concerns remain to be addressed in bounded slices.


Single-equation prose paragraphs now keep one qualified layout owner whether or
not their unchanged-format text carries a hyperlink. Native literal/fraction/
script regressions and a pinned-font compound-letter regression demonstrate that
annotations no longer change measurement or placement. Body and cell bookmarks
consume actual placed math geometry, including the body's separate equation item,
with source/block identity retained instead of resolving to following prose.
Independent two-story targets differ from Word by at most 1.45pt X and 0.94pt Y.

Blanket migration to the shared whole-object wrapper was evaluated and rejected.
Word breaks the overfull literal at an operator; whole-formula relocation adds a
line and reduces recall from 0.375 to 0.329. Both old relocation fallback assertions
failed. That prototype is retained separately and is not delivered. Operator
breaking remains a distinct capability required before retiring the legacy marker
path; the accepted annotation fix preserves both fallbacks.

Qualification passes 82 selected methods in one invocation and nine portable
rendering/inspection checks on Ubuntu. The eight DOCX portfolio outputs retain
page totals and show small recall/MAE improvements; both decks retain bytes.
All eight DOCX numerical gates still fail. Ten final PDFs match the compared
prototype, so independent comparisons reuse their verified identities. Seventeen
assets and four smoke PDFs transfer without root rebuild or suite replay; the
smoke includes exact parent output from the overfull fallback control.


`tools/AuditCompositionContent.py` screens complete public source prose and simple
literal/fraction expressions against the measured candidate and cached Office
PDFs. It pins frozen inputs/library, PDF/comparison identities and inspector
outputs. Referenced headers/footers repeat per reference page; only referenced
notes and mode-appropriate comments contribute. Complex source features remain
explicitly unresolved. This is a token-presence screen, not reading-order, image,
ink, clipping or geometry proof; numerical visual gates remain independent.

The first audit reuses frozen `1845342c` without conversions or library builds.
All ten source-prose/simple-math screens pass. Four review-mode Office token
screens flag the generated comment caption: Word prints `Commented [1]:`, while
the candidate substitutes `Commented [Public reviewer1]:` when initials are absent.
This is a confirmed caption difference, not missing body prose. It remains open.
The same audit functions identify exactly two missing `AAA` tokens in retained
198/200-word table output and no source-prose deficit after the accepted correction.


Word-compatible comment captions now use explicit initials and encounter-order
numbers, without substituting author names. A five-case Office control confirms
`[1]`, `[2]`, `[RV3]`, `[4]`, `[QA5]` for absent/empty/explicit initials, empty
authors and nonsequential source IDs. Ordinary author titles, palette and thread
rules stay on their existing paths. The portable regression fails on the parent.

Twenty-three selected title/wrapping/thread/palette/rendering methods pass in one
invocation; ten portable methods pass on Ubuntu. The four Final DOCX PDFs and both
decks retain bytes. Four review outputs improve MAE on the affected pages, while
all eight DOCX visual gates still fail. The complete source/Office token and simple
math screen now passes all ten outcomes: 4,234 source prose tokens, no token/math
deficits or unresolved source features. This does not establish image/ink/clipping
or reading-order fidelity. Seventeen assets/four smoke PDFs transfer without root
rebuild or suite replay. C1 remains open for geometry and composed-content limits.


Autofit column measurement now consumes prepared equation advances. Its minimum
protects each whole box; its maximum reuses the ordered inline measurement used
for wrapping, including prose, pictures, formulas and authored breaks. Fixed,
preferred-width and span fallbacks retain their existing paths. Cancellation flows
through the new measurement work; no font/size/document-specific rule is added.

An independent four-row control exposed a 10.7pt equation column and four omitted
fractions. All four now paint; added-prose and duplicated-formula transformations
also preserve their four/eight expressions and increase preferred width. Native
foreground recall improves from 0.275 to 0.851; other visual gates still fail. An
unreferenced template header continues to report an ignored equation, separately
from the visible body loss fixed here. Fixed-grid spacing controls show ordinary
full-row heights are close to Word; the held-out page-boundary discrepancy remains.

Thirty-two selected methods pass in one invocation and eleven portable checks
pass on Ubuntu. All ten portfolio PDFs remain exact and their content audit has
no deficits. Twelve independent comparisons reuse verified identities. Seventeen
assets/four smoke PDFs transfer without root rebuild or suite replay. The next
rendering investigation is independently controlled trailing-spacing admission
at a page boundary; the held-out memo remains an acceptance check.

Natural mathematical last rows now distinguish measured content from trailing
paragraph after-spacing at a page boundary. If all content fits atomically,
placement may discard only the overflowing unpainted tail. Existing measurements
are reused. Authored heights, keeps, merges, positioned content and shifted runs
retain their fallbacks; no font, size or whole-document qualification is added.

Independent 0/6/24pt spacing controls keep the equation on page one in Word and
the candidate. CantSplit and insufficient-space controls keep it on page two in
both modes. The held-out memo's Final equation row now stays on page one. Its
four source cell clips match, with no unresolved cells and a maximum bounds delta
of 1.333pt; the raster also shows the complete fraction inside the cell.

Forty-one selected methods pass in one invocation; twelve pinned-font rendering
methods pass on Ubuntu. Nine portfolio PDFs retain bytes, and fifteen independent
comparisons reuse verified identities. A page-local token screen detects the old
memo displacement despite matching page totals and passes all ten new outcomes.
It supplements global prose/simple-math presence checks, without proving ink,
image, clipping or reading-order fidelity. All eight DOCX visual gates still fail;
both decks pass. Memo Final page-two MAE improves from 1.172 to 0.429 and recall
from 0.541 to 0.840. Seventeen assets/four exact smoke PDFs transfer without root
rebuild or suite replay. C1 remains open for significant geometry and composition
limits.

Ordinary list numbering no longer excludes supported composed equations. Label
measurement and text-state planning are shared with ordinary paragraph layout;
the first measured line owns its label segments, while containers place those
segments in their own frames. Authored indentation/suffix rules are reused.
Numbered rows retain inline style and existing unsupported rich-content fallbacks.

A numbering-only transformation previously removed four equations from three
mixed paragraphs in both modes, while preserving prose. The candidate retains
four equations and three labels across body/cell/header/footer controls. Native
unchanged-format run splitting and pinned-font source transformations also retain
content. The body control's recall improves 0.141→0.671 and MAE 1.464→0.821;
all eight independent numbered-story visual gates still fail. Raster review shows
the labels, fractions, literals and illustration together.

Fifty-two selected methods pass in one invocation, including numbering emission,
indentation/suffix/continuation and table-math neighbors. Thirteen portable methods
pass on Ubuntu. All ten portfolio PDFs retain bytes; eighteen independent Office
comparisons reuse verified identities. Global/page-local prose and simple-math
screens pass. Seventeen assets/four exact smoke PDFs transfer without rebuilding
or repeating the suite. The portfolio DOCX gates remain failed; decks remain passed.

Geometry triage now preserves PDF page ownership in both nearest-position and
indexed text comparisons. Page-moved content previously passed both modes. A
second adversarial control exposed rounding-edge splits of collinear runs; anchored
0.1pt baseline bands now keep those runs in X order without chaining distinct
lines. Twenty-two adversarial cases pass. Optional structured reports retain input,
tool/helper hashes, matching options, tolerances, pages, text and measured deltas.
Existing raster thresholds are unchanged.

Cached portfolio inspection required no conversion or library build. All 334
Final merged pairs match normalized text and page, with maximum start differences
of 0.877pt X and 0.981pt Y. Review outcomes retain 28/33 unresolved operation pairs
per document, including differing mathematical operation segmentation; these are
not inferred content omissions. Exact review pairs reach 3.152pt Y. Global and
page-local content screens remain separate evidence; neither proves complete ink
or clipping fidelity. Sub-point Final refinements remain lower priority than
confirmed composition loss and the overfull operator-breaking contract.

Top-level inline literals now expose binary/relation break opportunities from
their existing measured glyph order. The ordered paragraph path wraps those
pieces, drops inter-piece glue at a line boundary and preserves source equation
identity. Structured boxes, display rows, nested and unary-only literals retain
their qualified fallbacks. No font/size/document-shape rule is added.

Word 16 keeps `α+` on the first line and `β=γ` on the next for default, before,
after and repeat settings in the independent inline control. This observed inline
behavior differs from the general default described in [Microsoft's brkBin
documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.breakbinary?view=openxml-2.10.0);
the implementation claim is limited to the validated inline scope. All five
glyphs now paint in a two-line source paragraph, including with bookmarks and run
splitting. The portable regression fails on the previous library.

Forty selected methods pass in one invocation; fourteen portable methods pass on
Ubuntu. The overfull control's recall improves 0.375→0.480 and MAE 0.290→0.280.
All eight setting/mode controls still fail foreground/MAE/SSIM gates. Its body
auto line height remains 16.968pt versus Word's 14.64pt; font/measurement ownership
needs separate evidence before changing it. Eight fitting literal/fraction/script
outcomes retain their comparison results. All ten portfolio rasters retain bytes
at the existing 144dpi qualification, while eight DOCX PDF byte streams change
with mathematical boxing. Content/page-local screens pass, but DOCX visual gates
remain failed. Seventeen assets/four exact smoke PDFs transfer without root build
or suite replay.

An ink-only inline-frame prototype worsened review portfolio MAE by up to 1.904
and recall by 0.0434. That measurement rule was rejected and reverted before
delivery; the accepted operator change retains the prior frames. The rejected
artifact and comparisons remain available as negative evidence.

Broken-literal content is now independently exercised in a cell, header and
footer, in both modes. All five glyphs and following prose survive; native
bookmarks/run splitting and pinned-font placement regressions pass. The production
library and its source tree remain identical to the accepted operator artifact:
only two new methods run in one invocation, plus the new Ubuntu method. The
earlier forty-method/fourteen-Ubuntu qualification and portfolio outputs are
reused by identity. Seventeen assets/four exact smoke PDFs transfer without root
rebuild, suite replay or repeated portfolio conversions.

The six new Office controls still fail their original visual gates. Final-mode
second-line differences are 2.334pt in the cell and 2.452pt in the header; the
footer's first line is 7.322pt high. Reader fallback currently supplies 278/240
when no line spacing resolves, while these sources use a single-line pitch in
Word. That discrepancy is confirmed; paragraph-mark metric ownership is a
hypothesis to test, not an accepted cause. Preserve older independent unstyled
controls while isolating document defaults, style inheritance and mark formatting.

Document paragraph defaults now retain their source ownership even when `pPr` is
empty. With no resolved line token, a declared document default supplies single
spacing; an absent owner retains implicit Normal 278/240. A uniquely declared
default paragraph style binds through the existing cascade, so its explicit line
setting overrides the fallback. Authored style IDs and cascade provenance remain
separate from inherited formatting. Two identical fallback constants and a
redundant qualification helper are removed. No font/size/document-shape guard is
added or indiscriminately removed.

The seven-source ownership matrix separates absent, run-only, empty/nonempty
paragraph defaults and explicit Normal inheritance. Its fourteen mode outcomes
preserve explicit Normal PDF bytes while affected recall improves 0.222→0.773 and
MAE 9.173→3.981. The source controls still fail their original visual gates. Plain
cell/header/footer controls improve recall from 0.377/0.286/0.353 to
0.836/0.850/0.831; MAE and changed-pixel gates pass, with foreground/SSIM failures
remaining. These controls disprove a global fallback replacement and do not
establish paragraph-mark metrics as the cause.

This shared-reader milestone qualifies 1,494 DOCX/neighbor methods once in one
invocation and sixteen portable methods on Ubuntu. An earlier artifact passed
1,493/1,494 but changed authored style identity; that defect was corrected before
delivery and the failed artifact/report retained. All ten portfolio PDFs and their
comparisons remain exact; full/global/page-local content screens pass. Seventeen
assets/four exact smoke PDFs transfer without a root build or suite replay. The
broader selection is a deliberate shared-contract milestone, not the selection
policy for routine local changes. Existing limits, cancellation, safe managed code
and visual thresholds remain intact.

PDF text comparisons now transform advances, character spacing and font height
consistently with effective coordinates. Page-local exact decoded matching reserves
content before geometric fallback, and decoded comparison is case-sensitive.
Twenty-eight adversarial controls pass, including scaled/sheared/rotated text,
adjacent columns and case changes. Existing tolerances and raster gates remain.
Cached inspection pairs all 676 normalized portfolio operations with no conversion
or library build: Final maxima are 0.877pt X / 0.981pt Y; review maxima are
0.484pt X / 3.166pt Y. These supersede the earlier unresolved segmentation counts.
Raw whitespace differences remain explicit: review output loses an authored space
after a comment reference in all four documents. This confirmed composition defect
is the next bounded correction; normalized matching does not establish exact text
or visible ink fidelity.

Authored whitespace now survives comment composition. The old review-only filter
suppressed every whitespace-only emission part in any commented paragraph; it
retained positioned gaps but damaged PDF text extraction. No generated comment
placeholder exists in this path. Removing the filter and its plumbing preserves
single/multiple/nonbreaking separators, same-run references, a following footnote
and an explicit break. Font cmap aliases may extract NBSP as ordinary space.

Twenty-two selected methods pass once in one invocation, and two affected portable
methods pass on Ubuntu. Six portfolio PDFs retain bytes; all seven pages of the
four changed review outcomes retain raster bytes at 144dpi. The two independent
Office source controls retain their visual metrics. All ten content screens pass;
eight DOCX numerical gates still fail and both decks pass. Seventeen assets/four
exact smoke PDFs transfer without root build or suite replay.

Freezing now rejects differing library copies in the library, tests and CLI output
directories. A test-only rebuild exposed a stale CLI copy; the coherent snapshot
retains the already-passed library/test hashes, so qualification is reused. Its
small portfolio outputs were rechecked once and all twelve comparisons reused by
PDF/reference/threshold identity. This extra check resolves the concrete coherence
failure rather than adding routine qualification work.

Printed inline mathematics now separates baseline placement from its qualified
occupied frame. The ordinary body path compared scaled mathematical extents with
an unscaled local text inset; the ordered path omitted the first-body baseline
anchor transition used by ordinary text. Both now place the local baseline in the
same physical scale and retain the source frame's minimum occupied height. Label
leading, glyph drawing and container fallbacks remain. No font/size/document-shape
qualification is added. Earlier experiments that also reduced occupied height
regressed three review documents and were rejected.

Ten source controls (literal/fraction/root, 12/18-point leading anchors, single or
multiple formulas) have twenty independent Word mode references. Mathematical
baseline gaps pass the 0.4pt placement bound; all twenty preserve simple math text.
A pinned-font transformation test proves that adding a fitting second formula
preserves its baseline and following advance. It fails on the previous library
with a 1.360pt review displacement. Twenty-three selected methods pass once in one
invocation, plus two affected Ubuntu methods.

Six portfolio PDFs retain bytes. Each review document's page-one MAE improves by
about 0.348; all 676 normalized page-local text pairs resolve, and maximum review
start Y difference falls from 3.166pt to 2.223pt. All ten source/Office content
screens pass. The numbered body neighbor keeps its visual metrics. All eight
DOCX numerical gates still fail; both decks pass. Seventeen assets/four exact
smoke PDFs transfer without root build or suite replay. Cached comparisons reuse
checked PDF/reference/threshold identities.

A broader following-line assertion revealed an existing literal automatic advance
discrepancy (2.219pt in one review control). It is retained as an unresolved
measurement issue; the accepted baseline correction preserves occupied height.
Current accepted-build checks also supersede older broken-literal story residuals:
Final cell/header/footer prose baselines differ from Word by at most 0.364pt after
the reader correction. These findings do not establish complete ink/clipping or
reading-order fidelity.

Fitting inline math/text rows now use the ordinary text frame when their
mathematical bounds fit above and below its physical baseline. They no longer
retain a second unscaled font frame. Taller/display math and image-containing
rows preserve their qualified occupied extents; no font/size/document-shape
qualification is added. The row retains its chosen text baseline separately from
object bounds and label leading.

The independent Word regression fails on the prior accepted library by 2.219pt.
Across twenty mode controls, the largest following-paragraph gap error is now
0.199pt (previous literal cases reached 2.297pt). All preserve their mathematical
text. Literal review MAE improves about 0.258–0.261; SSIM failures remain. The native
assertion now checks both the mathematical baseline and following advance. A
pinned-font ordinary-text substitution exercises the same line-box contract on
Windows/Ubuntu; it already passed with the prior library, so it is integration
coverage rather than a claimed reproduction of the native-font defect.

Twenty-eight selected methods pass once in one invocation, plus three Ubuntu
methods. All ten portfolio PDFs retain bytes and independent comparisons; global
and page-local content screens pass. All eight portfolio DOCX numerical gates
still fail and both decks pass. Seventeen assets/four exact smoke PDFs transfer
without root build or suite replay. The compound-document allocation baseline
was separately refreshed at the preceding 0fa503ac artifact; it does not establish
a repeated-layout cost or a speedup. Superseded compiled snapshots were archived
and hash-verified, retiring another 133.8MB while preserving inputs and evidence.

Isolated cell equations now scale the transition from the nominal cell baseline,
sharing the ordinary text rule. Models without a recorded nominal cell anchor retain the source-size fallback;
parsed empty runs normally carry the paragraph font metrics. Static stories with already physical
tops retain their mapping. No font/size/document-shape qualification is added.

Six new source controls cover literal/fraction/root content, alone in a cell or
between text paragraphs, in both presets. Review literal/fraction glyph-baseline
errors fall from about 2.6pt to at most 0.137pt; root/fraction retains about 0.59pt.
Final PDFs retain bytes, and all six review comparisons improve. Four independent
native baseline controls pass; the prior library fails by 2.682pt. Twenty-three
selected methods pass once in one invocation and three portable methods pass on
Ubuntu. The portable cell controls use normal-text fractions; the pinned face does
not cover the original mathematical italic Greek fixture.

Six portfolio PDFs retain bytes; all review comparisons improve and fractional
baseline differences are at most 0.704pt. All ten content screens and normalized
review pairs resolve. Eight DOCX numerical gates still fail and both decks pass.
Seventeen assets/four exact smoke PDFs transfer without root qualification build
or suite replay. One accidental root prototype build was corrected by restoring
all seventeen qualified assets by hash. The record distinguishes that incident
from the frozen qualification.

Visual inspection also confirms an existing composition defect in the new
zero-after-spacing controls: a table top border crosses preceding descenders,
where Word leaves a gap. The table frame's eligibility and coordinate ownership
remain the next bounded task. The baseline correction improves the equation and
preserves frame geometry; it does not claim to resolve that overlap or establish
complete clipping fidelity. Source-backed geometry, independent references and
original numerical failures remain retained.

Natural multi-paragraph cells now retain the same printed table origin as
single-paragraph cells. The old one-paragraph guard disabled the shared geometry
projection, making borders cross preceding descenders. Every paragraph must still
use natural line spacing; fixed paragraph spacing and the existing richer-content,
authored-row, alignment and fragment fallbacks remain. Page-admission eligibility
is separate and unchanged. Empty resolved runs already carry equation-only font
metrics; tracing rejected the earlier missing-owner hypothesis.

Independent controls place the table after plain text, fitting mixed mathematics
and an image, in both presets. Top-border error falls from about 2.6pt to at most
0.223pt. A portable transformation proves that adding cell paragraphs increases
occupied height without shifting the table top. Both it and the native Word gap
check fail on the previous library. Fourteen selected methods pass in one
invocation, plus three Ubuntu methods. Ten affected control outcomes retain their
mathematical text; original visual failures remain explicit.

The fixed portfolio contains only one-paragraph cells, where the predicate remains
equivalent. Its prior complete qualification is retained, with no full-portfolio
reconversion or new full-corpus claim. Letter and held-out memo review smoke PDFs
retain bytes. Seventeen assets/four exact integration PDFs transfer without root
build or suite replay. The generic measurement tool now calls custom-manifest
results pinned document/mode outcomes, distinguishing them from full portfolio
qualification. The remaining C1 task is the finite document severity and visible
clipping/overlap review, retaining independent thresholds and unresolved evidence.


The finite C1 portfolio severity review is complete. All sixteen candidate/Office
page pairs were inspected, using comparison-chain report hashes and exact PDF
identities. No missing visible content or serious pagination, overlap or clipping
defect was observed in these documents. Global and page-local prose/simple-math
screens remain clear, and both source PNG colors retain visible patches in all
eight DOCX outcomes. This is bounded visual evidence, not a proof of arbitrary
reading order, glyph ink bounds or unsupported combinations.

| Document | Final / review pages | Content / serious placement | Residual outcome |
|---|---|---|---|
| Report | 2 / 2 | No omission or serious defect observed | Typography/table spacing; review pane and connector geometry |
| Letter | 1 / 1 | No omission or serious defect observed | Typography/table spacing; review pane and connector geometry |
| Specification | 2 / 2 | No omission or serious defect observed | Typography/table spacing; review pane and connector geometry |
| Held-out technical memo | 2 / 2 | No omission or serious defect observed | Table/baseline spacing; review pane and connector geometry |
| Briefing deck | 1 | No omission or serious defect observed | Numerical gate passes |
| Held-out deck | 1 | No omission or serious defect observed | Numerical gate passes |

All eight DOCX numerical visual gates remain failed under their original
thresholds; neither manual review nor the bounded reliability checkpoint relabels
them. Review pane background extent is visibly different even where content is
clear. The held-out documents remain acceptance checks, outside derivation.
Detailed per-outcome severity, page raster hashes and comparison provenance are
retained in `artifacts/composition-c1/c1-portfolio-severity.json`. The source PNG
patch refresh is `portfolio-image-patches-current-frame.json`. Latest production
2ad7208a retains the 07b20914 full portfolio through the documented one-paragraph
predicate equivalence and exact integration smokes; no corpus reconversion was
needed for this review. CI 38820298013 (root 66d66fc1) is green on both platforms.

C1's bounded mixed-content reliability checkpoint is accepted with these explicit
visual residuals. Wider rich/related/floating stories and OMML remain on the
roadmap. The next maintenance step reuses unchanged immutable fixture snapshots;
subsequent rendering work returns to supported-content composition rather than
fitting the held-out portfolio or lowering visual thresholds.


Qualification freezes now share a read-only public fixture snapshot keyed by its
Git Cases tree across artifacts and linked worktrees. Every freeze still checks
all source/snapshot SHA values and file counts, then links its runner's Cases
directory to that snapshot. Changed input trees remain separate. Seven focused
checks pass on Windows and Ubuntu, including dirty inputs, corrupt snapshot and
mixed-library rejection, and linked-worktree reuse. CI runs the same small tool
checks; no library suite is added.

Applied artifact `frozen-shared-inputs` verifies 1,227 fixtures and preserves the
previously qualified 2ad7208a library bytes. One runner integration method and one
CLI conversion through the fixture link pass, the latter reproducing the accepted
PDF exactly. No library rebuild or qualification replay was needed. Evidence:
`artifacts/composition-c1/shared-fixture-evidence.json`. Historical fixture copies
and the previously rejected retirement targets remain untouched.


Ordinary default/left tabs now coexist with supported equations through the
ordered measurement path. The existing shared tab-stop resolver consumes preceding
text/math/image advances; body, cells, direct headers/footers, keep estimates and
autofit receive the document's default interval. Tab positions scale once. Right,
center, decimal, bar/leader and out-of-frame stop semantics keep explicit fallbacks.
Visible tabs prevent sole-formula display promotion. No font/size/document-ID
qualification is added.

Eight public source controls exercise an 18pt default interval and authored 72pt
left stops in four containers, separately in both presets. The prior library
omits both formulas in all sixteen outcomes; native and portable regressions
reproduce it. The accepted build retains both formulas, prose and one-page
placement. Independent Word relative field offsets differ by at most 0.052pt.
Review cell baselines remain about 2.5pt high; reviewed body/cell/header/footer
crops show visible unclipped formula content. All sixteen original numerical
control gates remain failed; this change establishes composition, not exact
typography.

Frozen 32f59eae / root b42436cb passes 23 selected methods in one invocation and
three Ubuntu methods. Source ordering, wrapping, scaled stops, cancellation and
explicit fallback controls pass. A read-only fixture workflow issue initially
blocked seven mutation neighbors: the helper copied the source attribute onto its
temporary DOCX. Only the temporary copy is now writable; the final artifact reuses
the verified 1,237-file shared snapshot. Failed preliminary evidence is retained.
All sixteen accepted control PDFs match the prototype and reuse pinned Office
comparisons. The portfolio has no authored inline tabs, so its complete previous
qualification is retained; letter and held-out memo review smokes retain bytes.
Seventeen assets/four integration PDFs transfer without a root rebuild or suite
replay. Evidence: `artifacts/composition-c1/math-left-tabs-evidence.json`.


A follow-up footprint check found 6.089GB, principally historical runtime
archives. Their obsolete compiled blobs and 86 Git-export duplicates were
retired after every exported file was checked against reachable Git history
(recording checkout CRLF normalization for text). Four exports lacking reachable
commits were retained. Compact archives retain original path maps and explicit
retirement metadata; all remaining blobs were SHA-verified. Current/rollback
runtimes, source inputs, PDF/raster evidence, uncommitted patches and Git refs
remain. No previously rejected retirement target was retried. The workspace is
4.606GB / 257,880 files with no scan errors; proof and compaction maps are under
`artifacts/workspace-maintenance/`, and the complete check is
`artifacts/workspace-footprint-after-archive-compaction.json`.


Equation/comment composition now keeps ordered source boundaries and measured
visible extents. Comment ranges attach as metadata; prose, formula-only, between-
formula and mixed selections preserve equations in body paragraphs and cells.
Highlights and connectors use placed source geometry, including the formula's
visible content rather than its padded layout advance. Adjacent prose owns the
inter-object whitespace. The line owns selection height independently of range
endpoints. Legacy annotation geometry remains on its existing path. Note/field/
revision and broader story combinations retain staged fallbacks.

Eight public sources have independent Word exports in both presets. Twelve
initial parent outcomes omitted both equations; native and portable regressions
fail on that prior library. The accepted base retains mathematical/prose content
in sixteen outcomes and passes eight independent selection-width checks, with
maximum horizontal error 0.080pt. Bookmarks and unchanged-format run splitting
retain source ranges. The sole-formula policy also retains display ownership.
Eight review crops show visible equations, prose and comment captions; cell
baselines, selection height, connectors and balloon/pane cosmetics remain
approximate. All sixteen original numerical visual gates remain failed.

Base artifact 535a5d8e / root c5130a51 passes fifty selected methods in one
invocation and four Ubuntu methods. Follow-up 6ff2f99e / root 9478690f corrects
line-owned highlight height; its regression fails on the base artifact. Eight
affected methods pass once in one invocation, plus one Ubuntu method. Only the
four prose/between source documents are reconverted for that follow-up (eight
mode outcomes); all four Final PDFs and their comparisons retain identity. Other
base qualification remains retained, rather than reporting a new complete run.
The 1,247-input snapshot is reused. Seventeen assets/four exact PDF smokes transfer
without a root rebuild or suite replay.

The fixed portfolio has no paragraphs combining comments and equations; retain
its prior full qualification and severity review, with exact letter and held-out
memo review integration PDFs. The held-out document remains outside derivation.
Evidence: `artifacts/composition-c1/math-comments-evidence.json`.


Continuous automatic footnote references now retain explicit resolved marker
owners in the ordered paragraph source. Equations and note glyphs coexist before
or after each other in body paragraphs and cells, with or without comment parts.
Shared authored runs and starts 1/9/12/103 retain ownership. Custom/unowned labels,
page/section/unknown restart requests and equation/endnote combinations retain
staged fallback. Plain-text correction eligibility now excludes equations, a
semantic ownership rule instead of another font/size/document-shape exception.
The prior predicate demonstrably accepted equation paragraphs; no separate
visible regression from that misclassification is asserted.

Ordered placement also scales source vertical run offsets once, preserving
superscript marker raises in physical review frames. Sixteen independent Word
mode controls retain both formulas, note text and destinations; relative marker
raises differ by at most 0.040pt. An earlier prototype fails this assertion by
about one point. Review cell baselines still differ by about 2.5pt, and all sixteen
original numerical control gates remain failed. Four reviewed review crops retain
visible equations and note content. The previous library omits both equations in
all sixteen mode probes; native/portable content regressions fail there.

Base 951e1370 / root eb790419 passes 54 selected methods in one invocation and
five Ubuntu methods. A rendered shared-run control exposed an additional link
ownership defect: an authored digit identical to the automatic label acquired a
second link. Follow-up 5ec7ac65 / root 41f224c8 retains the marker role through
formatting/emission and links only that owner. Its regression fails on the base;
14 focused methods and two Ubuntu methods pass. The controls cover paragraphs
with and without equations, plus existing page numbering, custom marks and
endnote behavior.

The base sixteen source/mode outputs and Office evidence are reused: canonical
reference runs contain no colliding authored digit. No full control/corpus replay
was needed for the role follow-up. Seventeen assets/four exact integration PDFs
transfer, including two canonical footnote controls and the portfolio letter/
held-out memo. The 1,258-input snapshot is reused. No root rebuild or suite replay
is performed. Portfolio paragraphs do not combine equations and footnote
references, and their tables/static stories/comments already exclude plain-text
corrections; prior complete qualification remains. Evidence:
`artifacts/composition-c1/math-footnotes-evidence.json`.

C2 adds a separately pinned, held-out composed annotation report. The existing
qualified library preserves all prose and equations in both DOCX presets, but
moves a complete review section to page 2 where Word retains it on page 1.
The checkpoint remains open. Independent small note-admission controls must
reproduce the cause before changing layout; the acceptance report stays outside
derivation. Reuse committed supplemental manifests with
`AuditCompositionContent.py --input-root . --manifest <public-manifest>` and the
existing frozen runtime. Input Git identity, source/PDF hashes and public-only
boundaries remain checked. Legacy ten-outcome compatibility and three negative
input controls pass without new conversions or library builds. Evidence is in
`artifacts/composition-c1/supplemental-audit-tool-evidence.json` and
`composed-report-content-audit.json`. CI through `be411f37` is green on both OSes.

An independent atomic-footnote reduction confirms unused future continuation
reserve can cause a spurious second page. `d2c5dde6` reserves only the complete
head for a new one-line note; split notes and continued tails keep their existing
reserve. Thirty-nine affected methods and five Ubuntu checks pass. The source-only
control-width follow-up `99991423` reuses all build assets and repeats only two
fixture-dependent methods (one on Ubuntu); all three new regressions fail with
the parent library. Four Word controls now match one-page membership and retain
all content, while their numerical gates still fail (note/separator Y about 6pt).
All ten original portfolio PDFs are exact, so Office comparisons and page reviews
are reused. Transfer verifies 17 assets and four PDFs, without root build/replay.
C2 Final now matches Word's page-local prose and its two page pairs show no serious
placement defect. Review remains unchanged and still moves section 7 to page 2;
both modes retain all equations/prose, but both numerical gates fail. C2 stays
open for an independent scaled-admission reduction. Evidence:
`artifacts/composition-c1/atomic-note-evidence.json` and
`composed-report-atomic-note-content.json`. Workspace is 4.941GB at the latest scan.

C2 now passes prose, simple-math and page-local content screens in both presets;
all four held-out page pairs are reviewed. Table-fraction/border clearance,
annotation geometry and typography remain residuals; both numerical gates fail.
`c433142b` converts note reserve once to printed paragraph points and uses shared
resolved script sizes for existing exact-line eligibility. Four valid standard-page
Word controls retain eighteen body lines and their note on one page; the parent
fails all three new regressions. Custom short-sheet review references were rejected
for out-of-page Word ink. The final frozen identity qualifies 44 affected methods
in one invocation and five Ubuntu methods. A test-only revision-metadata rebuild
required that final qualification; retain the qualified DLL for future test-only
builds with `--no-dependencies`. Four original review portfolio PDFs remain exact.
Transfer/remaining PDF identity smokes use no root build or suite replay.
`39d7beab` repairs the sole Windows CI failure at `fcc0ad96`: its synthetic column
test now retains the two-column frame and verifies all four note destinations on
the admitted page. Ubuntu was green; the next push is under CI. Evidence:
`artifacts/composition-c1/scaled-note-evidence.json`. C2 is accepted for bounded
composition with the numerical/geometry residuals; release preparation stays deferred.

`256ca0a4` corrects the shared review table-fraction clearance. Ordered cell frames
measure printed text struts and receive the row-origin translation without a
second nominal-inset conversion. Eight independent controls vary note-marker size
11/12 and script alignment; maximum review border-relative error falls from
2.652pt to 0.149pt. Four Final PDFs are exact. Twenty-eight focused methods in
one invocation and five Ubuntu checks pass; the parent fails both new regressions.
C2 retains all source/page-local content; its Final PDF and review page-2 raster
are exact, while the corrected review page 1 is reviewed. Both numerical C2 gates
and all eight control gates remain failed. Eight Office prose-token screens remain
unresolved because normal-text math words use Office's prose font; source prose,
simple math, direct glyph geometry and reviewed crops retain the content. The
fixed portfolio excludes changed ordered cells; prior Office reviews plus exact
letter/memo smokes remain accepted without full-corpus replay. Transfer checks
17 assets/four PDFs without root build or suite replay. Evidence:
`artifacts/composition-c1/ordered-cell-frame-evidence.json`. CI through `ed9566a9`
is green on Windows and Ubuntu. Workspace was 5.069GB at the last scan; automatic
review rejected retirement of 48 verified compiled duplicates (blocked by policy),
so no files were removed and those targets will not be retried.

`17b0e6cb` bounds future fixture growth by sharing verified read-only files between
published input trees. Checkout files remain independent; corruption is rejected
both when reusing a tree and when seeding a new one. Ten tool checks pass on
Windows and Ubuntu. A real snapshot shares 1,271 files / 12.09MB, copies 7.2KB of
byte-distinct metadata and rechecks all 1,274 input hashes after a temporary-OOXML
mutation test. All 17 compiled assets and one CLI PDF are exact, without library
build or suite replay. The footprint command retains its PowerShell interface and
uses Python standard-library file IDs to count hard links once; both-OS tests also
reject directory loops. Latest size: 5.159GB unique content / 5.171GB logical file
lengths, near the loose cap. Historical/denied retirement targets are untouched.
Evidence: `artifacts/composition-c1/cross-tree-inputs-evidence.json`. CI through
`3b7b8865` is green on both OSes; resume annotation geometry with existing references.

`1765e222` maps review-pane vertical bounds through the shared printed paper
frame, independently of body margins, and retires two fitted inset constants.
The same Y mapping preserves floating-textbox emission. Nine independent Word
controls cover six margin pairs, Letter/A4/landscape and mirrored pages: maximum
error is 0.438pt, down from a 377.571pt body-margin error. Twenty-six affected
methods in one invocation and five Ubuntu checks pass; both parent regressions
fail. The shared paint change justifies one fixed-corpus run: all ten content/
page-local screens pass; text operations and non-pane graphics remain exact;
seven review pages change only gray paper pixels, and six Final/deck PDFs are
exact. Eight DOCX gates still fail and two decks pass. C2 content/pagination remain
qualified in both presets, with both numerical gates failed and connector/range
geometry still residual. All nine new control gates remain failed. Transfer uses
17 checked assets/four PDFs, without root build or suite replay. `-DocxPreset Review`
selects only relevant DOCX outcomes for subsequent control batches. Evidence:
`artifacts/composition-c1/review-paper-frame-evidence.json`. CI at `c71ef4c8`
passed Ubuntu and Windows rendering but exposed a Windows long hard-link path;
`367656f1` uses extended paths, passing ten tool checks on both OSes and a
277-character shared-input path without renderer rebuild/requalification.

`3b79ab77` makes comment selections consume the occupied inline frame, including
its baseline strut, with visible bounds allowed to enlarge it. Eight independent
Word controls covering body/cell and prose/formula/mixed ranges reduce the height
deficit to at most 0.119pt; baseline-relative top/bottom error stays below 0.187pt.
Nineteen capability methods pass in one invocation, four portable Ubuntu checks
pass, and both new regressions fail with the previous library. All eight control
text-operation files remain exact; four paired crops are reviewed. No font/size/
document eligibility branch is added. All eight numerical control gates still
fail. C2 content and page-local screens pass in both presets; Final PDF and both
text-operation files are exact, with only five review selection operations changed.
The image/math selection crop is reviewed; both numerical C2 gates remain failed.
Connector and balloon attachment remain explicit residuals. Eight older Office
prose screens for normal-text equation words remain unresolved. Frozen `0d02535c`
qualifies once; root verifies 17 assets/four PDFs without rebuilding or replaying
the suite. Evidence: `artifacts/composition-c1/comment-occupied-frame-evidence.json`.
CI through `881f8871` is green on Windows and Ubuntu.

`9a14f5ac` shares measured selection depth with comment connectors while retaining
independent balloon placement. Eight Word controls reduce baseline-relative
connector error from 2.22pt to at most 0.112pt. Twenty-five methods pass in one
invocation and four portable Ubuntu checks pass; both new connector regressions
fail with the preceding library. All eight text-operation files, including balloon
text, stay exact; only two connector strokes change per control. A first prototype
coupled the two coordinates and displaced balloon titles; an independent title
guard demonstrates the rejected regression. Only corrected frozen `8a9db635`
transfers to root. Four control pairs are inspected; all eight numerical gates
still fail. C2 passes both content/page-local screens, retains exact Final PDF and
both text-operation files, and changes only two review connector strokes. Both
numerical C2 gates remain failed. Connector X/dash, balloon baseline and typography
remain approximate. Transfer checks 17 assets/four PDFs without root rebuild or
suite replay. Evidence: `artifacts/composition-c1/comment-connector-placement-evidence.json`.
CI through `6400abec` is green on both OSes.

`1ba76589` retains supported equations beside materialized cached fields. Shared
semantic source ownership replaces the blanket field exclusion; complex field
results packed with their control markers now enter the ordered source view even
after a formula. Eight controls / sixteen Word comparisons cover simple/packed
REF results in body, cells, headers and footers. Both fractions and source prose
remain in all outcomes, field advances agree within 0.144pt, and printed baselines
within 0.375pt. All eight simple/packed PDF pairs are exact. Twenty-nine methods
pass in one invocation and five portable Ubuntu checks pass; four new regressions
fail with the preceding library. Four paired crops are reviewed. All sixteen
numerical visual gates remain failed. The generic audit also retains sixteen
unresolved field-source/math-token classifications, despite zero source-prose and
simple-math deficits; these are not relabeled as passes. PAGE/NUMPAGES, unresolved
fields and revision owners retain their fallback. This adds cached-result display,
not field evaluation. Frozen `4609b4bb` transfers with 17 assets/four exact PDFs,
including accepted C2 review and held-out memo review, without root rebuild or
suite/corpus replay. The reference command adds an optional DOCX preset selector;
verified cached review references are reused without export. Evidence:
`artifacts/composition-c1/math-fields-evidence.json`. CI through `4ed093b6` is green
on both OSes. The latest workspace scan reports 5.365GB with four long-path fixture
traversal errors; those errors remain explicit pending the bounded tool correction.

The workspace counter now uses extended Windows paths for traversal while retaining
its ordinary reported root. A regression reproduces the missing long-path byte;
Windows wrapper and direct Ubuntu Python checks retain unique/logical byte, hard
link and loop behavior. The complete scan reports 5.379GB with zero errors. The
Ubuntu PowerShell wrapper remains a CI check; local Ubuntu has no PowerShell.
No library build or rendering suite is replayed. Evidence:
`artifacts/composition-c1/footprint-longpath-evidence.json`.

The bounded C3 cached-field checkpoint reuses frozen `4609b4bb` for a mechanical
held-out report variation: visible prose, six equations, one image and all package
parts except main XML remain unchanged. Simple/packed cached results coexist with
math, the image/comment range and an internal link. Both presets retain two pages
and pass all source/Office prose, simple-math and page-local screens. Existing four
link targets survive; review link end differs by 0.114pt. Merged text geometry has
zero failures (88 Final / 90 review lines), and graphics operations are exact
against accepted C2. Candidate and Office second-page rasters are exact; changed
first-page regions are inspected. Both numerical gates remain failed. Source
commit `43339a3d` and tool `b583a8d7` pin the independent input/audit work; all 17
runtime assets are unchanged, with no build or rendering-suite/corpus replay.
Twelve source-audit adversarial controls pass on Windows/Ubuntu; cached fields
resolve their old classification failures while dynamic/malformed/cross-paragraph
fields remain unresolved. Ten C1 audit screens still pass. Sixteen fraction probes
retain rate/time token deficits and all their visual failures. Initial acceptance
phrases were corrected to the actual baseline note/text; document bytes and both
candidate PDFs remain exact, without Office re-export. Evidence:
`artifacts/composition-c1/composed-fields-evidence.json`. CI through `373c99ad` is
green on both OSes; workspace traversal now reports 5.379GB with zero errors.

`934c79a1` supports PAGE/NUMPAGES beside equations in page-owned direct headers and
footers. Four Word controls / eight outcomes cover twelve pages and the 9/10 digit
boundary; every one of the 96 page outcomes retains the actual page values and
both fractions. Forty-seven methods pass in one invocation, six Ubuntu checks pass,
and two rendering regressions fail with the preceding library. Baseline and leading
field error stay within 0.375/0.120pt. Word's simple-field digits use Calibri while
cached source formatting retains Arial: count-field advance residual is 1.115pt,
and cumulative end offset reaches 2.396pt. Four pairs are reviewed; all eight control
gates remain failed. Complex-field cache ownership now ends at its end marker,
retaining following authored spaces/numbers; simple placeholders publish source
items. Shared substitution replaces three copies. Body/cell and related-story
dynamic math retain fallback; source placeholders cannot qualify without actual
page inputs. Ten fixed portfolio PDFs are exact, reusing Office comparisons and
ten passing content/page-local screens; eight DOCX gates still fail and decks pass.
Frozen `89d9d93a` transfers 17 assets/four exact PDF smokes, including accepted C3
review and held-out memo, with no root build or suite replay. The generic prose
audit deliberately leaves dynamic source fields unresolved; independent per-page
checks qualify this slice. Evidence: `artifacts/composition-c1/static-page-fields-evidence.json`.
CI through `cce36176` is green on Windows and Ubuntu.

`72ae0ebd` shares owned automatic-reference rules across footnotes and endnotes,
retaining restarted/custom/unowned fallbacks. Four Word controls / eight outcomes
keep both fractions, note text and the destination in body/cell, before/after
combinations. Twenty-nine methods pass in one invocation and five Ubuntu checks
pass; three regressions fail with the preceding library. Single/repeated formulas,
bookmarks, run splitting, wider Roman starts and adjoining marker-run text are
covered. Marker raise/font ratio agree within 0.040pt/0.000187; X residual is
1.243pt. All eight visual gates remain failed, as do the generic rate/time token
screens; source prose and simple-math deficits are zero. Four pairs reveal a
separate significant defect: endnote story gap is excessive by 8.163pt in Final
table cases and about 6.213pt in review body/table cases. Notes remain visible;
broader composition acceptance waits for semantic spacing work. No fitting branch
is added. Frozen `3b79f164` transfers 17 assets/four exact PDFs, including C3 review
and held-out memo Final, without root build or suite/corpus replay. Evidence:
`artifacts/composition-c1/math-endnotes-evidence.json`. CI through `500cdacf` is
green on Windows and Ubuntu.

The document-endnote start boundary now belongs to its trailing body paragraph,
with numeric after-spacing and empty-separator flow resolved locally. Earlier
tables and review print scaling no longer add 8.163pt / 6.213pt of excess gap.
Whole-note typography, continuation, mixed-footnote and between-note guards remain.
Frozen 2b320cd8 / root 4b811739 passes 42 affected methods in one invocation and
six Ubuntu checks; both new regressions fail with the preceding library. Eight
independent Word control outcomes agree within 0.163pt on the story-baseline gap.
Body text operations are exact and only note/separator baselines move; four visual
pairs were reviewed. All eight numerical visual gates and generic rate/time token
screens remain failed; source prose/simple-math deficits are zero. The checked
17-asset/four-PDF transfer preserves C3 review and held-out memo Final output,
without rebuilding or replaying a root suite. Evidence:
`artifacts/composition-c1/endnote-boundary-evidence.json`. A held-out complete
endnote report is the next acceptance step; reduced controls alone do not
establish general related-story composition.

The held-out composed-field report also passes the bounded endnote-composition
checkpoint after mechanically replacing its footnote with an endnote. Visible
prose, six equations, image bytes and unrelated package parts are preserved.
Both presets retain two pages, all source/Office global and page-local prose and
simple math, and four navigation targets; the note destination moves to page two.
Four page pairs were reviewed without a visible omission or serious pagination,
overlap or clipping defect. Note-baseline residual is 0.256pt Final / 0.023pt review.
Both numerical visual gates remain failed; typography and annotation residuals
remain explicit. Source 0b8ceeec reuses frozen 2b320cd8 with all 17 asset hashes
rechecked, no rebuild, suite replay or new fitting branch. The shared authoring
transformation also regenerates the four reduced controls byte for byte.
Evidence: `artifacts/composition-c1/composed-endnotes-evidence.json`.

Direct normal footnote/endnote paragraphs now retain supported mixed equations.
Preparation and diagnostics preserve their note-part source identity; ordered
measurement supplies occupied line boxes, and the common printed note map scales
math and annotation extents together. After-spacing belongs to the final line.
Frozen c57f2dd2 (compiled fd6abcfe) / root c24c1249 passes 56 affected methods in
one invocation and six Ubuntu checks; three regressions fail with the previous
library. Bookmarks, run splitting, repeated formulas, wrapping and real multipage
continuations retain source content and destinations. Eight independent Word
outcomes retain all four body/note fractions and prose, with first-note baseline/X
errors below 0.367pt / 0.062pt. Five crops were reviewed. All eight numerical gates
and generic rate/time token screens remain failed; source prose/simple-math
screens have zero deficits. A significant review-endnote width defect remains:
the wrapped note adds a line, shifting numerators 24.645pt horizontally / 11.439pt
vertically. Broader note-content placement acceptance waits for that correction.
Images, note tables and special-separator equations retain staged fallbacks.
The 17-asset/four-PDF transfer preserves held-out endnote-report review and memo
Final without rebuilding or replaying a suite. Evidence:
`artifacts/composition-c1/note-content-evidence.json`.

Normal footnotes and endnotes now share their owning section's full design text
width before the review print transform. This removes the extra wrapped endnote
line and 24.645pt X / 11.439pt Y formula displacement. Frozen 86b6dd92 / root
eb81f193 passes 60 affected methods in one invocation and four Ubuntu checks;
both new regressions fail with the preceding library. Eight Word outcomes retain
all four fractions and source prose/simple math; note numerator residuals are at
most 0.704pt X / 0.381pt Y. Six identical PDFs reuse comparisons, and both changed
review crops are reviewed. All eight numerical gates and generic math-token
screens remain failed. The complete C7 report changes only two note clip widths:
text operations and both raster pages remain exact. Its affected review outcome
passes source/page-local screens. The checked 17-asset/four-PDF transfer uses the
verified new PDF identity without a root build, suite or corpus replay. Audit
selection now honors measured Both/Final/Review modes; 14 adversarial controls
retain missing/extra/duplicate rejection on Windows/Ubuntu. Broader composed
note-math acceptance remains the next step. Evidence:
`artifacts/composition-c1/note-width-evidence.json`.

The held-out composed endnote report also retains two supported fractions inside
its note, alongside all six original equations, prose, image, cached fields and
annotations. Only the endnote package part changes. Both presets retain two
pages, pass global/page-local prose and simple-math presence screens, and preserve
four navigation targets with the note destination on page two. All four page
pairs were reviewed without visible omission or serious pagination, overlap or
clipping defects. Note-baseline residual reaches 0.302pt; note-numerator horizontal
residual reaches 1.682pt. Both numerical visual gates remain failed. Source
7ede0fb3 reuses frozen 86b6dd92 with 17 asset hashes checked, no library rebuild,
suite replay or new fitting branch. The authoring command's original mode still
regenerates the prior C7 source byte for byte. Evidence:
`artifacts/composition-c1/composed-note-math-evidence.json`.

Direct normal notes now compose inline images and equations through shared
measured frames. Note-image positions and dimensions join the text/math printed
map; owned inline frames also retain local endnote start-spacing ownership.
Existing whole-flow, table/floating and rich-story fallbacks remain. Frozen
65b9c4fe / root 2cc71652 passes 67 affected methods in one invocation and five
Ubuntu checks; three regressions fail with the previous library. Four Word
control documents in both presets cover inline/wrapped notes and two image sizes.
All four fractions, PNG, prose and destination survive; source prose/simple-math
deficits are zero. Eight exact prototype PDFs reuse comparisons, and four source/
mode identities reuse Word references. Four crops were reviewed. All eight
numerical gates and generic rate/time screens remain failed. PNG raster-bound
residuals reach 0.5pt top/size and 1pt horizontal prose advance. Native dimension
checks account for the two-edge 144 DPI sampling uncertainty; visual thresholds
are unchanged. The 17-asset/four-PDF transfer preserves complete note-math report
review and held-out memo Final without a root rebuild or suite/corpus replay.
Complete-document note-image acceptance remains the next finite step. Evidence:
`artifacts/composition-c1/note-images-evidence.json`.

The finite normal-note composition checkpoint is also accepted on the complete
held-out report with a note image beside its equations. Source 258b3328 adds only
the note drawing and its relationship, reusing the original PNG bytes. Eight
source equations, two image placements, prose and four navigation targets survive
both presets on two pages, with zero global/page-local content deficits. Both
first-page rasters are exact for candidate and Word, reusing their prior reviews;
the two new second-page pairs were reviewed without visible omission or serious
pagination, overlap or clipping defects. Note-image dimensions agree in the
sampled raster; top/X residuals reach 0.5pt / 1.5pt. Both numerical visual gates
remain failed. All 17 assets of frozen 65b9c4fe were rechecked with no library
build, suite replay or fitting branch. Further note variants are outside this
finite acceptance step. Evidence:
`artifacts/composition-c1/composed-note-image-math-evidence.json`.

Shared inline counter measurement now resolves PAGE from its actual page number
and NUMPAGES from its actual total, without requiring the unrelated input. Mixed
counter paragraphs still require both; source ownership and missing/noncanonical
fallbacks remain. Simple/packed body/cell source controls cover values 9/10/12 and
resolved advances. This is a dependency prerequisite; body/table page-aware math
admission remains deferred. Frozen 8fe300f6 / root 6190893d passes 26 capability/
neighbor methods in one invocation and five Ubuntu checks. The measurement
regression fails with the preceding library through existing APIs; the new
resolver-API unit is excluded from that parent check. The 17-asset/four-PDF
transfer preserves the complete note-image/math report in both presets, the
twelve-page static-header counter control and held-out memo Final. No root build,
suite/corpus replay or Office re-export; prior comparisons remain reusable with
unchanged failed numerical gates. Evidence:
`artifacts/composition-c1/counter-inputs-evidence.json`.

Reusable inline measurement and source-driven wrapping are now separate. Frozen
9550187c / root 3527ff81 passes 23 exact capability/neighbor methods in one
invocation and ten Ubuntu checks. A source cursor preserves authored breaks,
tabs, first-line numbering and wrapping frames; split counters retain the value
bound when their first digits were placed. Rebinding updates only changed counter
fragments and keeps canonical source ownership. All ten fixed-portfolio PDFs are
byte-identical to prior qualified outputs, reusing independent Office comparisons
and their explicit visual failures. Transfer verifies 17 assets/four exact PDFs
without a root rebuild or suite replay. This establishes the contract for actual
page-aware admission; body/table PAGE-math support and measured performance gains
are not claimed. Evidence: `artifacts/composition-c1/inline-cursors-evidence.json`.

Simple table cells now preserve supported equations beside PAGE counters. The
shared preparation policy distinguishes PAGE dependencies from an unavailable
total count; measurement and painting use actual placement inputs. Existing cell
height calculation and dynamic memo keys already carry those inputs, so no new
layout flag or cache policy was added. Frozen 2086a381 / root 05dff46f passes 25
capability/neighbor methods in one invocation, six Ubuntu checks and two genuine
parent-library rendering regressions. The independent Word control preserves four
fractions and actual values across pages nine and ten in both presets. Relevant
page pairs were reviewed through exact preset raster reuse, with no visible
omission or serious overlap/clipping. Baseline error is 0.029pt; counter/equation X
residuals reach 1.041pt / 2.349pt, including Word's Calibri field digits versus
canonical Arial. Both visual gates and the generic dynamic-field source screens
remain failed/unresolved; no simple-math deficit is observed. Transfer verifies
17 assets/four exact PDFs without a root build or suite replay. Body PAGE and
unknown total-count composition remain deferred. Evidence:
`artifacts/composition-c1/cell-page-counters-evidence.json`.

Body PAGE/math composition now binds counters after line admission. A source
cursor retries unchanged when pagination moves a line, including past an unfit
continuation column. A confirmed trailing-whitespace cursor defect was reproduced
and corrected. Fixed paragraphs reuse their measured wrap; counter indices/tab
stops are retained and rebinding occurs on page changes. No speedup is claimed.
Frozen 645d5508 / root c1f41899 passes 33 exact capability/neighbor checks in one
invocation, fifteen Ubuntu checks and four genuine parent behavior regressions.
Simple/packed source tests retain bookmarks and run splits. The independent Word
simple-field control places one paragraph across pages nine and ten with two
fractions and correct values 9 then 10. Relevant page pairs were reviewed through
exact preset raster reuse, with no visible omission or serious overlap/clipping.
Both visual gates remain failed: baseline residual reaches 0.105pt, authored
leading-space trimming shifts the second prose start 3.384pt, and field/numerator
X residuals reach 3.456pt/2.148pt. Generic dynamic-field source screens remain
unresolved, with no simple-math deficit. Ten fixed-portfolio PDFs are exact;
transfer verifies 17 assets/four exact PDFs without a root rebuild or repeated
suite/corpus. Complete compound-document acceptance is next; unknown total-count
composition remains staged. Evidence:
`artifacts/composition-c1/body-page-counters-evidence.json`.

The finite PAGE composition checkpoint is accepted on the complete held-out
report (source 7a8208ea), using frozen 645d5508 without a library rebuild or suite
replay. All 17 assets were rechecked. Existing cached field words remain, with
three PAGE fields added beside body math/images, inside an internal hyperlink and
in the math/note table cell. Eight source equations, two image placements and four
logical navigation targets survive both presets on two pages; the wrapped link
uses two rectangles, giving five painted navigation rectangles. The note target
remains on page two. All four page pairs were reviewed against independent Word;
no visible omission or serious pagination/overlap/clipping defect was observed.
Global/page-local prose and simple-math deficits are zero. Generic source screens
retain their explicit unresolved dynamic-field evaluation; both numerical visual
gates remain failed. Typography, spacing and annotation geometry remain cosmetic
residuals; PNG top/X residuals reach 1pt/1.5pt with half-point raster sampling.
Unknown total-count composition remains staged. Evidence:
`artifacts/composition-c1/composed-page-counter-evidence.json`.

Shared ordered wrapping now preserves authored leading spaces and the closing
paragraph line after a trailing hard break. The paragraph-end cursor represents
the source mark explicitly; soft-wrap whitespace remains discarded. A combined
body/cell/header control corrects accumulated occupied-height errors up to 29.175pt.
Eighteen independent Word anchor comparisons have maximum X/Y errors of
0.118pt/0.139pt. The existing cross-page PAGE control's former 3.384pt start offset
falls to 0.05pt. Frozen a08dcfa6 / root 323c0539 passes 38 capability/neighbor
methods in one invocation, sixteen Ubuntu checks and four genuine parent behavior
regressions. All eighteen equations and prose remain visible in the reviewed
control; both preset page pairs share exact raster identities. Both numerical
visual gates and generic rate/time whole-word screens remain failed; source prose
and simple-math deficits are zero. Ten fixed-portfolio PDFs plus both complete
PAGE-report PDFs remain exact. Transfer verifies 17 assets/four exact PDFs with
no root build or repeated suite. Evidence:
`artifacts/composition-c1/inline-whitespace-evidence.json`.

Parent regressions now use `tools/NewParentRegressionRuntime.ps1`, with independent
owned input copies and rejection of unowned/linked destinations. Thirteen Windows
fixture-sharing/ownership checks pass and run in existing CI on both platforms;
local Ubuntu lacks pwsh, while its .NET checks pass. This follows a confirmed
agent setup error: a legacy Cases junction received six unpinned additions and
one shared input lost its read-only flag. Existing pinned bytes remain exact and
protection is restored. Automatic review blocked removal of the additions and old
junction; both remain, with no retry. The historical 444f5a55 input set is excluded
from strict immutable identity claims. All 1,337 current frozen inputs retain
pinned hashes/protection and the six current parent inputs have independent file
ownership. Ownership repair/rejection evidence remains beside the delivery report.

The measured font-maintenance batch reduces warm portfolio allocation to
28.75–33.83MB for Final DOCX and 32.34–37.36MB for review, with unchanged PDFs.
Counted kerning source batches replace the expanded pair list, retaining bounded
source contributions, last-wins lookup values and cancellation. All 715,743 map
entries across the three actual faces match. Frozen f35fa653 / root 993543d6 passes
31 Windows methods in one invocation and 17 Ubuntu checks (four installed-font
checks skip); transfer verifies 17 assets/four PDFs. The ten allocation-probe
whole/staged PDFs provide corpus identity, without an extra CLI sweep. A generated
out-of-range class-matrix regression confirms an older managed bounds exception;
wide offsets now preserve the valid prefix. This is not evidence of corruption.
Matched sustained-warm timings overlap, so no speedup is claimed. Office output
and existing failed visual gates remain unchanged. See
`docs/OperatingEnvelope.md` and `artifacts/composition-c1/kerning-stream-evidence.json`.

Simple header/footer table cells now share body-cell math preparation and consume
prepared boxes and actual PAGE/NUMPAGES inputs in grid/height/painting. A confirmed
unchanged-content transformation omitted every header equation before this change.
Static table rows also join painted-equation diagnostic enumeration. Existing
simple-cell exclusions still apply to merged, fit-text, vertical, nested/rich and
note-table owners; no font/size/document guard was added.

The independent mixed/plain controls also exposed a footer table 38.664pt too low.
The complete measured block now shares bottom anchoring and existing row/image
translation; table top edges participate in the footer keep-out. Sixteen Word
anchors have maximum X/Y errors 0.024pt/0.140pt. All eight source equations, four
image placements and page fields survive both presets; the twelve-page portable
control verifies actual counters across 9/10. All four page pairs were reviewed
through exact preset raster reuse, with no visible omission or serious
pagination/overlap/clipping. Four numerical visual gates remain failed, and the
generic dynamic-field screens remain explicitly unresolved; prose/simple-math
deficits are zero.

Frozen f684dc29 / root 8bf5354e passes 32 exact Windows methods in one invocation,
15 Ubuntu methods and five genuine parent behavior regressions. Eight affected
fixed DOCX PDFs retain bytes; PPTX production is unchanged and the held-out deck
passes transfer smoke. Transfer verifies 17 assets/four PDFs with no root rebuild
or suite replay. Evidence: `artifacts/composition-c1/static-cell-evidence.json`.

The complete held-out static-cell report preserves every original package part
except its header/footer stories, adding mixed footer math/image/page fields.
It exposed two consequential integration defects beyond the simple controls:
a row's bottom Y was used as its top keep-out, and scaled review tables painted
below the page. The boundary now uses Y + Height. Simple static tables join the
existing complete-frame transform for text, math, images, clips, cells and borders;
printed paragraph admission converts the already resolved footer/note reserve
once. Shared simple-cell eligibility and the inherited physical-coordinate marker
are reused; no new flag, font/size/document guard or quota was added.

The accepted report retains ten source equations, three source image placements
and five navigation rectangles on two pages in each preset. All four page pairs
were reviewed against independent Word: the footer remains visible, section seven
starts on page two, and no serious omission/overlap/clipping is observed. Twelve
static prose anchors have maximum X/Y errors 0.024pt/0.422pt. Both numerical visual
gates still fail. Generic source screens remain unresolved: cached digit-one gaps
match Word, and whole-word rate/time extraction gaps remain; all simple expressions
are present. Typography, spacing, note-image alignment and annotation geometry
remain cosmetic residuals.

Frozen d0cc21fe / root 5b1af0af passes 29 exact Windows methods in one invocation;
22 Ubuntu checks pass and seven Office-font checks skip. The portable complete
report normalizes mathematical letters to the pinned face's supported ordinary
alphabet; the Windows Office control uses the original document. Two regressions
fail against the preceding artifact. Eight fixed DOCX PDFs and four prior static
controls remain exact. Transfer verifies 17 assets/four PDFs without a root rebuild
or suite replay. An intermediate keep-out artifact remains as the regression
parent; final promotion followed review of all presets. Evidence:
`artifacts/composition-c1/static-table-print-evidence.json` and
`artifacts/composition-c1/composed-static-cell-evidence.json`.

Ordered inline painting now rejoins adjacent word atoms within the same source
owner and identical font/spacing state. Wrapping retains word atoms and source
extents; painting retains glyph-run continuity with linear string assembly. This
improves all eight portfolio first-page Office comparisons (MAE down 0.04146 Final /
0.01897 review), with later-page metrics unchanged, and all four complete-report
page pairs. Eight classic failure sets remain unchanged; small cosmetic wrapping
metric tradeoffs are retained. All existing numerical thresholds remain.

A broader single-equation eligibility prototype is rejected: it discarded 0.348pt
of script occupied leading and added a failed visual gate. The original eligibility
and script/fraction PDFs remain unchanged; no script/font/size exception is added.
The next investigation is the shared prepared-frame versus ink-fit contract.

Frozen 94f6f425 / root 8994b85a passes 36 Windows methods in one invocation and
12 Ubuntu checks; 24 native/Office-font controls skip without their prerequisites.
Two missing test prerequisites reproduced parent failures and are now explicit.
Text-show anchor extraction accounts for leading spaces, preserving the original
Office values/tolerances. The test-only correction retains qualified library
SHA-256 44A76742D209D729F993E9BC692E9D53EAA81313178F942A29C9AF3499BBA354;
all completed conversion/comparison work is reused.

Ten portfolio outcomes retain zero prose/simple-math presence deficits; eight DOCX
gates fail and both decks pass. The complete report retains two pages in both
presets, all ten source equations, three source image placements and five navigation
rectangles; all four page pairs were reviewed without serious omission/overlap/
clipping. Its generic dynamic-field/rate-time screens remain unresolved, and both
numerical gates fail. Transfer verifies 17 assets/four PDFs with no root build or
suite replay. No speedup is claimed. Evidence:
`artifacts/composition-c1/ordered-glyph-run-evidence.json`.

### Prepared occupied frames (C26)

Ordered wrapping no longer substitutes ink bounds for a prepared equation frame
when the ink fits the ordinary text strut. It unions the complete occupied frame
with that strut once, preserving leading without a font, size or script exception.
Two new regressions fail on the parent and pass on frozen 23d10821 / root d72e938d.
Eight independent Word baseline gaps are within 0.15pt in both presets; all four
new script page pairs were reviewed without observed omission, overlap or clipping.

All ten fixed portfolio and both complete-report PDFs retain their accepted bytes.
The eight classic controls keep their failed gate sets. One wrapped literal
improves MAE by 0.03090; another loses 0.04248 foreground recall and adds 0.02014
MAE despite a closer formula baseline. Its following-prose line-height mismatch
remains. Script comparisons also retain cosmetic failures. Thresholds and
single-equation eligibility are unchanged; marker migration remains deferred.

Qualification runs 33 Windows methods once and 8 Ubuntu checks (25 font-specific
skips). Transfer verifies 17 assets and four exact PDFs without a root rebuild or
suite replay. The preceding test-only CI repair uses visible comment/tab text
anchors, preserving Office bounds and tolerances; Windows and Ubuntu CI are green
through 1ba9d797. Evidence: `artifacts/composition-c1/inline-occupied-frame-evidence.json`
and `artifacts/composition-c1/glyph-run-anchor-repair-evidence.json`.

### Ordinary-word emergency wrapping (C28)

Default/on paragraphs now emergency-break ordinary words wider than their text
frame, using the existing bounded Unicode-safe search. Protected spans retain
their existing fallback. Independent Word controls cover default/on/off in both
presets: default/on recall improves from 0.377 to 0.828 and MAE from 0.278 to 0.094;
explicit-off metrics remain unchanged. Relative line gaps are within 0.15pt.
Candidate `formu/la` versus Word `form/ula` remains a character-break residual.
All six visual gates still fail recall/SSIM; thresholds remain unchanged.

All ten fixed portfolio and both complete-report PDFs retain accepted bytes.
A formerly omitted narrow-column script now survives, as in Word, but sits one
line too high because of differing text breaks. This significant geometry gap
remains open; broad single-equation migration is rejected and remains deferred.
No observed omission, overlap or clipping in reviewed new page pairs; identical
rasters reuse reviews across presets/settings.

Frozen a6f1c557 / root d4bb5b02 passes 81 selected Windows methods once and 54 Ubuntu
checks (27 font-specific skips). Two new regressions fail on the accepted parent.
The test-only correction retains the qualified production DLL and tests the
actual whole-formula fallback boundary at an 8pt frame. Transfer checks 17 assets
and four exact PDFs without a root build or suite replay. Evidence:
`artifacts/composition-c1/emergency-word-wrap-evidence.json`.

The following CI repair restores image-driven admission before an emergency word
split and retains the remaining token's source range. The existing image-position
test now also checks complete prose preservation. A synthetic keep-lines test
retains its two-page result and now checks the complete three-line owner, reflecting
the corrected overwide-word behavior. Frozen c53a6a76 / root 626f1321 passes 93
Windows methods once and 62 Ubuntu checks (31 font-specific skips). All twelve
fixed/complete PDFs retain bytes; 17 assets/four smoke PDFs transfer without a root
build or suite replay. Evidence: `artifacts/composition-c1/emergency-image-ownership-evidence.json`.

### Balanced literal groups and ordered-math links (C31)

Balanced literal groups can break at binary/relation operators, matching independent
Word controls for parenthesized and nested groups. Previously omitted formulas now
retain all seven/nine glyphs on Word's two mathematical rows in both presets.
Matching delimiter kinds are checked; malformed/unary relocation fallbacks remain.
Hyperlinks retain ordered source ranges and merge mathematical bounds through the
existing rectangle builder. Tests cover both linked rows, bookmarks, run splitting
and the printed review transform. The optional test reader now applies graphics
transforms; production PDF validation is unchanged.

Maximum independent baseline/local-glyph differences are 0.113/0.060pt. Absolute
first-row X differs by up to 0.742pt; source-run splitting adds a 0.252pt row-origin
residual. All four new visual gates still fail recall/SSIM. Four preset pairs were
reviewed without observed omission, overlap or clipping; thresholds remain.

Frozen 14a11946 / root e390f2d7 passes 51 Windows methods once and 31 Ubuntu checks
(20 font-specific skips). Two new methods fail on the accepted parent. All twelve
fixed/complete PDFs retain accepted bytes; 17 assets/four exact PDFs transfer without
a root build or suite replay. Evidence: `artifacts/composition-c1/grouped-literal-math-evidence.json`.


### Qualified complete inline-box placement (C32)

Complete supported unary, fraction and script boxes share ordered placement and
source-backed mathematical hyperlink bounds. A fitting unary formula previously
omitted at its source-line boundary now moves to Word's next row. Measurement
marks relocation eligibility; malformed literal groups remain source-bound.
Alignment, whole-width and richer-content fallbacks remain explicit.

Independent unary origin/baseline differences are within 0.15pt; following visible
text X differs by less than 0.2pt, in both presets. Both new native regressions fail
on C31. All source prose survives; raw extraction warnings for a phrase split over
lines and mathematical Greek remain recorded. Native glyph/geometry checks cover
the corresponding visible content. Both new visual gates still fail recall/SSIM.

Frozen cab3445b / root 18aba4d0 passes 78 Windows methods once and 37 Ubuntu checks
(41 font-specific skips). Ten fixed outcomes pass prose/simple-math screens.
Changed DOCX PDFs receive new comparisons: first-page MAE changes are below 0.001,
confined to text beside a fraction; page counts and later pages remain unchanged.
The complete report retains its generic field/math extraction warnings. Reviewed
unary and changed table pairs show no new omission, overlap or clipping; existing
indentation differences and failed gates remain. Transfer verifies 17 assets/four
exact PDFs without a root build or suite replay. Evidence:
`artifacts/composition-c1/atomic-inline-placement-evidence.json`.


### Unchanged-format source-run measurement (C33)

Splitting a source run with unchanged formatting no longer resets ordered-text
boundary kerning. The confirmed 0.252pt grouped-formula origin drift is corrected;
source owners, offsets and annotation ranges remain separate. Different states,
dynamic fields and automatic markers retain their fallback. Global kerning defaults
and cross-source painting coalescence remain unchanged.

Two source controls, including the held-out nested group, retain absolute glyph
positions within 0.001pt in both presets. Four independently exported Word rasters
are identical before/after splitting; all four candidate rasters likewise match
their accepted unsplit output. The existing absolute Word-origin residual up to
0.742pt and recall/SSIM failures remain; thresholds are unchanged.

Frozen 16b0c7fe passes 28 Windows methods once and 14 Ubuntu checks (14 font skips).
Both regression methods fail on C32. All twelve fixed/complete PDFs retain bytes;
Office comparisons, content screens and visual reviews are reused through checked
identities. Transfer verifies 17 assets/four exact PDFs without a root build or suite
replay. Evidence: `artifacts/composition-c1/inline-run-boundaries-evidence.json`.


### Fitting aligned mathematical fields (C35)

Right/center tab fields now measure complete text/math content through the shared
ordinary tab-start rule. Previously omitted fractions survive. Overlapping or
out-of-frame fields, decimal/bar tabs and leaders retain explicit fallbacks.
Source extents scale with painted static/note geometry, and mathematical link
rectangles include denominator bounds even when a printed cell's text slot is
shallower. Default-tab regressions isolate those annotation defects separately.

Four public controls produce eight mode outcomes across body/cell/footer placement,
including bookmarks, unchanged-format run splitting and formula links. Independent
maximum text-field X error is 0.096pt; mathematical anchor X/relative-Y errors are
0.049/0.073pt. The absolute footer baseline residual is 0.375pt. All eight reviewed
mixed-field regions retain visible contents without new observed overlap/clipping;
this does not establish full-page clipping coverage. All eight numerical gates
still fail. Generic rate/time extraction warnings remain recorded alongside exact
native glyph/geometry checks; thresholds are unchanged.

Frozen 924eb51a / root cf4bf629 passes 28 Windows methods once and 16 Ubuntu checks
(12 font skips). Three new methods fail on the accepted parent. All twelve fixed
and complete-report PDFs retain bytes, comparisons, content screens and reviews.
Transfer verifies 17 assets/four PDFs without a root build or suite replay.
Evidence: `artifacts/composition-c1/aligned-inline-fields-evidence.json`.

A separate independent probe validates the existing sole-fraction hyperlink
styling guard: Word itself switches that combination to smaller left-aligned
inline styling. That guard remains; no broad eligibility deletion was made.


### Explicit centered display-source groups (C37)

Two supported formulas in explicit `m:oMathPara` source containers no longer
disappear. Preparation reuses the existing display-source-row ownership and ordered
layout while retaining original source objects and cancellation. Mixed-prose and
non-centered explicit groups retain their fallback.

Both modes and same/separate source containers retain all ten mathematical glyphs,
including bookmarks and unchanged-format run splitting. Independent baseline/origin/
operator errors are 0.009/0.193/0.248pt. The following-prose pitch residual is 0.599pt;
no font/size correction was added. Four page pairs share one reviewed raster pair,
with no observed omission/overlap/clipping. All four visual gates still fail recall/
SSIM. Raw joined-Greek extraction warnings remain alongside exact glyph/anchor checks.

Frozen d6083a42 passes 17 Windows methods once and seven Ubuntu checks (ten font
skips); both new methods fail on C35. All twelve fixed/complete PDFs retain bytes,
Office comparisons, content screens and reviews. Transfer verifies 17 assets/four
PDFs without a root rebuild or suite replay. Evidence:
`artifacts/composition-c1/explicit-display-groups-evidence.json`.


### Hyperlink-owned explicit mathematical style (C39)

Supported formulas inside hyperlinked explicit display containers no longer
disappear. Ordered source ranges resolve hyperlink ownership once; math style is
selected before measurement while original source objects remain unchanged.
Independent fractions establish inline argument sizing even for explicit display
source. Ordered layout consumes the resolved style, preserving the existing
implicit-formula hyperlink behavior without a font/size-specific correction.

Four inputs/eight outcomes cover body, cell, printed footer and a single fraction.
Exact mathematical content, Word anchors/sizing, clickable glyph bounds, source
identity and cancellation pass. Maximum local-X error is 0.381pt; absolute table-X
and footer-Y residuals remain 1.147/0.848pt. Eight regions reviewed through six unique
raster pairs retain formulas without new observed omission/overlap/clipping. All
numerical visual gates still fail. Six generic joined-Greek extraction warnings
remain alongside exact glyph checks; both fraction presence screens pass.

Frozen 0327073d passes 25 Windows methods once and nine Ubuntu checks (16 font skips).
The native regression fails on C37. Twelve fixed/complete PDFs and four earlier
implicit-style controls retain bytes. Transfer verifies 17 assets/four PDFs without
a root build or suite replay. Evidence:
`artifacts/composition-c1/explicit-display-links-evidence.json`.


### Authored display-group alignment (C40)

Uniform left/right explicit display groups no longer lose their formulas. Each
measured line owns one nullable display alignment; its display-role flag is derived.
Math-only paragraphs use that owner even when paragraph alignment is justified.
Unknown/mixed alignment and unsupported prose contexts retain explicit fallback.

Four inputs/eight modes retain ten mathematical glyphs, independent Word anchors,
bookmarks and unchanged-format run splitting across body/cell/footer placement.
Existing table/footer geometry and following-prose pitch residuals remain. Eight
reviewed regions show no new observed omission/overlap/clipping; all eight visual
gates and joined-Greek extraction warnings remain failed/unresolved. Numerical
thresholds are unchanged; exact native glyph/anchor checks cover those contents.

Frozen 0e3aa24c passes 18 Windows methods once and seven Ubuntu checks (11 font skips).
Two causal regressions fail on C39. Twelve fixed/complete PDFs retain accepted
bytes, comparisons, screens and reviews. Transfer verifies 17 assets/four PDFs
without a root build or suite replay. Evidence:
`artifacts/composition-c1/authored-display-alignment-evidence.json`.

### Finalized total-page fields alongside equations (C41)

Simple and complex NUMPAGES fields beside supported body/cell fractions no longer
discard the equations when pagination is stable. Preparation records deferred
dependencies without loading unused math fonts. One additional complete note-aware
coordinator pass receives the actual total; placement, keep estimates and table
rows share that input. Acceptance requires an unchanged final count and all
deferred equations placed. Changed counts retain the original fallback. Source
counter runs remain canonical; no guessed count or fixed-point loop is introduced.

Four inputs/eight outcomes distinguish actual total 2, current page 1 and stale
cache 777. Independent Word geometry, sixteen mathematical glyphs, bookmarks/run
splitting, portable normal-note composition, cancellation and unstable-count
fallback pass. Existing field-font/cell-origin residuals accumulate to 1.975pt;
fraction baseline/local-X errors remain within 0.158/0.574pt. Four simple first-page
pairs and both shared second-page pairs were reviewed without observed content
loss or new serious overlap/clipping. Complex first pages retain native checks.
All eight numerical visual gates fail at unchanged thresholds. Simple-math
presence screens pass; raw stale-cache and fragmented-prose token warnings remain.

Frozen 3eed2672 retains the production library from 5e838c1b and passes 25 Windows
methods in one invocation plus 17 Ubuntu checks (eight font skips). An earlier
incorrect fixture-count assertion was corrected without rebuilding the library or
repeating PDF conversions. Two content tests fail on C40. Twelve fixed/complete
PDFs retain accepted bytes/comparisons/screens/reviews; transfer verifies 17 assets
and four PDFs without a root build or suite replay. Focused warm allocation is
27.99–28.47MB Final and 31.47–31.90MB review; overlapping stages are not additive,
and no speedup is claimed. Evidence:
`artifacts/composition-c1/total-page-math-evidence.json`.

### Explicit formulas beside ordinary prose (C42)

An explicit display flag no longer discards supported formulas sharing a source
row with prose. One authored-row classifier distinguishes prose and math-only
roles; preparation resolves inline argument sizing before measurement while
preserving source objects. Body, simple cells and footers use the existing ordered
path. Math-only rows retain their display role across authored breaks; unknown or
mixed alignment and unsupported paragraph contexts retain fallback. Counter
availability belongs to measurement rather than source-row classification.

Six inputs/twelve modes cover one or both formulas marked explicit, literals and
fractions, bookmarks/run splitting and cell/footer placement. Both source flags
produce identical Word rasters and candidate rasters for each body pair. Exact
mathematical text, independent sizing and anchors pass; maximum absolute X/Y and
local-X errors are 1.094/0.367/0.274pt. Four complete page pairs were reviewed,
covering six outcomes through raster identities, with no new observed content
loss or serious overlap/clipping. Other outcomes retain native checks. All twelve
visual gates still fail; all source-prose/simple-math presence checks pass, while
eight generic mathematical extraction screens remain unresolved.

Frozen c972b1b3 passes 26 Windows methods once and 14 Ubuntu checks (twelve native
font skips); two causal regressions fail on C41. Portable tests use identically
upright Greek on both sides of the source transformation and preserve explicit
NUMPAGES composition. Twelve fixed/complete PDFs retain accepted identities and
comparisons. Transfer verifies 17 assets/four PDFs without a root build or suite
replay. Eight Office exports were unnecessarily repeated during probe expansion;
their references are retained, with no subsequent export repetition. Evidence:
`artifacts/composition-c1/mixed-display-rows-evidence.json`.

### Picture-owned inline formulas and equation boundaries (C43)

Explicit formulas beside visible inline pictures now retain inline argument
sizing instead of disappearing. The shared source-row classifier accepts visible
text/pictures while retaining tab-only and unsupported fallbacks. Fraction boxes
publish Inner atom edges, compound rows preserve their edge roles, and adjacent
equation boxes reuse the existing mathematical glue. This prevents two fraction
bars from joining. Source equations and images retain their identities.

Six inputs/twelve outcomes cover body/cell placement, pictures before/between two
fractions and separate/combined source representations. Word and candidate rasters
are identical across each before/combined pair. All sixteen mathematical glyphs,
inline sizing, picture painting/order, bookmarks/run splitting and portable source
equivalence pass. Maximum local-X error is 0.109pt; the existing table-X residual
remains below 0.87pt. All source-prose/simple-math screens pass; twelve generic
fraction-token screens and all twelve numerical visual gates remain failed or
unresolved at unchanged thresholds.

The printed-cell/table origin remains about 2.6pt high, with its upper border
crossing the preceding text line. An older C42 implicit control and independent
Word export reproduce this defect. It remains the next placement blocker; the
native bound records it without claiming fidelity. Four page pairs cover six
outcomes through raster identities and show retained images/distinct fraction
bars alongside that visible defect. Other outcomes retain native checks.

Frozen 61be3e0b passes 42 Windows methods once and 23 Ubuntu checks (19 font skips).
Two content regressions fail on C42. Twelve fixed/complete PDFs retain accepted
bytes/comparisons/screens/reviews. Transfer verifies 17 assets/four PDFs without
a root build or suite replay. Evidence:
`artifacts/composition-c1/display-image-rows-evidence.json`.

Microsoft documents [adjacent math-zone merging](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/1d77457b-2884-4749-9b4a-c150ca13cc19)
and [OfficeMath spacing](https://devblogs.microsoft.com/math-in-office/officemath/).
Using existing atom glue across prepared boxes is an inference confirmed by the
independent controls here.


## C44: owned picture and table frames

The C43 printed-cell placement blocker is corrected. Ordered pictures retain their
measured source-line owner; shared table translation moves the line, picture,
clip and optional picture frame together. Unowned pictures retain the existing
whole-table fallback before mutation. First-page pinning uses the measured text
strut rather than object-inflated ascent, and object-only paragraphs consume
retained paragraph-mark size. Removing a neighboring prose cell now preserves
the picture/math baseline without a font, size or document-shape exception.

Eight inputs/sixteen independent Word outcomes cover body/cell/first-table,
pictures before/between fractions, combined source formulas and neighboring-cell
removal. Bookmarks/run splitting retain content and placement. Maximum baseline
error falls from 2.567pt to 0.170pt; X/local-X remain within 0.916/0.109pt.
Four reviewed page pairs retain pictures and fraction rules, with the table top
clear of preceding prose. All sixteen visual gates still fail unchanged thresholds;
the three changed review controls now fail structural similarity only. Their mean
pixel error falls approximately 0.60 to 0.19, with foreground recall about 99.85%.
Source-prose/simple-math screens pass; generic rate/time extraction warnings remain.

Frozen 988936bd passes 60 Windows methods once, 47 Ubuntu checks (13 font skips),
two causal parent regressions and a 17-asset/four-PDF transfer. Twelve fixed/complete
PDFs retain accepted bytes/comparisons/reviews; no root build or suite replay.
C43 CI caught two older assertions assuming zero inter-formula spacing. Their
replacement checks exact spacing/centering and belongs to both capability sets.
Evidence: `artifacts/composition-c1/owned-picture-frames-evidence.json`.


## Next high-priority maintenance: PPTX picture masks

Reviewed synthetic evidence confirms that ordinary `p:pic` loses preset geometry
while shape picture fills retain it. Ellipse, rounded-rectangle and polygon pictures
therefore paint rectangular corners, independently of working source crop. The
next bounded batch preserves geometry/adjustments and reuses the existing shape
picture-fill masks. Acceptance covers ellipse/non-square, crop, transforms/grouping,
state restoration and adjustment guides against independent Office output. SVG,
outlines, shadows and unsupported/custom geometry require explicit scope and tests;
raster controls alone do not qualify them. Safe managed/dependency-free constraints
and existing visual thresholds remain. Detailed tasks are in `PLAN.md`; retained
public synthetic evidence is under `artifacts/pptx-picture-geometry/`.


## C45: ordinary PowerPoint picture masks

The high-priority picture handoff is implemented for raster and native SVG paint.
Resolved pictures retain preset geometry and adjustment guides; shared shape
readers and picture-fill clips apply the mask in transformed frame coordinates,
independently of source crop/fill and image alpha/recolor. Unknown/custom geometry
retains a diagnosed rectangle fallback. Outlines/shadows also retain explicit
rectangular effect warnings pending the next bounded geometry qualification.
Six-letter SVG color keywords no longer throw in the hex parser and drop paint.

Nine original pages and four new non-square/grouped/SVG/crop-fill-alpha controls
retain masks and following paint. Twelve of thirteen page gates pass the retained
image-family thresholds. The alpha control still fails SSIM (0.971024 versus
0.995; MAE 0.292673), with sampled saturated channels two levels below the
candidate in Office. Seven full-page pairs were reviewed, including transformed
footprints. Original ellipse frame MAE falls 28.27 to about 0.002; rectangles retain
pixels. The nine original Office raster references are reused with source-part
identity proof; only the four new pages required an export.

Frozen 2b91e23d passes 28 Windows methods once and 28 Ubuntu checks. Three parent
regressions fail on C44. Sixteen existing cached image-family cases pass their
individual thresholds; one cache miss retains the C44 PDF. All twelve fixed/
complete PDFs retain accepted bytes/comparisons/reviews. Transfer verifies 17
assets/four PDFs without rebuilding or repeating the root suite. Evidence:
`artifacts/pptx-picture-geometry/evidence.json`.


## C46: shared picture outlines and vector shadows

Picture outlines now reuse preset stroke geometry at Office's external half-width
centerline extent. Unblurred shadows reuse existing preset fill geometry. Five
of six independent Word pages pass retained image-family thresholds, with SSIM
0.9981–0.9999. Five complete pairs were reviewed, including the transformed
footprint and following marker. The blurred-shadow control still paints visible
rectangular gray corners and fails SSIM (0.9612); it remains the next high-priority
blocker. Unknown/custom geometry remains diagnosed.

Accepted test-only freeze 13586991 retains production 632dabf6. The 38-method set
passes on Windows and Ubuntu; an older CRLF-only assertion was normalized and
rechecked without rebuilding production or repeating conversions. Ubuntu CI now
runs the small picture-effects capability. One parent regression fails on C45.
The original thirteen mask pages, two portfolio decks and four image-neighbor PDFs
retain accepted bytes/comparisons. DOCX evidence carries forward from unchanged
rendering source, with only the small transfer smoke repeated. Transfer verifies
17 assets/four PDFs. Evidence:
`artifacts/pptx-picture-geometry/vector-effects-evidence.json`.


## C47: bounded ellipse and rounded picture blur

The existing bounded shadow raster now measures distance to an ellipse or rounded
rectangle, retaining its planes, limits, falloff and retained-byte charging.
Cancellation is checked per row. Rectangle blur retains identical PDF bytes.
Ellipse blur improves SSIM 0.9612 to 0.9988; all six effects pages and three new
rounded/transformed/grouped blur controls pass unchanged thresholds. All four
supported blur pairs were reviewed. Polygon blur retains a diagnosed rectangular
fallback and fails visibly (SSIM 0.8769); the earlier alpha-color residual remains.
Across the handoff's four decks/23 pages, 21 numerical page gates pass and two fail.
This count is independent of the 714 manifests passing integrity checks.

Frozen 7d9b1d77 passes 42 Windows methods once and 42 Ubuntu methods, including
transparent alpha corners, bounded dimensions, retained/live-image budgets and
empty output on failure. One parent regression fails on C46. Original thirteen
mask pages, five vector-effect rasters, two portfolio decks and four image-neighbor
PDFs retain bytes/comparisons/reviews. DOCX evidence is carried from unchanged
source; only the four-PDF transfer smoke repeats. Transfer verifies 17 assets.
Evidence: `artifacts/pptx-picture-geometry/bounded-blur-evidence.json`.
The bounded picture handoff is complete with explicit unsupported/cosmetic limits;
main rendering work returns to supported DOCX composition.


## C48: numbered explicit-math composition

A list label now owns the first authored inline row through shared source-row
classification. Explicit math zones retain inline sizing beside that visible
label/prose/picture; a later math-only row after an authored break retains display
sizing/alignment. This fixes four-formula omissions across body/cell/header/footer
without font, size or document-shape guards. Six public inputs/twelve Word outcomes
retain all 18/22 glyphs, three labels and one picture. Bookmarks and unchanged-format
run splitting retain geometry with native/portable fonts. Maximum Word X/Y errors
are 0.877/0.834pt. Six reviewed pairs cover both modes by raster identity. All twelve
numerical gates still fail unchanged thresholds; source-prose/simple-math presence
screens have no deficits. Unreferenced-header warnings remain separate from visible
content; that part is not printed by Word.

Frozen 1de30ea5 passes 18 Windows methods once, 11 Ubuntu checks (seven native-font
skips), two causal parent regressions and a 17-asset/four-PDF transfer. Eight fixed
DOCX/two complete-report PDFs retain accepted bytes/comparisons/reviews. An initial
unqualified probe built root by mistake; source moved to the isolated worktree and
17 accepted root assets were restored by SHA before qualification. The root suite
was not replayed. Evidence: `artifacts/composition-c1/numbered-display-evidence.json`.


Inserted text beside unchanged equations now shares source ownership between the
reader and preparation policy. Four containers retain all equations in both modes;
source-reference identity prevents equal-looking or nontext revisions from
qualifying. Deletions in Final retain their existing filtered path. Deletion
markup still has a confirmed balloon/frame mismatch and missing fraction; broad
admission was evaluated and rejected. Mark/property/move/equation revisions retain
explicit fallback. No font/size/document eligibility branch was added.

One frozen artifact passes 20 Windows methods once, 16 Ubuntu checks (four
native-font skips), two causal parent regressions and 17-asset/four-PDF transfer.
Twelve supported Word outcomes retain all 18 math glyphs within 0.877pt X and
0.768pt Y. All sixteen numerical page gates still fail unchanged thresholds.
Four inserted-markup pages were reviewed independently; inherited red/underlined
list labels and change-bar position remain defects. Twelve prose/simple-math
screens have no deficits, while generic revision flags and unused-part warnings
remain unresolved. Four deletion-markup controls retain actual deficits.
Eight fixed DOCX/two complete-report PDFs retain accepted bytes and comparisons;
no root rebuild or suite replay. Evidence: math-insertion-evidence.json.


Unchanged list labels and separators now inherit the first text run's formatting
before its inline revision decoration. Numbering-level overrides and runtime
font/spacing decisions remain intact; paragraph/mark revision owners keep their
existing path. Source tests cover four containers, both modes, bookmarks/run
splitting, authored formatting and level overrides.

One artifact passes 20 Windows methods once, 15 Ubuntu checks/five native-font
skips, two causal parent regressions and 17-asset/four-PDF transfer. All sixteen
math/text geometry controls retain C49 values; four Final insertion rasters are
identical despite separator coalescing in the PDF. Four markup insertion pages
were reviewed against Word and label styling improves. All sixteen numerical
gates still fail; change-bar position and deletion balloon/content defects remain.
Eight fixed DOCX/two complete-report PDFs retain accepted bytes/comparisons. A
wrong historical manifest also ran fourteen unnecessary spacing conversions;
those are excluded from acceptance and recorded. No root rebuild or suite replay.
Evidence: list-label-owner-evidence.json.


Correction to the deletion-markup investigation: the composition reference helper
left Word's RevisionsMode at 0 (balloons), while the library declares inline mode
1. The earlier 77–85pt comparison therefore does not establish a library frame
defect; that claim is withdrawn. Equation omission remains confirmed. Four new
independent deletion review references use the declared inline view; old exports
remain as provenance. View-agnostic historical comparisons are retained.

The helper now explicitly selects all six view settings, verifies observed export
logs and pins settings/log hashes. Cached source/PDF/view provenance is preflighted
before starting Office. Legacy pins require a matching retained export log;
mismatches reject without overwriting or global re-export. Eight tool checks pass
without Office or library builds. Four parent candidates reuse C50 PDF hashes,
with all four numerical gates still failed. The existing frozen library is reused;
no library rebuild or repeated suite. Evidence: reference-view-evidence.json.


Source-owned deletion text now shares the existing insertion composition path.
All sixteen revision controls retain unchanged formulas across body/cells/static
stories; mark/property/move/equation revisions retain fallback. Four affected
review outcomes use independently exported Word inline-mode references. No frame
correction or font/size/document eligibility branch was needed.

One artifact passes 24 Windows methods once, 20 Ubuntu checks/four native-font
skips, two causal parent regressions and 17-asset/four-PDF transfer. Four full-page
pairs were reviewed; frozen PDFs equal reviewed prototypes. Math residuals are
at most 0.877pt X/0.768pt Y. All four numerical gates still fail, though MAE and
recall improve. Prose/simple-math screens have no deficits; generic revision flags
and unreferenced-part warnings remain unresolved. The shared screen functions
verify an exact public-manifest subset; the whole-manifest CLI cannot accept that
subset and its rejected invocation is not an accepted report.

Only four affected review outputs were converted. Five fixed/complete inputs are
verified to contain no deletion elements, so their C50 comparisons/reviews carry
forward. Complete-report modes, memo and deck retain transfer-smoke PDF hashes.
No root rebuild or broad suite/corpus replay. Evidence: deletion-inline-evidence.json.


Revision bars now use the authored margin and printed-page scale. Their distance
from text is capped at 36pt, with centered placement in narrower margins and
scaled 0.72pt nominal thickness. Eight independent controls cover 36–144pt left
margins, right-margin/orientation changes and a fitted two-page report. Maximum
horizontal/thickness residuals are 0.024/0.018pt; vertical extents remain approximate.

One artifact passes 23 Windows methods once, 21 Ubuntu checks/two native-font
skips, two causal parent regressions and 17-asset/four-PDF transfer. Twelve text
inspections retain C52 bytes. Eight prose/simple-math screens have no deficits;
generic revision/unused-part warnings remain unresolved. Five representative
page pairs were reviewed. All twelve outcome gates still fail unchanged
thresholds. The composite report retains inherited comment-color and typography/geometry
differences, along with typography/geometry residuals. Fixed revision-free
documents carry accepted comparisons after a source audit and transfer smoke.

An initial unqualified probe built root CLI once and selected the old worktree
binary. Accepted root assets were restored by SHA with four smoke PDFs; the
intended worktree was rebuilt using an absolute project path. The wrong-stage
output is verified as C52 baseline only. No root suite replay. Evidence:
revision-bar-evidence.json and initial-prototype-correction.json.


Qualification builds now require an explicit committed linked worktree, full
revision and fresh artifact path. The reusable command builds absolute test/CLI
project paths, checks source identity between builds and delegates snapshot
publication to the existing freeze command. Ten fake-command checks cover main
checkout/subdirectory rejection, paths with spaces, revision/dirty source, output
reuse/location, build failure, source changes and successful publication. A real
main-checkout preflight rejects before builds and retains all seventeen accepted
assets. The C53 library artifact is reused; no library builds/tests were repeated.
The next rendering batch will exercise the real guarded build.

The composite author probe has identical Word rasters under revision modes0/1.
Both Word and candidate gray-lane samples are RGB242: the earlier gray-color
difference wording is withdrawn. The pink/blue comment author-color discrepancy
remains confirmed; source uses differently cased insertion/comment reviewer
names. Evidence: build-workspace-guard-evidence.json and review-author-view-probe.json.


Comment colors now register source-owned insertion/deletion authors before
comment authors, preserving exact name case and the qualified comment-ID order
and twenty-color cycle. Six Word controls distinguish author identity from anchor
and alphabetical order, including deletions. Word revision modes0/1 yield identical
baseline rasters. Balloon/range fills and strokes share the corrected slots; the
legacy revision hash keeps its normalization. Other revision-author classes and
inline multi-author ink remain approximate. The existing source paragraph walk
is shared with font purposes without changing its order.

One artifact passes25 Windows methods once,21 Ubuntu checks/four native-font
skips, two causal parent regressions and17-asset/four-PDF transfer. All six text
inspections retain parent bytes; two unchanged-color controls retain entire PDF
bytes. Eight fixed DOCX/two complete-report PDFs retain accepted bytes/comparisons.
Six prose/simple-math screens have no deficits; five generic revision flags stay
unresolved. All six numerical gates still fail unchanged thresholds. Three page
pairs reviewed; comment colors match while geometry/typography residuals remain.

The guarded command successfully built/froze the linked worktree from a main
checkout caller. An initial compile error stopped before freezing; a comparison
helper was restored to its font-plan owner before the accepted build. No root
build or suite replay. Evidence: review-author-evidence.json.
