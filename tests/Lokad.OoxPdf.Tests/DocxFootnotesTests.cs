using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxFootnotesTests
{

    public static void DocxPlainFootnoteKerningSharesContentThresholdsAndCompleteOwners()
    {
        foreach ((double size, string? minimum, bool enabled, bool nominal) in new (double, string?, bool, bool)[]
        {
            (10d, null, false, true), (12d, "0", false, true), (14d, "0", false, true),
            (10d, "20", true, true), (12d, "24", true, false), (14d, "28", true, true),
            (10d, "21", false, true), (12d, "25", false, true), (14d, "29", false, true),
        })
        {
            DocxDocument source = WithPlainFootnoteTestFooter(ReadWrappedFootnoteFlowFixture(100));
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    Runs = [p.Runs[0], p.Runs[1] with { FontSize = size, KerningMinimumHalfPointsValue = minimum,
                        PairKerningEnabled = !enabled }],
                })).ToArray(),
            } : s).ToArray() };
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var measurer = new LogicalBoundaryTextMeasurer();
            DocxLayout first = engine.Create(source, measurer, default);
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, first, out var adjusted, out var pages, default,
                DocxRelatedStoryKind.Footnote), "Layout-owned footnote frames share the authored note-content threshold policy.");
            TestAssert.True(pages.Values.Any(p => p.Slices > 1), "The control exercises actual continued notes.");
            TestAssert.True(ReferenceEquals(source.BodyElements, adjusted.BodyElements), "Body references are not rewritten.");
            foreach (var pair in source.RelatedStories.Zip(adjusted.RelatedStories).Where(p => p.First.IsNormalStoryType))
            {
                TestAssert.True(ReferenceEquals(pair.First.Paragraphs[0].Runs[0], pair.Second.Paragraphs[0].Runs[0]),
                    "Automatic marker runs retain exact source ownership and emission policy.");
                DocxTextRun oldRun = pair.First.Paragraphs[0].Runs[1], newRun = pair.Second.Paragraphs[0].Runs[1];
                TestAssert.Equal(oldRun with { PairKerningEnabled = enabled, UseLogicalSizeForPdfAdvances = nominal }, newRun);
                TestAssert.Equal(!enabled, oldRun.PairKerningEnabled);
            }
            DocxLayout final = engine.Create(adjusted, measurer, default);
            TestAssert.True(DocxNoteKerning.HasStableOwners(final, pages, default, DocxRelatedStoryKind.Footnote),
                "Changed wrapping retains complete, ordered source lines across rebased continuation slices.");
            foreach (var group in final.Pages.SelectMany(p => p.PlacedRelatedStories).Where(p => p.StoryLayout.Story.IsNormalStoryType)
                .GroupBy(p => p.StoryLayout.Story))
            {
                DocxRelatedStoryLayout canonical = final.RelatedStories.Single(s => ReferenceEquals(s.Story, group.Key));
                TestAssert.True(canonical.TextLines.Select(l => (l.SourceParagraph, l.SourceLineIndex, l.Text))
                    .SequenceEqual(group.SelectMany(p => p.StoryLayout.TextLines).Select(l => (l.SourceParagraph, l.SourceLineIndex, l.Text))),
                    "Every canonical line is emitted once, in source order.");
            }
        }
    }

    public static void DocxPlainFootnoteKerningRetainsIncompleteAndUnqualifiedFallbacks()
    {
        DocxDocument source = WithPlainFootnoteTestFooter(ReadWrappedFootnoteFlowFixture(100));
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new LogicalBoundaryTextMeasurer();
        DocxLayout layout = engine.Create(source, measurer, default);
        TestAssert.True(DocxNoteKerning.TryCreateDocument(source, layout, out _, out var owners, default, DocxRelatedStoryKind.Footnote),
            "The control has qualified source-owned footnote continuations.");
        foreach (DocxLayout invalid in new[]
        {
            layout with { Pages = layout.Pages.Select(p => p with { PlacedRelatedStories = p.PlacedRelatedStories
                .Select(s => s with { UsesPlainFootnoteFrame = false }).ToArray() }).ToArray() },
            layout with { Pages = layout.Pages.Take(layout.Pages.Count - 1).ToArray() },
            layout with { Pages = layout.Pages.Concat(layout.Pages.Take(1)).ToArray() },
        })
        {
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, invalid, out var unchanged, out _, default,
                DocxRelatedStoryKind.Footnote) && ReferenceEquals(source, unchanged),
                "Unqualified, missing or duplicate slices retain the original document.");
            TestAssert.True(!DocxNoteKerning.HasStableOwners(invalid, owners, default, DocxRelatedStoryKind.Footnote),
                "Incomplete source coverage cannot qualify the retry.");
        }
        foreach (string? fallback in new string?[] { "broken", "custom", "legacy" })
        {
            DocxDocument variant = fallback == "legacy" ? ReadWrappedFootnoteFlowFixture(100) : source with
            {
                RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
                {
                    BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    {
                        Runs = fallback == "custom" ? [p.Runs[0] with { IsAutomaticNoteStoryMarker = false }, p.Runs[1]] :
                            [p.Runs[0], p.Runs[1] with { KerningMinimumHalfPointsValue = fallback }],
                    })).ToArray(),
                } : s).ToArray(),
            };
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(variant, engine.Create(variant, measurer, default), out var unchanged,
                out _, default, DocxRelatedStoryKind.Footnote) && ReferenceEquals(variant, unchanged),
                "Malformed thresholds, custom marks and legacy frames retain their existing fallbacks.");
        }
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxNoteKerning.TryCreateDocument(source, layout, out _, out _,
            cancelled.Token, DocxRelatedStoryKind.Footnote));
    }

    private static DocxDocument WithPlainFootnoteTestHeader(DocxDocument source, bool centered = false)
    {
        DocxParagraph header = DocxTests.CreateDocxLayoutParagraph("public plain header", 12d, 14d) with
        {
            Alignment = centered ? DocxTextAlignment.Center : DocxTextAlignment.Left,
            LineSpacingPoints = null, LineSpacingFactor = 1d,
            Spacing = DocxParagraphSpacing.Empty with { LineRuleValue = "auto", LineValue = "240" },
        };
        var paragraphs = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [header] };
        var elements = new Dictionary<string, IReadOnlyList<DocxBodyElement>> { ["default"] = [new DocxParagraphElement(header)] };
        var settings = source.PageSettings with { HeaderDistancePoints = 36d,
            HeaderParagraphsByType = paragraphs, HeaderBodyElementsByType = elements };
        return source with { PageSettings = settings, FinalSectionBreak = source.FinalSectionBreak! with { PageSettings = settings },
            HeaderParagraphs = [header], HeaderParagraphsByType = paragraphs, HeaderBodyElementsByType = elements };
    }

    private static DocxDocument WithPlainFootnoteTestFooter(DocxDocument source)
    {
        DocxParagraph footer = DocxTests.CreateDocxLayoutParagraph("public footer frame", 12d, 14d) with
        {
            LineSpacingPoints = null,
            Spacing = DocxParagraphSpacing.Empty with { BeforeValue = "0", AfterValue = "0", LineValue = "240", LineRuleValue = "auto" },
        };
        var paragraphs = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [footer] };
        var elements = new Dictionary<string, IReadOnlyList<DocxBodyElement>> { ["default"] = [new DocxParagraphElement(footer)] };
        var settings = source.PageSettings with { FooterDistancePoints = 36d,
            FooterParagraphsByType = paragraphs, FooterBodyElementsByType = elements };
        return source with { PageSettings = settings, FinalSectionBreak = source.FinalSectionBreak! with { PageSettings = settings },
            FooterParagraphs = [footer], FooterParagraphsByType = paragraphs, FooterBodyElementsByType = elements };
    }

    public static void DocxAutomaticNoteMarkerSizeDoesNotInflateContentPitchOrInset()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool footnote in new[] { false, true })
        foreach (double contentSize in new[] { 10d, 12d, 14d })
        {
            DocxDocument source = footnote ? WithPlainFootnoteTestFooter(ReadWrappedFootnoteFlowFixture(100)) :
                ReadWrappedDocumentEndnoteFlowFixture(0, 0, 100);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    { Runs = [p.Runs[0] with { FontSize = 12d }, p.Runs[1] with { FontSize = contentSize }] })).ToArray(),
            } : s).ToArray() };
            DocxLayout layout = engine.Create(source, measurer, default);
            foreach (DocxRelatedStoryLayout story in layout.RelatedStories.Where(s => s.Story.IsNormalStoryType))
            {
                TestAssert.True(story.TextLines.Count > 2, "The control exercises repeated content line boxes.");
                TestAssert.True(story.TextLines.Skip(1).All(l => Math.Abs(l.LineHeight!.Value - contentSize * 1.2d) < .000001d),
                    "An automatic superscript marker does not set the nominal pitch of the content paragraph.");
                TestAssert.True(Math.Abs(story.TextLines[0].LineHeight!.Value - Math.Max(12d, contentSize) * 1.2d) < .000001d,
                    "The marked first line retains its original source metric box.");
                if (contentSize == 10d)
                    TestAssert.True(Math.Abs(story.TextLines[0].BaselineY + 12d * .94d) < .000001d &&
                        Math.Abs(story.TextLines[0].BaselineY - story.TextLines[1].BaselineY - 13.12d) < .000001d,
                        "The first-to-continuation advance accounts for both source-owned baseline insets.");
                TestAssert.True(story.TextLines.All(l => ReferenceEquals(l.SourceParagraph, story.Story.Paragraphs[0])),
                    "The metric correction retains the source paragraph and marker run.");
                TestAssert.Equal(12d, story.Story.Paragraphs[0].Runs[0].FontSize);
            }
        }
    }

    public static void DocxSingleLineAutomaticNoteMarkersRetainQualifiedMetrics()
    {
        DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(0, 0, 1);
        DocxDocument Compose(bool automatic) => source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
        {
            BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                { Runs = [p.Runs[0] with { FontSize = 12d, IsAutomaticNoteStoryMarker = automatic }, p.Runs[1] with { FontSize = 10d }] })).ToArray(),
        } : s).ToArray() };
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        DocxLayout marked = engine.Create(Compose(true), measurer, default), authored = engine.Create(Compose(false), measurer, default);
        TestAssert.True(marked.RelatedStories.Where(s => s.Story.IsNormalStoryType).All(s => s.TextLines.Count == 1),
            "The control exercises fitting marked lines without a continuation.");
        TestAssert.True(marked.RelatedStories.SelectMany(s => s.TextLines).Select(l => (l.BaselineY, l.LineHeight))
            .SequenceEqual(authored.RelatedStories.SelectMany(s => s.TextLines).Select(l => (l.BaselineY, l.LineHeight))),
            "A fitting single-line note retains its independently qualified marker and content box.");
    }

    public static void DocxSubordinateAutomaticNoteMarkersRetainQualifiedMixedInset()
    {
        DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(0, 0, 100);
        DocxDocument Compose(bool automatic) => source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
        {
            BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                { Runs = [p.Runs[0] with { FontSize = 12d, IsAutomaticNoteStoryMarker = automatic }, p.Runs[1] with { FontSize = 14d }] })).ToArray(),
        } : s).ToArray() };
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        DocxLayout marked = engine.Create(Compose(true), measurer, default), authored = engine.Create(Compose(false), measurer, default);
        TestAssert.True(marked.RelatedStories.SelectMany(s => s.TextLines).Select(l => (l.BaselineY, l.LineHeight))
            .SequenceEqual(authored.RelatedStories.SelectMany(s => s.TextLines).Select(l => (l.BaselineY, l.LineHeight))),
            "A subordinate marker retains the independently qualified mixed-content inset and take geometry.");
    }

    public static void DocxAuthoredScriptAndBaselineNoteMarkersRetainMetricOwnership()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool baseline in new[] { false, true })
        {
            DocxDocument source = WithPlainFootnoteTestFooter(ReadWrappedFootnoteFlowFixture(100));
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    Runs = [p.Runs[0] with { FontSize = 12d, IsAutomaticNoteStoryMarker = baseline,
                        VerticalAlignmentValue = baseline ? "baseline" : "superscript" }, p.Runs[1] with { FontSize = 10d }],
                })).ToArray(),
            } : s).ToArray() };
            DocxLayout layout = engine.Create(source, new NoteContentMetricMeasurer(), default);
            TestAssert.True(layout.RelatedStories.Where(s => s.Story.IsNormalStoryType).SelectMany(s => s.TextLines)
                .All(l => Math.Abs(l.LineHeight!.Value - 14.4d) < .000001d),
                "Authored superscript text and baseline note markers retain their existing nominal metric contribution.");
        }
    }

    public static void DocxMinimumNoteLinesHonorBoundsAndSourceMetricOwnership()
    {
        foreach (bool footnote in new[] { true, false })
        foreach (double minimum in new[] { 8d, 12d, 18d })
        {
            DocxDocument source = footnote ? ReadWrappedFootnoteFlowFixture(100) : ReadWrappedDocumentEndnoteFlowFixture(0, 0, 100);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    LineSpacingPoints = minimum, Spacing = p.Spacing with { LineRuleValue = "atLeast",
                        LineValue = (minimum * 20d).ToString(CultureInfo.InvariantCulture) },
                    Runs = [p.Runs[0] with { FontSize = 12d }, p.Runs[1] with { FontSize = 10d }],
                })).ToArray(),
            } : s).ToArray() };
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new NoteContentMetricMeasurer(), default);
            foreach (var story in layout.RelatedStories.Where(s => s.Story.IsNormalStoryType))
            {
                TestAssert.True(story.TextLines.Count > 2, "The minimum control contains actual unmarked continuation lines.");
                TestAssert.True(Math.Abs(story.TextLines[0].LineHeight!.Value - Math.Max(minimum, 14.4d)) < .000001d,
                    "The marked first line retains its source metric box and authored minimum.");
                TestAssert.True(story.TextLines.Skip(1).All(l => Math.Abs(l.LineHeight!.Value - Math.Max(minimum, 12d)) < .000001d),
                    "Minimum note continuation lines use content metrics and retain the authored lower bound.");
                TestAssert.True(Math.Abs(story.ContentHeight + story.TextLines[^1].BaselineY - 2d) < .000001d,
                    "Growing a minimum preserves the content's two-point bottom bearing.");
                TestAssert.True(story.TextLines.All(l => ReferenceEquals(l.SourceParagraph, story.Story.Paragraphs[0]) &&
                    l.SourceParagraph!.LineSpacingPoints == minimum), "Metric derivation preserves authored paragraph and minimum ownership.");
            }
        }
    }

    public static void DocxMinimumFootnoteFramesRetainCompleteContinuationOwners()
    {
        DocxDocument source = WithPlainFootnoteTestFooter(ReadWrappedFootnoteFlowFixture(100));
        source = source with
        {
            BodyElements = source.BodyElements.OfType<DocxParagraphElement>().Take(1).ToArray(),
            RelatedStories = source.RelatedStories.Where(s => !s.IsNormalStoryType || s.Id == "37").Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    LineSpacingPoints = 18d, Spacing = p.Spacing with { LineRuleValue = "atLeast", LineValue = "360" },
                    Runs = [p.Runs[0], p.Runs[1] with { FontSize = 10d }],
                })).ToArray(),
            } : s).ToArray(),
        };
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        DocxLayout layout = engine.Create(source, measurer, default);
        var slices = layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => s.StoryLayout.Story.IsNormalStoryType).ToArray();
        TestAssert.True(slices.Length > 1 && slices.All(s => s.UsesPlainFootnoteFrame),
            "The shared minimum footnote frame reaches the continuation drain with explicit frame ownership.");
        TestAssert.Equal(34, slices[0].TextLines.Count);
        var canonical = layout.RelatedStories.Single(s => s.Story.IsNormalStoryType);
        TestAssert.True(canonical.TextLines.Select(l => (l.Text, l.SourceParagraph, l.SourceLineIndex)).SequenceEqual(
            slices.SelectMany(s => s.StoryLayout.TextLines).Select(l => (l.Text, l.SourceParagraph, l.SourceLineIndex))),
            "Every source minimum line is placed exactly once and in order.");
        TestAssert.True(DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out var owners, default,
            DocxRelatedStoryKind.Footnote), "Minimum frames share the existing source-owned content kerning policy.");
        TestAssert.True(DocxNoteKerning.HasStableOwners(engine.Create(adjusted, measurer, default), owners, default,
            DocxRelatedStoryKind.Footnote), "The content retry retains complete minimum-flow owners.");
        TestAssert.True(ReferenceEquals(source.BodyElements, adjusted.BodyElements), "Body reference owners are retained.");
    }

    public static void DocxMinimumNoteBoxesPreserveSingleAndWrappedBottomBearings()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool footnote in new[] { false, true })
        foreach (int repeats in new[] { 1, 100 })
        foreach (double contentSize in new[] { 10d, 12d, 14d })
        foreach (double minimum in new[] { 8d, 18d, 24d })
        {
            DocxDocument source = footnote ? ReadWrappedFootnoteFlowFixture(repeats) :
                ReadWrappedDocumentEndnoteFlowFixture(0, 0, repeats);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    { Runs = [p.Runs[0] with { FontSize = 12d }, p.Runs[1] with { FontSize = contentSize }] })).ToArray(),
            } : s).ToArray() };
            DocxDocument bounded = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    LineSpacingPoints = minimum, Spacing = p.Spacing with { LineRuleValue = "atLeast",
                        LineValue = (minimum * 20d).ToString(CultureInfo.InvariantCulture) },
                })).ToArray(),
            } : s).ToArray() };
            var natural = engine.Create(source, measurer, default).RelatedStories.Where(s => s.Story.IsNormalStoryType).ToArray();
            var actual = engine.Create(bounded, measurer, default).RelatedStories.Where(s => s.Story.IsNormalStoryType).ToArray();
            TestAssert.Equal(natural.Length, actual.Length);
            foreach (var pair in natural.Zip(actual))
            {
                TestAssert.True(repeats == 1 ? pair.First.TextLines.Count == 1 : pair.First.TextLines.Count > 2,
                    "The controls exercise fitting marked lines and wrapped unmarked content.");
                TestAssert.Equal(pair.First.TextLines.Count, pair.Second.TextLines.Count);
                double naturalBottom = 0d, boundedBottom = 0d;
                foreach (var line in pair.First.TextLines.Zip(pair.Second.TextLines))
                {
                    naturalBottom += line.First.LineHeight!.Value; boundedBottom += line.Second.LineHeight!.Value;
                    TestAssert.True(line.Second.LineHeight!.Value >= minimum &&
                        Math.Abs(line.First.BaselineY + naturalBottom - line.Second.BaselineY - boundedBottom) < .000001d,
                        "Binding minimum boxes grow above the same source-owned bottom bearing, including a marked single line.");
                    TestAssert.True(ReferenceEquals(line.Second.SourceParagraph, pair.Second.Story.Paragraphs[0]),
                        "Minimum metric ownership retains the authored source paragraph.");
                }
                TestAssert.True(ReferenceEquals(pair.First.Story.Paragraphs[0].Runs[0], pair.Second.Story.Paragraphs[0].Runs[0]),
                    "The XML automatic marker retains its original source run.");
            }
        }
    }

    public static void DocxMinimumNoteBearingsRetainAuthoredRichAndScaledFallbacks()
    {
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool footnote in new[] { false, true })
        foreach (double scale in new[] { 1d, .5d })
        foreach (string variant in new[] { "automatic", "authored", "baseline", "rich" })
        {
            if (scale == 1d && variant == "automatic") continue;
            DocxDocument source = footnote ? ReadWrappedFootnoteFlowFixture(1) : ReadWrappedDocumentEndnoteFlowFixture(0, 0, 1);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    Alignment = variant == "rich" ? DocxTextAlignment.Center : p.Alignment,
                    Runs = [p.Runs[0] with { IsAutomaticNoteStoryMarker = variant != "authored",
                        VerticalAlignmentValue = variant == "baseline" ? "baseline" : "superscript" }, p.Runs[1]],
                })).ToArray(),
            } : s).ToArray() };
            DocxDocument bounded = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    { LineSpacingPoints = 24d, Spacing = p.Spacing with { LineRuleValue = "atLeast", LineValue = "480" } })).ToArray(),
            } : s).ToArray() };
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, scale);
            if (variant == "automatic") source = bounded with { RelatedStories = bounded.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    { Runs = [p.Runs[0] with { IsAutomaticNoteStoryMarker = false }, p.Runs[1]] })).ToArray(),
            } : s).ToArray() };
            var natural = engine.Create(source, measurer, default).RelatedStories.Where(s => s.Story.IsNormalStoryType).ToArray();
            var actual = engine.Create(bounded, measurer, default).RelatedStories.Where(s => s.Story.IsNormalStoryType).ToArray();
            TestAssert.Equal(natural.Length, actual.Length);
            TestAssert.True(natural.Zip(actual).All(pair => Math.Abs(pair.First.TextLines[0].BaselineY -
                pair.Second.TextLines[0].BaselineY) < .000001d),
                $"Authored superscripts, baseline markers, rich paragraphs and scaled stories retain their first-inset fallback ({footnote}, {scale}, {variant}).");
        }
    }

    public static void DocxExactNoteBoxesShareFixedBaselineWithoutNaturalExpansion()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool footnote in new[] { false, true })
        foreach (int repeats in new[] { 1, 100 })
        foreach (double contentSize in new[] { 10d, 12d, 14d })
        foreach (double fixedHeight in new[] { 8d, 14d, 24d })
        {
            DocxDocument source = footnote ? ReadWrappedFootnoteFlowFixture(repeats) :
                ReadWrappedDocumentEndnoteFlowFixture(0, 0, repeats);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    LineSpacingPoints = fixedHeight, Spacing = p.Spacing with { LineRuleValue = "exact",
                        LineValue = (fixedHeight * 20d).ToString(CultureInfo.InvariantCulture) },
                    Runs = [p.Runs[0] with { FontSize = 12d }, p.Runs[1] with { FontSize = contentSize }],
                })).ToArray(),
            } : s).ToArray() };
            foreach (var story in engine.Create(source, measurer, default).RelatedStories.Where(s => s.Story.IsNormalStoryType))
            {
                TestAssert.True(repeats == 1 ? story.TextLines.Count == 1 : story.TextLines.Count > 2,
                    "The controls exercise marked single lines and unmarked wrapped content.");
                TestAssert.True(Math.Abs(story.TextLines[0].BaselineY + fixedHeight * .8d) < .000001d,
                    "Plain exact notes use the independently qualified fixed-line first baseline, including uniform runs.");
                foreach (var line in story.TextLines)
                {
                    TestAssert.Equal(fixedHeight, line.LineHeight!.Value);
                    TestAssert.True(line.LineHeightSource == DocxLineHeightSource.ExactLineSpacing &&
                        ReferenceEquals(line.SourceParagraph, story.Story.Paragraphs[0]),
                        "Exact boxes retain fixed clipping and original source paragraphs without natural metric growth.");
                }
                TestAssert.True(story.TextLines.Zip(story.TextLines.Skip(1)).All(p =>
                    Math.Abs(p.First.BaselineY - p.Second.BaselineY - fixedHeight) < .000001d),
                    "Automatic markers and differing content sizes do not alter the fixed continuation pitch.");
            }
        }
    }

    public static void DocxExactNoteTypographyPreservesCompleteOwnersAndFixedSpacing()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool footnote in new[] { false, true })
        foreach (string threshold in new[] { "0", "20", "30" })
        {
            DocxDocument source = footnote ? WithPlainFootnoteTestFooter(ReadWrappedFootnoteFlowFixture(200)) :
                ReadWrappedDocumentEndnoteFlowFixture(8, 4, 200);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    LineSpacingPoints = 14d, Spacing = p.Spacing with { LineRuleValue = "exact", LineValue = "280" },
                    Runs = [p.Runs[0], p.Runs[1] with { FontSize = 10d, KerningMinimumHalfPointsValue = threshold,
                        PairKerningEnabled = threshold != "20" }],
                })).ToArray(),
            } : s).ToArray() };
            DocxLayout layout = engine.Create(source, measurer, default);
            TestAssert.True(layout.Pages.Count > 1, "The fixed-box control includes real note continuations.");
            var kind = footnote ? DocxRelatedStoryKind.Footnote : DocxRelatedStoryKind.Endnote;
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out var owners, default, kind),
                "Plain fixed-box notes share the source-owned absent, met and unmet threshold policy.");
            TestAssert.True(ReferenceEquals(source.BodyElements, adjusted.BodyElements), "Body reference owners remain unchanged.");
            foreach (var pair in source.RelatedStories.Zip(adjusted.RelatedStories).Where(p => p.First.IsNormalStoryType))
            {
                var before = pair.First.Paragraphs[0]; var after = pair.Second.Paragraphs[0];
                TestAssert.True(ReferenceEquals(before.Runs[0], after.Runs[0]), "The XML automatic marker remains immutable.");
                TestAssert.Equal(before.Spacing, after.Spacing);
                TestAssert.Equal(14d, after.LineSpacingPoints!.Value);
                TestAssert.Equal(threshold == "20", after.Runs[1].PairKerningEnabled);
            }
            TestAssert.True(DocxNoteKerning.HasStableOwners(engine.Create(adjusted, measurer, default), owners, default, kind),
                "Every canonical fixed-box line retains complete source ownership through the bounded typography retry.");
        }
    }

    public static void DocxExactNoteMetricsRetainRichAuthoredAndScaledFallbacks()
    {
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool footnote in new[] { false, true })
        foreach (double scale in new[] { 1d, .5d })
        foreach (string variant in new[] { "automatic", "authored", "baseline", "rich", "contextual", "keep", "malformed" })
        {
            if (scale == 1d && variant == "automatic") continue;
            DocxDocument source = footnote ? ReadWrappedFootnoteFlowFixture(1) : ReadWrappedDocumentEndnoteFlowFixture(0, 0, 1);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    LineSpacingPoints = 24d, Spacing = p.Spacing with { LineRuleValue = "exact", LineValue = "480",
                        ContextualSpacing = variant == "contextual", BeforeValue = variant == "malformed" ? "broken" : "0" },
                    KeepRules = p.KeepRules with { KeepNext = variant == "keep" },
                    Alignment = variant == "rich" ? DocxTextAlignment.Center : p.Alignment,
                    Runs = [p.Runs[0] with { IsAutomaticNoteStoryMarker = variant != "authored",
                        VerticalAlignmentValue = variant == "baseline" ? "baseline" : "superscript" }, p.Runs[1]],
                })).ToArray(),
            } : s).ToArray() };
            DocxDocument authored = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    { Runs = [p.Runs[0] with { IsAutomaticNoteStoryMarker = false }, p.Runs[1]] })).ToArray(),
            } : s).ToArray() };
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, scale);
            var before = engine.Create(authored, measurer, default).RelatedStories.Where(s => s.Story.IsNormalStoryType).ToArray();
            var after = engine.Create(source, measurer, default).RelatedStories.Where(s => s.Story.IsNormalStoryType).ToArray();
            TestAssert.True(before.Zip(after).All(p => p.First.TextLines[0].BaselineY == p.Second.TextLines[0].BaselineY),
                "Rich/contextual/kept paragraphs, malformed spacing, authored markers and originally scaled notes retain the qualified inset fallback.");
        }
    }

    public static void DocxCustomNoteLineMetricsPreserveMinimumBottomBearingsAndSourceOwners()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool footnote in new[] { false, true })
        foreach (double contentSize in new[] { 10d, 12d, 14d })
        foreach (double minimum in new[] { 8d, 18d })
        {
            DocxDocument source = ReadCustomNoteMetricFixture(footnote, contentSize);
            DocxDocument bounded = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    LineSpacingPoints = minimum, Spacing = p.Spacing with { LineRuleValue = "atLeast",
                        LineValue = (minimum * 20d).ToString(CultureInfo.InvariantCulture) },
                })).ToArray(),
            } : s).ToArray() };
            DocxLayout natural = engine.Create(source, measurer, default), actual = engine.Create(bounded, measurer, default);
            foreach (var pair in natural.RelatedStories.Where(s => s.Story.IsNormalStoryType)
                .Zip(actual.RelatedStories.Where(s => s.Story.IsNormalStoryType)))
            {
                TestAssert.True(pair.First.TextLines.Count > 2, "The control exercises actual custom-mark continuations.");
                TestAssert.Equal(pair.First.TextLines.Count, pair.Second.TextLines.Count);
                double naturalBottom = 0d, boundedBottom = 0d;
                foreach (var line in pair.First.TextLines.Zip(pair.Second.TextLines))
                {
                    naturalBottom += line.First.LineHeight!.Value; boundedBottom += line.Second.LineHeight!.Value;
                    double expectedNatural = (line.First.IsFirstParagraphLine == true ? Math.Max(12d, contentSize) : contentSize) * 1.2d;
                    TestAssert.True(Math.Abs(line.First.LineHeight.Value - expectedNatural) < .000001d,
                        "Custom note lines use only the source runs present on that wrapped line.");
                    TestAssert.True(Math.Abs(line.Second.LineHeight!.Value - Math.Max(minimum, expectedNatural)) < .000001d &&
                        Math.Abs(line.First.BaselineY + naturalBottom - line.Second.BaselineY - boundedBottom) < .000001d,
                        "Minimum boxes preserve the source-owned bottom bearing for marked and unmarked lines.");
                    TestAssert.True(ReferenceEquals(line.Second.SourceParagraph, pair.Second.Story.Paragraphs[0]),
                        "Metric views do not replace canonical source paragraphs.");
                }
                TestAssert.True(ReferenceEquals(pair.First.Story.Paragraphs[0].Runs[0], pair.Second.Story.Paragraphs[0].Runs[0]),
                    "The custom prefix retains its exact source run.");
            }
            var kind = footnote ? DocxRelatedStoryKind.Footnote : DocxRelatedStoryKind.Endnote;
            TestAssert.True(DocxNoteKerning.TryCreateDocument(bounded, actual, out var adjusted, out var owners, default, kind) &&
                DocxNoteKerning.HasStableOwners(engine.Create(adjusted, measurer, default), owners, default, kind),
                "The shared advance retry retains every canonical custom-note owner after metric correction.");
        }
    }

    public static void DocxMixedStaticNoteTypographyPreservesComplementaryPlacementAndBoundaryPolicy()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool custom in new[] { false, true })
        foreach ((bool header, bool footer) in new[] { (true, false), (false, true), (true, true) })
        foreach (int count in new[] { 1, 6 })
        {
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32), header, footer, count);
            DocxLayout before = engine.Create(source, measurer, default);
            TestAssert.True(!DocxPlainBodyFrame.IsEligible(source) && before.Pages.Any(p => p.StaticTextLines.Count != 0),
                "The mixed control contains actual static stories outside the original empty-frame admission.");
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, before, out _, out _, default, DocxRelatedStoryKind.Footnote),
                "The established mixed footnote frame admission remains unchanged.");
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, before, out var adjusted, out var owners, default),
                "Complete plain complementary sources share endnote typography in a usable static frame.");
            DocxLayout after = engine.Create(adjusted, measurer, default);
            TestAssert.True(DocxNoteKerning.HasStableOwners(after, owners, default) &&
                ReferenceEquals(source.PageSettings, adjusted.PageSettings) &&
                source.BodyElements.Zip(adjusted.BodyElements).All(p => ReferenceEquals(p.First, p.Second)),
                "The bounded retry retains both typed source owners, body references and static settings.");
            var previousFoot = before.Pages.SelectMany((p, page) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                .Select(s => (page, s.TopY, s.Height, s.ContentTopOffset, s.StoryLayout.Story.Paragraphs[0])));
            var currentFoot = after.Pages.SelectMany((p, page) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                .Select(s => (page, s.TopY, s.Height, s.ContentTopOffset, s.StoryLayout.Story.Paragraphs[0])));
            TestAssert.True(previousFoot.SequenceEqual(currentFoot),
                "Fitting and overflowing static frames preserve every complementary footnote box and canonical paragraph.");
            static object Geometry(DocxTextLineLayout line) => (line.Text, line.X, line.BaselineY, line.Width, line.FontSize,
                line.SourceParagraph, line.SourceBlockIndex, line.SourceParagraphIndex, line.SourceLineIndex);
            TestAssert.True(before.Pages.Count == after.Pages.Count && before.Pages.Zip(after.Pages).All(p =>
                p.First.Items.OfType<DocxTextLineLayout>().Select(Geometry).SequenceEqual(p.Second.Items.OfType<DocxTextLineLayout>().Select(Geometry)) &&
                p.First.StaticTextLines.Select(Geometry).SequenceEqual(p.Second.StaticTextLines.Select(Geometry))),
                "The typography correction preserves body and static geometry.");
            TestAssert.True(after.Pages.SelectMany(p => p.PlacedRelatedStories)
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote)
                .All(s => s.UsesPlainDocumentEndnoteTypography && !s.UsesCompleteDocumentEndnoteFlow),
                "Shared typography does not admit mixed static paragraph-boundary or pagination corrections.");
        }
    }

    public static void DocxMixedStaticNoteTypographyRetainsSourceShapeAndFrameFallbacks()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        DocxDocument plain = ReadMixedNoteRetryFixture(true, 32);
        DocxRelatedStory foot = plain.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Footnote);
        DocxParagraph paragraph = foot.Paragraphs[0];
        foreach (DocxParagraph invalid in new[]
        {
            paragraph with { KeepRules = paragraph.KeepRules with { KeepNext = true } },
            paragraph with { KeepRules = paragraph.KeepRules with { KeepLines = true } },
            paragraph with { Spacing = paragraph.Spacing with { ContextualSpacing = true } },
            paragraph with { Spacing = paragraph.Spacing with { AfterValue = "invalid" } },
            paragraph with { Alignment = DocxTextAlignment.Center },
        })
        {
            DocxDocument source = WithMixedNoteStaticFrame(plain with { RelatedStories = plain.RelatedStories.Select(s =>
                ReferenceEquals(s, foot) ? s with { BodyElements = [new DocxParagraphElement(invalid)] } : s).ToArray() }, true, true, 1);
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, engine.Create(source, measurer, default), out var unchanged, out _, default) &&
                ReferenceEquals(source, unchanged), "Kept, contextual, malformed and rich complementary sources retain their fallback.");
        }
        DocxParagraph staticParagraph = DocxTests.CreateDocxLayoutParagraph("public static frame", 12d, 14d) with
        {
            LineSpacingPoints = null, Spacing = DocxParagraphSpacing.Empty with { BeforeValue = "0", AfterValue = "0", LineValue = "240", LineRuleValue = "auto" },
        };
        foreach (DocxDocument source in new[]
        {
            WithMixedNoteStaticFrame(plain, true, true, 100),
            WithMixedNoteStaticFrame(plain, true, true, 1, staticParagraph with { LineSpacingPoints = 14d }),
            WithMixedNoteStaticFrame(plain, true, true, 1, staticParagraph with
                { Runs = [staticParagraph.Runs[0] with { Text = "{NUMPAGES}", FieldKind = DocxFieldKind.NumPages }] }),
        })
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, engine.Create(source, measurer, default), out _, out _, default),
                "Consumed, fixed and dynamic frames retain their fallback.");
        DocxDocument scaled = WithMixedNoteStaticFrame(plain, true, true, 1);
        DocxLayout scaledLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, .5d)
            .Create(scaled, measurer, default, measurer);
        TestAssert.True(!DocxNoteKerning.TryCreateDocument(scaled, scaledLayout, out _, out _, default),
            "Originally scaled mixed frames retain their fallback.");
    }

    public static void DocxMixedFootnoteFrameClearsOverflowingFooterWithoutChangingTypographyOwner()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool custom in new[] { false, true })
        {
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32), false, true, 6);
            DocxLayout layout = engine.Create(source, measurer, default);
            var footnotes = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                .Select(s => (Page: page, Index: index, Story: s))).ToArray();
            TestAssert.Equal(1, footnotes.Length);
            var foot = footnotes[0];
            double footerTop = layout.FooterContentTopByPage[foot.Index];
            TestAssert.True(footerTop > foot.Page.MarginBottom && foot.Story.TopY - foot.Story.Height >= footerTop - .000001d,
                "A mixed footnote's complete physical box clears the selected overflowing footer edge.");
            TestAssert.True(foot.Story.UsesPlainFootnoteFrame && foot.Story.ContentTopOffset == 0d &&
                foot.Story.TextLines.All(l => ReferenceEquals(l.SourceParagraph, foot.Story.StoryLayout.Story.Paragraphs[0])),
                "The shared physical frame retains the complete canonical footnote source.");
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, layout, out _, out _, default, DocxRelatedStoryKind.Footnote),
                "Physical footer clearance retains the original mixed footnote typography fallback.");
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out var owners, default),
                "Physical footer clearance preserves the existing endnote owner of the single successful typography retry.");
            DocxLayout after = engine.Create(adjusted, measurer, default);
            TestAssert.True(DocxNoteKerning.HasStableOwners(after, owners, default),
                "Both typed note kinds retain complete canonical owners after the retry.");
            var placed = after.Pages[foot.Index].PlacedRelatedStories.Single(s => s.StoryLayout.Story.IsNormalStoryType &&
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote);
            TestAssert.True(placed.TopY == foot.Story.TopY && placed.Height == foot.Story.Height &&
                ReferenceEquals(placed.StoryLayout.Story.Paragraphs[0], foot.Story.StoryLayout.Story.Paragraphs[0]),
                "The endnote retry leaves the corrected complementary physical placement and source unchanged.");
            DocxDocument continued = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 140), false, true, 6);
            DocxLayout continuedLayout = engine.Create(continued, measurer, default);
            var continuedEnds = continuedLayout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote)
                .Select(s => (Index: index, Story: s))).ToArray();
            var continuedFoot = continuedLayout.Pages[foot.Index].PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote);
            TestAssert.True(continuedEnds.Length > 1 && continuedEnds[0].Index == foot.Index &&
                continuedEnds[0].Story.TopY - continuedEnds[0].Story.Height >= continuedFoot.TopY &&
                continuedEnds.All(s => s.Story.UsesPlainDocumentEndnoteTypography),
                "The continued head uses body space above the footnote and clears its occupied physical frame.");
            TestAssert.True(DocxNoteKerning.TryCreateDocument(continued, continuedLayout, out _, out _, default),
                "An admitted partial head retains the existing endnote typography owner.");
        }
    }

    public static void DocxMixedFootnoteFrameRetainsRichComplementarySourceFallbacks()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        DocxDocument plain = ReadMixedNoteRetryFixture(true, 32);
        DocxRelatedStory end = plain.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Endnote);
        DocxParagraph paragraph = end.Paragraphs[0];
        foreach (DocxParagraph invalid in new[]
        {
            paragraph with { KeepRules = paragraph.KeepRules with { KeepNext = true } },
            paragraph with { KeepRules = paragraph.KeepRules with { KeepLines = true } },
            paragraph with { Spacing = paragraph.Spacing with { ContextualSpacing = true } },
            paragraph with { Spacing = paragraph.Spacing with { AfterValue = "invalid" } },
            paragraph with { Alignment = DocxTextAlignment.Center },
        })
        {
            DocxDocument source = WithMixedNoteStaticFrame(plain with { RelatedStories = plain.RelatedStories.Select(s =>
                ReferenceEquals(s, end) ? s with { BodyElements = [new DocxParagraphElement(invalid)] } : s).ToArray() }, false, true, 6);
            DocxLayout layout = engine.Create(source, measurer, default);
            TestAssert.True(layout.Pages.SelectMany(p => p.PlacedRelatedStories)
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                .All(s => !s.UsesPlainFootnoteFrame), $"The shared frame requires plain flow sources for both note kinds ({invalid.KeepRules.KeepNext}, {invalid.KeepRules.KeepLines}, {invalid.Spacing.ContextualSpacing}, {invalid.Spacing.AfterValue}, {invalid.Alignment}).");
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, layout, out _, out _, default, DocxRelatedStoryKind.Footnote),
                "Rich complementary sources retain the original mixed footnote typography fallback.");
        }
    }

    public static void DocxNoteSeparatorsUseTypedSettingsRegistration()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string registration in new[] { "none", "normal", "continuation", "other-kind", "wrong-id" })
        {
            DocxDocument source = ReadSectionNumberingFixture(kind, (_, settings, parts) =>
            {
                XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
                foreach (XElement special in notes.Root!.Elements(NoteWord + kind).Where(n =>
                    (string?)n.Attribute(NoteWord + "type") is "separator" or "continuationSeparator"))
                {
                    bool normal = (string?)special.Attribute(NoteWord + "type") == "separator";
                    special.ReplaceNodes(new XElement(NoteWord + "p",
                        new XElement(NoteWord + "pPr", new XElement(NoteWord + "spacing",
                            new XAttribute(NoteWord + "after", "0"), new XAttribute(NoteWord + "line", "240"),
                            new XAttribute(NoteWord + "lineRule", "auto"))),
                        new XElement(NoteWord + "r", new XElement(NoteWord + "rPr",
                            new XElement(NoteWord + "rFonts", new XAttribute(NoteWord + "ascii", "Arial"),
                                new XAttribute(NoteWord + "hAnsi", "Arial")),
                            new XElement(NoteWord + "sz", new XAttribute(NoteWord + "val", normal ? "16" : "32"))),
                            new XElement(NoteWord + "t", normal ? "Public normal separator" : "Public continuation separator"))));
                }
                parts["word/" + kind + "s.xml"] = notes.ToString();
                XElement properties = settings.Root!.Element(NoteWord + (kind + "Pr")) ??
                    new XElement(NoteWord + (kind + "Pr"));
                if (properties.Parent is null) settings.Root.Add(properties);
                properties.Elements(NoteWord + kind).Remove();
                string selected = registration == "continuation" ? "continuationSeparator" : "separator";
                string id = (string)notes.Root.Elements(NoteWord + kind).Single(n =>
                    (string?)n.Attribute(NoteWord + "type") == selected).Attribute(NoteWord + "id")!;
                if (registration == "wrong-id") id = "unrelated-id";
                if (registration == "other-kind")
                {
                    string other = kind == "footnote" ? "endnote" : "footnote";
                    XElement otherProperties = settings.Root.Element(NoteWord + (other + "Pr")) ??
                        new XElement(NoteWord + (other + "Pr"));
                    if (otherProperties.Parent is null) settings.Root.Add(otherProperties);
                    otherProperties.Add(new XElement(NoteWord + other, new XAttribute(NoteWord + "id", id)));
                }
                else if (registration != "none")
                    properties.Add(new XElement(NoteWord + kind, new XAttribute(NoteWord + "id", id)));
            });
            foreach (DocxRelatedStory special in source.RelatedStories.Where(s => !s.IsNormalStoryType))
            {
                bool normal = special.Type == DocxRelatedStoryType.Separator;
                bool authored = normal ? registration == "normal" : registration == "continuation";
                DocxParagraph paragraph = special.Paragraphs.Single();
                TestAssert.True(special.IsSynthetic != authored,
                    "Only a matching typed settings reference activates authored separator formatting.");
                TestAssert.Equal(authored ? (normal ? "Public normal separator" : "Public continuation separator") : "",
                    string.Concat(paragraph.Runs.Select(r => r.Text)));
                if (authored)
                {
                    TestAssert.Equal(normal ? 8d : 16d, paragraph.Runs.Single().FontSize);
                    TestAssert.Equal("Arial", paragraph.Runs.Single().FontFamily);
                }
                else
                    TestAssert.True(paragraph.Runs.All(r => r.FontSize != 8d && r.FontSize != 16d),
                        "Unregistered stories use the document cascade instead of ignored direct formatting.");
            }
            TestAssert.True(source.RelatedStories.Where(s => s.IsNormalStoryType).All(s =>
                s.Paragraphs.SelectMany(p => p.Runs).Any(r => r.Text.Length != 0)),
                "Separator registration preserves the normal note sources and their typed IDs.");
        }
    }

    public static void DocxFittingSeparatorPaintRetainsSourceBearingWithoutMovingNoteFlow()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool custom in new[] { false, true })
        foreach (bool matching in new[] { false, true })
        foreach (var profile in new[] { (Ascender: 2049, Descender: 423, Strikeout: 689,
            WordMixedGap: 17.28d, WordMatchingGap: 17.88d),
            (Ascender: 2059, Descender: 430, Strikeout: 679,
            WordMixedGap: 17.16d, WordMatchingGap: 17.76d) })
        {
            DocxDocument source = WithFittingSeparatorBearingSource(custom, matching);
            var measurer = new FittingSeparatorBearingMeasurer(profile.Ascender, profile.Descender, profile.Strikeout);
            DocxLayout layout = engine.Create(source, measurer, default);
            DocxLayoutPage page = layout.Pages.Single(p => p.PlacedRelatedStories.Any(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote));
            TestAssert.True(page.Items.Count > 0, "The independent metric profile witnesses a fitting occupied head.");
            DocxPlacedRelatedStoryLayout end = page.PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote);
            DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote &&
                s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
            // Independent Word references supply these local rule/first-glyph
            // distances; source profiles use their recorded font table metrics.
            double expected = matching ? profile.WordMatchingGap : profile.WordMixedGap;
            double actual = separator.SeparatorY!.Value - end.TextLines[0].BaselineY;
            TestAssert.True(Math.Abs(actual - expected) < .13d,
                $"Independent Word fitting rule-to-first distance: {actual}, expected={expected}, custom={custom}, matching={matching}.");
            if (!matching)
                TestAssert.True(Math.Abs(separator.TopY - separator.Height - end.TextLines[0].BaselineY -
                    14.6484375d) < .002d, "Rule painting retains the existing Calibri content flow origin and gap.");
            DocxRelatedStoryLayout canonical = layout.RelatedStories.Single(s =>
                s.Story.IsNormalStoryType && s.Story.Kind == DocxRelatedStoryKind.Endnote);
            TestAssert.True(end.TextLines.Select(l => l.SourceLineIndex).SequenceEqual(
                canonical.TextLines.Select(l => l.SourceLineIndex)),
                "The rule correction retains the complete canonical note source.");
        }
    }

    public static void DocxFittingSeparatorPaintKeepsMissingLineBearingFallback()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool custom in new[] { false, true })
        foreach (var profile in new[] { (Ascender: 2049, Descender: 423, Strikeout: 689, LegacyGap: 17.966d),
            (Ascender: 2059, Descender: 430, Strikeout: 679, LegacyGap: 17.907d) })
        {
            DocxDocument source = WithFittingSeparatorBearingSource(custom);
            var measurer = new FittingSeparatorBearingMeasurer(profile.Ascender, profile.Descender,
                profile.Strikeout, availableLineBearing: false);
            DocxLayout layout = engine.Create(source, measurer, default);
            DocxLayoutPage page = layout.Pages.Single(p => p.PlacedRelatedStories.Any(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote));
            DocxPlacedRelatedStoryLayout end = page.PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote);
            DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote &&
                s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
            TestAssert.True(Math.Abs(separator.SeparatorY!.Value - end.TextLines[0].BaselineY - profile.LegacyGap) < .02d,
                "Missing line bearings retain the qualified rule fallback even when a Windows ascent is available.");
            TestAssert.True(Math.Abs(separator.TopY - separator.Height - end.TextLines[0].BaselineY -
                14.6484375d) < .002d, "The missing-metric fallback retains its source content gap.");
        }
    }

    private static DocxDocument WithFittingSeparatorBearingSource(bool custom, bool matching = false)
    {
        DocxDocument source = WithNoteLineRule(
            WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32), false, true, 6), 8, exact: false);
        return source with
        {
            RelatedStories = source.RelatedStories.Select(s => s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    Runs = p.Runs.Select(r => r with
                    {
                        FontFamily = !s.IsNormalStoryType || matching ? "source-bearing" : "content-bearing",
                    }).ToArray(),
                })).ToArray(),
            }).ToArray(),
        };
    }

    public static void DocxPhysicalFootnoteHeadUsesSharedSourceGapAndRetainsFallbacks()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool custom in new[] { false, true })
        foreach (bool matching in new[] { false, true })
        foreach (int footerLines in new[] { 0, 1, 6 })
        foreach (int footRepeats in new[] { 0, 96, 100, 104, 120 })
        {
            var measurer = new NoteContentMetricMeasurer(10.8d, 1.15d,
                matching ? 1.15d : 1.220703125d, useStrikeoutMetrics: true);
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32, footRepeats),
                false, footerLines != 0, Math.Max(1, footerLines));
            DocxLayout layout = engine.Create(source, measurer, default);
            DocxLayoutPage page = layout.Pages.First(p => p.PlacedRelatedStories.Any(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote));
            DocxPlacedRelatedStoryLayout foot = page.PlacedRelatedStories.First(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote);
            DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
            // Word's Arial separator with Calibri content has a 16.44pt
            // rule-to-first-glyph distance. The unframed path keeps its existing
            // 17.15pt distance; matching metrics retain the mark-gap floor.
            double expected = matching ? 16.31d : footerLines == 0 ? 17.15d : 16.44d;
            double actual = separator.SeparatorY!.Value - foot.TextLines[0].BaselineY;
            TestAssert.True(Math.Abs(actual - expected) < .02d,
                $"Physical head source gap and fallback: {actual}, custom={custom}, matching={matching}, footer={footerLines}.");
            DocxRelatedStoryLayout canonical = layout.RelatedStories.Single(s => s.Story.IsNormalStoryType &&
                s.Story.Kind == DocxRelatedStoryKind.Footnote);
            TestAssert.True(layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                .SelectMany(s => s.TextLines).Select(l => l.SourceLineIndex).SequenceEqual(
                    canonical.TextLines.Select(l => l.SourceLineIndex)),
                "Physical head placement retains every canonical source owner and order.");
        }
    }

    public static void DocxMixedNoteHeadAdmissionUsesTheSameSeparatorOriginAsPlacement()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer(10.8d, 1.15d, 1.15d, true);
        int twoLineWitnesses = 0;
        foreach (bool custom in new[] { false, true })
        foreach (int repeats in Enumerable.Range(48, 20).Select(n => n * 2))
        {
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32, repeats), false, true, 6);
            DocxDocument WithWidow(bool? widow) => source with
            {
                RelatedStories = source.RelatedStories.Select(s =>
                    s.Kind == DocxRelatedStoryKind.Endnote && s.IsNormalStoryType ? s with
                    {
                        BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                        {
                            KeepRules = p.KeepRules with { WidowControl = widow,
                                WidowControlValue = widow is null ? null : widow.Value ? "1" : "0" },
                        })).ToArray(),
                    } : s).ToArray(),
            };
            DocxLayout off = engine.Create(WithWidow(false), measurer, default);
            for (int pageIndex = 0; pageIndex < off.Pages.Count; pageIndex++)
            {
                DocxLayoutPage page = off.Pages[pageIndex];
                if (!page.PlacedRelatedStories.Any(s => s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    s.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator)) continue;
                DocxPlacedRelatedStoryLayout? head = page.PlacedRelatedStories.FirstOrDefault(s =>
                    s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && s.StoryLayout.Story.IsNormalStoryType &&
                    s.TextLines[0].SourceLineIndex == 0);
                if (head is null || head.TextLines.Count < 2) continue;
                if (head.TextLines.Count == 2) twoLineWitnesses++;
                foreach (bool? widow in new bool?[] { null, true })
                {
                    DocxLayout guarded = engine.Create(WithWidow(widow), measurer, default);
                    DocxPlacedRelatedStoryLayout? guardedHead = guarded.Pages[pageIndex].PlacedRelatedStories.FirstOrDefault(s =>
                        s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && s.StoryLayout.Story.IsNormalStoryType &&
                        s.TextLines[0].SourceLineIndex == 0);
                    TestAssert.True(guardedHead is not null && guardedHead.TextLines.Count >= 2,
                        $"Shared separator head origin: placement fits {head.TextLines.Count} lines, but admission deferred the source head; custom={custom}, widow={widow}, repeats={repeats}.");
                }
            }
        }
        TestAssert.True(twoLineWitnesses > 0, "Nearby source lengths witness the exact two-line admission boundary.");
    }

    public static void DocxMixedNoteHeadClearsCanonicalSeparatorAscentAfterRemovingLatentPadding()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer(10.8d, 1.15d, 1.15d, true);
        int occupiedWitnesses = 0;
        foreach (bool custom in new[] { false, true })
        foreach (bool? widow in new bool?[] { null, true, false })
        foreach (int repeats in new[] { 118, 120, 122 })
        {
            DocxDocument source = WithEndnoteWidowControl(
                WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32, repeats), false, true, 6), widow);
            DocxLayout layout = engine.Create(source, measurer, default);
            foreach (DocxLayoutPage page in layout.Pages)
            {
                DocxPlacedRelatedStoryLayout? separator = page.PlacedRelatedStories.FirstOrDefault(s =>
                    s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    s.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator);
                if (separator is null) continue;
                foreach (DocxPlacedRelatedStoryLayout endnote in page.PlacedRelatedStories.Where(s =>
                    s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && s.StoryLayout.Story.IsNormalStoryType))
                {
                    occupiedWitnesses++;
                    // Independent Word source-length and widow witnesses show
                    // the separator's blank line still owns its font ascent.
                    // Latent after-spacing can be removed from the flow charge,
                    // but cannot make this canonical line's ascent available.
                    double canonicalTop = separator.TextLines[0].BaselineY + 10.8d;
                    TestAssert.True(endnote.TopY - endnote.Height >= canonicalTop - .001d,
                        $"Canonical separator ascent: endnote bottom={endnote.TopY - endnote.Height}, occupied={canonicalTop}, custom={custom}, widow={widow}, repeats={repeats}.");
                }
            }
            foreach (DocxRelatedStoryLayout story in layout.RelatedStories.Where(s => s.Story.IsNormalStoryType))
                TestAssert.True(layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s =>
                    s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == story.Story.Kind)
                    .SelectMany(s => s.TextLines).Select(l => l.SourceLineIndex)
                    .SequenceEqual(story.TextLines.Select(l => l.SourceLineIndex)),
                    "The shared occupied boundary retains every original typed source line in order.");
        }
        TestAssert.True(occupiedWitnesses > 0, "Source-length neighbors exercise an endnote head above a continued physical footnote.");
    }

    public static void DocxMinimumNoteFirstLineBoxAndBearingUseOneNaturalOwner()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool custom in new[] { false, true })
        foreach (bool typographic in new[] { false, true })
        foreach (int minimum in new[] { 8, 12, 20 })
        {
            DocxDocument source = WithNoteLineRule(ReadMixedNoteRetryFixture(custom, 32, 120), minimum, exact: false);
            DocxLayout layout = engine.Create(source, new NoteNaturalBoxMetricMeasurer(typographic), default);
            foreach (DocxRelatedStoryLayout story in layout.RelatedStories.Where(s => s.Story.IsNormalStoryType))
            {
                DocxTextLineLayout first = story.TextLines[0];
                // The resolved 12pt first box is 14.484375pt and its Windows
                // ascent is 12.005859375pt. Retained Word controls independently
                // witness this first-line bearing and the .12pt positioning grid.
                double natural = typographic ? 13.8d : 14.484375d;
                double bearing = typographic ? 11.28d : 12.005859375d;
                double expected = bearing + Math.Max(0d, minimum - natural);
                TestAssert.True(Math.Abs(first.LineHeight!.Value - Math.Max(minimum, natural)) < .000001d &&
                    Math.Abs(first.SingleLineHeight!.Value - natural) < .000001d &&
                    Math.Abs(-first.BaselineY - expected) < .000001d,
                    $"First natural note owner: box={first.LineHeight}, inset={-first.BaselineY}, custom={custom}, typo={typographic}.");
                TestAssert.True(ReferenceEquals(first.SourceParagraph, story.Story.Paragraphs[0]),
                    "The box and bearing keep the original paragraph owner rather than cloning a new source policy.");
                TestAssert.True(layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s =>
                    s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == story.Story.Kind)
                    .SelectMany(s => s.TextLines).Select(l => l.SourceLineIndex)
                    .SequenceEqual(story.TextLines.Select(l => l.SourceLineIndex)),
                    "First-line ownership retains every canonical typed source line and order.");
            }
        }
    }

    public static void DocxMinimumNoteFirstLineKeepsPartialAndUnenlargedBoxFallbacks()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool custom in new[] { false, true })
        foreach (bool complete in new[] { false, true })
        foreach (int minimum in new[] { 8, 12, 20 })
        {
            DocxDocument source = WithNoteLineRule(ReadMixedNoteRetryFixture(custom, 32, 120), minimum, exact: false);
            IDocxTextMeasurer measurer = complete ? new NoteNaturalBoxMetricMeasurer(false, 1.220703125d) :
                new NoteContentMetricMeasurer(10.8d, 1.15d, 1.15d);
            DocxLayout layout = engine.Create(source, measurer, default);
            foreach (DocxRelatedStoryLayout story in layout.RelatedStories.Where(s => s.Story.IsNormalStoryType))
            {
                DocxTextLineLayout first = story.TextLines[0];
                double natural = complete ? 14.6484375d : 13.8d;
                double inset = 11.28d + Math.Max(0d, minimum - natural);
                TestAssert.True(Math.Abs(first.LineHeight!.Value - Math.Max(minimum, natural)) < .000001d &&
                    Math.Abs(-first.BaselineY - inset) < .000001d,
                    "Partial face metrics and boxes without a larger natural owner retain their qualified first-line bearing.");
            }
        }
    }

    public static void DocxMinimumNoteContentUsesSourceOwnedNaturalMetrics()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool custom in new[] { false, true })
        foreach (bool typographic in new[] { false, true })
        foreach (int minimum in new[] { 8, 12, 20 })
        {
            DocxDocument source = WithNoteLineRule(ReadMixedNoteRetryFixture(custom, 32, 120), minimum, exact: false);
            var measurer = new NoteNaturalBoxMetricMeasurer(typographic);
            DocxLayout layout = engine.Create(source, measurer, default);
            foreach (DocxRelatedStoryLayout story in layout.RelatedStories.Where(s => s.Story.IsNormalStoryType))
            {
                TestAssert.True(story.TextLines.Count > 2, "Both note kinds exercise source-owned content-only lines.");
                DocxTextLineLayout first = story.TextLines[0];
                TestAssert.True(Math.Abs(first.LineHeight!.Value - Math.Max(minimum, typographic ? 13.8d : 14.484375d)) < .000001d,
                    "The first prefix uses its independently qualified source natural box.");
                // Tahoma's resolved Windows/hhea content box is 2472/2048 at
                // 10pt. Its independently observed Word pitches alternate on
                // the .12pt grid around this box. Typographic requests retain
                // the existing single-line floor under the minimum rule.
                double natural = typographic ? 11.5d : 12.0703125d;
                foreach (DocxTextLineLayout line in story.TextLines.Skip(2))
                {
                    TestAssert.True(line.LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing &&
                        Math.Abs(line.SingleLineHeight!.Value - natural) < .000001d &&
                        Math.Abs(line.LineHeight!.Value - Math.Max(minimum, natural)) < .000001d,
                        $"Source-owned minimum note box: {line.LineHeight}, natural={line.SingleLineHeight}, custom={custom}, typo={typographic}.");
                }
                TestAssert.True(layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s =>
                    s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == story.Story.Kind)
                    .SelectMany(s => s.TextLines).Select(l => l.SourceLineIndex)
                    .SequenceEqual(story.TextLines.Select(l => l.SourceLineIndex)),
                    "Natural boxes retain every canonical typed source line through pagination.");
            }
        }
    }

    public static void DocxMinimumNoteContentRetainsExactAndPartialMetricFallbacks()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool custom in new[] { false, true })
        foreach (bool exact in new[] { false, true })
        foreach (int minimum in new[] { 8, 12, 20 })
        {
            DocxDocument source = WithNoteLineRule(ReadMixedNoteRetryFixture(custom, 32, 120), minimum, exact);
            IDocxTextMeasurer measurer = exact ? new NoteNaturalBoxMetricMeasurer(false) :
                new NoteContentMetricMeasurer(10.8d, 1.15d, 1.15d);
            DocxLayout layout = engine.Create(source, measurer, default);
            foreach (DocxRelatedStoryLayout story in layout.RelatedStories.Where(s => s.Story.IsNormalStoryType))
            {
                TestAssert.True(story.TextLines.Count > 2, "Each fallback exercises wrapped content in both note kinds.");
                foreach (DocxTextLineLayout line in story.TextLines.Skip(2))
                    TestAssert.True(exact
                        ? line.LineHeightSource == DocxLineHeightSource.ExactLineSpacing && line.LineHeight == minimum
                        : line.LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing &&
                            line.SingleLineHeight == 11.5d && line.LineHeight == Math.Max(minimum, 11.5d),
                        "Exact authored boxes and providers without resolved face semantics retain their prior treatment.");
            }
        }
    }

    private static DocxDocument WithNoteLineRule(DocxDocument source, int points, bool exact) => source with
    {
        RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
        {
            BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
            {
                LineSpacingPoints = points,
                Spacing = p.Spacing with { LineRuleValue = exact ? "exact" : "atLeast", LineValue = (points * 20).ToString(CultureInfo.InvariantCulture) },
            })).ToArray(),
        } : s).ToArray(),
    };

    public static void DocxContinuedPhysicalFootnoteGapUsesSharedAdmissionAndPreservesFallbacks()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool custom in new[] { false, true })
        foreach (bool matching in new[] { false, true })
        foreach (int footerLines in new[] { 0, 1, 6 })
        {
            int occupiedContinuations = 0;
            foreach (int footRepeats in new[] { 96, 100, 104, 120 })
            {
                var measurer = new NoteContentMetricMeasurer(10.8d, 1.15d,
                    matching ? 1.15d : 1.220703125d, useStrikeoutMetrics: true);
                DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32, footRepeats),
                    false, footerLines != 0, Math.Max(1, footerLines));
                DocxLayout layout = engine.Create(source, measurer, default);
                DocxRelatedStoryLayout canonical = layout.RelatedStories.Single(s => s.Story.IsNormalStoryType &&
                    s.Story.Kind == DocxRelatedStoryKind.Footnote);
                var feet = layout.Pages.SelectMany((p, index) => p.PlacedRelatedStories.Where(s =>
                    s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                    .Select(s => (Page: index, Story: s))).ToArray();
                TestAssert.True(feet.SelectMany(s => s.Story.TextLines).Select(l => l.SourceLineIndex)
                    .SequenceEqual(canonical.TextLines.Select(l => l.SourceLineIndex)),
                    "Continuation gap correction retains every canonical source line and order.");
                foreach (var foot in feet.Skip(1))
                {
                    DocxLayoutPage page = layout.Pages[foot.Page];
                    if (page.Items.Count == 0) continue;
                    occupiedContinuations++;
                    DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(s =>
                        s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                        s.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator);
                    double gap = separator.SeparatorY!.Value - foot.Story.TextLines[0].BaselineY;
                    // Golden distances for this controlled metric profile retain the
                    // matching floor and unframed fallback. Independent Word PDF
                    // comparisons qualify the real-font continuation geometry.
                    double expected = matching ? 15.03d : footerLines == 0 ? 15.87d : 15.16d;
                    TestAssert.True(Math.Abs(gap - expected) < .02d,
                        $"Continued source gap and fallback: {gap}, custom={custom}, matching={matching}, footer={footerLines}.");
                    TestAssert.True(foot.Story.TextLines.All(l => ReferenceEquals(l.SourceParagraph,
                        canonical.Story.Paragraphs[0])), "Continuations retain the canonical source paragraph.");
                }
            }
            TestAssert.True(occupiedContinuations > 0,
                "Both marker and metric profiles exercise a continuation alongside an occupied body page.");
        }
    }

    public static void DocxContinuedPhysicalFootnoteGapRetainsRichComplementarySourceFallbacks()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer(10.8d, 1.15d, 1.220703125d, useStrikeoutMetrics: true);
        DocxDocument plain = ReadMixedNoteRetryFixture(true, 32, 120);
        DocxRelatedStory end = plain.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Endnote);
        DocxParagraph paragraph = end.Paragraphs[0];
        foreach (DocxParagraph invalid in new[]
        {
            paragraph with { KeepRules = paragraph.KeepRules with { KeepNext = true } },
            paragraph with { KeepRules = paragraph.KeepRules with { KeepLines = true } },
            paragraph with { Spacing = paragraph.Spacing with { ContextualSpacing = true } },
            paragraph with { Alignment = DocxTextAlignment.Center },
        })
        {
            DocxDocument source = WithMixedNoteStaticFrame(plain with { RelatedStories = plain.RelatedStories.Select(s =>
                ReferenceEquals(s, end) ? s with { BodyElements = [new DocxParagraphElement(invalid)] } : s).ToArray() },
                false, true, 6);
            DocxLayout layout = engine.Create(source, measurer, default);
            var feet = layout.Pages.SelectMany((p, index) => p.PlacedRelatedStories.Where(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                .Select(s => (Page: index, Story: s))).ToArray();
            TestAssert.True(feet.Length > 1 && feet.All(s => !s.Story.UsesPlainFootnoteFrame),
                "Rich original complementary sources retain the existing physical-frame fallback on every page.");
            DocxRelatedStoryLayout canonical = layout.RelatedStories.Single(s => s.Story.IsNormalStoryType &&
                s.Story.Kind == DocxRelatedStoryKind.Footnote);
            TestAssert.True(feet.SelectMany(s => s.Story.TextLines).Select(l => l.SourceLineIndex)
                .SequenceEqual(canonical.TextLines.Select(l => l.SourceLineIndex)),
                "Rich complementary fallbacks retain the complete typed footnote source.");
            int occupiedContinuations = 0;
            foreach (var foot in feet.Skip(1))
            {
                DocxLayoutPage page = layout.Pages[foot.Page];
                if (page.Items.Count == 0) continue;
                occupiedContinuations++;
                DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(s =>
                    s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    s.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator);
                double gap = separator.SeparatorY!.Value - foot.Story.TextLines[0].BaselineY;
                TestAssert.True(Math.Abs(gap - 15.87d) < .02d,
                    $"Unqualified complementary source keeps its existing continuation gap: {gap}.");
            }
            TestAssert.True(occupiedContinuations > 0, "Every rich-source guard exercises an occupied continuation.");
        }
    }

    public static void DocxJointMixedNoteTypographyPreservesTypedOwnersAndSourceMetadata()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool custom in new[] { true, false })
        foreach (string threshold in new[] { "0", "20", "22" })
        {
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32), false, true, 6);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    Runs = p.Runs.Select(r => r with { KerningMinimumHalfPointsValue = threshold }).ToArray(),
                })).ToArray(),
            } : s).ToArray() };
            DocxLayout first = engine.Create(source, measurer, default);
            TestAssert.True(DocxNoteKerning.TryCreateJointDocument(source, first, out DocxDocument adjusted, out var owners, default),
                "Both admitted note kinds share one source-threshold adjustment plan.");
            TestAssert.True(owners is not null && owners.Footnotes.ContainsKey("37") && owners.Endnotes.ContainsKey("37"),
                "Coincident numeric IDs retain separate typed owners.");
            TestAssert.True(ReferenceEquals(source.BodyElements, adjusted.BodyElements), "The plan preserves all body and inline-reference objects.");
            foreach (DocxRelatedStory story in adjusted.RelatedStories.Where(s => s.IsNormalStoryType))
            {
                DocxRelatedStory original = source.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == story.Kind && s.Id == story.Id);
                TestAssert.Equal(string.Concat(original.Paragraphs[0].Runs.Select(r => r.Text)),
                    string.Concat(story.Paragraphs[0].Runs.Select(r => r.Text)));
                DocxTextRun originalPrefix = original.Paragraphs[0].Runs[0];
                TestAssert.Equal(originalPrefix.IsAutomaticNoteStoryMarker ? originalPrefix :
                    originalPrefix with { Text = originalPrefix.Text[..1] }, story.Paragraphs[0].Runs[0]);
                TestAssert.True(story.Paragraphs[0].Runs.Skip(1).Where(r => r.Text.Length != 0).All(r =>
                    r.PairKerningEnabled == (threshold == "20" || threshold == "22" && r.FontSize == 12d) &&
                    r.UseLogicalSizeForPdfAdvances == (threshold == "0" || r.FontSize == 10d)),
                    $"Shared source policy: {story.Kind}, threshold {threshold}, states " + string.Join("; ", story.Paragraphs[0].Runs.Skip(1).Select(r => $"{r.FontSize}/{r.PairKerningEnabled}/{r.UseLogicalSizeForPdfAdvances}/{r.VerticalAlignment}/{r.KerningMinimumHalfPointsValue}")));
                TestAssert.True(original.Paragraphs[0].Runs.All(r => r.PairKerningEnabled && !r.UseLogicalSizeForPdfAdvances),
                    "Source typography metadata is immutable.");
            }
            DocxLayout changed = engine.Create(adjusted, measurer, default);
            TestAssert.True(DocxNoteKerning.HasStableJointOwners(changed, owners!, default),
                "The combined retry retains complete canonical owners for both kinds.");
            foreach (DocxRelatedStoryKind normalized in new[] { DocxRelatedStoryKind.Footnote, DocxRelatedStoryKind.Endnote })
            {
                DocxDocument partial = source with { RelatedStories = source.RelatedStories.Select((s, index) =>
                    s.Kind == normalized ? adjusted.RelatedStories[index] : s).ToArray() };
                TestAssert.True(DocxNoteKerning.TryCreateJointDocument(partial, engine.Create(partial, measurer, default),
                    out var completed, out var completeOwners, default),
                    "A normalized complementary kind still admits the remaining source correction.");
                TestAssert.True(DocxNoteKerning.HasStableJointOwners(engine.Create(completed, measurer, default), completeOwners!, default),
                    "Both complete typed owners survive a partially normalized source.");
                TestAssert.True(ReferenceEquals(partial.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == normalized).Paragraphs[0],
                    completed.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == normalized).Paragraphs[0]),
                    "The already normalized paragraph is retained exactly.");
            }
            TestAssert.True(!DocxNoteKerning.TryCreateJointDocument(adjusted, changed, out _, out _, default),
                "An already normalized document requires no additional typography retry.");
        }
    }

    public static void DocxJointMixedNoteTypographyRetainsSingleKindAndIncompleteOwnerFallbacks()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(true, 32), false, true, 6);
        DocxLayout first = engine.Create(source, measurer, default);
        TestAssert.True(DocxNoteKerning.TryCreateJointDocument(source, first, out var adjusted, out var owners, default),
            "The guard fixture has complete admitted owners.");
        foreach (DocxRelatedStoryKind lost in new[] { DocxRelatedStoryKind.Footnote, DocxRelatedStoryKind.Endnote })
        {
            DocxLayout RemoveKind(DocxLayout layout) => layout with { Pages = layout.Pages.Select(p => p with
            {
                PlacedRelatedStories = p.PlacedRelatedStories.Where(s =>
                    !s.StoryLayout.Story.IsNormalStoryType || s.StoryLayout.Story.Kind != lost).ToArray(),
            }).ToArray() };
            TestAssert.True(!DocxNoteKerning.TryCreateJointDocument(source, RemoveKind(first), out _, out _, default),
                "A missing typed kind rejects joint source admission.");
            TestAssert.True(!DocxNoteKerning.HasStableJointOwners(RemoveKind(engine.Create(adjusted, measurer, default)), owners!, default),
                "A lost complementary kind rejects the adjusted retry.");
        }
        DocxDocument noStatic = ReadMixedNoteRetryFixture(true, 32);
        DocxLayout fallback = engine.Create(noStatic, measurer, default);
        TestAssert.True(!DocxNoteKerning.TryCreateJointDocument(noStatic, fallback, out _, out _, default),
            "An unadmitted footnote frame keeps the single-kind fallback.");
        TestAssert.True(DocxNoteKerning.TryCreateDocument(noStatic, fallback, out _, out _, default),
            "The independently admitted endnote path remains available.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxNoteKerning.TryCreateJointDocument(source, first, out _, out _, cancelled.Token));
    }

    public static void DocxMixedNoteEmissionSharesSourceNominalAdvancesForBothOwners()
    {
        foreach (bool custom in new[] { true, false })
        {
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32), false, true, 6);
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            DocxTextEmissionSnapshot emitted = renderer.InspectTextEmission(source);
            foreach (string kind in new[] { "Footnote", "Endnote" })
            {
                var content = emitted.Lines.Where(l => l.StoryKind == kind && l.StoryVariantType is null)
                    .SelectMany(l => l.Segments).Where(s => s.TextLength > 1 && !s.IsTerminalLineSpace).ToArray();
                TestAssert.True(content.Length != 0 && content.Any(s => s.AdvanceProfile.NominalAdvanceAdjustmentTotal > 0d),
                    "Both source-owned mixed note kinds retain logical advances when the emitted font size rounds.");
            }
        }
    }

    public static void DocxMixedNoteThresholdEmissionKeepsAuthoredSuffixAndContentPolicies()
    {
        foreach (bool custom in new[] { true, false })
        foreach (string threshold in new[] { "20", "22" })
        foreach (int footerLines in new[] { 1, 6 })
        foreach (int footnoteRepeats in new[] { 0, 120 })
        {
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32, footnoteRepeats),
                false, true, footerLines);
            source = source with
            {
                RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
                {
                    BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    {
                        Runs = p.Runs.Select((r, index) => r with
                        {
                            // AB has a real pair adjustment in the deterministic test face.
                            Text = index == 0 ? custom ? "*AB" : r.Text : "AB " + r.Text,
                            KerningMinimumHalfPointsValue = threshold,
                        }).ToArray(),
                    })).ToArray(),
                } : s).ToArray(),
            };
            var authoredRuns = source.RelatedStories.Where(s => s.IsNormalStoryType)
                .SelectMany(s => s.Paragraphs[0].Runs).ToArray();
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            DocxTextEmissionSnapshot emitted = renderer.InspectTextEmission(source);
            foreach (string kind in new[] { "Footnote", "Endnote" })
            {
                var segments = emitted.Lines.Where(l => l.StoryKind == kind && l.StoryVariantType is null)
                    .SelectMany(l => l.Segments).Where(s => !s.IsTerminalLineSpace).ToArray();
                var content = segments.Where(s => s.FontSize == 10d && s.TextLength > 1).ToArray();
                TestAssert.True(content.Length != 0 && content.Any(s => s.AdvanceProfile.NominalAdvanceAdjustmentTotal > 0d),
                    "Rounded note content retains logical advances at either source threshold.");
                TestAssert.True(threshold == "20"
                    ? content.Any(s => s.AdvanceProfile.KerningAdjustmentTotal < 0d)
                    : content.All(s => s.AdvanceProfile.KerningAdjustmentTotal == 0d),
                    "Source logical size enables a met threshold and disables an unmet threshold in each note kind.");
                if (!custom) continue;
                var suffix = segments.Where(s => s.FontSize == 12d && s.TextLength == 2 &&
                    s.CharacterProfile.LetterCount == 2).ToArray();
                TestAssert.True(suffix.Length == 1 && suffix[0].PdfFontSize == 12d &&
                    suffix[0].AdvanceProfile.KerningAdjustmentTotal < 0d &&
                    Math.Abs(suffix[0].AdvanceProfile.NominalAdvanceAdjustmentTotal) < .000000001d,
                    "Text following the custom marker keeps its authored size and its independently met threshold.");
            }
            TestAssert.True(source.RelatedStories.Where(s => s.IsNormalStoryType)
                .SelectMany(s => s.Paragraphs[0].Runs).SequenceEqual(authoredRuns),
                "Emission does not split or rewrite the authored source record.");
        }
    }

    private static DocxDocument WithMixedNoteStaticFrame(DocxDocument source, bool header, bool footer, int count,
        DocxParagraph? paragraph = null)
    {
        paragraph ??= DocxTests.CreateDocxLayoutParagraph("public static frame", 12d, 14d) with
        {
            LineSpacingPoints = null, Spacing = DocxParagraphSpacing.Empty with { BeforeValue = "0", AfterValue = "0", LineValue = "240", LineRuleValue = "auto" },
        };
        var paragraphs = Enumerable.Repeat(paragraph, count).ToArray();
        var headers = new Dictionary<string, IReadOnlyList<DocxParagraph>>();
        var footers = new Dictionary<string, IReadOnlyList<DocxParagraph>>();
        if (header) headers["default"] = paragraphs;
        if (footer) footers["default"] = paragraphs;
        var headerElements = headers.ToDictionary(p => p.Key, p => (IReadOnlyList<DocxBodyElement>)p.Value.Select(v => (DocxBodyElement)new DocxParagraphElement(v)).ToArray());
        var footerElements = footers.ToDictionary(p => p.Key, p => (IReadOnlyList<DocxBodyElement>)p.Value.Select(v => (DocxBodyElement)new DocxParagraphElement(v)).ToArray());
        var settings = source.PageSettings with { HeaderDistancePoints = 36d, FooterDistancePoints = 36d,
            HeaderParagraphsByType = headers, FooterParagraphsByType = footers,
            HeaderBodyElementsByType = headerElements, FooterBodyElementsByType = footerElements };
        return source with { PageSettings = settings, FinalSectionBreak = source.FinalSectionBreak! with { PageSettings = settings },
            HeaderParagraphs = header ? paragraphs : [], FooterParagraphs = footer ? paragraphs : [],
            HeaderParagraphsByType = headers, FooterParagraphsByType = footers,
            HeaderBodyElementsByType = headerElements, FooterBodyElementsByType = footerElements };
    }

    private static DocxDocument ReadMixedNoteRetryFixture(bool custom, int repeats, int footnoteRepeats = 0,
        Action<XDocument, XDocument, Dictionary<string, string>>? modify = null)
    {
        return ReadSectionNumberingFixture("footnote", (document, settings, parts) =>
        {
            XElement section = new(document.Descendants(NoteWord + "sectPr").Last());
            section.Elements().Where(e => e.Name.LocalName is "footnotePr" or "endnotePr" or "headerReference" or "footerReference").Remove();
            section.Add(new XElement(NoteWord + "endnotePr", new XElement(NoteWord + "pos", new XAttribute(NoteWord + "val", "docEnd"))));
            XElement body = document.Root!.Element(NoteWord + "body")!; body.RemoveNodes();
            XElement Properties(int size, bool marker) => new(NoteWord + "rPr",
                new XElement(NoteWord + "rFonts", new XAttribute(NoteWord + "ascii", "Calibri"), new XAttribute(NoteWord + "hAnsi", "Calibri")),
                new XElement(NoteWord + "sz", new XAttribute(NoteWord + "val", size)),
                marker ? new XElement(NoteWord + "vertAlign", new XAttribute(NoteWord + "val", custom ? "baseline" : "superscript")) :
                    new XElement(NoteWord + "kern", new XAttribute(NoteWord + "val", "0")));
            XElement Spacing() => new(NoteWord + "pPr", new XElement(NoteWord + "spacing", new XAttribute(NoteWord + "before", "0"),
                new XAttribute(NoteWord + "after", "0"), new XAttribute(NoteWord + "line", "240"), new XAttribute(NoteWord + "lineRule", "auto")));
            foreach (string kind in new[] { "footnote", "endnote" })
            {
                XElement reference = new(NoteWord + kind + "Reference", new XAttribute(NoteWord + "id", "37"));
                if (custom) reference.Add(new XAttribute(NoteWord + "customMarkFollows", "1"));
                body.Add(new XElement(NoteWord + "p", Spacing(), new XElement(NoteWord + "r", Properties(24, false),
                    new XElement(NoteWord + "t", "Public " + kind + " body before ")),
                    new XElement(NoteWord + "r", Properties(24, true), reference, custom ? new XElement(NoteWord + "t", "*tail") : null),
                    new XElement(NoteWord + "r", Properties(24, false), new XElement(NoteWord + "t", " body after."))));
                XElement Note(string? type, string id, object content) => new(NoteWord + kind, new XAttribute(NoteWord + "id", id),
                    type is null ? null : new XAttribute(NoteWord + "type", type), new XElement(NoteWord + "p", Spacing(), content));
                parts["word/" + kind + "s.xml"] = new XDocument(new XElement(NoteWord + kind + "s",
                    Note("separator", "-1", new XElement(NoteWord + "r", new XElement(NoteWord + "separator"))),
                    Note("continuationSeparator", "0", new XElement(NoteWord + "r", new XElement(NoteWord + "continuationSeparator"))),
                    Note(null, "37", new[] { new XElement(NoteWord + "r", Properties(24, true), custom ?
                        new XElement(NoteWord + "t", "*tail") : new XElement(NoteWord + kind + "Ref")),
                        new XElement(NoteWord + "r", Properties(20, false), new XElement(NoteWord + "t", kind == "footnote" ?
                            " Public short footnote content." + string.Concat(Enumerable.Repeat(
                                " Public long footnote continuation keeps its canonical owner and public content.", footnoteRepeats)) :
                            string.Join(' ', Enumerable.Repeat("Public mixed endnote content continues.", repeats)))) }))).ToString();
            }
            body.Add(section); settings.Root!.RemoveNodes();
            XDocument relationships = XDocument.Parse(parts["word/_rels/document.xml.rels"]);
            XNamespace rel = relationships.Root!.Name.Namespace;
            relationships.Root.Add(new XElement(rel + "Relationship", new XAttribute("Id", "rIdMixedEndnotes"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"), new XAttribute("Target", "endnotes.xml")));
            parts["word/_rels/document.xml.rels"] = relationships.ToString();
            XDocument types = XDocument.Parse(parts["[Content_Types].xml"]);
            types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/word/endnotes.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")));
            parts["[Content_Types].xml"] = types.ToString();
            modify?.Invoke(document, settings, parts);
        });
    }

    public static void DocxCustomNoteSuffixMetricsFollowActualSourceSpans()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool footnote in new[] { false, true })
        foreach (string suffix in new[] { "*", "tail", string.Join(' ', Enumerable.Repeat(
            "Adjacent source text remains at its authored size and wraps across lines.", 8)) })
        foreach (double minimum in new[] { 8d, 18d })
        {
            DocxDocument source = ReadCustomNoteMetricFixture(footnote, 10d);
            DocxDocument WithMinimum(double points) => source with
            {
                RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
                {
                    BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    {
                        Runs = [p.Runs[0] with { Text = "*" + suffix }, p.Runs[1]],
                        LineSpacingPoints = points, Spacing = p.Spacing with { LineRuleValue = "atLeast",
                            LineValue = (points * 20d).ToString(CultureInfo.InvariantCulture) },
                    })).ToArray(),
                } : s).ToArray(),
            };
            var kind = footnote ? DocxRelatedStoryKind.Footnote : DocxRelatedStoryKind.Endnote;
            DocxDocument Correct(DocxDocument authored)
            {
                TestAssert.True(DocxNoteKerning.TryCreateDocument(authored, engine.Create(authored, measurer, default),
                    out var adjusted, out var owners, default, kind), "The complete note retry splits an owned custom prefix.");
                TestAssert.True(DocxNoteKerning.HasStableOwners(engine.Create(adjusted, measurer, default), owners, default, kind),
                    "Wrapped adjacent source text preserves all canonical note owners.");
                TestAssert.True(authored.BodyElements.Zip(adjusted.BodyElements).All(p => ReferenceEquals(p.First, p.Second)),
                    "The suffix correction preserves authored body references.");
                foreach (var pair in authored.RelatedStories.Zip(adjusted.RelatedStories).Where(p => p.First.IsNormalStoryType))
                {
                    DocxTextRun original = pair.First.Paragraphs[0].Runs[0];
                    DocxTextRun prefix = pair.Second.Paragraphs[0].Runs[0], tail = pair.Second.Paragraphs[0].Runs[1];
                    TestAssert.Equal(original with { Text = "*" }, prefix);
                    TestAssert.Equal(suffix, tail.Text);
                    TestAssert.Equal(original.SourceRunIndex, tail.SourceRunIndex);
                    TestAssert.Equal(original.SourceTextOffsetInRun + 1, tail.SourceTextOffsetInRun);
                    TestAssert.True(!tail.IsAutomaticNoteStoryMarker && !tail.PairKerningEnabled && tail.UseLogicalSizeForPdfAdvances,
                        "Only the first character is protected; its adjacent text keeps its own advance policy and XML offset.");
                    TestAssert.Equal("*" + suffix, original.Text);
                }
                return adjusted;
            }
            DocxLayout natural = engine.Create(Correct(WithMinimum(8d)), measurer, default);
            DocxLayout bounded = engine.Create(Correct(WithMinimum(minimum)), measurer, default);
            foreach (var pair in natural.RelatedStories.Where(s => s.Story.IsNormalStoryType)
                .Zip(bounded.RelatedStories.Where(s => s.Story.IsNormalStoryType)))
            {
                TestAssert.Equal(pair.First.TextLines.Count, pair.Second.TextLines.Count);
                double naturalBottom = 0d, boundedBottom = 0d;
                bool suffixContinues = false, contentContinues = false;
                foreach (var lines in pair.First.TextLines.Zip(pair.Second.TextLines))
                {
                    double expected = lines.First.Segments.Max(s => s.StyleRun.FontSize) * 1.2d;
                    naturalBottom += expected; boundedBottom += Math.Max(minimum, expected);
                    TestAssert.True(Math.Abs(lines.First.LineHeight!.Value - expected) < .000001d &&
                        Math.Abs(lines.Second.LineHeight!.Value - Math.Max(minimum, expected)) < .000001d &&
                        Math.Abs(lines.First.BaselineY + naturalBottom - lines.Second.BaselineY - boundedBottom) < .000001d,
                        "Each wrapped line uses its present source sizes and retains its bottom bearing inside a minimum box.");
                    TestAssert.True(ReferenceEquals(lines.Second.SourceParagraph, pair.Second.Story.Paragraphs[0]),
                        "Metric views retain the complete source paragraph owner.");
                    suffixContinues |= lines.First.IsFirstParagraphLine != true && lines.First.Segments.Any(s => s.StyleRun.FontSize == 12d);
                    contentContinues |= lines.First.IsFirstParagraphLine != true && expected == 12d;
                }
                TestAssert.True(contentContinues && (suffix.Length < 20 || suffixContinues),
                    "Long prefixes wrap before the smaller content continues with its own metric box.");
            }
        }
    }

    public static void DocxCustomNoteLineMetricsRetainFixedRichScaledAndAmbiguousFallbacks()
    {
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool footnote in new[] { false, true })
        foreach (string variant in new[] { "exact", "scaled", "rich", "kept", "contextual", "missing", "duplicate-reference", "duplicate-target" })
        {
            DocxDocument source = ReadCustomNoteMetricFixture(footnote, 10d);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    Alignment = variant == "rich" ? DocxTextAlignment.Center : p.Alignment,
                    KeepRules = p.KeepRules with { KeepLines = variant == "kept" },
                    LineSpacingPoints = variant == "exact" ? 14d : null,
                    Spacing = p.Spacing with { ContextualSpacing = variant == "contextual",
                        LineRuleValue = variant == "exact" ? "exact" : "auto", LineValue = variant == "exact" ? "280" : "240" },
                })).ToArray(),
            } : s).ToArray() };
            if (variant is "missing" or "duplicate-reference") source = source with
            {
                BodyElements = source.BodyElements.Select(e => e is DocxParagraphElement pe ? pe with
                {
                    Paragraph = pe.Paragraph with { InlineReferences = variant == "missing" ? pe.Paragraph.InlineReferences
                        .Select(r => r with { CustomMarkText = null }).ToArray() :
                        pe.Paragraph.InlineReferences.Concat(pe.Paragraph.InlineReferences).ToArray() },
                } : e).ToArray(),
            };
            if (variant == "duplicate-target") source = source with { RelatedStories = source.RelatedStories.Concat(
                source.RelatedStories.Where(s => s.IsNormalStoryType)).ToArray() };
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, variant == "scaled" ? .5d : 1d);
            DocxLayout layout = engine.Create(source, measurer, default);
            TestAssert.True(layout.RelatedStories.Where(s => s.Story.IsNormalStoryType).SelectMany(s => s.TextLines)
                .All(line => Math.Abs(line.LineHeight!.Value - (variant == "exact" ? 14d : 14.4d)) < .000001d),
                "Fixed, rich, kept, contextual, originally scaled and ambiguous custom notes retain their established metric boxes.");
        }
    }

    private static DocxDocument ReadCustomNoteMetricFixture(bool footnote, double contentSize)
    {
        DocxDocument source = ReadCustomDocumentEndnoteFlowFixture("*", footnote);
        return source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
        {
            BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
            {
                Runs = [p.Runs[0] with { FontSize = 12d }, p.Runs[1] with { FontSize = contentSize,
                    Text = string.Join(' ', Enumerable.Repeat("Public custom note source content continues.", 100)) }],
            })).ToArray(),
        } : s).ToArray() };
    }

    private sealed class NoteNaturalBoxMetricMeasurer(bool useTypographicMetrics, double singleLineEm = 1.15d) :
        IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider,
        IDocxHheaLineGapProvider, IDocxHheaDescenderProvider, IDocxTypographicMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => text.Length * fontSize * .5d;
        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => fontSize * singleLineEm;
        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => fontSize * 2472d / 2048d;
        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => fontSize * 2049d / 2048d;
        public double MeasureHheaDescender(DocxTextRun? run, double fontSize) => fontSize * 423d / 2048d;
        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => 0d;
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => MeasureHheaAscender(run, fontSize);
        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => MeasureHheaDescender(run, fontSize);
        public bool UseTypographicMetrics(DocxTextRun? run) => useTypographicMetrics;
        public double MeasureTypographicLineHeight(DocxTextRun? run, double fontSize) => fontSize;
    }

    private sealed class FittingSeparatorBearingMeasurer(int ascenderUnits, int descenderUnits,
        int strikeoutUnits, bool availableLineBearing = true) : IDocxTextMeasurer,
        IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxHheaLineGapProvider, IDocxHheaDescenderProvider,
        IDocxTypographicMetricsProvider
    {
        private static bool IsSourceBearing(DocxTextRun? run) => run?.FontFamily == "source-bearing";
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => text.Length * fontSize * .5d;
        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) =>
            fontSize * (IsSourceBearing(run) ? 1.15d : 1.220703125d);
        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = IsSourceBearing(run) ? 1.15d : 1.220703125d;
            return true;
        }
        public bool TryGetStrikeoutRuleMetrics(DocxTextRun? run, out double positionEm, out double thicknessEm)
        {
            positionEm = strikeoutUnits / 2048d;
            thicknessEm = .0625d;
            return IsSourceBearing(run);
        }
        public bool UseTypographicMetrics(DocxTextRun? run) => false;
        public double MeasureTypographicLineHeight(DocxTextRun? run, double fontSize) => fontSize *
            (IsSourceBearing(run) ? ascenderUnits == 2049 ? 1d : 2191d / 2048d : 1.220703125d);
        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => fontSize *
            (IsSourceBearing(run) ? (ascenderUnits + descenderUnits) / 2048d : 1.220703125d);
        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => fontSize *
            (IsSourceBearing(run) ? availableLineBearing ? ascenderUnits / 2048d : 0d : .75d);
        public double MeasureHheaDescender(DocxTextRun? run, double fontSize) => fontSize *
            (IsSourceBearing(run) ? descenderUnits / 2048d : .25d);
        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) =>
            IsSourceBearing(run) ? 0d : fontSize * .220703125d;
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize *
            (IsSourceBearing(run) ? ascenderUnits / 2048d : 1950d / 2048d);
        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => fontSize *
            (IsSourceBearing(run) ? descenderUnits / 2048d : 550d / 2048d);
    }

    private sealed class NoteContentMetricMeasurer(double? separatorWindowsAscender = null,
        double? separatorSingleLineEm = null, double? contentSingleLineEm = null,
        bool useStrikeoutMetrics = false) : IDocxTextMeasurer, IDocxLineMetricsProvider,
        IDocxStaticTextMetricsProvider, IDocxHheaLineGapProvider, IDocxHheaDescenderProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => text.Length * fontSize * .5d;
        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => fontSize *
            ((string.IsNullOrEmpty(run?.Text) ? separatorSingleLineEm : contentSingleLineEm) ?? 1.2d);
        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = (string.IsNullOrEmpty(run?.Text) ? separatorSingleLineEm : contentSingleLineEm) ?? 0d;
            return value > 0d;
        }
        public bool TryGetStrikeoutRuleMetrics(DocxTextRun? run, out double positionEm, out double thicknessEm)
        {
            positionEm = .2587890625d;
            thicknessEm = .05d;
            return useStrikeoutMetrics && string.IsNullOrEmpty(run?.Text);
        }
        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => fontSize * 1.2d;
        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => fontSize * .9d;
        public double MeasureHheaDescender(DocxTextRun? run, double fontSize) => fontSize * .2d;
        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => fontSize * .1d;
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) =>
            string.IsNullOrEmpty(run?.Text) && separatorWindowsAscender is double value ? value : fontSize * .9d;
        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => fontSize * .2d;
    }

    private static DocxDocument ReadWrappedFootnoteFlowFixture(int repeats)
    {
        DocxDocument source = ReadNotePrecedenceFixture("footnote", "decimal", null, null, true);
        return source with { RelatedStories = source.RelatedStories.Where(s => s.Kind == DocxRelatedStoryKind.Footnote).Select(s => s.IsNormalStoryType ? s with
        {
            BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
            {
                SpacingBeforePoints = 0d, SpacingAfterPoints = 0d, LineSpacingPoints = null, LineSpacingFactor = 1d,
                Spacing = DocxParagraphSpacing.Empty with { BeforeValue = "0", AfterValue = "0", LineValue = "240", LineRuleValue = "auto" },
                Runs = [p.Runs[0] with { FontSize = 12d }, p.Runs[1] with { FontSize = 12d,
                    Text = string.Join(' ', Enumerable.Repeat("Public continued footnote content.", repeats)) }],
            })).ToArray(),
        } : s).ToArray() };
    }

    private sealed class LogicalBoundaryTextMeasurer : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        private readonly DocxTests.FamilyWidthTextMeasurer inner = new();
        public double MeasureText(DocxTextRun? run, string text, double fontSize) =>
            inner.MeasureText(run, text, fontSize) * (run?.UseLogicalSizeForPdfAdvances == true ? 1.20d : 1d);
        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => inner.MeasureSingleLineHeight(run, fontSize);
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => inner.MeasureWindowsAscender(run, fontSize);
        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => inner.MeasureWindowsDescender(run, fontSize);
    }

    private static DocxDocument WithEndnoteText(DocxDocument source, string? id, string text) => source with
    {
        RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType && s.Id == id ? s with
        {
            BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                { Runs = [p.Runs[0], p.Runs[1] with { Text = text }] })).ToArray(),
        } : s).ToArray(),
    };

    private static DocxDocument WithEndnoteWidowControl(DocxDocument source, bool? widow) => source with
    {
        RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
        {
            BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
            {
                KeepRules = p.KeepRules with { WidowControl = widow, WidowControlValue = widow is null ? null : widow.Value ? "1" : "0" },
            })).ToArray(),
        } : s).ToArray(),
    };

    private static DocxDocument ReadWrappedDocumentEndnoteFlowFixture(int before, int after, int repeats)
    {
        DocxDocument source = ReadDocumentEndnoteFlowFixture(before, after);
        return source with { RelatedStories = source.RelatedStories.Select(s => s.Kind == DocxRelatedStoryKind.Endnote && s.IsNormalStoryType ? s with
        {
            BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
            {
                Runs = [p.Runs[0], p.Runs[1] with { Text = string.Join(' ', Enumerable.Repeat("Public wrapped endnote content.", repeats)) }],
            })).ToArray(),
        } : s).ToArray() };
    }

    private static DocxDocument ReadDocumentEndnoteFlowFixture(int before, int after)
    {
        DocxDocument document = ReadNotePrecedenceFixture("endnote", "decimal", null, null, true);
        return document with { RelatedStories = document.RelatedStories.Select(s =>
            s.Kind == DocxRelatedStoryKind.Endnote && s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    SpacingBeforePoints = before, SpacingAfterPoints = after, LineSpacingFactor = 1d, LineSpacingPoints = null,
                    Spacing = p.Spacing with { BeforeValue = (before * 20).ToString(CultureInfo.InvariantCulture),
                        AfterValue = (after * 20).ToString(CultureInfo.InvariantCulture), LineValue = "240", LineRuleValue = "auto" },
                })).ToArray()
            } : s).ToArray() };
    }

    private static DocxDocument ReadModifiedNoteNumberingFixture(string fixture, Action<XDocument> modify)
    {
        string seed = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", fixture + ".docx"));
        var parts = new Dictionary<string, string>();
        using (ZipArchive archive = ZipFile.OpenRead(seed))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                parts[entry.FullName] = reader.ReadToEnd();
            }
        }
        XDocument xml = XDocument.Parse(parts["word/document.xml"]);
        modify(xml);
        parts["word/document.xml"] = xml.ToString();
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(input); }
    }

    public static void DocxReaderRetainsAutomaticNoteStoryMarks()
    {
        foreach (string fixture in NoteNavigationFixtures)
        {
            DocxDocument document = ReadNoteNavigationFixture(fixture);
            foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements))
            foreach (DocxInlineReference reference in paragraph.InlineReferences)
            {
                if (reference.Kind is not (DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote)) continue;
                DocxRelatedStory story = document.RelatedStories.Single(story => story.Kind == reference.Kind && story.Id == reference.Id);
                TestAssert.Equal(reference.DisplayText!, story.Paragraphs[0].Runs[0].Text);
                TestAssert.Equal(DocxRunVerticalAlignment.Superscript, story.Paragraphs[0].Runs[0].VerticalAlignment);
            }
        }
    }

    public static void DocxNoteReferenceAnnotationsNavigateAcrossPages()
    {
        foreach (string fixture in NoteNavigationFixtures)
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            DocxDocument document = ReadNoteNavigationFixture(fixture);
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
            PdfPage[] pages = renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().ToArray();
            PdfLinkAnnotation[] links = pages.SelectMany(page => page.Annotations).ToArray();
            int expected = fixture is "note-nav-two-footnotes" or "note-nav-mixed-notes" or "note-nav-numbered-ids" ? 2 : 1;
            TestAssert.Equal(expected, links.Length);
            TestAssert.True(links.All(link => link.IsDestination && link.Width > 0d && link.Height > 0d),
                "Every visible note marker must have a measurable internal link.");
            TestAssert.Equal(expected, links.Select(link => link.Destination).Distinct().Count());
            foreach (PdfLinkAnnotation link in links)
            {
                PdfLinkDestination target = link.Destination!.Value;
                TestAssert.True(target.PageIndex >= 0 && target.PageIndex < pages.Length &&
                    target.Top > 0d && target.Top <= pages[target.PageIndex].Height,
                    "Note destinations must resolve to a rendered page location.");
            }
            if (fixture == "note-nav-endnote-page-two")
            {
                TestAssert.Equal(2, pages.Length);
                TestAssert.Equal(1, links[0].Destination!.Value.PageIndex);
            }
        }
    }

    public static void DocxNoteReferenceAnnotationsRequireVisibleSourceAndTarget()
    {
        DocxDocument source = ReadNoteNavigationFixture("note-nav-footnote-body");
        var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxDocument orphan = source with { RelatedStories = [] };
        TestAssert.Equal(0, renderer.RenderPagesAsync(orphan, null, CancellationToken.None).ToBlockingEnumerable().Sum(page => page.Annotations.Count));
        DocxParagraph paragraph = source.Paragraphs.Single();
        int markerRun = paragraph.InlineReferences.Single().SourceRunIndex;
        DocxDocument hidden = source with { BodyElements = [new DocxParagraphElement(paragraph with
            { Runs = paragraph.Runs.Where(run => run.SourceRunIndex != markerRun).ToArray() })] };
        TestAssert.Equal(0, renderer.RenderPagesAsync(hidden, null, CancellationToken.None).ToBlockingEnumerable().Sum(page => page.Annotations.Count));
    }

    public static void DocxReaderNoteStoryMarksFollowFilteredBodyLabels()
    {
        string seed = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-nav-two-footnotes.docx"));
        var parts = new Dictionary<string, string>();
        using (ZipArchive archive = ZipFile.OpenRead(seed))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                parts[entry.FullName] = reader.ReadToEnd();
            }
        }
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XDocument xml = XDocument.Parse(parts["word/document.xml"]);
        XElement marker = xml.Descendants(word + "footnoteReference").First().Parent!;
        marker.ReplaceWith(new XElement(word + "del", new XAttribute(word + "id", "9"),
            new XAttribute(word + "author", "Public reviewer"), new XElement(marker)));
        parts["word/document.xml"] = xml.ToString();
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try
        {
            foreach (OoxPdfDocxMarkupMode mode in new[]
                { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
            {
                DocxDocument document = DocxTests.ReadDocx(input, mode);
                DocxInlineReference[] references = document.Paragraphs.SelectMany(paragraph => paragraph.InlineReferences).ToArray();
                TestAssert.Equal(mode == OoxPdfDocxMarkupMode.Final ? 1 : 2, references.Length);
                foreach (DocxInlineReference reference in references)
                {
                    DocxRelatedStory story = document.RelatedStories.Single(story => story.Kind == reference.Kind && story.Id == reference.Id);
                    TestAssert.Equal(reference.DisplayText!, story.Paragraphs[0].Runs[0].Text);
                }
            }
        }
        finally { File.Delete(input); }
    }

    private static readonly string[] NoteNavigationFixtures =
    ["note-nav-footnote-body", "note-nav-footnote-table", "note-nav-endnote-body", "note-nav-endnote-page-two",
     "note-nav-two-footnotes", "note-nav-mixed-notes", "note-nav-numbered-ids", "note-nav-shared-run"];

    private static DocxDocument ReadNoteNavigationFixture(string fixture) => DocxTests.ReadDocx(
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", fixture + ".docx")),
        OoxPdfDocxMarkupMode.AllMarkup);

    public static void DocxWordCompatibleAllMarkupClampsCommentConnectorAnchorsAtPageEdge()
    {
        DocxTextRun run = new("Edge anchor", 10d, null, false, false, false, null, null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = 0
        };
        DocxParagraph paragraph = new(
            [run],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1.2d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null)
        {
            InlineReferences =
            [
                new DocxInlineReference(DocxRelatedStoryKind.Comment, "1", null, SourceRunIndex: 1, RunChildIndex: 0, TextOffsetInRun: 0, DisplayText: null)
            ],
            CommentRanges =
            [
                new DocxCommentRange("1", 0, 0, 0, 0, 1, 0)
            ]
        };
        DocxRelatedStory commentStory = new(
            DocxRelatedStoryKind.Comment,
            "/word/comments.xml",
            "1",
            [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public edge comment body", 10d, 12d))],
            [],
            [], null);
        // Near-edge clamp calibration: anchor = L - L * (1 - s) - 3.18 with s = 200 / 346.5,
        // so L = 4 lands at -0.87 and clamps to 0.5 (L = 10 no longer reaches the edge
        // under the lane-fit print scale).
        DocxDocument document = new(
            200d,
            240d,
            4d,
            120d,
            20d,
            20d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(paragraph)],
            [],
            [])
        {
            RelatedStories = [commentStory],
            MarkupMode = OoxPdfDocxMarkupMode.AllMarkup
        };
        var renderer = new DocxRenderer(
            fontResolver: null,
            markupMode: OoxPdfDocxMarkupMode.AllMarkup,
            markupGeometryMode: OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

        DocxMarkupBalloonPlacementSnapshot placement = renderer.InspectMarkupBalloons(document)
            .Single(item => item.Kind == "Comment");

        TestAssert.True(
            placement.AnchorConnectorX >= 0.5d &&
            placement.AnchorConnectorX < 1d,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Word-compatible all-markup should clamp near-edge comment connector anchors inside the page media box. AnchorX={placement.AnchorConnectorX}."));
        TestAssert.True(placement.AnchorConnectorClamped, "Near-edge connector placement snapshots should expose that the anchor was clamped.");
    }

    public static void DocxReaderSplitsNestedVisibleInlineContainerPageBreaks()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/break" TargetMode="External"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:hyperlink r:id="rIdLink">
                        <w:ins>
                          <w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After</w:t></w:r>
                        </w:ins>
                      </w:hyperlink>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal(3, document.BodyElements.Count);
        TestAssert.True(document.BodyElements[0] is DocxParagraphElement, "Text before the nested break should remain a paragraph fragment.");
        TestAssert.True(document.BodyElements[1] is DocxPageBreakElement, "Nested page breaks should become explicit body break elements.");
        TestAssert.True(document.BodyElements[2] is DocxParagraphElement, "Text after the nested break should remain a paragraph fragment.");
        DocxParagraph before = ((DocxParagraphElement)document.BodyElements[0]).Paragraph;
        DocxParagraph after = ((DocxParagraphElement)document.BodyElements[2]).Paragraph;
        TestAssert.Equal("Before", string.Concat(before.Runs.Select(run => run.Text)));
        TestAssert.Equal("After", string.Concat(after.Runs.Select(run => run.Text)));
        TestAssert.Equal(1, before.Hyperlinks.Count);
        TestAssert.Equal(1, after.Hyperlinks.Count);
    }

    public static void DocxReaderSplitsMoveToFinalViewPageBreaks()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:moveTo w:id="9" w:author="Author" w:date="2026-06-02T00:00:00Z">
                        <w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After</w:t></w:r>
                      </w:moveTo>
                    </w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        DocxDocument document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal(3, document.BodyElements.Count);
        TestAssert.True(document.BodyElements[0] is DocxParagraphElement, "Moved-to text before the break should remain a paragraph fragment.");
        TestAssert.True(document.BodyElements[1] is DocxPageBreakElement, "Moved-to page breaks should become explicit body break elements.");
        TestAssert.True(document.BodyElements[2] is DocxParagraphElement, "Moved-to text after the break should remain a paragraph fragment.");
        TestAssert.Equal("Before", string.Concat(((DocxParagraphElement)document.BodyElements[0]).Paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal("After", string.Concat(((DocxParagraphElement)document.BodyElements[2]).Paragraph.Runs.Select(run => run.Text)));
    }

    public static void DocxLayoutReservesPageFootnoteStoryArea()
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Anchor paragraph", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "7",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 6)
            ]
        };
        DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler paragraph keeps body close to the bottom", 10d, 12d);
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body line one wraps line two", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "7",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            112d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxParagraphElement(anchor),
                new DocxParagraphElement(filler),
                new DocxParagraphElement(filler),
                new DocxParagraphElement(filler),
                new DocxParagraphElement(filler)
            ],
            [],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutPageSnapshot footnotePage = snapshot.Pages.Single(page => page.PlacedFootnoteStoryCount == 1);
        double footnoteTop = footnotePage.PlacedRelatedItems.Max(item => item.Y + item.Height);
        double bodyBottom = footnotePage.Items.Min(item => item.Y);

        TestAssert.True(footnotePage.PlacedRelatedStoryTextLineCount >= 1, "Resolved footnote story text should be placed on the body-reference page.");
        TestAssert.True(bodyBottom >= footnoteTop, $"Body layout should reserve page space above placed footnotes; body bottom {bodyBottom} footnote top {footnoteTop}.");
    }

    public static void DocxFootnoteStoryLaysOutUnscaledDesignSpace()
    {
        // RV06 anchor probe (edge-footanchor-5, Word 16.0): footnote stories lay
        // out unscaled in design space while emission maps uniformly (WC first
        // baseline 236.62 = affine-mapped design). Pre-fix the word-compatible
        // story layout scaled metrics in place.
        DocxParagraph SizedLine(string text, double fontSize)
        {
            return DocxTests.CreateDocxLayoutParagraph(text, fontSize, 12d) with
            {
                LineSpacingPoints = null
            };
        }

        double[] Pitches(OoxPdfDocxMarkupGeometryMode mode, double scale)
        {
            DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d) with
            {
                InlineReferences =
                [
                    new DocxInlineReference(
                        DocxRelatedStoryKind.Footnote,
                        "7",
                        CustomMarkFollowsValue: null,
                        DisplayText: "1",
                        SourceRunIndex: 0,
                        RunChildIndex: 1,
                        TextOffsetInRun: 6)
                ]
            };
            var footnoteStory = new DocxRelatedStory(
                DocxRelatedStoryKind.Footnote,
                "/word/footnotes.xml",
                "7",
                [
                    new DocxParagraphElement(SizedLine("Alpha", 12d)),
                    new DocxParagraphElement(SizedLine("Beta", 15d)),
                    new DocxParagraphElement(SizedLine("Gamma", 12d))
                ],
                [],
                [],
                null);
            var document = new DocxDocument(
                220d,
                200d,
                10d,
                10d,
                10d,
                10d,
                DocxPageSettings.Empty,
                [],
                [],
                [],
                [new DocxParagraphElement(anchor)],
                [],
                [])
            {
                RelatedStories = [footnoteStory]
            };
            DocxTextLineLayout[] lines = new DocxLayoutEngine(mode, scale)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
                .Pages.SelectMany(page => page.PlacedRelatedStories).SelectMany(story => story.TextLines).ToArray();
            TestAssert.Equal(3, lines.Length);
            return [lines[0].BaselineY - lines[1].BaselineY, lines[1].BaselineY - lines[2].BaselineY];
        }

        // Footnote stories lay out unscaled in design space (emission maps
        // uniformly); the scaled layout keeps full unscaled transitions.
        double[] preserve = Pitches(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, 1d);
        double[] scaled = Pitches(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, 0.75d);
        TestAssert.True(
            Math.Abs(scaled[0] - preserve[0]) < 0.001d,
            $"Footnote story layout should stay design space. preserve={preserve[0]} scaled={scaled[0]}.");
        TestAssert.True(
            Math.Abs(scaled[1] - preserve[1]) < 0.001d,
            $"Footnote story layout should stay design space. preserve={preserve[1]} scaled={scaled[1]}.");
    }

    public static void DocxLayoutStacksMultipleFootnotesOnOnePage()
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Anchor paragraph with two footnote markers", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "31",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 0,
                    TextOffsetInRun: 8),
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "32",
                    CustomMarkFollowsValue: null,
                    DisplayText: "2",
                    SourceRunIndex: 0,
                    RunChildIndex: 0,
                    TextOffsetInRun: 22)
            ]
        };
        DocxParagraph firstFootnote = DocxTests.CreateDocxLayoutParagraph("First footnote body", 10d, 12d);
        DocxParagraph secondFootnote = DocxTests.CreateDocxLayoutParagraph("Second footnote body", 10d, 12d);
        var firstStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "31",
            [new DocxParagraphElement(firstFootnote)],
            [],
            [], null);
        var secondStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "32",
            [new DocxParagraphElement(secondFootnote)],
            [],
            [], null);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(anchor)], []) with
        {
            RelatedStories = [firstStory, secondStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutPageSnapshot footnotePage = snapshot.Pages.Single(page => page.PlacedFootnoteStoryCount == 2);
        DocxPlacedRelatedStoryLayoutSnapshot[] placedStories = footnotePage.PlacedRelatedStories.ToArray();
        DocxPlacedRelatedStoryLayoutSnapshot firstPlaced = placedStories.Single(story => story.Id == "31");
        DocxPlacedRelatedStoryLayoutSnapshot secondPlaced = placedStories.Single(story => story.Id == "32");

        TestAssert.True(firstPlaced.TopY > secondPlaced.TopY, "The first footnote body should be placed above the later footnote body instead of both stories sharing the page bottom.");
        TestAssert.True(secondPlaced.TopY <= firstPlaced.TopY - firstPlaced.Height + 0.001d, "Stacked footnote bodies should not overlap vertically.");
        TestAssert.True(firstPlaced.SeparatorY is not null && secondPlaced.SeparatorY is null, "Generic separator geometry should be emitted once above the stacked footnote group.");
    }

    public static void DocxRendererClipsOverlongPlacedFootnoteStoryToPageRegion()
    {
        DocxParagraph bodyParagraph = DocxTests.CreateDocxLayoutParagraph("body with long footnote", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "16",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote overflow body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "16",
            Enumerable.Range(0, 8).Select(_ => new DocxParagraphElement(footnoteParagraph)).Cast<DocxBodyElement>().ToArray(),
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(bodyParagraph)],
            [bodyParagraph],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxPlacedRelatedStoryLayoutSnapshot placedStory = snapshot.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .First(story => story.Kind == "Footnote" && story.Id == "16");
        // Overlong notes slice across pages now; the head slice keeps full-story height with a page-slice height.
        PdfPage placedPage = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable()
            .First(page => page.Content.Contains(" re W n", StringComparison.Ordinal));

        TestAssert.True(placedStory.ContentHeight > placedStory.Height, "The overlong footnote should retain full-story height while exposing the clipped page slice height.");
        TestAssert.Equal(0d, placedStory.ContentTopOffset);
        TestAssert.Contains(" re W n", placedPage.Content);
    }

    public static void DocxLayoutPlacesFootnoteOnRenderedMarkerRunPage()
    {
        var leadingRun = new DocxTextRun(
            string.Concat(Enumerable.Repeat("alpha beta gamma delta ", 12)),
            10d,
            null,
            false,
            false,
            false,
            null,
            null);
        var markerRun = new DocxTextRun(
            "marker",
            10d,
            null,
            false,
            false,
            false,
            null,
            null);
        DocxParagraph anchor = new(
            [leadingRun, markerRun],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null)
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "11",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 1,
                    RunChildIndex: 1,
                    TextOffsetInRun: 0)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "11",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(anchor)],
            [anchor],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        int footnotePageIndex = Array.FindIndex(snapshot.Pages.ToArray(), page => page.PlacedFootnoteStoryCount == 1);

        TestAssert.True(layout.Pages.Count > 1, "The marker paragraph should span pages so note ownership is not reducible to source-block ownership.");
        TestAssert.True(layout.Pages[0].Items.OfType<DocxTextLineLayout>().Any(line => line.SourceBlockIndex == 0), "The first page should contain the source paragraph before the marker run is rendered.");
        TestAssert.Equal(0, snapshot.Pages[0].PlacedFootnoteStoryCount);
        TestAssert.True(footnotePageIndex > 0, "The footnote story should be placed on the page where the marker run is rendered, not the first page containing the source paragraph.");
        TestAssert.True(layout.Pages[footnotePageIndex].Items.OfType<DocxTextLineLayout>().Any(line => line.Segments.Any(segment => segment.SourceTextRunIndex == 1)), "The selected footnote page should contain the marker source run.");
    }

    public static void DocxLayoutPlacesFootnoteOnRenderedMarkerOffsetPage()
    {
        const int markerOffset = 180;
        string text = string.Concat(Enumerable.Repeat("alpha beta gamma delta ", 12));
        var run = new DocxTextRun(
            text,
            10d,
            null,
            false,
            false,
            false,
            null,
            null);
        DocxParagraph anchor = new(
            [run],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null)
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "12",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: markerOffset)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "12",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(anchor)],
            [anchor],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        int footnotePageIndex = Array.FindIndex(snapshot.Pages.ToArray(), page => page.PlacedFootnoteStoryCount == 1);

        TestAssert.True(layout.Pages.Count > 1, "The single marker run should span pages so note ownership requires source offsets.");
        TestAssert.Equal(0, snapshot.Pages[0].PlacedFootnoteStoryCount);
        TestAssert.True(footnotePageIndex > 0, "The footnote story should be placed on the page containing the marker offset, not the first page containing the source run.");
        TestAssert.True(
            layout.Pages[footnotePageIndex].Items.OfType<DocxTextLineLayout>().Any(line =>
                line.Segments.Any(segment =>
                    segment.SourceTextRunIndex == 0 &&
                    markerOffset >= segment.SourceTextOffsetInRun &&
                    markerOffset <= segment.SourceTextOffsetInRun + segment.Text.Length)),
            "The selected footnote page should contain the marker source offset inside the rendered run segment.");
    }

    public static void DocxLayoutPlacesFootnoteOnReaderProvenanceMarkerOffsetPage()
    {
        string leadingText = string.Concat(Enumerable.Repeat("alpha beta gamma delta ", 12));
        const string markerText = "1";
        int markerOffset = leadingText.Length;
        var leadingRun = new DocxTextRun(
            leadingText,
            10d,
            null,
            false,
            false,
            false,
            null,
            null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = 0
        };
        var markerRun = new DocxTextRun(
            markerText,
            10d,
            null,
            false,
            false,
            false,
            null,
            null,
            0d,
            false,
            "superscript",
            false,
            null,
            false,
            null,
            null,
            null,
            null,
            null,
            false,
            null,
            false,
            null,
            null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = markerOffset
        };
        var trailingRun = new DocxTextRun(
            " trailing text after marker",
            10d,
            null,
            false,
            false,
            false,
            null,
            null)
        {
            SourceRunIndex = 0,
            SourceTextOffsetInRun = markerOffset
        };
        DocxParagraph anchor = new(
            [leadingRun, markerRun, trailingRun],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            12d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null)
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "13",
                    CustomMarkFollowsValue: null,
                    DisplayText: markerText,
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: markerOffset)
            ]
        };
        DocxParagraph footnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Footnote body", 10d, 12d);
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "13",
            [new DocxParagraphElement(footnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            220d,
            82d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(anchor)],
            [anchor],
            [])
        {
            RelatedStories = [footnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        int footnotePageIndex = Array.FindIndex(snapshot.Pages.ToArray(), page => page.PlacedFootnoteStoryCount == 1);

        TestAssert.True(layout.Pages.Count > 1, "The normalized run sequence should span pages so source provenance is required.");
        TestAssert.Equal(0, snapshot.Pages[0].PlacedFootnoteStoryCount);
        TestAssert.True(footnotePageIndex > 0, "The footnote story should be placed on the page containing the normalized marker run, not the first page containing earlier text from the same source run.");
        TestAssert.True(
            layout.Pages[footnotePageIndex].Items.OfType<DocxTextLineLayout>().Any(line =>
                line.Segments.Any(segment =>
                    segment.Text == markerText &&
                    segment.SourceTextRunIndex == 0 &&
                    segment.SourceTextOffsetInRun == markerOffset)),
            "The selected footnote page should contain the normalized marker display run with original OOXML source provenance.");
    }

    public static void RelatedStoryPageIndexMatchesLegacyMatching()
    {
        // R12: the once-per-pass page index answers reference/block queries exactly
        // like the legacy per-check page walks: null-tolerance and identity paragraph
        // matching, offset and run-fallback segment checks, vacuous truth for
        // never-rendered runs, and last-page section ranges.
        DocxParagraph source = MakeNoteParagraph();
        DocxParagraph other = MakeNoteParagraph();
        DocxTextLineLayout lineA = MakeOwnerLine(null, 0, 7, "hello");
        DocxTextLineLayout lineB = MakeOwnerLine(other, 0, 7, "world");
        DocxTextLineLayout lineC = MakeOwnerLine(null, 1, 7, "!");
        DocxLayoutPage page0 = MakeOwnerPage([lineA]);
        DocxLayoutPage page1 = MakeOwnerPage([lineB, lineC]);
        DocxLayoutEngine.RelatedStoryPageIndex index = DocxLayoutEngine.RelatedStoryPageIndex.Build([page0, page1], CancellationToken.None);

        DocxInlineReference reference = new(DocxRelatedStoryKind.Footnote, "1", null, null, SourceRunIndex: 7, RunChildIndex: 0, TextOffsetInRun: 2);
        var location = new DocxLayoutEngine.DocxInlineReferenceLocation(0, source, reference);
        TestAssert.True(index.IsReferenceRenderedOnPage(0, location), "Null-tolerance line must match any paragraph of its block.");
        TestAssert.True(!index.IsReferenceRenderedOnPage(1, location), "Identity mismatch and block mismatch must reject page 1.");
        TestAssert.Equal(0, index.FindFirstPageWithReference(location));
        TestAssert.Equal(0, index.FindFirstPageWithReference(location, requireRendered: true));

        var block1 = new DocxLayoutEngine.DocxInlineReferenceLocation(1, source, reference with { TextOffsetInRun = 0 });
        TestAssert.True(!index.IsReferenceRenderedOnPage(0, block1), "Block mismatch must reject page 0.");
        TestAssert.True(index.IsReferenceRenderedOnPage(1, block1), "Null-tolerance line must match block 1 on page 1.");
        TestAssert.Equal(1, index.FindFirstPageWithReference(block1, requireRendered: true));

        var missingBlock = new DocxLayoutEngine.DocxInlineReferenceLocation(9, source, reference);
        TestAssert.True(index.IsReferenceRenderedOnPage(0, missingBlock), "Blocks rendered nowhere stay vacuously true (legacy).");
        TestAssert.True(index.IsReferenceRenderedOnPage(1, missingBlock), "Blocks rendered nowhere stay vacuously true (legacy).");
        TestAssert.Equal(0, index.FindFirstPageWithReference(missingBlock));
        TestAssert.Equal(-1, index.FindFirstPageWithReference(missingBlock, requireRendered: true));

        var unrenderedRun = new DocxLayoutEngine.DocxInlineReferenceLocation(0, source, reference with { SourceRunIndex = 99 });
        TestAssert.True(index.IsReferenceRenderedOnPage(1, unrenderedRun), "Never-rendered runs stay vacuously true (legacy).");
        TestAssert.Equal(-1, index.FindFirstPageWithReference(unrenderedRun, requireRendered: true));
        var negativeRun = new DocxLayoutEngine.DocxInlineReferenceLocation(0, source, reference with { SourceRunIndex = -1 });
        TestAssert.True(index.IsReferenceRenderedOnPage(1, negativeRun), "Negative runs stay vacuously true (legacy).");
        TestAssert.Equal(-1, index.FindFirstPageWithReference(negativeRun, requireRendered: true));

        int[] page0Blocks = index.SortedBlocks(0);
        TestAssert.Equal(1, page0Blocks.Length);
        TestAssert.Equal(0, page0Blocks[0]);
        int[] page1Blocks = index.SortedBlocks(1);
        TestAssert.Equal(2, page1Blocks.Length);
        TestAssert.Equal(0, page1Blocks[0]);
        TestAssert.Equal(1, page1Blocks[1]);
        TestAssert.Equal(1, index.FindLastPageWithBlockInRange(0, 0));
        TestAssert.Equal(1, index.FindLastPageWithBlockInRange(0, 1));
        TestAssert.Equal(-1, index.FindLastPageWithBlockInRange(5, 9));
    }

    private static DocxParagraph MakeNoteParagraph()
    {
        return new DocxParagraph(
            [],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1.2d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
    }

    private static DocxTextLineLayout MakeOwnerLine(DocxParagraph? paragraph, int block, int run, string text)
    {
        var styleRun = new DocxTextRun(text, 10d, null, false, false, false, null, null);
        var segment = new DocxTextSegmentLayout(
            text,
            styleRun,
            0d,
            text.Length * 5d,
            10d,
            0d,
            0d,
            default,
            false,
            run,
            0,
            DocxTextSegmentRole.Text);
        return new DocxTextLineLayout(
            Text: text,
            StyleRun: styleRun,
            FontSize: 10d,
            X: 0d,
            BaselineY: 10d,
            Width: text.Length * 5d,
            Segments: [segment],
            SourceBlockIndex: block,
            SourceParagraphIndex: null,
            SourceLineIndex: null,
            Story: null,
            LineHeight: null,
            AppliedBeforeSpacing: null,
            IsFirstParagraphLine: null,
            EndsWithIntraTokenBreak: false,
            SingleLineHeight: null,
            ListLabelSingleLineHeight: null,
            BodyWindowsLineHeight: null,
            ListLabelWindowsLineHeight: null,
            EffectiveLineSpacingFactor: null,
            LineSpacingFactorFloorApplied: null,
            PendingAfterSpacing: null,
            ParagraphBeforeSpacing: null,
            ParagraphAfterSpacing: null,
            ContextualSpacingSuppressed: null,
            SourceParagraph: paragraph,
            LineHeightSource: null,
            EmitsTerminalParagraphMark: false);
    }

    private static DocxLayoutPage MakeOwnerPage(IReadOnlyList<DocxTextLineLayout> lines)
    {
        return new DocxLayoutPage(
            612d,
            792d,
            72d,
            72d,
            0d,
            72d,
            72d,
            DocxPageSettings.Empty,
            new DocxSectionLayoutProperties(null, null, null, null, null, null, []),
            [],
            [],
            [],
            [],
            [],
            lines);
    }

    public static void DocxFootnoteTrailingLinesShareDrawableLinePositions()
    {
        // RV06 footnote-align probe: Office centers/rights drawable footnote text and
        // keeps one row-end space beyond authored trailing in footnotes too.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">F trail   </w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">F trail</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxTextLineLayout[] footnoteLines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .SelectMany(story => story.TextLines)
            .ToArray();
        DocxTextLineLayout trailLine = footnoteLines.Single(line => line.Text.EndsWith("   ", StringComparison.Ordinal));
        DocxTextLineLayout cleanLine = footnoteLines.Single(line => line.Text == "F trail");
        TestAssert.Equal("F trail    ", trailLine.Text);
        TestAssert.Equal(cleanLine.X, trailLine.X);
    }

    public static void DocxFootnoteSeparatorMarkSitsAtRuleEnd()
    {
        // RV06 footnote probe: Office renders the separator mark as a space at the
        // rule end (body-left plus the 144pt rule width).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        TestAssert.Equal(1, separatorLine.Segments.Count);
        TestAssert.Equal(" ", separatorLine.Segments[0].Text);
        TestAssert.Equal(144d, separatorLine.Segments[0].X - separatorLine.X);
    }

    private sealed class StrikeoutFontResolver(int strikeoutSize = 50, int strikeoutPosition = 300) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            byte[] faceBytes = TestFontBuilder.CreateTestFont();
            PatchStrikeoutMetrics(faceBytes, (ushort)strikeoutSize, (short)strikeoutPosition);
            return new FontFaceResolution(
                request.FamilyName,
                "StrikeFace",
                new FontStyleKey(request.Bold, request.Italic),
                new MemoryFontProgramSource("test:strikeface", faceBytes),
                IsFallback: false);
        }

        private static void PatchStrikeoutMetrics(byte[] faceBytes, ushort size, short position)
        {
            // The shared synthetic face stores a placeholder strikeout size, so pin
            // real OS/2 strikeout geometry (0.05em size, 0.30em position) in place.
            int tableCount = (faceBytes[4] << 8) | faceBytes[5];
            for (int tableIndex = 0; tableIndex < tableCount; tableIndex++)
            {
                int record = 12 + 16 * tableIndex;
                if (faceBytes[record] == 0x4F && faceBytes[record + 1] == 0x53 && faceBytes[record + 2] == 0x2F && faceBytes[record + 3] == 0x32)
                {
                    int offset = (faceBytes[record + 8] << 24) | (faceBytes[record + 9] << 16) | (faceBytes[record + 10] << 8) | faceBytes[record + 11];
                    faceBytes[offset + 26] = (byte)(size >> 8);
                    faceBytes[offset + 27] = (byte)(size & 0xFF);
                    faceBytes[offset + 28] = (byte)((position >> 8) & 0xFF);
                    faceBytes[offset + 29] = (byte)(position & 0xFF);
                    return;
                }
            }

            throw new InvalidOperationException("Synthetic test font is missing the OS/2 table.");
        }
    }

    public static void DocxFootnoteSeparatorRuleFollowsStrikeoutMetrics()
    {
        // RV06 separator probes (Word 16.0, Times, Aptos and Calibri at 10, 12 and 14pt):
        // the footnote separator rule follows OS/2 strikeout geometry, with the rule top
        // at the strikeout position and the thickness at the strikeout size, instead of
                // The patched synthetic face pins strikeout 0.30em and 0.05em,
        // so the rule bottom must sit 0.25em above the separator bottom while the mark baseline
        // rides at the separator bottom with no extra ride.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = DocxFontPlanTextMeasurer.CreateAsync(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver).GetAwaiter().GetResult();
        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        double markSize = separatorLine.Segments[0].StyleRun.EffectiveProperties.FontSize;
        double ruleY = separator.SeparatorY ?? double.NaN;
        TestAssert.True(Math.Abs(ruleY - separatorLine.BaselineY - 0.25d * markSize) < 0.000001d, "Footnote rule bottom must sit 0.25em above the mark baseline with no ride.");
        TestAssert.True(Math.Abs(separator.SeparatorThickness - 0.05d * markSize) < 0.000001d, "Footnote rule thickness must follow the strikeout size.");
    }

    public static void DocxFootnoteSeparatorRuleSnapsThicknessToPixelGrid()
    {
        // RV06 rule-thickness probes (Word 16.0, Times grids plus Tahoma slash Calibri 14pt):
        // Office rule thickness snaps strikeout size to whole 600dpi pixels (seven points exact),
        // so a 0.067em synthetic face at explicit 10pt marks renders 0.60 instead of 0.67,
                // The patched synthetic face pins strikeout 0.25em position and 0.067em size at
        // explicit 10pt marks, so the thickness must snap 0.67 to the 5-pixel 0.60.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver(67, 250);
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = DocxFontPlanTextMeasurer.CreateAsync(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver).GetAwaiter().GetResult();
        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        double markSize = separatorLine.Segments[0].StyleRun.EffectiveProperties.FontSize;
        double ruleY = separator.SeparatorY ?? double.NaN;
        TestAssert.True(Math.Abs(markSize - 10d) < 0.000001d, "Separator marks must resolve at explicit 10pt; markSize=" + markSize.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(separator.SeparatorThickness - 0.72d) < 0.000001d, "Rule thickness must snap 0.67 to the 6-pixel 0.72; thickness=" + separator.SeparatorThickness.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteSeparatorRuleKeepsLegacyConstantsWithoutStrikeoutMetrics()
    {
        // Measurers without strikeout metrics keep the legacy footnote constants.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        DocxTextLineLayout separatorLine = separator.TextLines.Single();
        double ruleY = separator.SeparatorY ?? double.NaN;
        TestAssert.True(Math.Abs(ruleY - separatorLine.BaselineY - 2.1d) < 0.000001d, "Legacy footnote rule offset must stay 2.1pt above the separator bottom with no ride.");
        TestAssert.Equal(0.75d, separator.SeparatorThickness);
    }

    public static void DocxWordCompatibleAllMarkupMapsSeparatorRuleThickness()
    {
        // RV06 endnote probes: word-compatible separator rules map uniformly to
        // emission space like the surrounding story text, but the rule thickness
        // still emits at design size while the rule width is print-scaled.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:commentRangeStart w:id="0"/></w:r><w:r><w:t xml:space="preserve">commented</w:t></w:r><w:r><w:commentRangeEnd w:id="0"/></w:r><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>""",
            ["word/comments.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Reviewer" w:initials="R" w:date="2024-01-02T00:00:00Z"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"""
        });;
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.AllMarkup);
        }
        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = DocxFontPlanTextMeasurer.CreateAsync(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver).GetAwaiter().GetResult();
        DocxPlacedRelatedStoryLayout separator = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, measurer, CancellationToken.None)
            .Pages[0]
            .PlacedRelatedStories
            .Single(story => story.SeparatorY is not null);
        double designThickness = separator.SeparatorThickness;
        double designWidth = Math.Min(144d, separator.Width);
        var renderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        PdfPage page = renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single();
        Match[] ruleMatches = Regex.Matches(page.Content, @"(?<x>-?[0-9.]+) (?<y>-?[0-9.]+) (?<w>[0-9.]+) (?<h>[0-9.]+) re f[^*]")
            .Cast<Match>()
            .Where(match => double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture) < 200d && double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture) > 50d)
            .ToArray();
        TestAssert.Equal(1, ruleMatches.Length);
        double emitWidth = double.Parse(ruleMatches[0].Groups["w"].Value, CultureInfo.InvariantCulture);
        double emitHeight = double.Parse(ruleMatches[0].Groups["h"].Value, CultureInfo.InvariantCulture);
        double printScale = emitWidth / designWidth;
        TestAssert.True(printScale < 0.999d, "Separator rule width must carry the print scale, proving the emission map is active.");
        TestAssert.True(Math.Abs(emitHeight - designThickness * printScale) < 0.002d, "Word-compatible separator rule thickness must scale with the print map.");
    }



    public static void DocxWordCompatibleContinuationRuleSpansDesignBodyWidth()
    {
        // RV06 wclong probe (Word 16.0): word-compatible continuation rules span the
        // full design body mapped once (Office 354.98 emitted), while the renderer
        // emits the shrunk layout body mapped again (269.49 = shrunk times scale).
        // The continuation rule width must be the design body (layout body plus the
        // markup reserve) with the end mark at the rule end.
        var footnoteParas = new StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Note body line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:commentRangeStart w:id="0"/></w:r><w:r><w:t xml:space="preserve">commented</w:t></w:r><w:r><w:commentRangeEnd w:id="0"/></w:r><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>""",
            ["word/comments.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Reviewer" w:initials="R" w:date="2024-01-02T00:00:00Z"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.AllMarkup);
        }
        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = DocxFontPlanTextMeasurer.CreateAsync(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver).GetAwaiter().GetResult();
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .Create(document, measurer, CancellationToken.None);
        var continuations = layout.Pages.SelectMany(page => page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator &&
                story.SeparatorY is not null)
            .Select(story => (Page: page, Story: story))).ToArray();
        TestAssert.True(continuations.Length != 0, "Overflowing WC footnotes must place continuation separators.");
        foreach ((DocxLayoutPage page, DocxPlacedRelatedStoryLayout continuation) in continuations)
        {
            double designBody = page.Width - page.MarginLeft - page.MarginRight + page.MarkupMarginReservePoints;
            TestAssert.True(Math.Abs(continuation.SeparatorWidth - designBody) < 0.01d, "WC continuation rule width must be the design body width.");
            TestAssert.True(Math.Abs(continuation.Width - designBody) < 0.01d, "WC continuation story width must cover the design body for the rule clip.");
            if (continuation.TextLines.Count == 1 && continuation.TextLines[0].Segments.Count == 1 &&
                string.IsNullOrWhiteSpace(continuation.TextLines[0].Segments[0].Text))
            {
                double markX = continuation.TextLines[0].Segments[0].X;
                TestAssert.True(Math.Abs(markX - (continuation.X + designBody)) < 0.01d, "WC continuation end mark must sit at the rule end.");
            }
        }
        // Reserve-margin pages keep identity emission over the shrunk body, so the
        // reserve must not join the continuation rule back there.
        DocxLayout reserveLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin)
            .Create(document, measurer, CancellationToken.None);
        var reserveContinuations = reserveLayout.Pages.SelectMany(page => page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator &&
                story.SeparatorY is not null)
            .Select(story => (Page: page, Story: story))).ToArray();
        TestAssert.True(reserveContinuations.Length != 0, "Overflowing reserve-margin footnotes must place continuation separators.");
        foreach ((DocxLayoutPage reservePage, DocxPlacedRelatedStoryLayout reserveContinuation) in reserveContinuations)
        {
            double shrunkBody = reservePage.Width - reservePage.MarginLeft - reservePage.MarginRight;
            TestAssert.True(Math.Abs(reserveContinuation.SeparatorWidth - shrunkBody) < 0.01d, "Reserve-margin continuation rule width must stay in the shrunk body.");
        }
    }

    public static void DocxLongFootnoteContinuationReservesHeaderZone()
    {
        // RV06 wclong header probes (Word 16.0, h1/h5/nohdr): continued notes leave
        // the static header zone clear (h5 takes 22 vs 25 blind with the rule below
        // all header baselines; nohdr agrees 25/25), while the renderer starts
        // continuation pages at the full page top so the rule lands inside the
        // header text. Every continuation rule must sit below the header baselines
        // of its own page.
        var headerParas = new StringBuilder();
        for (int line = 0; line < 5; line++)
        {
            headerParas.Append("<w:p><w:r><w:t>Header line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        var footnoteParas = new StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Note body line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/header1.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">""" + headerParas.ToString() + """</w:hdr>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        var continuations = layout.Pages.SelectMany(page => page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator &&
                story.SeparatorY is not null)
            .Select(story => (Page: page, Story: story))).ToArray();
        TestAssert.True(continuations.Length != 0, "Overflowing footnotes with headers must place continuation separators.");
        foreach ((DocxLayoutPage page, DocxPlacedRelatedStoryLayout continuation) in continuations)
        {
            double[] headerBaselines = page.StaticTextLines.Select(line => line.BaselineY).ToArray();
            TestAssert.True(headerBaselines.Length != 0, "Continuation pages must carry the static header.");
            double lowestHeaderBaseline = headerBaselines.Min();
            TestAssert.True(continuation.SeparatorY <= lowestHeaderBaseline - 1d, "Continuation rules must sit below the header baselines of their page.");
        }
    }
    public static void DocxLongFootnoteHeadTakeMatchesOfficeWithAptosMetrics()
    {
        // RV06 p1-clamp probes (Word 16.0, h1/b-sweep): the head page takes 24 note
        // lines (2 mixed-size plus long 0..21) with full-width rules below the body,
        // with its document-default resolver explicitly pinned to Aptos. Synthetic
        // metrics take frame-exact counts that cannot see the real-advance shortfall,
        // so this pins the Office take with installed Aptos advances and skips without
        // a usable Aptos face like other font-environmental tests.
        if (!HasUsableAptosFootnoteFace())
        {
            TestAssert.Skip("Environmental precondition not met: Aptos is not installed.");
        }

        DocxLayout layout = CreateAptosContinuationHeadLayout();
        var headLines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .ToArray();
        var longIndexes = headLines
            .Select(line => line.Text)
            .Select(text => text.StartsWith("Long note line ", StringComparison.Ordinal) && int.TryParse(text.Substring(15).Split(' ')[0], out int index) ? index : -1)
            .Where(index => index >= 0)
            .ToArray();
        DocxTextLineLayout firstLine = headLines[0];        string diagnosis = "headLines=" + headLines.Length + ",long=[" + string.Join(",", longIndexes.Take(3)) + ".." + (longIndexes.Length == 0 ? "none" : longIndexes[^1].ToString(CultureInfo.InvariantCulture)) + "],firstLine(h=" + (firstLine.LineHeight?.ToString("F3", CultureInfo.InvariantCulture) ?? "?") + ",single=" + (firstLine.SingleLineHeight?.ToString("F3", CultureInfo.InvariantCulture) ?? "?") + ",factor=" + (firstLine.EffectiveLineSpacingFactor?.ToString("F4", CultureInfo.InvariantCulture) ?? "?") + ",after=" + (firstLine.ParagraphAfterSpacing?.ToString("F3", CultureInfo.InvariantCulture) ?? "?") + ")";
        TestAssert.Equal(24, headLines.Length);
        TestAssert.True(longIndexes.Length == 22 && longIndexes[0] == 0 && longIndexes[^1] == 21, "Head page must take long lines 0..21 like Office. " + diagnosis);
    }

    public static void DocxLongFootnoteHeadBlockSitsAtOfficePosition()
    {
        // RV06 p1-position probes (Word 16.0, h1/b-sweep): with takes exact, the head
        // block must also sit at Office absolute positions (h1-Final rule 682.54 with
        // mixed1 at 664.18). Fresh C153 exports retain both positions. Windows-only
        // Aptos pins this like the take test; the numerical thresholds stay unchanged.
        if (!HasUsableAptosFootnoteFace())
        {
            TestAssert.Skip("Environmental precondition not met: Aptos is not installed.");
        }

        DocxLayout layout = CreateAptosContinuationHeadLayout();
        DocxLayoutPage page = layout.Pages[0];
        DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(story => story.SeparatorY is not null);
        DocxTextLineLayout mixedFirst = page.PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .First(line => line.Text.StartsWith("Footnote twelve opening", StringComparison.Ordinal));
        TestAssert.True(Math.Abs((separator.SeparatorY ?? double.NaN) - 682.54d) < 1.5d, "Head separator rule must sit at the Office position; observed rule bottom=" + (separator.SeparatorY?.ToString(CultureInfo.InvariantCulture) ?? "?") + ".");
        TestAssert.True(Math.Abs(mixedFirst.BaselineY - 664.18d) < 1.5d, "Head footnote content must start at the Office position; observed mixed1 baseline=" + mixedFirst.BaselineY.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static DocxLayout CreateAptosContinuationHeadLayout()
    {
        // Keep the original style-less Word control. Resolve its document default
        // explicitly: a measurer fallback cannot override a resolved source face.
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "CompositionProbes", "aptos-continuation-head.docx"));
        TestAssert.Equal("12A7AF7C5EB0B12D20F05C383EFEFD8A312C4E327342BF485A368C2695DBCC33",
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(input))));
        var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
        var resolver = new AptosHeadControlFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        TestAssert.True(fontPlan.Runs.All(run => run.ResolvedFamily == "Aptos"),
            "The Office Aptos control must measure every source/default/separator owner with Aptos.");
        var measurer = DocxFontPlanTextMeasurer.CreateAsync(fontPlan, resolver.Resolve(new FontRequest("Aptos")), CancellationToken.None, resolver).GetAwaiter().GetResult();
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None);
    }

    private sealed class AptosHeadControlFontResolver : IFontResolver
    {
        private readonly WindowsFontResolver installed = new();

        public FontFaceResolution Resolve(FontRequest request) => installed.Resolve(
            request.FamilyName == DocxRenderer.DefaultDocumentTypefaceRequest
                ? request with { FamilyName = "Aptos" } : request);
    }

    private static bool HasUsableAptosFootnoteFace()
    {
        try
        {
            FontFaceResolution resolved = new WindowsFontResolver().Resolve(new FontRequest("Aptos"));
            if (resolved.IsFallback)
            {
                return false;
            }

            OpenTypeFont? font = FontProgramLoader.Load(resolved, CancellationToken.None);
            return font is not null && font.HasTrueTypeOutlines;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
        {
            return false;
        }
    }
    public static void DocxFootnoteStoriesTopHangsBelowBodyStart()
    {
        // RV06 p1-clamp probes (Word 16.0, b/ba/2para/bodyb sweeps): Office storiesTop
        // is independent of body size/after/length/before and follows body start with
        // footnote-side moves, while the renderer hung content one body em below the
        // lowest body baseline. Takes come from the same clamped capacity, but the
        // overflowing head block seats bottom-up instead of leaving capacity
        // slack above the margin. Trailing after-space can hang outside an
        // admitted top frame; occupied line content must remain inside it.
        var footnoteParas = new StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:t>Note body line " + line.ToString(CultureInfo.InvariantCulture) + " words here</w:t></w:r></w:p>");
        }
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutPage page = layout.Pages[0];
        DocxPlacedRelatedStoryLayout body = page.PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        // The portable filler measurer boxes Windows extents (1.2em); retain
        // its exact take while checking content separately from after-space.
        TestAssert.Equal(25, body.TextLines.Count);
        double bottom = body.TopY - body.Height;
        var last = body.StoryLayout.TextLines[^1];
        double after = Math.Max(0d, last.ParagraphAfterSpacing ?? last.PendingAfterSpacing ?? 0d);
        TestAssert.True(bottom <= page.MarginBottom + .000001d && bottom + after >= page.MarginBottom - .000001d &&
            body.TextLines[^1].BaselineY >= page.MarginBottom,
            "Only trailing after-space may hang below the margin; the full occupied last line remains inside its frame.");
    }

    public static void DocxFootnoteSeparatorGapFollowsSingleLineMetrics()
    {
        // RV06 separator-bottom probes (Word 16.0, Times/Aptos/Calibri): the Office
        // footnote gap above the body equals one single-spaced line box minus the
        // first-baseline inset, so it derives from the mark font instead of a fixed
        // 3pt constant. The patched synthetic face has a 1.15em single line, giving
        // a 0.21em gap.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = DocxFontPlanTextMeasurer.CreateAsync(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver).GetAwaiter().GetResult();
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is not null);
        DocxPlacedRelatedStoryLayout body = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is null);
        double markSize = separator.TextLines.Single().Segments[0].StyleRun.EffectiveProperties.FontSize;
        double separatorBottom = separator.TopY - separator.Height;
        TestAssert.True(Math.Abs(separatorBottom - body.TopY - 0.21d * markSize) < 0.000001d, "Footnote separator gap must follow the single-line box minus the first-baseline inset.");
    }

    public static void DocxFootnoteSeparatorGapKeepsLegacyConstantWithoutLineMetrics()
    {
        // Measurers without single-line metrics keep the legacy 3pt footnote gap.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        var resolver = new StrikeoutFontResolver();
        DocxFontPlan fontPlan = DocxFontPlan.Create(document, resolver, CancellationToken.None);
        var measurer = DocxFontPlanTextMeasurer.CreateAsync(fontPlan, resolver.Resolve(new FontRequest("StrikeFace")), CancellationToken.None, resolver).GetAwaiter().GetResult();
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is not null);
        DocxPlacedRelatedStoryLayout body = layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is null);
        double separatorBottom = separator.TopY - separator.Height;
        TestAssert.True(Math.Abs(separatorBottom - body.TopY - 3d) < 0.000001d, "Legacy footnote separator gap must stay 3pt.");
    }

    public static void DocxLongFootnoteSplitsAcrossPagesWithContinuationSeparator()
    {
        // RV06 footlong probe (Word 16.0): a 60-paragraph footnote splits 24/25/11
        // across 3 pages with full-width continuation rules plus end marks, while the
        // renderer stacks every line on the reference page off-page with no
        // continuation story placed. Continuation separators parse but never place.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body line 0 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 1 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 2 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 3 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 4 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 5 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 6 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 7 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 8 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 9 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 10 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 11 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 12 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 13 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 14 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 15 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 16 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 17 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 18 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 19 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 20 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 21 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 22 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 23 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 24 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 25 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 26 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 27 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 28 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 29 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 30 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 31 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 32 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 33 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 34 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 35 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 36 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 37 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 38 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 39 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 40 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 41 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 42 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 43 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 44 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 45 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 46 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 47 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 48 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 49 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 50 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 51 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 52 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 53 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 54 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 55 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 56 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 57 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 58 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 59 words here</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        foreach (DocxLayoutPage page in layout.Pages)
        {
            foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
            {
                if (story.StoryLayout.Story.Kind != DocxRelatedStoryKind.Footnote || story.StoryLayout.Story.Type != null)
                {
                    continue;
                }

                foreach (DocxTextLineLayout line in story.TextLines)
                {
                    TestAssert.True(line.BaselineY >= page.MarginBottom - 1d && line.BaselineY <= page.Height - page.MarginTop + 1d, "Footnote lines must stay inside their page.");
                }
            }
        }
        IReadOnlyList<DocxPlacedRelatedStoryLayout> continuations = layout.Pages.Skip(1).SelectMany(page => page.PlacedRelatedStories).Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator && story.SeparatorY is not null).ToArray();
        TestAssert.True(continuations.Count != 0, "Overflowing footnotes must place continuation separators on continuation pages.");
    }

    public static void DocxLongFootnoteWithoutSeparatorSplitsAcrossPages()
    {
        // RV06 footlong-absent probe (Word 16.0): separator-less footnotes (Word draws
        // default rules) still stack whole stories off-page on overflow. This pins
        // in-bounds lines plus continuation separators for the separator-less path.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body line 0 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 1 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 2 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 3 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 4 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 5 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 6 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 7 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 8 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 9 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 10 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 11 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 12 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 13 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 14 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 15 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 16 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 17 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 18 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 19 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 20 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 21 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 22 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 23 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 24 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 25 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 26 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 27 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 28 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 29 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 30 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 31 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 32 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 33 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 34 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 35 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 36 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 37 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 38 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 39 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 40 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 41 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 42 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 43 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 44 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 45 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 46 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 47 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 48 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 49 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 50 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 51 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 52 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 53 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 54 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 55 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 56 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 57 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 58 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 59 words here</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        var bodyLines = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal)).SelectMany(story => story.TextLines).ToArray();
        TestAssert.True(bodyLines.Length == 60, "All separator-less footnote lines must be placed.");
        TestAssert.True(bodyLines.All(line => layout.Pages.Any(page => line.BaselineY >= page.MarginBottom - 1d && line.BaselineY <= page.Height - page.MarginTop + 1d)), "Every separator-less footnote line must sit inside some page.");
        int contRules = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Count(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator && story.SeparatorY is not null);
        TestAssert.True(contRules != 0, "Overflowing separator-less footnotes must place continuation separators.");
    }

    public static void DocxFootnoteRemainderAccountsSeparatorUntilPlaced()
    {
        DocxRelatedStory story = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "31",
            [],
            [],
            [], null);
        var storyLayout = new DocxRelatedStoryLayout(story, 0, [], [], [], [], 0d);
        var location = new DocxLayoutEngine.DocxInlineReferenceLocation(
            0,
            DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d),
            new DocxInlineReference(DocxRelatedStoryKind.Footnote, "31", null, "1", 0, 1, 4));
        var inFlight = new DocxLayoutEngine.InFlightRelatedStory(storyLayout, null, null, 0, location);
        double[] lineBoxes = [10d, 10d, 10d, 10d, 10d];

        double remainder = DocxLayoutEngine.FootnoteRemainderHeight(inFlight, lineBoxes, null);

        TestAssert.Equal(53d, remainder);
    }

    public static void DocxFootnoteRemainderShrinksWithPlacedTakes()
    {
        DocxRelatedStory story = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "32",
            [],
            [],
            [], null);
        DocxTextRun lineRun = new("Shared note", 10d, null, false, false, false, null, null);
        DocxTextLineLayout RemainderLine(int index)
        {
            return new DocxTextLineLayout("Shared note line " + index.ToString(CultureInfo.InvariantCulture), lineRun, 10d, 72d, 700d - (12d * index), 200d, [], null, null, null, null, 12d, null, null, false, null, null, null, null, null, null, null, null, null, null, null, null, false);
        }

        var storyLayout = new DocxRelatedStoryLayout(story, 0, [RemainderLine(0), RemainderLine(1), RemainderLine(2), RemainderLine(3), RemainderLine(4)], [], [], [], 60d);
        var location = new DocxLayoutEngine.DocxInlineReferenceLocation(
            0,
            DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d),
            new DocxInlineReference(DocxRelatedStoryKind.Footnote, "32", null, "1", 0, 1, 4));
        var inFlight = new DocxLayoutEngine.InFlightRelatedStory(storyLayout, null, null, 0, location)
        {
            PlacedLineCount = 2,
            SeparatorPlaced = true
        };
        double[] lineBoxes = [10d, 10d, 10d, 10d, 10d];

        double remainder = DocxLayoutEngine.FootnoteRemainderHeight(inFlight, lineBoxes, null);

        TestAssert.Equal(30d, remainder);
        TestAssert.Equal(3, inFlight.RemainingLineCount);
        TestAssert.Equal(30d, inFlight.RemainingContentHeight(lineBoxes));
    }

    public static void DocxNarrowedFootnoteRemainderRebasesAtStoryOrigin()
    {
        DocxTextRun run = new("Shared note", 10d, null, false, false, false, null, null);
        DocxTextLineLayout Line(string text, double baseline)
        {
            return new DocxTextLineLayout(text, run, 10d, 72d, baseline, 200d, [], null, null, null, null, 12d, null, null, false, null, null, null, null, null, null, null, null, null, null, null, null, false);
        }

        DocxRelatedStory story = new(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "33",
            [],
            [],
            [], null);
        var storyLayout = new DocxRelatedStoryLayout(
            story, 0, [Line("Shared note line zero", 700d), Line("Shared note line one", 688d), Line("Shared note line two", 676d)], [], [], [], 36d);

        DocxRelatedStoryLayout narrowed = DocxLayoutEngine.NarrowStoryTextLinesForOffset(storyLayout, 1);

        TestAssert.Equal(2, narrowed.TextLines.Count);
        TestAssert.Equal(700d, narrowed.TextLines[0].BaselineY);
        TestAssert.Equal(688d, narrowed.TextLines[1].BaselineY);
        TestAssert.Equal(24d, narrowed.ContentHeight);
    }

    public static void DocxSharedFootnotePagesKeepBodyAboveNotes()
    {
        DocxDocument document = CreateInterleavingDocument(noteParagraphCount: 40, fillerParagraphCount: 60, footnoteId: "41");
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.Equal(40, placedTexts.Length);
        TestAssert.Equal(40, placedTexts.Distinct().Count());

        foreach (DocxLayoutPageSnapshot page in snapshot.Pages)
        {
            if (page.Items.Count == 0 || page.PlacedFootnoteStoryCount == 0)
            {
                continue;
            }

            double bodyBottom = page.Items.Min(item => item.Y);
            double footnoteTop = page.PlacedRelatedItems.Max(item => item.Y + item.Height);
            TestAssert.True(bodyBottom >= footnoteTop, "Shared pages must keep body above placed notes.");
        }
    }

    public static void DocxExhaustedFootnoteDrainUsesDedicatedPages()
    {
        DocxDocument document = CreateInterleavingDocument(noteParagraphCount: 200, fillerParagraphCount: 0, footnoteId: "42");
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.True(layout.Pages.Count > 2, "The long note must span several pages.");
        TestAssert.Equal(200, placedTexts.Length);
        TestAssert.Equal(200, placedTexts.Distinct().Count());
        TestAssert.Equal(0, layout.Pages[^1].Items.Count);
        TestAssert.True(layout.Pages[^1].PlacedRelatedStories.Count != 0, "The exhausted drain must append dedicated note pages.");
    }

    public static void DocxMidBodyFootnoteAnchorKeepsPrecedingFillersOnPage()
    {
        const int fillersBefore = 30;
        DocxDocument document = CreateMidBodyInterleavingDocument(fillersBefore, noteParagraphCount: 60, footnoteId: "43", fillersAfter: 30);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        int PageWithBlock(int blockIndex)
        {
            for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
            {
                foreach (DocxLayoutItem item in layout.Pages[pageIndex].Items)
                {
                    if (item is DocxTextLineLayout line && line.SourceBlockIndex == blockIndex)
                    {
                        return pageIndex;
                    }
                }
            }

            return -1;
        }

        TestAssert.True(PageWithBlock(fillersBefore - 1) >= 0, "Preceding fillers must be placed.");
        TestAssert.Equal(PageWithBlock(fillersBefore - 1), PageWithBlock(fillersBefore));

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.Equal(60, placedTexts.Length);
        TestAssert.Equal(60, placedTexts.Distinct().Count());
    }

    public static void DocxTableAnchoredFootnoteSharesBodyPages()
    {
        DocxParagraph cellAnchor = DocxTests.CreateDocxLayoutParagraph("Cell note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    "51",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        var cell = new DocxTableCell(
            string.Empty,
            [cellAnchor, DocxTests.CreateDocxLayoutParagraph("Cell trailing line", 10d, 12d)],
            null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [200d], [new DocxTableRow([cell], 30d)]);
        DocxParagraphElement[] noteElements = Enumerable.Range(0, 120)
            .Select(index => new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Shared note line " + index.ToString(CultureInfo.InvariantCulture) + " with trailing words", 10d, 12d)))
            .ToArray();
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            "51",
            noteElements,
            [],
            [], null);
        var bodyElements = new List<DocxParagraphElement> { };
        var paragraphs = new List<DocxParagraph> { cellAnchor };
        for (int index = 0; index < 60; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + index.ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

        var document = new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            new DocxBodyElement[] { new DocxTableElement(table) }.Concat(bodyElements).ToArray(),
            paragraphs.ToArray(),
            [])
        {
            RelatedStories = [footnoteStory]
        };
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        bool sharedPageExists = false;
        foreach (DocxLayoutPage page in layout.Pages)
        {
            bool hasFiller = page.Items.OfType<DocxTextLineLayout>().Any(line => line.Text.StartsWith("Filler body line", StringComparison.Ordinal));
            bool hasTable = page.Items.OfType<DocxTableRowLayout>().Any() || page.Items.OfType<DocxTextLineLayout>().Any(line => line.Text.StartsWith("Cell ", StringComparison.Ordinal));
            bool hasNote = page.PlacedRelatedStories
                .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
                .SelectMany(story => story.TextLines)
                .Any();
            if (hasFiller && hasNote && !hasTable)
            {
                sharedPageExists = true;
            }
        }

        TestAssert.True(sharedPageExists, "Table-anchored notes must share body pages instead of dedicating continuation pages.");

        string[] placedTexts = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.Equal(120, placedTexts.Length);
        TestAssert.Equal(120, placedTexts.Distinct().Count());
    }

    public static void DocxKeepChainBlockIndexesFollowKeepNext()
    {
        DocxParagraph Keep(string text)
        {
            return DocxTests.CreateDocxLayoutParagraph(text, 10d, 12d, keepRules: new DocxParagraphKeepRules(true, null, null, null, null, null));
        }

        DocxParagraph plain = DocxTests.CreateDocxLayoutParagraph("Plain", 10d, 12d);
        var table = new DocxTable(null, [60d], [new DocxTableRow([new DocxTableCell("Cell", [plain], null, null, null, null, [], DocxTableCellMargins.Empty)], 30d)]);
        IReadOnlyList<DocxBodyElement> elements = [new DocxParagraphElement(Keep("Keep zero")), new DocxParagraphElement(Keep("Keep one")), new DocxParagraphElement(plain), new DocxTableElement(table), new DocxParagraphElement(Keep("Keep four"))];

        TestAssert.True(new List<int> { 0, 1, 2 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 0)), "Keep chain from 0 must span both keep paras plus the plain follower.");
        TestAssert.True(new List<int> { 1, 2 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 1)), "Keep chain from 1 must span its keep para plus the plain follower.");
        TestAssert.True(new List<int> { 2 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 2)), "Plain paragraphs carry no chain.");
        TestAssert.True(new List<int> { 3 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 3)), "Table elements carry no chain.");
        TestAssert.True(new List<int> { 4 }.SequenceEqual(DocxLayoutEngine.KeepChainBlockIndexes(elements, 4)), "Trailing keep with no follower carries no chain.");
    }

    public static void DocxSharedContinuedFootnotePagesCarryContinuationRules()
    {
        DocxDocument document = CreateInterleavingDocument(noteParagraphCount: 120, fillerParagraphCount: 60, footnoteId: "44", includeSeparators: true);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        bool headPageSeen = false;
        foreach (DocxLayoutPage page in layout.Pages)
        {
            bool hasNormalLines = page.PlacedRelatedStories
                .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
                .SelectMany(story => story.TextLines)
                .Any();
            if (!hasNormalLines)
            {
                continue;
            }

            if (!headPageSeen)
            {
                headPageSeen = true;
                continue;
            }
        TestAssert.True(
                page.PlacedRelatedStories.Any(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote &&
                    story.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator),
                "Continued footnote pages must open with a continuation rule.");

        }

    }

    private static DocxDocument CreateMidBodyInterleavingDocument(int fillersBefore, int noteParagraphCount, string footnoteId, int fillersAfter)
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    footnoteId,
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraphElement[] noteElements = Enumerable.Range(0, noteParagraphCount)
            .Select(index => new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Shared note line " + index.ToString(CultureInfo.InvariantCulture) + " with trailing words", 10d, 12d)))
            .ToArray();
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            footnoteId,
            noteElements,
            [],
            [], null);
        var bodyElements = new List<DocxParagraphElement>();
        var paragraphs = new List<DocxParagraph>();
        for (int index = 0; index < fillersBefore; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + index.ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

        bodyElements.Add(new DocxParagraphElement(anchor));
        paragraphs.Add(anchor);
        for (int index = 0; index < fillersAfter; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + (fillersBefore + index).ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

        return new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            bodyElements.ToArray(),
            paragraphs.ToArray(),
            [])
        {
            RelatedStories = [footnoteStory]
        };
    }

    private static DocxDocument CreateInterleavingDocument(int noteParagraphCount, int fillerParagraphCount, string footnoteId, bool includeSeparators = false)
    {
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Body note marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Footnote,
                    footnoteId,
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 4)
            ]
        };
        DocxParagraphElement[] noteElements = Enumerable.Range(0, noteParagraphCount)
            .Select(index => new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Shared note line " + index.ToString(CultureInfo.InvariantCulture) + " with trailing words", 10d, 12d)))
            .ToArray();
        var footnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Footnote,
            "/word/footnotes.xml",
            footnoteId,
            noteElements,
            [],
            [], null);
        var bodyElements = new List<DocxParagraphElement> { new DocxParagraphElement(anchor) };
        DocxRelatedStory? separatorStory = null;
        DocxRelatedStory? continuationStory = null;
        if (includeSeparators)
        {
            separatorStory = new DocxRelatedStory(DocxRelatedStoryKind.Footnote, "/word/footnotes.xml", "0", [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Note separator mark", 10d, 12d))], [], [], DocxRelatedStoryType.Separator);
            continuationStory = new DocxRelatedStory(DocxRelatedStoryKind.Footnote, "/word/footnotes.xml", "1", [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Note continuation mark", 10d, 12d))], [], [], DocxRelatedStoryType.ContinuationSeparator);
        }

        var paragraphs = new List<DocxParagraph> { anchor };
        for (int index = 0; index < fillerParagraphCount; index++)
        {
            DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler body line " + index.ToString(CultureInfo.InvariantCulture) + " keeps body flowing", 10d, 12d);
            bodyElements.Add(new DocxParagraphElement(filler));
            paragraphs.Add(filler);
        }

        return new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            bodyElements.ToArray(),
            paragraphs.ToArray(),
            [])
        {
            RelatedStories = separatorStory is null || continuationStory is null ? [footnoteStory] : [footnoteStory, separatorStory, continuationStory]
        };
    }

    public static void DocxFootnoteSeparatorIgnoresDirectParagraphSpacing()
    {
        // Shared story-layout path with endnote separators (Office evidence is
        // endnote-only; no diverging footnote evidence): direct separator spacing
        // must not move footnote content, and the separator rule must not ride the
        // direct spacing either (the footnote content hangs off the stories top
        // while the rule sits in the separator story, so the rule is the moving
        // part here).
        double topAfter0 = LayoutFootnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:after=\"0\"/></w:pPr>");
        double topAfter24 = LayoutFootnoteBodyTopWithSeparatorSpacing("<w:pPr><w:spacing w:after=\"480\"/></w:pPr>");
        double topDefault = LayoutFootnoteBodyTopWithSeparatorSpacing(string.Empty);
        TestAssert.True(Math.Abs(topAfter0 - topAfter24) < 0.000001d, "Footnote separator after-spacing must not move the note block; observed shift=" + Math.Abs(topAfter0 - topAfter24).ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(topAfter0 - topDefault) < 0.000001d, "Footnote separator after-spacing must match latent defaults; observed shift=" + Math.Abs(topAfter0 - topDefault).ToString(CultureInfo.InvariantCulture) + ".");
        double ruleAfter0 = LayoutFootnoteSeparatorRuleWithSpacing("<w:pPr><w:spacing w:after=\"0\"/></w:pPr>");
        double ruleAfter24 = LayoutFootnoteSeparatorRuleWithSpacing("<w:pPr><w:spacing w:after=\"480\"/></w:pPr>");
        double ruleDefault = LayoutFootnoteSeparatorRuleWithSpacing(string.Empty);
        TestAssert.True(Math.Abs(ruleAfter0 - ruleAfter24) < 0.000001d, "Footnote separator rule must not ride direct after-spacing; observed shift=" + Math.Abs(ruleAfter0 - ruleAfter24).ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(ruleAfter0 - ruleDefault) < 0.000001d, "Footnote separator rule must match latent defaults; observed shift=" + Math.Abs(ruleAfter0 - ruleDefault).ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class HheaAscenderTextMeasurer(IDocxTextMeasurer inner, double hheaAscenderEm) : IDocxTextMeasurer, IDocxLineMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => inner is IDocxLineMetricsProvider lineMetrics ? lineMetrics.MeasureSingleLineHeight(run, fontSize) : fontSize;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => hheaAscenderEm * fontSize;
    }
    private static double LayoutEndnoteSeparatorHeightWithHheaAscender(int docDefaultsHalfPoints, double hheaAscenderEm)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new HheaAscenderTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), hheaAscenderEm), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).Height;
    }
    private static double LayoutEndnoteSeparatorHeightWithSingleLineEm(int docDefaultsHalfPoints, double singleLineEm)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new SingleLineEmTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), singleLineEm), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).Height;
    }
    private sealed class SingleLineEmTextMeasurer(IDocxTextMeasurer inner, double singleLineEm) : IDocxTextMeasurer
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = singleLineEm;
            return true;
        }
    }

    private static double LayoutEndnoteSeparatorContentGapWithMarkSize(int docDefaultsHalfPoints, double singleLineEm)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new SingleLineEmTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), singleLineEm), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }
    private static double LayoutEndnoteSeparatorHeightWithMarkSize(int docDefaultsHalfPoints)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).Height;
    }

    private sealed class WindowsAscenderTextMeasurer(IDocxTextMeasurer inner, double windowsAscenderEm) : IDocxTextMeasurer, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => windowsAscenderEm * fontSize;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => inner is IDocxStaticTextMetricsProvider staticMetrics ? staticMetrics.MeasureWindowsDescender(run, fontSize) : fontSize * 0.2d;
    }

    private static double LayoutEndnoteSeparatorHeightWithWindowsAscender(int docDefaultsHalfPoints, double windowsAscenderEm)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="SZ"/><w:szCs w:val="SZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("SZ", docDefaultsHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new WindowsAscenderTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), windowsAscenderEm), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).Height;
    }

    public static void DocxFootnoteLineHeightAddsDescDeficitExtra()
    {
        double tallFirst = LayoutFootnotePitchWithDescFaces("Tahoma", "Calibri");
        double tallSecond = LayoutFootnotePitchWithDescFaces("Calibri", "Tahoma");
        TestAssert.True(Math.Abs(tallFirst - 20.40d) < 0.01d, "Footnote line height must add the descender deficit extra; tallFirst=" + tallFirst.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(tallSecond - tallFirst) < 0.000001d, "Footnote descender extra must ignore run order; tallSecond=" + tallSecond.ToString(CultureInfo.InvariantCulture) + " tallFirst=" + tallFirst.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteLineHeightIgnoresWhitespaceOnlyRuns()
    {
        double pitch = LayoutFootnotePitchWithSpaceOnlyRun();
        TestAssert.True(Math.Abs(pitch - 19.58d) < 0.01d, "Footnote line height must ignore whitespace only runs; pitch=" + pitch.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static double LayoutFootnotePitchWithSpaceOnlyRun()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescDeficitTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY - lines[1].BaselineY;
    }

    public static void DocxLineHeightPrefersWindowsBoxOverHhea()
    {
        // RV06 line-box probes (Word 16.0): Office auto line boxes take max(Windows box, hhea box).
        (double footnotePitch, double bodyPitch) = LayoutWinBoxPitches();
        TestAssert.True(Math.Abs(footnotePitch - 23.98d) < 0.02d, "Footnote line height must prefer the Windows box.");
        TestAssert.True(Math.Abs(bodyPitch - 23.98d) < 0.02d, "Body line height must prefer the Windows box.");
    }

    private static (double FootnotePitch, double BodyPitch) LayoutWinBoxPitches()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="WinBox" w:hAnsi="WinBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Body winbox line one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="WinBox" w:hAnsi="WinBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Body winbox line two</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="WinBox" w:hAnsi="WinBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note winbox line one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="WinBox" w:hAnsi="WinBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note winbox line two</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new WinBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] footnoteLines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note winbox", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, footnoteLines.Length);
        DocxTextLineLayout[] bodyLines = layout.Pages[0].Items.OfType<DocxTextLineLayout>()
            .Where(line => line.Text.StartsWith("Body winbox", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, bodyLines.Length);
        return (footnoteLines[0].BaselineY - footnoteLines[1].BaselineY, bodyLines[0].BaselineY - bodyLines[1].BaselineY);
    }

    private sealed class WinBoxTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 10d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 12d;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize * 0.9d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => fontSize * 0.25d;
    }

    public static void DocxLineHeightPrefersTypographicBoxOverHhea()
    {
        // RV06 line-box probes (Word 16.0): Abadi runs box the unfloored typo box alone (typoFull 1.0698 against hhea 1.3047 with single floored at 1.15).
        (double footnotePitch, double bodyPitch) = LayoutTypoBoxPitches();
        TestAssert.True(Math.Abs(footnotePitch - 22.87d) < 0.02d, "Footnote line height must prefer the typographic box.");
        TestAssert.True(Math.Abs(bodyPitch - 22.87d) < 0.02d, "Body line height must prefer the typographic box.");
    }

    private static (double FootnotePitch, double BodyPitch) LayoutTypoBoxPitches()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="TypoBox" w:hAnsi="TypoBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Body typobox line one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="TypoBox" w:hAnsi="TypoBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Body typobox line two</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="TypoBox" w:hAnsi="TypoBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note typobox line one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="TypoBox" w:hAnsi="TypoBox"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note typobox line two</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new TypoBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] footnoteLines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note typobox", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, footnoteLines.Length);
        DocxTextLineLayout[] bodyLines = layout.Pages[0].Items.OfType<DocxTextLineLayout>()
            .Where(line => line.Text.StartsWith("Body typobox", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, bodyLines.Length);
        return (footnoteLines[0].BaselineY - footnoteLines[1].BaselineY, bodyLines[0].BaselineY - bodyLines[1].BaselineY);
    }

    private sealed class TypoBoxTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxTypographicMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 13.8d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 15.66d;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => 10.96d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => 2.89d;

        public bool UseTypographicMetrics(DocxTextRun? run) => true;

        public double MeasureTypographicLineHeight(DocxTextRun? run, double fontSize) => 12.84d;
    }

    private sealed class DescDeficitTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 10d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 10d;

        // Windows extents stay inside the 10pt hhea box so the deficit mechanism stays isolated under the line-box maximum (Tahoma ascender still leads for order-invariance).
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? fontSize * 0.6d : fontSize * 0.5d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? fontSize * 0.2d : fontSize * 0.27d;
    }

    private static double LayoutFootnotePitchWithDescFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescDeficitTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY - lines[1].BaselineY;
    }

    public static void DocxFootnoteFirstInsetFollowsDeeperDescender()
    {
        // RV06 descender probes (Word 16.0, Magneto/Calibri/Informal resolved sheets):
        // Office bottom bearing follows the deeper descender, so the inset is the hhea
        // box minus the maximum descender while equal descenders resolve identically
        // to the ascender-plus-gap form on both sides.
        double same = LayoutFootnoteFirstBaselineWithDescFaces("DescSame", "DescSame");
        double most = LayoutFootnoteFirstBaselineWithDescFaces("DescMost", "DescMost");
        TestAssert.True(Math.Abs((most - same) - 1.20d) < 0.02d, "Deeper-descender first baseline must sit above the equal-descender one by the inset difference; most=" + most.ToString(CultureInfo.InvariantCulture) + " same=" + same.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class DescMostTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxHheaLineGapProvider, IDocxHheaDescenderProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 12d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 14d;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "DescMost", StringComparison.Ordinal) ? 12d : 11.6d;

        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => 0d;

        public double MeasureHheaDescender(DocxTextRun? run, double fontSize) => 2.4d;

        // Windows extents stay inside the hhea box so the inset mechanism stays isolated under the line-box maximum.
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize * 0.8d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "DescMost", StringComparison.Ordinal) ? 3.6d : 2.4d;
    }

    private static double LayoutFootnoteFirstBaselineWithDescFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescMostTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY;
    }

    public static void DocxFootnoteFirstInsetFollowsWindowsDescender()
    {
        // RV06 windows-descender probes (Word 16.0, Informal Roman 10/12/14pt size
        // matrix plus Consolas): Office bottom bearing follows the Windows descender
        // even when the hhea descender runs deeper, so two synthetic faces sharing
        // the hhea box and Windows descender keep the same first baseline while
        // differing only in hhea descender depth.
        double deep = LayoutFootnoteFirstBaselineWithWinDescFaces("HheaDeep", "HheaDeep");
        double flat = LayoutFootnoteFirstBaselineWithWinDescFaces("HheaFlat", "HheaFlat");
        TestAssert.True(Math.Abs(deep - flat) < 0.02d, "Windows-descender first baselines must ignore hhea descender depth; deep=" + deep.ToString(CultureInfo.InvariantCulture) + " flat=" + flat.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class WinDescTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxHheaLineGapProvider, IDocxHheaDescenderProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 12d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 14d;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "HheaDeep", StringComparison.Ordinal) ? 12d : 11.6d;

        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => 0d;

        public double MeasureHheaDescender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "HheaDeep", StringComparison.Ordinal) ? 3.0d : 2.4d;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize * 0.9d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => 2.4d;
    }

    private static double LayoutFootnoteFirstBaselineWithWinDescFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new WinDescTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY;
    }

    public static void DocxTextBoxFirstInsetFollowsWindowsDescender()
    {
        // RV06 floatinset probes (Word 16.0, Tahoma slash Calibri slash Informal floating
        // textboxes): Office textbox first insets follow the hhea-box-minus-Windows-descender
        // tier, so a uniform synthetic face keeps its tier inset in textbox stories
        // instead of the legacy floor.
        double baseline = LayoutTextBoxFirstBaselineWithDescFaces("BoxTbx");
        TestAssert.True(Math.Abs(baseline - -11.6d) < 0.02d, "Textbox first baseline must sit at the tier inset; baseline=" + baseline.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxTextBoxFirstInsetFollowsMaxTier()
    {
        // RV06 floatmixsz probes (Word 16.0, Tahoma 10/14pt mixed with Calibri 12pt, both orders): Office
        // mixed-textbox insets take the max tier across runs (Tah10 reads Cal-tier, Tah14 reads Tah-tier),
        // unlike footnote mixed notes which keep the legacy floor so take capacities hold. The crossed
        // synthetic (high ascender on the low tier and vice versa) separates tier selection from
        // ascender selection. Pre-fix both orders read the ascender winner.
        double hiFirst = LayoutMixedTextBoxFirstBaseline("MixHi", "MixLo");
        double loFirst = LayoutMixedTextBoxFirstBaseline("MixLo", "MixHi");
        TestAssert.True(Math.Abs(hiFirst - -11.6d) < 0.02d, "Max-first mixed baselines must sit at the max tier inset; hiFirst=" + hiFirst.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(loFirst - -11.6d) < 0.02d, "Min-first mixed baselines must sit at the max tier inset; loFirst=" + loFirst.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class MixMaxTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxHheaLineGapProvider, IDocxHheaDescenderProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 12d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "MixHi", StringComparison.Ordinal) ? 11d : 14d;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "MixHi", StringComparison.Ordinal) ? 12d : 9d;

        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => 0d;

        public double MeasureHheaDescender(DocxTextRun? run, double fontSize) => 2.4d;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize * 0.9d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => 2.4d;
    }

    private static double LayoutMixedTextBoxFirstBaseline(string firstFamily, string secondFamily)
    {
        var runA = new DocxTextRun("Box note one ", 12d, null, false, false, false, null, firstFamily);
        var runB = new DocxTextRun("mixed tail", 12d, null, false, false, false, null, secondFamily);
        var paragraph = new DocxParagraph(
            [runA, runB],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var anchor = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d);
        var drawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(paragraph)]);
        var document = new DocxDocument(
            612d, 792d, 72d, 72d, 72d, 72d,
            DocxPageSettings.Empty,
            [drawing],
            [], [], [new DocxParagraphElement(anchor)], [], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new MixMaxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxFloatingDrawingLayout placed = layout.FloatingDrawings.Single();
        var textBoxLayout = TestAssert.NotNull(placed.TextBoxLayout, "Floating drawing must carry its textbox story layout.");
        DocxTextLineLayout first = textBoxLayout.TextLines.First(line => line.Text.StartsWith("Box note", StringComparison.Ordinal));
        return first.BaselineY;
    }

    private static double LayoutTextBoxFirstBaselineWithDescFaces(string family)
    {
        var run = new DocxTextRun("Box note one", 12d, null, false, false, false, null, family);
        var paragraph = new DocxParagraph(
            [run],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var anchor = DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d);
        var drawing = DocxTests.CreateFloatingTextBoxDrawing([new DocxParagraphElement(paragraph)]);
        var document = new DocxDocument(
            612d, 792d, 72d, 72d, 72d, 72d,
            DocxPageSettings.Empty,
            [drawing],
            [], [], [new DocxParagraphElement(anchor)], [], []);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescMostTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxFloatingDrawingLayout placed = layout.FloatingDrawings.Single();
        var textBoxLayout = TestAssert.NotNull(placed.TextBoxLayout, "Floating drawing must carry its textbox story layout.");
        DocxTextLineLayout first = textBoxLayout.TextLines.First(line => line.Text.StartsWith("Box note", StringComparison.Ordinal));
        return first.BaselineY;
    }

    public static void DocxFootnoteFirstInsetKeepsLegacyAboveMarkSingleEm()
    {
        // RV06 gap-slice guard (Palatino/Algerian first baselines): uniform notes whose
        // content single-em exceeds the mark reference stay on the legacy floor so
        // gap-overflowing notes keep validated placement while same-geometry Gate pairs
        // differ only by ascender.
        double big = LayoutFootnoteFirstBaselineWithFloorFaces("GateBig", "GateBig");
        double small = LayoutFootnoteFirstBaselineWithFloorFaces("GateSmall", "GateSmall");
        TestAssert.True(Math.Abs((small - big) - 0.72d) < 0.02d, "Gate pair must keep the legacy floor above mark single-em; small=" + small.ToString(CultureInfo.InvariantCulture) + " big=" + big.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteFirstInsetFollowsHheaAscenderPlusGap()
    {
        // RV06 box-law instruments (eighteen-family shape probes, Word 16.0): Office
        // first insets follow min-over-runs hhea ascender plus hhea gap with no 0.94em
        // floor, so a synthetic small face keeps its summed inset while a big-ascender
        // face keeps its own sum on both sides.
        double big = LayoutFootnoteFirstBaselineWithFloorFaces("BigTop", "BigTop");
        double small = LayoutFootnoteFirstBaselineWithFloorFaces("LowFloor", "LowFloor");
        TestAssert.True(Math.Abs((small - big) - 2.04d) < 0.02d, "Small-first minus big-first baselines must span the summed-inset difference; small=" + small.ToString(CultureInfo.InvariantCulture) + " big=" + big.ToString(CultureInfo.InvariantCulture) + ".");

    }

    private sealed class FloorInsetTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxHheaLineGapProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "GateBig", StringComparison.Ordinal) || string.Equals(run?.FontFamily, "GateSmall", StringComparison.Ordinal) ? 16d : 10d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "GateBig", StringComparison.Ordinal) || string.Equals(run?.FontFamily, "GateSmall", StringComparison.Ordinal) ? 16d : 10d;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "BigTop", StringComparison.Ordinal) || string.Equals(run?.FontFamily, "GateBig", StringComparison.Ordinal) ? fontSize : fontSize * 0.8d;

        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "LowFloor", StringComparison.Ordinal) ? fontSize * 0.03d : 0d;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => fontSize * 0.9d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => fontSize * 0.2d;
    }

    private static double LayoutFootnoteFirstBaselineWithFloorFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FloorInsetTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY;
    }

    public static void DocxFootnoteLineHeightFollowsHheaWithoutSingleFloor()
    {
        // RV06 box-law instruments (Corsiva/YiBaiti/Haettenschweiler/Arial/Ebrima
        // same-batch probes, Word 16.0): Office uniform pitches follow the per-family
        // hhea box with no single-line floor (Corsiva 13.50 against 13.46, YiBaiti
        // 12.36 against 12.32, Haett 12.88 against 12.80, Arial 13.81 against 13.80,
        // Ebrima 16.30 against 16.29), so a synthetic small-hhea face keeps its hhea
        // pitch while a big-hhea face keeps the hhea maximum on both sides.
        double big = LayoutFootnotePitchWithFloorFaces("BigBox", "BigBox");
        double small = LayoutFootnotePitchWithFloorFaces("SmallBox", "SmallBox");
        TestAssert.True(Math.Abs(big - 24.22d) < 0.01d, "Big-hhea pitch must keep the hhea maximum; big=" + big.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(small - 19.58d) < 0.01d, "Small-hhea pitch must drop the single-line floor; small=" + small.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class FloorBoxTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "SmallBox", StringComparison.Ordinal) || string.Equals(run?.FontFamily, "BigBox", StringComparison.Ordinal) ? 14d : 10d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "BigBox", StringComparison.Ordinal) ? 14d : 10d;

        // Windows extents stay inside the hhea box so the floor mechanism stays isolated under the line-box maximum (Tahoma ascender still leads).
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? fontSize * 0.6d : fontSize * 0.5d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? fontSize * 0.2d : fontSize * 0.27d;
    }

    private static double LayoutFootnotePitchWithFloorFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FloorBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY - lines[1].BaselineY;
    }

    public static void DocxFootnoteLineHeightSelfGatesWithoutSmallRun()
    {
        // RV06 box-law instruments: two big-ascender runs with matching descenders
        // carry no descender deficit, so the line keeps the uniform pitch while the
        // same helper with a small-ascender second run grows by the deficit extra.
        double control = LayoutFootnotePitchWithSelfGateFaces("Tahoma", "Calibri");
        double bigPair = LayoutFootnotePitchWithSelfGateFaces("Tahoma", "BigTop");
        TestAssert.True(Math.Abs(control - 20.40d) < 0.01d, "Mixed control pitch must keep the descender deficit extra; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(bigPair - 19.58d) < 0.01d, "Big-ascender pair pitch must self-gate to the uniform value; bigPair=" + bigPair.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteLineHeightSelfGatesWithDeepOwnDescender()
    {
        // RV06 box-law instruments (Palatino-like): the max-ascender run carrying the
        // deepest descender leaves a non-positive deficit, so the line keeps the
        // uniform pitch while the same helper with a shallow second run grows.
        double control = LayoutFootnotePitchWithSelfGateFaces("Tahoma", "Calibri");
        double deepOwn = LayoutFootnotePitchWithSelfGateFaces("DeepDsc", "Calibri");
        TestAssert.True(Math.Abs(control - 20.40d) < 0.01d, "Mixed control pitch must keep the descender deficit extra; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(deepOwn - 19.58d) < 0.01d, "Deep-descender max-ascender pitch must self-gate to the uniform value; deepOwn=" + deepOwn.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class SelfGateTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 10d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 10d;

        // Windows extents stay inside the hhea box so the self-gate mechanism stays isolated under the line-box maximum (ascender order preserved).
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => SelfGateAscenderEm(run) * fontSize * 0.45d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => SelfGateDescenderEm(run) * fontSize;

        private static double SelfGateAscenderEm(DocxTextRun? run)
        {
            if (string.Equals(run?.FontFamily, "DeepDsc", StringComparison.Ordinal))
            {
                return 1.05d;
            }

            if (string.Equals(run?.FontFamily, "Calibri", StringComparison.Ordinal))
            {
                return 0.9d;
            }

            return 1.0d;
        }

        private static double SelfGateDescenderEm(DocxTextRun? run)
        {
            if (string.Equals(run?.FontFamily, "DeepDsc", StringComparison.Ordinal))
            {
                return 0.35d;
            }

            if (string.Equals(run?.FontFamily, "Calibri", StringComparison.Ordinal))
            {
                return 0.27d;
            }

            return 0.2d;
        }
    }

    private static double LayoutFootnotePitchWithSelfGateFaces(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new SelfGateTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY - lines[1].BaselineY;
    }

    public static void DocxFootnoteContentGapFollowsFirstParagraph()
    {
        // RV06 mixed-gap probes (Word 16.0, Pal/Cal orders with r2f diff minus 1.20):
        // Office content gap follows the first paragraph while story-max would give
        // plus 0.4, so a Calibri-first mixed note keeps the Calibri-uniform gap while
        // Tahoma-first keeps the Tahoma-uniform gap.
        double tahUniform = LayoutFootnoteContentGapWithFamilies("Calibri", "Tahoma");
        double calUniform = LayoutFootnoteContentGapWithFamilies("Calibri", "Calibri");
        double tahFirst = LayoutMixedFootnoteContentGap("Tahoma", "Calibri");
        double calFirst = LayoutMixedFootnoteContentGap("Calibri", "Tahoma");
        TestAssert.True(Math.Abs(tahFirst - tahUniform) < 0.02d, "Tahoma-first mixed gap must keep the Tahoma-uniform gap; tahFirst=" + tahFirst.ToString(CultureInfo.InvariantCulture) + " tahUniform=" + tahUniform.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(calFirst - calUniform) < 0.02d, "Calibri-first mixed gap must keep the Calibri-uniform gap; calFirst=" + calFirst.ToString(CultureInfo.InvariantCulture) + " calUniform=" + calUniform.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static double LayoutMixedFootnoteContentGap(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyGapTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }

    public static void DocxFootnoteContentGapFollowsContentFont()
    {
        double control = LayoutFootnoteContentGapWithFamilies("Calibri", "Calibri");
        double mixed = LayoutFootnoteContentGapWithFamilies("Calibri", "Tahoma");
        TestAssert.True(Math.Abs(control - 3.37d) < 0.01d, "Uniform footnote gap must keep the shared value; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(mixed - 6.72d) < 0.01d, "Footnote gap must follow content font; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class FamilyGapTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? 1.5d : 1.2207d;
            return true;
        }
    }

    private static double LayoutFootnoteContentGapWithFamilies(string markFamily, string contentFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("MFAM", markFamily).Replace("CFAM", contentFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyGapTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }

    private static double LayoutEndnoteContentGapWithFamilies(string markFamily, string contentFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>""".Replace("MFAM", markFamily).Replace("CFAM", contentFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyGapTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }

    private static double LayoutEndnoteContentGapWithWindowsAscender(string markFamily, string contentFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>""".Replace("MFAM", markFamily).Replace("CFAM", contentFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyWascTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }

    public static void DocxFootnoteTakeExcludesTrailingAfterSpacing()
    {
        // RV06 after-spacing sweep (Word 16.0, seven after values): take needs exclude
        // the take last-line trailing after-spacing, so the head page takes one more
        // mixed line than the full-box fit allows. Synthetic desc-deficit metrics stand
        // in for the Calibri plus Tahoma mix with explicit 16pt note after-spacing, where
        // the legacy full-box fit takes 21 against 22 with the trailing-after exclusion.
        int headTake = LayoutFootnoteTakeWithAfterSpacingAndAfter(0, 320);
        TestAssert.Equal(22, headTake);
    }

    public static void DocxFootnoteContinuationTakeChargesRideTolerance()
    {
        // RV06 take-battery probes (Word 16.0, cal9-a8 laps 0.72 with cal-pal 0.36 while
        // a4 refuses the next line at lap 4.54): the take-side charge is the emitted rule
        // ride net of a 2.1pt overflow tolerance with full-box needs, rather than the full
        // rule block with trailing-after exclusion. At 24pt after-spacing the full-box
        // needs dominate, so the first continuation page takes 17 lines against 18 under
        // the legacy exclusion-plus-full-block model; the Office-anchored take-more
        // direction is pinned by the cal9-a8 29/31 and cal-pal 23/24/13 probe takes.
        // Same synthetic note shape as the overhead test at 24pt after-spacing.
        int continuationTake = LayoutFootnoteTakeWithAfterSpacingAndAfter(1, 480);
        TestAssert.Equal(17, continuationTake);
    }

    public static void DocxFootnoteContinuationTakeChargesRuleOverhead()
    {
        // RV06 continuation probes: the footnote bottom-anchor query charges the rule
        // overhead net of already-reserved top space, so the first continuation page
        // takes fewer lines than the full-frame fit allows. Same synthetic note shape
        // as the trailing-after test at 10pt after-spacing, counting the first
        // continuation page, where a zeroed overhead takes 29 against 28 charged.
        int continuationTake = LayoutFootnoteTakeWithAfterSpacingAndAfter(1, 200);
        TestAssert.Equal(28, continuationTake);
    }

    private static int LayoutFootnoteTakeWithAfterSpacingAndAfter(int pageIndex, int afterTwips)
    {
        var footnoteParas = new System.Text.StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:pPr><w:spacing w:after=\"" + afterTwips.ToString(CultureInfo.InvariantCulture) + "\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\">Note body line " + line.ToString(CultureInfo.InvariantCulture) + " start tail</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"Tahoma\" w:hAnsi=\"Tahoma\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\"> mixed tail</w:t></w:r></w:p>");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescDeficitTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        return layout.Pages[pageIndex].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Count();
    }

    public static void DocxFootnoteContentGapAddsWindowsAscenderExcess()
    {
        // RV06 fnwrap content-gap probes (Word 16.0): Office hangs big-ascender content
        // lower by the Windows-ascender excess over the mark, so a synthetic big-ascender
        // content face at 1.05em against a 0.9521em mark gains 0.94pt of gap at 12pt
        // while uniform content keeps the shared value.
        double control = LayoutFootnoteContentGapWithWindowsAscender("Calibri", "Calibri");
        double mixed = LayoutFootnoteContentGapWithWindowsAscender("Calibri", "BigAsc");
        TestAssert.True(Math.Abs(control - 3.37d) < 0.01d, "Uniform footnote gap must keep the shared value; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(mixed - 4.54d) < 0.01d, "Big-ascender content gap must add the Windows excess; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class FamilyWascTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = 1.2207d;
            return true;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => (string.Equals(run?.FontFamily, "BigAsc", StringComparison.Ordinal) ? 1.05d : 0.9521d) * fontSize;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => inner is IDocxStaticTextMetricsProvider staticMetrics ? staticMetrics.MeasureWindowsDescender(run, fontSize) : fontSize * 0.2d;
    }

    private sealed class TypoWascTextMeasurer(IDocxTextMeasurer inner, bool useTypographicMetrics) : IDocxTextMeasurer, IDocxStaticTextMetricsProvider, IDocxTypographicMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = 1.2207d;
            return true;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => (string.Equals(run?.FontFamily, "BigAsc", StringComparison.Ordinal) ? 1.05d : 0.9521d) * fontSize;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => inner is IDocxStaticTextMetricsProvider staticMetrics ? staticMetrics.MeasureWindowsDescender(run, fontSize) : fontSize * 0.2d;

        public bool UseTypographicMetrics(DocxTextRun? run) => useTypographicMetrics;
    }

    private static double LayoutFootnoteContentGapWithTypographicAscender(string markFamily, string contentFamily)
    {
        return LayoutFootnoteContentGapWithWindowsAscender(markFamily, contentFamily, true);
    }

    private static double LayoutFootnoteContentGapWithWindowsAscender(string markFamily, string contentFamily)
    {
        return LayoutFootnoteContentGapWithWindowsAscender(markFamily, contentFamily, false);
    }

    public static void DocxFootnoteContentGapScalesShelterWithRunSize()
    {
        // RV06 endnote mark-size probes (m28/m28c): the inset shelter must scale with the run size, so big marks on small content keep the run-sized shelter instead of releasing the supplement early. Same-size pairs behave identically either way.
        double same = LayoutFootnoteContentGapWithSizedWindowsAscender("Calibri", 28, "BigAsc", 28);
        double mixed = LayoutFootnoteContentGapWithSizedWindowsAscender("Calibri", 28, "BigAsc", 20);
        TestAssert.True(Math.Abs(same - 5.30d) < 0.02d, "Same-size content gap must keep the mark-sized shelter; same=" + same.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(mixed - 4.91d) < 0.02d, "Mixed-size content gap must shelter with the run size; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }
    private static double LayoutFootnoteContentGapWithSizedWindowsAscender(string markFamily, int markHalfPoints, string contentFamily, int contentHalfPoints)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rIdS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
            ["word/styles.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="MSZ"/><w:szCs w:val="MSZ"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>""".Replace("MFAM", markFamily).Replace("MSZ", markHalfPoints.ToString(CultureInfo.InvariantCulture)),
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="MSZ"/><w:szCs w:val="MSZ"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="CSZ"/><w:szCs w:val="CSZ"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="CSZ"/><w:szCs w:val="CSZ"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("MFAM", markFamily).Replace("CFAM", contentFamily).Replace("MSZ", markHalfPoints.ToString(CultureInfo.InvariantCulture)).Replace("CSZ", contentHalfPoints.ToString(CultureInfo.InvariantCulture))
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new TypoWascTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), false), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }
    private static double LayoutFootnoteContentGapWithWindowsAscender(string markFamily, string contentFamily, bool useTypographicMetrics)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="CFAM" w:hAnsi="CFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("MFAM", markFamily).Replace("CFAM", contentFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new TypoWascTextMeasurer(new DocxTests.FamilyWidthTextMeasurer(), useTypographicMetrics), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }

    public static void DocxFootnoteContentGapSkipsTypographicSupplement()
    {
        // RV06 Aptos probes (Word 16.0, uniform Aptos lacks the excess while Segoe keeps
        // 1.524 of it): only Aptos sets OS/2 USE_TYPO_METRICS among probed families, so a
        // typographic content face with big-ascender metrics keeps the shared gap value
        // while the same face without the flag gains the Windows excess.
        double control = LayoutFootnoteContentGapWithWindowsAscender("Calibri", "Calibri");
        double flagged = LayoutFootnoteContentGapWithTypographicAscender("Calibri", "BigAsc");
        double unflagged = LayoutFootnoteContentGapWithWindowsAscender("Calibri", "BigAsc");
        TestAssert.True(Math.Abs(control - 3.37d) < 0.01d, "Uniform footnote gap must keep the shared value; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(flagged - 3.37d) < 0.01d, "Typographic content gap must skip the supplement; flagged=" + flagged.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(unflagged - 4.54d) < 0.01d, "Unflagged big-ascender content gap must keep the Windows excess; unflagged=" + unflagged.ToString(CultureInfo.InvariantCulture) + ".");
    }

    public static void DocxFootnoteContentGapYieldsToCoveredInset()
    {
        // RV06 Tah-Verdana probes: the Windows excess must yield where the first-line
        // inset already covers it, so Tahoma-first Verdana-second content keeps the
        // shared value while uniform big-ascender content keeps its excess.
        double control = LayoutFootnoteContentGapWithCoveredInset("Calibri", "Calibri");
        double mixed = LayoutFootnoteContentGapWithCoveredInset("Calibri", "Tahoma", "Verdana");
        TestAssert.True(Math.Abs(control - 3.37d) < 0.01d, "Uniform footnote gap must keep the shared value; control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(mixed - 3.37d) < 0.01d, "Covered-inset content gap must yield to the shared value; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class FamilyCoveredInsetTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => inner is IDocxLineMetricsProvider lineMetrics ? lineMetrics.MeasureSingleLineHeight(run, fontSize) : fontSize;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize)
        {
            double em = string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? 1.0005d : string.Equals(run?.FontFamily, "Verdana", StringComparison.Ordinal) ? 1.0054d : 0.75d;
            return em * fontSize;
        }

        public bool TryGetSingleLineEm(DocxTextRun? run, out double value)
        {
            value = 1.2207d;
            return true;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
        {
            double em = string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? 1.0005d : string.Equals(run?.FontFamily, "Verdana", StringComparison.Ordinal) ? 1.0054d : 0.9521d;
            return em * fontSize;
        }

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => inner is IDocxStaticTextMetricsProvider staticMetrics ? staticMetrics.MeasureWindowsDescender(run, fontSize) : fontSize * 0.2d;
    }

    private static double LayoutFootnoteContentGapWithCoveredInset(string markFamily, string firstFamily, string secondFamily = "")
    {
        string secondRun = string.IsNullOrEmpty(secondFamily) ? "" : "<w:r><w:rPr><w:rFonts w:ascii=\"F2\" w:hAnsi=\"F2\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\"> mixed tail</w:t></w:r>";
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r>" + secondRun + "</w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r>" + secondRun + "</w:p></w:footnote></w:footnotes>""".Replace("MFAM", markFamily).Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyCoveredInsetTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxPlacedRelatedStoryLayout separator = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        DocxPlacedRelatedStoryLayout content = layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal));
        return (separator.TopY - separator.Height) - content.TopY;
    }

    public static void DocxFootnoteOverflowHeadSliceHangsContentByGap()
    {
        // RV06 mixlong probes (Word 16.0): overflowing head slices hang content below
        // the take-driven rule like fitting notes, so a synthetic big-ascender overflow
        // places its first baseline lower by the Windows excess at 12pt while takes,
        // take heights and uniform content stay put.
        double control = LayoutFootnoteOverflowHeadFirstBaseline("Calibri", "Calibri");
        double mixed = LayoutFootnoteOverflowHeadFirstBaseline("Calibri", "BigAsc");
        TestAssert.True(Math.Abs((control - mixed) - 1.17d) < 0.02d, "Overflow head slice must hang content by the Windows excess; control=" + control.ToString(CultureInfo.InvariantCulture) + " mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static double LayoutFootnoteOverflowHeadFirstBaseline(string firstFamily, string secondFamily)
    {
        var footnoteParas = new System.Text.StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:r><w:rPr><w:rFonts w:ascii=\"F1\" w:hAnsi=\"F1\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\">Note body line " + line.ToString(CultureInfo.InvariantCulture) + " start tail</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"F2\" w:hAnsi=\"F2\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\"> mixed tail</w:t></w:r></w:p>".Replace("F1", firstFamily).Replace("F2", secondFamily));
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:rPr><w:rFonts w:ascii="MFAM" w:hAnsi="MFAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>""".Replace("MFAM", firstFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new FamilyWascTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout firstLine = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .First();
        return firstLine.BaselineY;
    }

    private static int LayoutEndnoteTakeWithAfterSpacingAndAfter(int pageIndex, int afterTwips)
    {
        var noteParas = new System.Text.StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            noteParas.Append("<w:p><w:pPr><w:spacing w:after=\"" + afterTwips.ToString(CultureInfo.InvariantCulture) + "\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\">Note body line " + line.ToString(CultureInfo.InvariantCulture) + " start tail</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii=\"Tahoma\" w:hAnsi=\"Tahoma\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t xml:space=\"preserve\"> mixed tail</w:t></w:r></w:p>");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2">""" + noteParas.ToString() + """</w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescDeficitTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        return layout.Pages[pageIndex].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Count();
    }

    public static void DocxFootnoteNinePtHeadTakeReservesClampMargin()
    {
        // RV06 size sweep (Word 16.0, cal9-a0 takes 47 against 48 while 10pt and
        // first-10pt boundaries hold exact): uniform 9pt-after-0 head takes 53 under the
        // 4.0pt uniform reserve against 54 under the 2.5pt reserve; the Office-anchored
        // take-fewer direction is pinned by the cal9-a0 47/13 probe takes with size-10
        // 43/17 and first-10 36/24 holding exact.
        // This capacity regression exercises the authored separator paragraph.
        // Its fixture registers both special stories in settings.
        int ninePtHeadTake = LayoutFootnoteUniformHeadTakeWithAfterAndSize(0, 0);
        TestAssert.Equal(53, ninePtHeadTake);
    }

    public static void DocxFootnoteHeadTakeReservesClampMargin()
    {
        // RV06 take-battery probes (Word 16.0): Office head takes fit one more
        // line than clamp-minus-separator allows, so the clamp carries a 2.5pt
        // reserve shaved from the take side, pinned where a 10pt after-spacing
        // head take drops from 29 to 28 against the untrimmed fit.
        int headTake = LayoutFootnoteUniformHeadTakeWithAfter(200);
        TestAssert.Equal(28, headTake);
    }

    private static int LayoutFootnoteUniformHeadTakeWithAfter(int afterTwips)
    {
        return LayoutFootnoteUniformHeadTakeWithAfterAndSize(afterTwips, 24);
    }

    private static int LayoutFootnoteUniformHeadTakeWithAfterAndSize(int afterTwips, int sizeHalfPoints)
    {
        var footnoteParas = new System.Text.StringBuilder();
        for (int line = 0; line < 60; line++)
        {
            footnoteParas.Append("<w:p><w:pPr><w:spacing w:after=\"" + afterTwips.ToString(CultureInfo.InvariantCulture) + "\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"" + sizeHalfPoints.ToString(CultureInfo.InvariantCulture) + "\"/><w:szCs w:val=\"" + sizeHalfPoints.ToString(CultureInfo.InvariantCulture) + "\"/></w:rPr><w:t xml:space=\"preserve\">Note body line " + line.ToString(CultureInfo.InvariantCulture) + " uniform tail</w:t></w:r></w:p>");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>""",
            ["word/settings.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnotePr><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr></w:settings>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2">""" + footnoteParas.ToString() + """</w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DescDeficitTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Count();
    }

    public static void DocxFootnoteFirstBaselineIgnoresLaterRunFonts()
    {
        // RV05 fnmix probes (Word COM references edge-fnmix/edge-fnmix3: note first
        // baselines sit at 85.46 regardless of run order or family): related-story
        // content keeps legacy widest-run insets, so a trailing Tahoma run must not
        // move the first baseline. Pre-scoping (max-hhea) it drops by 0.73.
        double control = LayoutFootnoteFirstBaselineWithSecondRun("Calibri");
        double mixed = LayoutFootnoteFirstBaselineWithSecondRun("Tahoma");
        TestAssert.True(Math.Abs(mixed - control) < 0.000001d, "Footnote first baseline must ignore later-run fonts; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + " control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class StoryHheaTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => inner is IDocxLineMetricsProvider lineMetrics ? lineMetrics.MeasureSingleLineHeight(run, fontSize) : fontSize;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize)
        {
            double em = run?.FontFamily == "Tahoma" ? 1.0005d : 0.75d;
            return em * fontSize;
        }
    }

    private static double LayoutFootnoteFirstBaselineWithSecondRun(string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note start words </w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="FAM" w:hAnsi="FAM"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">middle mixed words</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("FAM", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new StoryHheaTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .First(line => line.Text.StartsWith("Note start words", StringComparison.Ordinal)).BaselineY;
    }
    public static void DocxFootnoteLineHeightUsesMaxAcrossRuns()
    {
        double control = LayoutFootnotePitchWithRuns("Calibri", "Calibri");
        double tall = LayoutFootnotePitchWithRuns("Tahoma", "Tahoma");
        double mixed = LayoutFootnotePitchWithRuns("Calibri", "Tahoma");
        double swapped = LayoutFootnotePitchWithRuns("Tahoma", "Calibri");
        TestAssert.True(Math.Abs(mixed - tall) < 0.000001d, "Footnote line height must use max across runs; mixed=" + mixed.ToString(CultureInfo.InvariantCulture) + " tall=" + tall.ToString(CultureInfo.InvariantCulture) + " control=" + control.ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs(swapped - tall) < 0.000001d, "Footnote line height must ignore run order; swapped=" + swapped.ToString(CultureInfo.InvariantCulture) + " tall=" + tall.ToString(CultureInfo.InvariantCulture) + ".");
    }

    private sealed class StoryHeightTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "Tahoma", StringComparison.Ordinal) ? 20d : 10d;
    }

    private static double LayoutFootnotePitchWithRuns(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with footnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body one start tail one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail one</w:t></w:r></w:p><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Note body two start tail two</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> mixed tail two</w:t></w:r></w:p></w:footnote></w:footnotes>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new StoryHeightTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].PlacedRelatedStories
            .Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal))
            .SelectMany(story => story.TextLines)
            .Where(line => line.Text.StartsWith("Note body", StringComparison.Ordinal))
            .Take(2)
            .ToArray();
        TestAssert.Equal(2, lines.Length);
        return lines[0].BaselineY - lines[1].BaselineY;
    }

    private static int CountSeparatorVariantTextShows(string separatorParagraph, DocxTests.SingleResolutionFontResolver resolver)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0">SEPPARA</w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>""".Replace("SEPPARA", separatorParagraph)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        PdfPage page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single();
        return DocxTests.CountPdfTextShows(page.Content);
    }
    public static void DocxFootnoteSeparatorIgnoresExactLineSpacing()
    {
        // Shared story-layout path with endnote separators (Office evidence is
        // endnote-only; no diverging footnote evidence): the footnote rule rides
        // the separator bottom, so an exact line rule that Office ignores would
        // move our rule with the taller story.
        double ruleExact = LayoutFootnoteSeparatorRuleWithSpacing("<w:pPr><w:spacing w:line=\"480\" w:lineRule=\"exact\"/></w:pPr>");
        double ruleAuto = LayoutFootnoteSeparatorRuleWithSpacing(string.Empty);
        TestAssert.True(Math.Abs(ruleExact - ruleAuto) < 0.000001d, "Footnote separator exact line spacing must match auto lines; observed shift=" + Math.Abs(ruleExact - ruleAuto).ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static double LayoutEndnoteBodyTopWithSeparatorSpacing(string separatorParagraphProperties)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p>""" + separatorParagraphProperties + """<w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal)).TopY;
    }

    private static double LayoutFootnoteBodyTopWithSeparatorSpacing(string separatorParagraphProperties)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p>""" + separatorParagraphProperties + """<w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && (story.StoryLayout.Story.Type is null || story.StoryLayout.Story.Type == DocxRelatedStoryType.Normal)).TopY;
    }

    private static double LayoutFootnoteSeparatorRuleWithSpacing(string separatorParagraphProperties)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""",
            ["word/footnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="0"><w:p>""" + separatorParagraphProperties + """<w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>Note body one</w:t></w:r></w:p><w:p><w:r><w:t>Note body two</w:t></w:r></w:p></w:footnote></w:footnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        return layout.Pages[0].PlacedRelatedStories.Single(story => story.SeparatorY is not null).SeparatorY ?? 0d;
    }
    }
