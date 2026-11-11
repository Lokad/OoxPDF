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

    public static void DocxContinuedFootnoteFooterFramesKeepHeadsSeparatorsAndTailClear()
    {
        DocxDocument plain = ReadWrappedFootnoteFlowFixture(100);
        DocxParagraph footer = DocxTests.CreateDocxLayoutParagraph("public footer frame", 12d, 14d) with
        {
            LineSpacingPoints = null,
            Spacing = DocxParagraphSpacing.Empty with { BeforeValue = "0", AfterValue = "0", LineValue = "240", LineRuleValue = "auto" },
        };
        DocxDocument Compose(int count, bool alternating)
        {
            var paragraphs = new Dictionary<string, IReadOnlyList<DocxParagraph>>
            { ["default"] = Enumerable.Repeat(footer, count).ToArray() };
            if (alternating) paragraphs["even"] = [footer];
            var elements = paragraphs.ToDictionary(p => p.Key,
                p => (IReadOnlyList<DocxBodyElement>)p.Value.Select(v => (DocxBodyElement)new DocxParagraphElement(v)).ToArray());
            var settings = plain.PageSettings with { FooterDistancePoints = 36d, EvenAndOddHeaders = alternating,
                FooterParagraphsByType = paragraphs, FooterBodyElementsByType = elements };
            return plain with { PageSettings = settings, FinalSectionBreak = plain.FinalSectionBreak! with { PageSettings = settings },
                FooterParagraphs = paragraphs["default"], FooterParagraphsByType = paragraphs, FooterBodyElementsByType = elements };
        }
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new LogicalBoundaryTextMeasurer();
        DocxLayout baseline = engine.Create(plain, measurer, default);
        var expected = baseline.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => s.StoryLayout.Story.IsNormalStoryType)
            .GroupBy(s => s.StoryLayout.Story.Id).ToDictionary(g => g.Key!, g => g.SelectMany(s => s.StoryLayout.TextLines)
                .Select(l => (l.Text, l.SourceParagraph, l.SourceLineIndex)).ToArray());
        // A fitting footer alone does not imply that the complete note batch fits.
        // Short notes also establish the physical head-rule gap for whole flow.
        DocxDocument fittingPlain = WithEndnoteText(plain, "37", "Short public note.");
        fittingPlain = WithEndnoteText(fittingPlain, "4", "Another short public note.");
        DocxLayout fittingBaseline = engine.Create(fittingPlain, measurer, default);
        DocxLayout fitting = engine.Create(Compose(1, false) with { RelatedStories = fittingPlain.RelatedStories }, measurer, default);
        DocxPlacedRelatedStoryLayout fittingHead = fitting.Pages[0].PlacedRelatedStories.Single(s =>
            s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        double fittingHeadRuleGap = fittingHead.SeparatorY!.Value - fitting.Pages[0].PlacedRelatedStories
            .First(s => s.StoryLayout.Story.IsNormalStoryType).TextLines[0].BaselineY;
        foreach ((int footerCount, bool alternating) in new[] { (10, false), (10, true), (1, false) })
        {
            DocxDocument source = Compose(footerCount, alternating);
            DocxLayout layout = engine.Create(source, measurer, default);
            TestAssert.True(ReferenceEquals(source.BodyElements, plain.BodyElements) && ReferenceEquals(source.RelatedStories, plain.RelatedStories),
                "Footer frame placement preserves authored body and note ownership.");
            var actual = layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => s.StoryLayout.Story.IsNormalStoryType)
                .GroupBy(s => s.StoryLayout.Story.Id).ToArray();
            TestAssert.Equal(expected.Count, actual.Length);
            foreach (var group in actual)
            {
                TestAssert.True(expected[group.Key!].SequenceEqual(group.SelectMany(s => s.StoryLayout.TextLines)
                    .Select(l => (l.Text, l.SourceParagraph, l.SourceLineIndex))), "Note lines remain source-owned, ordered and exactly once.");
                TestAssert.True(group.First().StoryLayout.TextLines.Count >= 2,
                    "An occupied footer frame defers a fresh note head that cannot fit its default widow pair.");
            }
            TestAssert.True(actual.Any(g => g.Count() > 1), "The fixture exercises actual continuation and final drain.");
            foreach (DocxLayoutPage page in layout.Pages)
            {
                double highestFooterBaseline = page.StaticTextLines.Max(l => l.BaselineY);
                foreach (var story in page.PlacedRelatedStories.Where(s => s.StoryLayout.Story.IsNormalStoryType))
                foreach (var line in story.TextLines)
                    TestAssert.True(line.BaselineY > highestFooterBaseline,
                        $"Footnote glyphs clear selected footer text through the final drain: note={story.StoryLayout.Story.Id}, line={line.SourceLineIndex}, baseline={line.BaselineY}, footer={highestFooterBaseline}.");
                TestAssert.True(page.PlacedRelatedStories.All(s => s.SeparatorY is null || s.SeparatorY > highestFooterBaseline),
                    "Head and continuation rules also stay above selected footer text.");
                if (page.PlacedRelatedStories.Any(s => s.StoryLayout.Story.IsNormalStoryType))
                    TestAssert.True(page.PlacedRelatedStories.Count(s => s.SeparatorY is not null) == 1,
                        "A shared footnote compartment emits one physical separator.");
                DocxPlacedRelatedStoryLayout? head = page.PlacedRelatedStories.FirstOrDefault(s =>
                    s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
                if (head is not null)
                {
                    DocxPlacedRelatedStoryLayout note = page.PlacedRelatedStories.First(s => s.StoryLayout.Story.IsNormalStoryType);
                    TestAssert.True(Math.Abs(head.SeparatorY!.Value - note.TextLines[0].BaselineY - fittingHeadRuleGap) < .000001d,
                        "Whole and split fresh heads retain the same physical separator gap above their first glyphs.");
                }
            }
        }
        DocxLayout consumed = engine.Create(Compose(100, true), measurer, default);
        DocxLayout uniformlyConsumed = engine.Create(Compose(100, false), measurer, default);
        TestAssert.Equal(baseline.Pages.Count, consumed.Pages.Count);
        TestAssert.True(uniformlyConsumed.Pages.SelectMany(p => p.PlacedRelatedStories).Select(s => (s.TopY, s.Height, s.ContentTopOffset))
            .SequenceEqual(consumed.Pages.SelectMany(p => p.PlacedRelatedStories).Select(s => (s.TopY, s.Height, s.ContentTopOffset))),
            "An unusable selected frame shares bounded fallback placement with uniformly unusable footer frames.");
        var consumedStories = consumed.Pages.SelectMany(p => p.PlacedRelatedStories)
            .Where(s => s.StoryLayout.Story.IsNormalStoryType).ToArray();
        TestAssert.True(consumedStories.All(s => !s.UsesPlainFootnoteFrame),
            "Unusable static variants do not admit note typography.");
        foreach (var group in consumedStories.GroupBy(s => s.StoryLayout.Story.Id))
            TestAssert.True(expected[group.Key!].SequenceEqual(group.SelectMany(s => s.StoryLayout.TextLines)
                .Select(l => (l.Text, l.SourceParagraph, l.SourceLineIndex))),
                "Fallback placement retains every canonical source line and paragraph owner exactly once.");
        TestAssert.Equal(1, fittingBaseline.Pages.Count);
        TestAssert.True(fittingBaseline.Pages.SelectMany(p => p.PlacedRelatedStories).Select(s => (s.TopY, s.Height, s.ContentTopOffset))
            .SequenceEqual(fitting.Pages.SelectMany(p => p.PlacedRelatedStories).Select(s => (s.TopY, s.Height, s.ContentTopOffset))),
            "A complete fitting footnote batch retains its original placements with a fitting footer.");
    }

    public static void DocxSharedFootnoteHeadFitRetainsPhysicalSeparatorSpacing()
    {
        DocxDocument source = WithPlainFootnoteTestFooter(ReadWrappedFootnoteFlowFixture(100));
        source = source with
        {
            BodyElements = source.BodyElements.OfType<DocxParagraphElement>().Take(1).ToArray(),
            RelatedStories = source.RelatedStories.Where(s => !s.IsNormalStoryType || s.Id == "37").Select(s =>
                s.IsNormalStoryType ? s with
                {
                    BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    {
                        Runs = [p.Runs[0], p.Runs[1] with { FontSize = 14d }],
                    })).ToArray(),
                } : s).ToArray(),
        };
        var settings = source.PageSettings with { HeightValue = "15760" };
        source = source with { PageSettings = settings, PageHeightPoints = 788d,
            FinalSectionBreak = source.FinalSectionBreak! with { PageSettings = settings } };
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        DocxLayout layout = engine.Create(source, measurer, default);
        var slices = layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => s.StoryLayout.Story.IsNormalStoryType).ToArray();
        TestAssert.True(slices.Select(s => s.TextLines.Count).SequenceEqual(new[] { 36, 22 }),
            "The shared frame admits the full 36-line fresh head before the 22-line continuation.");
        DocxDocument shortSource = WithEndnoteText(source, "37", "Short public footnote.");
        DocxLayout shortLayout = engine.Create(shortSource, measurer, default);
        var shortHead = shortLayout.Pages[0].PlacedRelatedStories.Single(s => s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        var head = layout.Pages[0].PlacedRelatedStories.Single(s => s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
        double gap = shortHead.SeparatorY!.Value - shortLayout.Pages[0].PlacedRelatedStories.First(s => s.StoryLayout.Story.IsNormalStoryType).TextLines[0].BaselineY;
        TestAssert.True(Math.Abs(head.SeparatorY!.Value - slices[0].TextLines[0].BaselineY - gap) < .000001d,
            "A split fresh head keeps the complete fitting head's physical separator spacing.");
        var canonical = layout.RelatedStories.Single(s => s.Story.IsNormalStoryType);
        TestAssert.True(canonical.TextLines.Select(l => (l.Text, l.SourceParagraph, l.SourceLineIndex)).SequenceEqual(
            slices.SelectMany(s => s.StoryLayout.TextLines).Select(l => (l.Text, l.SourceParagraph, l.SourceLineIndex))),
            "Every source line is placed exactly once through the continuation drain.");
        TestAssert.True(layout.Pages.All(p => p.PlacedRelatedStories.Where(s => s.StoryLayout.Story.IsNormalStoryType)
            .All(s => s.TopY - s.Height >= p.MarginBottom - .001d)), "All note clips clear the fitting footer frame.");
    }

    public static void DocxDeferredFootnoteHeadsHonorDefaultOnAndOffWithSingleSourceOwnership()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new LogicalBoundaryTextMeasurer();
        DocxParagraph footer = DocxTests.CreateDocxLayoutParagraph("public footer", 12d, 14d) with
        {
            LineSpacingPoints = null,
            Spacing = DocxParagraphSpacing.Empty with { BeforeValue = "0", AfterValue = "0", LineValue = "240", LineRuleValue = "auto" },
        };
        DocxDocument? control = null;
        DocxPlacedRelatedStoryLayout[] offSlices = [];
        int occupiedPage = -1;
        for (int pageHeight = 220; pageHeight <= 792 && control is null; pageHeight++)
        {
            DocxDocument candidate = WithEndnoteWidowControl(ReadWrappedFootnoteFlowFixture(40), false);
            var paragraphs = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [footer] };
            var elements = new Dictionary<string, IReadOnlyList<DocxBodyElement>> { ["default"] = [new DocxParagraphElement(footer)] };
            var settings = candidate.PageSettings with { FooterDistancePoints = 36d,
                HeightValue = (pageHeight * 20).ToString(CultureInfo.InvariantCulture),
                FooterParagraphsByType = paragraphs, FooterBodyElementsByType = elements };
            candidate = candidate with { PageHeightPoints = pageHeight, PageSettings = settings,
                FinalSectionBreak = candidate.FinalSectionBreak! with { PageSettings = settings },
                FooterParagraphs = [footer], FooterParagraphsByType = paragraphs, FooterBodyElementsByType = elements };
            DocxLayout layout = engine.Create(candidate, measurer, default);
            var placements = layout.Pages.SelectMany((p, i) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Id == "4").Select(s => (Story: s, Page: i))).ToArray();
            if (placements.Length != 2 || placements[0].Story.TextLines.Count != 1) continue;
            control = candidate; offSlices = placements.Select(p => p.Story).ToArray(); occupiedPage = placements[0].Page;
        }
        TestAssert.True(control is not null, "A bounded capacity scan finds a one-line fresh head with widow control disabled.");
        foreach (bool? widow in new bool?[] { null, true })
        {
            DocxDocument source = WithEndnoteWidowControl(control!, widow);
            var placements = engine.Create(source, measurer, default).Pages.SelectMany((p, i) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Id == "4").Select(s => (Story: s, Page: i))).ToArray();
            TestAssert.True(placements[0].Page > occupiedPage && placements[0].Story.TextLines.Count >= 2,
                "Default and explicit widow control defer the whole fresh head rather than publishing a one-line head.");
            TestAssert.True(offSlices.SelectMany(s => s.StoryLayout.TextLines).Select(l => (l.Text, l.SourceLineIndex))
                .SequenceEqual(placements.SelectMany(p => p.Story.StoryLayout.TextLines).Select(l => (l.Text, l.SourceLineIndex))),
                "The deferred head reaches the final drain exactly once and in source order.");
            TestAssert.True(ReferenceEquals(source.BodyElements, control!.BodyElements), "Deferral preserves the body reference owners.");
        }
    }

    public static void DocxHeaderFootnoteFramesPreserveCompleteCanonicalOwners()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (string rule in new[] { "auto", "atLeast", "exact" })
        {
            DocxDocument source = WithPlainFootnoteTestHeader(ReadWrappedFootnoteFlowFixture(200));
            source = source with
            {
                BodyElements = source.BodyElements.OfType<DocxParagraphElement>().Take(1).ToArray(),
                RelatedStories = source.RelatedStories.Where(s => !s.IsNormalStoryType || s.Id == "37").Select(s => s.IsNormalStoryType ? s with
                {
                    BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    {
                        LineSpacingPoints = rule == "auto" ? null : 18d,
                        Spacing = p.Spacing with { LineRuleValue = rule, LineValue = rule == "auto" ? "240" : "360" },
                        Runs = [p.Runs[0], p.Runs[1] with { FontSize = 10d }],
                    })).ToArray(),
                } : s).ToArray(),
            };
            DocxLayout layout = engine.Create(source, measurer, default);
            var slices = layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => s.StoryLayout.Story.IsNormalStoryType).ToArray();
            TestAssert.True(slices.Length > 1 && slices.All(s => s.UsesPlainFootnoteFrame),
                "Plain header-only frames share bounded fitting and continuation ownership with footer frames.");
            var canonical = layout.RelatedStories.Single(s => s.Story.IsNormalStoryType);
            TestAssert.True(canonical.TextLines.Select(l => (l.Text, l.SourceParagraph, l.SourceLineIndex)).SequenceEqual(
                slices.SelectMany(s => s.StoryLayout.TextLines).Select(l => (l.Text, l.SourceParagraph, l.SourceLineIndex))),
                "Every canonical line is placed once and in source order across header-only note continuations.");
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out var owners, default,
                DocxRelatedStoryKind.Footnote), "Plain header compositions share the source-owned content advance policy.");
            TestAssert.True(DocxNoteKerning.HasStableOwners(engine.Create(adjusted, measurer, default), owners, default,
                DocxRelatedStoryKind.Footnote), "The bounded content retry preserves complete header-frame owners.");
            TestAssert.True(ReferenceEquals(source.BodyElements, adjusted.BodyElements), "Body source reference owners remain unchanged.");
            TestAssert.True(ReferenceEquals(source.RelatedStories.Single(s => s.IsNormalStoryType).Paragraphs[0].Runs[0],
                adjusted.RelatedStories.Single(s => s.IsNormalStoryType).Paragraphs[0].Runs[0]), "The XML automatic marker remains immutable.");
        }
    }

    public static void DocxHeaderFootnoteFramesRetainRichKeepContextualAndScaledFallbacks()
    {
        var measurer = new NoteContentMetricMeasurer();
        foreach (string variant in new[] { "rich-header", "kept-note", "contextual-note", "scaled" })
        {
            DocxDocument source = ReadWrappedFootnoteFlowFixture(100);
            source = WithPlainFootnoteTestHeader(source, variant == "rich-header");
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    KeepRules = p.KeepRules with { KeepNext = variant == "kept-note" },
                    Spacing = p.Spacing with { ContextualSpacing = variant == "contextual-note" },
                })).ToArray(),
            } : s).ToArray() };
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, variant == "scaled" ? .5d : 1d);
            DocxLayout layout = engine.Create(source, measurer, default);
            TestAssert.True(layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => s.StoryLayout.Story.IsNormalStoryType)
                .All(s => !s.UsesPlainFootnoteFrame), "Rich headers, kept/contextual notes and scaled compositions retain static-frame fallback.");
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, layout, out var adjusted, out _, default,
                DocxRelatedStoryKind.Footnote) && ReferenceEquals(source, adjusted),
                "Unqualified header compositions retain the original source and typography policy.");
        }
    }

    public static void DocxContinuedFootnoteCapacityKeepsSeparatorInsetInsideAdmittedFrame()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        int middlePages = 0;
        foreach (bool custom in new[] { true, false })
        foreach (bool widow in new[] { true, false })
        for (int repeats = 100; repeats <= 140; repeats++)
        {
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32, repeats), false, true, 6);
            source = source with { RelatedStories = source.RelatedStories.Select(s =>
                s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Footnote ? s with
                {
                    BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    {
                        KeepRules = p.KeepRules with { WidowControl = widow, WidowControlValue = widow ? "1" : "0" },
                    })).ToArray(),
                } : s).ToArray() };
            DocxLayout layout = engine.Create(source, measurer, default);
            var feet = layout.Pages.SelectMany((p, index) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                .Select(s => (Page: index, Story: s))).ToArray();
            DocxRelatedStoryLayout canonical = layout.RelatedStories.Single(s => s.Story.IsNormalStoryType &&
                s.Story.Kind == DocxRelatedStoryKind.Footnote);
            TestAssert.Equal(string.Concat(canonical.TextLines.Select(l => l.Text)),
                string.Concat(feet.SelectMany(s => s.Story.StoryLayout.TextLines).Select(l => l.Text)));
            TestAssert.True(canonical.Story.Paragraphs.All(p => p.KeepRules.WidowControl == widow),
                "Frame capacity preserves authored widow metadata rather than replacing paragraph policy.");
            foreach (var foot in feet.Skip(1))
            {
                DocxLayoutPage page = layout.Pages[foot.Page];
                DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(s =>
                    s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && s.StoryLayout.Story.Type == DocxRelatedStoryType.ContinuationSeparator);
                TestAssert.True(separator.TopY <= page.Height - page.MarginTop + .000001d,
                    "An admitted continuation charges the complete separator inset before choosing a source take.");
                TestAssert.True(foot.Story.TextLines.All(l => ReferenceEquals(l.SourceParagraph, canonical.Story.Paragraphs[0])),
                    "The capacity correction retains complete canonical paragraph ownership.");
                if (foot != feet[^1]) middlePages++;
            }
        }
        TestAssert.True(middlePages > 0, "Both marker and widow states exercise middle continuation capacity boundaries.");
    }

    public static void DocxMiddleFootnoteContinuationAnchorsActualWidowAwareTakeAtFooter()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        int middlePages = 0;
        foreach (bool widow in new[] { true, false })
        for (int repeats = 100; repeats <= 140; repeats++)
        {
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(true, 32, repeats), false, true, 6);
            source = source with { RelatedStories = source.RelatedStories.Select(s =>
                s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Footnote ? s with
                {
                    BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    {
                        KeepRules = p.KeepRules with { WidowControl = widow, WidowControlValue = widow ? "1" : "0" },
                    })).ToArray(),
                } : s).ToArray() };
            DocxLayout layout = engine.Create(source, measurer, default);
            var feet = layout.Pages.SelectMany((p, index) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                .Select(s => (Page: index, Story: s))).ToArray();
            DocxRelatedStoryLayout canonical = layout.RelatedStories.Single(s => s.Story.IsNormalStoryType &&
                s.Story.Kind == DocxRelatedStoryKind.Footnote);
            TestAssert.Equal(string.Concat(canonical.TextLines.Select(l => l.Text)),
                string.Concat(feet.SelectMany(s => s.Story.StoryLayout.TextLines).Select(l => l.Text)));
            foreach (var foot in feet.Skip(1).SkipLast(1))
            {
                DocxLayoutPage page = layout.Pages[foot.Page];
                if (page.Items.Count != 0) continue;
                middlePages++;
                double bottom = Math.Max(page.MarginBottom, layout.FooterContentTopByPage[foot.Page]);
                TestAssert.True(Math.Abs(foot.Story.TopY - foot.Story.Height - bottom) < .000001d,
                    "An occupied middle continuation anchors its actual slice at the selected footer edge, including widow-tail adjustments.");
                TestAssert.True(foot.Story.TextLines.All(l => ReferenceEquals(l.SourceParagraph, canonical.Story.Paragraphs[0])),
                    "Every line retains the complete original canonical paragraph owner.");
                TestAssert.Equal(1, page.PlacedRelatedStories.Count(s => s.SeparatorY is not null));
            }
        }
        TestAssert.True(middlePages > 0, "The boundary sweep exercises actual empty middle continuation pages.");
    }

    public static void DocxScaledFootnoteExactAdvanceMatchesIndependentOffice()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("The Word control uses installed Arial metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        string input = Path.Combine(cases, "scaled-note-body.docx");
        using var reference = System.Text.Json.JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "scaled-note-office-advances.json")));
        TestAssert.Equal(reference.RootElement.GetProperty("InputSha256").GetString(),
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(input))));
        var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        var emission = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .InspectTextEmission(document).Lines.Where(line => line.ContainerStoryKind == "Body" && line.SourceBlockIndex == 1).ToArray();
        TestAssert.Equal(3, emission.Length);
        var gaps = reference.RootElement.GetProperty("FirstParagraphBaselineGapsPoints");
        for (int index = 0; index < 2; index++)
            TestAssert.True(Math.Abs(emission[index].Segments[0].BaselineY - emission[index + 1].Segments[0].BaselineY - gaps[index].GetDouble()) < .05d,
                "An automatic superscript must not change the independently measured printed exact-line advance.");
    }

    public static void DocxScaledFootnoteAdmissionMatchesIndependentPageMembership()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("The Office controls use installed Arial metrics.");
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
                "CompositionProbes", "scaled-note-" + container + ".docx")), mode);
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.True(pages.Length == 1, $"{container} / {mode}: expected one page, got {pages.Length}.");
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            for (int section = 1; section <= 6; section++)
            for (int line = 1; line <= 3; line++)
                TestAssert.True(text.Contains($"Section {section} line {line}.", StringComparison.Ordinal), "Body admission preserves all source lines.");
            TestAssert.True(text.Contains("Public calculation note.", StringComparison.Ordinal), "Note content remains beside its reference page.");
        }
    }

    public static void DocxAtomicFootnoteHeadIgnoresFutureContinuationRule()
    {
        var story = new DocxRelatedStory(DocxRelatedStoryKind.Footnote, "/word/footnotes.xml", "1", [], [], [], null);
        var run = new DocxTextRun("Note", 10d, null, false, false, false, null, null);
        var line = new DocxTextLineLayout("Note", run, 10d, 0d, -10d, 100d, [], null, null, null, null, 12d,
            null, null, false, null, null, null, null, null, null, null, null, null, null, null, null, false);
        var layout = new DocxRelatedStoryLayout(story, 0, [line], [], [], [], 12d);
        var separator = new DocxRelatedStoryLayout(story, 0, [line with { Text = " " }], [], [], [], 8d);
        var location = new DocxLayoutEngine.DocxInlineReferenceLocation(0,
            DocxTests.CreateDocxLayoutParagraph("Anchor", 10d, 12d),
            new DocxInlineReference(DocxRelatedStoryKind.Footnote, "1", null, "1", 0, 0, 0));
        var head = new DocxLayoutEngine.InFlightRelatedStory(layout, separator, separator, 0, location);
        var noFutureRule = new DocxLayoutEngine.InFlightRelatedStory(layout, separator, null, 0, location);
        double headHeight = DocxLayoutEngine.FootnoteRemainderHeight(head, [12d], null);
        TestAssert.Equal(DocxLayoutEngine.FootnoteRemainderHeight(noFutureRule, [12d], null), headHeight);
        head.SeparatorPlaced = true;
        TestAssert.True(DocxLayoutEngine.FootnoteRemainderHeight(head, [12d], null) > 12d,
            "A continued one-line tail must still charge its present continuation rule.");
        var split = new DocxLayoutEngine.InFlightRelatedStory(layout with { TextLines = [line, line] }, separator, separator, 0, location);
        TestAssert.True(DocxLayoutEngine.FootnoteRemainderHeight(split, [12d, 12d], null) > headHeight + 12d,
            "Potentially split multi-line notes retain their qualified conservative reserve.");
    }

    public static void DocxAtomicFootnotesMatchIndependentPageAdmission()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("The Office controls use installed Arial metrics.");
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes",
                "complete-note-" + container + ".docx"), mode);
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            for (int section = 1; section <= 6; section++)
            for (int line = 1; line <= 3; line++)
                TestAssert.True(text.Contains($"Section {section} line {line}.", StringComparison.Ordinal), "Every admitted line remains visible in the PDF stream.");
            TestAssert.True(text.Contains("Public calculation note.", StringComparison.Ordinal), "The complete note survives page admission.");
            TestAssert.Equal(1, pages.Single().Annotations.Count(annotation => annotation.Destination is not null));
        }
    }

    public static void DocxSimpleSectionPaginationUsesFollowingStartType()
    {
        foreach ((string? preceding, string? following, int expected) in new[]
        {
            ((string?)"nextPage", (string?)"continuous", 1),
            ((string?)"continuous", (string?)"nextPage", 2),
            ((string?)null, (string?)"continuous", 1),
            ((string?)"continuous", (string?)null, 2)
        })
        {
            DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, _, parts) =>
            {
                SetSectionTypes(xml, preceding, following);
                SetNotePresence(xml, parts, "footnote", false, false);
            });
            TestAssert.Equal(expected, RenderSectionPages(document).Length);
            TestAssert.Equal(preceding, document.BodyElements.OfType<DocxSectionBreakElement>().Single().TypeValue?.ToValueString());
            TestAssert.Equal(following, document.FinalSectionBreak!.TypeValue?.ToValueString());
        }
    }

    public static void DocxSimpleSectionPaginationSeparatesPageBottomFootnotes()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach ((bool first, bool second) in new[] { (true, true), (true, false), (false, true), (false, false) })
        foreach (string? position in kind == "footnote" ? new string?[] { "pageBottom", "beneathText", null } : ["docEnd"])
        {
            DocxDocument document = ReadSectionNumberingFixture(kind, (xml, _, parts) =>
            {
                SetSectionTypes(xml, "continuous", "continuous");
                SetNotePresence(xml, parts, kind, first, second);
                foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                {
                    XElement props = section.Element(NoteWord + kind + "Pr")!;
                    props.Elements(NoteWord + "pos").Remove();
                    if (position is not null) props.AddFirst(new XElement(NoteWord + "pos", new XAttribute(NoteWord + "val", position)));
                }
            });
            PdfPage[] pages = RenderSectionPages(document);
            bool splits = kind == "footnote" && first && position != "beneathText";
            TestAssert.Equal(splits ? 2 : 1, pages.Length);
            PdfLinkAnnotation[] links = pages.SelectMany(p => p.Annotations).ToArray();
            TestAssert.Equal((first ? 2 : 0) + (second ? 2 : 0), links.Length);
            TestAssert.True(links.All(l => l.IsDestination && l.Width > 0 && l.Height > 0), "Visible note markers remain linked.");
            if (splits && second)
            {
                TestAssert.Equal(2, pages[0].Annotations.Count);
                TestAssert.Equal(2, pages[1].Annotations.Count);
                TestAssert.True(pages[0].Annotations.All(l => l.Destination!.Value.PageIndex == 0) &&
                    pages[1].Annotations.All(l => l.Destination!.Value.PageIndex == 1), "Notes follow the section's source page.");
            }
        }
    }

    public static void DocxSimpleSectionPaginationIgnoresNotesAlreadyPlacedOnEarlierPages()
    {
        DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, _, parts) =>
        {
            SetSectionTypes(xml, "continuous", "continuous");
            XElement[] paragraphs = xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").ToArray();
            paragraphs[0].Add(new XElement(NoteWord + "r", new XElement(NoteWord + "br", new XAttribute(NoteWord + "type", "page"))));
            paragraphs[1].Descendants(NoteWord + "footnoteReference").Single().Parent!.Remove();
            XDocument notes = XDocument.Parse(parts["word/footnotes.xml"]);
            notes.Root!.Elements(NoteWord + "footnote").Single(n => (string?)n.Attribute(NoteWord + "id") == "4").Remove();
            parts["word/footnotes.xml"] = notes.ToString();
        });
        PdfPage[] pages = RenderSectionPages(document);
        TestAssert.Equal(2, pages.Length);
        TestAssert.Equal(1, pages[0].Annotations.Count);
        TestAssert.Equal(2, pages[1].Annotations.Count);
        TestAssert.True(pages[1].Annotations.All(l => l.Destination!.Value.PageIndex == 1), "An earlier note must not insert a third page.");
    }

    public static void DocxSimpleSectionPaginationRetainsExcludedFlow()
    {
        int variant = 0;
        foreach (Action<XDocument, XDocument, Dictionary<string, string>> modify in new Action<XDocument, XDocument, Dictionary<string, string>>[]
        {
            (xml, _, _) =>
            {
                XElement body = xml.Root!.Element(NoteWord + "body")!;
                XElement paragraph = body.Elements(NoteWord + "p").First();
                paragraph.Remove();
                body.AddFirst(new XElement(NoteWord + "tbl", new XElement(NoteWord + "tr", new XElement(NoteWord + "tc", paragraph))));
            },
            (xml, _, _) =>
            {
                foreach (XElement section in xml.Descendants(NoteWord + "sectPr"))
                    section.Add(new XElement(NoteWord + "cols", new XAttribute(NoteWord + "num", "2"), new XAttribute(NoteWord + "space", "720")));
            },
            (xml, _, _) => xml.Descendants(NoteWord + "sectPr").Last().Element(NoteWord + "pgMar")!.SetAttributeValue(NoteWord + "left", "1800"),
            (xml, _, _) => xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").Skip(2).Remove(),
            (xml, _, _) => xml.Root!.Element(NoteWord + "body")!.Elements(NoteWord + "p").First().Add(
                new XElement(NoteWord + "ins", new XElement(NoteWord + "r", new XElement(NoteWord + "t", "tracked"))))
        })
        {
            DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, settings, parts) =>
            {
                SetSectionTypes(xml, "continuous", "continuous");
                modify(xml, settings, parts);
            });
            PdfPage[] pages = RenderSectionPages(document);
            // Atomic heads no longer reserve unused future continuation rules.
            // Column-flow fallback must remain; the old second page was capacity
            // overhead, rather than an independently qualified section boundary.
            TestAssert.True(pages.Length == 1, $"Excluded variant {variant} must retain its admitted source page; got {pages.Length}.");
            if (variant == 1)
            {
                var layout = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).InspectLayout(document);
                TestAssert.True(layout.Pages.All(page => page.ColumnFrameCount == 2), "The excluded case must retain two-column placement.");
                TestAssert.Equal(4, pages.Single().Annotations.Count);
                TestAssert.True(pages.Single().Annotations.All(link => link.Destination is { PageIndex: 0 }),
                    "All four source markers retain note destinations on their admitted page.");
            }
            variant++;
        }
        DocxDocument ordinary = ReadSectionNumberingFixture("footnote", (xml, _, _) => SetSectionTypes(xml, "continuous", "continuous"));
        var scaled = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, wordCompatiblePrintScale: 0.5d);
        TestAssert.Equal(1, scaled.Create(ordinary, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None).Pages.Count);
    }

    public static void DocxSimpleSectionPaginationPreservesLabelsAndRepeatedDestinations()
    {
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[]
            { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            DocxDocument document = ReadSectionNumberingFixture("footnote", (xml, _, _) =>
            {
                SetSectionTypes(xml, "continuous", "continuous");
                SetNoteSection(xml.Descendants(NoteWord + "sectPr").Last(), "footnote", "lowerRoman", "4", "eachSect");
            });
            CheckNoteLabels(document, ["iv", "v", "iv", "v"]);
            PdfPage[] pages = RenderSectionPages(document, geometry);
            PdfLinkAnnotation[] links = pages.SelectMany(p => p.Annotations).ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Equal(4, links.Length);
            TestAssert.Equal(4, links.Select(l => l.Destination).Distinct().Count());
            TestAssert.Equal(105, document.Settings.FootnoteReferenceSettings.NumberStart!.Value);
            TestAssert.Equal(DocxSectionBreakType.Continuous, document.BodyElements.OfType<DocxSectionBreakElement>().Single().TypeValue!.Value);
        }
    }

    private static PdfPage[] RenderSectionPages(DocxDocument document,
        OoxPdfDocxMarkupGeometryMode geometry = OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup) =>
        new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry)
            .RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().ToArray();

    private static void SetSectionTypes(XDocument xml, string? preceding, string? following)
    {
        XElement[] sections = xml.Descendants(NoteWord + "sectPr").ToArray();
        for (int i = 0; i < sections.Length; i++)
        {
            sections[i].Elements(NoteWord + "type").Remove();
            string? value = i == 0 ? preceding : following;
            if (value is not null) sections[i].Add(new XElement(NoteWord + "type", new XAttribute(NoteWord + "val", value)));
        }
    }

    private static void SetNotePresence(XDocument xml, Dictionary<string, string> parts, string kind, bool first, bool second)
    {
        XElement[] references = xml.Descendants(NoteWord + kind + "Reference").ToArray();
        XDocument notes = XDocument.Parse(parts["word/" + kind + "s.xml"]);
        for (int i = 0; i < references.Length; i++)
        {
            if (i < 2 ? first : second) continue;
            string id = (string)references[i].Attribute(NoteWord + "id")!;
            references[i].Parent!.Remove();
            notes.Root!.Elements(NoteWord + kind).Single(n => (string?)n.Attribute(NoteWord + "id") == id).Remove();
        }
        parts["word/" + kind + "s.xml"] = notes.ToString();
    }
}
