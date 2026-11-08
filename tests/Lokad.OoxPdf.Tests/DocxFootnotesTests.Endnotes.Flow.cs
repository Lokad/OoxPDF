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
    public static void DocxDocumentEndnotesFollowTrailingBodySpacingAtNominalScale()
    {
        var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        double Top(string fixture) => renderer.InspectLayout(ReadNoteNavigationFixture(fixture)).Pages[0]
            .PlacedRelatedStories.Single(s => s.Kind == "Endnote" && s.Id == "1").TopY;
        double zero = Top("note-nav-endnote-body");
        TestAssert.True(Math.Abs(zero - Top("note-body-after-8") - 8d) < 0.001d,
            "Document endnotes must follow eight points of trailing body spacing.");
        TestAssert.True(Math.Abs(zero - Top("note-body-after-24") - 24d) < 0.001d,
            "Document endnotes must follow twenty-four points of trailing body spacing.");
        TestAssert.True(Math.Abs(zero - Top("note-body-default-after-8") - 8d) < 0.001d,
            "Inherited body spacing participates in document-endnote flow.");
        foreach (string fixture in new[] { "note-separator-after-8", "note-separator-after-24", "note-separator-before-24", "note-separator-default-after-8" })
        {
            TestAssert.True(Math.Abs(zero - Top(fixture)) < 0.001d,
                "The separator's own paragraph spacing must not replace body flow spacing.");
        }
    }

    public static void DocxDocumentEndnoteBoundariesCollapseNumericParagraphSpacing()
    {
        foreach ((int before, int after) in new[] { (0, 0), (0, 4), (4, 0), (8, 8), (8, 4), (4, 8) })
        {
            DocxDocument document = ReadDocumentEndnoteFlowFixture(before, after);
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var notes = renderer.InspectLayout(document).Pages.Single().PlacedRelatedStories
                .Where(s => s.Kind == "Endnote" && s.Type is null).ToArray();
            TestAssert.Equal(2, notes.Length);
            double expected = notes[0].Height - Math.Min(before, after);
            TestAssert.True(Math.Abs(notes[0].TopY - notes[1].TopY - expected) < .000001d,
                "Independent endnotes flow as adjacent paragraphs without an extra separator gap.");
            TestAssert.True(document.RelatedStories.Where(s => s.Kind == DocxRelatedStoryKind.Endnote && s.IsNormalStoryType)
                .All(s => s.Paragraphs[0].SpacingBeforePoints == before && s.Paragraphs[0].SpacingAfterPoints == after),
                "Collapsed placement retains each source paragraph's authored spacing.");
        }
    }

    public static void DocxDocumentEndnoteBoundariesRetainUnqualifiedStoryFallbacks()
    {
        foreach (Func<DocxParagraph, IReadOnlyList<DocxBodyElement>> modify in new Func<DocxParagraph, IReadOnlyList<DocxBodyElement>>[]
        {
            p => [new DocxParagraphElement(p with { KeepRules = p.KeepRules with { KeepNext = true } })],
            p => [new DocxParagraphElement(p with { Spacing = p.Spacing with { ContextualSpacing = true } })],
            p => [new DocxParagraphElement(p with { Spacing = p.Spacing with { BeforeAutoSpacingValue = "1" } })],
            p => [new DocxParagraphElement(p with { Spacing = p.Spacing with { BeforeValue = "broken" } })],
            p => [new DocxParagraphElement(p with { LineSpacingPoints = null,
                Spacing = p.Spacing with { LineValue = "480", LineRuleValue = "exact" } })],
            p => [new DocxParagraphElement(p with { LineSpacingPoints = null,
                Spacing = p.Spacing with { LineValue = "480", LineRuleValue = "atLeast" } })],
            p => [new DocxParagraphElement(p with { Spacing = p.Spacing with { LineRuleValue = "broken" } })],
            p => [new DocxParagraphElement(p), new DocxParagraphElement(p)],
        })
        {
            DocxDocument document = ReadDocumentEndnoteFlowFixture(0, 0);
            DocxRelatedStory first = document.RelatedStories.First(s => s.Kind == DocxRelatedStoryKind.Endnote && s.IsNormalStoryType);
            document = document with { RelatedStories = document.RelatedStories.Select(s => ReferenceEquals(s, first)
                ? s with { BodyElements = modify(s.Paragraphs[0]) } : s).ToArray() };
            var notes = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document).Pages.Single()
                .PlacedRelatedStories.Where(s => s.Kind == "Endnote" && s.Type is null).ToArray();
            TestAssert.Equal(2, notes.Length);
            TestAssert.True(Math.Abs(notes[0].TopY - notes[1].TopY - notes[0].Height - 3d) < .000001d,
                "Unqualified notes retain the legacy spacing and slicing path.");
        }
        DocxDocument scaledDocument = ReadDocumentEndnoteFlowFixture(0, 0);
        DocxLayout scaled = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, .5d)
            .Create(scaledDocument, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        var scaledNotes = scaled.Pages.Single().PlacedRelatedStories
            .Where(s => s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && s.StoryLayout.Story.IsNormalStoryType).ToArray();
        TestAssert.True(Math.Abs(scaledNotes[0].TopY - scaledNotes[1].TopY - scaledNotes[0].Height - 3d) < .000001d,
            "Scaled document endnotes retain their prior boundary geometry.");
        DocxDocument headerDocument = ReadDocumentEndnoteFlowFixture(0, 0) with
        {
            HeaderParagraphs = [DocxTests.CreateDocxLayoutParagraph("Public header", 12d, 14d)]
        };
        var headerNotes = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(headerDocument).Pages.Single()
            .PlacedRelatedStories.Where(s => s.Kind == "Endnote" && s.Type is null).ToArray();
        TestAssert.True(Math.Abs(headerNotes[0].TopY - headerNotes[1].TopY - headerNotes[0].Height - 3d) < .000001d,
            "The shared plain-body policy preserves composed header flow.");
    }

    public static void DocxDocumentEndnoteBoundariesEmitActualDestinationTops()
    {
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            DocxDocument document = ReadDocumentEndnoteFlowFixture(8, 4);
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
            var notes = renderer.InspectLayout(document).Pages.Single().PlacedRelatedStories
                .Where(s => s.Kind == "Endnote" && s.Type is null).ToArray();
            var links = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single().Annotations.ToArray();
            TestAssert.Equal(2, links.Length);
            for (int i = 0; i < links.Length; i++)
            {
                TestAssert.Equal(0, links[i].Destination!.Value.PageIndex);
                double targetTop = notes[i].TopY - (i == 0 ? 0d : 4d);
                TestAssert.True(Math.Abs(links[i].Destination!.Value.Top!.Value - targetTop) < .000001d,
                    "Emitted note links exclude the shared boundary overlap at their source-owned destination.");
            }
        }
    }

    public static void DocxDocumentEndnoteDestinationsUseNeighborSpacingWithoutMovingContent()
    {
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        foreach ((int previousAfter, int currentBefore) in new[] { (0, 8), (8, 0), (8, 4), (4, 8), (9, 2), (1, 7) })
        {
            DocxDocument document = ReadDocumentEndnoteFlowFixture(8, previousAfter);
            DocxRelatedStory current = document.RelatedStories.Last(s => s.Kind == DocxRelatedStoryKind.Endnote && s.IsNormalStoryType);
            DocxParagraph paragraph = current.Paragraphs.Single();
            document = document with { RelatedStories = document.RelatedStories.Select(s => ReferenceEquals(s, current) ? s with
            {
                BodyElements = [new DocxParagraphElement(paragraph with
                {
                    SpacingBeforePoints = currentBefore,
                    Spacing = paragraph.Spacing with { BeforeValue = (currentBefore * 20).ToString(CultureInfo.InvariantCulture) },
                })],
            } : s).ToArray() };
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
            var notes = renderer.InspectLayout(document).Pages.Single().PlacedRelatedStories
                .Where(s => s.Kind == "Endnote" && s.Type is null).ToArray();
            var links = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single().Annotations.ToArray();
            TestAssert.Equal(2, links.Length);
            TestAssert.True(Math.Abs(links[0].Destination!.Value.Top!.Value - notes[0].TopY) < .000001d,
                "The first note retains its target, including its own before spacing.");
            TestAssert.True(Math.Abs(links[1].Destination!.Value.Top!.Value - notes[1].TopY + Math.Min(previousAfter, currentBefore)) < .000001d,
                "Only the previous after-gap and current before-gap define the current destination overlap.");
            TestAssert.True(Math.Abs(notes[0].TopY - notes[1].TopY - notes[0].Height + Math.Min(previousAfter, currentBefore)) < .000001d,
                "Destination ownership preserves the independently qualified visible placement.");
            TestAssert.Equal(0, links[1].Destination!.Value.PageIndex);
        }
    }

    public static void DocxDocumentEndnoteDestinationsRetainUnqualifiedFrameTargets()
    {
        DocxDocument source = ReadDocumentEndnoteFlowFixture(8, 4);
        foreach (DocxDocument document in new[]
        {
            source with { HeaderParagraphs = [source.Paragraphs[0]] },
            source with { RelatedStories = source.RelatedStories.Select(s => s.Kind == DocxRelatedStoryKind.Endnote && s.IsNormalStoryType ? s with
            {
                BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                {
                    LineSpacingPoints = null, Spacing = p.Spacing with { LineRuleValue = "exact", LineValue = "480" },
                })).ToArray(),
            } : s).ToArray() },
        })
        {
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var notes = renderer.InspectLayout(document).Pages.Single().PlacedRelatedStories
                .Where(s => s.Kind == "Endnote" && s.Type is null).ToArray();
            var links = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single().Annotations.ToArray();
            TestAssert.Equal(2, links.Length);
            for (int i = 0; i < notes.Length; i++)
                TestAssert.True(Math.Abs(links[i].Destination!.Value.Top!.Value - notes[i].TopY) < .000001d,
                    "An unqualified frame keeps the established destination and visible placement.");
        }
    }

    public static void DocxDocumentEndnoteWrappedBoundariesUseCompleteParagraphFlow()
    {
        foreach (var geometry in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        foreach ((int before, int after) in new[] { (0, 0), (0, 4), (4, 0), (8, 8), (8, 4), (4, 8) })
        {
            DocxDocument document = ReadWrappedDocumentEndnoteFlowFixture(before, after, 8);
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
            var notes = renderer.InspectLayout(document).Pages.Single().PlacedRelatedStories
                .Where(s => s.Kind == "Endnote" && s.Type is null).ToArray();
            TestAssert.Equal(2, notes.Length);
            TestAssert.True(notes.All(n => n.TextLineCount > 1 && n.ContentTopOffset == 0d),
                "The shared flow owns complete wrapped paragraphs rather than partial story slices.");
            double overlap = Math.Min(before, after);
            TestAssert.True(Math.Abs(notes[0].TopY - notes[1].TopY - notes[0].Height + overlap) < .000001d,
                "Wrapped notes use the final paragraph after-gap and next paragraph before-gap.");
            var links = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single().Annotations.ToArray();
            TestAssert.Equal(2, links.Length);
            TestAssert.True(Math.Abs(links[0].Destination!.Value.Top!.Value - notes[0].TopY) < .000001d,
                "The first wrapped note retains its source-owned destination top.");
            TestAssert.True(Math.Abs(links[1].Destination!.Value.Top!.Value - notes[1].TopY + overlap) < .000001d,
                "Wrapped note destinations consume the same source-owned collapsed boundary.");
        }
    }

    public static void DocxDocumentEndnoteWrappedBoundariesOwnFirstSliceDestinations()
    {
        DocxDocument document = ReadWrappedDocumentEndnoteFlowFixture(8, 4, 200);
        var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxLayoutSnapshot snapshot = renderer.InspectLayout(document);
        var notes = snapshot.Pages.SelectMany((page, index) => page.PlacedRelatedStories
            .Where(s => s.Kind == "Endnote" && s.Type is null).Select(s => (Note: s, Page: index))).ToArray();
        TestAssert.True(notes.Any(n => n.Note.ContentTopOffset > 0d),
            "The oversized complete block must retain continued story slices.");
        var firstPlacements = notes.GroupBy(n => n.Note.Id).Select(g => g.First()).ToArray();
        var links = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().SelectMany(p => p.Annotations).ToArray();
        TestAssert.Equal(2, firstPlacements.Length);
        TestAssert.Equal(2, links.Length);
        for (int i = 0; i < links.Length; i++)
        {
            TestAssert.Equal(firstPlacements[i].Page, links[i].Destination!.Value.PageIndex);
            TestAssert.True(Math.Abs(links[i].Destination!.Value.Top!.Value - firstPlacements[i].Note.TopY + (i == 0 ? 0d : 4d)) < .000001d,
                "Continued notes navigate to the first source-owned slice below its shared paragraph boundary.");
        }
    }

    public static void DocxContinuedEndnoteParagraphAfterSpacingBelongsToTheFinalSourceLine()
    {
        DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(8, 4, 200);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
        DocxRelatedStoryLayout story = layout.RelatedStories.First(s => s.Story.IsNormalStoryType);
        DocxPlacedRelatedStoryLayout[] slices = layout.Pages.SelectMany(p => p.PlacedRelatedStories)
            .Where(p => ReferenceEquals(p.StoryLayout.Story, story.Story)).ToArray();
        TestAssert.True(slices.Length > 1 && slices.All(s => s.UsesPlainDocumentEndnoteTypography),
            "The regression exercises qualified continued paragraph slices.");
        TestAssert.Equal(4d, story.TextLines[0].ParagraphAfterSpacing!.Value);
        TestAssert.True(story.TextLines[^1].ParagraphAfterSpacing is null,
            "Source layout metadata remains on the first line; slice ownership does not rewrite it.");
        double consumed = 0d;
        for (int i = 0; i < slices.Length; i++)
        {
            DocxPlacedRelatedStoryLayout slice = slices[i];
            double expected = slice.StoryLayout.TextLines.Sum(line => line.LineHeight ?? line.SingleLineHeight ?? line.FontSize);
            if (i == 0) expected += 8d;
            if (i == slices.Length - 1) expected += 4d;
            TestAssert.True(Math.Abs(slice.Height - expected) < .000001d,
                "Only the final source-owned slice consumes the paragraph's after-space.");
            TestAssert.True(Math.Abs(slice.ContentTopOffset - consumed) < .000001d,
                "Continuation offsets are cumulative actual slice charges.");
            consumed += slice.Height;
        }
        TestAssert.True(Math.Abs(consumed - story.ContentHeight) < .000001d,
            "Deferring ownership preserves the full paragraph charge exactly once.");
        TestAssert.Equal(4d, source.RelatedStories.First(s => s.IsNormalStoryType).Paragraphs[0].SpacingAfterPoints);
    }

    public static void DocxContinuedEndnoteParagraphBoundariesShareNumericSpacingAndRetainKeepFallbacks()
    {
        foreach ((int before, int after) in new[] { (0, 0), (0, 4), (8, 0), (8, 4), (3, 7), (7, 3) })
        {
            DocxDocument source = ReadWrappedDocumentEndnoteFlowFixture(before, after, 200);
            DocxRelatedStory first = source.RelatedStories.First(s => s.IsNormalStoryType);
            source = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType && !ReferenceEquals(s, first)
                ? s with { BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    { Runs = [p.Runs[0], p.Runs[1] with { Text = "Public short following note." }] })).ToArray() } : s).ToArray() };
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var measurer = new DocxTests.FamilyWidthTextMeasurer();
            DocxLayout layout = engine.Create(source, measurer, default);
            var placements = layout.Pages.SelectMany((p, page) => p.PlacedRelatedStories.Where(s => s.StoryLayout.Story.IsNormalStoryType)
                .Select(s => (Story: s, Page: page))).ToArray();
            var previous = placements.Last(s => ReferenceEquals(s.Story.StoryLayout.Story, first));
            var current = placements.First(s => !ReferenceEquals(s.Story.StoryLayout.Story, first));
            TestAssert.True(placements.Count(s => ReferenceEquals(s.Story.StoryLayout.Story, first)) > 1,
                "The preceding source paragraph actually crosses pages.");
            TestAssert.Equal(previous.Page, current.Page);
            TestAssert.True(Math.Abs(previous.Story.TopY - current.Story.TopY - previous.Story.Height + Math.Min(before, after)) < .000001d,
                "The final slice and following paragraph share the same numeric boundary rule as fitting notes.");
            TestAssert.Equal((double)Math.Min(before, after), current.Story.NoteDestinationTopOffset);
            TestAssert.True(placements.Where(s => s.Story.ContentTopOffset > 0d).All(s => s.Story.NoteDestinationTopOffset == 0d),
                "Continued slices do not acquire a second first-placement navigation offset.");
            DocxDocument kept = source with { RelatedStories = source.RelatedStories.Select(s => s.IsNormalStoryType
                ? s with { BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                    { KeepRules = p.KeepRules with { KeepLines = true } })).ToArray() } : s).ToArray() };
            TestAssert.True(engine.Create(kept, measurer, default).Pages.SelectMany(p => p.PlacedRelatedStories)
                .Where(s => s.StoryLayout.Story.IsNormalStoryType).All(s => !s.UsesPlainDocumentEndnoteTypography && s.NoteDestinationTopOffset == 0d),
                "Kept continued paragraphs retain their explicit layout and destination fallback.");
        }
    }

    public static void DocxContinuedEndnoteWidowHeadsHonorDefaultOnAndOffWithoutChangingOwners()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        DocxDocument? control = null;
        DocxPlacedRelatedStoryLayout? oldHead = null;
        int oldPage = -1;
        string? targetId = null;
        for (int repeats = 120; repeats <= 340 && control is null; repeats++)
        {
            DocxDocument candidate = WithEndnoteWidowControl(ReadWrappedDocumentEndnoteFlowFixture(8, 4, repeats), false);
            DocxLayout layout = engine.Create(candidate, measurer, default);
            for (int page = 0; page < layout.Pages.Count && control is null; page++)
            foreach (DocxPlacedRelatedStoryLayout placed in layout.Pages[page].PlacedRelatedStories)
            {
                if (!placed.StoryLayout.Story.IsNormalStoryType || placed.ContentTopOffset != 0d || placed.TextLines.Count != 1 ||
                    layout.RelatedStories.Single(s => ReferenceEquals(s.Story, placed.StoryLayout.Story)).TextLines.Count <= 1) continue;
                control = candidate; oldHead = placed; oldPage = page; targetId = placed.StoryLayout.Story.Id;
                break;
            }
        }
        TestAssert.True(control is not null && oldHead is not null && targetId is not null,
            "The capacity scan finds a genuine one-line note head with authored widow control disabled.");
        foreach (bool? widow in new bool?[] { null, true })
        {
            DocxDocument source = WithEndnoteWidowControl(control!, widow);
            DocxLayout layout = engine.Create(source, measurer, default);
            var first = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.Id == targetId && s.StoryLayout.Story.IsNormalStoryType)
                .Select(s => (Story: s, Page: index))).First();
            TestAssert.Equal(oldPage + 1, first.Page);
            TestAssert.True(first.Story.TextLines.Count > 1 && first.Story.ContentTopOffset == 0d,
                "Default and explicit-on widow rules move the complete head onto an empty continuation frame.");
            TestAssert.Equal(oldHead!.NoteDestinationTopOffset, first.Story.NoteDestinationTopOffset);
            TestAssert.True(source.BodyElements.Zip(control!.BodyElements).All(p => ReferenceEquals(p.First, p.Second)),
                "Page-turn policy preserves source body and inline-reference owners.");
            TestAssert.True(source.RelatedStories.Where(s => s.IsNormalStoryType).All(s => s.Paragraphs[0].KeepRules.WidowControl == widow),
                "Layout never rewrites authored widow metadata.");
        }
    }

    public static void DocxContinuedEndnoteWidowHeadsMakeProgressOnOneLineContinuationFrames()
    {
        DocxDocument source = WithEndnoteWidowControl(ReadWrappedDocumentEndnoteFlowFixture(0, 0, 20), true);
        DocxPageSettings settings = source.PageSettings with { HeightValue = "1200", MarginTopValue = "200", MarginBottomValue = "200" };
        source = source with { PageHeightPoints = 60d, MarginTopPoints = 10d, MarginBottomPoints = 10d,
            PageSettings = settings, FinalSectionBreak = source.FinalSectionBreak! with { PageSettings = settings },
            BodyElements = source.BodyElements.OfType<DocxParagraphElement>().Where(p => p.Paragraph.InlineReferences.Count != 0)
                .Select(p => (DocxBodyElement)(p with { Paragraph = p.Paragraph with
                {
                    Runs = p.Paragraph.Runs.Select(r => r.VerticalAlignment == DocxRunVerticalAlignment.Baseline ? r with { Text = "body" } : r).ToArray(),
                    SpacingBeforePoints = 0d, SpacingAfterPoints = 0d, LineSpacingPoints = 8d,
                    Spacing = p.Paragraph.Spacing with { BeforeValue = "0", AfterValue = "0", LineRuleValue = "exact", LineValue = "160" },
                } })).ToArray() };
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
        var stories = layout.RelatedStories.Where(s => s.Story.IsNormalStoryType).ToArray();
        var placements = layout.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => s.StoryLayout.Story.IsNormalStoryType).ToArray();
        TestAssert.True(placements.Any(s => s.UsesPlainDocumentEndnoteTypography && s.TextLines.Count == 1),
            "The deliberately short empty continuation frame exercises mandatory one-line progress.");
        foreach (DocxRelatedStoryLayout story in stories)
        {
            var owned = placements.Where(p => ReferenceEquals(p.StoryLayout.Story, story.Story)).ToArray();
            TestAssert.Equal(story.TextLines.Count, owned.Sum(p => p.TextLines.Count));
            TestAssert.Equal(1, owned.Count(p => p.ContentTopOffset == 0d));
            TestAssert.True(owned.All(p => p.Height >= 0d && double.IsFinite(p.TopY)),
                "Every source line is staged once with finite bounded placement.");
        }
        TestAssert.True(layout.Pages.Count <= stories.Sum(s => s.TextLines.Count) + source.Paragraphs.Count + 2,
            "A fresh frame cannot be skipped repeatedly for its inability to hold two lines.");
    }

    public static void DocxContinuedEndnoteWidowTailsHonorDefaultOnAndOffWithSourceOwnedCharges()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        DocxDocument? control = null;
        DocxPlacedRelatedStoryLayout[] oldSlices = [];
        string? targetId = null;
        for (int repeats = 20; repeats <= 340 && control is null; repeats++)
        {
            DocxDocument candidate = WithEndnoteWidowControl(ReadWrappedDocumentEndnoteFlowFixture(8, 4, repeats), false);
            string? id = candidate.RelatedStories.First(s => s.IsNormalStoryType).Id;
            DocxPlacedRelatedStoryLayout[] slices = engine.Create(candidate, measurer, default).Pages
                .SelectMany(p => p.PlacedRelatedStories).Where(s => s.StoryLayout.Story.Id == id && s.StoryLayout.Story.IsNormalStoryType).ToArray();
            if (slices.Length != 2 || slices[0].TextLines.Count <= 2 || slices[^1].TextLines.Count != 1 ||
                !slices.All(s => s.UsesPlainDocumentEndnoteTypography)) continue;
            control = candidate; oldSlices = slices; targetId = id;
        }
        TestAssert.True(control is not null, "A bounded capacity scan finds a genuine one-line final continuation with widow control disabled.");
        foreach (bool? widow in new bool?[] { null, true })
        {
            DocxDocument source = WithEndnoteWidowControl(control!, widow);
            DocxLayout layout = engine.Create(source, measurer, default);
            DocxPlacedRelatedStoryLayout[] slices = layout.Pages.SelectMany(p => p.PlacedRelatedStories)
                .Where(s => s.StoryLayout.Story.Id == targetId && s.StoryLayout.Story.IsNormalStoryType).ToArray();
            TestAssert.Equal(2, slices.Length);
            TestAssert.Equal(oldSlices[0].TextLines.Count - 1, slices[0].TextLines.Count);
            TestAssert.Equal(2, slices[^1].TextLines.Count);
            TestAssert.Equal(oldSlices.Sum(s => s.TextLines.Count), slices.Sum(s => s.TextLines.Count));
            TestAssert.True(Math.Abs(oldSlices.Sum(s => s.Height) - slices.Sum(s => s.Height)) < .000001d,
                "Moving one source line preserves the exact cumulative paragraph charge, including final after-spacing.");
            TestAssert.True(Math.Abs(slices[^1].ContentTopOffset - slices[0].Height) < .000001d,
                "Continuation content offsets follow the source-owned slice charges.");
            TestAssert.Equal(oldSlices[0].TopY, slices[0].TopY);
            TestAssert.Equal(oldSlices[0].NoteDestinationTopOffset, slices[0].NoteDestinationTopOffset);
            TestAssert.Equal(0d, slices[^1].NoteDestinationTopOffset);
            TestAssert.True(source.BodyElements.Zip(control!.BodyElements).All(p => ReferenceEquals(p.First, p.Second)),
                "The fragment rule preserves source body and inline-reference ownership.");
            TestAssert.True(source.RelatedStories.Where(s => s.IsNormalStoryType).All(s => s.Paragraphs[0].KeepRules.WidowControl == widow),
                "Layout preserves default and authored widow metadata.");
        }
        DocxDocument kept = control! with { RelatedStories = control!.RelatedStories.Select(s => s.IsNormalStoryType
            ? s with { BodyElements = s.Paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p with
                { KeepRules = p.KeepRules with { KeepLines = true } })).ToArray() } : s).ToArray() };
        TestAssert.True(engine.Create(kept, measurer, default).Pages.SelectMany(p => p.PlacedRelatedStories)
            .Where(s => s.StoryLayout.Story.IsNormalStoryType).All(s => !s.UsesPlainDocumentEndnoteTypography),
            "Kept paragraphs retain their independently qualified fallback.");
    }

    public static void DocxDocumentEndnoteCapacityMovesUnfittingLinesBeforeWidowRules()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        DocxDocument? control = null;
        string? targetId = null;
        int occupiedPage = -1;
        for (int repeats = 20; repeats <= 340 && control is null; repeats++)
        {
            DocxDocument candidate = WithEndnoteWidowControl(ReadWrappedDocumentEndnoteFlowFixture(0, 0, repeats), false);
            DocxRelatedStory[] stories = candidate.RelatedStories.Where(s => s.IsNormalStoryType).ToArray();
            string? firstId = stories[0].Id;
            string? nextId = stories[1].Id;
            candidate = WithEndnoteText(candidate, nextId, "Short note.");
            DocxLayout layout = engine.Create(candidate, measurer, default);
            var first = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Id == firstId)
                .Select(s => (Story: s, Page: index))).ToArray();
            if (first.Length != 1 || !first[0].Story.UsesPlainDocumentEndnoteTypography) continue;
            double remaining = first[0].Story.TopY - first[0].Story.Height - layout.Pages[first[0].Page].MarginBottom;
            double nextHeight = layout.RelatedStories.Single(s => s.Story.Id == nextId && s.Story.IsNormalStoryType).ContentHeight;
            if (remaining <= .001d || remaining >= nextHeight - .001d) continue;
            control = candidate; targetId = nextId; occupiedPage = first[0].Page;
        }
        TestAssert.True(control is not null, "A bounded scan finds positive occupied-frame space smaller than one source line.");
        foreach (bool? widow in new bool?[] { null, true, false })
        foreach (string text in new[] { "Short note.", string.Join(' ', Enumerable.Repeat("Short plain note content.", 6)) })
        {
            DocxDocument source = WithEndnoteWidowControl(WithEndnoteText(control!, targetId, text), widow);
            DocxLayout layout = engine.Create(source, measurer, default);
            DocxRelatedStoryLayout story = layout.RelatedStories.Single(s => s.Story.Id == targetId && s.Story.IsNormalStoryType);
            var placed = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(s => ReferenceEquals(s.StoryLayout.Story, story.Story)).Select(s => (Story: s, Page: index))).ToArray();
            TestAssert.Equal(occupiedPage + 1, placed[0].Page);
            TestAssert.True(placed.All(p => p.Story.UsesPlainDocumentEndnoteTypography), "The regression exercises qualified plain note flow.");
            TestAssert.Equal(story.TextLines.Count, placed.Sum(p => p.Story.TextLines.Count));
            TestAssert.True(Math.Abs(story.ContentHeight - placed.Sum(p => p.Story.Height)) < .000001d,
                "An unfitting line moves intact instead of being clamped into an occupied page's residual clip.");
            TestAssert.Equal(0d, placed[0].Story.ContentTopOffset);
            TestAssert.True(source.BodyElements.Zip(control!.BodyElements).All(p => ReferenceEquals(p.First, p.Second)),
                "Capacity admission preserves body and inline-reference source ownership.");
            TestAssert.True(source.RelatedStories.Where(s => s.IsNormalStoryType).All(s => s.Paragraphs[0].KeepRules.WidowControl == widow),
                "Physical capacity does not rewrite authored widow rules.");
        }
    }

    public static void DocxContinuedEndnoteFooterFramesKeepSelectedStoriesClearAndPreserveSourceLines()
    {
        DocxDocument plain = ReadWrappedDocumentEndnoteFlowFixture(0, 0, 80);
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
        var expectedBody = baseline.Pages.SelectMany(p => p.Items.OfType<DocxTextLineLayout>())
            .Select(l => (l.Text, l.X, l.BaselineY, l.SourceParagraph, l.SourceLineIndex)).ToArray();
        var expectedNotes = baseline.Pages.SelectMany(p => p.PlacedRelatedStories)
            .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote)
            .GroupBy(s => s.StoryLayout.Story.Id).ToDictionary(g => g.Key!, g => g.SelectMany(s => s.StoryLayout.TextLines)
                .Select(l => (l.Text, l.SourceParagraph, l.SourceLineIndex)).ToArray());
        foreach (bool alternating in new[] { false, true })
        {
            DocxDocument source = Compose(10, alternating);
            DocxLayout actual = engine.Create(source, measurer, default);
            TestAssert.True(actual.Pages.Count >= baseline.Pages.Count, "Reducing the continuation frame does not remove pages.");
            TestAssert.True(expectedBody.SequenceEqual(actual.Pages.SelectMany(p => p.Items.OfType<DocxTextLineLayout>())
                .Select(l => (l.Text, l.X, l.BaselineY, l.SourceParagraph, l.SourceLineIndex))), "Body owners and geometry remain exact.");
            foreach (var group in actual.Pages.SelectMany(p => p.PlacedRelatedStories)
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote)
                .GroupBy(s => s.StoryLayout.Story.Id))
            {
                TestAssert.True(expectedNotes[group.Key!].SequenceEqual(group.SelectMany(s => s.StoryLayout.TextLines)
                    .Select(l => (l.Text, l.SourceParagraph, l.SourceLineIndex))), "Every source line remains ordered and appears exactly once.");
                TestAssert.Equal(1, group.Count(s => s.ContentTopOffset == 0d));
            }
            for (int i = 0; i < actual.Pages.Count; i++)
            {
                DocxLayoutPage page = actual.Pages[i];
                double bottom = Math.Max(page.MarginBottom, actual.FooterContentTopByPage[i]);
                bool fittingVariant = alternating && (i + 1) % 2 == 0;
                TestAssert.True(fittingVariant ? bottom == page.MarginBottom : bottom > page.MarginBottom,
                    "The selected default/even footer owns its own frame edge.");
                foreach (var story in page.PlacedRelatedStories.Where(s => s.StoryLayout.Story.IsNormalStoryType))
                    TestAssert.True(story.TopY - story.Height >= bottom - .001d, "Note clips stay above the selected footer edge.");
            }
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, actual, out _, out _, default),
                "A usable plain frame shares the existing note typography policy.");
        }
        DocxLayout fitting = engine.Create(Compose(1, false), measurer, default);
        TestAssert.True(baseline.Pages.SelectMany(p => p.PlacedRelatedStories).Select(s => (s.TopY, s.Height, s.ContentTopOffset))
            .SequenceEqual(fitting.Pages.SelectMany(p => p.PlacedRelatedStories).Select(s => (s.TopY, s.Height, s.ContentTopOffset))),
            "Fitting static edges leave the original note placement exact.");
    }

    public static void DocxFittingMixedEndnoteRetryPreservesTypedOwnersAndComplementaryGeometry()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool custom in new[] { false, true })
        {
            DocxDocument source = ReadMixedNoteRetryFixture(custom, 32);
            DocxLayout before = engine.Create(source, measurer, default);
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, before, out _, out _, default, DocxRelatedStoryKind.Footnote),
                "The complementary footnote retains its established empty-static frame fallback.");
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, before, out var adjusted, out var owners, default),
                "An admitted endnote frame uses the existing advance correction in a complete mixed document.");
            TestAssert.Equal(1, owners.Count);
            TestAssert.True(source.BodyElements.Zip(adjusted.BodyElements).All(p => ReferenceEquals(p.First, p.Second)),
                "The correction preserves both typed body references, including their equal IDs.");
            DocxRelatedStory foot = source.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Footnote);
            DocxRelatedStory end = source.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Endnote);
            TestAssert.Equal(foot.Id!, end.Id!);
            DocxRelatedStory adjustedFoot = adjusted.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Footnote);
            TestAssert.True(ReferenceEquals(foot.Paragraphs[0], adjustedFoot.Paragraphs[0]),
                "The complementary paragraph and every run remain the exact authored owners.");
            DocxLayout after = engine.Create(adjusted, measurer, default);
            TestAssert.True(DocxNoteKerning.HasStableOwners(after, owners, default), "Both kinds retain complete canonical slices after one retry.");
            DocxLayout lostKind = after with
            {
                RelatedStories = after.RelatedStories.Where(s => s.Story.Kind != DocxRelatedStoryKind.Footnote).ToArray(),
                Pages = after.Pages.Select(p => p with { PlacedRelatedStories = p.PlacedRelatedStories
                    .Where(s => s.StoryLayout.Story.Kind != DocxRelatedStoryKind.Footnote).ToArray() }).ToArray(),
            };
            TestAssert.True(!DocxNoteKerning.HasStableOwners(lostKind, owners, default),
                "The published owner map retains the original complementary sources even if an entire kind disappears.");
            var previous = before.Pages.SelectMany((p, index) => p.PlacedRelatedStories.Where(s => s.StoryLayout.Story.IsNormalStoryType &&
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote).Select(s => (Page: index, s.TopY, s.Height, s.ContentTopOffset))).ToArray();
            var current = after.Pages.SelectMany((p, index) => p.PlacedRelatedStories.Where(s => s.StoryLayout.Story.IsNormalStoryType &&
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote).Select(s => (Page: index, s.TopY, s.Height, s.ContentTopOffset))).ToArray();
            TestAssert.True(previous.SequenceEqual(current), "Endnote advances preserve the complementary note's page and physical box.");
            var endSlices = after.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => s.StoryLayout.Story.IsNormalStoryType &&
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote).ToArray();
            var previousPurposes = before.Pages.SelectMany(p => p.PlacedRelatedStories).Where(s => s.StoryLayout.Story.IsNormalStoryType &&
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote).Select(s => s.UsesCompleteDocumentEndnoteFlow);
            TestAssert.True(endSlices.All(s => s.UsesPlainDocumentEndnoteTypography) &&
                previousPurposes.SequenceEqual(endSlices.Select(s => s.UsesCompleteDocumentEndnoteFlow)),
                "Typography admission preserves the mixed document's existing pagination purposes.");
        }
    }

    public static void DocxMixedEndnoteRetryRequiresCompleteUniquePlainComplementaryOwners()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        DocxDocument source = ReadMixedNoteRetryFixture(true, 32);
        DocxLayout layout = engine.Create(source, measurer, default);
        var endPages = layout.Pages.SelectMany((p, page) => p.PlacedRelatedStories.Where(s =>
            s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && s.StoryLayout.Story.IsNormalStoryType).Select(s => page)).ToArray();
        TestAssert.True(endPages.Length > 0, "The fallback control has a real placed endnote owner.");
        var owners = new Dictionary<string, DocxNoteKerning.PageRange> { ["37"] = new(endPages[0], endPages[^1], endPages.Length) };
        var missing = layout with { Pages = layout.Pages.Select(p => p with
        {
            PlacedRelatedStories = p.PlacedRelatedStories.Where(s => s.StoryLayout.Story.Kind != DocxRelatedStoryKind.Footnote ||
                !s.StoryLayout.Story.IsNormalStoryType).ToArray(),
        }).ToArray() };
        TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, missing, out var unchanged, out _, default) &&
            ReferenceEquals(source, unchanged) && !DocxNoteKerning.HasStableOwners(missing, owners, default),
            "Missing complementary slices reject publication before and after the retry.");
        DocxRelatedStory foot = source.RelatedStories.Single(s => s.Kind == DocxRelatedStoryKind.Footnote && s.IsNormalStoryType);
        foreach (DocxDocument invalid in new[]
        {
            source with { RelatedStories = source.RelatedStories.Concat([foot with { }]).ToArray() },
            source with { RelatedStories = source.RelatedStories.Select(s => ReferenceEquals(s, foot) ? s with
                { BodyElements = [new DocxParagraphElement(foot.Paragraphs[0] with { Alignment = DocxTextAlignment.Center })] } : s).ToArray() },
            source with { BodyElements = source.BodyElements.Skip(1).ToArray() },
            ReadMixedNoteRetryFixture(true, 140),
            ReadMixedNoteRetryFixture(false, 140),
        })
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(invalid, engine.Create(invalid, measurer, default), out unchanged, out _, default) &&
                ReferenceEquals(invalid, unchanged), "Duplicate targets, rich complementary paragraphs, missing references and unadmitted continuation frames retain the fallback.");
    }

    public static void DocxContinuedMixedEndnoteHeadClearsFootnoteAndPreservesTypedOwners()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        foreach (bool custom in new[] { false, true })
        foreach (bool header in new[] { false, true })
        {
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 140), header, true, header ? 1 : 6);
            DocxLayout before = engine.Create(source, measurer, default);
            var foots = before.Pages.SelectMany((p, index) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                .Select(s => (Page: index, Story: s))).ToArray();
            var ends = before.Pages.SelectMany((p, index) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote)
                .Select(s => (Page: index, Story: s))).ToArray();
            TestAssert.Equal(1, foots.Length);
            TestAssert.True(ends.Length > 1 && ends[0].Page == foots[0].Page &&
                ends[0].Story.TopY - ends[0].Story.Height >= foots[0].Story.TopY,
                "A real partial endnote head sits above the footnote on the occupied body page.");
            DocxRelatedStoryLayout canonical = before.RelatedStories.Single(s => s.Story.IsNormalStoryType &&
                s.Story.Kind == DocxRelatedStoryKind.Endnote);
            TestAssert.True(ends[0].Story.StoryLayout.TextLines.Count >= 2 &&
                ends[0].Story.StoryLayout.TextLines.Count < canonical.TextLines.Count,
                "The admitted slice is a multi-line head of a continued source, not a fitting complete block.");
            TestAssert.Equal(string.Concat(canonical.TextLines.Select(l => l.Text)),
                string.Concat(ends.SelectMany(s => s.Story.StoryLayout.TextLines).Select(l => l.Text)));
            TestAssert.True(ends.All(s => !s.Story.UsesCompleteDocumentEndnoteFlow),
                "Head capacity does not opt mixed paragraphs into the complete-flow boundary policy.");
            TestAssert.Equal(foots[0].Story.StoryLayout.Story.Id!, canonical.Story.Id!);
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, before, out _, out _, default, DocxRelatedStoryKind.Footnote),
                "Physical head capacity preserves the complementary footnote typography fallback.");
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, before, out var adjusted, out var owners, default),
                "The existing endnote retry can resolve complete typed owners across the partial head and tail.");
            DocxLayout after = engine.Create(adjusted, measurer, default);
            TestAssert.True(DocxNoteKerning.HasStableOwners(after, owners, default),
                "The rewrapped head and continuations retain both complete canonical owners.");
            var currentFoot = after.Pages[foots[0].Page].PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote);
            TestAssert.True(currentFoot.TopY == foots[0].Story.TopY && currentFoot.Height == foots[0].Story.Height &&
                currentFoot.ContentTopOffset == foots[0].Story.ContentTopOffset &&
                ReferenceEquals(currentFoot.StoryLayout.Story.Paragraphs[0], foots[0].Story.StoryLayout.Story.Paragraphs[0]),
                "The endnote retry preserves the complementary physical box and original source paragraph.");
            TestAssert.True(source.BodyElements.Zip(adjusted.BodyElements).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "The head correction preserves both typed body references.");
            foreach (DocxLayoutPage page in after.Pages)
            {
                var normalFoots = page.PlacedRelatedStories.Where(s => s.StoryLayout.Story.IsNormalStoryType &&
                    s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote).ToArray();
                if (normalFoots.Length == 0) continue;
                TestAssert.True(page.PlacedRelatedStories.Where(s => s.StoryLayout.Story.IsNormalStoryType &&
                    s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote)
                    .All(s => s.TopY - s.Height >= normalFoots.Max(f => f.TopY)),
                    "The rewrapped endnote head cannot enter the complementary footnote box.");
            }
        }
    }

    public static void DocxContinuedMixedEndnoteHeadRetainsUnqualifiedSourceFallbacks()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        DocxDocument plain = ReadMixedNoteRetryFixture(true, 140);
        DocxRelatedStory end = plain.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Endnote);
        DocxParagraph paragraph = end.Paragraphs[0];
        var sources = new List<DocxDocument> { plain };
        foreach (DocxParagraph invalid in new[]
        {
            paragraph with { KeepRules = paragraph.KeepRules with { KeepNext = true } },
            paragraph with { KeepRules = paragraph.KeepRules with { KeepLines = true } },
            paragraph with { Spacing = paragraph.Spacing with { ContextualSpacing = true } },
            paragraph with { Alignment = DocxTextAlignment.Center },
        })
            sources.Add(WithMixedNoteStaticFrame(plain with { RelatedStories = plain.RelatedStories.Select(s =>
                ReferenceEquals(s, end) ? s with { BodyElements = [new DocxParagraphElement(invalid)] } : s).ToArray() }, false, true, 6));
        foreach (DocxDocument source in sources)
        {
            DocxLayout layout = engine.Create(source, measurer, default);
            int footPage = layout.Pages.Select((p, index) => (p, index)).First(x => x.p.PlacedRelatedStories.Any(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)).index;
            var ends = layout.Pages.SelectMany((p, index) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote)
                .Select(s => (Page: index, Story: s))).ToArray();
            TestAssert.True(ends[0].Page > footPage && ends.All(s => !s.Story.UsesPlainDocumentEndnoteTypography),
                "Absent static qualification and rich complementary flow retain the existing head deferral.");
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, layout, out _, out _, default),
                "Unqualified continued heads do not enable the typography retry.");
        }
    }

    public static void DocxFittingMixedEndnoteKeepsUnownedBearingAndFrameFallbacks()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        foreach (bool custom in new[] { false, true })
        foreach (int footRepeats in new[] { 0, 16 })
        foreach (double ascent in new[] { 0d, -1d, 12d, 10.8d })
        {
            bool footer = ascent != 10.8d;
            var measurer = new NoteContentMetricMeasurer(ascent, 1.15d, 1.220703125d, true);
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32, footRepeats),
                false, footer, 6);
            DocxLayout layout = engine.Create(source, measurer, default);
            DocxLayoutPage page = layout.Pages.Single(p => p.PlacedRelatedStories.Any(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote));
            TestAssert.True(page.Items.Count > 0, "The fallback witnesses a fitting occupied head.");
            DocxPlacedRelatedStoryLayout end = page.PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote);
            DocxPlacedRelatedStoryLayout separator = page.PlacedRelatedStories.Single(s =>
                s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote &&
                s.StoryLayout.Story.Type == DocxRelatedStoryType.Separator);
            double actual = separator.SeparatorY!.Value - end.TextLines[0].BaselineY;
            TestAssert.True(Math.Abs(actual - 17.15d) < .02d,
                $"Unavailable/unowned first bearings and absent static frames retain the qualified gap: {actual}, ascent={ascent}, custom={custom}.");
        }
    }

    public static void DocxDeferredMixedEndnoteKeepsOriginalPlainSourceAdmissionAndOwners()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        int deferred = 0;
        foreach (bool custom in new[] { false, true })
        foreach (int footRepeats in new[] { 96, 100, 104 })
        {
            DocxDocument source = WithMixedNoteStaticFrame(ReadMixedNoteRetryFixture(custom, 32, footRepeats), false, true, 6);
            DocxLayout before = engine.Create(source, measurer, default);
            var ends = before.Pages.SelectMany((p, index) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote)
                .Select(s => (Page: index, Story: s))).ToArray();
            if (before.Pages[ends[0].Page].Items.Count != 0) continue;
            deferred++;
            TestAssert.True(ends.All(s => s.Story.UsesPlainDocumentEndnoteTypography && !s.Story.UsesCompleteDocumentEndnoteFlow),
                "A deferred empty-page head retains plain source typography and the original mixed boundary policy.");
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, before, out _, out _, default, DocxRelatedStoryKind.Footnote),
                "Carried endnote admission preserves the independent mixed footnote typography fallback.");
            TestAssert.True(DocxNoteKerning.TryCreateDocument(source, before, out var adjusted, out var owners, default),
                "The deferred source remains eligible for the one bounded endnote retry.");
            DocxLayout after = engine.Create(adjusted, measurer, default);
            TestAssert.True(DocxNoteKerning.HasStableOwners(after, owners, default),
                "The deferred head and every complementary continuation retain complete typed canonical owners.");
            var previousFeet = before.Pages.SelectMany((p, index) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                .Select(s => (Page: index, s.TopY, s.Height, s.ContentTopOffset, s.SeparatorY, s.SeparatorWidth, s.SeparatorThickness))).ToArray();
            var currentFeet = after.Pages.SelectMany((p, index) => p.PlacedRelatedStories
                .Where(s => s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                .Select(s => (Page: index, s.TopY, s.Height, s.ContentTopOffset, s.SeparatorY, s.SeparatorWidth, s.SeparatorThickness))).ToArray();
            TestAssert.True(previousFeet.SequenceEqual(currentFeet),
                "The endnote retry preserves every complementary footnote clip, continuation and separator rule.");
            TestAssert.True(source.BodyElements.Zip(adjusted.BodyElements).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Both typed body references, including their equal numeric IDs, remain original source owners.");
        }
        TestAssert.True(deferred > 0, "The boundary controls include an actual endnote head deferred to an empty page.");
    }

    public static void DocxDeferredMixedEndnoteRetainsUnqualifiedSourceAndStaticFallbacks()
    {
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new NoteContentMetricMeasurer();
        DocxDocument plain = ReadMixedNoteRetryFixture(true, 32, 100);
        DocxRelatedStory foot = plain.RelatedStories.Single(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Footnote);
        DocxParagraph paragraph = foot.Paragraphs[0];
        foreach (DocxParagraph invalid in new[]
        {
            paragraph with { KeepRules = paragraph.KeepRules with { KeepNext = true } },
            paragraph with { Spacing = paragraph.Spacing with { ContextualSpacing = true } },
            paragraph with { Alignment = DocxTextAlignment.Center },
        })
        {
            DocxDocument source = WithMixedNoteStaticFrame(plain with { RelatedStories = plain.RelatedStories.Select(s =>
                ReferenceEquals(s, foot) ? s with { BodyElements = [new DocxParagraphElement(invalid)] } : s).ToArray() }, false, true, 6);
            TestAssert.True(!DocxNoteKerning.TryCreateDocument(source, engine.Create(source, measurer, default), out _, out _, default),
                "Carried source admission requires the existing plain complementary paragraph semantics.");
        }
        DocxDocument scaled = WithMixedNoteStaticFrame(plain, false, true, 6);
        DocxLayout scaledLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, .5d)
            .Create(scaled, measurer, default, measurer);
        TestAssert.True(!DocxNoteKerning.TryCreateDocument(scaled, scaledLayout, out _, out _, default),
            "A page turn does not admit an originally scaled source frame.");
    }

    public static void DocxFittingDocumentEndnotesStayAboveFootnoteArea()
    {
        DocxDocument document = ReadNoteNavigationFixture("note-nav-mixed-notes");
        var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxLayoutSnapshot fitting = renderer.InspectLayout(document);
        TestAssert.Equal(1, fitting.Pages.Count);
        var footnotes = fitting.Pages[0].PlacedRelatedStories.Where(story => story.Kind == "Footnote").ToArray();
        var endnotes = fitting.Pages[0].PlacedRelatedStories.Where(story => story.Kind == "Endnote").ToArray();
        TestAssert.True(endnotes.Length > 0 && endnotes.Min(story => story.TopY - story.Height) >= footnotes.Max(story => story.TopY),
            "Fitting document endnotes belong below the body and above the complete footnote area.");
        DocxRelatedStory endnote = document.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Endnote && story.IsNormalStoryType);
        DocxParagraph paragraph = endnote.Paragraphs[0];
        DocxRelatedStory longEndnote = endnote with { BodyElements = [new DocxParagraphElement(paragraph with
            { Runs = [paragraph.Runs[0], paragraph.Runs[1] with { Text = string.Join(' ', Enumerable.Repeat("Long endnote content", 500)) }] })] };
        DocxDocument overflowing = document with { RelatedStories = document.RelatedStories.Select(story => ReferenceEquals(story, endnote) ? longEndnote : story).ToArray() };
        DocxLayoutSnapshot fallback = renderer.InspectLayout(overflowing);
        TestAssert.True(fallback.Pages.Count > 1 && !fallback.Pages[0].PlacedRelatedStories.Any(story => story.Kind == "Endnote"),
            "Endnotes that exceed the free gap retain continuation placement and cannot overlap footnotes.");
    }

    public static void DocxEndnoteStartTopRecoversDesignBodyBottom()
    {
        // RV06 anchor probe (edge-endanchor-5, Word 16.0): scaled body bottom
        // 689.77 with frame top 720, first inset 11.28 and print scale 0.75874
        // recovers the design body end 683.75; scale 1.0 and degenerate frames
        // keep legacy behavior.
        TestAssert.True(
            Math.Abs(DocxLayoutEngine.ResolveDesignBodyBottomForEndnoteStart(689.77d, 720d, 11.28d, 0.75874d) - 683.75d) < 0.02d,
            "Design body end should recover 683.75.");
        TestAssert.Equal(100d, DocxLayoutEngine.ResolveDesignBodyBottomForEndnoteStart(100d, 720d, 11.28d, 1d));
        TestAssert.Equal(700d, DocxLayoutEngine.ResolveDesignBodyBottomForEndnoteStart(700d, 720d, null, 0.75d));
    }

    public static void DocxLayoutPlacesSectEndEndnoteOnOwningSectionEndPage()
    {
        DocxParagraph firstSectionParagraph = DocxTests.CreateDocxLayoutParagraph("first section endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "21",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 6)
            ]
        };
        DocxParagraph secondSectionParagraph = DocxTests.CreateDocxLayoutParagraph("second section body", 10d, 12d);
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Endnote body", 10d, 12d);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "21",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        DocxPageSettings firstSectionSettings = DocxPageSettings.Empty with
        {
            EndnoteReferenceSettings = DocxNoteReferenceSettings.Empty with { PositionValue = "sectEnd" }
        };
        var sectionBreak = new DocxSectionBreakElement(firstSectionSettings, DocxSectionBreakType.NextPage, null, null, null, []);
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
                new DocxParagraphElement(firstSectionParagraph),
                sectionBreak,
                new DocxParagraphElement(secondSectionParagraph)
            ],
            [firstSectionParagraph, secondSectionParagraph],
            [])
        {
            RelatedStories = [endnoteStory]
        };

        document = document with { Settings = document.Settings with { EndnoteReferenceSettings =
            document.Settings.EndnoteReferenceSettings with { PositionValue = "sectEnd" } } };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        DocxLayoutPageSnapshot endnotePage = snapshot.Pages.Single(page => page.PlacedEndnoteStoryCount == 1);

        TestAssert.True(snapshot.Pages.Count >= 2, "The next-page section break should separate the two sections.");
        TestAssert.Equal("sectEnd", snapshot.Pages[0].SectionEndnotePositionValue ?? string.Empty);
        TestAssert.Equal(1, snapshot.Pages[0].PlacedEndnoteStoryCount);
        TestAssert.Equal(0, snapshot.Pages[^1].PlacedEndnoteStoryCount);
        TestAssert.True(endnotePage.PlacedRelatedStories.Any(story => story.Kind == "Endnote" && story.SourceBlockIndex == 0), "A sectEnd endnote should be owned by the source section end page, not rewritten as a document-end story.");
    }

    public static void DocxLayoutWrapsDocumentEndEndnoteUsingFinalSectionPageWidth()
    {
        DocxPageSettings wideFirstSection = new(
            "12240",
            "12000",
            null,
            "1440",
            "1440",
            "1440",
            "1440",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        DocxPageSettings narrowFinalSection = new(
            "4400",
            "6000",
            null,
            "200",
            "200",
            "200",
            "200",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Narrow document-end endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "25",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 7)
            ]
        };
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("alpha beta gamma delta epsilon zeta eta theta", 10d, 12d);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "25",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            narrowFinalSection,
            [],
            [],
            [],
            [
                new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Wide first section", 10d, 12d)),
                new DocxSectionBreakElement(wideFirstSection, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(anchor)
            ],
            [],
            [])
        {
            RelatedStories = [endnoteStory]
        };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout placedStory = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.StoryLayout.Story.Id == "25");
        DocxTextLineLayout[] lines = placedStory.TextLines.ToArray();

        TestAssert.Equal(-1, placedStory.SourceBlockIndex);
        TestAssert.True(Math.Abs(placedStory.Width - 200d) < 0.001d, "The document-end endnote should inherit the final section body width.");
        TestAssert.True(lines.Length >= 2, "Document-end endnotes should wrap using the target page width, not the first section width.");
        TestAssert.True(lines.All(line => line.Width <= placedStory.Width + 0.001d), "Document-end endnote lines should not exceed the final section body width.");
    }

    public static void DocxLayoutWrapsSectEndEndnoteUsingOwningSectionPageWidth()
    {
        DocxPageSettings wideFirstSection = new(
            "12240",
            "12000",
            null,
            "1440",
            "1440",
            "1440",
            "1440",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null);
        DocxPageSettings narrowFinalSection = new(
            "4400",
            "6000",
            null,
            "200",
            "200",
            "200",
            "200",
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null)
        {
            EndnoteReferenceSettings = DocxNoteReferenceSettings.Empty with { PositionValue = "sectEnd" }
        };
        DocxParagraph anchor = DocxTests.CreateDocxLayoutParagraph("Narrow section-end endnote marker", 10d, 12d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "26",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 7)
            ]
        };
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("alpha beta gamma delta epsilon zeta eta theta", 10d, 12d);
        var endnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "26",
            [new DocxParagraphElement(endnoteParagraph)],
            [],
            [], null);
        var document = new DocxDocument(
            612d,
            792d,
            72d,
            72d,
            72d,
            72d,
            narrowFinalSection,
            [],
            [],
            [],
            [
                new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Wide first section", 10d, 12d)),
                new DocxSectionBreakElement(wideFirstSection, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(anchor)
            ],
            [],
            [])
        {
            RelatedStories = [endnoteStory]
        };

        document = document with { Settings = document.Settings with { EndnoteReferenceSettings =
            document.Settings.EndnoteReferenceSettings with { PositionValue = "sectEnd" } } };

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxPlacedRelatedStoryLayout placedStory = layout.Pages
            .SelectMany(page => page.PlacedRelatedStories)
            .Single(story => story.StoryLayout.Story.Id == "26");
        DocxTextLineLayout[] lines = placedStory.TextLines.ToArray();

        TestAssert.Equal(2, placedStory.SourceBlockIndex);
        TestAssert.True(Math.Abs(placedStory.Width - 200d) < 0.001d, "The section-end endnote should inherit the owning section body width.");
        TestAssert.True(lines.Length >= 2, "Section-end endnotes should wrap using the target section width, not the first section width.");
        TestAssert.True(lines.All(line => line.Width <= placedStory.Width + 0.001d), "Section-end endnote lines should not exceed the owning section body width.");
    }

    public static void DocxLayoutKeepsSectEndEndnoteOverflowPageBeforeFollowingSection()
    {
        DocxParagraph firstSectionParagraph = DocxTests.CreateDocxLayoutParagraph("first section endnote marker", 10d, 18d) with
        {
            InlineReferences =
            [
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "22",
                    CustomMarkFollowsValue: null,
                    DisplayText: "1",
                    SourceRunIndex: 0,
                    RunChildIndex: 1,
                    TextOffsetInRun: 6),
                new DocxInlineReference(
                    DocxRelatedStoryKind.Endnote,
                    "23",
                    CustomMarkFollowsValue: null,
                    DisplayText: "2",
                    SourceRunIndex: 0,
                    RunChildIndex: 2,
                    TextOffsetInRun: 12)
            ]
        };
        DocxParagraph secondSectionParagraph = DocxTests.CreateDocxLayoutParagraph("second section body", 10d, 12d);
        DocxParagraph endnoteParagraph = DocxTests.CreateDocxLayoutParagraph("Endnote overflow body", 10d, 18d);
        var firstEndnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "22",
            Enumerable.Range(0, 4).Select(_ => new DocxParagraphElement(endnoteParagraph)).Cast<DocxBodyElement>().ToArray(),
            [],
            [], null);
        var secondEndnoteStory = new DocxRelatedStory(
            DocxRelatedStoryKind.Endnote,
            "/word/endnotes.xml",
            "23",
            Enumerable.Range(0, 4).Select(_ => new DocxParagraphElement(endnoteParagraph)).Cast<DocxBodyElement>().ToArray(),
            [],
            [], null);
        DocxPageSettings firstSectionSettings = DocxPageSettings.Empty with
        {
            EndnoteReferenceSettings = DocxNoteReferenceSettings.Empty with { PositionValue = "sectEnd" }
        };
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
                new DocxParagraphElement(firstSectionParagraph),
                new DocxSectionBreakElement(firstSectionSettings, DocxSectionBreakType.NextPage, null, null, null, []),
                new DocxParagraphElement(secondSectionParagraph)
            ],
            [firstSectionParagraph, secondSectionParagraph],
            [])
        {
            RelatedStories = [firstEndnoteStory, secondEndnoteStory]
        };

        document = document with { Settings = document.Settings with { EndnoteReferenceSettings =
            document.Settings.EndnoteReferenceSettings with { PositionValue = "sectEnd" } } };

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None));
        int secondSectionPageIndex = snapshot.Pages
            .Select((page, pageIndex) => (page, pageIndex))
            .First(item => item.page.Items.Any(pageItem => pageItem.SourceBlockIndex == 2))
            .pageIndex;
        int overflowEndnotePageIndex = snapshot.Pages
            .Select((page, pageIndex) => (page, pageIndex))
            .Where(item => item.page.PlacedEndnoteStoryCount == 1)
            .Select(item => item.pageIndex)
            .Last();

        TestAssert.True(snapshot.Pages.Count >= 3, "Multiple sectEnd endnotes that cannot all fit on the section end page should create a section-owned continuation page.");
        TestAssert.True(overflowEndnotePageIndex > 0 && overflowEndnotePageIndex < secondSectionPageIndex, "Section-end endnote overflow must be inserted before the following section instead of falling back to document-end placement.");
        TestAssert.Equal("sectEnd", snapshot.Pages[overflowEndnotePageIndex].SectionEndnotePositionValue ?? string.Empty);
        TestAssert.True(snapshot.Pages[overflowEndnotePageIndex].PlacedRelatedStories.Any(story => story.Kind == "Endnote" && story.SourceBlockIndex == 0), "The inserted section-end continuation page should retain marker-owned endnote provenance.");
        TestAssert.Equal(0, snapshot.Pages[secondSectionPageIndex].PlacedEndnoteStoryCount);
    }

    public static void DocxSplitEndnotesKeepEachLineOnExactlyOnePage()
    {
        // RV06 endlong probe (Word 16.0): a 60-paragraph section-end endnote splits
        // 23/25/12 with full-width continuation rules, one text op per line, while the
        // renderer triplicates every line across 3 slice pages (180 ops for 60 lines)
        // because slices shift without filtering and only clip at emission.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t xml:space="preserve">Body with note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:endnotePr><w:pos w:val="sectEnd"/></w:endnotePr></w:sectPr></w:body></w:document>""",
            ["word/endnotes.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:t>Note body line 0 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 1 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 2 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 3 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 4 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 5 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 6 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 7 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 8 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 9 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 10 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 11 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 12 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 13 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 14 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 15 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 16 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 17 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 18 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 19 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 20 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 21 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 22 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 23 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 24 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 25 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 26 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 27 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 28 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 29 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 30 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 31 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 32 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 33 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 34 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 35 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 36 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 37 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 38 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 39 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 40 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 41 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 42 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 43 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 44 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 45 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 46 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 47 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 48 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 49 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 50 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 51 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 52 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 53 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 54 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 55 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 56 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 57 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 58 words here</w:t></w:r></w:p><w:p><w:r><w:t>Note body line 59 words here</w:t></w:r></w:p></w:endnote></w:endnotes>"""
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.True(layout.Pages.Count > 1, "Overflowing endnotes must paginate.");
        var lineKeys = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Where(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type != DocxRelatedStoryType.Separator && story.StoryLayout.Story.Type != DocxRelatedStoryType.ContinuationSeparator).SelectMany(story => story.TextLines).Select(line => (line.SourceBlockIndex, line.SourceParagraphIndex, line.SourceLineIndex)).ToArray();
        TestAssert.Equal(lineKeys.Distinct().Count(), lineKeys.Length);
    }
}
