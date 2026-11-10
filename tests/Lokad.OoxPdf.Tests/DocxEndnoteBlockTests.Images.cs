using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxEndnoteBlockTests
{
    public static void DocxPictureOnlyEndnoteRowsUseCompleteSourceFrames()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int height in new[] { 36, 560, 600, 900 })
        {
            var layout = PictureLayout(height, mode);
            var owners = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                .Where(story => story.StoryLayout.Story.Id == "37")
                .Select(story => (Page: index, Story: story))).ToArray();
            var prefix = owners.First(owner => owner.Story.TextLines.Any(line => line.Text.Contains("Public destination Alpha.", StringComparison.Ordinal)));
            var owner = owners.Single(item => item.Story.TableRows.Count != 0);
            var row = owner.Story.TableRows.Single();
            var picture = row.Cells.Single().InlineImages.Single();
            TestAssert.Equal((double)height, picture.Height);
            TestAssert.True(DocxTableCellClipGeometry.HasCompleteMeasuredRows([row], default),
                "The separate picture placement and measured carrier both survive their source cell clip.");
            if (height >= 600)
                TestAssert.True(owner.Page > prefix.Page, "A picture row defers from occupied space using the existing complete-block admission.");
            if (height < 900)
                TestAssert.True(picture.Y >= owner.Story.TopY - owner.Story.Height - .000001d &&
                    picture.Y + picture.Height <= owner.Story.TopY + .000001d,
                    "Every fitting picture is wholly inside its final note frame.");
            else TestAssert.True(owner.Story.Height > 0d && owner.Story.Height < row.Height &&
                owner.Story.Height <= layout.Pages[owner.Page].Height - layout.Pages[owner.Page].MarginTop - layout.Pages[owner.Page].MarginBottom,
                "A genuinely unfit picture retains the existing bounded clip, without fragmentation or rescaling.");
        }
    }

    public static void DocxPictureOnlyNoteRowsRejectIncompleteSourceViewports()
    {
        var layout = PictureLayout(36, OoxPdfDocxMarkupMode.Final);
        var row = layout.RelatedStories.Single(story => story.Story.Id == "37").TableRows.Single();
        var cell = row.Cells.Single();
        var image = cell.InlineImages.Single();
        TestAssert.True(DocxTableCellClipGeometry.HasCompleteMeasuredRows([row], default), "The control owns a complete picture frame.");
        foreach (var partial in new[] {
            image with { Y = cell.Y - 1d },
            image with { X = cell.X - 1d },
            image with { Width = double.NaN },
            image with { Height = 0d },
            image with { ClipBounds = new PdfRectangle(image.X, image.Y, image.Width, image.Height / 2d) } })
        {
            var clipped = row with { Cells = [cell with { InlineImages = [partial] }] };
            TestAssert.True(!DocxTableCellClipGeometry.HasCompleteMeasuredRows([clipped], default),
                "A surviving text carrier does not substitute for a missing, unmeasured or source-clipped picture.");
        }
        var borrowed = row with { Cells = [cell with { VisualOwnership = DocxTableCellVisualOwnership.VerticalMergeOwner }] };
        TestAssert.True(!DocxTableCellClipGeometry.HasCompleteMeasuredRows([borrowed], default), "Retain explicit cell ownership.");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() =>
            DocxTableCellClipGeometry.HasCompleteMeasuredRows([row], cancellation.Token)).CancellationToken);
    }

    public static void DocxPictureOnlyNoteRowsRetainAuthoredCellClipping()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var layout = PictureLayout(600, mode, exactHeight: 100);
            var canonical = layout.RelatedStories.Single(story => story.Story.Id == "37").TableRows.Single();
            TestAssert.True(!DocxTableCellClipGeometry.HasCompleteMeasuredRows([canonical], default),
                "An authored cell clips its picture independently of note-page admission.");
            var placed = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Where(story => story.StoryLayout.Story.Id == "37").ToArray();
            TestAssert.True(placed.Any(story => story.TableRows.Count != 0), "Preserve the existing authored clipping fallback.");
        }
    }

    public static void DocxPictureNoteStartDoesNotDependOnLaterNoteLength()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var shortLayout = PictureLayout(36, mode);
            var longLayout = PictureLayout(36, mode, laterNoteWords: 800);
            (int Page, DocxPlacedRelatedStoryLayout Story) First(DocxLayout layout) => layout.Pages
                .SelectMany((page, index) => page.PlacedRelatedStories
                    .Where(story => story.StoryLayout.Story.Id == "37")
                    .Select(story => (index, story))).First();
            var before = First(shortLayout);
            var after = First(longLayout);
            TestAssert.Equal(before.Page, after.Page);
            TestAssert.Equal(before.Story.TopY, after.Story.TopY);
            TestAssert.Equal(before.Story.TableRows.Single().Cells.Single().InlineImages.Single().Y,
                after.Story.TableRows.Single().Cells.Single().InlineImages.Single().Y);
            TestAssert.True(before.Story.UsesCompleteDocumentEndnoteFlow && !after.Story.UsesCompleteDocumentEndnoteFlow,
                "Start-boundary ownership does not promote an overflowing group to complete fitting flow.");
            string text = string.Join(" ", longLayout.Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(story => story.StoryLayout.Story.Id == "8").SelectMany(story => story.TextLines).Select(line => line.Text));
            foreach (int index in Enumerable.Range(0, 800))
                TestAssert.True(text.Contains($"LaterNote{index:000}", StringComparison.Ordinal), "Retain every later-note token.");
        }
    }

    public static void DocxPictureNoteFollowingBoundaryDoesNotDependOnLaterNoteLength()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var before = PictureLayout(36, mode);
            var after = PictureLayout(36, mode, laterNoteWords: 800);
            foreach (string id in new[] { "4", "73" })
            {
                var a = before.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                    .Where(story => story.StoryLayout.Story.Id == id).Select(story => (Page: index, Story: story))).Single();
                var b = after.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                    .Where(story => story.StoryLayout.Story.Id == id).Select(story => (Page: index, Story: story))).Single();
                TestAssert.Equal(a.Page, b.Page);
                TestAssert.Equal(a.Story.TopY, b.Story.TopY);
                TestAssert.Equal(a.Story.TextLines[0].BaselineY, b.Story.TextLines[0].BaselineY);
                TestAssert.True(!b.Story.UsesCompleteDocumentEndnoteFlow && !b.Story.UsesPlainDocumentEndnoteTypography,
                    "The local boundary does not promote the whole overflowing group or its typography.");
            }
        }
    }

    public static void DocxPictureNoteFollowingHeadUsesRemainingMeasuredSpace()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var layout = PictureLayout(570, mode);
            int pageIndex = Enumerable.Range(0, layout.Pages.Count).Single(index =>
                layout.Pages[index].PlacedRelatedStories.Any(story => story.StoryLayout.Story.Id == "37" && story.TableRows.Count != 0));
            var page = layout.Pages[pageIndex];
            var previous = page.PlacedRelatedStories.Single(story => story.StoryLayout.Story.Id == "37");
            var current = layout.RelatedStories.Single(story => story.Story.Id == "4");
            double available = previous.TopY - previous.Height - page.MarginBottom;
            TestAssert.True(current.ContentHeight <= available && current.ContentHeight + 3d > available,
                "The measured next note fits this remaining window; an extra separator gap would force a page turn.");
            var following = page.PlacedRelatedStories.Single(story => story.StoryLayout.Story.Id == "4");
            TestAssert.True(following.TopY - following.Height >= page.MarginBottom,
                "The actual next head stays wholly within the unchanged page frame.");
        }
    }

    public static void DocxMeasuredEndnoteBoundaryRejectsUnownedOrClippedSource()
    {
        var layout = PictureLayout(36, OoxPdfDocxMarkupMode.Final);
        var previous = layout.RelatedStories.Single(story => story.Story.Id == "37");
        var current = layout.RelatedStories.Single(story => story.Story.Id == "4");
        var inspect = typeof(DocxLayoutEngine).GetMethod("HasMeasuredEndnoteParagraphBoundary", BindingFlags.Static | BindingFlags.NonPublic)!
            .CreateDelegate<Func<DocxRelatedStoryLayout, DocxRelatedStoryLayout, CancellationToken, bool>>();
        TestAssert.True(inspect(previous, current, default), "A complete measured block owns its trailing source boundary.");
        var row = previous.TableRows.Single();
        var cell = row.Cells.Single();
        foreach (var unsupported in new[] {
            previous with { TextLines = [previous.TextLines[0]] },
            previous with { HasResolvedCounterInputs = false },
            previous with { TableRows = [row with { FragmentCount = 2 }] },
            previous with { TableRows = [row with { Cells = [cell with { Height = .001d }] }] },
            previous with { TableRows = [row with { Cells = [cell with { VisualOwnership = DocxTableCellVisualOwnership.VerticalMergeOwner }] }] } })
            TestAssert.True(!inspect(unsupported, current, default),
                "Missing tail ownership, unresolved inputs, source clipping and borrowed/fragmented rows keep the existing fallback.");
        TestAssert.True(!inspect(previous, current with { TextLines = [] }, default), "A source paragraph without a measured head is not admitted.");
        TestAssert.True(!inspect(previous, current with { TextLines = [current.TextLines[0] with {
            SourceParagraph = previous.TextLines[^1].SourceParagraph }] }, default), "Head ownership belongs to the next source paragraph.");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() =>
            inspect(previous, current, cancellation.Token)).CancellationToken);
    }

    private static DocxLayout PictureLayout(int height, OoxPdfDocxMarkupMode mode, int? exactHeight = null, int laterNoteWords = 0)
    {
        var source = DocxInlineCompositionTests.ReadMutation("CompositionProbes/mixed-note-endnote-table", mode, false, false,
            (xml, w, _) =>
            {
                var note = xml.Root!.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "id") == "37");
                var row = note.Element(w + "tbl")!.Element(w + "tr")!;
                var cell = row.Element(w + "tc")!;
                var drawing = new XElement(cell.Descendants(w + "drawing").Single());
                XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                foreach (var extent in drawing.Descendants().Where(node => node.Name == wp + "extent" || node.Name == a + "ext"))
                { extent.SetAttributeValue("cx", 36 * 12700); extent.SetAttributeValue("cy", height * 12700); }
                cell.Elements().Where(node => node.Name != w + "tcPr").Remove();
                cell.Add(new XElement(w + "p", new XElement(w + "pPr", new XElement(w + "spacing",
                    new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"), new XAttribute(w + "line", "240"),
                    new XAttribute(w + "lineRule", "auto"))), new XElement(w + "r", drawing)));
                if (exactHeight is { } exact)
                {
                    var properties = row.Element(w + "trPr");
                    if (properties is null) { properties = new XElement(w + "trPr"); row.AddFirst(properties); }
                    properties.Add(new XElement(w + "trHeight", new XAttribute(w + "val", exact * 20), new XAttribute(w + "hRule", "exact")));
                }
                if (laterNoteWords != 0)
                    xml.Root.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "id") == "8").Element(w + "p")!
                        .Add(new XElement(w + "r", new XElement(w + "t", " " + string.Join(" ",
                            Enumerable.Range(0, laterNoteWords).Select(index => $"LaterNote{index:000}")))));
            }, partOverride: "word/endnotes.xml");
        return new DocxLayoutEngine(mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup,
                wordCompatiblePrintScale: 1d)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
    }
}
