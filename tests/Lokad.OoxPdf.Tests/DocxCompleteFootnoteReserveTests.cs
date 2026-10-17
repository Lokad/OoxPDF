using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxCompleteFootnoteReserveTests
{
    public static void DocxCompleteFootnoteReserveReclaimsRowsWithoutMovingNotes()
    {
        foreach (int initialCount in new[] { 2, 3 })
        {
            DocxDocument source = ReadFixture(initialCount);
            DocxLayoutSnapshot before = Inspect(source, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            DocxLayoutSnapshot after = Inspect(source, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            int Spacers(DocxLayoutPageSnapshot p) => p.Items.Count(i => i.SourceBlockIndex >= initialCount && i.SourceBlockIndex < initialCount + 36);
            TestAssert.Equal(2, after.Pages.Count);
            TestAssert.Equal(36, after.Pages.Sum(Spacers));
            TestAssert.Equal(Spacers(before.Pages[0]), Spacers(after.Pages[0]));
            TestAssert.Equal(PageBodyPositions(before), PageBodyPositions(after));
            TestAssert.Equal(NotePlacements(before), NotePlacements(after));
            TestAssert.Equal("105", source.Paragraphs.First(p => p.InlineReferences.Count != 0).InlineReferences[0].DisplayText);
            var pages = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, CancellationToken.None).ToArray();
            TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
            for (int i = 0; i < pages.Length; i++)
            {
                TestAssert.Equal(i == 0 ? initialCount : 4 - initialCount, pages[i].Annotations.Count);
                TestAssert.Equal(pages[i].Annotations.Count, pages[i].Annotations.Select(a => a.Destination).Distinct().Count());
                TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Reclaimed rows retain every marker's actual note destination page.");
            }
        }
    }

    public static void DocxCompleteFootnoteReserveKeepsUnqualifiedStories()
    {
        Action<XDocument, XDocument>[] guards = [
            (_, notes) => notes.Root!.Elements(Word + "footnote").Single(n => (string?)n.Attribute(Word + "type") == "separator").Remove(),
            (_, notes) => notes.Root!.Elements(Word + "footnote").First(n => n.Attribute(Word + "type") is null)
                .Descendants(Word + "t").First().Value += string.Concat(Enumerable.Repeat(" Public wrapped note.", 25)),
            (xml, _) => xml.Descendants(Word + "p").First().Element(Word + "pPr")!.Add(new XElement(Word + "keepNext")),
            (xml, _) => xml.Descendants(Word + "footnotePr").Single().Element(Word + "pos")!.SetAttributeValue(Word + "val", "beneathText")
        ];
        foreach (var modify in guards)
        {
            DocxDocument source = ReadFixture(2, modify);
            var preserve = Inspect(source, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var word = Inspect(source, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            TestAssert.Equal(PageBodyPositions(preserve), PageBodyPositions(word));
            TestAssert.Equal(NotePlacements(preserve), NotePlacements(word));
        }
    }

    public static void DocxCompleteFootnoteReserveRejectsChangedCompleteOwnerPages()
    {
        DocxDocument document = ReadFixture(2);
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxLayout first = engine.Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.True(DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(document, first, first), "Witness must have complete same-page owners.");
        var swapped = first with { Pages = first.Pages.Reverse().ToArray() };
        TestAssert.True(DocxPageFootnoteNumbering.HasCompleteNoteOwners(document, swapped), "Moving a whole page keeps markers together with their notes.");
        TestAssert.True(!DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(document, first, swapped), "A reserve retry may reclaim rows but cannot relocate complete marker/note pairs to another page.");
        TestAssert.True(!DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(document, first, first with { Pages = [.. first.Pages, first.Pages[0]] }), "Page-count changes retain fallback.");
    }

    private static DocxLayoutSnapshot Inspect(DocxDocument source, OoxPdfDocxMarkupGeometryMode geometry) =>
        new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry).InspectLayout(source);

    private static string NotePlacements(DocxLayoutSnapshot snapshot) => JsonSerializer.Serialize(snapshot.Pages.SelectMany((p, i) =>
        p.PlacedRelatedStories.Select(n => new { Page = i, n.Id, n.Type, n.TopY, n.Height, n.ContentTopOffset, n.ContentHeight, n.SeparatorY })));

    private static string PageBodyPositions(DocxLayoutSnapshot snapshot) => JsonSerializer.Serialize(snapshot.Pages.Select(p =>
        p.Items.Select(i => new { i.SourceBlockIndex, i.SourceLineIndex, i.X, i.Y, i.Height })));

    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static DocxDocument ReadFixture(int initialCount, Action<XDocument, XDocument>? modify = null)
    {
        string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", $"note-complete-reserve-{initialCount}.docx"));
        if (modify is null) return DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
        var parts = new Dictionary<string, string>();
        using (ZipArchive archive = ZipFile.OpenRead(input))
            foreach (var entry in archive.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        XDocument xml = XDocument.Parse(parts["word/document.xml"]), notes = XDocument.Parse(parts["word/footnotes.xml"]);
        modify(xml, notes);
        parts["word/document.xml"] = xml.ToString(); parts["word/footnotes.xml"] = notes.ToString();
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(temporary); }
    }
}
