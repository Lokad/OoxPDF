using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxRegisteredSeparatorTests
{
    public static void DocxRegisteredEndnoteSeparatorsRetainSourceSpacingAndDefaultFallbacks()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string position in new[] { "sectEnd", "docEnd" })
        {
            var baseline = Place(0, 0, true);
            var before = Place(12, 0, true);
            var after = Place(0, 12, true);
            var inherited = Place(0, null, true);
            TestAssert.True(Math.Abs(baseline.Top - before.Top - 12d) < .001d,
                "Registered separator before-space survives measurement and " + position + " placement.");
            TestAssert.True(Math.Abs(baseline.Top - after.Top - 12d) < .001d,
                "Registered separator after-space is not replaced or refunded as latent padding.");
            TestAssert.True(Math.Abs(baseline.Top - inherited.Top - 8d) < .001d,
                "Omitted authored spacing follows the existing source style cascade.");
            TestAssert.Equal(baseline.Page, before.Page);
            TestAssert.Equal(baseline.Page, after.Page);
            TestAssert.True(!baseline.Synthetic, "Settings registration retains the authored separator.");
            var defaultA = Place(0, 0, false);
            var defaultB = Place(12, 24, false);
            TestAssert.True(defaultA.Synthetic && defaultB.Synthetic && defaultA == defaultB,
                "Unregistered formatting keeps the same synthetic default and note frame.");

            (int Page, double Top, bool Synthetic) Place(int beforePoints, int? afterPoints, bool registered)
            {
                var source = Read(mode, position, beforePoints, afterPoints, registered);
                var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
                var placed = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                    .Where(note => note.StoryLayout.Story.Id == "37")
                    .Select(note => (Page: index, Top: note.TopY))).Single();
                return (placed.Page, placed.Top, source.RelatedStories.Single(story =>
                    story.Type == DocxRelatedStoryType.Separator).IsSynthetic);
            }
        }
    }

    private static DocxDocument Read(OoxPdfDocxMarkupMode mode, string position,
        int before, int? after, bool registered)
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "CompositionProbes", "mixed-note-endnote-table.docx"));
        var parts = new Dictionary<string, byte[]>();
        using (var archive = ZipFile.OpenRead(fixture))
        foreach (var entry in archive.Entries)
        {
            using var input = entry.Open();
            using var bytes = new MemoryStream();
            input.CopyTo(bytes);
            parts.Add(entry.FullName, bytes.ToArray());
        }
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XDocument Load(string name) => XDocument.Parse(Encoding.UTF8.GetString(parts[name]));
        void Save(string name, XDocument xml) => parts[name] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        var document = Load("word/document.xml");
        var body = document.Root!.Element(w + "body")!;
        var section = new XElement(body.Element(w + "sectPr")!);
        XElement Paragraph(string text) => new(w + "p", new XElement(w + "pPr",
            new XElement(w + "spacing", new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"))),
            new XElement(w + "r", new XElement(w + "t", text)));
        var anchor = Paragraph("Public reference");
        anchor.Add(new XElement(w + "r", new XElement(w + "endnoteReference", new XAttribute(w + "id", "37"))));
        var closing = Paragraph("Public section end");
        closing.Element(w + "pPr")!.Add(new XElement(section));
        body.ReplaceNodes(anchor, closing, Paragraph("Public following section"), section);
        var settings = Load("word/settings.xml");
        var noteProperties = settings.Root!.Element(w + "endnotePr")!;
        noteProperties.Elements(w + "pos").Remove();
        noteProperties.Elements(w + "endnote").Remove();
        noteProperties.Add(new XElement(w + "pos", new XAttribute(w + "val", position)));
        var notes = Load("word/endnotes.xml");
        var separator = notes.Root!.Elements(w + "endnote").Single(note => (string?)note.Attribute(w + "type") == "separator");
        if (registered) noteProperties.Add(new XElement(w + "endnote", new XAttribute(w + "id", (string)separator.Attribute(w + "id")!)));
        var spacing = new XElement(w + "spacing", new XAttribute(w + "before", before * 20));
        if (after.HasValue) spacing.Add(new XAttribute(w + "after", after.Value * 20));
        separator.ReplaceNodes(new XElement(w + "p", new XElement(w + "pPr", spacing),
            new XElement(w + "r", new XElement(w + "separator"))));
        notes.Root.Elements(w + "endnote").Where(note => (string?)note.Attribute(w + "type") is null or "normal").Remove();
        notes.Root.Add(new XElement(w + "endnote", new XAttribute(w + "id", "37"), Paragraph("Public note body")));
        Save("word/document.xml", document);
        Save("word/settings.xml", settings);
        Save("word/endnotes.xml", notes);
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, mode); }
        finally { File.Delete(temporary); }
    }
}
