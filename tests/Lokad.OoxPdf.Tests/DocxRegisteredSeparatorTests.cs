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
            TestAssert.True(Math.Abs(baseline.Rule - before.Rule - 12d) < .001d &&
                Math.Abs(baseline.Rule - after.Rule) < .001d && Math.Abs(baseline.Rule - inherited.Rule) < .001d,
                "Before-space moves the line; after-space advances content below a stationary separator rule.");
            TestAssert.Equal(baseline.Page, before.Page);
            TestAssert.Equal(baseline.Page, after.Page);
            TestAssert.True(!baseline.Synthetic, "Settings registration retains the authored separator.");
            var defaultA = Place(0, 0, false);
            var defaultB = Place(12, 24, false);
            TestAssert.True(defaultA.Synthetic && defaultB.Synthetic && defaultA == defaultB,
                "Unregistered formatting keeps the same synthetic default and note frame.");
            var missing = Place(0, 0, false, absent: true);
            var inheritedDefault = Place(0, 0, false, inheritedBefore: true);
            foreach (var normalized in new[] { defaultA, missing, inheritedDefault })
                TestAssert.True(normalized.Synthetic && normalized.Page == baseline.Page &&
                    Math.Abs(normalized.Top - baseline.Top) < .001d && Math.Abs(normalized.Rule - baseline.Rule) < .001d,
                    "Absent/unregistered endnote separators normalize once to the same zero-spacing source frame.");

            (int Page, double Top, bool Synthetic, double Rule) Place(int beforePoints, int? afterPoints, bool registered,
                bool absent = false, bool inheritedBefore = false)
            {
                var source = Read(mode, position, beforePoints, afterPoints, registered,
                    absent: absent, inheritedBefore: inheritedBefore);
                var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
                var placed = layout.Pages.SelectMany((page, index) => page.PlacedRelatedStories
                    .Where(note => note.StoryLayout.Story.Id == "37")
                    .Select(note => (Page: index, Top: note.TopY))).Single();
                double rule = layout.Pages.SelectMany(page => page.PlacedRelatedStories)
                    .Single(note => note.StoryLayout.Story.Type == DocxRelatedStoryType.Separator).SeparatorY!.Value;
                return (placed.Page, placed.Top, source.RelatedStories.Single(story =>
                    story.Type == DocxRelatedStoryType.Separator).IsSynthetic, rule);
            }
        }
    }

    public static void DocxRegisteredEndnoteContinuationSpacingOwnsItsPageCapacity()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string position in new[] { "sectEnd", "docEnd" })
        {
            var before = Layout(0);
            var after = Layout(12);
            var first = before.Pages.SelectMany(page => page.PlacedRelatedStories.Where(note => note.StoryLayout.Story.Id == "37")).ToArray();
            var second = after.Pages.SelectMany(page => page.PlacedRelatedStories.Where(note => note.StoryLayout.Story.Id == "37")).ToArray();
            TestAssert.True(first.Length >= 3 && second.Length >= 3, "The note actually reaches repeated continuation frames.");
            TestAssert.Equal(first[0].TopY, second[0].TopY);
            TestAssert.True(first[0].TextLines.Select(line => line.Text).SequenceEqual(second[0].TextLines.Select(line => line.Text)),
                "Continuation-only formatting preserves the already placed head.");
            TestAssert.True(Math.Abs(first[1].TopY - second[1].TopY - 12d) < .001d,
                "The continued content frame owns the registered after-space.");
            TestAssert.True(second[1].TextLines.Count <= first[1].TextLines.Count,
                "Authored separator spacing also reduces the admitted continuation capacity.");
            TestAssert.True(first.SelectMany(note => note.TextLines).Select(line => line.Text)
                .SequenceEqual(second.SelectMany(note => note.TextLines).Select(line => line.Text)),
                "Every source line survives once and in order through changed page takes.");

            DocxLayout Layout(int points) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(Read(mode, position, 0, 0, true, continuationAfter: points, words: 900),
                    new DocxTests.FamilyWidthTextMeasurer(), default);
        }
    }

    private static DocxDocument Read(OoxPdfDocxMarkupMode mode, string position,
        int before, int? after, bool registered, int? continuationAfter = null, int words = 3,
        bool absent = false, bool inheritedBefore = false)
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
        if (absent) separator.Remove();
        if (inheritedBefore)
        {
            var styles = Load("word/styles.xml");
            styles.Root!.Element(w + "docDefaults")!.Add(new XElement(w + "pPrDefault", new XElement(w + "pPr",
                new XElement(w + "spacing", new XAttribute(w + "before", "240"), new XAttribute(w + "after", "160"),
                    new XAttribute(w + "line", "278"), new XAttribute(w + "lineRule", "auto")))));
            Save("word/styles.xml", styles);
        }
        if (continuationAfter.HasValue)
        {
            var continuation = notes.Root.Elements(w + "endnote")
                .Single(note => (string?)note.Attribute(w + "type") == "continuationSeparator");
            noteProperties.Add(new XElement(w + "endnote", new XAttribute(w + "id", (string)continuation.Attribute(w + "id")!)));
            continuation.ReplaceNodes(new XElement(w + "p", new XElement(w + "pPr", new XElement(w + "spacing",
                new XAttribute(w + "before", "0"), new XAttribute(w + "after", continuationAfter.Value * 20))),
                new XElement(w + "r", new XElement(w + "continuationSeparator"))));
        }
        notes.Root.Elements(w + "endnote").Where(note => (string?)note.Attribute(w + "type") is null or "normal").Remove();
        notes.Root.Add(new XElement(w + "endnote", new XAttribute(w + "id", "37"),
            Paragraph(string.Join(" ", Enumerable.Range(0, words).Select(index => "item" + index.ToString("D3", System.Globalization.CultureInfo.InvariantCulture))))));
        Save("word/document.xml", document);
        Save("word/settings.xml", settings);
        Save("word/endnotes.xml", notes);
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, mode); }
        finally { File.Delete(temporary); }
    }
}
