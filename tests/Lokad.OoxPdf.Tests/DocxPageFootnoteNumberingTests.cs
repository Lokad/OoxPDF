using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxPageFootnoteNumberingTests
{
    public static void DocxPageFootnoteNumbersRestartWithoutChangingAuthoredDigits()
    {
        foreach (bool runBreak in new[] { false, true })
        foreach ((string format, int start) in new[] { ("decimal", 1), ("decimal", 4), ("decimalZero", 9) })
        {
            DocxDocument source = ReadFixture(runBreak, format, start);
            DocxLayout first = CreateLayout(source);
            TestAssert.Equal(2, first.Pages.Count);
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, first, out var numbered, out var map),
                "A complete one-line note on each marker's physical page can restart.");
            TestAssert.Equal("1,1,2,2", string.Join(',', new[] { "37", "4", "73", "8" }.Select(id => map[id])));
            string[] labels = Enumerable.Range(0, 4).Select(i => (start + i % 2)
                .ToString(format == "decimalZero" ? "D2" : "D", CultureInfo.InvariantCulture)).ToArray();
            for (int i = 0; i < 4; i++)
            {
                DocxParagraph before = source.Paragraphs.First(p => p.InlineReferences.Any(r => r.Id == new[] { "37", "4", "73", "8" }[i]));
                DocxParagraph after = numbered.Paragraphs.First(p => p.InlineReferences.Any(r => r.Id == before.InlineReferences[0].Id));
                TestAssert.Equal(labels[i], after.InlineReferences[0].DisplayText);
                TestAssert.Equal(labels[i], numbered.RelatedStories.Single(s => s.Id == before.InlineReferences[0].Id).Paragraphs[0].Runs[0].Text);
                TestAssert.Equal((105 + i).ToString(CultureInfo.InvariantCulture), before.InlineReferences[0].DisplayText);
                TestAssert.Equal(before.Runs[0], after.Runs[0]);
                TestAssert.Equal(before.Runs[^1], after.Runs[^1]);
                TestAssert.Contains("105", after.Runs[0].Text);
                TestAssert.Contains("108", after.Runs[^1].Text);
                TestAssert.Equal(before.InlineReferences[0].SourceRunIndex, after.InlineReferences[0].SourceRunIndex);
            }
            TestAssert.Equal(source.Settings, numbered.Settings);
            TestAssert.Equal(source.FinalSectionBreak, numbered.FinalSectionBreak);
            TestAssert.True(DocxPageFootnoteNumbering.HasStablePages(numbered, CreateLayout(numbered), map), "Renumbering must keep complete notes with their markers.");
        }
    }

    public static void DocxPageFootnoteNumbersDistinguishBreakSpillFromWrapping()
    {
        DocxDocument source = ReadFixture(runBreak: true);
        DocxLayout layout = CreateLayout(source);
        TestAssert.True(layout.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>()
            .Any(l => l.Segments.Count > 0 && l.Segments.All(s => s.Role == DocxTextSegmentRole.BreakSpill)), "The run-break fixture must exercise a real spacing row.");
        TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, layout, out _, out _), "Spacing rows do not make the note marker multiline.");
        DocxDocument wrapped = ReadFixture(runBreak: true, modify: (xml, _, _) =>
            xml.Descendants(Word + "t").First().Value = string.Join(' ', Enumerable.Repeat("Public wrapped body text", 25)));
        DocxLayout wrappedLayout = CreateLayout(wrapped);
        DocxParagraph paragraph = wrapped.Paragraphs.First(p => p.InlineReferences.Count != 0);
        TestAssert.True(wrappedLayout.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>()
            .Count(l => ReferenceEquals(l.SourceParagraph, paragraph) && l.Segments.Any(s => s.Role == DocxTextSegmentRole.Text)) > 1,
            "The exclusion must contain actual wrapped text lines.");
        TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(wrapped, wrappedLayout, out var fallback, out _), "Wrapped marker paragraphs retain their fallback.");
        TestAssert.True(ReferenceEquals(wrapped, fallback), "Rejected numbering keeps the original document.");
    }

    public static void DocxPageFootnoteNumbersRejectMovedOrPartialNotes()
    {
        DocxDocument source = ReadFixture();
        DocxLayout first = CreateLayout(source);
        TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, first, out var numbered, out var pages), "Fixture must qualify initially.");
        DocxLayout final = CreateLayout(numbered);
        DocxPlacedRelatedStoryLayout note = final.Pages[0].PlacedRelatedStories.First(n => n.StoryLayout.Story.IsNormalStoryType);
        foreach (int variant in Enumerable.Range(0, 3))
        {
            var changedPages = final.Pages.ToArray();
            var notes = changedPages[0].PlacedRelatedStories.ToList();
            if (variant == 0)
            {
                notes.Remove(note);
                changedPages[1] = changedPages[1] with { PlacedRelatedStories = [.. changedPages[1].PlacedRelatedStories, note] };
            }
            if (variant == 1) notes[notes.IndexOf(note)] = note with { ContentTopOffset = 1d };
            if (variant == 2) notes.Add(note);
            changedPages[0] = changedPages[0] with { PlacedRelatedStories = notes };
            TestAssert.True(!DocxPageFootnoteNumbering.HasStablePages(numbered, final with { Pages = changedPages }, pages),
                "Moved, partial, and duplicate note placements must reject the retry.");
        }
        var moved = pages.ToDictionary(p => p.Key, p => p.Key == "37" ? 2 : p.Value);
        TestAssert.True(!DocxPageFootnoteNumbering.HasStablePages(numbered, final, moved), "Body marker pages are part of retry stability.");
    }

    public static void DocxPageFootnoteNumbersKeepExcludedSettingsAndStories()
    {
        Action<XDocument, XDocument, XDocument>[] exclusions = [
            (xml, _, _) => xml.Descendants(Word + "footnoteReference").First().SetAttributeValue(Word + "customMarkFollows", "1"),
            (xml, _, _) => xml.Descendants(Word + "footnotePr").Single().Element(Word + "pos")!.SetAttributeValue(Word + "val", "beneathText"),
            (xml, _, _) => xml.Descendants(Word + "numFmt").Single().SetAttributeValue(Word + "val", "upperRoman"),
            (xml, _, _) => xml.Descendants(Word + "numStart").Single().SetAttributeValue(Word + "val", "0"),
            (xml, _, _) => xml.Descendants(Word + "sectPr").Single().Add(new XElement(Word + "cols", new XAttribute(Word + "num", "2"))),
            (xml, _, _) => xml.Descendants(Word + "p").First().Add(new XElement(Word + "ins", new XAttribute(Word + "id", "9"),
                new XElement(Word + "r", new XElement(Word + "t", "public revision")))),
            (xml, _, _) => { XElement p = xml.Descendants(Word + "p").First(); p.ReplaceWith(new XElement(Word + "tbl", new XElement(Word + "tr", new XElement(Word + "tc", p)))); },
            (_, settings, _) => settings.Descendants(Word + "footnotePr").Single().Add(new XElement(Word + "numRestart", new XAttribute(Word + "val", "eachPage"))),
            (_, _, notes) => notes.Root!.Elements(Word + "footnote").First(n => n.Attribute(Word + "type") is null)
                .Add(new XElement(Word + "p", new XElement(Word + "r", new XElement(Word + "t", "another note paragraph"))))
        ];
        for (int i = 0; i < exclusions.Length; i++)
        {
            DocxDocument source = ReadFixture(modify: exclusions[i]);
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, CreateLayout(source), out var fallback, out _), "Excluded note scope " + i);
            TestAssert.True(ReferenceEquals(source, fallback), "Fallback preserves authored metadata and numbering.");
        }
    }

    public static void DocxPageFootnoteNumbersRenderLabelsAndDistinctDestinationsTogether()
    {
        foreach (OoxPdfDocxMarkupGeometryMode geometry in new[] { OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup })
        {
            DocxDocument source = ReadFixture(runBreak: true, format: "decimalZero", start: 9);
            var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, geometry);
            PdfPage[] pages = renderer.RenderBlankPages(source, null, CancellationToken.None).ToArray();
            TestAssert.Equal(2, pages.Length);
            for (int i = 0; i < pages.Length; i++)
            {
                string text = DecodeText(pages[i]);
                TestAssert.Contains("before note 105 09 and after note 108", text);
                TestAssert.Contains("before note 105 10 and after note 108", text);
                TestAssert.Contains("09 Public destination", text);
                TestAssert.Contains("10 Public destination", text);
                TestAssert.Equal(2, pages[i].Annotations.Count);
                TestAssert.Equal(2, pages[i].Annotations.Select(a => a.Destination).Distinct().Count());
                foreach (var link in pages[i].Annotations)
                {
                    TestAssert.Equal(i, link.Destination!.Value.PageIndex);
                    TestAssert.True(link.Width > 0d && link.Height > 0d && link.Destination.Value.Top > 0d &&
                        link.Destination.Value.Top < 150d, "Every renumbered marker points into its own rendered note area.");
                }
            }
            TestAssert.Equal("105", source.Paragraphs.First(p => p.InlineReferences.Count != 0).InlineReferences[0].DisplayText);
        }
    }

    private static DocxLayout CreateLayout(DocxDocument document) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
        .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

    private static string DecodeText(PdfPage page)
    {
        var text = new StringBuilder();
        foreach (Match region in Regex.Matches(page.Content, @"BT(.*?)ET", RegexOptions.Singleline))
        {
            Match name = Regex.Match(region.Value, @"/(\S+)\s+[\d.]+\s+Tf");
            PdfEmbeddedFont font = page.Fonts.Single(f => f.ResourceName == name.Groups[1].Value).Font;
            foreach (Match hex in Regex.Matches(region.Value, @"<([0-9A-Fa-f]+)>"))
                for (int i = 0; i < hex.Groups[1].Value.Length; i += 4)
                    text.Append(char.ConvertFromUtf32(font.UnicodeByCid[ushort.Parse(hex.Groups[1].Value.Substring(i, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture)]));
        }
        return text.ToString();
    }

    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static DocxDocument ReadFixture(bool runBreak = false, string format = "decimal", int start = 4,
        Action<XDocument, XDocument, XDocument>? modify = null)
    {
        string seed = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-nav-numbered-ids.docx"));
        var parts = new Dictionary<string, string>();
        using (ZipArchive archive = ZipFile.OpenRead(seed))
            foreach (var entry in archive.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        XDocument xml = XDocument.Parse(parts["word/document.xml"]);
        XElement body = xml.Root!.Element(Word + "body")!;
        XElement template = new(body.Elements(Word + "p").First());
        XElement section = new(xml.Descendants(Word + "sectPr").Last());
        section.Elements(Word + "footnotePr").Remove();
        section.AddFirst(new XElement(Word + "footnotePr", new XElement(Word + "pos", new XAttribute(Word + "val", "pageBottom")),
            new XElement(Word + "numFmt", new XAttribute(Word + "val", format)),
            new XElement(Word + "numStart", new XAttribute(Word + "val", start)),
            new XElement(Word + "numRestart", new XAttribute(Word + "val", "eachPage"))));
        body.RemoveNodes();
        XDocument notes = XDocument.Parse(parts["word/footnotes.xml"]);
        XElement noteTemplate = new(notes.Root!.Elements(Word + "footnote").First(n => n.Attribute(Word + "type") is null));
        notes.Root!.Elements(Word + "footnote").Where(n => n.Attribute(Word + "type") is null).Remove();
        for (int i = 0; i < 4; i++)
        {
            if (i == 2 && runBreak) body.Add(new XElement(Word + "p", new XElement(Word + "r", new XElement(Word + "br", new XAttribute(Word + "type", "page")))));
            XElement p = new(template);
            p.Descendants(Word + "sectPr").Remove();
            if (i == 2 && !runBreak) p.Element(Word + "pPr")!.Add(new XElement(Word + "pageBreakBefore"));
            p.Descendants(Word + "t").First().Value = new[] { "Alpha", "Bravo", "Charlie", "Delta" }[i] + " before note 105 ";
            p.Descendants(Word + "t").Last().Value = " and after note 108.";
            string id = new[] { "37", "4", "73", "8" }[i];
            p.Descendants(Word + "footnoteReference").Single().SetAttributeValue(Word + "id", id);
            body.Add(p);
            XElement note = new(noteTemplate); note.SetAttributeValue(Word + "id", id);
            note.Descendants(Word + "t").First().Value = " Public destination " + new[] { "Alpha", "Bravo", "Charlie", "Delta" }[i] + ".";
            notes.Root.Add(note);
        }
        body.Add(section);
        XDocument settings = new(new XElement(Word + "settings", new XElement(Word + "footnotePr",
            new XElement(Word + "numFmt", new XAttribute(Word + "val", "decimal")),
            new XElement(Word + "numStart", new XAttribute(Word + "val", "105")))));
        modify?.Invoke(xml, settings, notes);
        parts["word/document.xml"] = xml.ToString(); parts["word/settings.xml"] = settings.ToString(); parts["word/footnotes.xml"] = notes.ToString();
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(input); }
    }
}
