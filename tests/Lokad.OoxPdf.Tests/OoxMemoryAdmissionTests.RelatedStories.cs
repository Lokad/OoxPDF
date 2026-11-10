using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void LayoutRelatedStoryMeasurementsSkipEndnoteRegistrationAndKeepWidths()
    {
        foreach (bool keep in new[] { false, true })
            CheckRelatedStoryMeasurements(keep, footnote: false);
    }

    public static void LayoutRelatedStoryMeasurementsPreserveLiveFootnoteWidthAndReuse()
    {
        foreach (bool keep in new[] { false, true })
            CheckRelatedStoryMeasurements(keep, footnote: true);
    }

    private static void CheckRelatedStoryMeasurements(bool keep, bool footnote)
    {
        var measurer = LayoutAdmissionMeasurer();
        var pool = new TestMemoryBudget();
        using var input = RelatedStoryWidthInput(keep, footnote);
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            var package = OoxPackage.Open(input, default);
            var source = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
            var measurements = new Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>();
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                { PlacementRelatedStoryMeasurements = measurements };
            var first = engine.Create(source, measurer, default);
            TestAssert.Equal(6, first.Pages.Select(page => page.Width).Distinct().Count());
            TestAssert.True(measurements.Count == (footnote ? 3 : 2),
                "Only the initial/final endnote frames and an actual footnote owner need measurements.");
            TestAssert.Equal(6, first.Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(note => note.StoryLayout.Story.IsNormalStoryType && note.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote)
                .Select(note => note.StoryLayout.Story.Id).Distinct().Count());
            if (footnote)
            {
                var owner = first.Pages.Single(page => page.PlacedRelatedStories.Any(note =>
                    note.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && note.StoryLayout.Story.IsNormalStoryType));
                TestAssert.True(owner.Width == 11840d / 20d, "Keep lookahead and registration retain the reference's section.");
                TestAssert.True(measurements.ContainsKey(owner.Width - owner.MarginLeft - owner.MarginRight),
                    "A referenced footnote still measures its actual width.");
            }
            var cached = measurements.ToArray();
            long held = pool.Reserved;
            pool.Capacity = held;
            var second = engine.Create(source, measurer, default);
            TestAssert.Equal(first.Pages.Count, second.Pages.Count);
            TestAssert.Equal(held, pool.Reserved);
            TestAssert.Equal(cached.Length, measurements.Count);
            TestAssert.True(cached.All(pair => ReferenceEquals(pair.Value, measurements[pair.Key])),
                "The second preparation reuses source-owned measurements.");
            package.Dispose();
            TestAssert.True(pool.Reserved > 0, "Retained source and placements outlive package buffers.");
            TestAssert.True(measurements.Values.SelectMany(value => value).Any(note =>
                note.TextLines.Any(line => line.Text.Contains("Note text", StringComparison.Ordinal))),
                "Cached note content remains readable after package disposal.");
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    private static MemoryStream RelatedStoryWidthInput(bool keep, bool footnote)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string Section(int index) => $"<w:sectPr><w:pgSz w:w=\"{12240 - index * 200}\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr>";
        string body = string.Concat(Enumerable.Range(0, 6).Select(index =>
            (keep ? "<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>Lead</w:t></w:r></w:p>" : "") +
            $"<w:p><w:pPr>{(index < 5 ? Section(index) : "")}</w:pPr><w:r><w:t>Body</w:t></w:r><w:r><w:endnoteReference w:id=\"{index + 1}\"/></w:r>" +
            (footnote && index == 2 ? "<w:r><w:footnoteReference w:id=\"1\"/></w:r>" : "") + "</w:p>"));
        return TestFixtures.CreateZipPackage(new Dictionary<string, string> {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{w}\"><w:body>{body}{Section(5)}</w:body></w:document>",
            ["word/endnotes.xml"] = $"<w:endnotes xmlns:w=\"{w}\">" + string.Concat(Enumerable.Range(1, 6).Select(id =>
                $"<w:endnote w:id=\"{id}\"><w:p><w:r><w:endnoteRef/><w:t>Note text {id}</w:t></w:r></w:p></w:endnote>")) + "</w:endnotes>",
            ["word/footnotes.xml"] = $"<w:footnotes xmlns:w=\"{w}\"><w:footnote w:id=\"1\"><w:p><w:r><w:footnoteRef/><w:t>Live footnote</w:t></w:r></w:p></w:footnote></w:footnotes>",
            ["word/_rels/document.xml.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"e1\" Type=\"{rel}/endnotes\" Target=\"endnotes.xml\"/>" +
                (footnote ? $"<Relationship Id=\"f1\" Type=\"{rel}/footnotes\" Target=\"footnotes.xml\"/>" : "") + "</Relationships>"
        });
    }
}
