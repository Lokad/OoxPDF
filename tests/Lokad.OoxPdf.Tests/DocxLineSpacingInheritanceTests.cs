using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxLineSpacingInheritanceTests
{
    public static void DocxLineSpacingInheritanceAutomaticOverridesReplacePointHeights()
    {
        foreach (string origin in Origins)
        foreach (string parent in new[] { "exact", "atLeast" })
        foreach (var (leaf, factor) in new[] { ("w:line=\"240\" w:lineRule=\"auto\"", 1d), ("w:line=\"480\"", 2d) })
        {
            var actual = Read(origin, parent, leaf);
            var control = Read(origin, "auto", leaf);
            foreach (var paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(actual))
            {
                TestAssert.True(paragraph.LineSpacingPoints is null, "An automatic child cannot retain a point height from its parent.");
                TestAssert.Equal(factor, paragraph.LineSpacingFactor);
                TestAssert.Equal(leaf.Contains("lineRule") ? "auto" : "", paragraph.Spacing.LineRuleValue ?? "");
            }
            TestAssert.Equal(LayoutSignature(control), LayoutSignature(actual));
        }
    }

    public static void DocxLineSpacingInheritanceRuleOnlyLeavesRetainUnitsAndRawTokens()
    {
        foreach (string origin in Origins)
        foreach (string parent in new[] { "exact", "atLeast", "auto" })
        foreach (bool table in new[] { false, true })
        {
            var actual = Read(origin, parent, "w:lineRule=\"auto\"", table);
            var control = Read(origin, parent, "", table);
            foreach (var paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(actual))
            {
                TestAssert.Equal("auto", paragraph.Spacing.LineRuleValue ?? "");
                TestAssert.Equal(parent == "auto" ? "480" : "320", paragraph.Spacing.LineValue ?? "");
                TestAssert.Equal(parent == "auto" ? null : (double?)16d, paragraph.LineSpacingPoints);
            }
            if (parent != "atLeast")
                TestAssert.Equal(LayoutSignature(control), LayoutSignature(actual));
        }
    }

    public static void DocxLineSpacingInheritancePointOverridesReplaceAutomaticFactors()
    {
        foreach (string origin in Origins)
        foreach (string rule in new[] { "exact", "atLeast" })
        {
            string leaf = $"w:line=\"320\" w:lineRule=\"{rule}\"";
            var actual = Read(origin, "auto", leaf, lineBasedSpacing: true);
            var control = Read(origin, rule, leaf, lineBasedSpacing: true);
            foreach (var paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(actual))
            {
                TestAssert.Equal(16d, paragraph.LineSpacingPoints!.Value);
                TestAssert.True(paragraph.LineSpacingFactor != 2d, "A point override clears the inherited double-spacing factor.");
                TestAssert.Equal(8d, paragraph.SpacingBeforePoints);
                TestAssert.Equal(16d, paragraph.SpacingAfterPoints);
            }
            TestAssert.Equal(LayoutSignature(control), LayoutSignature(actual));
        }
    }

    public static void DocxLineSpacingInheritanceDefaultOverridesKeepNotesAndNavigation()
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "note-minimum-body-line-16.docx"));
        var parts = new Dictionary<string, string>();
        using (var archive = ZipFile.OpenRead(fixture))
            foreach (var entry in archive.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                parts[entry.FullName] = reader.ReadToEnd();
            }
        var styles = XDocument.Parse(parts["word/styles.xml"]);
        var defaults = styles.Root!.Element(Word + "docDefaults")!;
        defaults.Element(Word + "pPrDefault")?.Remove();
        defaults.Add(new XElement(Word + "pPrDefault", new XElement(Word + "pPr", new XElement(Word + "spacing",
            new XAttribute(Word + "line", "320"), new XAttribute(Word + "lineRule", "exact")))));
        parts["word/styles.xml"] = styles.ToString();
        foreach (string name in new[] { "word/document.xml", "word/footnotes.xml" })
        {
            var xml = XDocument.Parse(parts[name]);
            foreach (var spacing in xml.Descendants(Word + "spacing"))
            {
                spacing.SetAttributeValue(Word + "line", "240");
                spacing.SetAttributeValue(Word + "lineRule", "auto");
            }
            parts[name] = xml.ToString();
        }
        var document = ReadParts(parts);
        TestAssert.True(DocxBlockTraversal.EnumerateBodyParagraphs(document).All(p => p.LineSpacingPoints is null), "Body overrides use automatic metrics.");
        TestAssert.True(document.RelatedStories.SelectMany(DocxBlockTraversal.EnumerateBodyParagraphs).All(p => p.LineSpacingPoints is null), "Notes and separator overrides also clear the inherited point height.");
        var renderer = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var pages = renderer.RenderBlankPages(document, null, CancellationToken.None).ToArray();
        TestAssert.Equal(4, pages.Sum(p => p.Annotations.Count));
        TestAssert.Equal(4, pages.SelectMany(p => p.Annotations).Select(a => a.Destination).Distinct().Count());
        for (int i = 0; i < pages.Length; i++)
            TestAssert.True(pages[i].Annotations.All(a => a.Destination!.Value.PageIndex == i), "Actual note links retain their same-page owners.");
        TestAssert.Equal("105", document.Paragraphs.First(p => p.InlineReferences.Count != 0).InlineReferences[0].DisplayText);
        TestAssert.Equal("240", document.Paragraphs[0].Spacing.LineValue ?? "");
    }

    private static readonly string[] Origins = ["default", "style", "chain"];
    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string LayoutSignature(DocxDocument document)
    {
        var snapshot = new DocxRenderer(new TestFaceFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectLayout(document);
        return JsonSerializer.Serialize(snapshot.Pages.SelectMany(p => p.Items).SelectMany(item => item.TextLines is null ? new[] { item } : item.TextLines)
            .Where(item => item.Kind == "TextLine").Select(item => new { item.X, item.Y, item.Width, item.Height, item.LineHeightPoints, item.LineHeightSource, item.SingleLineHeightPoints }));
    }

    private static DocxDocument Read(string origin, string parent, string leaf, bool table = false, bool lineBasedSpacing = false)
    {
        string parentSpacing = $"<w:spacing w:line=\"{(parent == "auto" ? "480" : "320")}\" w:lineRule=\"{parent}\"/>";
        string leafSpacing = lineBasedSpacing ? $"<w:spacing w:beforeLines=\"50\" w:afterLines=\"100\" {leaf}/>" : $"<w:spacing w:before=\"0\" w:after=\"0\" {leaf}/>";
        string styles = origin == "default"
            ? $"<w:docDefaults><w:pPrDefault><w:pPr>{parentSpacing}</w:pPr></w:pPrDefault></w:docDefaults>"
            : $"<w:style w:type=\"paragraph\" w:styleId=\"Parent\"><w:pPr>{parentSpacing}</w:pPr></w:style>";
        if (origin == "chain") styles += $"<w:style w:type=\"paragraph\" w:styleId=\"Child\"><w:basedOn w:val=\"Parent\"/><w:pPr>{leafSpacing}</w:pPr></w:style>";
        string style = origin == "default" ? "" : $"<w:pStyle w:val=\"{(origin == "chain" ? "Child" : "Parent")}\"/>";
        string paragraphs = string.Concat(Enumerable.Range(0, 3).Select(i => $"<w:p><w:pPr>{style}{(origin == "chain" ? "" : leafSpacing)}</w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"24\"/></w:rPr><w:t>Public spacing {i}.</w:t></w:r></w:p>"));
        if (table) paragraphs = $"<w:tbl><w:tblPr><w:tblW w:w=\"8640\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"8640\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"8640\" w:type=\"dxa\"/></w:tcPr>{paragraphs}</w:tc></w:tr></w:tbl>";
        return ReadParts(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/></Types>",
            ["_rels/.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/_rels/document.xml.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>",
            ["word/styles.xml"] = $"<w:styles xmlns:w=\"{Word}\">{styles}</w:styles>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{Word}\"><w:body>{paragraphs}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr></w:body></w:document>"
        });
    }

    private static DocxDocument ReadParts(Dictionary<string, string> parts)
    {
        string path = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.AllMarkup); }
        finally { File.Delete(path); }
    }
}
