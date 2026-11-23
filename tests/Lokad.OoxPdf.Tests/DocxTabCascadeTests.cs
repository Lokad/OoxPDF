using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTabCascadeTests
{
    public static void DocxTabCascadeAddsAcrossDefaultsStylesAndDirectProperties()
    {
        foreach (string kind in new[] { "defaults-add", "style-add", "direct-add" })
        {
            var paragraph = Owner(Read(kind, false));
            TestAssert.Equal("90,180,270", string.Join(",", paragraph.TabStops.OrderBy(tab => tab.PositionPoints).Select(tab => tab.PositionPoints)));
            TestAssert.Equal(12d, paragraph.SpacingBeforePoints);
            TestAssert.Equal(DocxTextAlignment.Left, paragraph.Alignment);
        }
    }

    public static void DocxTabCascadeClearsAndReplacesOnlyAddressedPositions()
    {
        var cleared = Owner(Read("clear-one", false));
        TestAssert.True(cleared.TabStops.Any(tab => tab.PositionPoints == 90d && tab.Value == "clear"),
            "The clear declaration remains inspectable after removing its inherited positioning stop.");
        TestAssert.True(cleared.TabStops.Count(tab => tab.PositionPoints is 180d or 270d && tab.Value == "left") == 2,
            "Clearing one stop preserves the other inherited stops.");
        var replaced = Owner(Read("override-one", false));
        TestAssert.Equal(3, replaced.TabStops.Count);
        TestAssert.Equal("center", replaced.TabStops.Single(tab => tab.PositionPoints == 180d).Value);
        var retained = replaced.TabStops.Single(tab => tab.PositionPoints == 90d);
        TestAssert.Equal("01800", retained.PositionValue);
        TestAssert.Equal("dot", retained.LeaderValue);
    }

    public static void DocxTabCascadeLayeredSourcesMatchFlattenedPaint()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string container in new[] { "body", "cell", "header" })
        foreach (string kind in new[] { "defaults-add", "style-add", "direct-add", "clear-one", "override-one", "logical" })
        {
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final ?
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var layered = renderer.RenderPagesAsync(Read(kind, false, container, mode), null, default).ToBlockingEnumerable().Single();
            var flattened = renderer.RenderPagesAsync(Read(kind, true, container, mode), null, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(flattened.Content, layered.Content);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(layered, true, includeTextTransforms: true).ToArray();
            TestAssert.Equal("ABCD", string.Concat(shows.Select(show => show.Text)).Replace(" ", "", StringComparison.Ordinal));
            var letters = shows.Where(show => show.Text.Contains('B') || show.Text.Contains('C') || show.Text.Contains('D')).ToArray();
            TestAssert.True(letters.Length != 0 && letters.All(show => show.X >= 72d && show.X < 540d),
                "Every tab field remains present inside its container: " + container + "/" + kind);
        }
    }

    private static DocxParagraph Owner(DocxDocument document) =>
        DocxBlockTraversal.EnumerateSourceParagraphs(document).Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance)
            .Single(paragraph => paragraph.Runs.Any(run => run.Text.Contains('A')));

    private static DocxDocument Read(string kind, bool flattened, string container = "body",
        OoxPdfDocxMarkupMode mode = OoxPdfDocxMarkupMode.Final)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        string Stop(string position, string alignment = "left", string? leader = null) =>
            $"<w:tab w:pos=\"{position}\" w:val=\"{alignment}\"" + (leader is null ? "" : $" w:leader=\"{leader}\"") + "/>";
        string baseline = Stop("01800", leader: "dot") + Stop("3600");
        string higher = Stop("5400");
        string expected = baseline + higher;
        if (kind == "logical")
        {
            baseline = Stop("1800", "start") + Stop("3600", "END");
            expected = Stop("1800") + Stop("3600", "right") + higher;
        }
        if (kind is "clear-one" or "override-one")
        {
            baseline += Stop("5400");
            higher = kind == "clear-one" ? Stop("1800", "clear") : Stop("+3600", "center");
            expected = kind == "clear-one" ? Stop("3600") + Stop("5400") : Stop("01800", leader: "dot") + higher + Stop("5400");
        }
        const string spacing = "<w:spacing w:before=\"240\"/>";
        string defaults = kind == "defaults-add" ? "<w:docDefaults><w:pPrDefault><w:pPr>" + spacing +
            (flattened ? "" : "<w:tabs>" + baseline + "</w:tabs>") + "</w:pPr></w:pPrDefault></w:docDefaults>" : "";
        string baseTabs = kind == "defaults-add" ? higher : baseline;
        string styles = $"<w:styles xmlns:w=\"{w}\">" + defaults +
            "<w:style w:type=\"paragraph\" w:styleId=\"Base\"><w:name w:val=\"Base\"/><w:pPr>" + spacing +
            "<w:tabs>" + baseTabs + "</w:tabs></w:pPr></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Derived\"><w:name w:val=\"Derived\"/><w:basedOn w:val=\"Base\"/><w:pPr><w:tabs>" + higher +
            "</w:tabs></w:pPr></w:style></w:styles>";
        string properties = flattened ? spacing + "<w:tabs>" + expected + "</w:tabs>" : kind switch
        {
            "defaults-add" => "<w:pStyle w:val=\"Base\"/>",
            "style-add" or "logical" => "<w:pStyle w:val=\"Derived\"/>",
            _ => "<w:pStyle w:val=\"Base\"/><w:tabs>" + higher + "</w:tabs>"
        };
        string paragraph = "<w:p><w:pPr>" + properties + "</w:pPr><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:tab/><w:t>C</w:t><w:tab/><w:t>D</w:t></w:r></w:p>";
        string body = container switch
        {
            "cell" => "<w:tbl><w:tblPr><w:tblW w:w=\"9360\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"9360\"/></w:tblGrid><w:tr><w:tc>" + paragraph + "</w:tc></w:tr></w:tbl>",
            "header" => "<w:p/>",
            _ => paragraph
        };
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["_rels/.rels"] = $"<Relationships xmlns=\"{rel}\"><Relationship Id=\"main\" Type=\"{r}/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/_rels/document.xml.rels"] = $"<Relationships xmlns=\"{rel}\"><Relationship Id=\"styles\" Type=\"{r}/styles\" Target=\"styles.xml\"/>" +
                (container == "header" ? $"<Relationship Id=\"header\" Type=\"{r}/header\" Target=\"header.xml\"/>" : "") + "</Relationships>",
            ["word/styles.xml"] = styles,
            ["word/document.xml"] = $"<w:document xmlns:w=\"{w}\" xmlns:r=\"{r}\"><w:body>" + body + "<w:sectPr>" +
                (container == "header" ? "<w:headerReference w:type=\"default\" r:id=\"header\"/>" : "") +
                "<w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr></w:body></w:document>"
        };
        if (container == "header") parts["word/header.xml"] = $"<w:hdr xmlns:w=\"{w}\">{paragraph}</w:hdr>";
        using var stream = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(stream, default);
        return new DocxReader().Read(package, null, default, mode);
    }
}
