using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxBreakCompositionTests
{
    public static void PptxManualBreaksKeepAnchoredTextboxContent()
    {
        foreach (var (anchor, autofit) in new[] { ("b", false), ("ctr", false), ("t", true), ("b", true) })
            CheckPair(table: false, anchor, autofit, "ordinary");
    }

    public static void PptxManualBreaksKeepAnchoredTableContent()
    {
        foreach (string anchor in new[] { "b", "ctr" })
            CheckPair(table: true, anchor, autofit: true, "ordinary");
    }

    public static void PptxManualBreaksKeepBlankAndSplitRunFrames()
    {
        foreach (bool table in new[] { false, true })
        foreach (string variant in new[] { "crlf", "split", "blank", "mixed-size" })
            CheckPair(table, "b", autofit: false, variant);
    }

    public static void PptxManualBreaksPreserveAuthoredCrops()
    {
        foreach (bool table in new[] { false, true })
        {
            PdfPage element = Render(table, "b", autofit: false, "ordinary", literal: false, height: 24);
            PdfPage literal = Render(table, "b", autofit: false, "ordinary", literal: true, height: 24);
            TestAssert.Equal(element.Content, literal.Content);
            TestAssert.Equal("Alpha", string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(literal)
                .Select(show => show.Text)));
        }
    }

    public static void PptxBottomAnchorKeepsDescendersInsideFixedLineSpacingClip()
    {
        foreach (string spacing in new[] { "fixed-prose", "natural-prose" })
        {
            PdfPage top = Render(table: false, "t", autofit: true, spacing, literal: false, height: 96);
            PdfPage bottom = Render(table: false, "b", autofit: true, spacing, literal: false, height: 96);
            PdfRectangle frame = new(72, 372, 288, 96);
            var topInk = PortableRenderingTests.ReadPaintedSourceLabel(top, "Alpha Beta Gamma", requiredClip: frame);
            var bottomInk = PortableRenderingTests.ReadPaintedSourceLabel(bottom, "Alpha Beta Gamma", requiredClip: frame);
            TestAssert.Equal(14, bottomInk.Count);
            TestAssert.True(bottomInk.All(glyph => glyph.Size == 24d),
                "Bottom alignment preserves the authored font size.");
            TestAssert.True(topInk[0].Origin.Y - bottomInk[0].Origin.Y > 60d,
                "Bottom alignment still places the fitting line near the bottom of the same clip.");
        }
    }

    private static void CheckPair(bool table, string anchor, bool autofit, string variant)
    {
        double height = variant == "ordinary" ? 96 : 224;
        PdfPage element = Render(table, anchor, autofit, variant, literal: false, height);
        PdfPage literal = Render(table, anchor, autofit, variant, literal: true, height);
        var baselines = new List<double>();
        foreach (string word in new[] { "Alpha", "Beta", "Gamma" })
        {
            var expected = PortableRenderingTests.ReadPaintedSourceLabel(element, word);
            var actual = PortableRenderingTests.ReadPaintedSourceLabel(literal, word);
            baselines.Add(expected[0].Origin.Y);
            TestAssert.Equal(expected.Count, actual.Count);
            for (int i = 0; i < actual.Count; i++)
                TestAssert.True(System.Numerics.Vector2.Distance(expected[i].Origin, actual[i].Origin) < .002f &&
                    Math.Abs(expected[i].Size - actual[i].Size) < .001d,
                    "Equivalent breaks retain every source glyph's size and anchor: " + word);
        }

        if (variant == "mixed-size")
        {
            TestAssert.True(baselines[0] > baselines[1] && baselines[1] > baselines[2],
                "Mixed-size spans retain three separate lines in source order.");
        }
        else
        {
            double firstStep = variant == "blank" ? 48d : 24d;
            TestAssert.True(Math.Abs(baselines[0] - baselines[1] - firstStep) < .002d &&
                Math.Abs(baselines[1] - baselines[2] - 24d) < .002d,
                "The authored line grid includes the intermediate blank line when present.");
        }
    }

    private static PdfPage Render(bool table, string anchor, bool autofit, string variant, bool literal, double height)
    {
        string text = variant switch
        {
            "crlf" when literal => Run("Alpha&#xD;&#xA;Beta&#xD;&#xA;Gamma"),
            "split" when literal => Run("Al") + Run("pha&#xA;Be") + Run("ta&#xA;Gamma"),
            "split" => Run("Al") + Run("pha") + Break() + Run("Be") + Run("ta") + Break() + Run("Gamma"),
            "blank" when literal => Run("&#xA;Alpha&#xA;&#xA;Beta&#xA;Gamma&#xA;"),
            "blank" => Break() + Run("Alpha") + Break() + Break() + Run("Beta") + Break() + Run("Gamma") + Break(),
            "mixed-size" when literal => Run("Alpha&#xA;", 12) + Run("Beta&#xA;", 18) + Run("Gamma"),
            "mixed-size" => Run("Alpha", 12) + Break(12) + Run("Beta", 18) + Break(18) + Run("Gamma"),
            "fixed-prose" or "natural-prose" => Run("Alpha Beta Gamma"),
            _ when literal => Run("Alpha&#xA;Beta&#xA;Gamma"),
            _ => Run("Alpha") + Break() + Run("Beta") + Break() + Run("Gamma")
        };
        string spacing = variant == "natural-prose" ? "" : """<a:lnSpc><a:spcPts val="2400"/></a:lnSpc>""";
        string body = $$"""
            <a:bodyPr anchor="{{anchor}}" vertOverflow="clip" lIns="0" rIns="0" tIns="0" bIns="0">{{(autofit ? "<a:noAutofit/>" : "")}}</a:bodyPr><a:lstStyle/>
            <a:p><a:pPr>{{spacing}}<a:defRPr sz="2400"><a:latin typeface="DejaVu Sans"/></a:defRPr></a:pPr>{{text}}</a:p>
            """;
        string emuHeight = ((int)(height * 12700)).ToString(System.Globalization.CultureInfo.InvariantCulture);
        string content = table ? $$"""
            <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Break composition table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="{{emuHeight}}"/></p:xfrm>
            <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/>
            <a:tblGrid><a:gridCol w="3657600"/></a:tblGrid><a:tr h="{{emuHeight}}"><a:tc><a:txBody>{{body}}</a:txBody>
            <a:tcPr anchor="{{anchor}}" marL="0" marR="0" marT="0" marB="0"/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
            """ : $$"""
            <p:sp><p:nvSpPr><p:cNvPr id="2" name="Break composition"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="{{emuHeight}}"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
            <p:txBody>{{body}}</p:txBody></p:sp>
            """;
        var entries = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <p:cSld><p:spTree>{{content}}</p:spTree></p:cSld></p:sld>
                """
        };
        using var input = TestFixtures.CreateZipPackage(entries);
        using var package = OoxPackage.Open(input, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default)
            .ToBlockingEnumerable().Single();
        TestAssert.Equal(0, diagnostics.Count);
        return page;
    }

    private static string Run(string text, int size = 24) =>
        $"<a:r><a:rPr sz=\"{size * 100}\"><a:latin typeface=\"DejaVu Sans\"/></a:rPr><a:t>{text}</a:t></a:r>";

    private static string Break(int size = 24) =>
        $"<a:br><a:rPr sz=\"{size * 100}\"><a:latin typeface=\"DejaVu Sans\"/></a:rPr></a:br>";
}
