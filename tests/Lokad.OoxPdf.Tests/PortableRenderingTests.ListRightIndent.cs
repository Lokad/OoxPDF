using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingListRightIndentFollowsFormattingAcrossContainersAndBreaks()
    {
        ListIndentPaint[][] Paint(int? levelRight, int? styleRight = null, string directAttributes = "") =>
            PaintListIndents("w:left=\"720\"" + Attribute("right", levelRight), Attribute("right", styleRight), directAttributes);

        ListIndentPaint[][] zero = Paint(0);
        ListIndentPaint[][] tight = Paint(1440);
        ListIndentPaint[][] wide = Paint(-1440);
        for (int i = 0; i < ListIndentOwners.Length; i++)
        {
            TestAssert.True(tight[i].Length > zero[i].Length, ListIndentOwners[i] + " wraps within the tighter frame.");
            TestAssert.True(!wide[i].SequenceEqual(zero[i]), ListIndentOwners[i] + " responds to the signed inset.");
        }
        AssertListIndentPaint(tight, Paint(0, directAttributes: "w:right=\"1440\""), "Direct right overrides a level zero");
        AssertListIndentPaint(zero, Paint(1440, directAttributes: "w:right=\"0\""), "Explicit zero resets a level inset");
        AssertListIndentPaint(wide, Paint(0, 1440, "w:end=\"-1440\" w:right=\"1440\""), "Logical end owns the direct signed inset");
        AssertListIndentPaint(zero, Paint(0, 1440), "The numbered level wins over the paragraph style");
        AssertListIndentPaint(tight, Paint(null, 1440), "The paragraph style fills an absent level inset");
    }

    public static void PortableRenderingListLeftIndentFollowsFormattingAcrossContainersAndBreaks()
    {
        ListIndentPaint[][] zero = PaintListIndents("w:left=\"0\"");
        ListIndentPaint[][] tight = PaintListIndents("w:left=\"1440\"");
        for (int i = 0; i < ListIndentOwners.Length; i++)
            TestAssert.True(tight[i][0].X > zero[i][0].X + 40d, ListIndentOwners[i] + " responds to the level inset.");

        AssertListIndentPaint(zero, PaintListIndents("w:left=\"0\"", "w:left=\"1440\""),
            "A level zero overrides an inherited left inset");
        AssertListIndentPaint(tight, PaintListIndents("w:left=\"1440\"", "w:left=\"0\""),
            "An inherited zero does not replace a level inset");
        AssertListIndentPaint(zero, PaintListIndents("w:left=\"1440\"", directAttributes: "w:left=\"0\""),
            "Direct zero resets a level inset");
        AssertListIndentPaint(tight, PaintListIndents("w:left=\"0\"", directAttributes: "w:start=\"1440\" w:left=\"0\""),
            "Logical start owns the direct inset");
        AssertListIndentPaint(tight, PaintListIndents("", "w:left=\"1440\""),
            "The paragraph style fills an absent level inset");
    }

    private static void AssertListIndentPaint(ListIndentPaint[][] expected, ListIndentPaint[][] actual, string rule)
    {
        for (int i = 0; i < ListIndentOwners.Length; i++)
            TestAssert.True(expected[i].SequenceEqual(actual[i]), rule + " in " + ListIndentOwners[i] + ".");
    }

    private static string Attribute(string side, int? twips) => twips is { } value ? $" w:{side}=\"{value}\"" : "";

    private static readonly string[] ListIndentOwners = ["Body", "Cell", "Tail", "Footer"];
    private readonly record struct ListIndentPaint(double X, string Words);

    private static ListIndentPaint[][] PaintListIndents(string levelAttributes, string styleAttributes = "", string directAttributes = "")
    {
        string[] words = ["Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel", "India", "Juliet"];
        var renderer = new DocxRenderer(new PinnedResolver(), OoxPdfDocxMarkupMode.Final,
            OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);

        XNamespace w = OoxNamespaces.WordprocessingNamespace;
        XElement Paragraph(string owner) => new(w + "p", new XElement(w + "pPr",
            new XElement(w + "pStyle", new XAttribute(w + "val", "RightInset")),
            new XElement(w + "numPr", new XElement(w + "numId", new XAttribute(w + "val", "1"))),
            XElement.Parse($"""<w:ind xmlns:w="{w}" {directAttributes}/>""")), Run(owner));
        XElement Run(string owner) => new(w + "r", new XElement(w + "t", string.Join(' ', words.Select(word => owner + word))));

        XElement first = Paragraph("Body");
        first.Add(new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page"))), Run("Tail"));
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                  <Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                  <Relationship Id="footer" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = $"""
                <w:numbering xmlns:w="{w}">
                  <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="space"/>
                  <w:pPr><w:ind w:hanging="360" {levelAttributes}/></w:pPr></w:lvl></w:abstractNum>
                  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
                </w:numbering>
                """,
            ["word/styles.xml"] = $"""
                <w:styles xmlns:w="{w}"><w:style w:type="paragraph" w:styleId="RightInset"><w:pPr><w:ind {styleAttributes}/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:style></w:styles>
                """,
            ["word/footer1.xml"] = new XElement(w + "ftr", Paragraph("Footer")).ToString(),
            ["word/document.xml"] = $"""
                <w:document xmlns:w="{w}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
                  {first}
                  <w:tbl><w:tblPr><w:tblW w:w="5760" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
                  <w:tblGrid><w:gridCol w:w="5760"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5760" w:type="dxa"/></w:tcPr>{Paragraph("Cell")}</w:tc></w:tr></w:tbl>
                  <w:sectPr><w:footerReference w:type="default" r:id="footer"/><w:pgSz w:w="7200" w:h="8640"/>
                  <w:pgMar w:top="720" w:bottom="1440" w:left="720" w:right="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>
                </w:body></w:document>
                """
        });
        using var package = OoxPackage.Open(input, default);
        DocxDocument source = new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
        PdfPage[] pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
        TestAssert.Equal(2, pages.Length);

        var output = pages.Select(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
            .GroupBy(show => Math.Round(show.Y, 3))
            .Select(line => (X: line.First().X,
                Words: string.Concat(line.Select(show => show.Text)).Split(' ', StringSplitOptions.RemoveEmptyEntries)))
            .ToArray()).ToArray();
        return ListIndentOwners.Select(owner => Lines(owner)).ToArray();

        ListIndentPaint[] Lines(string owner)
        {
            string[] expected = words.Select(word => owner + word).ToArray();
            var lines = new List<ListIndentPaint>();
            for (int page = 0; page < output.Length; page++)
            {
                var selected = output[page].Select(line => (line.X, Words: line.Words
                        .Where(word => word.StartsWith(owner, StringComparison.Ordinal)).ToArray()))
                    .Where(line => line.Words.Length != 0).ToArray();
                bool shouldPaint = owner == "Footer" || page == (owner == "Body" ? 0 : 1);
                TestAssert.True(selected.SelectMany(line => line.Words).SequenceEqual(shouldPaint ? expected : []),
                    $"All {owner} words paint in order on their owning page. Page {page + 1}.");
                lines.AddRange(selected.Select(line => new ListIndentPaint(Math.Round(line.X, 3), string.Join(' ', line.Words))));
            }
            return lines.ToArray();
        }
    }
}
