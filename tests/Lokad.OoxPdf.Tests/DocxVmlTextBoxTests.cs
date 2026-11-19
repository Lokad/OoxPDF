using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxVmlTextBoxTests
{
    private static readonly XNamespace W = OoxNamespaces.WordprocessingNamespace;
    private static readonly XNamespace V = OoxNamespaces.VmlNamespace;

    public static void DocxVmlTextBoxesShareHiddenAndRevisionSelection()
    {
        foreach (bool vml in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original })
        {
            var paragraph = new XElement(W + "p", Text("VISIBLE "),
                Text("HIDDEN ", new XElement(W + "rPr", new XElement(W + "vanish"))),
                new XElement(W + "del", new XAttribute(W + "id", 1), new XAttribute(W + "author", "A"),
                    new XElement(W + "r", new XElement(W + "delText", new XAttribute(XNamespace.Xml + "space", "preserve"), "DELETED "))),
                new XElement(W + "ins", new XAttribute(W + "id", 2), new XAttribute(W + "author", "A"), Text("INSERTED ")));
            With(vml ? new XElement(W + "p", Box(paragraph)) : paragraph, mode, null, document =>
                TestAssert.Equal(mode == OoxPdfDocxMarkupMode.Final ? "VISIBLE INSERTED" : "VISIBLE DELETED", RenderText(document, mode)));
        }
    }

    public static void DocxVmlTextBoxesKeepPicturesBetweenSurroundingRunText()
    {
        var box = Box(new XElement(W + "p", Text("BEFORE "), Picture(), Text(" AFTER")));
        box.AddFirst(new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "OUTER-BEFORE "));
        box.Add(new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " OUTER-AFTER"));
        With(new XElement(W + "p", box), OoxPdfDocxMarkupMode.Final, null, document =>
        {
            var paragraph = document.Paragraphs.Single();
            TestAssert.Equal(1, paragraph.Images.Count);
            var items = TestAssert.NotNull(paragraph.InlineContent);
            int picture = items.Select((item, index) => (item, index)).Single(pair => pair.item is DocxInlinePicture).index;
            TestAssert.True(string.Concat(items.Take(picture).OfType<DocxInlineText>().Select(item => item.Run.Text)).Contains("OUTER-BEFORE BEFORE", StringComparison.Ordinal),
                "Surrounding and inner prefix text precede the picture.");
            TestAssert.True(string.Concat(items.Skip(picture + 1).OfType<DocxInlineText>().Select(item => item.Run.Text)).Contains("AFTER OUTER-AFTER", StringComparison.Ordinal),
                "Inner and surrounding suffix text follow the picture.");
            var pages = Render(document, OoxPdfDocxMarkupMode.Final);
            TestAssert.Equal(1, pages.Sum(page => page.Images.Count));
            TestAssert.Equal("OUTER-BEFORE BEFORE AFTER OUTER-AFTER", Normalize(string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText))));
        });
    }

    public static void DocxVmlTextBoxesOwnInnerStylesWithoutLeakingToSiblings()
    {
        var paragraph = new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", "Outer"))),
            Text("BEFORE "), Box(new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", "Inner"))),
                Text("INNER "), Text("HIDDEN", new XElement(W + "rPr", new XElement(W + "rStyle", new XAttribute(W + "val", "Hidden")))))), Text(" AFTER"));
        var styles = new XElement(W + "styles",
            Style("Outer", 48, "FF0000"), Style("Inner", 24, "0000FF"),
            new XElement(W + "style", new XAttribute(W + "type", "character"), new XAttribute(W + "styleId", "Hidden"),
                new XElement(W + "rPr", new XElement(W + "vanish"))));
        With(paragraph, OoxPdfDocxMarkupMode.Final, styles, document =>
        {
            var runs = document.Paragraphs.Single().Runs;
            TestAssert.Equal(24d, runs.Single(run => run.Text.Contains("BEFORE", StringComparison.Ordinal)).FontSize);
            TestAssert.Equal(12d, runs.Single(run => run.Text.Contains("INNER", StringComparison.Ordinal)).FontSize);
            TestAssert.Equal(24d, runs.Single(run => run.Text.Contains("AFTER", StringComparison.Ordinal)).FontSize);
            TestAssert.True(runs.Single(run => run.Text == "HIDDEN").Hidden, "The inner character style shares ordinary hidden-text resolution.");
            TestAssert.Equal("BEFORE INNER AFTER", RenderText(document, OoxPdfDocxMarkupMode.Final));
        });
    }

    public static void DocxVmlTextBoxesKeepBlockWrappedParagraphsAndNavigation()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            XElement[] paragraphs = [
                new(W + "p", Text("STORED TERMS "),
                    new XElement(W + "hyperlink", new XAttribute(W + "anchor", "Target"), Text("REFERENCE LINK"))),
                new(W + "p", Text("SECOND PARAGRAPH "), Picture(), Text(" AFTER PICTURE"))
            ];
            XElement[][] variants = [
                paragraphs,
                [new(W + "sdt", new XElement(W + "sdtPr", new XElement(W + "id", new XAttribute(W + "val", 123))),
                    new XElement(W + "sdtContent", paragraphs))],
                [new(W + "customXml", new XAttribute(W + "uri", "urn:ooxpdf:terms"),
                    new XAttribute(W + "element", "terms"), paragraphs)]
            ];
            PdfPage? baseline = null;
            foreach (var blocks in variants)
            {
                var paragraph = new XElement(W + "p", Text("BEFORE BOX "), Box(blocks),
                    new XElement(W + "bookmarkStart", new XAttribute(W + "id", 1), new XAttribute(W + "name", "Target")),
                    Text(" AFTER BOX"), new XElement(W + "bookmarkEnd", new XAttribute(W + "id", 1)));
                With(paragraph, mode, null, document =>
                {
                    PdfPage page = Render(document, mode).Single();
                    const string expected = "BEFORE BOX STORED TERMS REFERENCE LINK SECOND PARAGRAPH AFTER PICTURE AFTER BOX";
                    string text = DocxPageFootnoteNumberingTests.DecodeText(page);
                    TestAssert.Equal(expected.Replace(" ", ""), string.Concat(text.Where(character => !char.IsWhiteSpace(character))));
                    TestAssert.Equal(1, page.Images.Count);
                    var link = page.Annotations.Single();
                    TestAssert.True(link.IsDestination, "The stored link must resolve the surrounding bookmark.");
                    TestAssert.Equal(0, link.Destination!.Value.PageIndex);
                    TestAssert.True(link.Width > 0 && link.Height > 0, "The stored link must have a visible source rectangle.");
                    if (baseline is { } direct)
                    {
                        TestAssert.Equal(direct.Content, page.Content);
                        TestAssert.Equal(direct.Annotations.Single(), link);
                    }
                    else baseline = page;
                });
            }
        }
    }

    public static void DocxVmlTextBoxesShareDynamicAndLockedFieldResults()
    {
        var inner = new XElement(W + "p", Text("PAGE:"),
            new XElement(W + "fldSimple", new XAttribute(W + "instr", "PAGE"), Text("99")), Text(" LOCK:"),
            new XElement(W + "fldSimple", new XAttribute(W + "instr", "PAGE"), new XAttribute(W + "fldLock", "1"), Text("92")));
        With(new XElement(W + "p", Box(inner)), OoxPdfDocxMarkupMode.Final, null, document =>
        {
            TestAssert.Equal(2, document.Paragraphs.Single().FieldReferences.Count);
            TestAssert.Equal("PAGE:1 LOCK:92", RenderText(document, OoxPdfDocxMarkupMode.Final));
        });
    }

    public static void DocxVmlTextBoxesKeepBlockBookmarkDestinations()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            PdfPage? baseline = null;
            foreach (string placement in new[] { "inline", "block", "wrapped", "closed", "skipped-table" })
            {
                var start = new XElement(W + "bookmarkStart", new XAttribute(W + "id", 1), new XAttribute(W + "name", "BoxTarget"));
                var end = new XElement(W + "bookmarkEnd", new XAttribute(W + "id", 1));
                var target = new XElement(W + "p", Text("TARGET "), Picture(), Text(" AFTER"));
                XElement[] blocks = placement switch
                {
                    "inline" => [new(W + "p", start, target.Elements(), end)],
                    "block" => [start, target, end],
                    "wrapped" => [new(W + "sdt", new XElement(W + "sdtContent", start, target, end))],
                    "closed" => [start, end, target],
                    _ => [start, new(W + "tbl", new XElement(W + "tr", new XElement(W + "tc",
                        new XElement(W + "p", Text("TABLE"))))), target, end]
                };
                var paragraph = new XElement(W + "p", Box([new(W + "p", Text("FIRST")), .. blocks]),
                    new XElement(W + "hyperlink", new XAttribute(W + "anchor", "BoxTarget"), Text(" JUMP")));
                With(paragraph, mode, null, document =>
                {
                    var page = Render(document, mode).Single();
                    TestAssert.Equal(1, page.Images.Count);
                    if (baseline is { } direct) TestAssert.Equal(direct.Content, page.Content);
                    else baseline = page;
                    if (placement is "closed" or "skipped-table")
                    {
                        TestAssert.Equal(0, page.Annotations.Count);
                    }
                    else
                    {
                        var link = page.Annotations.Single();
                        TestAssert.True(link.IsDestination, "The link must reach the placed second paragraph.");
                        TestAssert.Equal(baseline.Value.Annotations.Single(), link);
                    }
                });
            }
        }
    }

    public static void DocxVmlHiddenDrawingRunSuppressesItsTextAndPictures()
    {
        var box = Box(new XElement(W + "p", Text("HIDDEN"), Picture()));
        box.AddFirst(new XElement(W + "rPr", new XElement(W + "vanish")));
        With(new XElement(W + "p", Text("BEFORE "), box, Text("AFTER")), OoxPdfDocxMarkupMode.Final, null, document =>
        {
            TestAssert.Equal(0, document.Paragraphs.Single().Images.Count);
            TestAssert.Equal("BEFORE AFTER", RenderText(document, OoxPdfDocxMarkupMode.Final));
        });
    }

    private static XElement Text(string text, XElement? properties = null) => new(W + "r", properties,
        new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));

    private static XElement Box(params XElement[] blocks) => new(W + "r", new XElement(W + "pict",
        new XElement(V + "shape", new XAttribute("id", "box"), new XAttribute("style", "width:300pt;height:100pt"),
            new XElement(V + "textbox", new XElement(W + "txbxContent", blocks)))));

    private static XElement Style(string id, int halfPoints, string color) => new(W + "style",
        new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", id), new XElement(W + "rPr",
            new XElement(W + "sz", new XAttribute(W + "val", halfPoints)), new XElement(W + "color", new XAttribute(W + "val", color))));

    private static XElement Picture() => XElement.Parse("""
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:drawing><wp:inline><wp:extent cx="152400" cy="152400"/><wp:docPr id="1" name="Picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="152400" cy="152400"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic>
          </a:graphicData></a:graphic></wp:inline></w:drawing>
        </w:r>
        """);

    private static void With(XElement paragraph, OoxPdfDocxMarkupMode mode, XElement? styles, Action<DocxDocument> action)
    {
        var parts = new Dictionary<string, byte[]> {
            ["[Content_Types].xml"] = TestFixtures.Utf8("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"png\" ContentType=\"image/png\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>"),
            ["_rels/.rels"] = TestFixtures.Utf8("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"),
            ["word/document.xml"] = TestFixtures.Utf8(new XElement(W + "document", new XElement(W + "body", paragraph)).ToString()),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"image\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image.png\"/><Relationship Id=\"styles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>"),
            ["word/styles.xml"] = TestFixtures.Utf8((styles ?? new XElement(W + "styles")).ToString()),
            ["word/media/image.png"] = TestFixtures.CreateTinyStoredRgbPng()
        };
        using var input = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(input, default);
        var xml = package.LoadXml(TestAssert.NotNull(package.GetPart("/word/document.xml")), default);
        string original = xml.ToString();
        action(new DocxReader().Read(package, null, default, mode));
        TestAssert.Equal(original, xml.ToString());
    }

    private static IReadOnlyList<PdfPage> Render(DocxDocument document, OoxPdfDocxMarkupMode mode) =>
        new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();

    private static string RenderText(DocxDocument document, OoxPdfDocxMarkupMode mode) =>
        Normalize(string.Concat(Render(document, mode).Select(DocxPageFootnoteNumberingTests.DecodeText)));

    private static string Normalize(string text) => string.Join(' ', text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
