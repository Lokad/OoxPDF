using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxGlyphlessAnchorTests
{
    public static void DocxGlyphlessFloatingImagePaintsWithoutFontPrograms()
    {
        string input = CreateDocument();
        foreach (var (mode, geometry) in Views())
        {
            PdfPage page = Render(input, new EmptyProgramResolver(), mode, geometry).Single();
            TestAssert.Equal(0, page.Fonts.Count);
            double[] frame = PaintedFrame(page);
            double[] expected = [144d, 72d, 234d, page.Height - 149d];
            TestAssert.True(frame.Zip(expected).All(p => Math.Abs(p.First - p.Second) < .001d), "Fontless drawing placement keeps the source extent, centered column and5pt paragraph offset.");
        }
    }

    public static void DocxGlyphlessFloatingFramesMatchAnExplicitFontControl()
    {
        string input = CreateDocument();
        foreach (var (mode, geometry) in Views())
        {
            double[] missing = PaintedFrame(Render(input, new EmptyProgramResolver(), mode, geometry).Single());
            double[] available = PaintedFrame(Render(input, new PortableRenderingTests.PinnedResolver(), mode, geometry).Single());
            TestAssert.True(missing.Zip(available).All(p => Math.Abs(p.First - p.Second) < .001d), "An unavailable font does not change a glyphless floating picture's physical frame.");
        }
    }

    public static void DocxGlyphlessTablePicturesSurviveNestedCellsAndHeaders()
    {
        byte[] png = TestFixtures.CreateRgbPng(1, 1, [0, 255, 0]);
        string nested = TestFixtures.WriteDocxInlinePng($"""
            <w:tbl>
              <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
              <w:tr><w:tc>
                <w:tbl>
                  <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                  <w:tr><w:tc><w:p>{TestFixtures.DocxInlinePngRun}</w:p></w:tc></w:tr>
                </w:tbl>
                <w:p/>
              </w:tc></w:tr>
            </w:tbl>
            """, png);
        XNamespace w = OoxNamespaces.WordprocessingNamespace;
        XNamespace r = OoxNamespaces.RelationshipsNamespace;
        XElement source = XElement.Parse(System.Text.Encoding.UTF8.GetString(ReadParts(nested)["word/document.xml"]));
        XElement body = source.Element(w + "body")!;
        XElement table = body.Element(w + "tbl")!;
        var header = new XElement(w + "hdr", source.Attributes(), table.Descendants(w + "tbl").Single());
        table.ReplaceWith(new XElement(w + "p"));
        XElement section = body.Element(w + "sectPr")!;
        section.AddFirst(new XElement(w + "headerReference",
            new XAttribute(w + "type", "default"), new XAttribute(r + "id", "rIdHeader1")));
        section.Add(new XElement(w + "pgMar", new XAttribute(w + "top", "2880"), new XAttribute(w + "header", "360")));
        string staticCell = TestFixtures.WriteDocxHeaderImage(source.ToString(SaveOptions.DisableFormatting),
            header.ToString(SaveOptions.DisableFormatting), png);

        foreach (var (mode, geometry) in Views())
        foreach (var (input, expectedClips) in new[] { (nested, 2), (staticCell, 1) })
        {
            PdfPage page = Render(input, new EmptyProgramResolver(), mode, geometry).Single();
            TestAssert.Equal(0, page.Fonts.Count);
            double[] frame = PaintedFrame(page);
            TestAssert.Equal(72d, frame[0]);
            TestAssert.Equal(72d, frame[1]);
            MatchCollection clips = Regex.Matches(page.Content,
                @"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re W n");
            TestAssert.Equal(expectedClips, clips.Count);
            foreach (Match clip in clips)
            {
                double x = Coordinate(1), y = Coordinate(2);
                TestAssert.True(frame[2] >= x - .001d && frame[3] >= y - .001d &&
                    frame[2] + frame[0] <= x + Coordinate(3) + .001d &&
                    frame[3] + frame[1] <= y + Coordinate(4) + .001d,
                    "Every containing cell clip must retain the full picture.");

                double Coordinate(int index) => double.Parse(clip.Groups[index].Value, CultureInfo.InvariantCulture);
            }
        }
    }

    public static void DocxGlyphlessInlineBoxPicturesSurviveWithoutFontPrograms()
    {
        byte[] png = TestFixtures.CreateRgbPng(1, 1, [0, 255, 0]);
        string box = PictureBoxParagraph(floating: false, height: 90);
        string body = TestFixtures.WriteDocxInlinePng(box, png);
        string cell = TestFixtures.WriteDocxInlinePng($"""
            <w:tbl>
              <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
              <w:tr><w:tc>{box}</w:tc></w:tr>
            </w:tbl>
            """, png);
        foreach (var (mode, geometry) in Views())
        foreach (string input in new[] { body, cell })
        {
            ReadGreenPictureFrame(Render(input, new EmptyProgramResolver(), mode, geometry).Single());
        }
    }

    public static void DocxGlyphlessFloatingBoxPicturesKeepTheirViewport()
    {
        byte[] png = TestFixtures.CreateRgbPng(1, 1, [0, 255, 0]);
        var controls = new[] { 96, 24 }.Select(height =>
            (Input: TestFixtures.WriteDocxInlinePng(PictureBoxParagraph(floating: true, height), png), Height: height)).ToArray();
        foreach (var (mode, geometry) in Views())
        foreach (var (input, height) in controls)
        {
            PdfPage page = Render(input, new EmptyProgramResolver(), mode, geometry).Single();
            double[] frame = ReadGreenPictureFrame(page);
            TestAssert.Equal(145d, frame[2]);
            TestAssert.Equal(page.Height - 217d, frame[3]);
            string clip = FormattableString.Invariant($"145 {page.Height - 143d - height} 94 {height - 2} re W n");
            TestAssert.Contains(clip, page.Content);
        }
    }

    public static void DocxGlyphlessInlinePicturesAndBoxesFollowSourceOrder()
    {
        var parts = ReadParts(TestFixtures.WriteDocxInlinePng(PictureBoxParagraph(floating: false, height: 90),
            TestFixtures.CreateRgbPng(1, 1, [0, 255, 0])));
        XNamespace w = OoxNamespaces.WordprocessingNamespace;
        XNamespace a = OoxNamespaces.DrawingNamespace;
        XNamespace r = OoxNamespaces.RelationshipsNamespace;
        XNamespace rel = OoxNamespaces.PackageRelationshipsNamespace;
        const string documentPart = "word/document.xml";
        const string relationshipsPart = "word/_rels/document.xml.rels";
        XElement relationships = XElement.Parse(System.Text.Encoding.UTF8.GetString(parts[relationshipsPart]));
        relationships.Add(new XElement(rel + "Relationship",
            new XAttribute("Id", "rIdBlue"), new XAttribute("Type", r.NamespaceName + "/image"),
            new XAttribute("Target", "media/blue.png")));
        parts[relationshipsPart] = TestFixtures.Utf8(relationships.ToString(SaveOptions.DisableFormatting));
        parts["word/media/blue.png"] = TestFixtures.CreateRgbPng(1, 1, [0, 0, 255]);
        XElement seed = XElement.Parse(System.Text.Encoding.UTF8.GetString(parts[documentPart]));
        foreach (var (boxFirst, width) in new[] { (true, 0), (false, 0), (true, 4000), (false, 4000), (true, 2800) })
        {
            var document = new XElement(seed);
            var host = document.Element(w + "body")!.Element(w + "p")!;
            var box = host.Element(w + "r")!;
            var picture = new XElement(box.Descendants(w + "txbxContent").Single().Element(w + "p")!.Element(w + "r")!);
            picture.Descendants(a + "blip").Single().SetAttributeValue(r + "embed", "rIdBlue");
            if (boxFirst) box.AddAfterSelf(picture);
            else box.AddBeforeSelf(picture);
            if (width != 0)
                host.ReplaceWith(new XElement(w + "tbl",
                    new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", width))),
                    new XElement(w + "tr", new XElement(w + "tc", host))));
            parts[documentPart] = TestFixtures.Utf8(document.ToString(SaveOptions.DisableFormatting));
            string input = TestFixtures.WriteTempPackage(".docx", parts);
            foreach (var (mode, geometry) in Views())
            {
                PdfPage page = Render(input, new EmptyProgramResolver(), mode, geometry).Single();
                TestAssert.Equal(0, page.Fonts.Count);
                var frames = new Dictionary<int, double[]>();
                foreach (PdfImageResource image in page.Images)
                {
                    using var compressed = new MemoryStream(image.Image.Bytes);
                    using var rgb = new ZLibStream(compressed, CompressionMode.Decompress);
                    TestAssert.Equal(0, rgb.ReadByte());
                    int green = rgb.ReadByte(), blue = rgb.ReadByte();
                    TestAssert.True((green, blue) is (255, 0) or (0, 255), "Each authored picture color is distinct.");
                    frames.Add(green, PaintedFrame(page, image));
                }
                double[] boxed = frames[255], direct = frames[0];
                TestAssert.Equal(72d, boxed[0]);
                TestAssert.Equal(72d, direct[0]);
                if (width == 2800)
                {
                    TestAssert.True(boxed[3] >= direct[3] + direct[1], "The later picture wraps below the source-first box.");
                }
                else
                {
                    TestAssert.True(Math.Abs(boxed[3] - direct[3]) < 72d, "Both pictures fit the same line.");
                    TestAssert.True(boxFirst ? boxed[2] + boxed[0] <= direct[2] : direct[2] + direct[0] <= boxed[2],
                        "The pictured objects follow the authored run order.");
                }
            }
        }
    }

    private static string PictureBoxParagraph(bool floating, int height)
    {
        string placement = floating ? "anchor" : "inline";
        string attributes = floating ? """ simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" """ : "";
        string position = floating ? """
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>
            """ : "";
        return $"""
            <w:p><w:r><w:drawing><wp:{placement}{attributes}>
              {position}<wp:extent cx="1219200" cy="{height * 12700}"/>
              {(floating ? "<wp:wrapNone/>" : "")}<wp:docPr id="1" name="Picture box"/>
              <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                  <wps:txbx><w:txbxContent><w:p>{TestFixtures.DocxInlinePngRun}</w:p></w:txbxContent></wps:txbx>
                  <wps:bodyPr lIns="12700" tIns="12700" rIns="12700" bIns="12700"/>
                </wps:wsp>
              </a:graphicData></a:graphic>
            </wp:{placement}></w:drawing></w:r></w:p>
            """;
    }

    private static double[] ReadGreenPictureFrame(PdfPage page)
    {
        TestAssert.Equal(0, page.Fonts.Count);
        double[] frame = PaintedFrame(page);
        TestAssert.Equal(72d, frame[0]);
        TestAssert.Equal(72d, frame[1]);
        using var compressed = new MemoryStream(page.Images.Single().Image.Bytes);
        using var rgb = new ZLibStream(compressed, CompressionMode.Decompress);
        TestAssert.Equal(0, rgb.ReadByte());
        TestAssert.Equal(255, rgb.ReadByte());
        TestAssert.Equal(0, rgb.ReadByte());
        TestAssert.Equal(-1, rgb.ReadByte());
        return frame;
    }

    public static void DocxGlyphlessAnchorsFollowBlankSlotsAndPageBreaks()
    {
        var parts = ReadParts(CreateDocument());
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XDocument xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["word/document.xml"]));
        XElement anchor = xml.Descendants(w + "p").First();
        XElement Blank() => new(w + "p", new XElement(w + "pPr", new XElement(w + "spacing",
            new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"), new XAttribute(w + "line", "400"), new XAttribute(w + "lineRule", "exact"))));
        anchor.AddBeforeSelf(Blank(), new XElement(w + "p", new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page")))), Blank(), Blank());
        parts["word/document.xml"] = TestFixtures.Utf8(xml.ToString(SaveOptions.DisableFormatting));
        string input = TestFixtures.WriteTempPackage(".docx", parts);
        PdfPage[] pages = Render(input, new EmptyProgramResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        TestAssert.Equal(2, pages.Length);
        TestAssert.Equal(0, pages[0].Images.Count);
        double[] frame = PaintedFrame(pages[1]);
        TestAssert.True(Math.Abs(pages[1].Height - frame[3] - frame[1] - 117d) < .001d, "Two authored20pt blank slots move the paragraph anchor40pt on its own page.");
        TestAssert.True(pages.All(p => p.Fonts.Count == 0), "Empty slots do not import a font program.");
    }

    public static void DocxFloatingImagesFollowTheirEmittedFlowBlocks()
    {
        var parts = ReadParts(CreateDocument());
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        string original = System.Text.Encoding.UTF8.GetString(parts["word/document.xml"]);
        foreach (string source in new[] { "page-break-before", "content-control", "run-break" })
        {
            XDocument xml = XDocument.Parse(original);
            XElement anchor = xml.Descendants(w + "p").Single();
            XElement authoredBreak = new(w + "br", new XAttribute(w + "type", "page"));
            if (source == "run-break")
            {
                anchor.Element(w + "r")!.AddFirst(new XElement(w + "t", "First page"), authoredBreak);
            }
            else
            {
                anchor.AddBeforeSelf(new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "First page"))));
                if (source == "page-break-before")
                {
                    anchor.AddFirst(new XElement(w + "pPr", new XElement(w + "pageBreakBefore")));
                }
                else
                {
                    anchor.AddBeforeSelf(new XElement(w + "p", new XElement(w + "r", authoredBreak)));
                    var control = new XElement(w + "sdt", new XElement(w + "sdtPr"),
                        new XElement(w + "sdtContent", new XElement(anchor)));
                    anchor.ReplaceWith(control);
                }
            }

            parts["word/document.xml"] = TestFixtures.Utf8(xml.ToString(SaveOptions.DisableFormatting));
            string input = TestFixtures.WriteTempPackage(".docx", parts);
            foreach (var (mode, geometry) in Views())
            {
                PdfPage[] pages = Render(input, new PortableRenderingTests.PinnedResolver(), mode, geometry);
                TestAssert.True(pages.Length == 2 && pages[0].Images.Count == 0 && pages[1].Images.Count == 1,
                    source + " must keep the picture on the second page with the owning paragraph.");
                double[] frame = PaintedFrame(pages[1]);
                TestAssert.Equal((144d, 72d), (frame[0], frame[1]));
            }
        }
    }

    public static void DocxFloatingPictureLinksFollowPaintedViewports()
    {
        const string uri = "https://example.invalid/floating-picture";
        var parts = ReadParts(CreateDocument());
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XDocument rels = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));
        rels.Root!.Add(new XElement(rels.Root.Name.Namespace + "Relationship",
            new XAttribute("Id", "rIdPictureLink"), new XAttribute("Type", r.NamespaceName + "/hyperlink"),
            new XAttribute("Target", uri), new XAttribute("TargetMode", "External")));
        parts["word/_rels/document.xml.rels"] = TestFixtures.Utf8(rels.ToString(SaveOptions.DisableFormatting));
        XDocument xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["word/document.xml"]));
        XElement anchor = xml.Descendants(wp + "anchor").Single();
        anchor.Element(wp + "wrapSquare")!.AddAfterSelf(new XElement(wp + "docPr", new XAttribute("id", "1"), new XAttribute("name", "Linked picture"),
            new XElement(a + "hlinkClick", new XAttribute(r + "id", "rIdPictureLink"))));
        XElement horizontal = new(wp + "posOffset", "914400");
        anchor.Element(wp + "positionH")!.ReplaceWith(new XElement(wp + "positionH", new XAttribute("relativeFrom", "page"), horizontal));
        anchor.Element(wp + "positionV")!.ReplaceWith(new XElement(wp + "positionV",
            new XAttribute("relativeFrom", "page"), new XElement(wp + "posOffset", "914400")));
        XElement transform = new(a + "xfrm", new XElement(a + "ext", new XAttribute("cx", "1828800"), new XAttribute("cy", "914400")));
        anchor.Descendants(pic + "pic").Single().Add(new XElement(pic + "spPr", transform));
        foreach (var (rotation, offset, missing, expected) in new (int, int, bool, PdfRectangle?)[]
        {
            (0, 914400, false, new PdfRectangle(72d, 648d, 144d, 72d)),
            (90, 914400, false, new PdfRectangle(108d, 612d, 72d, 144d)),
            (30, 914400, false, null),
            (0, 6985000, false, null),
            (0, 914400, true, null)
        })
        {
            horizontal.Value = offset.ToString(CultureInfo.InvariantCulture);
            transform.SetAttributeValue("rot", rotation * 60000);
            parts["word/document.xml"] = TestFixtures.Utf8(xml.ToString(SaveOptions.DisableFormatting));
            if (missing) parts.Remove("word/media/image1.png");

            PdfPage page = Render(TestFixtures.WriteTempPackage(".docx", parts), new EmptyProgramResolver(),
                OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Single();
            TestAssert.Equal(missing ? 0 : 1, page.Images.Count);
            if (expected is { } frame)
            {
                var link = page.Annotations.Single();
                TestAssert.Equal(uri, link.Uri);
                TestAssert.Equal(frame, new PdfRectangle(link.X, link.Y, link.Width, link.Height));
            }
            else
            {
                TestAssert.True(page.Annotations.Count == 0, "A URI requires successful paint and a complete rectangular viewport.");
            }
        }
    }

    public static void DocxRotatedInlinePictureLinksFollowPaintedViewports()
    {
        const string uri = "https://example.invalid/inline-picture";
        var parts = ReadParts(CreateDocument());
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace wp = OoxNamespaces.WordprocessingDrawingNamespace;
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XDocument relationships = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));
        relationships.Root!.Add(new XElement(relationships.Root.Name.Namespace + "Relationship",
            new XAttribute("Id", "rIdPictureLink"), new XAttribute("Type", r.NamespaceName + "/hyperlink"),
            new XAttribute("Target", uri), new XAttribute("TargetMode", "External")));
        parts["word/_rels/document.xml.rels"] = TestFixtures.Utf8(relationships.ToString(SaveOptions.DisableFormatting));
        XDocument source = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["word/document.xml"]));
        XElement graphic = source.Descendants(a + "graphic").Single();
        var resolver = new PortableRenderingTests.PinnedResolver();

        foreach (var (rotation, cell, exactHeight, merged, expected) in new[]
        {
            (0, false, 0, false, true),
            (90, false, 0, false, true),
            (180, false, 0, false, true),
            (90, true, 0, false, true),
            (180, true, 0, false, true),
            (90, false, 420, false, false),
            (90, true, 780, false, false),
            (90, true, 0, true, false),
            (30, false, 0, false, false)
        })
        {
            const string width = "304800"; // 24 pt
            const string height = "228600"; // 18 pt
            bool clipped = exactHeight > 0;
            bool exactLine = clipped && !cell;
            var transform = new XElement(a + "xfrm", new XAttribute("rot", rotation * 60000),
                new XElement(a + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                new XElement(a + "ext", new XAttribute("cx", width), new XAttribute("cy", height)));
            var inline = new XElement(wp + "inline",
                new XElement(wp + "extent", new XAttribute("cx", width), new XAttribute("cy", height)),
                new XElement(wp + "docPr", new XAttribute("id", "1"), new XAttribute("name", "Linked inline picture"),
                    new XElement(a + "hlinkClick", new XAttribute(r + "id", "rIdPictureLink"))),
                new XElement(graphic));
            inline.Descendants(pic + "pic").Single().Add(new XElement(pic + "spPr", transform));
            var paragraph = new XElement(w + "p",
                new XElement(w + "pPr", new XElement(w + "spacing",
                    new XAttribute(w + "line", exactLine ? exactHeight : 960),
                    new XAttribute(w + "lineRule", exactLine ? "exact" : "atLeast"))),
                new XElement(w + "r", new XElement(w + "drawing", inline)),
                new XElement(w + "r", new XElement(w + "t", "Inline picture")));
            var document = new XDocument(source);
            XElement body = document.Descendants(w + "body").Single();
            if (cell)
            {
                var row = new XElement(w + "tr",
                    clipped ? new XElement(w + "trPr", new XElement(w + "trHeight",
                        new XAttribute(w + "val", exactHeight), new XAttribute(w + "hRule", "exact"))) : null,
                    new XElement(w + "tc", merged ? new XElement(w + "tcPr",
                        new XElement(w + "vMerge", new XAttribute(w + "val", "restart"))) : null, paragraph));
                var table = new XElement(w + "tbl",
                    new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "7200"))), row);
                if (merged)
                    table.Add(new XElement(w + "tr", new XElement(w + "tc",
                        new XElement(w + "tcPr", new XElement(w + "vMerge")), new XElement(w + "p"))));
                body.Element(w + "p")!.ReplaceWith(table);
            }
            else
                body.Element(w + "p")!.ReplaceWith(paragraph);
            parts["word/document.xml"] = TestFixtures.Utf8(document.ToString(SaveOptions.DisableFormatting));
            PdfPage page = Render(TestFixtures.WriteTempPackage(".docx", parts), resolver,
                OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Single();
            string context = $"rotation={rotation}, cell={cell}, exactHeight={exactHeight}, merged={merged}";
            TestAssert.True(page.Images.Count == 1, "The fixture paints one image: " + context);
            string imagePaint = "/" + page.Images.Single().ResourceName + " Do";
            var links = page.Annotations.Where(link => link.Uri == uri).ToArray();
            if (rotation == 30)
            {
                TestAssert.Contains(imagePaint, page.Content);
                TestAssert.Equal(0, links.Length);
                continue;
            }

            var paint = DocxNoteClippingTests.ReadPhysicalPaintFrames(page).Single(frame => frame.Block == imagePaint);
            var matrix = paint.Transform;
            var frame = new PdfRectangle(
                matrix.M31 + Math.Min(0f, matrix.M11) + Math.Min(0f, matrix.M21),
                matrix.M32 + Math.Min(0f, matrix.M12) + Math.Min(0f, matrix.M22),
                Math.Abs(matrix.M11) + Math.Abs(matrix.M21), Math.Abs(matrix.M12) + Math.Abs(matrix.M22));
            TestAssert.True(rotation switch
            {
                90 => matrix.M12 < -23d && matrix.M21 > 17d,
                180 => matrix.M11 < -23d && matrix.M22 < -17d,
                _ => matrix.M11 > 23d && matrix.M22 > 17d
            }, "The picture paints with the authored orientation: " + context);
            bool containsFrame = paint.Clip.X <= frame.X + .002d && paint.Clip.Y <= frame.Y + .002d &&
                paint.Clip.X + paint.Clip.Width >= frame.X + frame.Width - .002d &&
                paint.Clip.Y + paint.Clip.Height >= frame.Y + frame.Height - .002d;
            TestAssert.True(containsFrame != clipped, "The PDF clip distinguishes fitting and clipped pictures: " + context);
            if (clipped)
                TestAssert.True(paint.Clip.Height > 0d && paint.Clip.Width > 0d, "The clipped control retains partial paint.");
            TestAssert.True(links.Length == (expected ? 1 : 0), "The inline picture has the expected URI count: " + context);
            if (expected)
            {
                var link = links[0];
                TestAssert.True(Math.Abs(link.X - frame.X) < .002d && Math.Abs(link.Y - frame.Y) < .002d &&
                    Math.Abs(link.Width - frame.Width) < .002d && Math.Abs(link.Height - frame.Height) < .002d,
                    "The link rectangle matches the emitted rotation: " + context);
            }
        }
    }

    public static void DocxGlyphlessPictureFilePublishesAndRefundsWithoutFonts()
    {
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        using var budget = new TestMemoryBudget();
        try
        {
            File.WriteAllBytes(output, [1, 2, 3]);
            OoxPdfConverter.Convert(CreateDocument(), output, new OoxPdfOptions { FontResolver = new EmptyProgramResolver(), MemoryBudget = budget });
            string pdf = File.ReadAllText(output, System.Text.Encoding.ASCII);
            TestAssert.Contains("/Subtype /Image", pdf);
            TestAssert.True(Regex.IsMatch(pdf, @"/[A-Za-z0-9]+ Do\b"), "The published PDF paints an XObject.");
            TestAssert.DoesNotContain("/FontFile", pdf);
            TestAssert.Equal(0L, budget.Reserved);
            TestAssert.True(budget.Peak > 0, "The caller funds the conversion's existing resource owners.");
            TestAssert.Equal(0, budget.CallerDisposals);
        }
        finally { File.Delete(output); }
    }

    public static void DocxGlyphlessPictureLateCancellationPreservesPublication()
    {
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        byte[] original = [1, 2, 3];
        using var cancellation = new CancellationTokenSource();
        using var budget = new TestMemoryBudget();
        bool completedRendering = false;
        try
        {
            File.WriteAllBytes(output, original);
            TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(CreateDocument(), output,
                new OoxPdfOptions { FontResolver = new EmptyProgramResolver(), MemoryBudget = budget, ReportResourceUsage = true,
                    DiagnosticSink = diagnostic => { if (diagnostic.Id == "CONVERSION_RESOURCE_SUMMARY") { completedRendering = true; cancellation.Cancel(); } } }, cancellation.Token));
            TestAssert.True(completedRendering, "Cancellation occurs after the fontless image page is rendered.");
            TestAssert.True(original.AsSpan().SequenceEqual(File.ReadAllBytes(output)), "Late cancellation preserves the published destination.");
            TestAssert.Equal(0L, budget.Reserved);
            TestAssert.Equal(0, budget.CallerDisposals);
        }
        finally { File.Delete(output); }
    }

    internal static string CreateDocument()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """),
            ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
                </Relationships>
                """),
            ["word/document.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body>
                    <w:p>
                      <w:r>
                        <w:drawing>
                          <wp:anchor distT="114300" distB="228600" distL="342900" distR="457200"
                                     simplePos="0" relativeHeight="251658240" behindDoc="0"
                                     locked="1" layoutInCell="1" allowOverlap="0">
                            <wp:extent cx="1828800" cy="914400"/>
                            <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
                            <wp:positionV relativeFrom="paragraph"><wp:posOffset>63500</wp:posOffset></wp:positionV>
                            <wp:wrapSquare wrapText="bothSides"/>
                            <a:graphic>
                              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic><pic:blipFill><a:blip r:embed="rIdImage1"/></pic:blipFill></pic:pic>
                              </a:graphicData>
                            </a:graphic>
                          </wp:anchor>
                        </w:drawing>
                      </w:r>
                    </w:p>
                    <w:sectPr>
                      <w:pgSz w:w="12240" w:h="15840"/>
                      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """),
            ["word/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
    }

    private static (OoxPdfDocxMarkupMode, OoxPdfDocxMarkupGeometryMode)[] Views() =>
    [
        (OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout),
        (OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout),
        (OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
    ];

    private static PdfPage[] Render(string input, IFontResolver resolver, OoxPdfDocxMarkupMode mode, OoxPdfDocxMarkupGeometryMode geometry)
    {
        using var stream = File.OpenRead(input);
        using var package = OoxPackage.Open(stream, default);
        var document = new DocxReader().Read(package, null, default, markupMode: mode);
        return new DocxRenderer(resolver, mode, geometry).RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
    }

    private static double[] PaintedFrame(PdfPage page) => PaintedFrame(page, page.Images.Single());

    private static double[] PaintedFrame(PdfPage page, PdfImageResource image)
    {
        var draw = Regex.Match(page.Content, @"([-\d.]+) 0 0 ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\s+/" + Regex.Escape(image.ResourceName) + @" Do");
        TestAssert.True(draw.Success, "The picture resource must be painted.");
        return Enumerable.Range(1, 4).Select(i => double.Parse(draw.Groups[i].Value, CultureInfo.InvariantCulture)).ToArray();
    }

    private static Dictionary<string, byte[]> ReadParts(string input)
    {
        using var stream = File.OpenRead(input);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var source = entry.Open(); using var output = new MemoryStream(); source.CopyTo(output); return output.ToArray();
        });
    }

    internal sealed class EmptyProgramResolver : IFontResolver
    {
        private static readonly MemoryFontProgramSource Source = new("glyphless-empty-font", ReadOnlyMemory<byte>.Empty);
        public FontFaceResolution Resolve(FontRequest request) =>
            new(request.FamilyName, "Helvetica", new FontStyleKey(request.Bold, request.Italic), Source, IsFallback: true);
    }
}
