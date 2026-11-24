using System.Globalization;
using System.Numerics;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionPreferredBreakKeepsTailInsideExactCell()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var control = Paint(hardBreak: true, split: false);
            var wrapped = Paint(hardBreak: false, split: false);
            var split = Paint(hardBreak: false, split: true);
            var prefix = PortableRenderingTests.ReadPaintedSourceLabel(control, "scope-");
            var tail = PortableRenderingTests.ReadPaintedSourceLabel(control, "schedule followup");
            TestAssert.True(prefix[0].Origin.Y > tail[0].Origin.Y, "The hard break reserves the second line for the complete tail.");
            foreach (var actual in new[] { wrapped, split })
            {
                TestAssert.True(prefix.Select(glyph => glyph.Origin).SequenceEqual(
                    PortableRenderingTests.ReadPaintedSourceLabel(actual, "scope-").Select(glyph => glyph.Origin)),
                    "The hyphenated prefix paints at the end of the occupied line.");
                TestAssert.True(tail.Select(glyph => glyph.Origin).SequenceEqual(
                    PortableRenderingTests.ReadPaintedSourceLabel(actual, "schedule followup").Select(glyph => glyph.Origin)),
                    "The complete tail paints inside the same second-line frame as the control.");
                TestAssert.Equal(control.Images.Single().Image.ResourceKey, actual.Images.Single().Image.ResourceKey);
                TestAssert.Equal(PortableRenderingTests.ReadPaintedSourceLabel(control, "Following body.")[0].Origin,
                    PortableRenderingTests.ReadPaintedSourceLabel(actual, "Following body.")[0].Origin);
            }

            PdfPage Paint(bool hardBreak, bool split)
            {
                var document = ReadMutation("docx-run-picture-visible", mode, false, false, (xml, w, _) =>
                {
                    var (paragraph, drawing) = CreateExactPictureCell(xml, w, 2360, 600, 280, "Following body.");
                    XElement Run(params object[] content) => PictureCellRun(w, content);
                    XElement Text(string value) => PictureCellText(w, value);
                    paragraph.Add(Run(Text("Review "), drawing, Text(" ")));
                    if (hardBreak)
                        paragraph.Add(Run(Text("scope-"), new XElement(w + "br"), Text("schedule followup")));
                    else if (split)
                        paragraph.Add(Run(Text("scope")), Run(Text("-schedule followup")));
                    else
                        paragraph.Add(Run(Text("scope-schedule followup")));
                });
                return renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single();
            }
        }
    }

    public static void DocxHangingContinuationKeepsPrecedingFractionAndPicture()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            PdfPage manual = Paint(authoredSplit: true);
            PdfPage wrapped = Paint(authoredSplit: false);
            foreach (var (label, mathOnly) in new[] { ("437", true), ("293", true), ("RENDERINGORDER", false), ("After table.", false) })
            {
                var expected = PortableRenderingTests.ReadPaintedSourceLabel(manual, label, mathOnly);
                var actual = PortableRenderingTests.ReadPaintedSourceLabel(wrapped, label, mathOnly);
                TestAssert.True(expected.SequenceEqual(actual), "Complete glyph ink matches authored placement for " + label);
            }
            var continuation = PortableRenderingTests.ReadPaintedSourceLabel(wrapped, "RENDERINGORDER");
            TestAssert.True(continuation[0].Origin.Y > continuation[^1].Origin.Y,
                "The word wraps inside the narrower continuation rather than rejecting the preceding math.");
            TestAssert.Equal(ReadPaintedPictureFrame(manual), ReadPaintedPictureFrame(wrapped));
            TestAssert.Equal(manual.Images.Single().Image.ResourceKey, wrapped.Images.Single().Image.ResourceKey);
            TestAssert.Equal(ReadPictureFractionRule(manual), ReadPictureFractionRule(wrapped));

            PdfPage Paint(bool authoredSplit)
            {
                var document = ReadMutation("docx-run-picture-visible", mode, false, false, (xml, w, m) =>
                {
                    var (paragraph, drawing) = CreateExactPictureCell(xml, w, 2800, 1320, 320, "After table.");
                    var properties = paragraph.Element(w + "pPr")!;
                    properties.AddFirst(new XElement(w + "wordWrap", new XAttribute(w + "val", "0")));
                    properties.Add(new XElement(w + "ind", new XAttribute(w + "left", "1340"), new XAttribute(w + "hanging", "1340")));
                    AddPictureFractionLabel(paragraph, drawing, w, m);
                    paragraph.Add(PictureCellRun(w, new XElement(w + "br")));
                    if (authoredSplit)
                        paragraph.Add(PictureCellRun(w, PictureCellText(w, "RENDERING"), new XElement(w + "br")),
                            PictureCellRun(w, PictureCellText(w, "ORDER")));
                    else
                        paragraph.Add(PictureCellRun(w, PictureCellText(w, "RENDERINGORDER")));
                });
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = new DocxRenderer(resolver, mode, CachedMathGeometry(mode))
                    .RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
                TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION") &&
                    diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "The original fraction is placed.");
                return page;
            }
        }
    }

    public static void DocxFittingDecimalTabsRetainMathPicturesAndNumericAnchors()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            PdfPage right = PaintPictureTabField(renderer, mode, "right", ["123"]);
            PdfPage integer = PaintPictureTabField(renderer, mode, "decimal", ["123"]);
            TestAssert.Equal(right.Content, integer.Content);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(integer, "123").Count);
            PdfRectangle rule = ReadPictureFractionRule(right);

            foreach (var (number, parts) in new[] { ("12.34", new[] { "12.34" }), ("12,34", new[] { "12", ",34" }) })
            {
                PdfPage page = PaintPictureTabField(renderer, mode, "decimal", parts);
                var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(page, number);
                var label = PortableRenderingTests.ReadPaintedSourceLabel(page, "Evidence");
                TestAssert.Equal(5, glyphs.Count);
                TestAssert.True(Math.Abs(glyphs[2].Origin.X - label[0].Origin.X - 132d) < .02d,
                    "The decimal separator paints at the authored 132pt tab stop.");
                foreach (string value in new[] { "437", "293" })
                    TestAssert.True(PortableRenderingTests.ReadPaintedSourceLabel(right, value, mathOnly: true).SequenceEqual(
                        PortableRenderingTests.ReadPaintedSourceLabel(page, value, mathOnly: true)),
                        "Decimal positioning must preserve the preceding fraction's complete glyph paint.");
                TestAssert.Equal(ReadPaintedPictureFrame(right), ReadPaintedPictureFrame(page));
                TestAssert.Equal(right.Images.Single().Image.ResourceKey, page.Images.Single().Image.ResourceKey);
                TestAssert.Equal(rule, ReadPictureFractionRule(page));
            }
        }
    }

    public static void DocxInlineCompositionAlignedTabsContinueWithBothFractions()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (string id in new[] { "math-tabs-right108-body", "math-tabs-right108-cell", "math-tabs-center108-footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var manual = MathPaint(authoredBreak: true);
            var wrapped = MathPaint(authoredBreak: false);
            TestAssert.Equal("ratetimeratetime", string.Concat(wrapped.Select(glyph => glyph.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(16, wrapped.Length);
            TestAssert.True(manual.SequenceEqual(wrapped),
                "Both complete fractions retain the authored continuation's glyph positions in " + id + "/" + mode);
            TestAssert.True(wrapped[0].Y > wrapped[8].Y,
                "The second fraction moves to the continuation before the aligned field can overlap preceding content.");

            (string Text, double X, double Y)[] MathPaint(bool authoredBreak)
            {
                var document = ReadMutation("CompositionProbes/" + id, mode, false, false, (xml, w, m) =>
                {
                    XElement paragraph = xml.Descendants(w + "p").Single(p => p.Elements(m + "oMath").Any());
                    // With the pinned font, the complete field fits within this
                    // stop only when it starts on a fresh line.
                    paragraph.Element(w + "pPr")!.Element(w + "tabs")!.Elements(w + "tab").First()
                        .SetAttributeValue(w + "pos", "1440");
                    if (authoredBreak)
                        paragraph.Elements(w + "r").First(run => run.Element(w + "tab") is not null)
                            .AddBeforeSelf(new XElement(w + "r", new XElement(w + "br")));
                }, partOverride: id.EndsWith("footer", StringComparison.Ordinal) ? "word/math-tabs-footer.xml" : null);
                var diagnostics = new List<OoxPdfDiagnostic>();
                PdfPage page = new DocxRenderer(resolver, mode, CachedMathGeometry(mode))
                    .RenderPagesAsync(document, diagnostics.Add, CancellationToken.None).ToBlockingEnumerable().Single();
                TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    $"A fitting continuation must preserve both equations in {id}/{mode} (authored break: {authoredBreak}).");
                var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                    .Select(resource => resource.Font).ToHashSet();
                return DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Where(glyph => fonts.Contains(glyph.Font)).Select(glyph => (glyph.Text, VisibleTextStartX(glyph), glyph.Y)).ToArray();
            }
        }
    }

    public static void DocxLogicalTabsPreserveComposedFractionsAndFieldPlacement()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            foreach (bool body in new[] { false, true })
            foreach (var (physical, logical) in new[] { ("left", "start"), ("right", "end") })
            {
                PdfPage control = PaintPictureTabField(renderer, mode, physical, ["123"], body);
                PdfPage aliased = PaintPictureTabField(renderer, mode, logical, ["123"], body);
                TestAssert.Equal(control.Content, aliased.Content);
                TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(aliased, "123").Count);
                TestAssert.Equal(ReadPictureFractionRule(control), ReadPictureFractionRule(aliased));
                TestAssert.Equal(ReadPaintedPictureFrame(control), ReadPaintedPictureFrame(aliased));
                TestAssert.Equal(control.Images.Single().Image.ResourceKey, aliased.Images.Single().Image.ResourceKey);
            }

            PdfPage plainRight = PaintPictureTabField(renderer, mode, "right", ["123"], body: true, composed: false);
            PdfPage plainEnd = PaintPictureTabField(renderer, mode, "end", ["123"], body: true, composed: false);
            PdfPage plainLeft = PaintPictureTabField(renderer, mode, "left", ["123"], body: true, composed: false);
            var rightNumber = PortableRenderingTests.ReadPaintedSourceLabel(plainRight, "123");
            var leftNumber = PortableRenderingTests.ReadPaintedSourceLabel(plainLeft, "123");
            TestAssert.True(rightNumber.SequenceEqual(PortableRenderingTests.ReadPaintedSourceLabel(plainEnd, "123")),
                "The plain end tab uses the right-aligned numeric field geometry.");
            TestAssert.True(rightNumber[0].Origin.X < leftNumber[0].Origin.X,
                "The plain field moves left of the same stop when end-aligned.");
        }
    }

    public static void DocxImplicitDecimalCellTabsMatchExplicitFieldPlacement()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            foreach (var (composed, breaks) in new[] { (false, false), (false, true), (true, false) })
            {
                PdfPage explicitTab = Paint(authored: true);
                PdfPage implicitTab = Paint(authored: false);
                var number = PortableRenderingTests.ReadPaintedSourceLabel(implicitTab, "12.34");
                var following = PortableRenderingTests.ReadPaintedSourceLabel(implicitTab, "After table.");
                TestAssert.True(Math.Abs(number[2].Origin.X - following[0].Origin.X - 145d) < .02d,
                    "The implicit leading field places its decimal separator at the cell's authored stop.");
                TestAssert.True(number.SequenceEqual(PortableRenderingTests.ReadPaintedSourceLabel(explicitTab, "12.34")),
                    "The whole implicit numeric field uses the explicit field's geometry.");
                TestAssert.True(following.SequenceEqual(PortableRenderingTests.ReadPaintedSourceLabel(explicitTab, "After table.")),
                    "The fitting tab keeps the following paragraph in its original flow.");
                if (breaks)
                {
                    var continuation = PortableRenderingTests.ReadPaintedSourceLabel(implicitTab, "Second line.");
                    TestAssert.True(continuation.SequenceEqual(PortableRenderingTests.ReadPaintedSourceLabel(explicitTab, "Second line.")),
                        "Only the leading field consumes the implicit stop.");
                }
                if (composed)
                {
                    TestAssert.Equal(ReadPictureFractionRule(explicitTab), ReadPictureFractionRule(implicitTab));
                    TestAssert.Equal(ReadPictureBounds(explicitTab).Single(), ReadPictureBounds(implicitTab).Single());
                    TestAssert.Equal(explicitTab.Images.Single().Image.ResourceKey, implicitTab.Images.Single().Image.ResourceKey);
                }

                PdfPage Paint(bool authored)
                {
                    var document = ReadMutation("docx-run-picture-visible", mode, false, false, (xml, w, m) =>
                    {
                        var (paragraph, drawing) = CreateExactPictureCell(xml, w, 3600, 1320, 320, "After table.");
                        paragraph.Element(w + "pPr")!.AddFirst(new XElement(w + "tabs", new XElement(w + "tab",
                            new XAttribute(w + "val", "decimal"), new XAttribute(w + "pos", "2900"))));
                        if (authored) paragraph.Add(PictureCellRun(w, new XElement(w + "tab")));
                        if (composed) AddPictureFractionLabel(paragraph, drawing, w, m);
                        paragraph.Add(PictureCellRun(w, PictureCellText(w, "12.34")));
                        if (breaks)
                            paragraph.Add(PictureCellRun(w, new XElement(w + "br"), PictureCellText(w, "Second line.")));
                    });
                    return renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
                }
            }
        }
    }

    public static void DocxLogicalParagraphEdgesMatchPhysicalAlignment()
    {
        const string styleId = "LogicalParagraphTest";
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool body in new[] { false, true })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            PdfPage left = Paint("left");
            PdfPage right = Paint("right");
            foreach (var (control, logical) in new[] { (left, "start"), (right, "end") })
            {
                PdfPage actual = Paint(logical);
                TestAssert.Equal(control.Content, actual.Content);
                TestAssert.Equal(ReadPictureFractionRule(control), ReadPictureFractionRule(actual));
                TestAssert.Equal(ReadPictureBounds(control).Single(), ReadPictureBounds(actual).Single());
                TestAssert.Equal(control.Images.Single().Image.ResourceKey, actual.Images.Single().Image.ResourceKey);
                TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual, "123").Count);
            }
            var before = PortableRenderingTests.ReadPaintedSourceLabel(left, "Evidence");
            var after = PortableRenderingTests.ReadPaintedSourceLabel(right, "Evidence");
            TestAssert.True(after[0].Origin.X > before[0].Origin.X + 10d,
                "The right-edge source moves complete paint within its owning paragraph frame.");

            PdfPage Paint(string alignment)
            {
                var document = ReadMutation("docx-run-picture-visible", mode, false, false, (xml, w, m) =>
                {
                    var (paragraph, drawing) = CreateExactPictureCell(xml, w, 3600, 1320, 320, "After table.");
                    var properties = paragraph.Element(w + "pPr")!;
                    AddPictureFractionLabel(paragraph, drawing, w, m);
                    paragraph.Add(PictureCellRun(w, PictureCellText(w, "123")));
                    if (body)
                    {
                        properties.Add(Justification(alignment));
                        MovePictureParagraphToBody(paragraph, w);
                    }
                    else
                        properties.AddFirst(new XElement(w + "pStyle", new XAttribute(w + "val", styleId)));
                }, packageChange: body ? null : zip =>
                {
                    var entry = zip.GetEntry("word/styles.xml")!;
                    XDocument styles;
                    using (var stream = entry.Open()) styles = XDocument.Load(stream);
                    styles.Root!.Add(new XElement(word + "style", new XAttribute(word + "type", "paragraph"),
                        new XAttribute(word + "styleId", styleId), new XElement(word + "name", new XAttribute(word + "val", styleId)),
                        new XElement(word + "pPr", Justification(alignment))));
                    entry.Delete();
                    using var output = zip.CreateEntry("word/styles.xml").Open();
                    styles.Save(output);
                });
                return renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
            }
        }

        XElement Justification(string value) => new(word + "jc", new XAttribute(word + "val", value));
    }

    private static PdfPage PaintPictureTabField(DocxRenderer renderer, OoxPdfDocxMarkupMode mode,
        string alignment, IReadOnlyList<string> number, bool body = false, bool composed = true)
    {
        var document = ReadMutation("docx-run-picture-visible", mode, false, false, (xml, w, m) =>
        {
            var (paragraph, drawing) = CreateExactPictureCell(xml, w, 3600, 1320, 320, "After table.");
            paragraph.Element(w + "pPr")!.AddFirst(new XElement(w + "tabs", new XElement(w + "tab",
                new XAttribute(w + "val", alignment), new XAttribute(w + "pos", "2640"))));
            if (composed) AddPictureFractionLabel(paragraph, drawing, w, m);
            paragraph.Add(PictureCellRun(w, new XElement(w + "tab")));
            foreach (string part in number)
                paragraph.Add(PictureCellRun(w, PictureCellText(w, part)));
            if (body) MovePictureParagraphToBody(paragraph, w);
        });
        var diagnostics = new List<OoxPdfDiagnostic>();
        var page = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
        if (composed)
        {
            var equation = diagnostics.Single(d => d.Feature == "equation");
            TestAssert.True(equation.Id == "DOCX_APPROXIMATED_EQUATION",
                $"The {alignment} stop keeps the fraction beside {string.Concat(number)} in {mode}. {equation.Message}");
        }
        return page;
    }

    private static void MovePictureParagraphToBody(XElement paragraph, XNamespace word)
    {
        var table = paragraph.Ancestors(word + "tbl").Single();
        paragraph.Remove();
        table.ReplaceWith(paragraph);
    }

    private static PdfRectangle ReadPictureFractionRule(PdfPage page)
    {
        var numerator = PortableRenderingTests.ReadPaintedSourceLabel(page, "437", mathOnly: true);
        var denominator = PortableRenderingTests.ReadPaintedSourceLabel(page, "293", mathOnly: true);
        var rules = new List<PdfRectangle>();
        foreach (var (block, clip, transform) in DocxNoteClippingTests.ReadPhysicalPaintFrames(page, includeVisibleFills: true))
        {
            if (!block.EndsWith(" re f", StringComparison.Ordinal)) continue;
            string[] tokens = block.Split(' ');
            float Number(int index) => float.Parse(tokens[index], CultureInfo.InvariantCulture);
            var start = Vector2.Transform(new Vector2(Number(0), Number(1)), transform);
            var end = Vector2.Transform(new Vector2(Number(0) + Number(2), Number(1) + Number(3)), transform);
            if (start.X <= Math.Min(numerator[0].Origin.X, denominator[0].Origin.X) &&
                end.X > Math.Max(numerator[^1].Origin.X, denominator[^1].Origin.X) &&
                start.Y > denominator[0].Origin.Y && end.Y < numerator[0].Origin.Y &&
                start.X >= clip.X && end.X <= clip.X + clip.Width &&
                start.Y >= clip.Y && end.Y <= clip.Y + clip.Height &&
                end.Y > start.Y && end.X > start.X)
                rules.Add(new PdfRectangle(start.X, start.Y, end.X - start.X, end.Y - start.Y));
        }
        TestAssert.True(rules.Count == 1, "The fraction paints one complete visible rule between its arguments.");
        return rules[0];
    }

    private static void AddPictureFractionLabel(XElement paragraph, XElement drawing, XNamespace w, XNamespace m)
    {
        XElement FractionPart(string kind, string value) =>
            new(m + kind, new XElement(m + "r", new XElement(m + "t", value)));
        paragraph.Add(PictureCellRun(w, PictureCellText(w, "Evidence ")), PictureCellRun(w, drawing),
            new XElement(m + "oMath", new XElement(m + "f", FractionPart("num", "437"), FractionPart("den", "293"))));
    }

    private static (XElement Paragraph, XElement Drawing) CreateExactPictureCell(XDocument xml, XNamespace w,
        int width, int height, int lineSpacing, string followingText)
    {
        var body = xml.Root!.Element(w + "body")!;
        var section = body.Element(w + "sectPr")!;
        var drawing = body.Descendants(w + "drawing").Single();
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (var extent in drawing.Descendants(wp + "extent").Concat(drawing.Descendants(a + "ext")
            .Where(e => e.Parent?.Name == a + "xfrm")))
        {
            extent.SetAttributeValue("cx", 279400);
            extent.SetAttributeValue("cy", 139700);
        }
        var table = XElement.Parse($$"""
            <w:tbl xmlns:w="{{w}}">
              <w:tblPr><w:tblW w:w="{{width}}" w:type="dxa"/><w:tblLayout w:type="fixed"/>
                <w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>
                  <w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>
              </w:tblPr>
              <w:tblGrid><w:gridCol w:w="{{width}}"/></w:tblGrid>
              <w:tr><w:trPr><w:trHeight w:val="{{height}}" w:hRule="exact"/></w:trPr>
                <w:tc><w:tcPr><w:tcW w:w="{{width}}" w:type="dxa"/></w:tcPr>
                  <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{{lineSpacing}}" w:lineRule="exact"/></w:pPr></w:p>
                </w:tc>
              </w:tr>
            </w:tbl>
            """);
        body.ReplaceNodes(table, new XElement(w + "p", PictureCellRun(w, PictureCellText(w, followingText))), section);
        return (table.Descendants(w + "p").Single(), drawing);
    }

    private static XElement PictureCellRun(XNamespace w, params object[] content) => new(w + "r",
        new XElement(w + "rPr", new XElement(w + "sz", new XAttribute(w + "val", "24"))), content);

    private static XElement PictureCellText(XNamespace w, string value) =>
        new(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value);
}
