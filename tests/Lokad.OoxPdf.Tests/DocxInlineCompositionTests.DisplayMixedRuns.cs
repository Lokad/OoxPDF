using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxDisplayMathKeepsOperandsAroundDrawingCompanions()
    {
        foreach (bool withBreak in new[] { false, true })
            CheckDisplayDrawingCompanions(withBreak, explicitRow: false);
    }

    public static void DocxDisplayRowKeepsAdjacentOperandsAroundDrawingCompanions() =>
        CheckDisplayDrawingCompanions(withBreak: false, explicitRow: true);

    public static void DocxInlineEquationKeepsDrawingCompanionsBetweenTextRuns() =>
        CheckDisplayDrawingCompanions(withBreak: false, explicitRow: true, inlineEquation: true);

    public static void DocxEquationZonesKeepDrawingsWithoutMathOperands()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (bool inDisplay in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var expected = Render(packed: false);
            var actual = Render(packed: true);
            TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "QA").Count);
            TestAssert.Equal(1, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "Ω").Count);
            TestAssert.Contains($"/{actual.Page.Images.Single().ResourceName} Do", actual.Page.Content);
            TestAssert.Equal(expected.Page.Content, actual.Page.Content);
            TestAssert.Equal(expected.Equations, actual.Equations);
            TestAssert.True(expected.Diagnostics.SequenceEqual(actual.Diagnostics) &&
                !actual.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION"),
                "A drawing with no operands keeps its ordinary paint and contributes no equation to the count.");

            (PdfPage Page, int Equations, List<string> Diagnostics) Render(bool packed)
            {
                var diagnostics = new List<string>();
                void Record(Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic diagnostic) => diagnostics.Add(diagnostic.Id);
                var source = ReadMutation("CompositionProbes/math-display-image-" + container + "-between", mode,
                    false, false, (xml, w, m) =>
                    {
                        var display = xml.Descendants(m + "oMathPara").First();
                        var picture = display.Parent!.Elements(w + "r").Single();
                        var drawing = picture.Element(w + "drawing")!;
                        var companion = new XElement((packed ? m : w) + "r", new XElement(w + "t", "QA"),
                            new XElement(w + "sym", new XAttribute(w + "font", "Arial"), new XAttribute(w + "char", "03A9")), drawing);
                        XElement representation = packed ? new XElement(m + "oMath", companion) : companion;
                        if (inDisplay)
                        {
                            display.Add(representation);
                            picture.Remove();
                        }
                        else
                        {
                            picture.ReplaceWith(representation);
                        }
                    }, diagnosticSink: Record);
                var page = PictureRangeRenderer(mode).RenderPagesAsync(source, Record, default).ToBlockingEnumerable().Single();
                return (page, source.EquationParts.Sum(part => part.SourceCount), diagnostics);
            }
        }
    }

    public static void DocxStaticInlineEquationKeepsDrawingCompanions()
    {
        foreach (string story in new[] { "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = "static-inline-ordinary-" + story + "-picture";
            var source = ReadMutation("InlineBoxMath/" + fixture, mode, false, false, (xml, w, m) =>
            {
                var paragraph = xml.Root!.Element(w + "p")!;
                var drawing = paragraph.Elements(w + "r").Select(run => run.Element(w + "drawing")).OfType<XElement>().First();
                paragraph.ReplaceNodes(
                    new XElement(w + "r", new XElement(w + "t", "Static before")),
                    new XElement(m + "oMath", new XElement(m + "r",
                        new XElement(m + "t", "314"), new XElement(w + "t", "QA"), drawing, new XElement(m + "t", "793"))),
                    new XElement(w + "r", new XElement(w + "t", "Static after")));
            }, partOverride: "word/" + fixture + ".xml");
            var page = PictureRangeRenderer(mode).RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
            foreach (string label in new[] { "Static before", "314", "QA", "793", "Static after" })
            {
                var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(page, label, mathOnly: label is "314" or "793");
                TestAssert.True(glyphs.All(glyph => InStory(glyph.Ink)), $"'{label}' must paint in the {story}.");
            }
            var (picture, clip) = ReadPictureBounds(page).Single();
            TestAssert.True(InStory(picture) && picture.X >= clip.X - .01d && picture.Y >= clip.Y - .01d &&
                picture.X + picture.Width <= clip.X + clip.Width + .01d && picture.Y + picture.Height <= clip.Y + clip.Height + .01d,
                $"The complete picture must paint inside its {story} clip.");

            bool InStory(PdfRectangle bounds) => story == "header" ? bounds.Y > page.Height / 2d :
                bounds.Y + bounds.Height < page.Height / 2d;
        }
    }

    public static void DocxDisplayRowKeepsBookmarkAcrossDrawingCompanions() =>
        CheckDisplayDrawingCompanions(withBreak: false, explicitRow: true, withBookmark: true);

    public static void DocxDisplayRowKeepsDrawingCompanionsInTrackedChanges()
    {
        foreach (var revision in new[] { DocxRevisionKind.Insertion, DocxRevisionKind.Deletion })
            CheckDisplayDrawingCompanions(withBreak: false, explicitRow: true, revision: revision);
    }

    public static void DocxEquationDrawingCompanionsKeepInnerRevisionOwners()
    {
        foreach ((string container, bool display) in new[] { ("body", false), ("cell", true) })
        foreach (bool partial in new[] { false, true })
        foreach (bool deletion in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
        {
            bool visible = mode == OoxPdfDocxMarkupMode.AllMarkup || mode ==
                (deletion ? OoxPdfDocxMarkupMode.Original : OoxPdfDocxMarkupMode.Final);
            var expected = Render(packed: false);
            var actual = Render(packed: true);
            TestAssert.Equal(expected.Page.Content, actual.Page.Content);
            TestAssert.Equal(expected.Equations, actual.Equations);
            if (visible)
                TestAssert.True(expected.Diagnostics.SequenceEqual(actual.Diagnostics), "Equivalent revision owners report the same diagnostics.");
            TestAssert.True(!actual.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION"), "Every selected math fragment is accounted for.");
            TestAssert.Equal(visible ? 1 : 0, actual.Page.Images.Count);
            if (visible)
            {
                TestAssert.Contains($"/{actual.Page.Images.Single().ResourceName} Do", actual.Page.Content);
                TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "QA").Count);
                TestAssert.Equal(1, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "Ω").Count);
            }
            string painted = DocxPageFootnoteNumberingTests.DecodeText(actual.Page);
            foreach (string label in new[] { "17", "314", "793", "28" })
            {
                bool neighbor = label is "17" or "28";
                if (visible || partial && neighbor)
                    TestAssert.Equal(label.Length, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, label, mathOnly: true).Count);
                else
                    TestAssert.DoesNotContain(label, painted);
            }
            if (!visible)
            {
                TestAssert.DoesNotContain("QA", painted);
                TestAssert.DoesNotContain("Ω", painted);
            }

            (PdfPage Page, int Equations, List<string> Diagnostics) Render(bool packed)
            {
                var diagnostics = new List<string>();
                void Record(Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic item) => diagnostics.Add(item.Id);
                var source = ReadMutation("CompositionProbes/math-display-image-" + container + "-between", mode,
                    false, false, (xml, w, m) =>
                    {
                        var displayRow = xml.Descendants(m + "oMathPara").First();
                        var picture = displayRow.Parent!.Elements(w + "r").Single();
                        var drawing = picture.Element(w + "drawing")!;
                        picture.Remove();
                        var before = MathRun("17");
                        var after = MathRun("28");
                        var label = new XElement(w + (deletion ? "delText" : "t"), "QA");
                        var symbol = new XElement(w + "sym", new XAttribute(w + "font", "Arial"), new XAttribute(w + "char", "03A9"));
                        var owner = new XElement(w + (deletion ? "del" : "ins"), new XAttribute(w + "id", "528"),
                            new XAttribute(w + "author", "QA Reviewer"));
                        XElement[] items;
                        if (packed)
                        {
                            owner.Add(partial ? null : before,
                                new XElement(m + "r", new XElement(m + "t", "314"), label, symbol, drawing, new XElement(m + "t", "793")),
                                partial ? null : after);
                            items = [new XElement(m + "oMath", partial ? before : null, owner, partial ? after : null)];
                        }
                        else if (visible)
                        {
                            owner.Add(new XElement(m + "oMath", partial ? null : before, MathRun("314")),
                                new XElement(w + "r", label, symbol, drawing),
                                new XElement(m + "oMath", MathRun("793"), partial ? null : after));
                            items = partial ? [new XElement(m + "oMath", before), owner, new XElement(m + "oMath", after)] : [owner];
                        }
                        else
                        {
                            items = partial ? [new XElement(m + "oMath", before, after)] : [];
                        }
                        if (display) displayRow.Add(items);
                        else displayRow.AddAfterSelf(items);

                        XElement MathRun(string text) => new(m + "r", new XElement(m + "t", text));
                    }, diagnosticSink: Record);
                var page = PictureRangeRenderer(mode).RenderPagesAsync(source, Record, default).ToBlockingEnumerable().Single();
                return (page, source.EquationParts.Sum(part => part.SourceCount), diagnostics);
            }
        }
    }

    public static void DocxEquationDrawingCompanionsKeepInnerHyperlinkCoverage()
    {
        foreach ((string container, bool display) in new[] { ("body", false), ("cell", false), ("cell", true) })
        foreach (bool partial in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var expected = Render(inner: false);
            var actual = Render(inner: true);
            TestAssert.Equal(expected.Page.Content, actual.Page.Content);
            TestAssert.Equal(expected.Equations, actual.Equations);
            TestAssert.True(expected.Diagnostics.SequenceEqual(actual.Diagnostics) &&
                !actual.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION"), "Moving a link across the equation wrapper preserves every visible item.");

            var link = actual.Page.Annotations.Single();
            var outer = expected.Page.Annotations.Single();
            TestAssert.Equal((outer.X, outer.Y, outer.Width, outer.Height), (link.X, link.Y, link.Width, link.Height));
            TestAssert.Equal(outer.Destination, link.Destination);
            TestAssert.True(link.Destination is { PageIndex: 0, Left: not null, Top: not null }, "The link goes to the authored heading.");
            var picture = ReadPictureBounds(actual.Page).Single().Bounds;
            TestAssert.True(picture.X >= link.X - .01d && picture.Y >= link.Y - .01d &&
                picture.X + picture.Width <= link.X + link.Width + .01d && picture.Y + picture.Height <= link.Y + link.Height + .01d,
                "The link covers the full picture.");
            foreach (string label in new[] { "17", "314", "793", "28" })
            {
                var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, label, mathOnly: true);
                TestAssert.Equal(label.Length, glyphs.Count);
                bool linked = !partial || label is "314" or "793";
                // The ordinary link pads the hit area at its ends. Use the
                // operand center and keep the exact rectangle comparison above.
                double center = (glyphs.Min(glyph => glyph.Ink.X) + glyphs.Max(glyph => glyph.Ink.X + glyph.Ink.Width)) / 2d;
                TestAssert.True(linked == (center >= link.X && center <= link.X + link.Width),
                    $"The link keeps the authored ownership of operand {label} ({container}, display={display}, partial={partial}, {mode}).");
            }
            TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "QA").Count);
            TestAssert.Equal(1, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "Ω").Count);

            (PdfPage Page, int Equations, List<string> Diagnostics) Render(bool inner)
            {
                var diagnostics = new List<string>();
                void Record(Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic item) => diagnostics.Add(item.Id);
                var source = ReadMutation("CompositionProbes/math-display-image-" + container + "-between", mode,
                    false, false, (xml, w, m) =>
                    {
                        const string target = "InnerLinkDestination";
                        var headingRun = xml.Descendants(w + "p").First().Elements(w + "r").First();
                        headingRun.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "529"), new XAttribute(w + "name", target)));
                        headingRun.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "529")));
                        var displayRow = xml.Descendants(m + "oMathPara").First();
                        var pictureRun = displayRow.Parent!.Elements(w + "r").Single();
                        var inlinePicture = pictureRun.Element(w + "drawing")!;
                        pictureRun.Remove();
                        var left = MathRun("17");
                        var right = MathRun("28");
                        var mixed = new XElement(m + "r", new XElement(m + "t", "314"),
                            new XElement(w + "t", "QA"),
                            new XElement(w + "sym", new XAttribute(w + "font", "Arial"), new XAttribute(w + "char", "03A9")),
                            inlinePicture, new XElement(m + "t", "793"));
                        var owner = new XElement(w + "hyperlink", new XAttribute(w + "anchor", target));
                        var properties = new XElement(m + "oMathParaPr", new XElement(m + "jc", new XAttribute(m + "val", "left")));
                        XElement Frame(params object?[] items)
                        {
                            var math = new XElement(m + "oMath", items);
                            return display ? new XElement(m + "oMathPara", properties, math) : math;
                        }
                        if (inner)
                        {
                            owner.Add(partial ? null : left, mixed, partial ? null : right);
                            displayRow.AddAfterSelf(Frame(partial ? left : null, owner, partial ? right : null));
                        }
                        else
                        {
                            owner.Add(Frame(partial ? null : left, mixed, partial ? null : right));
                            displayRow.AddAfterSelf(partial ? Frame(left) : null, owner, partial ? Frame(right) : null);
                        }

                        XElement MathRun(string text) => new(m + "r", new XElement(m + "t", text));
                    }, diagnosticSink: Record);
                var page = PictureRangeRenderer(mode).RenderPagesAsync(source, Record, default).ToBlockingEnumerable().Single();
                return (page, source.EquationParts.Sum(part => part.SourceCount), diagnostics);
            }
        }
    }

    public static void DocxEquationDrawingCompanionsKeepCombinedLinkAndRevisionScopes()
    {
        foreach ((string container, bool display, string revisionName) in new[] { ("body", false, "ins"), ("cell", true, "del") })
        foreach (bool linkOutside in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup })
        {
            bool visible = mode == OoxPdfDocxMarkupMode.AllMarkup ||
                mode == (revisionName == "ins" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.Original);
            var expected = Render(inner: false);
            var actual = Render(inner: true);
            var wrapped = Render(inner: true, withMetadata: true);
            TestAssert.True(actual.Page.Content == wrapped.Page.Content && actual.Page.Annotations.SequenceEqual(wrapped.Page.Annotations),
                $"Metadata preserves the linked revision's paint and navigation for {container}, linkOutside={linkOutside}, {mode}.");
            TestAssert.Equal(actual.Equations, wrapped.Equations);
            TestAssert.True(!wrapped.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION"), "Metadata preserves the drawing projection's source count.");
            TestAssert.True(expected.Page.Content == actual.Page.Content,
                $"Packed and split paint agree for {container}, {revisionName}, linkOutside={linkOutside}, {mode}.");
            TestAssert.Equal(expected.Equations, actual.Equations);
            TestAssert.True(!actual.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION"), "Visible fragments and omitted revisions are counted by their source scopes.");
            if (visible)
                TestAssert.True(expected.Diagnostics.SequenceEqual(actual.Diagnostics), "The visible owners retain the split control's diagnostics.");
            TestAssert.True(expected.Page.Annotations.SequenceEqual(actual.Page.Annotations), "The hyperlink keeps the split control's destination and hit area.");
            TestAssert.Equal(visible || linkOutside ? 1 : 0, actual.Page.Annotations.Count);
            TestAssert.Equal(visible ? 1 : 0, actual.Page.Images.Count);
            string painted = DocxPageFootnoteNumberingTests.DecodeText(actual.Page);
            foreach (string label in new[] { "17", "314", "793", "28", "QA", "Ω" })
            {
                bool neighbor = label is "17" or "28";
                if (visible || linkOutside && neighbor)
                    TestAssert.Equal(label.Length, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, label, mathOnly: label is not ("QA" or "Ω")).Count);
                else
                    TestAssert.DoesNotContain(label, painted);
            }
            if (visible)
            {
                var link = actual.Page.Annotations.Single();
                var picture = ReadPictureBounds(actual.Page).Single().Bounds;
                TestAssert.True(link.Destination is { PageIndex: 0 } && picture.X >= link.X - .01d &&
                    picture.X + picture.Width <= link.X + link.Width + .01d && picture.Y >= link.Y - .01d &&
                    picture.Y + picture.Height <= link.Y + link.Height + .01d,
                    "The revised picture retains full link coverage and the original target.");
            }

            (PdfPage Page, int Equations, List<string> Diagnostics) Render(bool inner, bool withMetadata = false)
            {
                var diagnostics = new List<string>();
                void Record(Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic item) => diagnostics.Add(item.Id);
                var source = ReadMutation("CompositionProbes/math-display-image-" + container + "-between", mode,
                    false, false, (xml, w, m) =>
                    {
                        const string target = "LinkedRevisionTarget";
                        var heading = xml.Descendants(w + "p").First().Elements(w + "r").First();
                        heading.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "534"), new XAttribute(w + "name", target)));
                        heading.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "534")));
                        var displayRow = xml.Descendants(m + "oMathPara").First();
                        var pictureRun = displayRow.Parent!.Elements(w + "r").Single();
                        var drawing = pictureRun.Element(w + "drawing")!;
                        pictureRun.Remove();

                        var link = new XElement(w + "hyperlink", new XAttribute(w + "anchor", target));
                        var revision = new XElement(w + revisionName, new XAttribute(w + "id", "535"), new XAttribute(w + "author", "QA Reviewer"));
                        var caption = new XElement(w + (revisionName == "del" ? "delText" : "t"), "QA");
                        var symbol = new XElement(w + "sym", new XAttribute(w + "font", "Arial"), new XAttribute(w + "char", "03A9"));
                        XElement MathRun(string text) => new(m + "r", new XElement(m + "t", text));
                        XElement Frame(params XElement[] children)
                        {
                            var math = new XElement(m + "oMath", children);
                            return display ? new XElement(m + "oMathPara", math) : math;
                        }

                        if (!inner && !visible)
                        {
                            if (linkOutside)
                                displayRow.AddAfterSelf(new XElement(link.Name, link.Attributes(), Frame(MathRun("17"), MathRun("28"))));
                            return;
                        }

                        XElement first = MathRun("17"), last = MathRun("28");
                        XElement[] middle = inner
                            ? [new XElement(m + "r", new XElement(m + "t", "314"), caption, symbol, drawing, new XElement(m + "t", "793"))]
                            : [Frame(MathRun("314")), new XElement(w + "r", caption, symbol, drawing), Frame(MathRun("793"))];
                        if (withMetadata)
                        {
                            var run = middle.Single();
                            middle[0] = display ? MetadataWrapper(w, "smartTag", run) :
                                new XElement(w + "sdt", new XElement(w + "sdtPr", new XElement(w + "id", new XAttribute(w + "val", "541"))),
                                    new XElement(w + "sdtContent", MetadataWrapper(w, "customXml", run)));
                        }
                        XElement innerOwner = linkOutside ? revision : link;
                        XElement outerOwner = linkOutside ? link : revision;
                        innerOwner.Add(middle);
                        outerOwner.Add(inner ? first : Frame(first), innerOwner, inner ? last : Frame(last));
                        displayRow.AddAfterSelf(inner ? Frame(outerOwner) : outerOwner);
                    }, diagnosticSink: Record);
                var page = PictureRangeRenderer(mode).RenderPagesAsync(source, Record, default).ToBlockingEnumerable().Single();
                return (page, source.EquationParts.Sum(part => part.SourceCount), diagnostics);
            }
        }
    }

    public static void DocxDisplayRowKeepsCommentRangeThroughDrawing()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var expected = Render(inside: false);
            var actual = Render(inside: true);
            TestAssert.Equal(expected.Page.Content, actual.Page.Content);
            TestAssert.Equal(1, actual.Page.Images.Count);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "314", mathOnly: true).Count);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "793", mathOnly: true).Count);
            TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "QA").Count);
            TestAssert.True(expected.Diagnostics.SequenceEqual(actual.Diagnostics) &&
                !actual.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION"), "Range metadata does not drop any equation fragments.");
            string text = DocxPageFootnoteNumberingTests.DecodeText(actual.Page);
            TestAssert.Equal(mode == OoxPdfDocxMarkupMode.AllMarkup, text.Contains("Check printed geometry.", StringComparison.Ordinal));

            (PdfPage Page, List<string> Diagnostics) Render(bool inside)
            {
                var diagnostics = new List<string>();
                void Record(Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic item) => diagnostics.Add(item.Id);
                var document = ReadMutation("CompositionProbes/math-display-image-" + container + "-between", mode,
                    false, false, (xml, w, m) =>
                    {
                        var display = xml.Descendants(m + "oMathPara").First();
                        var originalPicture = display.Parent!.Elements(w + "r").Single();
                        var drawing = originalPicture.Element(w + "drawing")!;
                        var start = xml.Descendants(w + "commentRangeStart").Single();
                        var end = xml.Descendants(w + "commentRangeEnd").Single();
                        var referenceRun = xml.Descendants(w + "commentReference").Single().Parent!;
                        start.Remove();
                        end.Remove();
                        referenceRun.Remove();
                        var picture = new XElement((inside ? m : w) + "r", new XElement(w + "t", "QA"), drawing);
                        originalPicture.Remove();
                        if (inside)
                            display.Add(new XElement(m + "oMath", start, MathRun("314"), picture, end, MathRun("793")));
                        else
                            display.Add(new XElement(m + "oMath", start, MathRun("314")), picture, end,
                                new XElement(m + "oMath", MathRun("793")));
                        display.AddAfterSelf(referenceRun);

                        XElement MathRun(string text) => new(m + "r", new XElement(m + "t", text));
                    }, diagnosticSink: Record);
                var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).Single(p => p.Images.Count != 0);
                var range = paragraph.CommentRanges.Single();
                var items = TestAssert.NotNull(paragraph.InlineContent);
                int start = range.StartInlineItemIndex!.Value;
                int end = range.EndInlineItemIndex!.Value;
                TestAssert.True(items[start] is DocxInlineEquation { Equation.Content.Children: [DocxMathRun { Text.Text: "314" }] } &&
                    items[end - 1] is DocxInlinePicture &&
                    items[end] is DocxInlineEquation { Equation.Content.Children: [DocxMathRun { Text.Text: "793" }] },
                    "The range closes after the picture and before the following operand.");
                var page = PictureRangeRenderer(mode).RenderPagesAsync(document, Record, default).ToBlockingEnumerable().Single();
                return (page, diagnostics);
            }
        }
    }

    private static void CheckDisplayDrawingCompanions(bool withBreak, bool explicitRow, bool withBookmark = false,
        DocxRevisionKind? revision = null, bool inlineEquation = false)
    {
        OoxPdfDocxMarkupMode[] modes = revision is null
            ? [OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup]
            : [OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original, OoxPdfDocxMarkupMode.AllMarkup];
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in modes)
        {
            bool visible = (revision, mode) is not
                ((DocxRevisionKind.Insertion, OoxPdfDocxMarkupMode.Original) or (DocxRevisionKind.Deletion, OoxPdfDocxMarkupMode.Final));
            var expected = Render(packed: false);
            var actual = Render(packed: true);
            TestAssert.Equal(expected.Page.Content, actual.Page.Content);
            TestAssert.True(expected.Diagnostics.SequenceEqual(actual.Diagnostics),
                "Packing should preserve the diagnostics for the same revision view.");
            TestAssert.Equal(expected.SourceEquationCount, actual.SourceEquationCount);
            if (!visible)
            {
                TestAssert.Equal(0, actual.Page.Images.Count);
                string painted = DocxPageFootnoteNumberingTests.DecodeText(actual.Page);
                foreach (string label in new[] { "17", "314", "793", "28", "QA", "Ω" })
                    TestAssert.DoesNotContain(label, painted);
                continue;
            }

            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "314", mathOnly: true).Count);
            TestAssert.Equal(3, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "793", mathOnly: true).Count);
            if (explicitRow)
            {
                var leading = PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "17", mathOnly: true);
                TestAssert.Equal(2, leading.Count);
                TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "28", mathOnly: true).Count);
                if (withBookmark)
                {
                    var link = actual.Page.Annotations.Single();
                    var outer = expected.Page.Annotations.Single();
                    TestAssert.Equal((outer.X, outer.Y, outer.Width, outer.Height),
                        (link.X, link.Y, link.Width, link.Height));
                    var ink = leading[0].Ink;
                    TestAssert.True(link.Destination is { PageIndex: 0, Left: double left, Top: double top } &&
                        outer.Destination is { Left: double outerLeft, Top: double outerTop } &&
                        Math.Abs(left - outerLeft) < .002d &&
                        top >= ink.Y + ink.Height - .002d && top <= outerTop + .002d,
                        "The inner bookmark opens above the first operand, within the containing line's viewport.");
                }
            }
            TestAssert.Equal(2, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "QA").Count);
            TestAssert.Equal(1, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "Ω").Count);
            if (inlineEquation)
            {
                TestAssert.Equal(6, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "before").Count);
                TestAssert.Equal(5, PortableRenderingTests.ReadPaintedSourceLabel(actual.Page, "after").Count);
            }
            TestAssert.Equal(1, actual.Page.Images.Count);
            TestAssert.Contains($"/{actual.Page.Images.Single().ResourceName} Do", actual.Page.Content);
            TestAssert.True(!actual.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION"),
                "Every projected equation fragment must have a matching source count.");

            (PdfPage Page, List<string> Diagnostics, int SourceEquationCount) Render(bool packed)
            {
                var diagnostics = new List<string>();
                void Record(Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic item) => diagnostics.Add(item.Id);
                var source = ReadMutation("CompositionProbes/math-display-image-" + container + "-between", mode,
                    false, false, (xml, w, m) =>
                    {
                        var display = xml.Descendants(m + "oMathPara").First();
                        var originalPictureRun = display.Parent!.Elements(w + "r").Single();
                        var drawing = originalPictureRun.Element(w + "drawing")!;
                        var mathProperties = new XElement(m + "rPr", new XElement(m + "sty", new XAttribute(m + "val", "p")));
                        var wordProperties = new XElement(w + "rPr", new XElement(w + "color", new XAttribute(w + "val", "005599")));
                        var leftOperand = new XElement(m + "t", "314");
                        var rightOperand = new XElement(m + "t", "793");
                        var label = new XElement(w + (revision == DocxRevisionKind.Deletion ? "delText" : "t"), "QA");
                        var symbol = new XElement(w + "sym", new XAttribute(w + "font", "Arial"), new XAttribute(w + "char", "03A9"));
                        var lineBreak = withBreak ? new XElement(w + "br") : null;
                        var leftNeighbor = explicitRow ? MathRun(new XElement(m + "t", "17")) : null;
                        var rightNeighbor = explicitRow ? MathRun(new XElement(m + "t", "28")) : null;
                        XElement? bookmarkStart = null;
                        XElement? bookmarkEnd = null;
                        if (withBookmark)
                        {
                            const string name = "EquationDrawing";
                            bookmarkStart = new XElement(w + "bookmarkStart", new XAttribute(w + "id", "412"), new XAttribute(w + "name", name));
                            bookmarkEnd = new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "412"));
                            xml.Descendants(w + "p").First().Add(
                                new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " ")),
                                new XElement(w + "hyperlink", new XAttribute(w + "anchor", name),
                                    new XElement(w + "r", new XElement(w + "t", "Jump row"))));
                        }
                        XElement owner = inlineEquation ? new XElement(w + "p") : display;
                        if (revision is not null)
                        {
                            owner = new XElement(w + (revision == DocxRevisionKind.Insertion ? "ins" : "del"),
                                new XAttribute(w + "id", "412"), new XAttribute(w + "author", "QA Reviewer"));
                            display.Add(owner);
                        }
                        if (packed)
                        {
                            var run = MathRun(leftOperand, lineBreak, label, symbol, drawing, rightOperand);
                            owner.Add(explicitRow ? new XElement(m + "oMath", bookmarkStart, leftNeighbor, run, rightNeighbor, bookmarkEnd) : run);
                        }
                        else
                        {
                            var leftRun = MathRun(leftOperand);
                            var rightRun = MathRun(rightOperand);
                            owner.Add(bookmarkStart,
                                explicitRow ? new XElement(m + "oMath", leftNeighbor, leftRun) : leftRun,
                                new XElement(w + "r", wordProperties, lineBreak, label, symbol, drawing),
                                explicitRow ? new XElement(m + "oMath", rightRun, rightNeighbor) : rightRun,
                                bookmarkEnd);
                        }
                        if (inlineEquation)
                            display.AddAfterSelf(new XElement(w + "r", new XElement(w + "t", "Run before")),
                                owner.Elements(), new XElement(w + "r", new XElement(w + "t", "Run after")));
                        originalPictureRun.Remove();

                        XElement MathRun(params XElement?[] content) => new(m + "r",
                            new XElement(mathProperties), new XElement(wordProperties), content);
                    }, diagnosticSink: Record);

                if (!visible)
                {
                    TestAssert.True(!DocxBlockTraversal.EnumerateBodyParagraphs(source).Any(p => p.Images.Count != 0),
                        "The excluded revision cannot contribute a picture to layout.");
                }
                else
                {
                    CheckInlineOrder(source, packed);
                }
                var page = PictureRangeRenderer(mode).RenderPagesAsync(source, Record, default).ToBlockingEnumerable().Single();
                return (page, diagnostics, source.EquationParts.Sum(part => part.SourceCount));
            }

            void CheckInlineOrder(DocxDocument source, bool packed)
            {
                var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(source).Single(p => p.Images.Count != 0);
                var items = TestAssert.NotNull(paragraph.InlineContent).ToArray();
                string[] before = explicitRow ? ["17", "314"] : ["314"];
                string[] after = explicitRow ? ["793", "28"] : ["793"];
                int first = Array.FindIndex(items, item => Matches(item, before));
                int picture = Array.FindIndex(items, item => item is DocxInlinePicture);
                int last = Array.FindIndex(items, item => Matches(item, after));
                TestAssert.True(first >= 0 && first < picture && picture < last,
                    "Both operands and the drawing retain their source order in the paragraph.");
                TestAssert.Equal(withBreak, items[(first + 1)..picture].Any(item => item is DocxInlineBreak));
                if (revision is not null)
                {
                    TestAssert.Equal(revision, ((DocxInlineEquation)items[first]).Equation.Revision?.Kind);
                    TestAssert.Equal(revision, ((DocxInlineEquation)items[last]).Equation.Revision?.Kind);
                    var pictureRevision = ((DocxInlinePicture)items[picture]).Image.Revisions.Single(r => r.Id == "412");
                    TestAssert.Equal(revision, pictureRevision.Kind);
                }
                if (withBookmark && packed)
                {
                    var owner = ((DocxInlineEquation)items[first]).Equation.Content.Bookmarks.Single();
                    TestAssert.Equal(0, owner.ChildIndex);
                    var anchor = paragraph.BookmarkAnchors.Single(bookmark => bookmark.Name == "EquationDrawing");
                    TestAssert.True(ReferenceEquals(anchor.MathSource, owner.Source),
                        "The original bookmark remains bound to the mathematical source before the drawing.");
                }
                static bool Matches(DocxInlineItem item, string[] labels) => item is DocxInlineEquation math &&
                    math.Equation.Content.Children.Count == labels.Length &&
                    math.Equation.Content.Children.OfType<DocxMathRun>().Select(run => run.Text.Text).SequenceEqual(labels);
            }
        }
    }
}
