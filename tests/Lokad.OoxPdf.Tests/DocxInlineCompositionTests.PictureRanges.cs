using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxMixedPictureRangeKeepsIndependentWordBounds()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("The independent Word control uses installed Windows fonts.");
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "mixed-picture-range-office.json")));
        string input = Path.Combine(directory, "mixed-picture-equation-link-cell.docx");
        TestAssert.Equal(gold.RootElement.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        var page = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderPagesAsync(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final), null, default).ToBlockingEnumerable().Single();
        var link = page.Annotations.Single();
        var rect = gold.RootElement.GetProperty("WordRect");
        TestAssert.True(Math.Abs(link.X - rect[0].GetDouble()) < 1d &&
            Math.Abs(link.X + link.Width - rect[2].GetDouble()) < 1d &&
            Math.Abs(page.Height - link.Y - link.Height - rect[1].GetDouble()) < 1d &&
            Math.Abs(page.Height - link.Y - rect[3].GetDouble()) < 1d,
            "All four mixed-link edges match the independent one-point Word rectangle.");
    }

    public static void DocxMixedPictureRangesRetainAuthoredAfterSpacingAcrossContainers()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int height in new[] { 6, 18, 36 })
        {
            var renderer = PictureRangeRenderer(mode);
            var control = renderer.RenderPagesAsync(ReadPictureRange(container, mode, height,
                includePicture: true, afterSpacing: 0), null, default).ToBlockingEnumerable().Single();
            var page = renderer.RenderPagesAsync(ReadPictureRange(container, mode, height,
                includePicture: true, afterSpacing: 9), null, default).ToBlockingEnumerable().Single();
            var controlImage = ReadPictureRangeFrame(control);
            var image = ReadPictureRangeFrame(page);
            double controlBottom = control.Annotations.Where(link => link.Destination is not null).Min(link => link.Y);
            double bottom = page.Annotations.Where(link => link.Destination is not null).Min(link => link.Y);
            // Compare against painted placement: a footer can move as its complete
            // story grows. The last link still owns the authored paragraph gap.
            TestAssert.True(Math.Abs((controlBottom - controlImage.Y) - (bottom - image.Y) - 9d) < .01d,
                container + "/" + mode + "/" + height + ": the last link retains its nine-point after-gap.");
            TestAssert.Equal(control.Annotations.Count, page.Annotations.Count);
            TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(control), DocxPageFootnoteNumberingTests.DecodeText(page));
        }
    }

    public static void DocxMixedPictureRangesIncludePaintAcrossContainersAndPresets()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int height in new[] { 6, 18, 36 })
        {
            try
            {
                var renderer = PictureRangeRenderer(mode);
                var control = renderer.RenderPagesAsync(ReadPictureRange(container, mode, height, includePicture: false), null, default)
                    .ToBlockingEnumerable().Single();
                var page = renderer.RenderPagesAsync(ReadPictureRange(container, mode, height, includePicture: true), null, default)
                    .ToBlockingEnumerable().Single();
                TestAssert.Equal(control.Content, page.Content);
                TestAssert.Equal(control.Images.Count, page.Images.Count);
                TestAssert.Equal(control.Annotations.Count, page.Annotations.Count);
                var frame = ReadPictureRangeFrame(page);
                var covered = page.Annotations.Select((link, index) => (Link: link, Index: index)).Where(pair =>
                    pair.Link.X <= frame.X && pair.Link.X + pair.Link.Width >= frame.X + frame.Width &&
                    pair.Link.Y <= frame.Y + .001d && pair.Link.Y + pair.Link.Height >= frame.Y + frame.Height - .001d).ToArray();
                TestAssert.True(covered.Length == 1, "One existing link covers its actually painted picture.");
                var (link, index) = covered.Single();
                TestAssert.Equal(control.Annotations[index].Destination, link.Destination);
                TestAssert.True(link.X < control.Annotations[index].X - 20d,
                    "Extending the source range makes its preceding picture clickable.");
                // A formula may wrap into further cell lines. Those link rectangles
                // remain independently placed and must not absorb this picture.
                for (int other = 0; other < page.Annotations.Count; other++)
                    if (other != index) TestAssert.Equal(control.Annotations[other], page.Annotations[other]);
                TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(control), DocxPageFootnoteNumberingTests.DecodeText(page));
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(container + "/" + mode + "/" + height + ": " + exception, exception);
            }
        }
    }

    public static void DocxMixedPictureRangesKeepDecodeRotationClipAndChildTargetFallbacks()
    {
        foreach (string variant in new[] { "decode", "rotation", "clip", "child-uri", "missing-target" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string container = variant == "clip" ? "cell" : "body";
            var renderer = PictureRangeRenderer(mode);
            var controlSource = ReadPictureRange(container, mode, 18, includePicture: false, variant);
            var source = ReadPictureRange(container, mode, 18, includePicture: true, variant);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var control = renderer.RenderPagesAsync(controlSource, null, default).ToBlockingEnumerable().Single();
            var page = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(control.Content, page.Content);
            TestAssert.Equal(control.Annotations.Count, page.Annotations.Count);
            for (int index = 0; index < page.Annotations.Count; index++)
                TestAssert.Equal(control.Annotations[index], page.Annotations[index]);
            if (variant == "decode")
            {
                TestAssert.Equal(0, page.Images.Count);
                TestAssert.Equal(1, diagnostics.Count(d => d.Id == "IMAGE_UNSUPPORTED_FORMAT"));
            }
            if (variant == "child-uri")
            {
                TestAssert.Equal(2, page.Annotations.Count);
                TestAssert.True(page.Annotations[0].Destination is not null &&
                    page.Annotations[1].Uri == "https://example.invalid/child-picture",
                    "The child drawing keeps its own target and annotation order.");
            }
            if (variant == "missing-target") TestAssert.Equal(0, page.Annotations.Count);
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, canceled.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxPictureOnlySourceRangesRetainUnlinkedFallback()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadPictureRange(container, mode, 18, includePicture: true, "picture-only");
            var page = PictureRangeRenderer(mode).RenderPagesAsync(source, null, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(1, page.Images.Count);
            TestAssert.Equal(0, page.Annotations.Count);
        }
    }

    public static void DocxNoteCellPictureRangesFollowExistingPaintAndTargets()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (int height in new[] { 6, 18, 36 })
        {
            var renderer = PictureRangeRenderer(mode);
            var control = renderer.RenderPagesAsync(ReadPictureRange(kind, mode, height, includePicture: false), null, default)
                .ToBlockingEnumerable().ToArray();
            var pages = renderer.RenderPagesAsync(ReadPictureRange(kind, mode, height, includePicture: true), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(control.Length, pages.Length);
            for (int index = 0; index < pages.Length; index++)
            {
                var page = pages[index];
                TestAssert.Equal(control[index].Content, page.Content);
                TestAssert.Equal(control[index].Annotations.Count, page.Annotations.Count);
                if (page.Images.Count == 0)
                {
                    TestAssert.True(control[index].Annotations.SequenceEqual(page.Annotations), "Other note pages keep their navigation.");
                    continue;
                }
                var frame = ReadPictureRangeFrame(page);
                var covered = page.Annotations.Select((link, number) => (Link: link, Number: number)).Where(pair =>
                    pair.Link.X <= frame.X && pair.Link.X + pair.Link.Width >= frame.X + frame.Width &&
                    pair.Link.Y <= frame.Y + .001d && pair.Link.Y + pair.Link.Height >= frame.Y + frame.Height - .001d).ToArray();
                TestAssert.True(covered.Length == 1, "The complete note cell's source link covers its painted picture: " + kind + "/" + mode + "/" + height);
                var (link, number) = covered.Single();
                TestAssert.Equal(control[index].Annotations[number].Destination, link.Destination);
                for (int other = 0; other < page.Annotations.Count; other++)
                    if (other != number) TestAssert.Equal(control[index].Annotations[other], page.Annotations[other]);
            }
        }
    }

    public static void DocxNoteCellPictureRangesRetainClippedAndMergedFallbacks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string variant in new[] { "clip", "merge" })
        {
            var renderer = PictureRangeRenderer(mode);
            var control = renderer.RenderPagesAsync(ReadPictureRange(kind, mode, 18, includePicture: false, variant), null, default)
                .ToBlockingEnumerable().ToArray();
            var source = ReadPictureRange(kind, mode, 18, includePicture: true, variant);
            var pages = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(control.Length, pages.Length);
            TestAssert.Equal(1, pages.Sum(page => page.Images.Count));
            for (int index = 0; index < pages.Length; index++)
            {
                TestAssert.Equal(control[index].Content, pages[index].Content);
                TestAssert.True(control[index].Annotations.SequenceEqual(pages[index].Annotations),
                    "An unresolved or clipped note-cell owner retains its prior link bounds: " + kind + "/" + mode + "/" + variant);
            }
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(source, null, canceled.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    private static DocxRenderer PictureRangeRenderer(OoxPdfDocxMarkupMode mode) =>
        new(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, mode == OoxPdfDocxMarkupMode.Final
            ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

    private static PdfRectangle ReadPictureRangeFrame(PdfPage page)
    {
        var draw = Regex.Match(page.Content, @"([-\d.]+) 0 0 ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\s+/" +
            Regex.Escape(page.Images.Single().ResourceName) + @" Do");
        TestAssert.True(draw.Success, "The source range is compared with the actual axis-aligned image draw.");
        double Number(int index) => double.Parse(draw.Groups[index].Value, CultureInfo.InvariantCulture);
        return new(Number(3), Number(4), Number(1), Number(2));
    }

    private static DocxDocument ReadPictureRange(string container, OoxPdfDocxMarkupMode mode, int height,
        bool includePicture, string variant = "ordinary", int? afterSpacing = null)
    {
        bool note = container is "footnote" or "endnote";
        var source = ReadMutation("CompositionProbes/" + (note ? "mixed-note-" + container + "-table" : container == "cell" ? "body" : container), mode, false, false,
            (xml, w, m) =>
            {
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                var paragraph = xml.Descendants(w + "p").Single(p => p.Elements(m + "oMath").Any());
                if (afterSpacing is int after)
                {
                    var properties = paragraph.Element(w + "pPr");
                    if (properties is null) { properties = new XElement(w + "pPr"); paragraph.AddFirst(properties); }
                    var spacing = properties.Element(w + "spacing");
                    if (spacing is null) { spacing = new XElement(w + "spacing"); properties.Add(spacing); }
                    spacing.SetAttributeValue(w + "after", (after * 20).ToString(CultureInfo.InvariantCulture));
                }
                string target = paragraph.Elements(w + "bookmarkStart").First().Attribute(w + "name")!.Value;
                var picture = paragraph.Elements(w + "r").Single(run => run.Descendants(w + "drawing").Any());
                foreach (var extent in picture.Descendants().Where(e => e.Name.LocalName is "extent" or "ext" && e.Attribute("cy") is not null))
                    extent.SetAttributeValue("cy", height * 12700);
                if (variant == "rotation") picture.Descendants(a + "xfrm").First().SetAttributeValue("rot", "5400000");
                var equation = paragraph.Elements(m + "oMath").Last();
                var link = new XElement(w + "hyperlink", new XAttribute(w + "anchor", variant == "missing-target" ? "AbsentTarget" : target));
                if (variant == "picture-only")
                {
                    picture.ReplaceWith(link);
                    link.Add(picture);
                }
                else
                {
                    equation.ReplaceWith(link);
                    if (includePicture) { picture.Remove(); link.Add(picture); }
                    link.Add(equation);
                }
                if (note && variant == "clip")
                {
                    var row = paragraph.Ancestors(w + "tr").Single();
                    var properties = row.Element(w + "trPr");
                    if (properties is null) { properties = new XElement(w + "trPr"); row.AddFirst(properties); }
                    properties.Elements(w + "trHeight").Remove();
                    properties.Add(new XElement(w + "trHeight", new XAttribute(w + "val", "20"), new XAttribute(w + "hRule", "exact")));
                }
                if (note && variant == "merge")
                {
                    var cell = paragraph.Ancestors(w + "tc").Single();
                    var properties = cell.Element(w + "tcPr");
                    if (properties is null) { properties = new XElement(w + "tcPr"); cell.AddFirst(properties); }
                    properties.Add(new XElement(w + "vMerge", new XAttribute(w + "val", "restart")));
                }
                if (container == "cell") paragraph.ReplaceWith(new XElement(w + "tbl",
                    new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "8640"))),
                    new XElement(w + "tr", variant == "clip" ? new XElement(w + "trPr", new XElement(w + "trHeight",
                        new XAttribute(w + "val", "20"), new XAttribute(w + "hRule", "exact"))) : null,
                        new XElement(w + "tc", new XElement(paragraph)))));
            }, partOverride: note ? "word/" + container + "s.xml" : container == "header" ? "word/header1.xml" : container == "footer" ? "word/footer1.xml" : null);
        if (variant is not ("decode" or "child-uri")) return source;
        var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(source.BodyElements).Single(p => p.Images.Count != 0);
        var image = variant == "decode" ? paragraph.Images.Single() with { Bytes = [0, 1, 2, 3] }
            : paragraph.Images.Single() with { ClickTarget = "https://example.invalid/child-picture" };
        var replacement = paragraph with { Images = [image], InlineContent = TestAssert.NotNull(paragraph.InlineContent)
            .Select(item => item is DocxInlinePicture ? new DocxInlinePicture(image) : item).ToArray() };
        return source with { BodyElements = source.BodyElements.Select(item => item is DocxParagraphElement p &&
            ReferenceEquals(p.Paragraph, paragraph) ? new DocxParagraphElement(replacement) : item).ToArray(),
            FallbackParagraphs = source.FallbackParagraphs.Select(p => ReferenceEquals(p, paragraph) ? replacement : p).ToArray() };
    }
}
