using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxStaticPictureLinksReuseIndependentCompoundFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent static picture origins use installed Windows fonts.");
        CheckStaticPictureLinks(null, compareWord: true);
    }

    internal static void CheckStaticPictureLinks(IFontResolver? resolver, bool compareWord)
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "Composition"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "composed-picture-navigation-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var renderer = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var source = ReadMutation("Composition/composed-picture-navigation-report", mode, false, false);
            var parent = ReadMutation("Composition/composed-inline-frame-report", mode, false, false);
            var pages = renderer.RenderBlankPages(source, null, default).ToArray();
            var before = renderer.RenderBlankPages(parent, null, default).ToArray();
            TestAssert.Equal(row.GetProperty("WordPages").GetInt32(), pages.Length);
            TestAssert.Equal(before.Length, pages.Length);
            var targets = row.GetProperty("TargetUris").EnumerateArray().Select(item => item.GetString()!).ToHashSet(StringComparer.Ordinal);
            for (int index = 0; index < pages.Length; index++)
            {
                var page = pages[index];
                TestAssert.Equal(before[index].Content, page.Content);
                TestAssert.Equal(before[index].Images.Count, page.Images.Count);
                var links = page.Annotations.Where(link => link.Uri is { } uri && targets.Contains(uri)).ToArray();
                var expected = row.GetProperty("WordLinks").EnumerateArray().Where(link => link.GetProperty("Page").GetInt32() == index + 1 &&
                    !link.GetProperty("Uri").GetString()!.Contains("/endnotes/", StringComparison.Ordinal)).ToArray();
                TestAssert.Equal(expected.Length, links.Length);
                foreach (var link in links)
                {
                    CheckPictureLinkDraw(page, link);
                    var word = expected.Single(item => item.GetProperty("Uri").GetString() == link.Uri).GetProperty("Rect");
                    if (compareWord && final)
                    {
                        double[] rect = [link.X, page.Height - link.Y - link.Height, link.X + link.Width, page.Height - link.Y];
                        for (int edge = 0; edge < rect.Length; edge++)
                            TestAssert.True(Math.Abs(rect[edge] - word[edge].GetDouble()) < 1d, "Final static URI bounds retain the original Word gate.");
                    }
                }
                TestAssert.Equal(before[index].Annotations.Count(link => link.IsUri), page.Annotations.Count(link => link.IsUri) - links.Length);
            }
        }
    }

    public static void DocxStaticPictureLinksKeepDeclaredClipsAndCancellation()
    {
        foreach (string fixture in new[] { "docx-static-image-header", "docx-static-image-footer", "docx-static-fixed-footer", "docx-static-mixed-fixed-footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", fixture + ".docx"));
            var source = AddStaticPictureTargets(DocxTests.ReadDocx(path, mode));
            var renderer = new DocxRenderer(new MapFontResolver([], "Courier"), mode, mode == OoxPdfDocxMarkupMode.Final ?
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderBlankPages(source, null, default).Single();
            TestAssert.True(page.Images.Count > 0, "A clipped picture is still decoded and emitted; navigation must respect its clip.");
            var links = page.Annotations.Where(link => link.Uri == "https://example.invalid/static-picture").ToArray();
            TestAssert.Equal(fixture.Contains("fixed", StringComparison.Ordinal) ? 0 : 1, links.Length);
            foreach (var link in links) CheckPictureLinkDraw(page, link);
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderBlankPages(source, null, cancellation.Token).ToArray());
        }
    }

    private static void CheckPictureLinkDraw(PdfPage page, PdfLinkAnnotation link)
    {
        bool matches = page.Images.Any(image =>
        {
            var draw = Regex.Match(page.Content, @"([-\d.]+) 0 0 ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\s+/" + Regex.Escape(image.ResourceName) + @" Do");
            if (!draw.Success) return false;
            double Number(int index) => double.Parse(draw.Groups[index].Value, CultureInfo.InvariantCulture);
            return Math.Abs(link.X - Number(3)) <= .001d && Math.Abs(link.Y - Number(4)) <= .001d &&
                Math.Abs(link.Width - Number(1)) <= .001d && Math.Abs(link.Height - Number(2)) <= .001d;
        });
        TestAssert.True(matches, "The clickable region follows an actual physical raster draw.");
    }

    private static DocxDocument AddStaticPictureTargets(DocxDocument source)
    {
        var cache = new Dictionary<DocxParagraph, DocxParagraph>(ReferenceEqualityComparer.Instance);
        DocxParagraph Rewrite(DocxParagraph paragraph)
        {
            if (cache.TryGetValue(paragraph, out var found)) return found;
            var images = new Dictionary<DocxInlineImage, DocxInlineImage>(ReferenceEqualityComparer.Instance);
            foreach (var image in paragraph.Images) images.Add(image, image with { ClickTarget = "https://example.invalid/static-picture" });
            return cache[paragraph] = paragraph with { Images = paragraph.Images.Select(image => images[image]).ToArray(),
                InlineContent = paragraph.InlineContent?.Select(item => item is DocxInlinePicture picture ? new DocxInlinePicture(images[picture.Image]) : item).ToArray() };
        }
        IReadOnlyList<DocxParagraph> Paragraphs(IReadOnlyList<DocxParagraph> paragraphs) => paragraphs.Select(Rewrite).ToArray();
        IReadOnlyList<DocxBodyElement> Blocks(IReadOnlyList<DocxBodyElement> blocks) => blocks.Select(item =>
            item is DocxParagraphElement paragraph ? new DocxParagraphElement(Rewrite(paragraph.Paragraph)) : item).ToArray();
        return source with
        {
            HeaderParagraphs = Paragraphs(source.HeaderParagraphs), FooterParagraphs = Paragraphs(source.FooterParagraphs),
            HeaderParagraphsByType = source.HeaderParagraphsByType.ToDictionary(item => item.Key, item => Paragraphs(item.Value)),
            FooterParagraphsByType = source.FooterParagraphsByType.ToDictionary(item => item.Key, item => Paragraphs(item.Value)),
            HeaderBodyElementsByType = source.HeaderBodyElementsByType.ToDictionary(item => item.Key, item => Blocks(item.Value)),
            FooterBodyElementsByType = source.FooterBodyElementsByType.ToDictionary(item => item.Key, item => Blocks(item.Value))
        };
    }
}
