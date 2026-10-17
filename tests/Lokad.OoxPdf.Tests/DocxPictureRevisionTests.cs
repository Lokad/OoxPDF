using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using System.Text;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class DocxPictureRevisionTests
{
    public static void DocxPictureRevisionBodyCellFramesMatchIndependentWordDimensions()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native picture controls use installed Windows fonts.");
        CheckBodyCellPictureFrames(null);
    }

    internal static void CheckBodyCellPictureFrames(IFontResolver? resolver)
    {
        using var single = ReadGolden();
        using var shared = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases(), "PictureRevisions", "shared-picture-runs-office.json")));
        using var standalone = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases(), "PictureRevisions", "picture-size-variants-office.json")));
        var rows = single.RootElement.GetProperty("Rows").EnumerateArray()
            .Concat(shared.RootElement.GetProperty("Rows").EnumerateArray()).Concat(standalone.RootElement.GetProperty("Rows").EnumerateArray())
            .Where(row => row.GetProperty("Id").GetString() is { } id && !id.EndsWith("-header", StringComparison.Ordinal) && !id.EndsWith("-footer", StringComparison.Ordinal));
        foreach (var row in rows)
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string context = row.GetProperty("Id").GetString() + "/" + mode;
            var document = ReadSource(row, mode);
            var geometry = mode == OoxPdfDocxMarkupMode.AllMarkup
                ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
            var page = new DocxRenderer(resolver, mode, geometry).RenderBlankPages(document, null, default).Single();
            int expected = row.TryGetProperty("ImageClips", out var clips) ? clips.GetArrayLength() : 1;
            TestAssert.Equal(expected, page.Images.Count);
            var officeClip = expected == 1 ? row.GetProperty("ImageClip") : clips[0];
            double scale = mode == OoxPdfDocxMarkupMode.AllMarkup ? 1d : 1d / .7586666666666667d;
            foreach (var resource in page.Images)
            {
                var picture = Regex.Match(page.Content, @"(?<w>[-\d.]+) 0 0 (?<h>[-\d.]+) (?<x>[-\d.]+) (?<y>[-\d.]+) cm\r?\n/" + Regex.Escape(resource.ResourceName) + " Do");
                TestAssert.True(picture.Success, "Both presets must retain their source picture: " + context);
                double x = Number(picture, "x"), y = Number(picture, "y"), width = Number(picture, "w"), height = Number(picture, "h");
                TestAssert.True(Math.Abs(width - (officeClip[2].GetDouble() - officeClip[0].GetDouble()) * scale) < .1d &&
                    Math.Abs(height - (officeClip[3].GetDouble() - officeClip[1].GetDouble()) * scale) < .1d,
                    "Wrapping and painting must share the source picture's print scale: " + context);
                TestAssert.True(x >= 0d && y >= 0d && x + width <= page.Width && y + height <= page.Height,
                    "Body/cell picture ink must remain inside the physical page: " + context);
            }
            TestAssert.Contains("Public comment anchor.", ReadControlText(page));
            if (resolver is null && mode == OoxPdfDocxMarkupMode.Final && row.TryGetProperty("FinalFollowingTextBaselineY", out var finalBaseline))
            {
                var following = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true)
                    .Single(show => show.Text.Trim() == "Public comment anchor.");
                TestAssert.True(Math.Abs(following.Y - finalBaseline.GetDouble()) < .75d,
                    "Source-positioned image paragraphs must preserve Word Final's following baseline: " + context);
            }
            if (resolver is null && mode == OoxPdfDocxMarkupMode.AllMarkup)
            {
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true).ToArray();
                foreach (var pair in new[] { (Text: "Public owner paragraph.", Key: "OwnerTextBaselineY"), (Text: "Public comment anchor.", Key: "FollowingTextBaselineY") })
                {
                    if (row.GetProperty(pair.Key).ValueKind == JsonValueKind.Null) continue;
                    var show = shows.Single(show => show.Text.Trim() == pair.Text);
                    TestAssert.True(Math.Abs(show.Y - row.GetProperty(pair.Key).GetDouble()) < .75d,
                        "Picture scaling must preserve the independent source-text baseline: " + context + ", " + pair.Text + ", y=" + show.Y.ToString(CultureInfo.InvariantCulture));
                }
            }
        }
    }

    public static void DocxPictureRevisionEligibilityRequiresActualImageOwnership()
    {
        using var golden = ReadGolden();
        var row = golden.RootElement.GetProperty("Rows")[0];
        var source = ReadSource(row, OoxPdfDocxMarkupMode.AllMarkup);
        var paragraph = DocxBlockTraversal.EnumerateSourceParagraphs(source).First(paragraph => paragraph.Images.Count != 0);
        TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(paragraph),
            "A retained inserted image owns its source even without text in the same run.");
        TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default,
                paragraph.Revisions.Select(revision => revision with { }).ToArray()),
            "Value-equal aggregate revision metadata must not invent image ownership.");
        TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with
            { Images = paragraph.Images.Select(image => image with { SourceRunStyle = null }).ToArray() }, default),
            "An image without retained source style must keep its conservative fallback.");
        TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with
            { Images = paragraph.Images.Select(image => image with { SourceRunIndex = -1 }).ToArray() }, default),
            "A style from another source position must not claim the picture's revisions.");
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, canceled.Token));
    }

    public static void DocxPictureRevisionStaticFramesKeepContentInsideThePage()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent static geometry uses installed Windows fonts.");
        CheckStaticPictureFrames(null, compareGeometry: true);
    }

    internal static void CheckStaticPictureFrames(IFontResolver? resolver, bool compareGeometry)
    {
        using var single = ReadGolden();
        using var shared = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases(), "PictureRevisions", "shared-picture-runs-office.json")));
        var rows = single.RootElement.GetProperty("Rows").EnumerateArray()
            .Concat(shared.RootElement.GetProperty("Rows").EnumerateArray())
            .Where(row => row.GetProperty("Id").GetString() is { } id && (id.EndsWith("-header", StringComparison.Ordinal) || id.EndsWith("-footer", StringComparison.Ordinal)));
        foreach (var row in rows)
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string context = row.GetProperty("Id").GetString() + "/" + mode;
            var document = ReadSource(row, mode);
            var geometry = mode == OoxPdfDocxMarkupMode.AllMarkup
                ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
            var page = new DocxRenderer(resolver, mode, geometry).RenderBlankPages(document, null, default).Single();
            int expected = row.TryGetProperty("ImageClips", out var clips) ? clips.GetArrayLength() : 1;
            TestAssert.Equal(expected, page.Images.Count);
            var officeClip = expected == 1 ? row.GetProperty("ImageClip") : clips[0];
            foreach (var resource in page.Images)
            {
                var picture = Regex.Match(page.Content, @"(?<w>[-\d.]+) 0 0 (?<h>[-\d.]+) (?<x>[-\d.]+) (?<y>[-\d.]+) cm\r?\n/" + Regex.Escape(resource.ResourceName) + " Do");
                TestAssert.True(picture.Success, "The source picture must be painted: " + context);
                double x = Number(picture, "x"), y = Number(picture, "y"), width = Number(picture, "w"), height = Number(picture, "h");
                TestAssert.True(x >= 0d && y >= 0d && x + width <= page.Width && y + height <= page.Height,
                    "Static picture ink must remain inside the physical page: " + context);
                if (mode == OoxPdfDocxMarkupMode.AllMarkup)
                {
                    TestAssert.True(Math.Abs(width - officeClip[2].GetDouble() + officeClip[0].GetDouble()) < .1d &&
                        Math.Abs(height - officeClip[3].GetDouble() + officeClip[1].GetDouble()) < .1d,
                        "Review picture dimensions must use Word's complete static frame scale: " + context);
                    if (compareGeometry)
                        TestAssert.True(Math.Abs(y - officeClip[1].GetDouble()) < 1.5d,
                            "Static picture origin must retain the independently measured story frame: " + context + ", y=" + y.ToString(CultureInfo.InvariantCulture));
                }
            }
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true).ToArray();
            TestAssert.True(shows.Where(show => show.Text.Contains("Public owner", StringComparison.Ordinal) || show.Text.Contains("Attached source.", StringComparison.Ordinal))
                .All(show => show.Y > 0d && show.Y < page.Height), "Source static text must remain within the viewport: " + context);
            TestAssert.Contains("Public owner paragraph.", ReadControlText(page));
        }
    }

    public static void DocxPictureRevisionStylesBelongToTheImageRun()
    {
        using var golden = ReadGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadSource(row, mode);
            var image = DocxBlockTraversal.EnumerateSourceParagraphs(source).SelectMany(paragraph => paragraph.Images)
                .Distinct<DocxInlineImage>(ReferenceEqualityComparer.Instance).Single();
            var style = image.SourceRunStyle;
            string variant = row.GetProperty("Variant").GetString()!;
            TestAssert.True(style is not null && style.SourceRunIndex == image.SourceRunIndex &&
                ReferenceEquals(style.Revision, image.Revisions.Single()), "The inserted picture must retain its actual owning revision and run position.");
            TestAssert.Equal(string.Empty, style!.Text);
            TestAssert.Equal(variant == "arial14" ? "Arial" : "Calibri", style.FontFamily!);
            TestAssert.Equal(variant == "arial14" ? 14d : variant == "calibri18" ? 18d : 11d, style.FontSize);
        }
    }

    public static void DocxPictureRevisionOneRunKeepsTextAndBothPictures()
    {
        string input = Path.Combine(Cases(), "ReviewAuthors", "nontext-author-image-insertion.docx");
        using var bytes = new MemoryStream();
        using (var source = ZipFile.OpenRead(input))
        using (var target = new ZipArchive(bytes, ZipArchiveMode.Create, leaveOpen: true))
        foreach (var entry in source.Entries)
        {
            using var from = entry.Open();
            using var to = target.CreateEntry(entry.FullName).Open();
            if (entry.FullName != "word/document.xml") { from.CopyTo(to); continue; }
            var xml = XDocument.Load(from);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var run = xml.Descendants(w + "r").Single(node => node.Element(w + "drawing") is not null);
            var copy = new XElement(run.Element(w + "drawing")!);
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            copy.Descendants(wp + "docPr").Single().SetAttributeValue("id", "2");
            run.Add(copy, new XElement(w + "t", " Attached source."));
            xml.Save(to);
        }
        bytes.Position = 0;
        var document = new DocxReader().Read(OoxPackage.Open(bytes, default), null, default, markupMode: OoxPdfDocxMarkupMode.AllMarkup);
        var paragraph = document.BodyElements.OfType<DocxParagraphElement>().First().Paragraph;
        var owner = paragraph.Runs.Single(run => run.Text.Contains("Attached source.", StringComparison.Ordinal));
        TestAssert.Equal(2, paragraph.Images.Count);
        TestAssert.True(paragraph.Images.All(image => ReferenceEquals(image.SourceRunStyle, owner)),
            "Pictures and text from the same source run must share their resolved owning style.");
        var page = new DocxRenderer(new MissingPictureFontResolver(), OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).Single();
        TestAssert.Equal(2, page.Images.Count);
        TestAssert.True(ReferenceEquals(page.Images[0].Image, page.Images[1].Image),
            "Repeated source pictures must retain the existing decoded-image cache owner.");
        TestAssert.Contains("Attached source.", ReadControlText(page));
        double previousEnd = double.NegativeInfinity;
        foreach (var resource in page.Images)
        {
            var bitmap = Regex.Match(page.Content, @"(?<w>[-\d.]+) 0 0 (?<h>[-\d.]+) (?<x>[-\d.]+) (?<y>[-\d.]+) cm\r?\n/" + Regex.Escape(resource.ResourceName) + " Do");
            TestAssert.True(bitmap.Success, "Both source pictures must be painted.");
            TestAssert.True(Number(bitmap, "x") >= previousEnd - .03d,
                "Two adjacent source pictures must not overlap or replace one another.");
            previousEnd = Number(bitmap, "x") + Number(bitmap, "w");
            bool marked = Regex.Matches(page.Content, @"0\.82 0\.204 0\.22 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f")
                .Cast<Match>().Any(rect => Math.Abs(Number(rect, "x") - Number(bitmap, "x")) < .03d &&
                    Math.Abs(Number(rect, "w") - Number(bitmap, "w")) < .03d && Number(rect, "y") < Number(bitmap, "y"));
            TestAssert.True(marked, "Both inserted pictures must keep their own visible revision indicator.");
        }
    }

    public static void DocxPictureRevisionSharedRunsMatchIndependentWordPlacement()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent metrics use installed Windows fonts.");
        CheckSharedPictures(null);
    }

    internal static void CheckSharedPictures(IFontResolver? resolver)
    {
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases(), "PictureRevisions", "shared-picture-runs-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string context = row.GetProperty("Id").GetString() + "/" + mode;
            var source = ReadSource(row, mode);
            var geometry = mode == OoxPdfDocxMarkupMode.AllMarkup
                ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
            var page = new DocxRenderer(resolver, mode, geometry).RenderBlankPages(source, null, default).Single();
            TestAssert.Equal(2, page.Images.Count);
            var pictures = page.Images.Select(resource => Regex.Match(page.Content,
                @"(?<w>[-\d.]+) 0 0 (?<h>[-\d.]+) (?<x>[-\d.]+) (?<y>[-\d.]+) cm\r?\n/" + Regex.Escape(resource.ResourceName) + " Do")).ToArray();
            TestAssert.True(pictures.All(picture => picture.Success), "Both inline pictures must be painted: " + context);
            double firstX = Number(pictures[0], "x"), width = Number(pictures[0], "w"), secondX = Number(pictures[1], "x");
            TestAssert.True(Math.Abs(secondX - firstX - width) < .03d,
                "Adjacent pictures must retain one picture advance without overlap: " + context);
            // Word scales picture frames in its print-review preview. That existing
            // size difference is separate from the ordering/overlap contract here.
            var officeClip = row.GetProperty("ImageClips")[0];
            double officeWidth = officeClip[2].GetDouble() - officeClip[0].GetDouble();
            TestAssert.True(Math.Abs((secondX - firstX) / width - row.GetProperty("FirstToSecondAdvance").GetDouble() / officeWidth) < .001d,
                "Each picture must advance by its frame width, as independent Word clips establish: " + context);
            var attached = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true)
                .Where(show => show.Text.Contains("Attached source.", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(1, attached.Length);
            TestAssert.True(attached[0].X >= secondX + Number(pictures[1], "w") - .03d,
                "Following text must clear both source pictures: " + context);
            TestAssert.Contains("Public owner paragraph.", ReadControlText(page));
        }
    }

    public static void DocxPictureRevisionMarkersMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent metrics use installed Windows fonts.");
        CheckRendering(null, compareGeometry: true);
    }

    public static void DocxPictureRevisionFallbackRetainsTheImageAndAuthorInk() => CheckRendering(new MissingPictureFontResolver(), false);

    internal static void CheckRendering(IFontResolver? resolver, bool compareGeometry)
    {
        using var golden = ReadGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var source = ReadSource(row, OoxPdfDocxMarkupMode.AllMarkup);
            var imageOwners = DocxBlockTraversal.EnumerateSourceParagraphs(source).SelectMany(paragraph => paragraph.Images)
                .Distinct<DocxInlineImage>(ReferenceEqualityComparer.Instance).ToArray();
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(source, null, default).Single();
            TestAssert.Equal(1, page.Images.Count);
            string name = Regex.Escape(page.Images.Single().ResourceName);
            Match bitmap = Regex.Match(page.Content, @"(?<w>[-\d.]+) 0 0 (?<h>[-\d.]+) (?<x>[-\d.]+) (?<y>[-\d.]+) cm\r?\n/" + name + " Do");
            TestAssert.True(bitmap.Success, "The source PNG must survive real body/cell/static painting.");
            double x = Number(bitmap, "x"), y = Number(bitmap, "y"), width = Number(bitmap, "w");
            var rgb = row.GetProperty("FillRgb").EnumerateArray().Select(value => value.GetByte()).ToArray();
            string ink = Regex.Escape(DocxTests.FormatPdfRgb((rgb[0], rgb[1], rgb[2]), "rg"));
            var bars = Regex.Matches(page.Content, ink + @"\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f")
                .Cast<Match>().Where(rect => Math.Abs(Number(rect, "x") - x) < .03d &&
                    Math.Abs(Number(rect, "w") - width) < .03d && Number(rect, "y") < y && Number(rect, "h") > 0d).ToArray();
            TestAssert.Equal(1, bars.Length);
            if (compareGeometry)
            {
                var bar = bars.Single();
                TestAssert.True(Math.Abs((Number(bar, "x") - x) - row.GetProperty("OffsetX").GetDouble()) < .25d,
                    "Picture indicator X must follow its independently measured image frame.");
                TestAssert.True(Math.Abs((Number(bar, "y") - y) - row.GetProperty("OffsetY").GetDouble()) < .75d,
                    "Picture indicator Y must use the owning run's independently measured font metrics: " + row.GetProperty("Id").GetString());
                TestAssert.True(Math.Abs(Number(bar, "h") - row.GetProperty("Marker")[3].GetDouble()) < .15d,
                    "Picture indicator thickness must follow the owning font, including the 14/18-point controls.");
                if (row.GetProperty("Variant").GetString() == "arial14")
                    TestAssert.True(page.Fonts.All(resource => resource.Font.Font.FamilyName != "Arial"),
                        "A picture-only run needs font metrics without embedding an unused face.");
            }
            TestAssert.Contains("Public owner paragraph.", ReadControlText(page));
            TestAssert.True(imageOwners.All(image => DocxBlockTraversal.EnumerateSourceParagraphs(source)
                .SelectMany(paragraph => paragraph.Images).Any(current => ReferenceEquals(current, image))),
                "Picture decoration must not replace the source picture or its metadata.");
        }
    }

    private static string ReadControlText(PdfPage page)
    {
        var text = new StringBuilder();
        foreach (Match block in Regex.Matches(page.Content, @"(?s)BT\s+.*?\s+ET"))
        {
            string name = Regex.Match(block.Value, @"/(?<name>\S+) [\d.]+ Tf").Groups["name"].Value;
            if (page.Fonts.Any(resource => resource.ResourceName == name))
                foreach (var show in DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, block.Value, page.Fonts)))
                    text.Append(show.Text);
            else if (page.FallbackFonts.Any(resource => resource.ResourceName == name))
                foreach (Match glyph in Regex.Matches(block.Value, @"<(?<byte>[0-9A-F]{2})> Tj"))
                {
                    byte value = byte.Parse(glyph.Groups["byte"].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    TestAssert.True(value < 128, "These public fallback controls contain ASCII text only.");
                    text.Append((char)value);
                }
        }
        return text.ToString();
    }

    private static double Number(Match match, string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);
    private static JsonDocument ReadGolden() => JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases(), "PictureRevisions", "picture-revisions-office.json")));
    private static string Cases() => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases"));
    private static DocxDocument ReadSource(JsonElement row, OoxPdfDocxMarkupMode mode)
    {
        string relative = row.GetProperty("Input").GetString()!.Replace("tests/Lokad.OoxPdf.Tests/Cases/", string.Empty, StringComparison.Ordinal);
        string path = Path.Combine(Cases(), relative);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        return DocxTests.ReadDocx(path, mode);
    }

    private sealed class MissingPictureFontResolver : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, "Missing", new FontStyleKey(),
            new MemoryFontProgramSource("missing-picture-style", ReadOnlyMemory<byte>.Empty), IsFallback: false);
    }
}
