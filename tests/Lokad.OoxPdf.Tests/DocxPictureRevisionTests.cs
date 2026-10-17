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
