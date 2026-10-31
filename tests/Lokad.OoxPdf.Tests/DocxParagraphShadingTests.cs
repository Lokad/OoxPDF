using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxParagraphShadingTests
{
    private static string Cases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "ParagraphShading"));

    public static void DocxParagraphShadingRetainsDirectAndInheritedStyleValues()
    {
        var document = DocxTests.ReadDocx(Path.Combine(Cases, "flow.docx"), OoxPdfDocxMarkupMode.Final);
        var paragraphs = document.BodyElements.OfType<DocxParagraphElement>().Select(item => item.Paragraph)
            .ToDictionary(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)));
        TestAssert.Equal("224466", paragraphs["SIZE24"].Shading!.Value.FillHex);
        TestAssert.Equal("222222", paragraphs["INHERITED"].Shading!.Value.FillHex);
        TestAssert.Equal("222222", paragraphs["CLEAR"].Shading!.Value.FillHex);
        TestAssert.True(paragraphs["EMPTY_BEFORE"].Shading is null, "Unshaded paragraphs keep the containing background.");
        TestAssert.Equal(paragraphs["INHERITED"].Shading, paragraphs["INHERITED"].EffectiveProperties.Shading);
    }

    public static void DocxParagraphShadingPaintsWordColorsAcrossFlowCellsStaticAndMath()
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string input = Path.Combine(Cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            foreach (var reference in control.GetProperty("References").EnumerateArray())
            {
                var samples = reference.GetProperty("Colors").EnumerateArray().ToArray();
                bool review = reference.GetProperty("Mode").GetString() == "all-word-compatible";
                var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
                var document = DocxTests.ReadDocx(input, mode);
                var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                    review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
                var found = new Dictionary<string, (string Ink, string? Fill)>(StringComparer.Ordinal);
                int pictures = 0;
                bool pictureOnlyBackground = false;
                foreach (var page in renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable())
                {
                    var rectangles = ReadRectangles(page.Content).ToArray();
                    pictures += ImagePaintCount(page);
                    pictureOnlyBackground |= rectangles.Any(rect => rect.Color == "552233");
                    foreach (Match block in Regex.Matches(page.Content, @"(?s)BT\s+(.*?)\s+ET"))
                    {
                        var fragment = new PdfPage(page.Width, page.Height, block.Value, page.Fonts);
                        var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(fragment).ToArray();
                        string text = string.Concat(shows.Select(show => show.Text));
                        foreach (var sample in samples)
                        {
                            string label = sample.GetProperty("Label").GetString()!;
                            if (!text.Contains(label, StringComparison.Ordinal)) continue;
                            var show = shows.First();
                            double y = show.Y + show.Size * .25d;
                            var backing = rectangles.LastOrDefault(rect => show.X + .1d >= rect.X && show.X + .1d <= rect.X + rect.Width &&
                                y >= rect.Y && y <= rect.Y + rect.Height);
                            TestAssert.True(found.TryAdd(label, (DocxAutomaticColorTests.ReadInk(block.Value), backing?.Color)),
                                "Each source label is painted once: " + label);
                        }
                    }
                }
                if (control.GetProperty("Input").GetString() == "objects.docx")
                {
                    TestAssert.Equal(4, pictures);
                    TestAssert.True(pictureOnlyBackground, "A picture-only paragraph retains its own background without a synthetic text glyph.");
                }
                foreach (var sample in samples)
                {
                    string label = sample.GetProperty("Label").GetString()!;
                    TestAssert.True(found.TryGetValue(label, out var actual), "Source content survives paragraph shading: " + label);
                    TestAssert.Equal(sample.GetProperty("Ink").GetString(), actual.Ink);
                    TestAssert.Equal(sample.GetProperty("Fill").GetString(), actual.Fill);
                }
            }
        }
    }

    public static void DocxParagraphShadingUsesSourceIndentsAndFragmentLineBoxes()
    {
        var paragraph = DocxTests.CreateDocxLayoutParagraph(string.Concat(Enumerable.Repeat("word ", 100)), 12d, 18d,
            indent: DocxParagraphIndent.Empty with { LeftPoints = 18d, RightPoints = 36d, FirstLinePoints = 18d }) with
        {
            Shading = new("222222", "clear", null), SpacingAfterPoints = 9d,
            Spacing = DocxParagraphSpacing.Empty with { LineValue = "360", LineRuleValue = "exact", AfterValue = "180" }
        };
        var document = DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []);
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default);
        var lines = layout.Pages.SelectMany(page => page.Items).OfType<DocxTextLineLayout>().ToArray();
        TestAssert.True(layout.Pages.Count > 1 && lines.Length > 2, "A single shaded paragraph must survive page fragmentation.");
        for (int index = 0; index < lines.Length; index++)
        {
            var line = lines[index];
            TestAssert.True(line.ParagraphShadingFrame is not null, "Every placed source line owns its background frame.");
            var frame = line.ParagraphShadingFrame!.Value;
            TestAssert.True(Math.Abs(line.X + frame.OffsetX - (document.MarginLeftPoints + 18d - 1.44d)) < .001d,
                "First-line text indentation cannot narrow the paragraph background.");
            TestAssert.True(Math.Abs(frame.Width - (document.PageWidthPoints - document.MarginLeftPoints - document.MarginRightPoints - 18d - 36d + 2.88d)) < .001d,
                "The container width and paragraph indents own the background extent.");
            TestAssert.True(Math.Abs(frame.Height - 18d) < .001d,
                "A terminal paragraph keeps after-spacing outside every fragment background.");
        }
    }

    public static void DocxParagraphShadingAfterSpacingFollowsAdjacentShadedContent()
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "spacing-office.json")));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string input = Path.Combine(Cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default);
            var lines = layout.Pages.SelectMany(page => page.Items.OfType<DocxTextLineLayout>()
                .Concat(page.StaticTextLines)
                .Concat(page.Items.OfType<DocxTableRowLayout>().SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines)))
                .ToArray();
            foreach (var sample in control.GetProperty("Samples").EnumerateArray())
            {
                string label = sample.GetProperty("Label").GetString()!;
                var line = lines.Single(line => line.Text.Trim() == label);
                TestAssert.True(line.ParagraphShadingFrame is not null && line.LineHeight is not null,
                    "The placed line retains its paragraph paint frame: " + label);
                double paintedAfter = line.ParagraphShadingFrame!.Value.Height - line.LineHeight!.Value;
                TestAssert.True(Math.Abs(paintedAfter - sample.GetProperty("PaintedAfterSpacing").GetDouble()) < .001d,
                    "Word paints after-spacing only towards an adjacent shaded paragraph in the same container: " + label);
            }
        }
    }

    public static void DocxParagraphShadingBeforeSpacingFollowsPlacedContainerBoundaries()
    {
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "before-office.json")));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string input = Path.Combine(Cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final);
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default);
            var lines = layout.Pages.SelectMany(page => page.Items.OfType<DocxTextLineLayout>()
                .Concat(page.StaticTextLines)
                .Concat(page.Items.OfType<DocxTableRowLayout>().SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines)))
                .ToArray();
            foreach (var sample in control.GetProperty("Samples").EnumerateArray())
            {
                string label = sample.GetProperty("Label").GetString()!;
                var first = lines.Single(line => line.SourceLineIndex == 0 && line.Text.TrimStart().StartsWith(label, StringComparison.Ordinal));
                var paragraphLines = lines.Where(line => ReferenceEquals(line.SourceParagraph, first.SourceParagraph)).ToArray();
                foreach (var line in paragraphLines)
                {
                    double expected = ReferenceEquals(line, first) ? sample.GetProperty("PaintedBeforeSpacing").GetDouble() : 0d;
                    TestAssert.True(line.ParagraphShadingFrame is not null && line.LineHeight is not null,
                        "Each supported source line retains its paragraph background: " + label);
                    double extra = line.ParagraphShadingFrame!.Value.Height - line.LineHeight!.Value;
                    TestAssert.True(Math.Abs(extra - expected) < .001d,
                        "Only the first placed line joins excess spacing to a shaded predecessor in the same container: " + label);
                }
            }
        }
    }

    public static void DocxParagraphShadingBeforeSpacingPreservesPictureAndMathContainers()
    {
        var pictures = DocxTests.ReadDocx(Path.Combine(Cases, "objects.docx"), OoxPdfDocxMarkupMode.Final);
        var picture = pictures.BodyElements.OfType<DocxParagraphElement>().Select(item => item.Paragraph)
            .Single(p => p.Images.Count == 1 && p.Runs.All(run => run.Text.Length == 0));
        var math = DocxTests.ReadDocx(Path.Combine(Cases, "containers.docx"), OoxPdfDocxMarkupMode.Final)
            .BodyElements.OfType<DocxParagraphElement>().Select(item => item.Paragraph).Single(p => p.Equations.Count != 0);
        var previous = DocxTests.CreateDocxLayoutParagraph("PRECEDING", 12d, 13.8d) with
        {
            Shading = new("224466", "clear", null), SpacingAfterPoints = 18d,
            Spacing = DocxParagraphSpacing.Empty with { AfterValue = "360" }
        };
        foreach (var source in new[] { picture, math })
        foreach (bool inCell in new[] { false, true })
        {
            var rectangles = new List<PaintRectangle>();
            foreach (double before in new[] { 18d, 24d })
            {
                var paragraph = source with { Shading = new("552233", "clear", null), SpacingBeforePoints = before,
                    Spacing = source.Spacing with { BeforeValue = (before * 20d).ToString(CultureInfo.InvariantCulture) } };
                var cell = new DocxTableCell(string.Empty, [previous, paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
                var table = new DocxTable(null, [120d], [new DocxTableRow([cell], 90d)]);
                var document = inCell ? DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table])
                    : DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(previous), new DocxParagraphElement(paragraph)], []);
                var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: source.Equations.Count != 0),
                    OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
                var pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(source.Images.Count, pages.Sum(ImagePaintCount));
                rectangles.Add(pages.SelectMany(page => ReadRectangles(page.Content)).Single(rect => rect.Color == "552233"));
            }
            TestAssert.True(Math.Abs(rectangles[1].Height - rectangles[0].Height - 6d) < .001d,
                "Picture/math backgrounds include the same six-point excess spacing as prose.");
            TestAssert.True(Math.Abs(rectangles[1].Y + rectangles[1].Height - rectangles[0].Y - rectangles[0].Height) < .001d,
                "The shared background stays joined to the predecessor while its source object moves down.");
        }
    }

    public static void DocxParagraphShadingChangesPaintWithoutChangingTextGeometry()
    {
        var document = DocxTests.ReadDocx(Path.Combine(Cases, "flow.docx"), OoxPdfDocxMarkupMode.Final);
        var replacements = document.BodyElements.OfType<DocxParagraphElement>().Select(item => item.Paragraph)
            .ToDictionary(paragraph => paragraph, paragraph => paragraph with { Shading = null });
        var replace = new DocxParagraphReplacements(replacements, default);
        var plain = document with { BodyElements = replace.ReplaceElements(document.BodyElements),
            FallbackParagraphs = document.FallbackParagraphs.Select(replace.Replace).ToArray() };
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var shadedPages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
        var plainPages = renderer.RenderPagesAsync(plain, null, default).ToBlockingEnumerable().ToArray();
        TestAssert.Equal(plainPages.Length, shadedPages.Length);
        for (int page = 0; page < shadedPages.Length; page++)
        {
            var shaded = shadedPages[page];
            TestAssert.True(DocxCommentsTests.ReadEmbeddedGlyphTextShows(shaded).Select(show => (show.Text, show.Size, show.X, show.Y))
                .SequenceEqual(DocxCommentsTests.ReadEmbeddedGlyphTextShows(plainPages[page]).Select(show => (show.Text, show.Size, show.X, show.Y))),
                "Paragraph shading cannot change text, fonts, placement or pagination.");
            int firstText = shaded.Content.IndexOf("BT", StringComparison.Ordinal);
            TestAssert.True(ReadRectangles(shaded.Content).All(rect => rect.Position < firstText),
                "Paragraph backgrounds paint before glyphs so exact/overlapping lines cannot erase ink.");
        }
    }

    public static void DocxParagraphShadingSurvivesPictureOnlyCellPlacement()
    {
        var source = DocxTests.ReadDocx(Path.Combine(Cases, "objects.docx"), OoxPdfDocxMarkupMode.Final);
        var paragraph = source.BodyElements.OfType<DocxParagraphElement>().Select(item => item.Paragraph)
            .Single(p => p.Images.Count == 1 && p.Runs.All(run => run.Text.Length == 0));
        var cell = new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [120d], [new DocxTableRow([cell], 60d)]);
        foreach (var document in new[] {
            DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(paragraph)], []),
            DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]) })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(),
                OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, pages.Sum(ImagePaintCount));
            TestAssert.True(pages.SelectMany(page => ReadRectangles(page.Content)).Any(rect => rect.Color == "552233"),
                "Moving a supported picture-only paragraph into a cell must preserve its background and picture.");
        }
    }

    private static int ImagePaintCount(PdfPage page) => Regex.Matches(page.Content, @"/(?<name>[^\s/]+) Do")
        .Count(match => page.Images.Any(image => image.ResourceName == match.Groups["name"].Value));

    private sealed record PaintRectangle(string Color, double X, double Y, double Width, double Height, int Position);
    private static IEnumerable<PaintRectangle> ReadRectangles(string content)
    {
        foreach (Match match in Regex.Matches(content,
            @"(?m)^(?<r>[\d.]+)(?: (?<g>[\d.]+) (?<b>[\d.]+))? (?<op>rg|g)\r?\n(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f\r?$"))
        {
            double Read(string name) => double.Parse(match.Groups[name].Value, CultureInfo.InvariantCulture);
            byte red = (byte)Math.Round(Read("r") * 255d);
            bool gray = match.Groups["op"].Value == "g";
            byte green = gray ? red : (byte)Math.Round(Read("g") * 255d);
            byte blue = gray ? red : (byte)Math.Round(Read("b") * 255d);
            yield return new($"{red:X2}{green:X2}{blue:X2}", Read("x"), Read("y"), Read("w"), Read("h"), match.Index);
        }
    }
}
