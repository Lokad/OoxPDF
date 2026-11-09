using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxColumnAdmissionTests
{
    public static void DocxColumnAdmissionMatchesIndependentWordHorizontalOrigins()
    {
        if (!OperatingSystem.IsWindows() || !File.Exists(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "calibri.ttf")))
            TestAssert.Skip("Independent column origins require installed Windows fonts.");
        using var gold = Gold();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            bool review = row.GetProperty("Mode").GetString() != "final-preserve";
            var pages = Renderer(review).RenderPagesAsync(Read(row.GetProperty("Input").GetString()!, review), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            var chars = Characters(pages);
            TestAssert.Equal(row.GetProperty("ExpectedBodyText").GetString(), new string(chars.Select(c => c.Character).ToArray()));
            var anchors = row.GetProperty("Anchors");
            TestAssert.True(anchors.GetArrayLength() > 0, "Native X coverage requires real anchors.");
            foreach (var anchor in anchors.EnumerateArray())
            {
                var actual = chars[anchor.GetProperty("Offset").GetInt32()];
                TestAssert.True(actual.Character.ToString() == anchor.GetProperty("Character").GetString() &&
                    actual.Page == anchor.GetProperty("Page").GetInt32() &&
                    Math.Abs(actual.X - anchor.GetProperty("X").GetDouble()) <= 1d &&
                    Math.Abs(actual.Size - anchor.GetProperty("Size").GetDouble()) <= .1d,
                    $"Admitted column X matches Word: {row.GetProperty("Id").GetString()}/{actual}.");
                // Native Y is retained in the gold and full comparison reports;
                // minimum-leading placement in columns remains a separate fallback.
            }
        }
    }

    public static void DocxColumnAdmissionSharesLocalPlacementAcrossSpacingAndAlignment()
    {
        foreach (string rule in new[] { "atLeast", "exact", "auto" })
        foreach (string alignment in new[] { "left", "center", "right", "both" })
        foreach (bool indent in new[] { false, true })
        {
            void Properties(XDocument xml)
            {
                foreach (var p in Paragraphs(xml))
                {
                    var properties = p.Element(Word + "pPr")!;
                    properties.Add(new XElement(Word + "jc", new XAttribute(Word + "val", alignment)));
                    if (indent) properties.Add(new XElement(Word + "ind", new XAttribute(Word + "left", "240"), new XAttribute(Word + "firstLine", "120")));
                    var spacing = properties.Element(Word + "spacing")!;
                    spacing.SetAttributeValue(Word + "lineRule", rule);
                    spacing.SetAttributeValue(Word + "line", rule == "auto" ? "240" : "560");
                }
            }
            var source = Read("two-pages.docx", modify: Properties);
            var single = Read("two-pages.docx", modify: xml =>
            {
                Properties(xml);
                var section = xml.Descendants(Word + "sectPr").Single();
                section.Element(Word + "pgSz")!.SetAttributeValue(Word + "w", "7200");
                section.Element(Word + "cols")!.SetAttributeValue(Word + "num", "1");
            });
            var layout = Layout(source); var singleLayout = Layout(single);
            var admitted = FrameLines(layout).ToArray(); var control = FrameLines(singleLayout).ToArray();
            TestAssert.Equal(control.Length, admitted.Length);
            TestAssert.True(admitted.Any(line => line.Column > 0), "The control must actually cross a column boundary.");
            foreach (var pair in admitted.Zip(control))
            {
                TestAssert.Equal(pair.Second.Line.Text, pair.First.Line.Text);
                TestAssert.Equal(pair.Second.Line.SourceBlockIndex, pair.First.Line.SourceBlockIndex);
                TestAssert.Equal(pair.Second.Line.SourceLineIndex, pair.First.Line.SourceLineIndex);
                TestAssert.Equal(pair.Second.Line.Width, pair.First.Line.Width);
                TestAssert.True(Math.Abs(pair.First.Line.X - pair.First.FrameX - pair.Second.Line.X + pair.Second.FrameX) < .00001d,
                    "Identical column constraints retain the same local line placement after admission.");
                TestAssert.Equal(pair.First.Column, pair.First.Line.ColumnIndex);
            }
        }
    }

    public static void DocxColumnAdmissionKeepsRunSplitsAndNavigation()
    {
        foreach (bool review in new[] { false, true })
        {
            var source = Read("two-pages.docx", review);
            var original = Renderer(review, new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(source, null, default)
                .ToBlockingEnumerable().ToArray();
            foreach (bool navigation in new[] { false, true })
            {
                var changed = Read("two-pages.docx", review, xml =>
                {
                    var paragraphs = Paragraphs(xml); var target = paragraphs[4];
                    var run = target.Elements(Word + "r").First(); string text = run.Element(Word + "t")!.Value;
                    int split = text.IndexOf(' ') + 1;
                    var first = new XElement(run); first.Element(Word + "t")!.Value = text[..split];
                    var second = new XElement(run); second.Element(Word + "t")!.Value = text[split..];
                    first.Element(Word + "t")!.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                    second.Element(Word + "t")!.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                    run.ReplaceWith(first, second);
                    if (!navigation) return;
                    target.Element(Word + "pPr")!.AddAfterSelf(new XElement(Word + "bookmarkStart",
                        new XAttribute(Word + "id", "7"), new XAttribute(Word + "name", "ColumnTarget")));
                    target.Add(new XElement(Word + "bookmarkEnd", new XAttribute(Word + "id", "7")));
                    foreach (var p in new[] { paragraphs[0], target })
                    {
                        var runs = p.Elements(Word + "r").ToArray(); foreach (var item in runs) item.Remove();
                        var link = new XElement(Word + "hyperlink", new XAttribute(Word + "anchor", "ColumnTarget"), runs);
                        if (p.Element(Word + "bookmarkEnd") is { } end) end.AddBeforeSelf(link);
                        else p.Add(link);
                    }
                });
                var pages = Renderer(review, new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(changed, null, default)
                    .ToBlockingEnumerable().ToArray();
                TestAssert.Equal(original.Length, pages.Length);
                TestAssert.True(Characters(original).SequenceEqual(Characters(pages)),
                    "Source-preserving run splits and navigation keep every glyph and its placement.");
                if (!navigation) continue;
                var target = FrameLines(Layout(changed)).First(line => line.Line.SourceBlockIndex == 4 && line.Line.SourceLineIndex == 0);
                var destinations = pages.SelectMany(page => page.Annotations).Where(a => a.Destination is not null).ToArray();
                TestAssert.True(destinations.Length > 1, "Both the source jump and wrapped target link remain clickable.");
                foreach (var annotation in destinations)
                {
                    TestAssert.Equal(target.Page, annotation.Destination!.Value.PageIndex);
                    TestAssert.True(Math.Abs(annotation.Destination.Value.Left!.Value - target.FrameX + 3d) < .00001d,
                        "The bookmark viewport uses the admitted column frame.");
                }
            }
        }
    }

    private static IEnumerable<(DocxTextLineLayout Line, int Page, int Column, double FrameX)> FrameLines(DocxLayout layout)
    {
        for (int pi = 0; pi < layout.Pages.Count; pi++)
        {
            var page = layout.Pages[pi]; int column = 0; double previous = double.PositiveInfinity;
            foreach (var line in page.Items.OfType<DocxTextLineLayout>())
            {
                if (line.BaselineY > previous + .001d) column++;
                TestAssert.True(column < page.ColumnFrames.Count, "Only the authored physical columns are admitted.");
                yield return (line, pi, column, page.ColumnFrames[column].X);
                previous = line.BaselineY;
            }
        }
    }

    private static (char Character, int Page, double X, double Y, double Size)[] Characters(PdfPage[] pages) =>
        pages.SelectMany((page, pi) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
            .SelectMany(show => Compact(show.Text).Select(c => (c, pi, show.X, show.Y, show.Size)))).ToArray();
    private static string Compact(string text) => new(text.Normalize(NormalizationForm.FormKC).Where(c => !char.IsWhiteSpace(c)).ToArray());
    private static DocxLayout Layout(DocxDocument document)
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        var measurer = DocxFontPlanTextMeasurer.CreateAsync(DocxFontPlan.Create(document, resolver, default),
            resolver.Resolve(new FontRequest("Calibri", false, false)), default, resolver).GetAwaiter().GetResult();
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, default);
    }
    private static DocxRenderer Renderer(bool review, IFontResolver? resolver = null) => new(resolver,
        review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final,
        review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static XElement[] Paragraphs(XDocument xml) => xml.Root!.Element(Word + "body")!.Elements(Word + "p").ToArray();
    private static string Cases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "ColumnAdmission"));
    private static JsonDocument Gold() => JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
    private static DocxDocument Read(string name, bool review = false, Action<XDocument>? modify = null)
    {
        string input = Path.Combine(Cases, name);
        using var gold = Gold();
        string hash = gold.RootElement.GetProperty("Rows").EnumerateArray().First(r => r.GetProperty("Input").GetString() == name)
            .GetProperty("InputSha256").GetString()!;
        TestAssert.Equal(hash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        if (modify is null) return DocxTests.ReadDocx(input, review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final);
        var parts = new Dictionary<string, string>();
        using (var zip = ZipFile.OpenRead(input))
            foreach (var entry in zip.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        var document = XDocument.Parse(parts["word/document.xml"]); modify(document); parts["word/document.xml"] = document.ToString();
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final); }
        finally { File.Delete(temporary); }
    }
}
