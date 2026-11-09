using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxColumnAdmissionTests
{
    public static void DocxMixedColumnFlowMatchesIndependentWordOrigins()
    {
        if (!OperatingSystem.IsWindows() || !File.Exists(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "calibri.ttf")))
            TestAssert.Skip("Independent mixed-column origins require installed Windows fonts.");
        using var gold = MixedGold();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            bool review = row.GetProperty("Mode").GetString() != "final-preserve";
            var pages = Renderer(review).RenderPagesAsync(ReadMixed(row.GetProperty("Input").GetString()!, review), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            var shows = pages.SelectMany((page, pi) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                .Select(show => (Show: show, Page: pi))).ToArray();
            // Normal-font fractions use ordinary glyph shows. Their paint order
            // follows their line, independently of the source prose order.
            foreach (string kind in new[] { "Prose", "Math" })
            {
                double size = row.GetProperty(kind + "Size").GetDouble();
                var chars = shows.Where(s => Math.Abs(s.Show.Size - size) < .1d)
                    .SelectMany(s => Compact(s.Show.Text).Select((c, i) =>
                        (Character: c, s.Page, s.Show.X, s.Show.Y, s.Show.Size, Start: i == 0))).ToArray();
                TestAssert.Equal(row.GetProperty(kind + "Text").GetString(), new string(chars.Select(c => c.Character).ToArray()));
                foreach (var anchor in row.GetProperty(kind + "Anchors").EnumerateArray())
                {
                    var actual = chars[anchor.GetProperty("Offset").GetInt32()];
                    TestAssert.True(actual.Start && actual.Character.ToString() == anchor.GetProperty("Character").GetString() &&
                        actual.Page == anchor.GetProperty("Page").GetInt32() &&
                        Math.Abs(actual.X - anchor.GetProperty("X").GetDouble()) <= 1d &&
                        Math.Abs(actual.Y - anchor.GetProperty("Y").GetDouble()) <= 1d &&
                        Math.Abs(actual.Size - anchor.GetProperty("Size").GetDouble()) <= .1d,
                        $"Mixed source origins match Word: {row.GetProperty("Id").GetString()}/{kind}/{actual}.");
                }
            }
        }
    }

    public static void DocxMixedColumnFlowKeepsRunSplitsAndNavigation()
    {
        foreach (string input in new[] { "equal.docx", "wide-narrow.docx", "narrow-wide.docx" })
        foreach (bool review in new[] { false, true })
        {
            var resolver = new PortableRenderingTests.PinnedResolver();
            var before = Renderer(review, resolver).RenderPagesAsync(ReadMixed(input, review), null, default)
                .ToBlockingEnumerable().ToArray();
            var changed = ReadMixed(input, review, xml =>
            {
                var target = Paragraphs(xml)[3];
                foreach (var run in target.Elements(Word + "r").ToArray())
                {
                    string value = run.Element(Word + "t")!.Value;
                    int split = value.IndexOf(' ') + 1;
                    var first = new XElement(run); first.Element(Word + "t")!.Value = value[..split];
                    var last = new XElement(run); last.Element(Word + "t")!.Value = value[split..];
                    foreach (var text in new[] { first, last }.Select(r => r.Element(Word + "t")!))
                        text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                    run.ReplaceWith(new XElement(Word + "hyperlink", new XAttribute(Word + "anchor", "MixedColumn"), first, last));
                }
                target.Element(Word + "pPr")!.AddAfterSelf(new XElement(Word + "bookmarkStart",
                    new XAttribute(Word + "id", "7"), new XAttribute(Word + "name", "MixedColumn")));
                target.Add(new XElement(Word + "bookmarkEnd", new XAttribute(Word + "id", "7")));
            });
            var after = Renderer(review, resolver).RenderPagesAsync(changed, null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(before.Length, after.Length);
            TestAssert.True(Characters(before).Select(c => (c.Character, c.Page, c.Y, c.Size))
                .SequenceEqual(Characters(after).Select(c => (c.Character, c.Page, c.Y, c.Size))),
                "Run splitting and source navigation preserve prose/math, physical pages, baselines and sizes.");
            TestAssert.True(after.SelectMany(page => page.Annotations).Any(a => a.Destination is not null),
                "The supported mixed-content source remains linked after reflow.");
        }
    }

    private static string MixedCases => Path.GetFullPath(Path.Combine(Cases, "..", "MixedColumnFlow"));
    private static JsonDocument MixedGold() => JsonDocument.Parse(File.ReadAllText(Path.Combine(MixedCases, "office.json")));
    private static DocxDocument ReadMixed(string name, bool review, Action<XDocument>? modify = null)
    {
        string path = Path.Combine(MixedCases, name);
        using var gold = MixedGold();
        string hash = gold.RootElement.GetProperty("Rows").EnumerateArray().First(r => r.GetProperty("Input").GetString() == name)
            .GetProperty("InputSha256").GetString()!;
        TestAssert.Equal(hash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        if (modify is null) return DocxTests.ReadDocx(path, review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final);
        var parts = new Dictionary<string, string>();
        using (var zip = ZipFile.OpenRead(path))
            foreach (var entry in zip.Entries) { using var reader = new StreamReader(entry.Open()); parts[entry.FullName] = reader.ReadToEnd(); }
        var xml = XDocument.Parse(parts["word/document.xml"]); modify(xml); parts["word/document.xml"] = xml.ToString();
        string temporary = TestFixtures.WriteTempPackage(".docx", parts);
        try { return DocxTests.ReadDocx(temporary, review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final); }
        finally { File.Delete(temporary); }
    }
}
