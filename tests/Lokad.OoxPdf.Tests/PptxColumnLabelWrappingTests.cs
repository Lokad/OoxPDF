using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxColumnLabelWrappingTests
{
    private static string Cases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases"));

    internal static void CheckAutomatic()
    {
        using var controls = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "ColumnLabelWrapping/office-labels.json")));
        foreach (var control in controls.RootElement.EnumerateArray().Where(item => item.GetProperty("Automatic").GetBoolean()))
        {
            string id = control.GetProperty("Id").GetString()!;
            PdfPage page = Render(File.ReadAllBytes(Path.Combine(Cases, "ColumnLabelWrapping", id + ".pptx")));
            string text = Normalize(string.Concat(Shows(page).Select(show => show.Text)));
            foreach (var label in control.GetProperty("Labels").EnumerateArray())
                TestAssert.True(text.Contains(Normalize(label.GetProperty("Text").GetString()!), StringComparison.Ordinal),
                    id + ": every Office source label survives automatic wrapping.");
            var shape = control.GetProperty("Shape");
            CheckFrames(page, shape.GetProperty("ShapeLeft").GetDouble(),
                page.Height - shape.GetProperty("ShapeTop").GetDouble() - shape.GetProperty("ShapeHeight").GetDouble(),
                shape.GetProperty("ShapeWidth").GetDouble(), shape.GetProperty("ShapeHeight").GetDouble(), id);
        }
    }

    internal static void CheckComposition()
    {
        foreach (string name in new[] { "keys-point-override-valid", "fixed-series-negative-keys", "authored-series-point-keys" })
        {
            byte[] input = File.ReadAllBytes(Path.Combine(Cases, "ComposedBarPalettes", name + ".pptx"));
            PdfPage page = Render(input);
            string text = string.Concat(Shows(page).Select(show => show.Text));
            TestAssert.True(text.Contains("96", StringComparison.Ordinal) && text.Contains("145", StringComparison.Ordinal),
                name + ": both final automatic value suffixes remain encoded.");
            CheckFrames(page, 72d, 72d, 720d, 432d, name, positiveOnly: true);
        }
    }

    public static void PptxColumnLabelSourceRetainsExplicitWrappingConstraints()
    {
        foreach (string wrap in new[] { "none", "square", "future" })
        {
            XElement labels = XElement.Parse($"""
                <c:dLbls xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:txPr><a:bodyPr wrap="{wrap}"/></c:txPr>
                </c:dLbls>
                """);
            TestAssert.Equal(wrap, PptxSceneBuilder.ReadChartTextBodyProperties(labels).WrapValue);
        }
        TestAssert.Equal(string.Empty, PptxSceneBuilder.ReadChartTextBodyProperties(null).WrapValue);
    }

    public static void PptxAutomaticColumnOverfullGlyphsKeepFittedReferenceWidths()
    {
        const double cap = 12d; // 90pt chart /5, less the two existing3pt insets.
        foreach (string source in new[] { "WWWL", "LWWW" })
        {
            PdfPage page = Render(OverfullColumnFixture(source));
            var shows = Shows(page).ToArray();
            TestAssert.True(shows.Length == 4 && string.Concat(shows.Select(x => x.Text)) == source,
                "Forced overfull glyphs preserve all source text and four actual wrapped lines.");
            var narrow = shows.Single(x => x.Text == "L");
            var font = narrow.Font.Font;
            double scale = narrow.Size / font.UnitsPerEm;
            double narrowAdvance = font.GetAdvanceWidth(font.MapCodePoint('L')) * scale;
            double wideAdvance = font.GetAdvanceWidth(font.MapCodePoint('W')) * scale;
            TestAssert.True(narrowAdvance < cap && wideAdvance > cap,
                "The pinned source challenges a fitting glyph and a forced overfull text element.");
            TestAssert.True(shows.Where(x => x.Text == "W").All(x =>
                Math.Abs(narrow.X - x.X - (cap - narrowAdvance) / 2d) < .004),
                "The fitted source width owns line alignment while full overfull glyphs remain painted.");
            double advance = (font.Hhea.HorizontalAscender - font.Hhea.HorizontalDescender + font.Hhea.HorizontalLineGap) * scale;
            for (int index = 1; index < shows.Length; index++)
                TestAssert.True(Math.Abs(shows[index - 1].Y - shows[index].Y - advance) < .004,
                    "Overfull fragments keep the existing four-line normal advance.");
            PdfContentValidator.ValidatePage(page, 0, default);
        }
    }

    public static void PptxAutomaticColumnNoWrapKeepsCenteredSourceParagraphs()
    {
        var reference = Shows(Render(OverfullColumnFixture("L", "none"))).Single();
        double center = reference.X + ColumnAdvance(reference.Font, reference.Size, "L") / 2d;
        foreach (string source in new[] { "WWWL", "LWWW", "W W W L", "WWWL\nLWWW" })
        {
            PdfPage page = Render(OverfullColumnFixture(source, "none"));
            var shows = Shows(page).ToArray();
            TestAssert.True(string.Join("\n", shows.Select(x => x.Text)) == source,
                "Authored no-wrap keeps complete source lines and mandatory paragraph breaks.");
            TestAssert.True(shows.All(x => Math.Abs(x.X + ColumnAdvance(x.Font, x.Size, x.Text) / 2d - center) < .004),
                "Unwrapped overfull text centers its full advance on the source reference.");
            PdfContentValidator.ValidatePage(page, 0, default);
        }
    }

    public static void PptxAutomaticColumnNoWrapKeysKeepParagraphAndSwatchReferences()
    {
        var plain = Shows(Render(OverfullColumnFixture("WWWL", "none"))).Single();
        double center = plain.X + ColumnAdvance(plain.Font, plain.Size, plain.Text) / 2d;
        string? firstSwatch = null;
        foreach (string source in new[] { "WWWL", "LLLL" })
        {
            PdfPage page = Render(OverfullColumnFixture(source, "none", key: true));
            var show = Shows(page).Single();
            TestAssert.Equal(source, show.Text);
            TestAssert.True(Math.Abs(show.X + ColumnAdvance(show.Font, show.Size, source) / 2d - center - 22.7d) < .004,
                "The existing18pt key reserve belongs to the complete unwrapped paragraph span.");
            var swatch = System.Text.RegularExpressions.Regex.Match(page.Content, @"[-0-9.]+ [-0-9.]+ 9\.9 9\.9 re f");
            TestAssert.True(swatch.Success, "Expected the source-sized key swatch.");
            firstSwatch ??= swatch.Value;
            TestAssert.Equal(firstSwatch, swatch.Value);
            PdfContentValidator.ValidatePage(page, 0, default);
        }
    }

    private static double ColumnAdvance(PdfEmbeddedFont embedded, double size, string text)
    {
        var font = embedded.Font;
        double units = 0d;
        ushort previous = 0;
        foreach (var rune in text.EnumerateRunes())
        {
            ushort glyph = font.MapCodePoint(rune.Value);
            units += font.GetAdvanceWidth(glyph) + (previous == 0 ? 0 : font.GetKerning(previous, glyph));
            previous = glyph;
        }
        return units * size / font.UnitsPerEm;
    }

    private static byte[] OverfullColumnFixture(string text, string wrap = "square", bool key = false)
    {
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(Path.Combine(Cases, "ColumnLabelWrapping/single-long-word.pptx")));
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            void Edit(string name, Action<XDocument> edit)
            {
                var entry = archive.GetEntry(name)!;
                XDocument xml;
                using (var input = entry.Open()) xml = XDocument.Load(input);
                edit(xml); entry.Delete();
                using var output = archive.CreateEntry(name).Open(); xml.Save(output);
            }
            Edit("ppt/charts/chart1.xml", xml =>
            {
                XElement series = xml.Descendants(c + "ser").First();
                series.Element(c + "tx")!.Element(c + "v")!.Value = text;
                foreach (string name in new[] { "cat", "val" })
                {
                    XElement literal = series.Element(c + name)!.Elements().Single();
                    literal.Element(c + "ptCount")!.SetAttributeValue("val", 1);
                    literal.Elements(c + "pt").Where(point => (int)point.Attribute("idx")! != 0).Remove();
                    literal.Element(c + "pt")!.Element(c + "v")!.Value = name == "cat" ? "A" : "20";
                }
                XElement labels = xml.Descendants(c + "dLbls").First();
                foreach (string name in new[] { "showVal", "showCatName", "showLegendKey" })
                    labels.Element(c + name)!.SetAttributeValue("val", 0);
                if (key) labels.Element(c + "showLegendKey")!.SetAttributeValue("val", 1);
                labels.Element(c + "txPr")!.Element(a + "bodyPr")!.SetAttributeValue("wrap", wrap);
                xml.Descendants(c + "legend").Remove();
            });
            Edit("ppt/slides/slide1.xml", xml => xml.Descendants(p + "graphicFrame").First()
                .Element(p + "xfrm")!.Element(a + "ext")!.SetAttributeValue("cx", 90 * 12700));
        }
        return stream.ToArray();
    }

    internal static void CheckUnicode()
    {
        byte[] source = File.ReadAllBytes(Path.Combine(Cases, "ColumnLabelWrapping/single-long-word.pptx"));
        using var stream = new MemoryStream();
        stream.Write(source); stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
            XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            xml.Descendants(c + "barChart").First().Element(c + "ser")!.Element(c + "tx")!.Element(c + "v")!.Value =
                "AnnualRevenueForecast\U0001f600e\u0301\nNext line";
            entry.Delete();
            using var output = archive.CreateEntry("ppt/charts/chart1.xml").Open();
            xml.Save(output);
        }
        var shows = Shows(Render(stream.ToArray())).ToArray();
        string text = string.Concat(shows.Select(show => show.Text));
        TestAssert.True(Normalize(text).Contains("AnnualRevenueForecast\U0001f600e\u0301Nextline", StringComparison.Ordinal),
            "Explicit breaks and supported supplementary/combining characters survive wrapping.");
        TestAssert.Equal(8, shows.Count(show => show.Text.Contains("e\u0301", StringComparison.Ordinal)));
        TestAssert.True(shows.All(show => !show.Text.StartsWith('\u0301')),
            "A combining sequence remains on one measured line.");
    }

    private static void CheckFrames(PdfPage page, double x, double y, double width, double height, string id, bool positiveOnly = false)
    {
        foreach (var show in Shows(page))
        {
            // Count advances independently from layout. Ignoring negative kerning
            // gives a conservative upper bound; 2pt covers that difference, not
            // a visual threshold. Whole-value clipping exceeds it by many points.
            double advance = show.Text.EnumerateRunes().Sum(rune =>
                (double)show.Font.Font.GetAdvanceWidth(show.Font.Font.MapCodePoint(rune.Value))) * show.Size / show.Font.Font.UnitsPerEm;
            if (positiveOnly && show.Text.Contains('-')) continue;
            TestAssert.True(show.X >= x - 2d && show.X + advance <= x + width + 2d &&
                show.Y >= y && show.Y <= y + height,
                id + ": emitted label text must fit inside its chart frame: " + show.Text);
        }
    }

    private static string Normalize(string text) => string.Concat(text.Where(c => !char.IsWhiteSpace(c)));
    private static IEnumerable<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)> Shows(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false, fontResourceFilter: resource => resource.StartsWith("CBD", StringComparison.Ordinal));

    private static PdfPage Render(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
            new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
