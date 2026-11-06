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

    public static void PptxPositionedColumnLabelsMoveMeasuredTextFramesAndKeysTogether()
    {
        foreach (bool key in new[] { false, true })
        foreach (string wrap in new[] { "square", "none" })
        {
            PdfPage before = Render(PositionedColumnFixture("WWWL\nLWWW", wrap, key, 0d, 0d, "factor"));
            PdfPage after = Render(PositionedColumnFixture("WWWL\nLWWW", wrap, key, .02d, .02d, "factor"));
            var a = Shows(before).ToArray();
            var b = Shows(after).ToArray();
            TestAssert.Equal(a.Length, b.Length);
            TestAssert.Equal("WWWLLWWW", string.Concat(b.Select(show => show.Text)));
            for (int index = 0; index < a.Length; index++)
            {
                TestAssert.Equal(a[index].Text, b[index].Text);
                CheckTranslation(a[index].X, a[index].Y, b[index].X, b[index].Y, 1.8d, -8.64d);
            }
            var frameA = PaintedLabelFrame(before);
            var frameB = PaintedLabelFrame(after);
            CheckTranslation(frameA.X, frameA.Y, frameB.X, frameB.Y, 1.8d, -8.64d);
            if (key)
            {
                var swatchA = LabelSwatch(before);
                var swatchB = LabelSwatch(after);
                CheckTranslation(swatchA.X, swatchA.Y, swatchB.X, swatchB.Y, 1.8d, -8.64d);
            }
            PdfContentValidator.ValidatePage(after, 0, default);
        }
    }

    public static void PptxPositionedColumnLabelsUseFullChartEdgesAndFrameAdmission()
    {
        foreach (bool key in new[] { false, true })
        {
            PdfPage first = Render(PositionedColumnFixture("LLLL", "none", key, .1d, .2d, "edge"));
            PdfPage second = Render(PositionedColumnFixture("LLLL", "none", key, .3d, .3d, "edge"));
            var a = Shows(first).Single();
            var b = Shows(second).Single();
            TestAssert.Equal("LLLL", b.Text);
            CheckTranslation(a.X, a.Y, b.X, b.Y, 18d, -43.2d);
            var frame = PaintedLabelFrame(first);
            TestAssert.True(Math.Abs(frame.X - 81d) < .004 &&
                Math.Abs(frame.Y + frame.Height - (504d - .2d * 432d)) < .004,
                "Edge coordinates locate the complete occupied frame in the full90x432pt chart.");
            PdfPage clipped = Render(PositionedColumnFixture("LLLL", "none", key, 1d, 0d, "edge"));
            frame = PaintedLabelFrame(clipped);
            TestAssert.True(Math.Abs(frame.X + frame.Width - 162d) < .004 &&
                Math.Abs(frame.Y + frame.Height - 504d) < .004,
                "The admitted occupied frame, including key space, fits the physical chart edges.");
            PdfContentValidator.ValidatePage(clipped, 0, default);
        }
    }

    public static void PptxSizedColumnLabelsKeepTextCenterAndKeyTopReferences()
    {
        PdfPage first = Render(SizedColumnFixture("WWWW LLLL", "square", true, .8d, .3d));
        PdfPage second = Render(SizedColumnFixture("WWWW LLLL", "square", true, .8d, .4d));
        var a = Shows(first).ToArray();
        var b = Shows(second).ToArray();
        TestAssert.Equal(a.Length, b.Length);
        TestAssert.True(a.Length > 1 && string.Concat(a.Select(x => x.Text)).Replace(" ", "", StringComparison.Ordinal) == "WWWWLLLL",
            "The sized body retains every source word across measured lines.");
        for (int i = 0; i < a.Length; i++)
        {
            TestAssert.Equal(a[i].Text, b[i].Text);
            CheckTranslation(a[i].X, a[i].Y, b[i].X, b[i].Y, 0d, -21.6d);
        }
        var frame = PaintedLabelFrame(first);
        TestAssert.True(Math.Abs(frame.Width - 72d) < .004 && Math.Abs(frame.Height - 129.6d) < .004 &&
            Math.Abs(frame.X - 81d) < .004 && Math.Abs(frame.Y + frame.Height - 417.6d) < .004,
            "Authored dimensions and edge positions belong to the complete chart frame.");
        var keyA = LabelSwatch(first);
        var keyB = LabelSwatch(second);
        CheckTranslation(keyA.X, keyA.Y, keyB.X, keyB.Y, 0d, 0d);
        PdfContentValidator.ValidatePage(second, 0, default);
    }

    public static void PptxSizedColumnLabelsElideWordsAndKeepTextElementsWhole()
    {
        foreach (var (source, width, expected) in new[]
        {
            ("Alpha Beta Gamma", 1d, "Alpha\u2026"),
            ("WWWWWWWW", .6d, "W\u2026"),
            (string.Concat(Enumerable.Repeat("\U0001f600\u0301", 8)), .6d, "\U0001f600\u0301\u2026")
        })
        {
            PdfPage page = Render(SizedColumnFixture(source, "square", true, width, .08d));
            var show = Shows(page).Single();
            TestAssert.Equal(expected, show.Text);
            PdfContentValidator.ValidatePage(page, 0, default);
        }
        PdfPage tiny = Render(SizedColumnFixture("WWWWWWWW", "square", true, .005d, .02d));
        TestAssert.Equal("W", Shows(tiny).Single().Text);
    }

    public static void PptxSizedColumnNoWrapKeepsWholeMandatoryParagraphs()
    {
        PdfPage page = Render(SizedColumnFixture("WWWL\nLWWW", "none", true, .4d, .2d));
        TestAssert.Equal("WWWL\nLWWW", string.Join("\n", Shows(page).Select(x => x.Text)));
        PdfContentValidator.ValidatePage(page, 0, default);
    }

    public static void PptxSizedColumnLabelsDistinguishMissingZeroAndRejectedExtents()
    {
        PdfPage keyOnly = Render(SizedColumnFixture("LLLL", "square", true, null, .1d));
        TestAssert.Equal(0, Shows(keyOnly).Count());
        TestAssert.True(Math.Abs(PaintedLabelFrame(keyOnly).Width - 22.7d) < .004,
            "An absent width retains only the source key reference.");
        LabelSwatch(keyOnly);
        foreach (var (width, height, key) in new (double?, double?, bool)[]
        {
            (0d, .1d, true), (.4d, 0d, true), (.4d, null, true), (null, .1d, false)
        })
        {
            PdfPage page = Render(SizedColumnFixture("LLLL", "square", key, width, height));
            TestAssert.Equal(0, Shows(page).Count());
            TestAssert.True(!page.Content.Contains("1 0 1 rg", StringComparison.Ordinal) &&
                !System.Text.RegularExpressions.Regex.IsMatch(page.Content, @"[-0-9.]+ [-0-9.]+ 9\.9 9\.9 re f"),
                "An empty declared area does not paint a label shape or key.");
        }
        PdfPage automatic = Render(PositionedColumnFixture("LLLL", "square", true, 0d, 0d, "factor"));
        foreach (var (width, height) in new[] { (-.1d, .1d), (.4d, -.1d), (1.2d, .1d) })
        {
            PdfPage page = Render(SizedColumnFixture("LLLL", "square", true, width, height));
            TestAssert.Equal(automatic.Content, page.Content);
            PdfContentValidator.ValidatePage(page, 0, default);
        }
    }

    private static byte[] SizedColumnFixture(string text, string wrap, bool key, double? width, double? height)
    {
        using var stream = new MemoryStream();
        stream.Write(PositionedColumnFixture(text, wrap, key, .1d, .2d, "edge"));
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
            XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XElement layout = xml.Descendants(c + "manualLayout").Single();
            if (width is { } w) layout.Add(new XElement(c + "wMode", new XAttribute("val", "factor")),
                new XElement(c + "w", new XAttribute("val", w)));
            if (height is { } h) layout.Add(new XElement(c + "hMode", new XAttribute("val", "factor")),
                new XElement(c + "h", new XAttribute("val", h)));
            entry.Delete();
            using var output = archive.CreateEntry("ppt/charts/chart1.xml").Open();
            xml.Save(output);
        }
        return stream.ToArray();
    }

    private static void CheckTranslation(double ax, double ay, double bx, double by, double dx, double dy) =>
        TestAssert.True(Math.Abs(bx - ax - dx) < .004 && Math.Abs(by - ay - dy) < .004,
            "Text, frame and swatch must share chart-wide position offsets.");

    private static (double X, double Y, double Width, double Height) PaintedLabelFrame(PdfPage page) =>
        Rectangle(page, @"1 0 1 rg\s+([-0-9.]+) ([-0-9.]+) ([-0-9.]+) ([-0-9.]+) re f");

    private static (double X, double Y) LabelSwatch(PdfPage page)
    {
        var rectangle = Rectangle(page, @"([-0-9.]+) ([-0-9.]+) (9\.9) (9\.9) re f");
        return (rectangle.X, rectangle.Y);
    }

    private static (double X, double Y, double Width, double Height) Rectangle(PdfPage page, string pattern)
    {
        var match = System.Text.RegularExpressions.Regex.Match(page.Content, pattern);
        TestAssert.True(match.Success, "Expected a painted label frame or legend-key swatch.");
        double Read(int group) => double.Parse(match.Groups[group].Value, System.Globalization.CultureInfo.InvariantCulture);
        return (Read(1), Read(2), Read(3), Read(4));
    }

    private static byte[] PositionedColumnFixture(string text, string wrap, bool key, double x, double y, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(OverfullColumnFixture(text, wrap, key));
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
            XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            xml.Descendants(c + "dLbls").First().Add(new XElement(c + "dLbl",
                new XElement(c + "idx", new XAttribute("val", 0)),
                new XElement(c + "layout", new XElement(c + "manualLayout",
                    new XElement(c + "xMode", new XAttribute("val", mode)),
                    new XElement(c + "yMode", new XAttribute("val", mode)),
                    new XElement(c + "x", new XAttribute("val", x)),
                    new XElement(c + "y", new XAttribute("val", y)))),
                new XElement(c + "spPr", new XElement(a + "solidFill",
                    new XElement(a + "srgbClr", new XAttribute("val", "FF00FF")))),
                new XElement(c + "showVal", new XAttribute("val", 0)),
                new XElement(c + "showCatName", new XAttribute("val", 0)),
                new XElement(c + "showSerName", new XAttribute("val", 1)),
                new XElement(c + "showLegendKey", new XAttribute("val", key ? 1 : 0))));
            entry.Delete();
            using var output = archive.CreateEntry("ppt/charts/chart1.xml").Open();
            xml.Save(output);
        }
        return stream.ToArray();
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
