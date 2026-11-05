using System.Xml.Linq;
using System.Text.Json;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    public static void PptxManualOuterEmptyCaptionsRemoveOnlyTheirTextLine()
    {
        using var oracle = OuterOracle();
        var row = OuterRow(oracle, "outer-width35-left12-y03");
        // Independent Office ordinary/empty/hidden triplets across four faces,
        // both sides, zero offset and enlarged numeric labels establish this
        // relation. Use the pinned face's measured line, not native font sizes
        // or the renderer's baseline coefficients, for portable geometry.
        foreach (var (size, high, numericSize, offset) in new[] {
            (12d, false, 18d, 100), (18d, false, 18d, 100), (24d, false, 18d, 100),
            (18d, true, 18d, 100), (18d, false, 36d, 100), (18d, false, 18d, 0) })
        {
            PdfPage Render(string mode) => RenderCaption(row, size, high, numericSize, offset, mode);
            var ordinary = Render("ordinary"); var empty = Render("empty"); var hidden = Render("hidden");
            var o = AdmissionBounds(ordinary); var e = AdmissionBounds(empty); var h = AdmissionBounds(hidden);
            var font = AdmissionNumericShows(ordinary)[0].Font.Font;
            var os2 = font.Os2;
            double units = (os2.SelectionFlags & 0x80) != 0
                ? os2.TypographicAscender - os2.TypographicDescender + os2.TypographicLineGap
                : os2.WindowsAscender + os2.WindowsDescender;
            double line = units * size / font.UnitsPerEm;
            double expected = high ? Math.Max(h.Top, o.Top - line) : Math.Min(h.Bottom, o.Bottom + line);
            TestAssert.True(Math.Abs((high ? e.Top : e.Bottom) - expected) < .002d,
                $"Empty captions retain spacing after removing one selected normal line (size={size}, high={high}, numeric={numericSize}, offset={offset}).");
            TestAssert.True(Math.Abs(e.Left - h.Left) < .002d && Math.Abs(e.Right - h.Right) < .002d &&
                Math.Abs((high ? e.Bottom - h.Bottom : e.Top - h.Top)) < .002d,
                "Only the category-caption edge changes; the numeric owner and opposite endcap remain fixed.");
            var numeric = AdmissionNumericShows(ordinary).Select(show => show.Text).ToArray();
            foreach (var page in new[] { empty, hidden })
            {
                TestAssert.True(AdmissionNumericShows(page).Select(show => show.Text).SequenceEqual(numeric),
                    "Removing category text preserves every numeric caption.");
                TestAssert.Equal(0, CategoryShows(page).Length);
            }
            TestAssert.Equal("IIIIIIII", string.Concat(CategoryShows(ordinary)));
            static string[] CategoryShows(PdfPage page) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                usePageCoordinates: false, fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal))
                .Select(show => show.Text).ToArray();
        }
    }

    public static void PptxManualOuterRotatedEmptyCaptionsKeepTheirMinimumBand()
    {
        using var oracle = OuterOracle();
        var row = OuterRow(oracle, "outer-width35-left12-y03");
        foreach (bool high in new[] { false, true })
        {
            PdfPage Render(bool hidden) => RenderCaption(row, 18d, high, 18d, 100, hidden ? "hidden" : "empty", rotation: -45d);
            var empty = AdmissionBounds(Render(hidden: false)); var hidden = AdmissionBounds(Render(hidden: true));
            TestAssert.True(new[] { empty.Left - hidden.Left, empty.Top - hidden.Top,
                empty.Right - hidden.Right, empty.Bottom - hidden.Bottom }.All(delta => Math.Abs(delta) < .002d),
                "Independent Office controls reserve the same minimum for rotated empty captions and hidden labels.");
        }
    }

    public static void PptxManualOuterWhitespaceCaptionsReserveMeasuredLines()
    {
        using var oracle = OuterOracle();
        var row = OuterRow(oracle, "outer-width35-left12-y03");
        foreach (var (size, high, rotation) in new[] { (18d, false, 0d), (24d, false, 0d), (18d, true, 0d), (18d, false, -45d) })
        {
            PdfPage Render(string mode) => RenderCaption(row, size, high, 18d, 100, mode, rotation);
            var spaces = Render("spaces"); var ordinary = Render("ordinary"); var empty = Render("empty");
            var s = AdmissionBounds(spaces); var o = AdmissionBounds(ordinary); var e = AdmissionBounds(empty);
            if (rotation == 0d)
                TestAssert.True(new[] { s.Left - o.Left, s.Top - o.Top, s.Right - o.Right, s.Bottom - o.Bottom }
                    .All(delta => Math.Abs(delta) < .002d),
                    "Native flat space-only labels retain the same normal-line frame as short visible labels.");
            else
                TestAssert.True(e.Bottom - s.Bottom > 10d,
                    "Native rotated space-only labels project a measured text line; rotated empty sources do not.");
            TestAssert.True(AdmissionNumericShows(spaces).Select(show => show.Text)
                .SequenceEqual(AdmissionNumericShows(ordinary).Select(show => show.Text)),
                "Space-only category measurement preserves the complete numeric interval.");
            TestAssert.Equal(0, DocxCommentsTests.ReadEmbeddedGlyphTextShows(spaces, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).Count());
        }
    }

    private static PdfPage RenderCaption(JsonElement row, double size, bool high, double numericSize,
        int offset, string mode, double rotation = 0d, string target = "outer", int? skip = 1,
        string? text = null) => RenderAdmission(row, caseDirectory: OuterCases,
        horizontalCategories: true, editChart: xml =>
        {
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            if (target == "automatic") xml.Descendants(c + "plotArea").Single().Element(c + "layout")!.Remove();
            else xml.Descendants(c + "manualLayout").Single().Element(c + "layoutTarget")!.SetAttributeValue("val", target);
            foreach (var value in xml.Descendants(c + "cat").Descendants(c + "v"))
                value.Value = text ?? (mode == "empty" ? string.Empty : mode == "spaces" ? " " : "I");
            var axis = xml.Descendants(c + "catAx").Single();
            axis.Element(c + "tickLblSkip")?.Remove();
            if (skip is { } authoredSkip) axis.Add(new XElement(c + "tickLblSkip", new XAttribute("val", authoredSkip)));
            axis.Element(c + "tickLblPos")!.SetAttributeValue("val", mode == "hidden" ? "none" : high ? "high" : "nextTo");
            axis.Element(c + "lblOffset")!.SetAttributeValue("val", offset);
            SetSize(axis, size);
            axis.Element(c + "txPr")!.Element(a + "bodyPr")!.SetAttributeValue("rot", (int)(rotation * 60000d));
            SetSize(xml.Descendants(c + "valAx").Single(), numericSize);

            void SetSize(XElement owner, double points)
            {
                var tx = owner.Element(c + "txPr");
                if (tx is null) { tx = new XElement(c + "txPr", new XElement(a + "bodyPr")); owner.Add(tx); }
                tx.Elements(a + "p").Remove();
                tx.Add(new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "defRPr",
                    new XAttribute("sz", (int)(points * 100d)), new XElement(a + "latin", new XAttribute("typeface", "Calibri"))))));
            }
        });
}
