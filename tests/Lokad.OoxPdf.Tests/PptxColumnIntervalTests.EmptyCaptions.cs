using System.Xml.Linq;
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
            PdfPage Render(string mode) => RenderAdmission(row, caseDirectory: OuterCases,
                horizontalCategories: true, editChart: xml =>
                {
                    XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                    XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                    foreach (var value in xml.Descendants(c + "cat").Descendants(c + "v"))
                        value.Value = mode == "empty" ? string.Empty : "I";
                    var axis = xml.Descendants(c + "catAx").Single();
                    axis.Element(c + "tickLblPos")!.SetAttributeValue("val", mode == "hidden" ? "none" : high ? "high" : "nextTo");
                    axis.Element(c + "lblOffset")!.SetAttributeValue("val", offset);
                    SetSize(axis, size);
                    SetSize(xml.Descendants(c + "valAx").Single(), numericSize);

                    void SetSize(XElement axis, double points)
                    {
                        var tx = axis.Element(c + "txPr");
                        if (tx is null) { tx = new XElement(c + "txPr", new XElement(a + "bodyPr", new XAttribute("rot", 0))); axis.Add(tx); }
                        tx.Elements(a + "p").Remove();
                        tx.Add(new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "defRPr",
                            new XAttribute("sz", (int)(points * 100d)), new XElement(a + "latin", new XAttribute("typeface", "Calibri"))))));
                    }
                });
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
            PdfPage Render(bool hidden) => RenderAdmission(row, caseDirectory: OuterCases,
                horizontalCategories: true, editChart: xml =>
                {
                    XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                    XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                    foreach (var value in xml.Descendants(c + "cat").Descendants(c + "v")) value.Value = string.Empty;
                    var axis = xml.Descendants(c + "catAx").Single();
                    axis.Element(c + "txPr")!.Element(a + "bodyPr")!.SetAttributeValue("rot", -45 * 60000);
                    axis.Element(c + "tickLblPos")!.SetAttributeValue("val", hidden ? "none" : high ? "high" : "nextTo");
                });
            var empty = AdmissionBounds(Render(hidden: false)); var hidden = AdmissionBounds(Render(hidden: true));
            TestAssert.True(new[] { empty.Left - hidden.Left, empty.Top - hidden.Top,
                empty.Right - hidden.Right, empty.Bottom - hidden.Bottom }.All(delta => Math.Abs(delta) < .002d),
                "Independent Office controls reserve the same minimum for rotated empty captions and hidden labels.");
        }
    }
}
