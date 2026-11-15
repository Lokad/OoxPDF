using System.Xml.Linq;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingTableFractionsKeepFollowingNestedContentBelowTheirInk()
    {
        foreach (string fixture in new[] { "maintenance-matrix", "capacity-schedule" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxInlineCompositionTests.ReadMutation("MixedTables/" + fixture, mode, false, false,
                (xml, w, m) =>
                {
                    var math = xml.Descendants(m + "oMath").Single(equation => equation.Value == "a+b");
                    XElement Run(string text) => new(m + "r", new XElement(m + "rPr", new XElement(m + "nor")), new XElement(m + "t", text));
                    math.ReplaceNodes(new XElement(m + "f", new XElement(m + "num", Run("a+b")), new XElement(m + "den", Run("c+d"))));
                });
            var pages = RenderAlignmentProbe(document, mode);
            TestAssert.Equal(1, pages.Length);
            var equationInk = ReadPaintedSourceLabel(pages[0], "a+bc+d", mathOnly: true);
            var nestedInk = ReadPaintedSourceLabel(pages[0], "Nested record");
            var followingInk = ReadPaintedSourceLabel(pages[0], "After nested record");
            double clearance = equationInk.Min(glyph => glyph.Ink.Y) - nestedInk.Max(glyph => glyph.Ink.Y + glyph.Ink.Height);
            TestAssert.True(clearance > 0d,
                $"The following nested block starts below the complete displayed fraction: {fixture}/{mode}, gap={clearance}.");
            TestAssert.True(nestedInk.Min(glyph => glyph.Ink.Y) > followingInk.Max(glyph => glyph.Ink.Y + glyph.Ink.Height),
                "Text after the nested table retains its source block order.");
        }
    }
}
