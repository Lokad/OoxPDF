using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionAdmittedMergeContentAppearsOnceAcrossRowsAndPages()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent merged-source anchors require installed Windows fonts.");
        CheckMergedSourceConsumption(null);
    }

    internal static void CheckMergedSourceConsumption(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("merged-source-consumption-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(
                SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            int equations = row.GetProperty("SourceEquations").GetInt32();
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            List<(int Page, string Text, double X, double Y)>? originalMath = null;
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed);
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, diagnostics.Add, default).ToArray();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Merge continuations cannot remove their supported source mathematics.");
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
                TestAssert.Equal(equations, text.Split("Before inline", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(equations, text.Split("after equation.", StringSplitOptions.None).Length - 1);
                foreach (string expected in new[] { "Before public math table.", "After public math table." })
                    TestAssert.Contains(expected, text);
                if (id.Contains("page-boundary", StringComparison.Ordinal))
                    foreach (string marker in new[] { "Side row 1.", "Side row 2." })
                        TestAssert.Equal(1, text.Split(marker, StringSplitOptions.None).Length - 1);
                if (id.Contains("two-owners", StringComparison.Ordinal)) TestAssert.Contains("Tail after first owner.", text);
                var math = new List<(int Page, string Text, double X, double Y)>();
                for (int page = 0; page < pages.Length; page++)
                {
                    var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[page], usePageCoordinates: true)
                        .Where(show => pages[page].Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                            ReferenceEquals(font.Font, show.Font))).ToArray();
                    math.AddRange(shows.Select(show => (page + 1, show.Text, show.X, show.Y)));
                }
                TestAssert.Equal(4 * equations, math.Count);
                if (resolver is null)
                {
                    TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
                    foreach (var pair in math.Zip(row.GetProperty("MathAnchors").EnumerateArray()))
                        TestAssert.True(pair.First.Page == pair.Second.GetProperty("Page").GetInt32() &&
                            pair.First.Text == pair.Second.GetProperty("Text").GetString() &&
                            Math.Abs(pair.First.Y - pair.Second.GetProperty("Y").GetDouble()) < 1d,
                            "Each source fraction is painted once at its independent Word page/baseline. Absolute cell-left/row-flow residuals remain explicit.");
                }
                if (originalMath is not null)
                    foreach (var pair in originalMath.Zip(math))
                        TestAssert.True(pair.First.Page == pair.Second.Page && pair.First.Text == pair.Second.Text &&
                            Math.Abs(pair.First.X - pair.Second.X) < .002d && Math.Abs(pair.First.Y - pair.Second.Y) < .002d,
                            "Bookmarks and unchanged-format run splitting cannot reset merged-source ownership.");
                originalMath = math;
            }
        }
    }

    public static void DocxInlineCompositionMergeConsumptionRequiresCompleteAdmittedFrames()
    {
        foreach (string variant in new[] { "complete", "unmeasured", "center", "fixed", "dynamic", "oversized" })
        {
            var source = ReadMutation("CompositionProbes/merged-span-three-rows", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, _) =>
                {
                    var table = xml.Root!.Element(w + "body")!.Element(w + "tbl")!;
                    var first = table.Element(w + "tr")!;
                    if (variant == "center")
                        first.Element(w + "tc")!.Element(w + "tcPr")!.Element(w + "vAlign")!.SetAttributeValue(w + "val", "center");
                    if (variant == "fixed")
                        foreach (var row in table.Elements(w + "tr"))
                            row.AddFirst(new XElement(w + "trPr", new XElement(w + "trHeight",
                                new XAttribute(w + "val", "100"), new XAttribute(w + "hRule", "exact"))));
                    if (variant == "dynamic")
                        first.Element(w + "tc")!.Element(w + "p")!.Add(new XElement(w + "fldSimple", new XAttribute(w + "instr", " PAGE "),
                            new XElement(w + "r", new XElement(w + "t", "1"))));
                    if (variant == "oversized")
                    {
                        var section = xml.Root!.Element(w + "body")!.Element(w + "sectPr")!;
                        section.Element(w + "pgSz")!.SetAttributeValue(w + "h", "3000");
                        section.Element(w + "pgMar")!.SetAttributeValue(w + "top", "200");
                        section.Element(w + "pgMar")!.SetAttributeValue(w + "bottom", "200");
                    }
                });
            var equation = source.Tables.Single().Rows[0].Cells[0].Paragraphs[0].Equations.Single();
            var prepared = variant == "unmeasured" ? null :
                new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
                { [equation] = variant == "oversized"
                    ? new(equation, new(18d, 180d, 20d, []), 180d, 200d)
                    : new(equation, new(18d, 12d, 4d, []), 16d, 20d) };
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: prepared);
            var rows = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
            TestAssert.True(rows[0].Cells[0].TextLines.Count != 0, "The restart retains its original source content.");
            foreach (var cell in rows.Skip(1).Select(row => row.Cells[0]))
            {
                TestAssert.True(ReferenceEquals(source.Tables.Single().Rows[0].Cells[0], cell.VisualCell),
                    "Consumption retains visual ownership for shading, borders and source inspection.");
                TestAssert.True((cell.TextLines.Count == 0) == (variant == "complete"),
                    "Only a complete, static, top-aligned source frame inside its admitted row/page is consumed. Unmeasured/fixed/shifted/dynamic/oversized sources retain fallback.");
            }
        }
    }
}
