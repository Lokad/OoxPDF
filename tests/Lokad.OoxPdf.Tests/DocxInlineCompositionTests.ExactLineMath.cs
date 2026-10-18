using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionExactLineMathKeepsWordSlotsAcrossContainers()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent exact-slot origins require installed Windows fonts.");
        CheckDisplayIntegrals(null, "exact-line-math-office.json");
    }

    public static void DocxInlineCompositionExactLineReportKeepsIndependentScaledContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent composed exact-slot origins require installed Windows fonts.");
        CheckExactLineReport(null);
    }

    internal static void CheckExactLineReport(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("exact-line-math-report-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            string? originalOrigins = null;
            foreach (bool transformed in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed);
                var source = document.Paragraphs.First(p => p.Equations.Count != 0);
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).ToArray();
                TestAssert.Equal(2, pages.Length);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "An exact slot must not omit its supported formula in the composed document.");
                var math = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Where(show => page.Fonts.Any(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                        ReferenceEquals(resource.Font, show.Font)))).ToArray();
                TestAssert.Equal(row.GetProperty("MathGlyphCount").GetInt32(), math.Length);
                string origins = string.Join("|", math.Select(show => show.X.ToString("R", CultureInfo.InvariantCulture) + "," +
                    show.Y.ToString("R", CultureInfo.InvariantCulture)));
                if (originalOrigins is not null) TestAssert.Equal(originalOrigins, origins);
                originalOrigins = origins;
                string prose = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
                foreach (string text in new[] { "Quarterly operations report", "Inventory availability", "All quantities in this synthetic example are illustrative." })
                    TestAssert.Contains(text, prose);
                TestAssert.Equal(48d, source.LineSpacingPoints);
                if (resolver is not null) continue;
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0], usePageCoordinates: true).ToArray();
                var summary = shows.GroupBy(show => show.Y).Select(line => line.OrderBy(show => show.X).ToArray())
                    .First(line => string.Concat(line.Select(show => show.Text)).StartsWith("Summary:", StringComparison.Ordinal))[0];
                TestAssert.True(Math.Abs(summary.Y - row.GetProperty("SummaryY").GetDouble()) < 1d,
                    "Exact source baseline retains the independent one-point Word gate, including review scaling.");
                var target = shows.Where(show => pages[0].Fonts.Any(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                    ReferenceEquals(resource.Font, show.Font)) && Math.Abs(show.Y - summary.Y) < summary.Size).ToArray();
                TestAssert.Equal(row.GetProperty("TargetMathText").GetString(),
                    string.Concat(target.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(row.GetProperty("MathAnchors").GetArrayLength(), target.Length);
                // The glyph reader maps origins but intentionally retains raw Tf
                // sizes. Compare Word's physical size after the page print map.
                double printScale = final ? 1d : DocxRenderer.ResolveWordCompatiblePrintScale(document,
                    DocxMarkupContext.FromMode(mode, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
                foreach (var pair in target.Zip(row.GetProperty("MathAnchors").EnumerateArray()))
                    TestAssert.True(Math.Abs(pair.First.X - pair.Second.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(pair.First.Y - pair.Second.GetProperty("Y").GetDouble()) < 1d &&
                        Math.Abs(pair.First.Size * printScale - pair.Second.GetProperty("Size").GetDouble()) < .1d,
                        "Restored formula glyphs keep independent Word geometry and authored size in both presets.");
            }
        }
    }

    public static void DocxInlineCompositionExactLineSlotSeparatesInkFromAdvanceAndKeepsFallbacks()
    {
        var document = ReadMutation("CompositionProbes/math-minimum-line-72", OoxPdfDocxMarkupMode.Final, false, false);
        var source = document.Paragraphs.First(p => p.Equations.Count != 0);
        var equation = source.Equations.Single();
        var prepared = new DocxPreparedMathParagraph(equation, new(18d, 50d, 40d, []), 52d, 96d) { UsesDisplayStyle = false };
        var equations = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance) { [equation] = prepared };
        foreach (double height in new[] { 12d, 48d, 72d })
        foreach (double scale in new[] { 1d, .75d })
        {
            var paragraph = source with { LineSpacingPoints = height,
                Spacing = source.Spacing with { LineRuleValue = "exact", LineValue = (height * 20d).ToString(System.Globalization.CultureInfo.InvariantCulture) } };
            TestAssert.Equal(1, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new TrackingMeasurer(),
                12d, 9d * scale, 60d * scale, scale, default, out var atoms), "A finite fixed slot retains its source equation.");
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(TestAssert.NotNull(atoms), 600d, 600d, default, out var lines),
                "Mixed content still uses the shared wrapping path.");
            var line = TestAssert.NotNull(lines).Single();
            TestAssert.True(Math.Abs(line.OccupiedHeight - height * scale) < 1e-9 &&
                Math.Abs(line.BaselineInset - height * scale * .8d) < 1e-9,
                "The authored slot owns advance and the existing exact text baseline, independent of incoming natural or ink heights.");
            TestAssert.True(line.Fragments.Any(fragment => ReferenceEquals(fragment.Math, prepared) &&
                fragment.OccupiedHeight > line.OccupiedHeight), "Tall ink remains present without enlarging the fixed slot.");
            TestAssert.Equal(height, paragraph.LineSpacingPoints);
        }
        foreach (double? invalid in new double?[] { null, 0d, -1d, double.NaN, double.PositiveInfinity })
        {
            var paragraph = source with { LineSpacingPoints = invalid, Spacing = source.Spacing with { LineRuleValue = "exact" } };
            TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
        }
    }
}
