using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMinimumRowsKeepIndependentWordContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent row geometry uses installed Windows fonts.");
        CheckEquationRevisions(null, false, "minimum-row-placement-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionMinimumRowsKeepIndependentWordLabelFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent row labels use installed Windows fonts.");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "minimum-row-placement-labels-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            if (!row.GetProperty("CheckNativeLabelGeometry").GetBoolean()) continue;
            string id = row.GetProperty("Id").GetString()!;
            var document = ReadMutation("CompositionProbes/" + id, OoxPdfDocxMarkupMode.AllMarkup, false, false);
            var page = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).Single();
            string label = row.GetProperty("Kind").GetString() == "ins" ? "Inserted Cells" : "Deleted Cells";
            var title = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Single(show => show.Text.Contains(label, StringComparison.Ordinal));
            TestAssert.True(Math.Abs(title.Y - row.GetProperty("LabelBaselineY").GetDouble()) < 1d,
                $"The label consumes the independently measured printed row frame: {id}/{title.Y}.");
        }
    }

    public static void DocxInlineCompositionMinimumRowsScaleAuthoredExtentOnce()
    {
        var source = ReadMutation("CompositionProbes/cell-structure-placement-tall-ins", OoxPdfDocxMarkupMode.AllMarkup, false, false);
        var sourceTable = source.Tables.Single();
        foreach (string? rule in new string?[] { null, "atLeast", "exact" })
        foreach (double scale in new[] { 1d, .75d })
        {
            var table = sourceTable with { Rows = [sourceTable.Rows.Single() with { HeightRuleValue = rule }] };
            var document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
            var row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, scale)
                .Create(document, new TrackingMeasurer(), default).Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.True(Math.Abs(row.Height - 144d * scale) < .001d && row.DeclaredHeightPoints == 144d,
                "Authored minima and exact extents enter printed space once, while source metadata remains unchanged.");
        }
    }
}
