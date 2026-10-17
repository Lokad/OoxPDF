using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionScaledCellStrutsKeepPhysicalBoundsAcrossMetricsAndDepths()
    {
        foreach (double size in new[] { 9d, 11d, 12d, 17d })
        foreach (double scale in new[] { .5d, .75d, .9d, 1d })
        foreach (bool nested in new[] { false, true })
        {
            var paragraph = DocxTests.CreateDocxLayoutParagraph("Public cell line", size, size) with { LineSpacingPoints = null };
            var cell = new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
            var table = new DocxTable(null, [240d], [new DocxTableRow([cell], null)]);
            if (nested) table = new DocxTable(null, [240d], [new DocxTableRow([cell with
                { Paragraphs = [], BodyElements = [new DocxTableElement(table)] }], null)]);
            var source = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
            var natural = new DocxTests.FamilyWidthTextMeasurer();
            var measurer = new DocxTests.ScaledLayoutTextMeasurer(natural, scale, scale);
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, wordCompatiblePrintScale: scale)
                .Create(source, measurer, default, unscaledTextMeasurer: natural);
            var lines = layout.Pages.SelectMany(page => NestedMathRows(page.Items.OfType<DocxTableRowLayout>()))
                .SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines).ToArray();
            TestAssert.True(lines.Length != 0, "Every ordinary source line remains measured.");
            foreach (var line in lines)
            {
                TestAssert.True(line.BodyLineBoxHeightPoints is null && line.OccupiedLineBox is { Height: > 0d },
                    "Physical measurement does not grant the narrower plain-body painting policy.");
                TestAssert.True(DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double bottom, out double top) &&
                    bottom <= line.BaselineY && top >= line.BaselineY,
                    $"Scaled struts contain their actual baseline: size={size}/scale={scale}/nested={nested}.");
            }
        }
    }

    public static void DocxInlineCompositionScaledCellStrutsMatchIndependentWordNoteRectangles()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent note rectangles use installed Windows fonts.");
        CheckScaledCellNoteStruts(null, independentGeometry: true);
    }

    internal static void CheckScaledCellNoteStruts(IFontResolver? resolver, bool independentGeometry = false)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("scaled-cell-notes-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderBlankPages(DocxTests.ReadDocx(NestedMathCasePath(id + ".docx"), mode), null, default).Single();
            string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
            foreach (string prose in new[] { "Public cell before note", "and after note.", "Public review anchor", "Public note target." })
                TestAssert.Contains(prose, text);
            var link = page.Annotations.Single();
            TestAssert.True(link.Destination is { PageIndex: 0, Left: not null, Top: not null } && link.Width > 0d && link.Height > 0d,
                "The custom note retains its owning body-cell link and placed destination.");
            if (!independentGeometry) continue;
            double[] actual = [link.X, link.Y, link.X + link.Width, link.Y + link.Height];
            double[] expected = row.GetProperty("Rect").EnumerateArray().Select(value => value.GetDouble()).ToArray();
            for (int i = 0; i < actual.Length; i++)
                TestAssert.True(Math.Abs(actual[i] - expected[i]) < .5d,
                    $"Physical cell note bounds match independent Word: {id}/{mode}/edge={i}.");
        }
    }
}
