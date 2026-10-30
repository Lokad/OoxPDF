using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNormalTableStyleTests
{
    public static void DocxNormalTableIdentityAndOverridesMatchIndependentWord()
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "Cases", "NormalTableStyle"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "office.json")));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string path = Path.Combine(cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var document = DocxTests.ReadDocx(path, mode);
                var cells = document.BodyElements.OfType<DocxTableElement>().Single().Table.Rows
                    .SelectMany(row => row.Cells).ToArray();
                TestAssert.Equal(9, cells.Length);
                foreach (var cell in cells)
                {
                    var expected = control.GetProperty("Cells").GetProperty(cell.Text);
                    TestAssert.Equal(expected.GetProperty("TextColor").GetString(),
                        cell.Paragraphs[0].Runs[0].ColorHex ?? "000000");
                    var fill = expected.GetProperty("Fill");
                    TestAssert.Equal(fill.ValueKind == JsonValueKind.Null ? null : fill.GetString(), cell.FillHex);
                    TestAssert.Equal(expected.GetProperty("LeftPadding").GetDouble(), cell.Margins.LeftPoints);
                    TestAssert.Equal(expected.GetProperty("RightPadding").GetDouble(), cell.Margins.RightPoints);
                    TestAssert.Equal(control.GetProperty("HasBorders").GetBoolean(), cell.Borders.Count != 0);
                }

                var layout = new DocxLayoutEngine(mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, wordCompatiblePrintScale: 1d)
                    .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default);
                var placed = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
                    .SelectMany(row => row.Cells).ToArray();
                TestAssert.Equal(9, placed.Length);
                foreach (var cell in placed)
                {
                    double expectedX = control.GetProperty("Cells").GetProperty(cell.Cell.Text)
                        .GetProperty("OfficeX").GetDouble();
                    TestAssert.True(Math.Abs(cell.X + cell.ContentPaddingLeft - expectedX) <= 1d,
                        "The cell keeps the one-point Word origin gate: " + control.GetProperty("Id").GetString());
                }
            }
        }
    }
}
