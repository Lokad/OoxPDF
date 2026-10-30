using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTableLookTests
{
    public static void DocxTableLookDeclarationsMatchIndependentWordFills()
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "Cases", "TableLook"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "office.json")));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string path = Path.Combine(cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var document = DocxTests.ReadDocx(path, mode);
                int count = 0;
                foreach (var table in document.Tables)
                {
                    var raw = control.GetProperty("LookTokens");
                    TestAssert.Equal(raw.GetProperty("val").GetString(), table.Look?.Value);
                    string?[] values = [table.Look?.FirstRowValue, table.Look?.LastRowValue,
                        table.Look?.FirstColumnValue, table.Look?.LastColumnValue,
                        table.Look?.NoHorizontalBandValue, table.Look?.NoVerticalBandValue];
                    string[] names = ["firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand"];
                    for (int i = 0; i < names.Length; i++)
                        TestAssert.Equal(raw.GetProperty(names[i]).GetString(), values[i]);
                    foreach (var cell in table.Rows.SelectMany(row => row.Cells))
                    {
                        TestAssert.Equal(control.GetProperty("Fills").GetProperty(cell.Text).GetString(), cell.FillHex);
                        count++;
                    }
                }
                TestAssert.Equal(27, count);
            }
        }
    }
}
