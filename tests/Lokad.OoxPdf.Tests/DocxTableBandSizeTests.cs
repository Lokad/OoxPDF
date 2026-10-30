using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTableBandSizeTests
{
    public static void DocxTableBandSizesMatchIndependentWordFills()
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "Cases", "TableBandSizes"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "office.json")));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string path = Path.Combine(cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var document = DocxTests.ReadDocx(path, mode);
                foreach (var cell in document.BodyElements.OfType<DocxTableElement>()
                    .SelectMany(element => element.Table.Rows).SelectMany(row => row.Cells))
                {
                    var expected = control.GetProperty("Fills").GetProperty(cell.Text);
                    TestAssert.Equal(expected.ValueKind == JsonValueKind.Null ? null : expected.GetString(), cell.FillHex);
                }
            }
        }
    }
}
