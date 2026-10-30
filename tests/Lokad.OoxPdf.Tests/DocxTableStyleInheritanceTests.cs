using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTableStyleInheritanceTests
{
    public static void DocxConditionalInheritanceMatchesIndependentWordProperties()
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "Cases", "TableStyleInheritance"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "office.json")));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string path = Path.Combine(cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
            {
                var document = DocxTests.ReadDocx(path, mode);
                var table = document.BodyElements.OfType<DocxTableElement>().Single().Table;
                var cell = table.Rows[0].Cells[0];
                TestAssert.Equal(control.GetProperty("LeftPadding").GetDouble(), cell.Margins.LeftPoints);
                TestAssert.Equal(control.GetProperty("LeftPadding").GetDouble(), cell.Margins.RightPoints);
                TestAssert.Equal(control.GetProperty("TextColor").GetString(),
                    cell.Paragraphs[0].Runs[0].ColorHex ?? "000000");
                TestAssert.Equal(control.GetProperty("HasInheritedBorders").GetBoolean(), cell.Borders.Count != 0);
            }
        }
    }
}
