using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTableLookTests
{
    public static void DocxTableLookLegacyFlagsMatchIndependentWordFills()
    {
        foreach (var (control, document) in ReadWordControls("legacy-office.json"))
            AssertWordFills(control, document.Tables, expectedCellCount: 22);
    }

    public static void DocxTableLookDeclarationsMatchIndependentWordFills()
    {
        foreach (var (control, document) in ReadWordControls("office.json"))
        {
            var tables = document.Tables;
            var raw = control.GetProperty("LookTokens");
            foreach (var table in tables)
            {
                TestAssert.Equal(raw.GetProperty("val").GetString(), table.Look?.Value);
                string?[] values = [table.Look?.FirstRowValue, table.Look?.LastRowValue,
                    table.Look?.FirstColumnValue, table.Look?.LastColumnValue,
                    table.Look?.NoHorizontalBandValue, table.Look?.NoVerticalBandValue];
                string[] names = ["firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand"];
                for (int i = 0; i < names.Length; i++)
                    TestAssert.Equal(raw.GetProperty(names[i]).GetString(), values[i]);
            }
            AssertWordFills(control, tables, expectedCellCount: 27);
        }
    }

    private static IEnumerable<(JsonElement Control, DocxDocument Document)> ReadWordControls(string manifest)
    {
        string cases = TestFixtures.PublicCasePath("TableLook");
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, manifest)));
        foreach (var control in oracle.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string path = Path.Combine(cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
                yield return (control, DocxTests.ReadDocx(path, mode));
        }
    }

    private static void AssertWordFills(JsonElement control, IReadOnlyList<DocxTable> tables, int expectedCellCount)
    {
        var remaining = control.GetProperty("Fills").EnumerateObject()
            .ToDictionary(fill => fill.Name, fill => fill.Value.GetString(), StringComparer.Ordinal);
        TestAssert.Equal(expectedCellCount, remaining.Count);
        string? id = control.GetProperty("Id").GetString();
        foreach (var cell in tables.SelectMany(table => table.Rows).SelectMany(row => row.Cells))
        {
            string label = DocxTests.GetCellText(cell);
            TestAssert.True(remaining.Remove(label, out string? expected),
                $"Word table look {id}: unexpected or duplicate cell label {label}.");
            TestAssert.True(expected == cell.FillHex,
                $"Word table look {id}/{label}: expected {expected}, got {cell.FillHex}.");
        }
        TestAssert.True(remaining.Count == 0, $"Word table look {id}: missing cells {string.Join(", ", remaining.Keys)}.");
    }
}
