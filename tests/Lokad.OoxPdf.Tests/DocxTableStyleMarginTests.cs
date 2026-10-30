using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTableStyleMarginTests
{
    private static readonly string Cases = Path.GetFullPath(Path.Combine(
        AppContext.BaseDirectory, "..", "..", "..", "Cases", "TableStyleMargins"));

    public static void DocxCustomTableMarginsMatchIndependentWordContentOrigins()
    {
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var control in reference.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string input = Path.Combine(Cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            foreach (var expected in control.GetProperty("Rows").EnumerateArray())
            {
                bool final = expected.GetProperty("Mode").GetString() == "final-preserve";
                var source = DocxTests.ReadDocx(input, final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup);
                var cell = FirstCell(source, final);
                double left = control.GetProperty("LeftMargin").ValueKind == JsonValueKind.Null
                    ? double.NaN : control.GetProperty("LeftMargin").GetDouble();
                double right = control.GetProperty("RightMargin").ValueKind == JsonValueKind.Null
                    ? double.NaN : control.GetProperty("RightMargin").GetDouble();
                TestAssert.True(double.IsNaN(left) ? cell.Cell.Margins.LeftPoints is null : cell.Cell.Margins.LeftPoints == left,
                    "The resolved left padding keeps its source/style owner.");
                TestAssert.True(double.IsNaN(right) ? cell.Cell.Margins.RightPoints is null : cell.Cell.Margins.RightPoints == right,
                    "The resolved right padding keeps its source/style owner.");
                double expectedLeft = control.GetProperty("RightAligned").GetBoolean()
                    ? 77.424d : expected.GetProperty("FirstX").GetDouble();
                TestAssert.True(Math.Abs(cell.X + cell.ContentPaddingLeft - expectedLeft) <= 1d,
                    "Cell content uses the unchanged one-point independent Word origin gate: " + control.GetProperty("Id").GetString());
            }
        }
    }

    public static void DocxCustomTableMarginsKeepIncompleteInheritanceFallbacks()
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        foreach (string parent in new[] { "MissingParent", "edges" })
        {
            var parts = new Dictionary<string, string>();
            using (var archive = ZipFile.OpenRead(Path.Combine(Cases, "named.docx")))
                foreach (var entry in archive.Entries)
                {
                    using var reader = new StreamReader(entry.Open());
                    parts[entry.FullName] = reader.ReadToEnd();
                }
            var styles = XDocument.Parse(parts["word/styles.xml"]);
            styles.Root!.Elements(w + "style").First().Element(w + "name")!
                .AddAfterSelf(new XElement(w + "basedOn", new XAttribute(w + "val", parent)));
            parts["word/styles.xml"] = styles.ToString();
            string input = TestFixtures.WriteTempPackage(".docx", parts);
            try
            {
                var cell = FirstCell(DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.Final), true);
                TestAssert.True(cell.Cell.Margins.LeftPoints is null, "An incomplete/cyclic chain retains the unset-margin fallback.");
                TestAssert.True(Math.Abs(cell.X + cell.ContentPaddingLeft - 72.48d) < 1e-9,
                    "Unknown inheritance keeps the previous bare-table anchor.");
            }
            finally { File.Delete(input); }
        }
    }

    private static DocxTableCellLayout FirstCell(DocxDocument source, bool final)
    {
        var layout = new DocxLayoutEngine(final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
            : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup,
            // These sources have no revisions or balloons. Like the renderer,
            // retain their full-size page frame in Word-compatible mode.
            wordCompatiblePrintScale: 1d)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
        return layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>().Concat(page.StaticTableRows))
            .SelectMany(row => row.Cells).First(cell => cell.Cell.Paragraphs.Any(paragraph =>
                paragraph.Runs.Any(run => run.Text == "edges-00")));
    }
}
