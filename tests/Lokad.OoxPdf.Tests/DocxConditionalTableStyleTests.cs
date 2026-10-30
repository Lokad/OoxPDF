using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxConditionalTableStyleTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly string Cases = Path.GetFullPath(Path.Combine(
        AppContext.BaseDirectory, "..", "..", "..", "Cases", "ConditionalTableStyles"));

    public static void DocxConditionalTableStylesMatchIndependentOfficeFills()
    {
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var control in references.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string path = Path.Combine(Cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            AssertOfficeFills(Read(path), control);
        }
    }

    public static void DocxConditionalTableStylesRetainCascadeAcrossInheritanceAndSourceOrder()
    {
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var control in references.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string path = Path.Combine(Cases, control.GetProperty("Input").GetString()!);
            var transformed = Read(path, (document, styles) =>
            {
                foreach (var style in styles.Root!.Elements(W + "style").ToArray())
                {
                    string id = (string)style.Attribute(W + "styleId")!;
                    var regions = style.Elements(W + "tblStylePr").Reverse().ToArray();
                    foreach (var region in regions) region.Remove();
                    styles.Root.Add(new XElement(W + "style", new XAttribute(W + "type", "table"),
                        new XAttribute(W + "styleId", id + "Base"), regions));
                    style.AddFirst(new XElement(W + "basedOn", new XAttribute(W + "val", id + "Base")));
                }
            });
            AssertOfficeFills(transformed, control);
        }
    }

    public static void DocxConditionalTableStylesKeepFalseFlagsAndDirectFormatting()
    {
        var source = Read(Path.Combine(Cases, "explicit.docx"), (document, styles) =>
        {
            var style = styles.Root!.Elements(W + "style").First();
            style.Element(W + "tblPr")!.AddAfterSelf(new XElement(W + "tcPr",
                new XElement(W + "shd", new XAttribute(W + "fill", "123456"))));
            // All explicit false values suppress positional membership, even
            // at a real first-row/first-column intersection.
            var cells = document.Descendants(W + "tc").Take(2).ToArray();
            foreach (var flag in cells[0].Descendants(W + "cnfStyle").Single().Attributes())
                flag.Value = "0";
            cells[1].Element(W + "tcPr")!.Add(new XElement(W + "shd", new XAttribute(W + "fill", "FEDCBA")));
            // The same cascade also merges run properties. Direct formatting
            // remains above the resolved conditional style.
            foreach (var region in style.Elements(W + "tblStylePr"))
                region.AddFirst(new XElement(W + "rPr", new XElement(W + "color",
                    new XAttribute(W + "val", (string)region.Descendants(W + "shd").Single().Attribute(W + "fill")!))));
            cells[1].Descendants(W + "rPr").Single().Add(new XElement(W + "color", new XAttribute(W + "val", "010203")));
        });
        var cells = source.Tables[0].Rows[0].Cells;
        TestAssert.Equal("123456", cells[0].FillHex);
        TestAssert.True(cells[0].ConditionalFormat?.FirstRow == false &&
            cells[0].ConditionalFormat?.FirstColumn == false, "Explicit false source flags remain false.");
        TestAssert.Equal("FEDCBA", cells[1].FillHex);
        TestAssert.Equal("010203", cells[1].Paragraphs[0].Runs[0].ColorHex);
        TestAssert.Equal("0066CC", cells[2].Paragraphs[0].Runs[0].ColorHex);
    }

    public static void DocxConditionalTableStylesKeepTrailingPropertiesUnderLeadingEdges()
    {
        var document = Read(Path.Combine(Cases, "single-cell.docx"), (_, styles) =>
        {
            var regions = styles.Root!.Elements(W + "style").First()
                .Elements(W + "tblStylePr").ToDictionary(e => (string)e.Attribute(W + "type")!);
            regions["lastRow"].AddFirst(new XElement(W + "rPr",
                new XElement(W + "b", new XAttribute(W + "val", "0")), new XElement(W + "i")));
            regions["firstRow"].AddFirst(new XElement(W + "rPr", new XElement(W + "b")));
            regions["lastCol"].Element(W + "tcPr")!.Add(new XElement(W + "vAlign", new XAttribute(W + "val", "bottom")),
                new XElement(W + "tcMar", new XElement(W + "left", new XAttribute(W + "w", "100"), new XAttribute(W + "type", "dxa"))));
            regions["firstCol"].Element(W + "tcPr")!.Add(new XElement(W + "vAlign", new XAttribute(W + "val", "top")));
        });
        var cell = document.Tables[0].Rows[0].Cells[0];
        var run = cell.Paragraphs[0].Runs[0];
        TestAssert.Equal("0066CC", cell.FillHex);
        TestAssert.Equal("top", cell.VerticalAlignmentValue);
        TestAssert.Equal(5d, cell.Margins.LeftPoints!.Value);
        TestAssert.True(run.Bold && run.Italic,
            "Leading formatting overrides conflicting properties while retaining independent trailing properties.");
    }

    private static void AssertOfficeFills(DocxDocument document, JsonElement control)
    {
        var expected = control.GetProperty("Fills");
        int count = 0;
        foreach (var cell in document.Tables.SelectMany(table => table.Rows).SelectMany(row => row.Cells))
        {
            TestAssert.Equal(expected.GetProperty(cell.Text).GetString(), cell.FillHex);
            count++;
        }
        TestAssert.Equal(expected.EnumerateObject().Count(), count);
    }

    private static DocxDocument Read(string path, Action<XDocument, XDocument>? transform = null)
    {
        string? temporary = null;
        try
        {
            if (transform is not null)
            {
                var parts = new Dictionary<string, string>();
                using (var archive = ZipFile.OpenRead(path))
                    foreach (var entry in archive.Entries)
                    {
                        using var reader = new StreamReader(entry.Open());
                        parts[entry.FullName] = reader.ReadToEnd();
                    }
                var document = XDocument.Parse(parts["word/document.xml"]);
                var styles = XDocument.Parse(parts["word/styles.xml"]);
                transform(document, styles);
                parts["word/document.xml"] = document.ToString();
                parts["word/styles.xml"] = styles.ToString();
                path = temporary = TestFixtures.WriteTempPackage(".docx", parts);
            }
            using var stream = File.OpenRead(path);
            var package = OoxPackage.Open(stream, default);
            return new DocxReader().Read(package, null, default, OoxPdfDocxMarkupMode.Final);
        }
        finally { if (temporary is not null) File.Delete(temporary); }
    }
}
