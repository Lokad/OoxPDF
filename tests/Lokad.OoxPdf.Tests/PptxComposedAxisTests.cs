using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxComposedAxisTests
{
    internal static void CheckLabels(bool hidePrimary)
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "ComposedAxes"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "office-labels.json")));
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            byte[] source = File.ReadAllBytes(Path.Combine(cases, row.GetProperty("Input").GetString()!));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(source)));
            using var stream = new MemoryStream();
            stream.Write(source); stream.Position = 0;
            if (hidePrimary)
            {
                using var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true);
                var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
                XDocument xml;
                using (var input = entry.Open()) xml = XDocument.Load(input);
                XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                var axis = xml.Descendants(c + "valAx").Single(axis =>
                    (string?)axis.Element(c + "axId")?.Attribute("val") == row.GetProperty("PrimaryAxisId").GetString());
                axis.Element(c + "tickLblPos")!.SetAttributeValue("val", "none");
                entry.Delete();
                using var output = archive.CreateEntry("ppt/charts/chart1.xml").Open();
                xml.Save(output);
            }
            stream.Position = 0;
            using var package = OoxPackage.Open(stream, default);
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
                new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
            string[] actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, fontResourceFilter:
                name => name.StartsWith("CVA", StringComparison.Ordinal)).Select(show => show.Text).Order(StringComparer.Ordinal).ToArray();
            string[] expected = row.GetProperty(hidePrimary ? "SecondaryLabels" : "OfficeNumericLabels")
                .EnumerateArray().Select(label => label.GetString()!).Order(StringComparer.Ordinal).ToArray();
            TestAssert.True(expected.SequenceEqual(actual),
                "Only actual visible source axes may emit numeric strips: " + row.GetProperty("Id").GetString() +
                "/hidePrimary=" + hidePrimary + "; expected " + string.Join(",", expected) + "; actual " + string.Join(",", actual));
        }
    }
}
