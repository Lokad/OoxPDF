using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxCommentsTests
{
    public static void DocxWordCompatiblePlainParagraphCountsMatchIndependentOffice()
    {
        if (!HasUsableAptosBodyFace() || MeasureSegoeTitleDesignAdvance() < 78d)
            TestAssert.Skip("The Office paragraph controls require installed Aptos and matching Segoe UI Bold.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "plain-comment-paragraphs-office.json")));
        string template = Path.Combine(cases, reference.RootElement.GetProperty("Template").GetString()!);
        TestAssert.Equal(reference.RootElement.GetProperty("TemplateSha256").GetString(),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(template))));
        double tolerance = reference.RootElement.GetProperty("MaximumOriginErrorPoints").GetDouble();
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        foreach (JsonElement row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var parts = new Dictionary<string, byte[]>();
            using (var archive = ZipFile.OpenRead(template))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    using Stream source = entry.Open();
                    using var data = new MemoryStream();
                    source.CopyTo(data);
                    parts.Add(entry.FullName, data.ToArray());
                }
            }
            XDocument comments = XDocument.Parse(Encoding.UTF8.GetString(parts["word/comments.xml"]));
            XElement comment = comments.Root!.Elements(w + "comment").ElementAt(1);
            comment.RemoveNodes();
            foreach (JsonElement text in row.GetProperty("ParagraphTexts").EnumerateArray())
                comment.Add(new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", text.GetString()))));
            parts["word/comments.xml"] = Encoding.UTF8.GetBytes(comments.ToString(SaveOptions.DisableFormatting));
            using MemoryStream stream = TestFixtures.CreateZipPackage(parts);
            var document = new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null,
                CancellationToken.None, markupMode: OoxPdfDocxMarkupMode.AllMarkup);
            var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderPagesAsync(document, null, CancellationToken.None).ToBlockingEnumerable().Single();
            var actual = ReadEmbeddedGlyphTextShows(page)
                .Where(show => Regex.IsMatch(show.Text, @"^Paragraph\s+[0-9]+\."))
                .OrderBy(show => int.Parse(Regex.Match(show.Text, @"^Paragraph\s+([0-9]+)\.").Groups[1].Value))
                .ToArray();
            JsonElement[] expected = row.GetProperty("Origins").EnumerateArray().ToArray();
            TestAssert.Equal(expected.Length, actual.Length);
            for (int i = 0; i < expected.Length; i++)
            {
                TestAssert.True(Math.Abs(actual[i].X - expected[i].GetProperty("X").GetDouble()) <= tolerance &&
                    Math.Abs(page.Height - actual[i].Y - expected[i].GetProperty("Y").GetDouble()) <= tolerance,
                    $"{row.GetProperty("Id").GetString()} paragraph {i + 1} must retain its independent Office origin.");
            }
        }
    }
}
