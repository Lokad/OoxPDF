using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionListLabelStylesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent emitted label faces use installed Word-compatible fonts.");
        CheckListLabelStyles(null);
    }

    internal static void CheckListLabelStyles(IFontResolver? resolver)
    {
        string folder = TestFixtures.PublicCasePath("CompositionProbes");
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "list-styles-office.json")));
        foreach (var row in RequireListLabelReferenceRows(golden, "list-styles-office.json"))
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = Path.Combine(folder, row.GetProperty("File").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = DocxTests.ReadDocx(path, mode);
            var paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements)
                .Concat(document.HeaderParagraphs).Concat(document.FooterParagraphs).Where(p => p.ListLabel is not null).ToArray();
            var labels = row.GetProperty("Labels").EnumerateArray().ToArray();
            TestAssert.Equal(3, paragraphs.Length);
            foreach (var paragraph in paragraphs)
            {
                var label = paragraph.ListLabel!;
                var expected = labels.Single(value => value.GetProperty("Text").GetString() == label.Text);
                var run = DocxLayoutEngine.CreateListLabelRun(label, paragraph.Runs.FirstOrDefault(), 11d);
                TestAssert.Equal(expected.GetProperty("Bold").GetBoolean(), run.Bold);
                TestAssert.Equal(expected.GetProperty("Italic").GetBoolean(), run.Italic);
            }
            var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
            var paints = ReadListLabelPaintFonts(page, labels, id);
            if (resolver is null)
            {
                for (int i = 0; i < labels.Length; i++)
                {
                    TestAssert.Equal(labels[i].GetProperty("Bold").GetBoolean(), paints[i].Font.Os2.WeightClass >= 600);
                    TestAssert.Equal(labels[i].GetProperty("Italic").GetBoolean(), Math.Abs(paints[i].Font.Post.ItalicAngle) > .01d);
                }
            }
        }
    }
}
