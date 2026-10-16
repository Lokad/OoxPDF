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
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "list-styles-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = Path.Combine(folder, row.GetProperty("File").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = DocxTests.ReadDocx(path, mode);
            var paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements)
                .Concat(document.HeaderParagraphs).Concat(document.FooterParagraphs).Where(p => p.ListLabel is not null).ToArray();
            var owners = paragraphs.SelectMany(p => p.Runs).ToArray();
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
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            foreach (var expected in labels)
            {
                string text = expected.GetProperty("Text").GetString()!;
                var paints = shows.Where(show => show.Text.Trim() == text).ToArray();
                TestAssert.True(paints.Length != 0 && paints.All(show => show.Y > 0d && show.Y < page.Height),
                    "Resolved labels keep their painted source row: " + id);
                if (resolver is null)
                    foreach (var paint in paints)
                    {
                        TestAssert.Equal(expected.GetProperty("Bold").GetBoolean(), paint.Font.Font.Os2.WeightClass >= 600);
                        TestAssert.Equal(expected.GetProperty("Italic").GetBoolean(), Math.Abs(paint.Font.Font.Post.ItalicAngle) > .01d);
                    }
            }
            TestAssert.True(owners.Zip(paragraphs.SelectMany(p => p.Runs)).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Resolving label faces cannot rewrite source text formatting or revision owners.");
        }
    }
}
