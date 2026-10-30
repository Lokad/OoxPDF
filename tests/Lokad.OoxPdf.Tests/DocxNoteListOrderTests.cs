using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNoteListOrderTests
{
    private static readonly string Cases = Path.GetFullPath(Path.Combine(
        AppContext.BaseDirectory, "..", "..", "..", "Cases", "NoteListOrder"));

    public static void DocxNoteListLabelsFollowIndependentWordReferenceOrder()
    {
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var control in reference.RootElement.GetProperty("Controls").EnumerateArray())
        {
            string input = Path.Combine(Cases, control.GetProperty("Input").GetString()!);
            TestAssert.Equal(control.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
                AssertLabels(DocxTests.ReadDocx(input, mode), control);
        }
    }

    public static void DocxNoteListOrderKeepsCachedXmlAndSourceStoryOrder()
    {
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var control in reference.RootElement.GetProperty("Controls").EnumerateArray())
        {
            using var stream = File.OpenRead(Path.Combine(Cases, control.GetProperty("Input").GetString()!));
            var package = OoxPackage.Open(stream, default);
            var part = package.GetPart(control.GetProperty("Part").GetString()!)!;
            string before = package.LoadXml(part, default).ToString(SaveOptions.DisableFormatting);
            foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupMode.Final })
            {
                var document = new DocxReader().Read(package, null, default, mode);
                AssertLabels(document, control);
                TestAssert.Equal(string.Join(',', control.GetProperty("SourceNormalIds").EnumerateArray().Select(id => id.GetString())),
                    string.Join(',', document.RelatedStories.Where(story => story.IsNormalStoryType).Select(story => story.Id)));
                TestAssert.Equal(before, package.LoadXml(part, default).ToString(SaveOptions.DisableFormatting));
            }
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => new DocxReader().Read(package, null, cancellation.Token, OoxPdfDocxMarkupMode.Final));
        }
    }

    internal static void CheckRendered(IFontResolver resolver)
    {
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));
        foreach (var control in reference.RootElement.GetProperty("Controls").EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = DocxTests.ReadDocx(Path.Combine(Cases, control.GetProperty("Input").GetString()!), mode);
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(source, null, default).ToArray();
            TestAssert.Equal(control.GetProperty("ExpectedPages").GetInt32(), pages.Length);
            string text = string.Concat(pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)).Select(show => show.Text));
            TestAssert.Equal(1, Regex.Matches(text, @"(?<![0-9])1\.").Count);
            TestAssert.Equal(control.GetProperty("ExpectedLabels").GetArrayLength() - 1,
                Regex.Matches(text, @"(?<![0-9])2\.").Count);
            foreach (string expected in new[] { "Public note calculation", "rate", "time", "flow", "days" })
                TestAssert.Contains(expected, text);
            TestAssert.True(!text.Contains("Unused list entry", StringComparison.Ordinal), "Unreferenced note content remains unpainted.");
            TestAssert.Equal(1, pages.Sum(page => page.Images.Sum(image =>
                Regex.Matches(page.Content, "/" + Regex.Escape(image.ResourceName) + @"\s+Do\b").Count)));
        }
    }

    private static void AssertLabels(DocxDocument document, JsonElement control)
    {
        foreach (var expected in control.GetProperty("ExpectedLabels").EnumerateArray())
        {
            var story = document.RelatedStories.Single(story => story.Id == expected.GetProperty("Id").GetString());
            var labels = DocxBlockTraversal.EnumerateBodyParagraphs(story).Where(paragraph => paragraph.ListLabel is not null)
                .Select(paragraph => paragraph.ListLabel!.Text);
            TestAssert.Equal(string.Join(',', expected.GetProperty("Labels").EnumerateArray().Select(label => label.GetString())), string.Join(',', labels));
        }
    }
}
