using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionFormattingBalloonDescriptionsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent descriptions use installed Word-compatible text fonts.");
        CheckRunFormattingBalloonDescriptions(null);
    }

    internal static void CheckRunFormattingBalloonDescriptions(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "formatting-balloons-office-text.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string path = Path.Combine(folder, row.GetProperty("Id").GetString()! + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            var document = DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.AllMarkup);
            var owners = document.Paragraphs.SelectMany(paragraph => paragraph.Runs).ToArray();
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var title = shows.Single(show => show.Text.Trim() == row.GetProperty("Title").GetString());
            var body = shows.Single(show => show.Text.Trim() == row.GetProperty("Body").GetString());
            TestAssert.True(Math.Abs(title.Y - body.Y) < .001d, "The formatting title and description share a row.");
            if (resolver is null)
            {
                TestAssert.True(Math.Abs(title.X - row.GetProperty("TitleX").GetDouble()) < 1d,
                    "Formatting title keeps the independent physical review-lane X.");
                TestAssert.True(Math.Abs(body.X - row.GetProperty("BodyX").GetDouble()) < 1d,
                    "The description starts after the title and its visible separator space.");
            }
            TestAssert.True(owners.Zip(document.Paragraphs.SelectMany(paragraph => paragraph.Runs))
                .All(pair => ReferenceEquals(pair.First, pair.Second)), "Balloon descriptions cannot rewrite source owners.");
        }
    }

    public static void DocxInlineCompositionFormattingBalloonDescriptionsResolveInheritedState()
    {
        foreach (string token in new[] { "1", "true", "on" })
        {
            var document = ReadMutation("CompositionProbes/math-run-format-body-bold", OoxPdfDocxMarkupMode.AllMarkup,
                false, false, (xml, w, m) =>
                {
                    var change = xml.Descendants(w + "rPrChange").Single();
                    change.Parent!.Element(w + "b")!.SetAttributeValue(w + "val", token);
                    change.Element(w + "rPr")!.Element(w + "b")!.SetAttributeValue(w + "val", "off");
                });
            var paragraph = document.Paragraphs.First(item => item.Equations.Count != 0);
            TestAssert.Equal("Font: Bold", DocxRenderer.BuildWordCompatibleRunFormattingPreview(paragraph, default));
        }
        var inherited = ReadMutation("CompositionProbes/math-run-format-body-bold", OoxPdfDocxMarkupMode.AllMarkup,
            false, false, (xml, w, m) =>
            {
                var normal = new XElement(w + "style", new XAttribute(w + "type", "paragraph"), new XAttribute(w + "styleId", "Normal"));
                xml.Root!.Add(normal);
                var properties = normal.Element(w + "rPr");
                if (properties is null) { properties = new XElement(w + "rPr"); normal.Add(properties); }
                properties.Elements(w + "i").Remove(); properties.Add(new XElement(w + "i"));
            }, partOverride: "word/styles.xml");
        var source = inherited.Paragraphs.First(item => item.Equations.Count != 0);
        TestAssert.True(source.Runs.First(run => run.Revisions.Count != 0).Italic, "The paragraph style supplies inherited italic.");
        TestAssert.Equal("Font: Bold", DocxRenderer.BuildWordCompatibleRunFormattingPreview(source, default));
    }

    public static void DocxInlineCompositionFormattingBalloonDescriptionsKeepExplicitFallbacks()
    {
        foreach (string variant in new[] { "unknown-property", "malformed", "other-change", "duplicate-id", "character-style" })
        {
            var document = ReadMutation("CompositionProbes/math-run-format-body-bold", OoxPdfDocxMarkupMode.AllMarkup,
                false, false, (xml, w, m) =>
                {
                    var change = xml.Descendants(w + "rPrChange").Single(); var properties = change.Parent!;
                    if (variant == "unknown-property") properties.Add(new XElement(w + "effect", new XAttribute(w + "val", "blinkBackground")));
                    if (variant == "malformed") properties.Element(w + "b")!.SetAttributeValue(w + "val", "perhaps");
                    if (variant == "other-change") properties.Add(new XElement(w + "spacing", new XAttribute(w + "val", "40")));
                    if (variant == "duplicate-id") properties.Add(new XElement(change));
                    if (variant == "character-style") properties.Add(new XElement(w + "rStyle", new XAttribute(w + "val", "Strong")));
                });
            var paragraph = document.Paragraphs.First(item => item.Equations.Count != 0);
            TestAssert.True(DocxRenderer.BuildWordCompatibleRunFormattingPreview(paragraph, default) is null,
                "An incomplete formatting description retains its generic summary: " + variant);
        }
        var original = ReadMutation("CompositionProbes/math-run-format-body-bold", OoxPdfDocxMarkupMode.AllMarkup, false, false)
            .Paragraphs.First(item => item.Equations.Count != 0);
        TestAssert.True(DocxRenderer.BuildWordCompatibleRunFormattingPreview(original with { Revisions = [original.Revisions.Single() with { }] }, default) is null,
            "A value-equal revision copy cannot acquire source ownership.");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxRenderer.BuildWordCompatibleRunFormattingPreview(original, canceled.Token));
    }
}
