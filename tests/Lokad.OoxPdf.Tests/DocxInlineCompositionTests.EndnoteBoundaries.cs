using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionFittingEndnotesOwnBoundariesWithoutChangingTypography()
    {
        foreach (string container in new[] { "paragraph", "table" })
        {
            var source = ReadMutation("CompositionProbes/mixed-note-endnote-" + container, OoxPdfDocxMarkupMode.Final, false, false);
            var math = DocxBlockTraversal.EnumerateBodyParagraphs(source.RelatedStories.Single(story => story.Id == "37"))
                .SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                    equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            var notes = layout.Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(story => story.StoryLayout.Story.IsNormalStoryType).ToArray();
            TestAssert.Equal(4, notes.Length);
            TestAssert.True(notes.All(story => story.UsesCompleteDocumentEndnoteFlow),
                "Complete measured paragraph/table blocks own their source boundaries.");
            TestAssert.True(notes.All(story => !story.UsesPlainDocumentEndnoteTypography),
                "Placement does not promote compound content into the narrower typography contract.");
        }
    }

    public static void DocxInlineCompositionFittingEndnotesRetainComplexFrameFallbacks()
    {
        foreach (string variant in new[] { "keep", "overflow", "static", "scaled", "clipped", "nested" })
        {
            bool table = variant is "clipped" or "nested";
            var source = ReadMutation("CompositionProbes/mixed-note-endnote-" + (table ? "table" : "paragraph"),
                OoxPdfDocxMarkupMode.Final, false, false, (xml, w, _) =>
                {
                    var story = xml.Root!.Elements(w + "endnote").Single(story => (string?)story.Attribute(w + "id") == "37");
                    if (variant == "keep") story.Elements(w + "p").First().Element(w + "pPr")!.Add(new XElement(w + "keepLines"));
                    if (variant == "overflow") story.Descendants(w + "t").First(text => text.Value.Contains("Public destination Alpha", StringComparison.Ordinal))
                        .Value += string.Concat(Enumerable.Repeat(" Additional note words.", 200));
                    if (variant == "clipped") story.Element(w + "tbl")!.Element(w + "tr")!.AddFirst(new XElement(w + "trPr",
                        new XElement(w + "trHeight", new XAttribute(w + "val", "20"), new XAttribute(w + "hRule", "exact"))));
                    if (variant == "nested")
                    {
                        var paragraph = story.Element(w + "tbl")!.Descendants(w + "tc").Single().Element(w + "p")!;
                        paragraph.ReplaceWith(new XElement(w + "tbl", new XElement(w + "tr", new XElement(w + "tc", paragraph))));
                    }
                }, partOverride: "word/endnotes.xml");
            if (variant == "static") source = source with { HeaderParagraphs = [DocxTests.CreateDocxLayoutParagraph("Public header", 12d, 14d)] };
            var math = source.RelatedStories.SelectMany(story => DocxBlockTraversal.EnumerateBodyParagraphs(story))
                .SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                    equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var layout = new DocxLayoutEngine(variant == "scaled" ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout, wordCompatiblePrintScale: .8d)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            TestAssert.True(!layout.Pages.SelectMany(page => page.PlacedRelatedStories).Any(story => story.UsesCompleteDocumentEndnoteFlow),
                "Unqualified whole blocks retain the existing continuation/frame policy: " + variant);
        }
    }

    public static void DocxInlineCompositionFittingEndnotesMatchIndependentBodyAndStoryBoundaries()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent endnote boundaries use installed Windows fonts.");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "endnote-fitting-office.json")));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(folder, row.GetProperty("CaseRelativePath").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = DocxTests.ReadDocx(input, mode);
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(source, null, default).ToArray();
            TestAssert.Equal(row.GetProperty("ExpectedPages").GetInt32(), pages.Length);
            var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Select(show => (Page: index + 1, Show: show))).ToArray();
            foreach (var expected in row.GetProperty("NoteAnchors").EnumerateArray().Append(row.GetProperty("BodyAnchor")))
            {
                string text = expected.GetProperty("Text").GetString()!;
                var actual = shows.Single(item => item.Show.Text.Trim() == text);
                TestAssert.Equal(expected.GetProperty("PageNumber").GetInt32(), actual.Page);
                TestAssert.True(Math.Abs(actual.Show.Y - expected.GetProperty("Y").GetDouble()) <= .5d,
                    "Actual source/story boundaries match independent Word within half a point: " + text);
            }
            string content = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(row.GetProperty("ExpectedEquations").GetInt32(), content.Split("ratetime", StringSplitOptions.None).Length - 1);
            var mathShows = pages.SelectMany((page, index) =>
            {
                var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                    .Select(resource => resource.Font).ToHashSet();
                return DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Where(show => fonts.Contains(show.Font)).Select(show => (Page: index + 1, Show: show));
            }).ToArray();
            var mathAnchors = row.GetProperty("MathAnchors").EnumerateArray().ToArray();
            TestAssert.Equal(mathAnchors.Length * 4, mathShows.Length);
            for (int index = 0; index < mathAnchors.Length; index++)
            {
                var actual = mathShows[index * 4];
                TestAssert.Equal(mathAnchors[index].GetProperty("Text").GetString(), string.Concat(mathShows.Skip(index * 4).Take(4)
                    .Select(item => item.Show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(mathAnchors[index].GetProperty("PageNumber").GetInt32(), actual.Page);
                TestAssert.True(Math.Abs(actual.Show.X - mathAnchors[index].GetProperty("X").GetDouble()) <= .5d &&
                    Math.Abs(actual.Show.Y - mathAnchors[index].GetProperty("Y").GetDouble()) <= .5d,
                    "Final mathematical origins follow the independently validated source/story boundaries.");
            }
            var links = pages.SelectMany(page => page.Annotations).ToArray();
            TestAssert.Equal(4, links.Length);
            TestAssert.Equal(4, links.Select(link => link.Destination).Distinct().Count());
        }
    }
}
