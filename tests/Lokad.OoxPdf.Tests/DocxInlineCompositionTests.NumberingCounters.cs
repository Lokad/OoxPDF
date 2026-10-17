using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNumberingCountersRetainSourceTimelines()
    {
        using var golden = ReadNumberingCounterGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            // Both independent-list Word outcomes stay in the retained visual
            // batch. Concrete-ID continuation is an inherited, separate defect.
            if (!row.GetProperty("QualifiesInsertedMarkTimeline").GetBoolean()) continue;
            var document = ReadNumberingCounterDocument(row);
            var paragraphs = NumberedCounterParagraphs(document);
            var expected = row.GetProperty("Paragraphs").EnumerateArray().ToArray();
            TestAssert.Equal(3, paragraphs.Length);
            for (int i = 0; i < paragraphs.Length; i++)
            {
                var label = paragraphs[i].ListLabel!;
                var paints = expected[i].GetProperty("Paints").EnumerateArray().ToArray();
                TestAssert.Equal(paints[^1].GetProperty("Text").GetString(), label.Text);
                TestAssert.Equal(string.Concat(paints.Select(paint => paint.GetProperty("Text").GetString())), label.DisplayText);
                TestAssert.Equal(paints.Length == 2 ? paints[0].GetProperty("Text").GetString() : null,
                    label.NumberingChange?.PreviousText);
                if (label.NumberingChange is { } change)
                {
                    TestAssert.True(paragraphs.Take(i).Any(p => p.ParagraphMarkRun?.Revisions.Any(
                        revision => ReferenceEquals(revision, change.SourceMark)) == true),
                        "A changed counter retains a preceding actual mark as its source trigger.");
                    TestAssert.True(!paragraphs[i].Revisions.Any(revision => ReferenceEquals(revision, change.SourceMark)),
                        "Derived counters must not invent paragraph revisions or formatting balloons.");
                }
            }
        }
        // A different mark operation cannot establish an original/current counter
        // timeline. Preserve the existing fallback instead of inventing numbers.
        var unsupported = ReadMutation("CompositionProbes/numbering-counter-second-mark", OoxPdfDocxMarkupMode.AllMarkup,
            false, false, (xml, w, m) =>
            {
                var insertion = xml.Descendants(w + "ins").Single();
                insertion.Name = w + "del";
            }, partOverride: "word/footer1.xml");
        TestAssert.True(NumberedCounterParagraphs(unsupported).All(p => p.ListLabel!.NumberingChange is null),
            "Unsupported original numbering ownership keeps the existing current-counter path.");
    }

    public static void DocxInlineCompositionNumberingCounterPaintsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native numbering frames use installed Word-compatible fonts.");
        CheckNumberingCounterContent(null);
    }

    internal static void CheckNumberingCounterContent(IFontResolver? resolver)
    {
        using var golden = ReadNumberingCounterGolden();
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            // Both independent-list Word outcomes stay in the retained visual
            // batch. Concrete-ID continuation is an inherited, separate defect.
            if (!row.GetProperty("QualifiesInsertedMarkTimeline").GetBoolean()) continue;
            var document = ReadNumberingCounterDocument(row);
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var paragraphs = NumberedCounterParagraphs(document);
            var owners = paragraphs.SelectMany(p => p.Runs).ToArray();
            var page = new DocxRenderer(resolver, final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup,
                final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true).ToArray();
            var numbers = shows.Where(show => Regex.IsMatch(show.Text.Trim(), @"^\d+\.$")).ToArray();
            var expectedParagraphs = row.GetProperty("Paragraphs").EnumerateArray().ToArray();
            var expectedNumbers = expectedParagraphs.SelectMany(p => p.GetProperty("Paints").EnumerateArray()).ToArray();
            TestAssert.Equal(string.Join(" ", expectedNumbers.Select(value => value.GetProperty("Text").GetString())),
                string.Join(" ", numbers.Select(show => show.Text.Trim())));
            if (resolver is null)
            {
                for (int i = 0; i < numbers.Length; i++)
                    TestAssert.True(Math.Abs(numbers[i].X - expectedNumbers[i].GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(numbers[i].Y - expectedNumbers[i].GetProperty("Y").GetDouble()) < 1d,
                        "Old/current counters retain the independent one-point Word frame: " + row.GetProperty("Id").GetString());
                foreach (var expected in expectedParagraphs)
                {
                    string prefix = expected.GetProperty("TextPrefix").GetString()!;
                    var text = shows.Single(show => show.Text.StartsWith(prefix, StringComparison.Ordinal));
                    TestAssert.True(Math.Abs(text.X - expected.GetProperty("TextX").GetDouble()) < 1d &&
                        Math.Abs(text.Y - expected.GetProperty("TextY").GetDouble()) < 1d,
                        "The whole measured prefix, suffix and next tab stop own text placement: " + row.GetProperty("Id").GetString());
                }
            }
            TestAssert.True(numbers.All(show => show.X >= 0d && show.X < page.Width && show.Y > 0d && show.Y < page.Height),
                "Original/current numbering must remain on the page.");
            TestAssert.True(owners.Zip(paragraphs.SelectMany(p => p.Runs)).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Counter layout retains source text objects.");
        }
    }

    private static JsonDocument ReadNumberingCounterGolden() =>
        JsonDocument.Parse(File.ReadAllText(NumberingMarkCase("numbering-counters-office.json")));

    private static DocxDocument ReadNumberingCounterDocument(JsonElement row)
    {
        string path = NumberingMarkCase(row.GetProperty("File").GetString()!);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        return DocxTests.ReadDocx(path, row.GetProperty("Mode").GetString() == "final-preserve"
            ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup);
    }

    private static DocxParagraph[] NumberedCounterParagraphs(DocxDocument document) =>
        DocxBlockTraversal.EnumerateSourceParagraphs(document).Where(p => p.ListLabel is not null)
            .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance).ToArray();
}
