using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionSpanningPageNotesUseMarkerPagesAndKeepTypographyScope()
    {
        foreach (string typography in new[] { "calibri-12-zero", "arial-11-zero" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/spanning-page-notes-" + typography, mode, false, false);
            var paragraph = source.Paragraphs.Single();
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var first = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(first.Pages.Count > 1 && paragraph.InlineReferences.Count == 4,
                "The control has four actual references in one naturally paginated paragraph.");
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, first, out _, out _, allowWrappedBody: true),
                "Single-reference callers retain their independently qualified scope.");
            TestAssert.True(DocxPageFootnoteNumbering.TryCreateDocument(source, first, out var numbered, out var owners,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
                "Each complete marker/note pair owns a page independently of its paragraph's other lines.");
            TestAssert.True(owners.Values.Distinct().Count() > 1, "Actual reference ownership spans physical pages.");
            var counters = new Dictionary<int, int>();
            foreach (var reference in numbered.Paragraphs.Single().InlineReferences)
            {
                int page = owners[reference.Id!];
                counters.TryGetValue(page, out int prior);
                string expected = (9 + prior).ToString("D2", CultureInfo.InvariantCulture);
                counters[page] = prior + 1;
                TestAssert.Equal(expected, reference.DisplayText);
                TestAssert.Equal(expected, numbered.RelatedStories.Single(story => story.Id == reference.Id).Paragraphs[0].Runs[0].Text);
            }
            var generated = paragraph.InlineReferences.Select(reference => reference.SourceRunIndex).ToHashSet();
            TestAssert.True(paragraph.Runs.Where(run => !generated.Contains(run.SourceRunIndex)).All(run =>
                numbered.Paragraphs.Single().Runs.Any(retained => ReferenceEquals(run, retained))),
                "All authored prose and its source runs retain their identities.");
            TestAssert.Equal("105,106,107,108", string.Join(',', paragraph.InlineReferences.Select(reference => reference.DisplayText)));
            var final = engine.Create(numbered, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.True(DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, final, owners,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
                "A bounded numbering retry retains every physical marker/note page.");
            TestAssert.True(!DocxPageFootnoteNumbering.HasStablePages(numbered, final, owners,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
                "Typography admission retains the existing same-paragraph-page requirement.");
            TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(numbered, final, out var unchanged, out _,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences) && ReferenceEquals(numbered, unchanged),
                "Correct labels avoid another retry.");

            // A paragraph spanning pages cannot excuse an absent or duplicated
            // source marker, or a note placed on another physical page.
            var referenceOwner = paragraph.InlineReferences[0];
            int ownerIndex = owners[referenceOwner.Id!] - 1;
            var line = first.Pages[ownerIndex].Items.OfType<DocxTextLineLayout>().Single(item =>
                item.Segments.Any(segment => segment.SourceTextRunIndex == referenceOwner.SourceRunIndex &&
                    segment.SourceTextOffsetInRun == referenceOwner.TextOffsetInRun));
            var marker = line.Segments.Single(segment => segment.SourceTextRunIndex == referenceOwner.SourceRunIndex &&
                segment.SourceTextOffsetInRun == referenceOwner.TextOffsetInRun);
            foreach (bool duplicate in new[] { false, true })
            {
                var pages = first.Pages.ToArray();
                var segments = duplicate ? line.Segments.Append(marker).ToArray()
                    : line.Segments.Where(segment => !ReferenceEquals(segment, marker)).ToArray();
                pages[ownerIndex] = pages[ownerIndex] with {
                    Items = pages[ownerIndex].Items.Select(item => ReferenceEquals(item, line) ? line with { Segments = segments } : item).ToArray()
                };
                TestAssert.True(!DocxPageFootnoteNumbering.TryCreateDocument(source, first with { Pages = pages }, out var fallback, out _,
                    allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences) && ReferenceEquals(source, fallback),
                    "Missing or ambiguous source glyphs reject the entire numbering replacement.");
            }
            var movedPages = final.Pages.ToArray();
            var note = movedPages[ownerIndex].PlacedRelatedStories.Single(item => item.StoryLayout.Story.Id == referenceOwner.Id);
            int otherIndex = (ownerIndex + 1) % movedPages.Length;
            movedPages[ownerIndex] = movedPages[ownerIndex] with {
                PlacedRelatedStories = movedPages[ownerIndex].PlacedRelatedStories.Where(item => !ReferenceEquals(item, note)).ToArray()
            };
            movedPages[otherIndex] = movedPages[otherIndex] with { PlacedRelatedStories = [.. movedPages[otherIndex].PlacedRelatedStories, note] };
            TestAssert.True(!DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, final with { Pages = movedPages }, owners,
                allowWrappedBody: true, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences),
                "A moved note invalidates retry stability even when its paragraph spans both pages.");
        }
    }

    public static void DocxInlineCompositionSpanningPageNotesRenderCountersAndDistinctDestinations()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Natural pagination controls use installed Windows fonts.");
        CheckSpanningPageNotes(null);
    }

    public static void DocxInlineCompositionSpanningPageNotesMatchIndependentPageCapacity()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent capacity controls use installed Windows fonts.");
        CheckSpanningPageNotes(null, requireIndependentGeometry: true);
    }

    internal static void CheckSpanningPageNotes(IFontResolver? resolver, bool requireIndependentGeometry = false,
        string referenceFile = "spanning-page-notes-office.json")
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, referenceFile)));
        foreach (var row in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(folder, row.GetProperty("Id").GetString()! + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var source = DocxTests.ReadDocx(input, mode);
            var paragraph = source.Paragraphs.Single();
            double size = paragraph.Runs.Where(run => run.VerticalAlignment == DocxRunVerticalAlignment.Baseline && run.Text.Length != 0).Min(run => run.FontSize);
            var pages = new DocxRenderer(resolver, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(source, null, default).ToArray();
            TestAssert.True(pages.Length > 1, "The paragraph must cross a physical page boundary.");
            if (requireIndependentGeometry)
            {
                TestAssert.Equal(row.GetProperty("ExpectedPages").GetInt32(), pages.Length);
                var actual = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Where(show => show.Size < size * .8d && show.Text.Length > 0 && show.Text.All(char.IsAsciiDigit))
                    .Select(show => (Page: index + 1, Show: show))).ToArray();
                var expected = row.GetProperty("Markers").EnumerateArray().ToArray();
                TestAssert.Equal(expected.Length, actual.Length);
                for (int i = 0; i < actual.Length; i++)
                {
                    TestAssert.Equal(expected[i].GetProperty("PageNumber").GetInt32(), actual[i].Page);
                    TestAssert.Equal(expected[i].GetProperty("Text").GetString(), actual[i].Show.Text);
                    TestAssert.True(Math.Abs(expected[i].GetProperty("X").GetDouble() - actual[i].Show.X) <= .5d &&
                        Math.Abs(expected[i].GetProperty("Y").GetDouble() - actual[i].Show.Y) <= .5d,
                        "Source and target glyphs retain independent Word page geometry within half a point.");
                }
            }
            if (row.TryGetProperty("ExpectedEquations", out var expectedEquations))
            {
                string mathText = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText))
                    .Normalize(System.Text.NormalizationForm.FormKC);
                TestAssert.Equal(expectedEquations.GetInt32(), mathText.Split("ratetime", StringSplitOptions.None).Length - 1);
                int images = pages.Sum(page => page.Images.Sum(image => System.Text.RegularExpressions.Regex.Matches(page.Content,
                    "/" + System.Text.RegularExpressions.Regex.Escape(image.ResourceName) + @"\s+Do\b").Count));
                TestAssert.Equal(row.GetProperty("ImagePlacements").GetInt32(), images);
                if (requireIndependentGeometry)
                {
                    var glyphs = pages.SelectMany((page, index) =>
                    {
                        var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                            .Select(resource => resource.Font).ToHashSet();
                        return DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                            .Where(show => fonts.Contains(show.Font)).Select(show => (Page: index + 1, Show: show));
                    }).ToArray();
                    var anchors = row.GetProperty("MathAnchors").EnumerateArray().ToArray();
                    TestAssert.Equal(anchors.Length * 4, glyphs.Length);
                    for (int i = 0; i < anchors.Length; i++)
                    {
                        var anchor = glyphs[i * 4];
                        string mathWord = string.Concat(glyphs.Skip(i * 4).Take(4).Select(item => item.Show.Text))
                            .Normalize(System.Text.NormalizationForm.FormKC);
                        TestAssert.Equal(anchors[i].GetProperty("Text").GetString(), mathWord);
                        TestAssert.Equal(anchors[i].GetProperty("PageNumber").GetInt32(), anchor.Page);
                        TestAssert.True(Math.Abs(anchors[i].GetProperty("X").GetDouble() - anchor.Show.X) <= .5d &&
                            Math.Abs(anchors[i].GetProperty("Y").GetDouble() - anchor.Show.Y) <= .5d,
                            "Mixed-content fractions retain independent Word page geometry within half a point.");
                    }
                }
            }
            var destinations = pages.SelectMany(page => page.Annotations).Select(link => link.Destination).ToArray();
            TestAssert.Equal(4, destinations.Length);
            TestAssert.Equal(4, destinations.Distinct().Count());
            for (int index = 0; index < pages.Length; index++)
            {
                var page = pages[index];
                var links = page.Annotations.ToArray();
                var expected = Enumerable.Range(row.GetProperty("NumberStart").GetInt32(), links.Length)
                    .Select(value => value.ToString("D2", CultureInfo.InvariantCulture)).ToArray();
                var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Where(show => show.Size < size * .8d && show.Text.Length > 0 && show.Text.All(char.IsAsciiDigit))
                    .Select(show => show.Text).ToArray();
                TestAssert.Equal(string.Join(',', expected.Concat(expected)), string.Join(',', actual));
                TestAssert.True(links.All(link => link.Destination is { } destination && destination.PageIndex == index &&
                    link.Width > 0d && link.Height > 0d), "Every source marker retains its own complete same-page destination.");
            }
            // The independent Word data establishes eachPage counter semantics.
            // Portable fonts have their own line widths; only the native
            // capacity observer above compares independent Word pagination.
            TestAssert.Equal("decimalZero", row.GetProperty("NumberFormat").GetString());
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            foreach (string content in new[] { "105", "108", "Alpha", "Bravo", "Charlie", "Delta" }) TestAssert.Contains(content, text);
            TestAssert.Equal("105,106,107,108", string.Join(',', DocxBlockTraversal.EnumerateBodyParagraphs(source)
                .SelectMany(item => item.InlineReferences).Select(reference => reference.DisplayText)));
        }
    }
}
