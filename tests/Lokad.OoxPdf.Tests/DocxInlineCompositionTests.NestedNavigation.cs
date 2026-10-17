using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNestedNavigationKeepsIndependentContentAndViewportContext()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent navigation controls use installed Windows fonts.");
        CheckNestedNavigation(null, independentGeometry: true);
    }

    public static void DocxInlineCompositionNestedNavigationKeepsOriginalSourceAndCellOwner()
    {
        foreach (string direction in new[] { "within", "incoming", "outgoing" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/nested-math-navigation-" + direction, mode, false, false);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).Single(paragraph => paragraph.Equations.Count != 0);
            TestAssert.True(ReferenceEquals(paragraph, DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, default)
                .Single(paragraph => paragraph.Equations.Count != 0)),
                "Hyperlinks and bookmark targets retain the same nested mathematical source owner.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.Equal(1, DocxBlockTraversal.EnumerateBodyParagraphs(document).Sum(source => source.Hyperlinks.Count));
            TestAssert.Equal(1, DocxBlockTraversal.EnumerateBodyParagraphs(document).Sum(source => source.BookmarkAnchors
                .Count(bookmark => bookmark.Name == "PublicNestedTarget")));
        }
    }

    internal static void CheckNestedNavigation(IFontResolver? resolver, bool independentGeometry = false)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("nested-math-navigation-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            string direction = row.GetProperty("Direction").GetString()!;
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var baseline = renderer.RenderBlankPages(ReadMutation("CompositionProbes/mixed-nested-cell-inside", mode, false, false), null, default).Single();
            foreach (bool transformed in new[] { false, true })
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = renderer.RenderBlankPages(ReadMutation("CompositionProbes/" + id, mode, transformed, transformed), diagnostics.Add, default).Single();
                string text = DocxPageFootnoteNumberingTests.DecodeText(page).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
                foreach (string prose in new[] { "Public outer head", "Public outer tail", "Following prose." }) TestAssert.Contains(prose, text);
                TestAssert.Equal(1, page.Images.Sum(image => Regex.Matches(page.Content,
                    "/" + Regex.Escape(image.ResourceName) + @"\s+Do\b").Count));
                TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Navigation cannot suppress supported nested mathematical content.");
                var annotation = page.Annotations.Single();
                TestAssert.True(annotation.Destination is { PageIndex: 0, Left: not null, Top: not null } &&
                    annotation.Width > 0d && annotation.Height > 0d,
                    "Each source link retains a finite placed rectangle and its concrete target.");
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                var math = shows.Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                    ReferenceEquals(font.Font, show.Font))).ToArray();
                TestAssert.Equal(16, math.Length);
                if (direction != "incoming")
                    TestAssert.True(math.Take(8).All(show => annotation.X <= show.X + .01d &&
                        annotation.X + annotation.Width >= show.X - .01d && annotation.Y <= show.Y + .01d &&
                        annotation.Y + annotation.Height >= show.Y - .01d),
                        "The nested formula link includes both painted numerator and denominator origins.");
                if (!transformed)
                    TestAssert.Equal(baseline.Content, page.Content);
                if (!independentGeometry || transformed || direction != "outgoing") continue;
                var target = shows.Single(show => show.Text.Trim() == "Public outer tail");
                double context = target.X - annotation.Destination!.Value.Left!.Value;
                double expected = row.GetProperty("TargetTextX").GetDouble() - row.GetProperty("DestinationX").GetDouble();
                TestAssert.True(Math.Abs(context - expected) < .5d,
                    "The body-cell viewport uses Word's independently measured source-relative context.");
                if (mode == OoxPdfDocxMarkupMode.Final)
                    TestAssert.True(Math.Abs(annotation.Destination.Value.Left.Value - row.GetProperty("DestinationX").GetDouble()) < .5d,
                        "Final body-cell destination X matches the independent Word control.");
            }
        }
    }
}
