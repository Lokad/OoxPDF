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
    public static void DocxInlineBoxMathKeepsIndependentContentAcrossContainers()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native inline-box content uses installed Windows fonts.");
        CheckInlineBoxMath(null);
    }

    public static void DocxInlineBoxMathPreparationRetainsChildSourceIdentityAndSiblingFallback()
    {
        using var gold = ReadInlineBoxMathGold();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray()
            .Where(row => row.GetProperty("Mode").GetString() == "final-preserve"))
        {
            var source = ReadMutation("InlineBoxMath/" + row.GetProperty("Id").GetString(),
                OoxPdfDocxMarkupMode.Final, false, false);
            var expected = new HashSet<DocxEquation>(DocxBlockTraversal.EnumerateSourceParagraphs(source)
                .SelectMany(paragraph => paragraph.InlineTextBoxes)
                .SelectMany(box => DocxBlockTraversal.EnumerateBodyParagraphs(box.BodyElements))
                .SelectMany(paragraph => paragraph.Equations), ReferenceEqualityComparer.Instance);
            var actual = DocxMathParagraphPolicy.EnumerateInlineBoxPreparationCandidates(source, default)
                .SelectMany(paragraph => DocxMathParagraphPolicy.GetPreparationEquations(paragraph)).ToArray();
            TestAssert.Equal(row.GetProperty("BoxEquationCount").GetInt32(), actual.Length);
            TestAssert.Equal(expected.Count, actual.Length);
            TestAssert.True(actual.All(expected.Contains),
                "Prepared math retains actual child owners across body, cell and static-story frames.");
            foreach (var parent in DocxBlockTraversal.EnumerateBodyParagraphs(source)
                .Where(paragraph => paragraph.InlineTextBoxes.Count != 0 && paragraph.Equations.Count != 0))
                TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(parent).Count);
        }
    }

    public static void DocxInlineBoxMathPreparationObservesCancellationBetweenChildren()
    {
        var source = ReadMutation("InlineBoxMath/inline-box-mixed-source-report", OoxPdfDocxMarkupMode.Final, false, false);
        using var cancellation = new CancellationTokenSource();
        using var candidates = DocxMathParagraphPolicy.EnumerateInlineBoxPreparationCandidates(source, cancellation.Token).GetEnumerator();
        TestAssert.True(candidates.MoveNext(), "The mixed frame has multiple source children.");
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => candidates.MoveNext());
    }

    public static void DocxInlineBoxMathReaderKeepsPicturesWithTheirChildSourceFrame()
    {
        var source = ReadMutation("InlineBoxMath/inline-box-mixed-source-report", OoxPdfDocxMarkupMode.Final, false, false);
        var owner = DocxBlockTraversal.EnumerateBodyParagraphs(source).Single(paragraph => paragraph.InlineTextBoxes.Count != 0);
        TestAssert.Equal(0, owner.Images.Count);
        var children = DocxBlockTraversal.EnumerateBodyParagraphs(owner.InlineTextBoxes.Single().BodyElements).ToArray();
        TestAssert.Equal(3, children.Sum(paragraph => paragraph.Equations.Count));
        TestAssert.Equal(1, children.Sum(paragraph => paragraph.Images.Count));
        TestAssert.True(children.SelectMany(paragraph => paragraph.Images).All(image => image.SourceRunIndex >= 0),
            "An image retains its child run position instead of being duplicated at the outer box's extent.");
    }

    internal static void CheckInlineBoxMath(IFontResolver? resolver, string goldenFile = "inline-box-math-office.json")
    {
        using var gold = ReadInlineBoxMathGold(goldenFile);
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath", id + ".docx"));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            bool hasMath = row.GetProperty("WordMathText").GetString()!.Length != 0;
            foreach (bool transformed in hasMath ? new[] { false, true } : new[] { false })
            {
                var document = ReadMutation("InlineBoxMath/" + id, mode, transformed, transformed,
                    partOverride: row.GetProperty("MathPart").GetString());
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).ToArray();
                var shows = pages.SelectMany((page, pageIndex) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                    usePageCoordinates: true).Select(show => (Page: page, PageIndex: pageIndex, Show: show))).ToArray();
                var math = shows.Where(item => !string.IsNullOrWhiteSpace(item.Show.Text) &&
                    item.Page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                    ReferenceEquals(font.Font, item.Show.Font))).ToArray();
                TestAssert.Equal(row.GetProperty("ExpectedPaintedMathText").GetString(),
                    string.Concat(math.Select(item => item.Show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(row.GetProperty("ExpectedPaintedMathGlyphs").GetInt32(), math.Length);
                TestAssert.True(math.All(item => item.Show.X >= 0d && item.Show.X <= item.Page.Width &&
                    item.Show.Y > 0d && item.Show.Y < item.Page.Height),
                    id + ": prepared glyph origins must remain inside the printed page, including footer boxes.");
                var mathFonts = math.Select(item => item.Show.Font).ToHashSet();
                string prose = Regex.Replace(string.Join(" ", shows.Where(item => !mathFonts.Contains(item.Show.Font))
                    .GroupBy(item => (item.PageIndex, item.Show.Y))
                    .Select(line => string.Concat(line.OrderBy(item => item.Show.X).Select(item => item.Show.Text)))), @"\s+", " ").Trim();
                foreach (var text in row.GetProperty("ExpectedText").EnumerateArray())
                    TestAssert.True(prose.Contains(text.GetString()!, StringComparison.Ordinal),
                        id + "/" + row.GetProperty("Mode").GetString() + ": source prose missing: " + text.GetString());
                TestAssert.Equal(row.GetProperty("WordImageCount").GetInt32(), pages.Sum(page => page.Images.Count));
                TestAssert.Equal(row.GetProperty("IgnoredEquationDiagnostics").GetInt32(),
                    diagnostics.Count(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION" && diagnostic.Fallback == "Ignored"));
                TestAssert.True(diagnostics.Where(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION" && diagnostic.Fallback == "Ignored")
                    .Select(diagnostic => diagnostic.PartName).OrderBy(part => part, StringComparer.Ordinal).SequenceEqual(
                        row.GetProperty("IgnoredEquationParts").EnumerateArray().Select(part => part.GetString()).OrderBy(part => part, StringComparer.Ordinal)),
                    id + ": preserve staged sibling or unused-part diagnostics without reporting painted child math as ignored.");
                if (math.Length != 0)
                    TestAssert.True(diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_APPROXIMATED_EQUATION"),
                        "Painted child equations report approximation instead of an ignored source.");
            }
        }
    }

    private static JsonDocument ReadInlineBoxMathGold(string goldenFile = "inline-box-math-office.json") => JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath", goldenFile))));
}
