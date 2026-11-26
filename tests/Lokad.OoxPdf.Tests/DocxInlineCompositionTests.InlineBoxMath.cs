using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineBoxMathKeepsIndependentContentAcrossContainers()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native inline-box content uses installed Windows fonts.");
        CheckInlineBoxMath(null);
    }

    public static void DocxInlineBoxMathPreparationRetainsChildAndSiblingSourceIdentity()
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
            var actual = DocxTests.EnumerateInlineBoxMathCandidates(source, default)
                .SelectMany(paragraph => DocxMathParagraphPolicy.GetPreparationEquations(paragraph)).ToArray();
            TestAssert.Equal(row.GetProperty("BoxEquationCount").GetInt32(), actual.Length);
            TestAssert.Equal(expected.Count, actual.Length);
            TestAssert.True(actual.All(expected.Contains),
                "Prepared math retains actual child owners across body, cell and static-story frames.");
            foreach (var parent in DocxBlockTraversal.EnumerateBodyParagraphs(source)
                .Where(paragraph => paragraph.InlineTextBoxes.Count != 0 && paragraph.Equations.Count != 0))
                TestAssert.Equal(parent.Equations.Count, DocxMathParagraphPolicy.GetPreparationEquations(parent).Count);
        }
    }

    public static void DocxInlineBoxMathPreparationObservesCancellationBetweenChildren()
    {
        var source = ReadMutation("InlineBoxMath/inline-box-mixed-source-report", OoxPdfDocxMarkupMode.Final, false, false);
        using var cancellation = new CancellationTokenSource();
        using var candidates = DocxTests.EnumerateInlineBoxMathCandidates(source, cancellation.Token).GetEnumerator();
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

        var mixed = ReadMutation("InlineBoxMath/inline-box-mixed-source-report", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, _) =>
            {
                XElement outer = xml.Descendants(w + "drawing").Single(drawing => drawing.Descendants(w + "txbxContent").Any());
                outer.AddBeforeSelf(new XElement(outer.Descendants(w + "drawing").Single()));
            });
        var sharedRunOwner = DocxBlockTraversal.EnumerateBodyParagraphs(mixed).Single(p => p.InlineTextBoxes.Count != 0);
        var picture = sharedRunOwner.Images.Single();
        var box = sharedRunOwner.InlineTextBoxes.Single();
        TestAssert.Equal(picture.SourceRunIndex, box.SourceRunIndex);
        TestAssert.Equal(picture.SourceTextOffsetInRun, box.SourceTextOffsetInRun);
        TestAssert.Equal(picture.SourceOrderInRun + 1, box.SourceOrderInRun);
        TestAssert.Equal(1, DocxBlockTraversal.EnumerateBodyParagraphs(box.BodyElements).Sum(p => p.Images.Count));
    }

    [Flags]
    private enum InlineBoxChecks
    {
        Content = 1,
        MathOrigins = 2,
        ProseOrigins = 4,
        FollowingOrigins = 8,
        All = Content | MathOrigins | ProseOrigins | FollowingOrigins
    }

    internal static void CheckInlineBoxMath(IFontResolver? resolver, string goldenFile = "inline-box-math-office.json",
        bool compareFieldOrigins = false) =>
        CheckInlineBoxReference(goldenFile, InlineBoxChecks.Content, resolver, compareFieldOrigins);

    private static void CheckInlineBoxReference(string goldenFile, InlineBoxChecks checks, IFontResolver? resolver = null,
        bool compareFieldOrigins = false)
    {
        bool checkContent = checks.HasFlag(InlineBoxChecks.Content);
        bool checkMathOrigins = checks.HasFlag(InlineBoxChecks.MathOrigins);
        using var gold = ReadInlineBoxMathGold(goldenFile);
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            bool checkProseOrigins = checks.HasFlag(InlineBoxChecks.ProseOrigins) && row.TryGetProperty("WordProseAnchors", out _);
            bool checkFollowingOrigins = checks.HasFlag(InlineBoxChecks.FollowingOrigins) && row.TryGetProperty("WordFollowingAnchor", out _);
            if (!checkContent && !checkMathOrigins && !checkProseOrigins && !checkFollowingOrigins) continue;

            string id = row.GetProperty("Id").GetString()!;
            string modeName = row.GetProperty("Mode").GetString()!;
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath", id + ".docx"));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = modeName == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var geometry = final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;

            // All baseline assertions read the same pinned package and page data.
            Check(DocxTests.ReadDocx(path, mode), transformed: false);
            if (checkContent && row.GetProperty("WordMathText").GetString()!.Length != 0)
            {
                var transformed = ReadMutation("InlineBoxMath/" + id, mode, bookmark: true, split: true,
                    partOverride: row.GetProperty("MathPart").GetString());
                Check(transformed, transformed: true);
            }

            void Check(DocxDocument document, bool transformed)
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, geometry).RenderPagesAsync(document, diagnostics.Add, default)
                    .ToBlockingEnumerable().ToArray();
                var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Select(show => new InlineBoxShow(page, index + 1, show.Font, show.Text, show.Size, show.X, show.Y))).ToArray();
                var math = checkContent || checkMathOrigins
                    ? shows.Where(show => show.Page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                        ReferenceEquals(font.Font, show.Font))).ToArray()
                    : Array.Empty<InlineBoxShow>();
                string label = id + "/" + modeName;
                if (checkContent)
                {
                    CheckInlineBoxContent(row, label, pages, shows, math, diagnostics, compareFieldOrigins);
                }

                if (transformed) return;
                if (checkMathOrigins) CheckInlineBoxMathOrigins(row, label, math);
                if (checkProseOrigins || checkFollowingOrigins)
                    CheckInlineBoxTextOrigins(row, label, shows, checkProseOrigins, checkFollowingOrigins);
            }
        }
    }

    private readonly record struct InlineBoxShow(PdfPage Page, int PageNumber, PdfEmbeddedFont Font,
        string Text, double Size, double X, double Y);

    private static void CheckInlineBoxContent(JsonElement row, string label, PdfPage[] pages,
        InlineBoxShow[] shows, InlineBoxShow[] mathShows, List<OoxPdfDiagnostic> diagnostics, bool compareFieldOrigins)
    {
        var math = mathShows.Where(show => !string.IsNullOrWhiteSpace(show.Text)).ToArray();
        TestAssert.Equal(row.GetProperty("ExpectedPaintedMathText").GetString(),
            string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
        TestAssert.Equal(row.GetProperty("ExpectedPaintedMathGlyphs").GetInt32(), math.Length);
        if (row.TryGetProperty("WordFieldGlyphs", out var fieldGlyphs))
        {
            foreach (var anchor in fieldGlyphs.EnumerateArray())
            {
                string value = anchor.GetProperty("Text").GetString()!;
                var field = shows.Single(show => show.Text.Trim() == value);
                TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), field.PageNumber);
                if (compareFieldOrigins)
                    TestAssert.True(Math.Abs(field.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(field.Page.Height - field.Y - anchor.GetProperty("Y").GetDouble()) < 1d &&
                        Math.Abs(field.Size - anchor.GetProperty("Size").GetDouble()) < .1d,
                        label + ": source field glyph retains its independent Word origin and size.");
            }
        }

        TestAssert.True(math.All(show => show.X >= 0d && show.X <= show.Page.Width && show.Y > 0d && show.Y < show.Page.Height),
            label + ": prepared glyph origins must remain inside the printed page, including footer boxes.");
        var mathFonts = math.Select(show => show.Font).ToHashSet();
        string prose = Regex.Replace(string.Join(" ", shows.Where(show => !mathFonts.Contains(show.Font))
            .GroupBy(show => (show.PageNumber, show.Y))
            .Select(line => string.Concat(line.OrderBy(show => show.X).Select(show => show.Text)))), @"\s+", " ").Trim();
        foreach (var text in row.GetProperty("ExpectedText").EnumerateArray())
            TestAssert.True(prose.Contains(text.GetString()!, StringComparison.Ordinal),
                label + ": source prose missing: " + text.GetString());
        TestAssert.Equal(row.GetProperty("WordImageCount").GetInt32(), pages.Sum(page => page.Images.Count));
        var ignored = diagnostics.Where(item => item.Id == "DOCX_UNSUPPORTED_EQUATION" && item.Fallback == "Ignored").ToArray();
        TestAssert.Equal(row.GetProperty("IgnoredEquationDiagnostics").GetInt32(), ignored.Length);
        TestAssert.True(ignored.Select(item => item.PartName).OrderBy(part => part, StringComparer.Ordinal).SequenceEqual(
            row.GetProperty("IgnoredEquationParts").EnumerateArray().Select(part => part.GetString()).OrderBy(part => part, StringComparer.Ordinal)),
            label + ": preserve staged sibling or unused-part diagnostics without reporting painted child math as ignored.");
        if (math.Length != 0)
            TestAssert.True(diagnostics.Any(item => item.Id == "DOCX_APPROXIMATED_EQUATION"),
                label + ": painted child equations report approximation instead of an ignored source.");
    }

    private static void CheckInlineBoxMathOrigins(JsonElement row, string label, InlineBoxShow[] math)
    {
        var expected = row.GetProperty("WordMathAnchors").EnumerateArray().ToArray();
        TestAssert.Equal(expected.Length, math.Length);
        for (int index = 0; index < expected.Length; index++)
        {
            var actual = math[index];
            var anchor = expected[index];
            TestAssert.Equal(anchor.GetProperty("Text").GetString(), actual.Text.Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.PageNumber);
            foreach (var axis in row.GetProperty("GeometryAxes").EnumerateArray().Select(axis => axis.GetString()))
            {
                double position = axis == "X" ? actual.X : actual.Page.Height - actual.Y;
                TestAssert.True(Math.Abs(position - anchor.GetProperty(axis!).GetDouble()) < 1d,
                    label + ": child glyph " + axis + " retains the one-point independent Word gate.");
            }
        }
    }

    private static void CheckInlineBoxTextOrigins(JsonElement row, string label, InlineBoxShow[] shows,
        bool checkProseOrigins, bool checkFollowingOrigins)
    {
        var lines = shows.GroupBy(show => (show.PageNumber, show.Y)).Select(line =>
            (Page: line.Key.PageNumber, Height: line.First().Page.Height, Y: line.Key.Y, X: line.Min(show => show.X),
            Text: string.Concat(line.OrderBy(show => show.X).Select(show => show.Text)))).ToArray();

        if (checkProseOrigins)
        {
            var normalizedLines = lines.Select(line => (Line: line, Text: line.Text.Normalize(NormalizationForm.FormKC).Trim())).ToArray();
            foreach (var anchor in row.GetProperty("WordProseAnchors").EnumerateArray())
            {
                var actual = normalizedLines.Single(line => line.Text.StartsWith(anchor.GetProperty("Text").GetString()!, StringComparison.Ordinal)).Line;
                TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
                foreach (var axis in row.GetProperty("ProseGeometryAxes").EnumerateArray().Select(axis => axis.GetString()))
                {
                    double position = axis == "X" ? actual.X : actual.Height - actual.Y;
                    TestAssert.True(Math.Abs(position - anchor.GetProperty(axis!).GetDouble()) < 1d,
                        label + ": ordinary child " + axis + " retains the one-point independent Word gate.");
                }
            }
        }

        if (checkFollowingOrigins)
        {
            var expected = row.GetProperty("WordFollowingAnchor");
            string text = row.TryGetProperty("FollowingText", out var value) ? value.GetString()! : "After public math table.";
            var following = lines.Single(line => line.Text.Trim() == text);
            TestAssert.Equal(expected.GetProperty("Page").GetInt32(), following.Page);
            TestAssert.True(Math.Abs(following.X - expected.GetProperty("X").GetDouble()) < 1d &&
                Math.Abs(following.Height - following.Y - expected.GetProperty("Y").GetDouble()) < 1d,
                label + ": following prose retains one-point independent Word origins.");
        }
    }

    private static JsonDocument ReadInlineBoxMathGold(string goldenFile = "inline-box-math-office.json") => JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath", goldenFile))));
}
