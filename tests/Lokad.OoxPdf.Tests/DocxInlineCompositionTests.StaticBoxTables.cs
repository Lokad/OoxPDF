using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxStaticBoxTablesKeepPaintAndNavigationInsideTheirContainer()
    {
        using var cases = ReadStaticBoxTableCases();
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var row in cases.RootElement.EnumerateArray())
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
            CheckStaticBoxTablePaint(RenderStaticBoxTable(row, mode, resolver), row.GetProperty("Id").GetString() + "/" + mode);
    }

    public static void DocxStaticBoxTablesKeepIndependentWordVerticalOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent static-box table origins use installed Windows fonts.");
        using var cases = ReadStaticBoxTableCases();
        using var golden = JsonDocument.Parse(File.ReadAllText(StaticBoxTablePath("static-box-tables-office.json")));
        foreach (var gold in golden.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var row = cases.RootElement.EnumerateArray().Single(item =>
                item.GetProperty("Id").GetString() == gold.GetProperty("Id").GetString());
            TestAssert.Equal(row.GetProperty("Sha256").GetString(), gold.GetProperty("InputSha256").GetString());
            var mode = gold.GetProperty("Mode").GetString() == "final-preserve"
                ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var page = RenderStaticBoxTable(row, mode, null);
            CheckStaticBoxTablePaint(page, row.GetProperty("Id").GetString() + "/" + mode);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            foreach (var anchor in gold.GetProperty("WordTextOrigins").EnumerateArray())
            {
                var actual = shows.Single(show => show.Text.Trim() == anchor.GetProperty("Text").GetString());
                TestAssert.True(Math.Abs(page.Height - actual.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                    "The complete static table retains its independent Word vertical origin.");
            }
            var math = shows.Where(show => show.Text.Trim().Length != 0 && page.Fonts.Any(font =>
                font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))).ToArray();
            var native = gold.GetProperty("WordMath").EnumerateArray().ToArray();
            TestAssert.Equal(native.Length, math.Length);
            for (int i = 0; i < math.Length; i++)
            {
                TestAssert.Equal(native[i].GetProperty("Text").GetString(), math[i].Text.Normalize(NormalizationForm.FormKC));
                TestAssert.True(Math.Abs(page.Height - math[i].Y - native[i].GetProperty("Y").GetDouble()) < 1d,
                    "Mathematical children share the same independently measured vertical frame.");
            }
        }
    }

    public static void DocxStaticBoxTablesPreserveSourcePreparationAndCancellation()
    {
        using var cases = ReadStaticBoxTableCases();
        foreach (var row in cases.RootElement.EnumerateArray())
        {
            var document = ReadStaticBoxTableSource(row, OoxPdfDocxMarkupMode.Final);
            var expected = new HashSet<DocxEquation>(DocxBlockTraversal.EnumerateSourceParagraphs(document)
                .SelectMany(paragraph => paragraph.InlineTextBoxes)
                .SelectMany(box => DocxBlockTraversal.EnumerateBodyParagraphs(box.BodyElements))
                .SelectMany(paragraph => paragraph.Equations), ReferenceEqualityComparer.Instance);
            var actual = DocxMathParagraphPolicy.EnumerateInlineBoxPreparationCandidates(document, default)
                .SelectMany(paragraph => DocxMathParagraphPolicy.GetPreparationEquations(paragraph)).ToArray();
            TestAssert.Equal(1, expected.Count);
            TestAssert.Equal(1, actual.Length);
            TestAssert.True(ReferenceEquals(expected.Single(), actual[0]), "Table preparation retains its actual boxed equation owner.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
                    OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(document, null, canceled.Token).ToBlockingEnumerable().ToArray());
        }
    }

    internal static void CheckContainedTablePrintedBaselines()
    {
        using var cases = ReadStaticBoxTableCases();
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        double Gap(string id)
        {
            var source = cases.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == id);
            var page = RenderStaticBoxTable(source, OoxPdfDocxMarkupMode.AllMarkup, resolver);
            CheckStaticBoxTablePaint(page, id);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            return shows.Single(show => show.Text.Trim() == "BOX TABLE NORTH").Y -
                shows.Single(show => show.Text.Trim() == "BOX TABLE SOUTH").Y;
        }
        TestAssert.True(Math.Abs(Gap("static-box-table-body") - Gap("static-box-table-header")) < .05d,
            "The same plain/math rows retain their printed baseline spacing across body and header box placement.");
    }

    private static PdfPage RenderStaticBoxTable(JsonElement row, OoxPdfDocxMarkupMode mode, IFontResolver? resolver)
    {
        var geometry = mode == OoxPdfDocxMarkupMode.AllMarkup
            ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
        return new DocxRenderer(resolver, mode, geometry)
            .RenderPagesAsync(ReadStaticBoxTableSource(row, mode), null, default).ToBlockingEnumerable().Single();
    }

    private static void CheckStaticBoxTablePaint(PdfPage page, string context)
    {
        var outlines = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re\s+S")
            .Select(match => Enumerable.Range(1, 4).Select(i => double.Parse(match.Groups[i].Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(rect => rect[2] > 100d && rect[3] > 30d).ToArray();
        TestAssert.Equal(1, outlines.Length);
        var box = outlines[0];
        bool Inside(double x, double y) => x >= box[0] - .01d && x <= box[0] + box[2] + .01d &&
            y >= box[1] - .01d && y <= box[1] + box[3] + .01d;
        var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
        foreach (string label in new[] { "BOX TABLE NORTH", "BOX TABLE SOUTH" })
        {
            var show = shows.Single(item => item.Text.Trim() == label);
            TestAssert.True(Inside(show.X, show.Y), "Table text must paint inside its containing box clip: " + label);
        }
        var math = shows.Where(show => show.Text.Trim().Length != 0 && page.Fonts.Any(font =>
            font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))).ToArray();
        TestAssert.Equal("x+2y", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
        TestAssert.True(math.All(show => Inside(show.X, show.Y)), "All mathematical glyph origins remain inside the box clip.");
        var link = page.Annotations.Single(annotation => annotation.Uri == "https://example.invalid/static-box-table");
        TestAssert.True(Inside(link.X, link.Y) && Inside(link.X + link.Width, link.Y + link.Height),
            "The source link follows the painted table rather than the body emission origin.");
        var destination = page.Annotations.Single(annotation => annotation.IsDestination).Destination!.Value;
        var target = shows.Single(show => show.Text.Trim() == "BOX TABLE NORTH");
        // A viewport can include ascender/context space above the box. Its top
        // must reveal the painted target, rather than hide it below the viewport.
        TestAssert.True(destination.PageIndex == 0 && destination.Left is double x &&
            x >= box[0] && x <= box[0] + box[2] && destination.Top is double y && y >= target.Y && y <= page.Height,
            "The table bookmark viewport must reveal the painted target: " + context +
            "; destination=" + destination + "; box=" + string.Join(",", box));
    }

    private static DocxDocument ReadStaticBoxTableSource(JsonElement row, OoxPdfDocxMarkupMode mode)
    {
        string id = row.GetProperty("Id").GetString()!;
        TestAssert.Equal(row.GetProperty("Sha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(StaticBoxTablePath(id + ".docx")))));
        return ReadMutation("InlineBoxMath/" + id, mode, false, false, partOverride: row.GetProperty("MathPart").GetString());
    }

    private static JsonDocument ReadStaticBoxTableCases() => JsonDocument.Parse(File.ReadAllText(StaticBoxTablePath("static-box-tables.json")));
    private static string StaticBoxTablePath(string name) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "InlineBoxMath", name));
}
