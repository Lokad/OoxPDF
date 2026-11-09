using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxFixedStaticFramesMatchIndependentWordOriginsAndBoxFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent fixed static origins use installed Windows fonts.");
        using var cases = ReadFixedStaticFramesJson("fixed-static.json");
        using var gold = ReadFixedStaticFramesJson("fixed-static-office.json");
        foreach (var expected in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var source = cases.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == expected.GetProperty("Id").GetString());
            TestAssert.Equal(source.GetProperty("Sha256").GetString(), expected.GetProperty("InputSha256").GetString());
            var mode = expected.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var page = RenderContainerGrid(source, mode, null);
            string context = source.GetProperty("Id").GetString() + "/" + mode;
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            foreach (var origin in expected.GetProperty("TextOrigins").EnumerateArray())
            {
                var actual = shows.Single(show => show.Text.Trim() == origin.GetProperty("Text").GetString());
                TestAssert.True(Math.Abs(actual.X - origin.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(page.Height - actual.Y - origin.GetProperty("Y").GetDouble()) < 1d,
                    "Declared line flow retains the independent static origin: " + context);
            }
            var math = shows.Where(show => show.Text.Trim().Length != 0 && page.Fonts.Any(font =>
                font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))).ToArray();
            var native = expected.GetProperty("Math").EnumerateArray().ToArray();
            TestAssert.Equal(native.Length, math.Length);
            for (int i = 0; i < math.Length; i++)
            {
                TestAssert.Equal(native[i].GetProperty("Text").GetString(), math[i].Text.Normalize(NormalizationForm.FormKC));
                TestAssert.True(Math.Abs(math[i].X - native[i].GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(page.Height - math[i].Y - native[i].GetProperty("Y").GetDouble()) < 1d,
                    "Mathematics retains its independent placed frame: " + context);
            }
            var box = expected.GetProperty("Box");
            if (box.GetArrayLength() == 0) continue;
            var rectangle = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re\s+S")
                .Select(match => Enumerable.Range(1, 4).Select(index => double.Parse(match.Groups[index].Value, CultureInfo.InvariantCulture)).ToArray())
                .Single(rect => rect[2] > 100d && rect[3] > 30d);
            double[] edges = [rectangle[0], page.Height - rectangle[1] - rectangle[3], rectangle[0] + rectangle[2], page.Height - rectangle[1]];
            for (int i = 0; i < edges.Length; i++)
                TestAssert.True(Math.Abs(edges[i] - box[i].GetDouble()) < 1d, "The declared child cannot move its container frame: " + context);
        }
    }

    public static void DocxFixedStaticFramesPreservePortableContentNavigationAndCancellation()
    {
        using var cases = ReadFixedStaticFramesJson("fixed-static.json");
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var source in cases.RootElement.EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var page = RenderContainerGrid(source, mode, resolver);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            string text = string.Concat(shows.Select(show => show.Text));
            foreach (var label in source.GetProperty("ExpectedText").EnumerateArray())
            {
                TestAssert.True(text.Contains(label.GetString()!, StringComparison.Ordinal), "Supported content survives its declared static frame: " + source.GetProperty("Id").GetString());
                var actual = shows.Single(show => show.Text.Trim() == label.GetString());
                TestAssert.True(actual.Y >= 0d && actual.Y <= page.Height,
                    "The declared frame must keep its text on the physical page: " + source.GetProperty("Id").GetString() + "/" + mode);
            }
            if (source.GetProperty("ExpectedEquationCount").GetInt32() != 0)
                CheckStaticBoxTablePaint(page, source.GetProperty("Id").GetString() + "/" + mode);
        }
        var footer = cases.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == "footer-terminal-text12");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => RenderContainerGrid(footer, OoxPdfDocxMarkupMode.AllMarkup, resolver, canceled.Token));
    }

    private static JsonDocument ReadFixedStaticFramesJson(string name) => JsonDocument.Parse(File.ReadAllText(Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "StaticDeclaredFrames", name))));
}
