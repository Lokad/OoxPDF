using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionEmptyFractionRevisionsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent glyphless fraction rules use installed Windows fonts.");
        CheckEquationRevisions(null, true, "empty-fraction-revisions-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionEmptyFractionRevisionsKeepActualNeutralOwners()
    {
        foreach (string shape in new[] { "whole", "fragment" })
        foreach (string kind in new[] { "ins", "del" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/empty-fraction-revision-" + shape + "-" + kind, mode, false, false);
            var paragraph = source.Paragraphs.First();
            var equation = paragraph.Equations.Single();
            TestAssert.True(DocxMathLayout.IsEmptyFraction(equation) && DocxMathLayout.TryGetRuns(equation, default, out var runs) && runs.Count == 0 &&
                ReferenceEquals(DocxMathLayout.GetStandaloneControl(equation, default), equation.Content.Children.Single()) &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default) &&
                DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph),
                "A glyphless fraction keeps its actual structural owner without inventing revised literal ink.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.TryGetEquationRevisionOwners(equation, canceled.Token, out _));
        }
    }

    public static void DocxInlineCompositionEmptyFractionRevisionsKeepMetricOnlyPublication()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native metric-only fractions use installed Windows fonts.");
        CheckEmptyFractionRevisionPublication(null);
    }

    internal static void CheckEmptyFractionRevisionPublication(IFontResolver? resolver)
    {
        var source = ReadMutation("CompositionProbes/empty-fraction-revision-whole-ins", OoxPdfDocxMarkupMode.AllMarkup,
            false, false, (xml, w, m) =>
            {
                var body = xml.Root!.Element(w + "body")!;
                body.ReplaceNodes(new XElement(body.Element(w + "p")!), new XElement(body.Element(w + "sectPr")!));
            });
        var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderBlankPages(source, null, default).Single();
        TestAssert.True(page.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
            "The bar uses metric-only preparation without a dummy glyph or mathematical PDF subset.");
        TestAssert.True(Regex.IsMatch(page.Content, @"(?m)^0 g\r?\n[-\d.]+ [-\d.]+ [\d.]+ [\d.]+ re f\r?$"),
            "The owned bar remains authored black without mathematical text shows.");
    }

    private static void CheckPinnedMathRuleBoxes(PdfPage page, JsonElement row, bool compareGeometry)
    {
        const string rectangle = @"(?<x>[-\d.]+) (?<y>[-\d.]+) (?<width>[\d.]+) (?<height>[\d.]+) re f\r?$";
        var rectangles = Regex.Matches(page.Content, @"(?m)^" + rectangle).Cast<Match>().ToArray();
        var rules = rectangles.Where(IsThinRule).ToArray();
        int total = row.GetProperty("TotalThinRules").GetInt32();
        var blackRules = Regex.Matches(page.Content, @"(?m)^0 g\r?\n" + rectangle).Cast<Match>().Where(IsThinRule).ToArray();
        int black = row.TryGetProperty("BlackThinRules", out var blackCount) ? blackCount.GetInt32() : total;
        TestAssert.Equal(black, blackRules.Length);
        if (total == black) TestAssert.Equal(total, rules.Length);
        else
            // Word can overlap a run-wide strike with an operator strike;
            // the builder emits one strike per glyph. Count structural rules
            // exactly, and retain revision ink without requiring PDF segmentation.
            TestAssert.True(rules.Length > blackRules.Length, "Visible revised literals retain strike ink.");
        TestAssert.Equal(row.GetProperty("RevisionBarCount").GetInt32(), rectangles.Count(rule =>
            Number(rule, "x") is > 20d and < 40d && Number(rule, "width") is > 0d and < 2d && Number(rule, "height") > 5d));
        if (!compareGeometry) return;
        var physicalRules = blackRules.Select(rule => ResolveGeneratedRuleBounds(page.Content, rule)).ToArray();
        foreach (var expected in row.GetProperty("RuleBoxes").EnumerateArray())
        {
            TestAssert.Equal("g:0", expected.GetProperty("FillColor").GetString());
            var rule = physicalRules.OrderBy(rule => Math.Abs(rule.X - expected.GetProperty("X").GetDouble()) +
                Math.Abs(rule.Y - expected.GetProperty("Y").GetDouble())).First();
            TestAssert.True(Math.Abs(rule.X - expected.GetProperty("X").GetDouble()) < 1d &&
                Math.Abs(rule.Y - expected.GetProperty("Y").GetDouble()) < 1d &&
                Math.Abs(rule.Width - expected.GetProperty("Width").GetDouble()) < .1d &&
                Math.Abs(rule.Height - expected.GetProperty("Height").GetDouble()) < .03d,
                "The glyphless structural rule retains independent Word bounds and thickness.");
        }
        static bool IsThinRule(Match rule) => Number(rule, "width") is > 1d and < 50d && Number(rule, "height") is > 0d and < 2d;
        static double Number(Match rule, string field) => double.Parse(rule.Groups[field].Value, CultureInfo.InvariantCulture);
    }

    private static (double X, double Y, double Width, double Height) ResolveGeneratedRuleBounds(string content, Match rule)
    {
        // Only the builder's q/Q/cm operators: inspect the same physical frame
        // as the independent PDF inspector, rather than raw local rectangle data.
        (double A, double B, double C, double D, double E, double F) matrix = (1, 0, 0, 1, 0, 0);
        var saved = new Stack<(double A, double B, double C, double D, double E, double F)>();
        foreach (Match operation in Regex.Matches(content,
            @"(?m)^(?:(?<save>q)|(?<restore>Q)|(?<matrix>[-+\d.eE]+(?: [-+\d.eE]+){5}) cm)\r?$"))
        {
            if (operation.Index >= rule.Index) break;
            if (operation.Groups["save"].Success) saved.Push(matrix);
            else if (operation.Groups["restore"].Success) matrix = saved.Pop();
            else
            {
                var values = operation.Groups["matrix"].Value.Split(' ').Select(Parse).ToArray();
                matrix = (matrix.A * values[0] + matrix.C * values[1], matrix.B * values[0] + matrix.D * values[1],
                    matrix.A * values[2] + matrix.C * values[3], matrix.B * values[2] + matrix.D * values[3],
                    matrix.A * values[4] + matrix.C * values[5] + matrix.E, matrix.B * values[4] + matrix.D * values[5] + matrix.F);
            }
        }
        double x = Parse(rule.Groups["x"].Value), y = Parse(rule.Groups["y"].Value);
        double width = Parse(rule.Groups["width"].Value), height = Parse(rule.Groups["height"].Value);
        var points = new[] { (X: x, Y: y), (X: x + width, Y: y), (X: x, Y: y + height), (X: x + width, Y: y + height) }
            .Select(point => (X: matrix.A * point.X + matrix.C * point.Y + matrix.E,
                Y: matrix.B * point.X + matrix.D * point.Y + matrix.F)).ToArray();
        double left = points.Min(point => point.X), bottom = points.Min(point => point.Y);
        return (left, bottom, points.Max(point => point.X) - left, points.Max(point => point.Y) - bottom);
        static double Parse(string value) => double.Parse(value, CultureInfo.InvariantCulture);
    }
}
