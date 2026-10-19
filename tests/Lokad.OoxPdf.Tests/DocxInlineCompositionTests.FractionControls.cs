using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionFractionControlsMatchIndependentWordFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent fraction anchors require installed Windows fonts.");
        CheckFractionControlFrames(null);
    }

    internal static void CheckFractionControlFrames(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("fraction-control-frames-office.json")));
        var equivalent = new Dictionary<string, (string Text, double X, double Y, double Size)[]>();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string modeName = row.GetProperty("Mode").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(
                SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = modeName == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            (string Text, double X, double Y, double Size)[]? original = null;
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed);
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, diagnostics.Add, default).Single();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Authored controls retain supported fraction content.");
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                var lines = shows.GroupBy(show => show.Y).Select(group =>
                {
                    var ordered = group.OrderBy(show => show.X).ToArray();
                    return (Text: string.Concat(ordered.Select(show => show.Text)), First: ordered[0]);
                }).ToArray();
                foreach (string expected in new[] { "Before public math table.", "Before inline", "after equation.", "After public math table." })
                    TestAssert.Equal(1, lines.Count(line => line.Text.Contains(expected, StringComparison.Ordinal)));
                var prose = lines.Single(line => line.Text.StartsWith("Before inline", StringComparison.Ordinal)).First;
                var after = lines.Single(line => line.Text.StartsWith("After public", StringComparison.Ordinal)).First;
                var math = shows.Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                    ReferenceEquals(font.Font, show.Font))).Select(show => (show.Text, show.X, show.Y, show.Size)).ToArray();
                TestAssert.Equal(4, math.Length);
                TestAssert.True(math.All(show => show.Y > after.Y), "The complete fraction clears following prose.");
                if (resolver is null)
                {
                    TestAssert.True(Math.Abs(prose.Y - row.GetProperty("CellProseY").GetDouble()) < 1d,
                        id + ": the control owns the independent Word line frame, without rescaling its arguments.");
                    foreach (var pair in math.Zip(row.GetProperty("MathAnchors").EnumerateArray()))
                        TestAssert.True(pair.First.Text == pair.Second.GetProperty("Text").GetString() &&
                            Math.Abs(pair.First.Y - pair.Second.GetProperty("Y").GetDouble()) < 1d &&
                            Math.Abs(pair.First.X - prose.X - pair.Second.GetProperty("RelativeX").GetDouble()) < 1d &&
                            Math.Abs(pair.First.Y - prose.Y - pair.Second.GetProperty("RelativeY").GetDouble()) < 1d &&
                            Math.Abs(pair.First.Size - pair.Second.GetProperty("Size").GetDouble()) < .1d,
                            id + ": child size and control-relative geometry retain independent one-point/.1-point Office thresholds.");
                    string container = row.GetProperty("Container").GetString()!;
                    if (container is "body" or "cell")
                        TestAssert.True(Math.Abs(after.Y - row.GetProperty("AfterY").GetDouble()) < 1d,
                            "Ordinary containers use the measured control frame for following prose. Merged-span flow remains a separate residual.");
                    if (container == "body")
                        foreach (var pair in math.Zip(row.GetProperty("MathAnchors").EnumerateArray()))
                            TestAssert.True(Math.Abs(pair.First.X - pair.Second.GetProperty("X").GetDouble()) < 1d,
                                "Body fraction glyphs retain absolute independent Word X anchors.");
                }
                if (original is not null) SameGeometry(original, math);
                original = math;
            }
            if (id is "merged-span-tall-fraction" or "fraction-control-style-24" or "fraction-control-style-font-24")
            {
                if (equivalent.TryGetValue(modeName, out var direct)) SameGeometry(direct, original!);
                else equivalent.Add(modeName, original!);
            }
        }

        static void SameGeometry((string Text, double X, double Y, double Size)[] expected,
            (string Text, double X, double Y, double Size)[] actual)
        {
            TestAssert.Equal(expected.Length, actual.Length);
            foreach (var pair in expected.Zip(actual))
                TestAssert.True(pair.First.Text == pair.Second.Text && Math.Abs(pair.First.X - pair.Second.X) < .002d &&
                    Math.Abs(pair.First.Y - pair.Second.Y) < .002d && pair.First.Size == pair.Second.Size,
                    "Bookmarks, run splitting and equivalent direct/style control declarations preserve measured geometry.");
        }
    }

    public static void DocxInlineCompositionFractionSizeOwnershipRequiresResolvedSourceControls()
    {
        foreach (string variant in new[] { "direct", "style", "missing-style", "implicit" })
        {
            var source = ReadMutation("CompositionProbes/" + (variant is "style" or "missing-style" ?
                "fraction-control-style-24" : "fraction-control-body-24"), OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, _) =>
                {
                    var m = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/math";
                    var control = xml.Descendants(m + "ctrlPr").Single();
                    if (variant == "missing-style") control.Descendants(w + "rStyle").Single().SetAttributeValue(w + "val", "MissingPublicStyle");
                    if (variant == "implicit") control.Remove();
                });
            var equation = (variant is "style" or "missing-style" ? source.Tables.Single().Rows[0].Cells[0].Paragraphs :
                source.Paragraphs).SelectMany(paragraph => paragraph.Equations).Single();
            var fraction = (DocxMathFraction)equation.Content.Children.Single();
            TestAssert.Equal(variant is "direct" or "style" ? 24d : (double?)null, fraction.ControlFontSize);
            TestAssert.True(DocxMathLayout.TryGetRuns(equation, default, out var runs) && runs.All(run => run.Run.Text.FontSize == 11d),
                "Resolved control size must not overwrite the authored child run sizes.");
        }
    }
}
