using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionSourceAdvancesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent controls use installed Windows font metrics.");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "source-text-advances-office-lines.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string id = row.GetProperty("Id").GetString()!;
            string input = Path.Combine(folder, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var document = DocxTests.ReadDocx(input, mode);
            var owners = document.Paragraphs.SelectMany(p => p.Runs).ToArray();
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderBlankPages(document, null, default).Single();
            var lines = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).GroupBy(show => show.Y)
                .OrderByDescending(group => group.Key).Select(group =>
                    (Baseline: group.Key, Text: string.Concat(group.Select(show => show.Text)).Trim())).ToArray();
            var expected = row.GetProperty("Lines").EnumerateArray().ToArray();
            TestAssert.Equal(expected.Length, lines.Length);
            for (int index = 0; index < expected.Length; index++)
            {
                TestAssert.Equal(expected[index].GetProperty("Text").GetString(), lines[index].Text);
                TestAssert.True(Math.Abs(lines[0].Baseline - lines[index].Baseline - expected[index].GetProperty("Gap").GetDouble()) < .15d,
                    "Source thresholds retain independent Word line gaps: " + id);
            }
            var segments = renderer.InspectTextEmission(document).Lines.SelectMany(line => line.Segments)
                .Where(segment => !segment.IsTerminalLineSpace).ToArray();
            double firstY = segments.Max(segment => segment.BaselineY);
            double end = segments.Where(segment => Math.Abs(segment.BaselineY - firstY) < .01d)
                .Max(segment => segment.X + segment.AdvanceProfile.PlannedEmittedAdvance);
            TestAssert.True(Math.Abs(end - row.GetProperty("FirstLineEnd").GetDouble()) < .1d,
                "Source advances and physical glyph size retain the independent first-line endpoint: " + id);
            TestAssert.True(owners.Zip(document.Paragraphs.SelectMany(p => p.Runs)).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Measurement and emission cannot replace parsed source owners.");
        }
    }

    public static void DocxInlineCompositionSourceAdvancesShareMeasurementAndEmissionWithoutMutatingRuns()
    {
        var document = ReadMutation("CompositionProbes/overwide-word-default", OoxPdfDocxMarkupMode.Final, false, false);
        var run = document.Paragraphs[0].Runs[0];
        var resolver = new TestFaceFontResolver();
        var plan = DocxFontPlan.Create(document, resolver, default);
        var measurer = new DocxFontPlanTextMeasurer(plan, resolver.Resolve(new FontRequest("Calibri", false, false)), default,
            resolver, resolveSourceAdvances: true);
        double nominal = measurer.MeasureText(run, "AB", 12d);
        var met = run with { KerningMinimumHalfPointsValue = "16" };
        var metPlan = new DocxFontPlan(plan.Runs.Select(item => ReferenceEquals(item.Run, run) ? item with { Run = met } : item).ToArray());
        var metMeasurer = new DocxFontPlanTextMeasurer(metPlan, resolver.Resolve(new FontRequest("Calibri", false, false)), default,
            resolver, resolveSourceAdvances: true);
        TestAssert.True(Math.Abs(nominal - metMeasurer.MeasureText(met, "AB", 12d) - .6d) < .00001d,
            "The real synthetic -50-unit pair follows the source threshold in measurement.");
        foreach (var source in new[] { run, met })
        {
            var emission = DocxTextEmissionPlanner.CreateForEmissionSegment(source, 12d, 0d,
                DocxTextStateCharacterSpacingSource.None, false, false);
            TestAssert.Equal(source.KerningMinimumHalfPointsValue == "16", emission.PairKerningEnabled);
            TestAssert.Equal(12d, emission.NominalAdvanceFontSize!.Value);
            TestAssert.True(source.PairKerningEnabled && !source.UseLogicalSizeForPdfAdvances,
                "Resolved rendering purposes leave raw/derived eligibility state unchanged.");
        }
    }

    public static void DocxInlineCompositionSourceAdvancesRetainUnownedShiftedAndMarkerPurposes()
    {
        var run = new DocxTextRun("AB", 11d, null, false, false, false, null, "Calibri") { SourceRunIndex = 0 };
        foreach (var fallback in new[] { run with { SourceRunIndex = -1 }, run with { VerticalAlignmentValue = "superscript" },
            run with { VerticalAlignmentValue = "subscript" }, run with { IsAutomaticNoteReferenceMarker = true },
            run with { IsAutomaticNoteStoryMarker = true }, run with { KerningMinimumHalfPointsValue = "bad" } })
            TestAssert.Equal(new DocxTextAdvances(true, false), DocxTextAdvances.Resolve(fallback));
        TestAssert.Equal(new DocxTextAdvances(false, true), DocxTextAdvances.Resolve(run));
        TestAssert.Equal(new DocxTextAdvances(false, true), DocxTextAdvances.Resolve(run with {
            PairKerningEnabled = false, KerningMinimumHalfPointsValue = "16" }));
    }
}
