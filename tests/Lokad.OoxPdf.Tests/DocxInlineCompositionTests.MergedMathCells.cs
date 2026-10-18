using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMergedCellMathRetainsIndependentRelativeWordGeometry()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent merged-cell geometry requires installed Windows fonts.");
        CheckMergedMathCells(null);
    }

    internal static void CheckMergedMathCells(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("merged-math-cells-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            byte[] input = File.ReadAllBytes(NestedMathCasePath(id + ".docx"));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(
                SHA256.HashData(input)));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed);
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, diagnostics.Add, default).Single();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "A merge owner's supported mathematics must not disappear.");
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
                var lines = shows.GroupBy(show => show.Y).Select(group =>
                {
                    var ordered = group.OrderBy(show => show.X).ToArray();
                    return (Text: string.Concat(ordered.Select(show => show.Text)), First: ordered[0]);
                }).ToArray();
                foreach (string prefix in new[] { "Before public math table.", "Before inline", "after equation.", "After public math table." })
                    TestAssert.True(lines.Any(line => line.Text.Contains(prefix, StringComparison.Ordinal)),
                        "Every surrounding source prose fragment remains visible.");
                var math = shows.Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                    ReferenceEquals(font.Font, show.Font))).ToArray();
                TestAssert.Equal(4, math.Length);
                var prose = lines.Single(line => line.Text.StartsWith("Before inline", StringComparison.Ordinal)).First;
                TestAssert.True(math.All(show => show.Y > lines.Single(line =>
                    line.Text.StartsWith("After public", StringComparison.Ordinal)).First.Y),
                    "Every fraction glyph clears the following paragraph.");
                if (resolver is not null) continue;
                foreach (var pair in math.Zip(row.GetProperty("MathAnchors").EnumerateArray()))
                    TestAssert.True(pair.First.Text == pair.Second.GetProperty("Text").GetString() &&
                        Math.Abs(pair.First.X - prose.X - pair.Second.GetProperty("RelativeX").GetDouble()) < 1d &&
                        Math.Abs(pair.First.Y - prose.Y - pair.Second.GetProperty("RelativeY").GetDouble()) < 1d &&
                        Math.Abs(pair.First.Size - pair.Second.GetProperty("Size").GetDouble()) < .1d,
                        "The fraction retains independent one-point geometry relative to its cell prose. Raw cell-left/after-flow residuals remain separate.");
            }
        }
    }

    public static void DocxInlineCompositionMergedMathPreparationKeepsOwningSourcesAndPlacementFallbacks()
    {
        var source = DocxTests.ReadDocx(NestedMathCasePath("merged-math-cell.docx"), OoxPdfDocxMarkupMode.Final);
        var table = source.Tables.Single();
        var owner = table.Rows[0].Cells.Single();
        var continuation = table.Rows[1].Cells.Single();
        var paragraph = owner.Paragraphs.Single();
        TestAssert.True(owner.IsVerticalMergeRestart && !continuation.IsVerticalMergeRestart,
            "Only an explicit restart owns the merged source content.");
        TestAssert.True(!DocxMathParagraphPolicy.IsSupportedCellProperties(owner),
            "Source preparation does not relax existing placement/print-frame eligibility.");
        foreach (string variant in new[] { "restart", "continue", "unknown", "missing", "fit", "vertical" })
        {
            var changed = owner with { VerticalMergeValue = variant == "missing" ? null :
                variant is "fit" or "vertical" ? "restart" : variant,
                FitText = variant == "fit", TextDirectionValue = variant == "vertical" ? "tbRl" : owner.TextDirectionValue };
            var changedTable = table with { Rows = [table.Rows[0] with { Cells = [changed] }, table.Rows[1]] };
            var elements = source.BodyElements.Select(element => element is DocxTableElement ?
                new DocxTableElement(changedTable) : element).ToArray();
            var document = source with { BodyElements = elements, Tables = [changedTable] };
            var candidates = DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, default)
                .Where(p => p.Equations.Count != 0).ToArray();
            TestAssert.Equal(variant == "restart" ? 1 : 0, candidates.Length);
            if (candidates.Length != 0) TestAssert.True(ReferenceEquals(paragraph, candidates[0]), "Preparation retains the exact visible source paragraph.");
            var staticSource = document with
            {
                BodyElements = [], Tables = [],
                HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>
                    { ["default"] = [new DocxTableElement(changedTable)] },
                EquationParts = [new DocxEquationPartSource("/word/header1.xml", 1, [new DocxTableElement(changedTable)])]
            };
            var staticCandidates = DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(staticSource, default).ToArray();
            TestAssert.Equal(variant == "restart" ? 1 : 0, staticCandidates.Length);
            if (staticCandidates.Length != 0)
                TestAssert.True(ReferenceEquals(paragraph, staticCandidates[0].Paragraph), "Static preparation shares the same visible merge source.");
        }
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() =>
            DocxMathParagraphPolicy.EnumeratePreparationCandidates(source, cancelled.Token).ToArray());
    }
}
