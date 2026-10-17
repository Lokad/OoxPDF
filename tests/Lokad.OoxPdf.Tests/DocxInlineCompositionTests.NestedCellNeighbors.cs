using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNestedCellNeighborsKeepDirectPreparationOwnersAndFrames()
    {
        foreach (string placement in new[] { "plain", "before", "after" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/mixed-nested-cell-" + placement, mode, false, false);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(source).Single(p => p.Equations.Count != 0);
            TestAssert.True(ReferenceEquals(paragraph,
                DocxMathParagraphPolicy.EnumeratePreparationCandidates(source, default).Single(p => p.Equations.Count != 0)),
                "A neighboring table must not replace the actual direct paragraph owner.");
            if (placement != "plain")
                TestAssert.True(!DocxMathParagraphPolicy.IsSupportedSimpleCell(source.Tables[0].Rows[0].Cells[0]),
                    "Preparing direct paragraphs does not relax whole-cell fitting/placement eligibility.");
            var math = paragraph.Equations.ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 20d, 8d, []), 28d, 36d));
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            var cell = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).Single().Cells.Single();
            var lines = cell.TextLines.Where(line => ReferenceEquals(line.SourceParagraph, paragraph)).ToArray();
            TestAssert.Equal(2, lines.Sum(line => line.InlineMath.Count));
            TestAssert.Equal(1, cell.InlineImages.Count);
            if (placement == "after")
            {
                double bottom = lines.Min(line =>
                {
                    TestAssert.True(DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double lower, out _),
                        "The direct math line owns a complete occupied frame.");
                    return lower;
                });
                var nested = cell.NestedRows.Single();
                TestAssert.True(nested.Y + nested.Height <= bottom + .000001d,
                    "A following nested table consumes the measured occupied math frame, without overlap.");
            }
        }
    }

    public static void DocxInlineCompositionNestedCellNeighborsRetainUnqualifiedAndNestedSourceFallbacks()
    {
        foreach (string variant in new[] { "merge", "fit", "vertical", "height", "move", "inside" })
        {
            var source = ReadMutation("CompositionProbes/mixed-nested-cell-" + (variant == "inside" ? "inside" : "after"),
                OoxPdfDocxMarkupMode.Final, false, false, (xml, w, _) =>
                {
                    var table = xml.Root!.Element(w + "body")!.Element(w + "tbl")!;
                    var row = table.Element(w + "tr")!;
                    var cell = row.Element(w + "tc")!;
                    if (variant == "height")
                        row.AddFirst(new XElement(w + "trPr", new XElement(w + "trHeight",
                            new XAttribute(w + "val", "320"), new XAttribute(w + "hRule", "unknown"))));
                    else if (variant == "move")
                        cell.Element(w + "p")!.ReplaceWith(new XElement(w + "moveTo",
                            new XAttribute(w + "id", "71"), new XAttribute(w + "author", "Public author"), cell.Element(w + "p")));
                    else if (variant != "inside")
                        cell.Element(w + "tcPr")!.Add(variant switch
                        {
                            "merge" => new XElement(w + "vMerge", new XAttribute(w + "val", "restart")),
                            "fit" => new XElement(w + "tcFitText"),
                            _ => new XElement(w + "textDirection", new XAttribute(w + "val", "tbRl"))
                        });
                });
            TestAssert.True(!DocxMathParagraphPolicy.EnumeratePreparationCandidates(source, default).Any(p => p.Equations.Count != 0),
                "Unqualified cell properties and nested math retain explicit fallback: " + variant);
        }
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        var control = ReadMutation("CompositionProbes/mixed-nested-cell-after", OoxPdfDocxMarkupMode.Final, false, false);
        TestAssert.Throws<OperationCanceledException>(() =>
            DocxMathParagraphPolicy.EnumeratePreparationCandidates(control, cancelled.Token).ToArray());
    }

    public static void DocxInlineCompositionNestedCellNeighborsKeepIndependentContentAndFinalVerticalPlacement()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent nested-cell geometry uses installed Windows fonts.");
        CheckNestedCellNeighbors(null, checkGeometry: true);
    }

    internal static void CheckNestedCellNeighbors(IFontResolver? resolver, bool checkGeometry = false)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "mixed-nested-cells-office.json")));
        var controlX = new Dictionary<OoxPdfDocxMarkupMode, double[]>();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(folder, id + ".docx")))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            string? originalText = null;
            foreach (bool transformed in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
                int expected = row.GetProperty("AcceptedPaintedEquations").GetInt32();
                TestAssert.Equal(expected, text.Split("ratetime", StringSplitOptions.None).Length - 1);
                foreach (var prose in row.GetProperty("ExpectedText").EnumerateArray()) TestAssert.Contains(prose.GetString()!, text);
                TestAssert.Equal(row.GetProperty("ImagePlacements").GetInt32(), pages.Sum(page => page.Images.Sum(image =>
                    Regex.Matches(page.Content, "/" + Regex.Escape(image.ResourceName) + @"\s+Do\b").Count)));
                TestAssert.Equal(expected == 0, diagnostics.Any(diagnostic => diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION"));
                // PDF run boundaries can add extraction whitespace between
                // paragraphs; the prose assertions above keep authored spaces.
                if (originalText is not null) TestAssert.Equal(Regex.Replace(originalText, @"\s", ""), Regex.Replace(text, @"\s", ""));
                originalText = text;
                if (!checkGeometry || expected == 0 || transformed) continue;
                var shows = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                    .Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font)))).ToArray();
                TestAssert.Equal(16, shows.Length);
                var words = row.GetProperty("Math").EnumerateArray().ToArray();
                var control = gold.RootElement.GetProperty("Rows").EnumerateArray().Single(item =>
                    item.GetProperty("Id").GetString() == "mixed-nested-cell-plain" &&
                    item.GetProperty("Mode").GetString() == row.GetProperty("Mode").GetString()).GetProperty("Math").EnumerateArray().ToArray();
                if (id == "mixed-nested-cell-plain") controlX[mode] = shows.Where((_, index) => index % 4 == 0).Select(show => show.X).ToArray();
                for (int i = 0; i < words.Length; i++)
                {
                    TestAssert.Equal(words[i].GetProperty("Text").GetString(), string.Concat(shows.Skip(i * 4).Take(4).Select(show => show.Text)));
                    TestAssert.True(Math.Abs(shows[i * 4].X - controlX[mode][i] - words[i].GetProperty("X").GetDouble() +
                        control[i].GetProperty("X").GetDouble()) < .5d,
                        "A neighboring table preserves Word's independently measured horizontal change.");
                    // Absolute X and Review Y retain known failed visual gates.
                    // Final vertical placement has separate independent evidence.
                    if (mode == OoxPdfDocxMarkupMode.Final)
                        TestAssert.True(Math.Abs(shows[i * 4].Y - words[i].GetProperty("Y").GetDouble()) < .5d,
                            $"Direct math retains independently measured Final Y: {id}/word={i}.");
                }
            }
        }
    }
}
