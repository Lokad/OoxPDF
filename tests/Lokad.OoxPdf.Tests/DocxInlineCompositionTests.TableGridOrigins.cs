using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxAuthoredZeroTableIndentsKeepIndependentWordAnchors()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent authored-zero origins use installed Windows fonts.");
        CheckTableGridOrigins(null, "table-zero-indents-office.json");
    }

    public static void DocxAuthoredTableMarginsKeepIndependentBorderedWordOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent bordered-margin origins use installed Windows fonts.");
        CheckTableGridOrigins(null, "table-implicit-border-origins-office.json");
    }

    public static void DocxAuthoredTableMarginPresenceKeepsUnsetAndStaticGridFallbacks()
    {
        foreach (double? margin in new double?[] { null, 0d, 10d })
        foreach (string width in new[] { "4", "16" })
        {
            var paragraph = DocxTests.CreateDocxLayoutParagraph("Public cell", 11d, 16d);
            var borders = new[] { new DocxTableCellBorder("left", "single", "000000", width) };
            var cell = new DocxTableCell("Public cell", [paragraph], null, null, null, null, borders,
                DocxTableCellMargins.Empty with { LeftPoints = margin, RightPoints = margin });
            var basis = new DocxTable("fixed", [200d], [new DocxTableRow([cell], null)]);
            foreach (bool authored in new[] { false, true })
            {
                var page = LayoutSharedTable(authored ? basis with { IndentPoints = 0d } : basis);
                var flowing = page.Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
                var stationary = page.StaticTableRows.Single().Cells.Single();
                TestAssert.True(Math.Abs(stationary.X - page.MarginLeft) < .001d,
                    "Static containers preserve their grid anchor with absent or authored zero indentation.");
                TestAssert.Equal(flowing.Width, stationary.Width);
                if (authored || margin is not null)
                    TestAssert.True(Math.Abs(flowing.TextLines.Single().X - page.MarginLeft) < .001d,
                        "Authored indentation or margins own flowing first text, including a visible starting border.");
                else
                    TestAssert.Equal(stationary.X, flowing.X);
            }
        }
    }

    public static void DocxInlineCompositionTableGridOriginsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent table-grid typography requires installed Windows fonts.");
        CheckTableGridOrigins(null);
    }

    public static void DocxInlineCompositionSignedTableIndentsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent signed-table typography requires installed Windows fonts.");
        CheckTableGridOrigins(null, "table-signed-indents-office.json");
    }

    public static void DocxInlineCompositionSignedTableIndentsKeepContainerWidthAndOneGrid()
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "TableGridOrigins"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "table-signed-indents-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray()
            .Where(row => row.GetProperty("Mode").GetString() == "final-preserve"))
        {
            string id = row.GetProperty("Id").GetString()!;
            var source = ReadMutation("TableGridOrigins/" + id, OoxPdfDocxMarkupMode.Final, false, false);
            TestAssert.True(source.BodyElements.OfType<DocxTableElement>().Single().Table.IndentPoints < 0d,
                "The reader preserves the authored signed indentation.");
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            var body = layout.Pages[0].Items.OfType<DocxTextLineLayout>().First();
            var rows = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
            var first = rows[0].Cells[0];
            var text = first.TextLines.First(line => line.Text.StartsWith("Before inline", StringComparison.Ordinal));
            TestAssert.True(Math.Abs(text.X - body.X - row.GetProperty("LogicalFirstCellOffsetPoints").GetDouble()) < .002d,
                id + ": the signed first-text offset belongs to the shared grid.");
            TestAssert.True(rows.All(item => Math.Abs(item.Cells[0].X - first.X) < .002d),
                "Rows keep one grid when the table hangs into the margin.");
            TestAssert.True(Math.Abs(first.Width - row.GetProperty("WordCellWidthPoints").GetDouble()) < .1d,
                id + ": hanging placement preserves the independent Word fixed/auto/percentage width.");
        }
    }

    public static void DocxInlineCompositionAuthoredCellMarginsKeepOneGridAcrossRows()
    {
        foreach (string id in new[] { "table-origin-direct-wide", "table-origin-cell-wide", "table-origin-style-zero" })
        {
            var source = ReadMutation("TableGridOrigins/" + id, OoxPdfDocxMarkupMode.Final, false, false);
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default);
            var body = layout.Pages[0].Items.OfType<DocxTextLineLayout>().First();
            var rows = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
            var cellText = rows[0].Cells[0].TextLines.First(line => line.Text.StartsWith("Before inline", StringComparison.Ordinal));
            TestAssert.True(Math.Abs(body.X - cellText.X) < 1e-9d,
                "An authored starting margin belongs to the shared grid, so first text stays at the container origin.");
            TestAssert.True(rows.All(row => Math.Abs(row.Cells[0].X - rows[0].Cells[0].X) < 1e-9d),
                "All rows keep the same first grid edge even when cell margins differ.");
        }
    }

    internal static void CheckTableGridOrigins(IFontResolver? resolver, string goldenFile = "table-grid-origins-office.json")
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "TableGridOrigins"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, goldenFile)));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(
                SHA256.HashData(File.ReadAllBytes(Path.Combine(directory, id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            List<(string Text, double X, double Y)>? original = null;
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation("TableGridOrigins/" + id, mode, transformed, transformed);
                var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, default).ToArray();
                var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                    usePageCoordinates: true).Select(show => (Page: index + 1, Height: page.Height, Show: show))).ToArray();
                var lines = shows.GroupBy(item => (item.Page, item.Show.Y)).Select(group => {
                    var ordered = group.OrderBy(item => item.Show.X).ToArray();
                    return (group.Key.Page, Text: string.Concat(ordered.Select(item => item.Show.Text)), First: ordered[0]);
                }).ToArray();
                foreach (string prefix in new[] { "Before public math table.", "Before inline", "after equation.", "After public math table." })
                    TestAssert.Equal(1, lines.Count(line => line.Text.Contains(prefix, StringComparison.Ordinal)));
                TestAssert.Equal(4, pages.Sum(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Count(show =>
                    page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font)))));
                var anchors = new List<(string Text, double X, double Y)>();
                foreach (var anchor in row.GetProperty("Anchors").EnumerateArray())
                {
                    string text = anchor.GetProperty("Text").GetString()!.TrimEnd();
                    var actual = lines.Single(line => line.Text.StartsWith(text, StringComparison.Ordinal));
                    var first = actual.First;
                    anchors.Add((text, first.Show.X, first.Show.Y));
                    TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
                    if (resolver is not null) continue;
                    TestAssert.True(Math.Abs(first.Show.X - anchor.GetProperty("X").GetDouble()) < 1d,
                        id + ": first text and surrounding flow retain independent one-point Word X geometry.");
                    TestAssert.True(Math.Abs(first.Height - first.Show.Y - anchor.GetProperty("Y").GetDouble()) < 1d &&
                        Math.Abs(first.Show.Size - anchor.GetProperty("Size").GetDouble()) < .1d,
                        id + ": the grid-origin change preserves independent Word baseline/size geometry.");
                }
                if (resolver is not null)
                {
                    double expectedDelta = row.TryGetProperty("LogicalFirstCellOffsetPoints", out var logicalOffset)
                        ? logicalOffset.GetDouble() : row.GetProperty("Anchors")[1].GetProperty("X").GetDouble() -
                        row.GetProperty("Anchors")[0].GetProperty("X").GetDouble();
                    TestAssert.True(Math.Abs(anchors[1].X - anchors[0].X - expectedDelta) < .002d,
                        "Starting-margin ownership is independent of the selected portable font.");
                }
                if (original is not null)
                    foreach (var pair in original.Zip(anchors))
                        TestAssert.True(pair.First.Text == pair.Second.Text && Math.Abs(pair.First.X - pair.Second.X) < .002d &&
                            Math.Abs(pair.First.Y - pair.Second.Y) < .002d, "Bookmarks and run splitting preserve the shared table grid and source content.");
                original = anchors;
            }
        }
    }
}
