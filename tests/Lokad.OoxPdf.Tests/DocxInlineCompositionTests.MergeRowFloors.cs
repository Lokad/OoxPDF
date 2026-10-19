using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNonMathMergeSpansKeepIndependentFloorsAndPropertyFallbacks()
    {
        foreach (string variant in new[] { "complete", "fit", "vertical", "center", "unmeasured", "oversized" })
        {
            var source = ReadMutation("CompositionProbes/merge-row-floor-wrap-run", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, _) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    var properties = body.Element(w + "tbl")!.Element(w + "tr")!.Element(w + "tc")!.Element(w + "tcPr")!;
                    if (variant == "fit") properties.Add(new XElement(w + "tcFitText"));
                    if (variant == "vertical") properties.Add(new XElement(w + "textDirection", new XAttribute(w + "val", "tbRl")));
                    if (variant == "center") properties.Element(w + "vAlign")!.SetAttributeValue(w + "val", "center");
                    if (variant == "oversized")
                    {
                        var section = body.Element(w + "sectPr")!;
                        section.Element(w + "pgSz")!.SetAttributeValue(w + "h", "1000");
                        section.Element(w + "pgMar")!.SetAttributeValue(w + "top", "200");
                        section.Element(w + "pgMar")!.SetAttributeValue(w + "bottom", "200");
                    }
                });
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var layout = engine.Create(source, variant == "unmeasured" ? null : new DocxTests.FamilyWidthTextMeasurer(), default);
            var rows = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
            TestAssert.True(rows.All(row => (row.Table.AdmittedMergeSpanHeights.Count != 0) == (variant == "complete")),
                "Plain source spans share the existing static/complete/property fallback contract without requiring an unrelated equation.");
            if (variant != "complete") continue;
            var table = source.Tables.Single();
            var empty = table.Rows[0].Cells[0] with { Text = string.Empty, Paragraphs = [], BodyElements = [],
                HasVerticalMerge = false, VerticalMergeValue = null };
            var controlTable = table with { Rows = [table.Rows[0] with { Cells = [empty, table.Rows[0].Cells[1]] }, table.Rows[1]] };
            var control = source with { BodyElements = [new DocxTableElement(controlTable)] };
            double independentFloor = engine.Create(control, new DocxTests.FamilyWidthTextMeasurer(), default).Pages
                .SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).First().Height;
            TestAssert.True(Math.Abs(rows[0].Height - independentFloor) < 1e-9d,
                "Wrapped source and direct run formatting cannot expand an independently owned row floor.");
            TestAssert.True(rows[0].Cells[0].TextLines.Any(line => line.InlineFrame is null) &&
                rows[1].Cells[0].TextLines.Count != 0,
                "The line-only model measurer cannot prove complete source bounds; continuation painting retains its existing staged fallback. Real-font observers independently verify complete source consumption.");
        }
    }

    public static void DocxInlineCompositionMergedSourceRowFloorsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent merged-source flow requires installed Windows fonts.");
        CheckMergeSourceRowFloors(null);
    }

    public static void DocxInlineCompositionMergeCellMarksMatchIndependentWordFloors()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent cell-terminal marker metrics require installed Windows fonts.");
        CheckMergeSourceRowFloors(null, "merge-cell-mark-floors-office.json");
    }

    internal static void CheckMergeSourceRowFloors(IFontResolver? resolver,
        string goldenFile = "merge-source-row-floors-office.json")
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath(goldenFile)));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(
                SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            string? original = null;
            List<(string Text, int Page, double Y)>? originalFlow = null;
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation("CompositionProbes/" + id, mode, false, transformed,
                    (xml, w, _) =>
                    {
                        if (!transformed) return;
                        var paragraph = xml.Root!.Element(w + "body")!.Element(w + "tbl")!
                            .Element(w + "tr")!.Element(w + "tc")!.Element(w + "p")!;
                        var first = paragraph.Elements().First(element => element.Name != w + "pPr");
                        first.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "902"),
                            new XAttribute(w + "name", "PublicMergedProse")));
                        paragraph.Add(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "902")));
                    });
                var pages = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, default).ToArray();
                var shows = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                    usePageCoordinates: true).Select(show => (Page: index + 1, show.Text, show.X, show.Y))).ToArray();
                string visible = Compact(string.Concat(shows.Select(show => show.Text)));
                foreach (var expected in row.GetProperty("SourceProse").EnumerateArray())
                    TestAssert.Equal(1, visible.Split(Compact(expected.GetString()!), StringSplitOptions.None).Length - 1);
                foreach (string expected in new[] { "Before public math table.", "After public math table." }
                    .Concat(row.GetProperty("FlowAnchors").EnumerateArray().Select(anchor => anchor.GetProperty("Text").GetString()!))
                    .Distinct(StringComparer.Ordinal))
                    TestAssert.Equal(1, visible.Split(Compact(expected), StringSplitOptions.None).Length - 1);
                TestAssert.Equal(row.GetProperty("Images").GetInt32(), pages.Sum(page => page.Images.Count));
                var lines = shows.GroupBy(show => (show.Page, show.Y)).Select(group => (group.Key.Page, group.Key.Y,
                    Text: string.Concat(group.OrderBy(show => show.X).Select(show => show.Text)))).ToArray();
                var flow = new List<(string Text, int Page, double Y)>();
                foreach (var anchor in row.GetProperty("FlowAnchors").EnumerateArray())
                {
                    string text = anchor.GetProperty("Text").GetString()!;
                    var actual = lines.Single(line => line.Text.Contains(text, StringComparison.Ordinal));
                    flow.Add((text, actual.Page, actual.Y));
                    if (resolver is null)
                        TestAssert.True(actual.Page == anchor.GetProperty("Page").GetInt32() &&
                            Math.Abs(actual.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                            id + ": independent non-spanning cells own row floors; the merged source owns its complete span.");
                }
                if (resolver is null) TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
                if (original is not null)
                {
                    TestAssert.Equal(original, visible);
                    foreach (var pair in originalFlow!.Zip(flow))
                        TestAssert.True(pair.First.Text == pair.Second.Text && pair.First.Page == pair.Second.Page &&
                            Math.Abs(pair.First.Y - pair.Second.Y) < .002d,
                            "Bookmarks and unchanged-format run splitting preserve source content and complete-span placement.");
                }
                original = visible;
                originalFlow = flow;
            }
        }
        static string Compact(string text) => new(text.Normalize(NormalizationForm.FormKC).Where(c => !char.IsWhiteSpace(c)).ToArray());
    }
}
