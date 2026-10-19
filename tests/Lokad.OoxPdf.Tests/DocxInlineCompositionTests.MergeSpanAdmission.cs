using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionOverlappingMergeSpansPreserveIndependentOwnersAndFloors()
    {
        foreach (bool offset in new[] { false, true })
        {
            var source = ReadMutation("CompositionProbes/merged-span-two-owners", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, _) =>
                {
                    if (!offset) return;
                    var rows = xml.Root!.Element(w + "body")!.Element(w + "tbl")!.Elements(w + "tr").ToArray();
                    var first = rows[0].Elements(w + "tc").Last();
                    var next = new XElement(first);
                    rows[1].Elements(w + "tc").Last().ReplaceWith(next);
                    first.Element(w + "tcPr")!.Element(w + "vMerge")!.Remove();
                    first.Element(w + "p")!.ReplaceWith(new XElement(w + "p"));
                });
            var table = source.Tables.Single();
            var owners = table.Rows.SelectMany((row, index) => row.Cells.Select((cell, column) => (cell, index, column)))
                .Where(item => item.cell.IsVerticalMergeRestart).ToArray();
            TestAssert.Equal(2, owners.Length);
            var prepared = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance);
            foreach (var owner in owners)
            {
                var equation = owner.cell.Paragraphs[0].Equations.Single();
                double height = owner.column == 0 ? 50d : 90d;
                prepared.Add(equation, new(equation, new(18d, height - 8d, 8d, []), height - 8d, height + 4d));
            }
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var layout = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: prepared);
            var rows = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
            TestAssert.Equal(1, layout.Pages.Count);
            TestAssert.Equal(2, rows[0].Table.AdmittedMergeSpanHeights.Count);
            var baseline = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default).Pages
                .SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
            TestAssert.True(rows.Zip(baseline).All(pair => pair.First.Height + 1e-9d >= pair.Second.Height),
                "Overlapping constraints cannot spend another row's existing floor.");
            foreach (var owner in owners)
            {
                int count = owner.column == 0 ? 2 : offset ? 2 : 3;
                var cell = owner.cell with { HasVerticalMerge = false, VerticalMergeValue = null };
                var standalone = source with { BodyElements = [new DocxTableElement(table with
                    { Rows = [table.Rows[owner.index] with { Cells = [cell] }] })] };
                double required = engine.Create(standalone, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: prepared)
                    .Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).Single().Height;
                TestAssert.True(rows.Skip(owner.index).Take(count).Sum(row => row.Height) >= required,
                    "Every independent owner fits its complete span, including overlapping spans that begin on different rows.");
            }
            TestAssert.Equal(2, rows.SelectMany(row => row.Cells).SelectMany(cell => cell.TextLines)
                .Count(line => line.Text.StartsWith("Before inline", StringComparison.Ordinal)));
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => engine.Create(source,
                new DocxTests.FamilyWidthTextMeasurer(), cancelled.Token, mathEquations: prepared));
        }
    }

    public static void DocxInlineCompositionMergeSpansRequireCompleteStaticAdmissionAndKeepRowFloors()
    {
        foreach (string variant in new[] { "complete", "large", "minimum", "exact", "small-page", "forced-break",
            "dynamic", "unmeasured", "center", "continuation-source", "header", "note-reference" })
        {
            var source = ReadMutation("CompositionProbes/merged-span-three-rows", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, _) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    var rows = body.Element(w + "tbl")!.Elements(w + "tr").ToArray();
                    var owner = rows[0].Element(w + "tc")!;
                    if (variant is "minimum" or "exact")
                        foreach (var row in rows)
                            row.AddFirst(new XElement(w + "trPr", new XElement(w + "trHeight",
                                new XAttribute(w + "val", variant == "minimum" ? "800" : "400"),
                                new XAttribute(w + "hRule", variant == "minimum" ? "atLeast" : "exact"))));
                    if (variant == "small-page")
                    {
                        var section = body.Element(w + "sectPr")!;
                        section.Element(w + "pgSz")!.SetAttributeValue(w + "h", "3000");
                        section.Element(w + "pgMar")!.SetAttributeValue(w + "top", "200");
                        section.Element(w + "pgMar")!.SetAttributeValue(w + "bottom", "200");
                    }
                    if (variant == "center") owner.Element(w + "tcPr")!.Element(w + "vAlign")!.SetAttributeValue(w + "val", "center");
                    if (variant == "dynamic") owner.Element(w + "p")!.Add(new XElement(w + "fldSimple",
                        new XAttribute(w + "instr", " PAGE "), new XElement(w + "r", new XElement(w + "t", "1"))));
                    if (variant == "forced-break") owner.Element(w + "p")!.Add(new XElement(w + "r",
                        new XElement(w + "br", new XAttribute(w + "type", "page"))));
                    if (variant == "note-reference") owner.Element(w + "p")!.Add(new XElement(w + "r",
                        new XElement(w + "footnoteReference", new XAttribute(w + "id", "1"))));
                    if (variant == "continuation-source") rows[1].Element(w + "tc")!.Element(w + "p")!
                        .Add(new XElement(w + "r", new XElement(w + "t", "Continuation source.")));
                    if (variant == "header") rows[0].AddFirst(new XElement(w + "trPr", new XElement(w + "tblHeader")));
                });
            var equation = source.Tables.Single().Rows[0].Cells[0].Paragraphs[0].Equations.Single();
            double height = variant == "small-page" ? 200d : variant == "large" ? 70d : 28d;
            var prepared = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
            { [equation] = new(equation, new(18d, height - 8d, 8d, []), height - 8d, height + 4d) };
            var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var layout = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default,
                mathEquations: variant == "unmeasured" ? null : prepared);
            var rows = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
            bool admitted = variant is "complete" or "large" or "minimum";
            TestAssert.True(rows.All(row => (row.Table.AdmittedMergeSpanHeights.Count != 0) == admitted),
                "Only complete static measured spans may borrow row space. Fixed, partial, dynamic and source-bearing continuations retain staging.");
            if (!admitted) continue;
            TestAssert.True(rows.All(row => ReferenceEquals(row.Table, rows[0].Table)),
                "All rows reuse the complete frame admitted at the original origin.");
            var baseline = engine.Create(source, new DocxTests.FamilyWidthTextMeasurer(), default).Pages
                .SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
            TestAssert.True(rows.Zip(baseline).All(pair => pair.First.Height + 1e-9d >= pair.Second.Height),
                variant + ": measured span ownership preserves every pre-existing row floor (within floating-point subtraction precision).");
            if (variant == "minimum") TestAssert.True(rows.All(row => row.Height >= 40d), "Authored minimum heights remain floors.");
            var table = source.Tables.Single();
            var cell = table.Rows[0].Cells[0] with { HasVerticalMerge = false, VerticalMergeValue = null };
            var standalone = source with { BodyElements = [new DocxTableElement(table with
                { Rows = [table.Rows[0] with { Cells = [cell] }] })] };
            double required = engine.Create(standalone, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: prepared)
                .Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).Single().Height;
            TestAssert.True(Math.Abs(rows.Sum(row => row.Height) - Math.Max(baseline.Sum(row => row.Height), required)) < .000001d,
                "A complete span pays for its measured owner once, beyond existing row floors, rather than expanding its restart plus every continuation.");
            TestAssert.True(rows.Skip(1).All(row => row.Cells[0].TextLines.Count == 0), "Admitted complete source is painted only at its restart.");
        }
    }
}
