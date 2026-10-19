using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMergeCellMarkFloorsRetainMetricOwnersAndUnknownFallback()
    {
        var source = ReadMutation("CompositionProbes/merge-no-floor-source-mark", OoxPdfDocxMarkupMode.Final, false, false);
        var table = source.Tables.Single();
        var owner = table.Rows[0].Cells[0];
        var paragraph = owner.Paragraphs.Single();
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxTableRowLayout[] Rows(DocxParagraph changed)
        {
            var cell = owner with { Paragraphs = [changed], BodyElements = [new DocxParagraphElement(changed)] };
            var changedTable = table with { Rows = [table.Rows[0] with { Cells = [cell] }, table.Rows[1]] };
            var document = source with { BodyElements = [new DocxTableElement(changedTable)] };
            return engine.Create(document, new DocxTests.FamilyWidthTextMeasurer(), default).Pages
                .SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
        }
        foreach (double marker in new[] { 10d, 18d, 26d })
        {
            var changed = paragraph with { ParagraphMarkFontSize = marker };
            var empty = changed with { Runs = [], Images = [], Equations = [], ListLabel = null };
            var controlCell = owner with { HasVerticalMerge = false, VerticalMergeValue = null,
                Paragraphs = [empty], BodyElements = [new DocxParagraphElement(empty)] };
            var control = source with { BodyElements = [new DocxTableElement(table with
                { Rows = [table.Rows[0] with { Cells = [controlCell] }] })] };
            double expected = engine.Create(control, new DocxTests.FamilyWidthTextMeasurer(), default).Pages
                .SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).Single().Height;
            var actual = Rows(changed);
            TestAssert.True(Math.Abs(actual[0].Height - expected) < 1e-9d,
                "The terminal mark owns the same minimum strut as an ordinary empty cell; full content remains a span requirement.");
            TestAssert.Equal(1, actual[0].Table.AdmittedMergeSpanHeights.Count);
        }
        double floor = Rows(paragraph)[0].Height;
        foreach (double visibleSize in new[] { 12d, 20d, 28d })
        {
            var changed = paragraph with { Runs = paragraph.Runs.Select(run => run with { FontSize = visibleSize }).ToArray() };
            TestAssert.True(Math.Abs(Rows(changed)[0].Height - floor) < 1e-9d,
                "Visible run formatting cannot replace the terminal marker's independent minimum while its complete source fits the span.");
        }
        double legacy = Rows(paragraph with { ParagraphMarkFontSize = null })[0].Height;
        foreach (double invalid in new[] { 0d, -1d, double.NaN, double.PositiveInfinity })
            TestAssert.Equal(legacy, Rows(paragraph with { ParagraphMarkFontSize = invalid })[0].Height);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => engine.Create(source,
            new DocxTests.FamilyWidthTextMeasurer(), cancelled.Token));
    }
}
