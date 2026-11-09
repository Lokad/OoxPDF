using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxNaturalBodyLineTests
{
    public static void DocxNaturalBodyLinesKeepActualLineMetricsAndSourceOwners()
    {
        foreach (bool reverse in new[] { false, true })
        foreach (bool split in new[] { false, true })
        {
            var paragraph = Mixed(reverse, split);
            var actual = Layout([paragraph]).Pages.Single().Items.OfType<DocxTextLineLayout>().ToArray();
            // Split runs stay on their line; the unsplit control supplies the same
            // visible content in separate zero-spacing paragraphs.
            var control = Layout(Mixed(reverse, false).Runs.Select(run =>
                paragraph with { Runs = [run with { Text = run.Text.Trim() }] }).ToArray());
            var expected = control.Pages.Single().Items.OfType<DocxTextLineLayout>().ToArray();
            TestAssert.Equal(2, actual.Length);
            for (int index = 0; index < actual.Length; index++)
            {
                TestAssert.Equal(expected[index].LineHeight, actual[index].LineHeight);
                TestAssert.True(Math.Abs(expected[index].BaselineY - actual[index].BaselineY) < 1e-9,
                    "A break shares the same natural metrics as the same zero-gap paragraph boundary.");
                TestAssert.True(ReferenceEquals(paragraph, actual[index].SourceParagraph), "Metric views cannot replace source/navigation owners.");
                TestAssert.True(actual[index].Segments.All(segment => paragraph.Runs.Any(run => ReferenceEquals(run, segment.StyleRun))),
                    "Placed segments retain their original run owners.");
            }
        }
        var sameLine = Mixed(false, false);
        sameLine = sameLine with { Runs = sameLine.Runs.Select(run => run with { Text = run.Text.Replace('\n', ' ') }).ToArray() };
        TestAssert.Equal(18d, Layout([sameLine]).Pages.Single().Items.OfType<DocxTextLineLayout>().Single().FontSize);
    }

    public static void DocxNaturalBodyLinesKeepAndWidowAdmissionUseActualAdvances()
    {
        var filler = DocxTests.CreateDocxLayoutParagraph("Fill", 10d, 10d);
        var next = DocxTests.CreateDocxLayoutParagraph("Next", 10d, 10d);
        var kept = Mixed(false, false) with { KeepRules = DocxParagraphKeepRules.Empty with { KeepLines = true, KeepNext = true } };
        foreach ((double height, int pages) in new[] { (78d, 1), (72d, 2) })
        {
            var layout = Layout([filler, kept, next], height);
            TestAssert.Equal(pages, layout.Pages.Count);
            var owner = layout.Pages.Single(page => page.Items.OfType<DocxTextLineLayout>().Any(line => ReferenceEquals(line.SourceParagraph, kept)));
            TestAssert.Equal(2, owner.Items.OfType<DocxTextLineLayout>().Count(line => ReferenceEquals(line.SourceParagraph, kept)));
            TestAssert.True(owner.Items.OfType<DocxTextLineLayout>().Any(line => ReferenceEquals(line.SourceParagraph, next)),
                "The kept pair fits or moves using its actual occupied line heights.");
        }
        var widow = Mixed(false, false);
        widow = widow with { Runs = [widow.Runs[0], widow.Runs[1] with { Text = "small\nlast" }] };
        TestAssert.Equal(1, Layout([filler, widow], 83d).Pages.Count);
    }

    public static void DocxNaturalBodyLinesPreserveDeclaredAndUnmeasuredFallbacks()
    {
        foreach (string rule in new[] { "exact", "atLeast" })
        {
            var paragraph = Mixed(false, false) with { LineSpacingPoints = 40d,
                Spacing = DocxParagraphSpacing.Empty with { LineValue = "800", LineRuleValue = rule } };
            var lines = Layout([paragraph]).Pages.Single().Items.OfType<DocxTextLineLayout>().ToArray();
            TestAssert.Equal(2, lines.Length);
            TestAssert.True(lines.All(line => line.LineHeight == 40d), "Declared line slots retain their independent contract.");
        }
        var source = Mixed(false, false);
        var document = Document([source], 200d);
        var fallback = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new AdvanceOnlyMeasurer(), CancellationToken.None);
        TestAssert.True(fallback.Pages.Single().Items.OfType<DocxTextLineLayout>().All(line => line.FontSize == 18d),
            "Providers without font metrics retain the previous paragraph fallback.");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), canceled.Token));
    }

    private static DocxParagraph Mixed(bool reverse, bool split)
    {
        var paragraph = DocxTests.CreateDocxLayoutParagraph("large\n", reverse ? 12d : 18d, 20d) with
        { LineSpacingPoints = null, Spacing = DocxParagraphSpacing.Empty with { LineValue = "240", LineRuleValue = "auto" } };
        var first = paragraph.Runs[0];
        var second = first with { Text = "small", FontSize = reverse ? 18d : 12d };
        return paragraph with { Runs = split ? [first, second with { Text = "sm" }, second with { Text = "all" }] : [first, second] };
    }

    private static DocxDocument Document(DocxParagraph[] paragraphs, double height) =>
        DocxTests.CreateLayoutTestDocument(paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p)).ToArray(), [])
            with { PageHeightPoints = height };

    private static DocxLayout Layout(DocxParagraph[] paragraphs, double height = 200d) =>
        new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(Document(paragraphs, height), new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

    private sealed class AdvanceOnlyMeasurer : IDocxTextMeasurer
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => text.Length * 5d;
    }
}
