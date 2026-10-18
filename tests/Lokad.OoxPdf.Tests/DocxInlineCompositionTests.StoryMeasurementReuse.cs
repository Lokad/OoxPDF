using System.Reflection;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionStoryMeasurementsSurvivePlacementRetryWithoutRemeasuringNotes()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = DocxTests.ReadDocx(NestedMathCasePath("two-plain-notes-report.docx"), mode);
            var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var runs = new HashSet<DocxTextRun>(source.RelatedStories.Where(story => story.IsNormalStoryType &&
                story.Kind == DocxRelatedStoryKind.Footnote).SelectMany(story => story.Paragraphs).SelectMany(p => p.Runs),
                ReferenceEqualityComparer.Instance);
            var measurer = new NoteMeasurementCounter(runs);
            var engine = new DocxLayoutEngine(geometry);
            // Reflection permits the same regression to run against the parent
            // library, where the retry still has no shared measurement owner.
            typeof(DocxLayoutEngine).GetProperty("PlacementRelatedStoryMeasurements", BindingFlags.Instance | BindingFlags.NonPublic)
                ?.SetValue(engine, new Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>());
            var first = engine.Create(source, measurer, default);
            TestAssert.True(measurer.NoteCalls > 0, "The source note fragments are measured before placement.");
            measurer.NoteCalls = 0;
            var shifted = engine.Create(source, measurer, default,
                headerOverflowDisplacementByPage: new Dictionary<int, double> { [0] = 8d });
            TestAssert.Equal(0, measurer.NoteCalls);
            TestAssert.True(first.Pages[0].Items.OfType<DocxTextLineLayout>().First().BaselineY >
                shifted.Pages[0].Items.OfType<DocxTextLineLayout>().First().BaselineY,
                "Body placement still consumes the changed header constraint.");
            foreach (var canonical in first.RelatedStories)
            {
                var reused = shifted.RelatedStories.Single(story => story.StoryIndex == canonical.StoryIndex);
                TestAssert.True(canonical.TextLines.Select(line => (line.Text, line.SourceParagraph, line.SourceLineIndex))
                    .SequenceEqual(reused.TextLines.Select(line => (line.Text, line.SourceParagraph, line.SourceLineIndex))),
                    "Measurement reuse retains source text and paragraph/line provenance.");
            }
        }
    }

    private sealed class NoteMeasurementCounter(HashSet<DocxTextRun> noteRuns) : IDocxTextMeasurer
    {
        private readonly DocxTests.FamilyWidthTextMeasurer inner = new();
        public int NoteCalls { get; set; }
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            if (run is not null && noteRuns.Contains(run)) NoteCalls++;
            return inner.MeasureText(run, text, fontSize);
        }
    }
}
