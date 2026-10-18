using System.Text;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMixedMathPictureNotesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent mixed-note frames require installed Windows fonts.");
        CheckPictureNoteBatches(null, "mixed-math-picture-note-report");
    }

    public static void DocxInlineCompositionMixedNoteFramesRequireCompleteMeasuredEquationOwners()
    {
        var source = DocxTests.ReadDocx(NestedMathCasePath("mixed-math-picture-note-report.docx"), OoxPdfDocxMarkupMode.Final);
        var originalStory = source.RelatedStories.Single(candidate => candidate.Kind == DocxRelatedStoryKind.Footnote && candidate.Id == "2");
        var original = originalStory.Paragraphs[0];
        var run = new DocxTextRun("A", 10d, null, false, false, false, null, "TestFont");
        var equation = original.Equations.Single();
        var font = new PdfFontResource("Eq1", PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(),
            run.Text.EnumerateRunes().Select(rune => rune.Value), default));
        // One atomic portable measurement keeps the immutable source owner;
        // rendering checks measure the fraction. Fragmented formulas retain
        // their existing fallback in this bounded batch migration.
        var literal = new DocxMathRow([new DocxMathRun(run, "p", null)]);
        TestAssert.True(DocxMathRowLayout.TryMeasure(literal, _ => font, default, out var measured, out _), "Portable literal measurement is available.");
        var prepared = new DocxPreparedMathParagraph(equation, TestAssert.NotNull(measured), 12d, 24d);
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(source,
            new DocxTests.FamilyWidthTextMeasurer(), default,
            mathEquations: new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance) { [equation] = prepared });
        var canonical = layout.RelatedStories.Single(s => ReferenceEquals(s.Story, originalStory));
        var mathLine = canonical.TextLines.Single(line => line.InlineMath.Count != 0);
        TestAssert.Equal(1, mathLine.InlineMath.Count);
        TestAssert.True(canonical.InlineImages.Single().LineFrame is null && mathLine.InlineFrame is not null,
            "An ordered paragraph owns its picture through the measured inline frame, without a second legacy image-growth charge.");
        TestAssert.True(DocxLayoutEngine.HasParagraphNoteFlowStory(canonical, default),
            "Every measured atomic equation and picture belongs to its original source paragraph.");
        foreach (var invalid in new[]
        {
            mathLine with { InlineMath = [] },
            mathLine with { InlineMath = [.. mathLine.InlineMath, mathLine.InlineMath[0]] },
            mathLine with { InlineMath = [new(prepared with { Equation = equation with { SourceRunIndex = 999 } }, 0d)] },
            mathLine with { InlineFrame = new(double.NaN, 24d) }
        })
        {
            var candidate = canonical with { TextLines = canonical.TextLines.Select(line => ReferenceEquals(line, mathLine) ? invalid : line).ToArray() };
            TestAssert.True(!DocxLayoutEngine.HasParagraphNoteFlowStory(candidate, default),
                "Missing, duplicated, foreign or unmeasured mathematical owners retain the fallback.");
        }
        var missingPictureOwner = canonical with { InlineImages = [canonical.InlineImages.Single() with { SourceLineIndex = null }] };
        TestAssert.True(!DocxLayoutEngine.HasParagraphNoteFlowStory(missingPictureOwner, default), "An unowned picture cannot borrow a measured mathematical frame.");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.HasParagraphNoteFlowStory(canonical, cancelled.Token));
    }
}
