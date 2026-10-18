using System.Text;
using System.Xml.Linq;
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
        var canonical = MeasureAtomicNoteBatchOwnershipControl("mixed-math-picture-note-report");
        var mathLine = canonical.TextLines.Single(line => line.InlineMath.Count != 0);
        var prepared = mathLine.InlineMath[0].Math;
        var equation = prepared.Equation;
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

    private static DocxRelatedStoryLayout MeasureAtomicNoteBatchOwnershipControl(string fixture,
        Action<XDocument, XNamespace, XNamespace>? change = null)
    {
        var source = change is null
            ? DocxTests.ReadDocx(NestedMathCasePath(fixture + ".docx"), OoxPdfDocxMarkupMode.Final)
            : ReadMutation("CompositionProbes/" + fixture, OoxPdfDocxMarkupMode.Final, false, false, change,
                partOverride: "word/footnotes.xml");
        var story = source.RelatedStories.Single(candidate => candidate.Kind == DocxRelatedStoryKind.Footnote && candidate.Id == "2");
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? equations = null;
        if (story.Paragraphs[0].Equations.SingleOrDefault() is { } equation)
        {
            // One atomic portable measurement preserves immutable source owners;
            // independent rendering checks measure the actual source fraction.
            var run = new DocxTextRun("A", 10d, null, false, false, false, null, "TestFont");
            var font = new PdfFontResource("Eq1", PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(),
                run.Text.EnumerateRunes().Select(rune => rune.Value), default));
            TestAssert.True(DocxMathRowLayout.TryMeasure(new DocxMathRow([new DocxMathRun(run, "p", null)]),
                _ => font, default, out var measured, out _), "Portable literal measurement is available.");
            equations = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
                { [equation] = new(equation, TestAssert.NotNull(measured), 12d, 24d) };
        }
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: equations);
        return layout.RelatedStories.Single(candidate => ReferenceEquals(candidate.Story, story));
    }
}
