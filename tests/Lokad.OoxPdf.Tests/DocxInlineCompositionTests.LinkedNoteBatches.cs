using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionLinkedMixedNoteFramesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent linked-note frames require installed Windows fonts.");
        CheckPictureNoteBatches(null, "linked-mixed-note-report");
    }

    public static void DocxInlineCompositionLinkedNoteFramesRequireOwnedMeasuredSourceRanges()
    {
        var canonical = MeasureAtomicNoteBatchOwnershipControl("linked-mixed-note-report");
        var line = canonical.TextLines.Single(candidate => candidate.InlineMath.Count != 0);
        var paragraph = line.SourceParagraph!;
        var link = paragraph.Hyperlinks.Single();
        TestAssert.True(DocxLayoutEngine.HasParagraphNoteFlowStory(canonical, default),
            "A hyperlink annotates its existing measured content without changing paragraph placement.");
        foreach (var invalid in new[]
        {
            link with { StartInlineItemIndex = null },
            link with { StartInlineItemIndex = -1 },
            link with { EndInlineItemIndex = paragraph.InlineContent!.Count + 1 },
            link with { EndInlineItemIndex = link.StartInlineItemIndex - 1 },
            link with { TextRunCount = int.MaxValue }
        })
        {
            var changed = paragraph with { Hyperlinks = [invalid] };
            TestAssert.True(!DocxInlineContent.HasValidHyperlinkSourceRanges(changed, default),
                "Incomplete or foreign annotation ranges retain their source fallback.");
        }
        var unmeasured = canonical with { TextLines = canonical.TextLines.Select(candidate =>
            ReferenceEquals(candidate, line) ? line with { InlineFrame = null } : candidate).ToArray() };
        TestAssert.True(!DocxLayoutEngine.HasParagraphNoteFlowStory(unmeasured, default),
            "Annotations cannot establish an unmeasured placement frame.");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() =>
            DocxInlineContent.HasValidHyperlinkSourceRanges(paragraph, cancelled.Token));
    }
}
