using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxOrdinaryPictureBoxesRetainOrderedOwnersAndReuseChildren()
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (double scale in new[] { 1d, .75d })
        {
            var source = ReadOrdinaryPictureBox(container);
            TestAssert.Equal(0, source.Equations.Count);
            var items = TestAssert.NotNull(source.InlineContent);
            TestAssert.True(ReferenceEquals(source.Images.Single(), items.OfType<DocxInlinePicture>().Single().Image),
                "The source picture owner survives removal of parent mathematics.");
            TestAssert.True(ReferenceEquals(source.InlineTextBoxes.Single(), items.OfType<DocxInlineBox>().Single().Box),
                "The complete textbox keeps its independent source owner.");
            var measurer = new SiblingFrameMeasurer();
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(source, null, measurer,
                11d, 10d * scale, 16d * scale, scale, default, out var atoms),
                "Ordinary pictures and boxes share ordered measurement without a parent equation.");
            var measured = TestAssert.NotNull(atoms);
            var frame = TestAssert.NotNull(measured.Fragments.Single(fragment => fragment.TextBox is not null).TextBox);
            int calls = measurer.Calls;
            foreach (double width in new[] { 450d, 300d, 220d })
            {
                TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(measured, width, width, default, out var rows),
                    "The container wraps complete source frames.");
                var fragments = TestAssert.NotNull(rows).SelectMany(row => row.Fragments).ToArray();
                TestAssert.Equal(1, fragments.Count(fragment => ReferenceEquals(frame, fragment.TextBox)));
                TestAssert.Equal(1, fragments.Count(fragment => fragment.Source is DocxInlinePicture));
                TestAssert.True(Array.FindIndex(fragments, fragment => fragment.Source is DocxInlinePicture) <
                    Array.FindIndex(fragments, fragment => fragment.TextBox is not null), "The image precedes its sibling box.");
                TestAssert.Equal(0, fragments.Count(fragment => fragment.Math is not null));
            }
            TestAssert.Equal(calls, measurer.Calls);
        }
    }

    public static void DocxOrdinaryPictureBoxesKeepCancellationAndPurePictureFallback()
    {
        var source = ReadOrdinaryPictureBox("body");
        using var cancellation = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryMeasureOrderedInlineAtoms(source,
            null, new SiblingFrameMeasurer(cancellation), 11d, 10d, 16d, 1d, cancellation.Token, out _));
        var pictureOnly = source with { InlineTextBoxes = [], InlineContent = TestAssert.NotNull(source.InlineContent)
            .Where(item => item is not DocxInlineBox).ToArray() };
        foreach (bool staticFallback in new[] { false, true })
            TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(pictureOnly, null, new SiblingFrameMeasurer(),
                11d, 10d, 16d, 1d, default, out _, allowStaticFallback: staticFallback),
                "Pure-picture paragraphs retain their established placement path.");
    }

    public static void DocxOrdinaryPictureBoxesMatchIndependentWordOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent ordinary picture/box origins use installed Windows fonts.");
        const string golden = "inline-box-ordinary-pictures-office.json";
        CheckInlineBoxMath(null, golden);
        CheckInlineBoxWordOrigins(golden);
        CheckInlineBoxFollowingOrigins(golden);
        CheckInlineBoxProseOrigins(golden);
    }

    private static DocxParagraph ReadOrdinaryPictureBox(string container) =>
        DocxBlockTraversal.EnumerateSourceParagraphs(ReadMutation("InlineBoxMath/inline-box-picture-" + container,
            OoxPdfDocxMarkupMode.Final, false, false)).Single(paragraph => paragraph.InlineTextBoxes.Count != 0);
}
