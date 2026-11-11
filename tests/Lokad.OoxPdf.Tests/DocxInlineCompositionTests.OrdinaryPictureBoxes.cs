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
            int childCalls = measurer.ChildCalls;
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
            TestAssert.Equal(childCalls, measurer.ChildCalls);
        }
    }

    public static void DocxOrdinaryPictureBoxesKeepCancellationAndPurePictureFallback()
    {
        var source = ReadOrdinaryPictureBox("body");
        using var cancellation = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryMeasureOrderedInlineAtoms(source,
            null, new SiblingFrameMeasurer(cancellation), 11d, 10d, 16d, 1d, cancellation.Token, out _));
        var withoutBox = source with { InlineTextBoxes = [], InlineContent = TestAssert.NotNull(source.InlineContent)
            .Where(item => item is not DocxInlineBox).ToArray() };
        // The historical method name describes the old conservative fallback.
        // Removing a box does not remove its sibling content or source owners.
        // Automatic-flow containers share ordinary picture/prose measurement.
        TestAssert.True(withoutBox.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text)),
            "This admission witness contains both prose and a picture.");
        foreach (bool staticFallback in new[] { false, true })
        {
            bool measured = DocxLayoutEngine.TryMeasureOrderedInlineAtoms(withoutBox, null, new SiblingFrameMeasurer(),
                11d, 10d, 16d, 1d, default, out var atoms, allowStaticFallback: staticFallback);
            TestAssert.True(measured, "Owned pictures and prose share the ordered row contract.");
            var fragments = TestAssert.NotNull(atoms).Fragments;
            TestAssert.Equal(1, fragments.Count(fragment => fragment.Source is DocxInlinePicture picture &&
                ReferenceEquals(picture.Image, withoutBox.Images.Single())));
            TestAssert.True(fragments.All(fragment => fragment.TextBox is null), "The removed box supplies no frame.");
            string expected = string.Concat(TestAssert.NotNull(withoutBox.InlineContent).OfType<DocxInlineText>()
                .Select(text => text.Run.Text.Substring(text.Start, text.Length)));
            TestAssert.Equal(expected, string.Concat(fragments.SelectMany(fragment => fragment.Text).Select(span => span.Text)));
        }
        var proseOnly = withoutBox with { Images = [], InlineContent = TestAssert.NotNull(withoutBox.InlineContent)
            .Where(item => item is not DocxInlinePicture).ToArray() };
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(proseOnly, null, new SiblingFrameMeasurer(),
            11d, 10d, 16d, 1d, default, out _), "Ordinary prose keeps its existing text path.");
        var unknownImage = withoutBox.Images.Single() with { SourceRunIndex = -1 };
        var unknownOwner = withoutBox with { Images = [unknownImage], InlineContent = TestAssert.NotNull(withoutBox.InlineContent)
            .Select(item => item is DocxInlinePicture picture ? picture with { Image = unknownImage } : item).ToArray() };
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(unknownOwner, null, new SiblingFrameMeasurer(),
            11d, 10d, 16d, 1d, default, out _), "Unowned pictures retain their explicit fallback.");
        DocxParagraph DeclaredFlow(DocxParagraph value, string rule) => value with
        {
            LineSpacingPoints = 32d, LineSpacingFactor = 1d,
            Spacing = value.Spacing with { LineValue = "640", LineRuleValue = rule }
        };
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(DeclaredFlow(withoutBox, "atLeast"), null,
            new SiblingFrameMeasurer(), 11d, 10d, 32d, 1d, default, out _),
            "Content-driven ordinary rows share their measured source owners.");
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(
            DeclaredFlow(withoutBox with { InlineContent = null }, "atLeast"), null,
            new SiblingFrameMeasurer(), 11d, 10d, 32d, 1d, default, out _),
            "Reconstructed minimum-row owners retain their independently qualified fallback.");
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(DeclaredFlow(source, "atLeast"), null,
            new SiblingFrameMeasurer(), 11d, 10d, 32d, 1d, default, out _),
            "Box hosts retain their separately qualified automatic-flow eligibility.");
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(DeclaredFlow(withoutBox, "exact"), null,
            new SiblingFrameMeasurer(), 11d, 10d, 32d, 1d, default, out _),
            "Captured ordinary fixed-line pictures share the contained-slot rule.");
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(
            DeclaredFlow(withoutBox with { InlineContent = null }, "exact"), null,
            new SiblingFrameMeasurer(), 11d, 10d, 32d, 1d, default, out _),
            "Reconstructed fixed-line owners retain their independently qualified fallback.");
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
