using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxIntrinsicFrameEndsOwnWrappedRowsAndTrimmedWhitespace()
    {
        foreach (bool picture in new[] { false, true })
        {
            var atoms = MakeIntrinsicFrameAtoms(picture);
            var text = IntrinsicFrameText("T", 40d);
            var space = IntrinsicFrameText(" ", 5d);
            atoms = atoms with { Fragments = [text, atoms.Fragments.Single(), space] };
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 60d, 60d, default, out var rows), "The frame wraps atomically.");
            var lines = TestAssert.NotNull(rows);
            TestAssert.Equal(2, lines.Count);
            TestAssert.Equal(12d, lines[0].OccupiedHeight);
            TestAssert.Equal(40d, lines[1].OccupiedHeight);
            TestAssert.Equal(40d, lines[1].BaselineInset);
            TestAssert.Equal(1, lines[1].Fragments.Count);
            TestAssert.True(lines[1].EndCursor.IsComplete(atoms), "Trimmed source whitespace is consumed once.");
            var cursor = new DocxInlineSourceCursor();
            foreach (var expected in lines)
            {
                TestAssert.True(DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, cursor, 60d, 60d, default, out var next),
                    "Resumable placement shares the whole-paragraph frame.");
                var actual = TestAssert.NotNull(next);
                TestAssert.Equal(expected.OccupiedHeight, actual.OccupiedHeight);
                TestAssert.Equal(expected.BaselineInset, actual.BaselineInset);
                cursor = actual.EndCursor;
            }
            TestAssert.True(cursor.IsComplete(atoms), "Both wrapping paths consume the closing paragraph mark.");
        }
    }

    public static void DocxIntrinsicFrameTextRetainsPhysicalDescentAtEitherScale()
    {
        foreach (bool picture in new[] { false, true })
        foreach (double scale in new[] { 1d, .75d })
        {
            var atoms = MakeIntrinsicFrameAtoms(picture);
            atoms = atoms with { Fragments = [IntrinsicFrameText("T", 10d), atoms.Fragments.Single()] };
            atoms = atoms with { Scale = scale, BaselineInset = 10d * scale, LineHeight = 12d * scale,
                Fragments = atoms.Fragments.Select(fragment => fragment with { Advance = fragment.Advance * scale,
                    BaselineInset = fragment.BaselineInset * scale, OccupiedHeight = fragment.OccupiedHeight * scale }).ToArray() };
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 100d, 100d, default, out var rows,
                physicalTextBaselineInset: 7d * scale), "Text and the intrinsic frame share a physical row.");
            var line = TestAssert.NotNull(rows).Single();
            TestAssert.Equal(45d * scale, line.OccupiedHeight);
            TestAssert.Equal(40d * scale, line.BaselineInset);
            TestAssert.Equal(10d * scale, line.TextBaselineInset);
        }
    }

    public static void DocxIntrinsicFrameEndsPreserveShortMinimumAndExactSlots()
    {
        var atoms = MakeIntrinsicFrameAtoms(false);
        var frame = atoms.Fragments.Single();
        var shortFrame = atoms with { Fragments = [frame with { BaselineInset = 8d, OccupiedHeight = 8d }] };
        TestAssert.Equal(12d, Wrap(shortFrame).OccupiedHeight);
        var minimum = atoms with { Paragraph = atoms.Paragraph with { LineSpacingPoints = 60d,
            Spacing = atoms.Paragraph.Spacing with { LineRuleValue = "atLeast" } } };
        TestAssert.Equal(60d, Wrap(minimum).OccupiedHeight);
        TestAssert.Equal(60d, Wrap(minimum).BaselineInset);
        var exact = atoms with { Paragraph = atoms.Paragraph with { LineSpacingPoints = 12d,
            Spacing = atoms.Paragraph.Spacing with { LineRuleValue = "exact" } } };
        TestAssert.Equal(12d, Wrap(exact).OccupiedHeight);
        TestAssert.Equal(10d, Wrap(exact).BaselineInset);
        TestAssert.Equal(40d, Wrap(exact).Fragments.Single().OccupiedHeight);

        static DocxMeasuredInlineLine Wrap(DocxMeasuredInlineAtoms input)
        {
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(input, 100d, 100d, default, out var rows), "The authored slot wraps.");
            return TestAssert.NotNull(rows).Single();
        }
    }

    public static void DocxIntrinsicFrameEndsRetainDominantMathematicalFrames()
    {
        foreach (double mathAscent in new[] { 20d, 50d })
        {
            var atoms = MakeIntrinsicFrameAtoms(false);
            var equation = atoms.Paragraph.Equations.Single();
            var prepared = new DocxPreparedMathParagraph(equation, new(18d, mathAscent, 2d, []), mathAscent, mathAscent + 2d)
                { UsesDisplayStyle = false };
            atoms = atoms with { Fragments = [IntrinsicFrameText("T", 10d), atoms.Fragments.Single(),
                new(new DocxInlineEquation(equation), 18d, mathAscent, mathAscent + 2d, [], prepared)] };
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 100d, 100d, default, out var rows,
                physicalTextBaselineInset: 7d), "Prepared mathematical rows retain their accepted shared strut.");
            var line = TestAssert.NotNull(rows).Single();
            TestAssert.Equal(Math.Max(40d, mathAscent), line.BaselineInset);
            TestAssert.Equal(Math.Max(40d, mathAscent) + 2d, line.OccupiedHeight);
            TestAssert.True(ReferenceEquals(prepared, line.Fragments.Single(fragment => fragment.Math is not null).Math),
                "The mathematical frame survives beside both larger and smaller intrinsic objects without another descent.");
        }
    }

    public static void DocxIntrinsicFrameTrailingBreakRetainsParagraphMarkAndCancellation()
    {
        var atoms = MakeIntrinsicFrameAtoms(false);
        var run = new DocxTextRun("\n", 12d, null, false, false, false, null, null);
        atoms = atoms with { Fragments = [atoms.Fragments.Single(), new(new DocxInlineBreak(run, 0), 0d, 10d, 12d, [])] };
        TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 100d, 100d, default, out var rows), "The trailing hard break wraps.");
        var lines = TestAssert.NotNull(rows);
        TestAssert.Equal(2, lines.Count);
        TestAssert.Equal(40d, lines[0].OccupiedHeight);
        TestAssert.Equal(12d, lines[1].OccupiedHeight);
        TestAssert.Equal(0, lines[1].Fragments.Count);
        TestAssert.True(lines[1].EndCursor.IsComplete(atoms), "The empty closing row consumes the paragraph mark.");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms,
            100d, 100d, cancelled.Token, out _));
    }

    public static void DocxIntrinsicFrameEndsRetainIndependentWordContentAndOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent frame-end geometry uses installed Windows fonts.");
        CheckInlineBoxMath(null, "inline-frame-ends-office.json");
        CheckInlineBoxWordOrigins("inline-frame-ends-office.json");
        CheckInlineBoxFollowingOrigins("inline-frame-ends-office.json");
    }

    private static DocxMeasuredInlineAtoms MakeIntrinsicFrameAtoms(bool picture)
    {
        var paragraph = ReadSiblingFrameSource() with { LineSpacingPoints = null,
            Spacing = DocxParagraphSpacing.Empty, ListLabel = null };
        DocxInlineItem owner = picture
            ? new DocxInlinePicture(ReadMutation("InlineBoxMath/inline-sibling-mixed", OoxPdfDocxMarkupMode.Final,
                false, false).Paragraphs.Single(p => p.Images.Count != 0).Images.Single())
            : new DocxInlineBox(paragraph.InlineTextBoxes.Single());
        return new(paragraph, [new(owner, 50d, 40d, 40d, [])], new TrackingMeasurer(),
            12d, 10d, 12d, 1d, 36d, [], []);
    }

    private static DocxMeasuredInlineFragment IntrinsicFrameText(string value, double advance)
    {
        var run = new DocxTextRun(value, 12d, null, false, false, false, null, null);
        return new(new DocxInlineText(run, 0, value.Length), advance, 10d, 12d, [new(value, run, 0, 0)]);
    }
}
