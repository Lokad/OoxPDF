using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxOwnedBoxParagraphsRetainSourceOwnershipWithoutParentMath()
    {
        foreach (string id in new[] { "inline-box-boxed-text-valid", "inline-box-boxed-equation-valid" })
        {
            var source = ReadOwnedBoxParagraph(id);
            TestAssert.Equal(0, source.Equations.Count);
            TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(source), "A box retains its source owner independently of parent mathematics.");
            TestAssert.True(DocxMathParagraphPolicy.NeedsOrderedInlineLayout(source), "The existing ordered path owns the complete frame.");
            TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(source).Count);
            TestAssert.True(ReferenceEquals(source.InlineTextBoxes.Single(),
                TestAssert.NotNull(source.InlineContent).OfType<DocxInlineBox>().Single().Box), "The parsed frame is not reconstructed from child text.");
            var replaced = source with { InlineTextBoxes = source.InlineTextBoxes.ToArray() };
            TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(replaced), "Replacing owners invalidates the source view.");
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedSourceOwner(source, token: cancelled.Token));
        }
    }

    public static void DocxOwnedBoxParagraphsReuseCompleteChildrenAtEitherScale()
    {
        foreach (string id in new[] { "inline-box-boxed-text-valid", "inline-box-boxed-equation-valid" })
        foreach (double scale in new[] { 1d, .75d })
        {
            var source = ReadOwnedBoxParagraph(id);
            var prepared = PrepareSiblingFrameMath(source);
            var measurer = new SiblingFrameMeasurer();
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(source, prepared, measurer,
                11d, 10d * scale, 16d * scale, scale, default, out var atoms), "Automatic source-owned box paragraphs measure without a synthetic equation.");
            var measured = TestAssert.NotNull(atoms);
            var fragment = measured.Fragments.Single(item => item.TextBox is not null);
            var frame = TestAssert.NotNull(fragment.TextBox);
            int calls = measurer.Calls;
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(measured, 300d, 300d, default, out var rows), "The measured frame wraps once.");
            var line = TestAssert.NotNull(rows).Single();
            TestAssert.Equal(60d * scale, line.OccupiedHeight);
            TestAssert.True(ReferenceEquals(frame, line.Fragments.Single(item => item.TextBox is not null).TextBox), "Wrapping reuses the complete measured child.");
            TestAssert.Equal(calls, measurer.Calls);
            TestAssert.Equal(0, line.Fragments.Count(item => item.Math is not null));
            TestAssert.Equal(id == "inline-box-boxed-equation-valid" ? 1 : 0, frame.TextLines.Sum(child => child.InlineMath.Count));
        }
    }

    public static void DocxOwnedBoxParagraphsRetainExplicitSlotRevisionAndNestedFallbacks()
    {
        var source = ReadOwnedBoxParagraph("inline-box-boxed-text-valid");
        foreach (string rule in new[] { "exact", "atLeast" })
        {
            var explicitSlot = source with { LineSpacingPoints = 12d, Spacing = source.Spacing with { LineRuleValue = rule } };
            TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(explicitSlot, null, new TrackingMeasurer(),
                11d, 10d, 12d, 1d, default, out _), "Unqualified explicit slots retain their existing box placement.");
        }
        var original = source.InlineTextBoxes.Single();
        var nested = original with { BodyElements = [new DocxParagraphElement(source)] };
        var unsupported = source with { InlineTextBoxes = [nested], InlineContent = TestAssert.NotNull(source.InlineContent)
            .Select(item => item is DocxInlineBox ? (DocxInlineItem)new DocxInlineBox(nested) : item).ToArray() };
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(unsupported, null, new TrackingMeasurer(),
            11d, 10d, 16d, 1d, default, out _), "Nested frames retain their existing staged path.");
        var revised = source with { InlineContent = null };
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(revised, null, new TrackingMeasurer(),
            11d, 10d, 16d, 1d, default, out _), "A missing trusted source view cannot silently reconstruct revised owners.");
        var inserted = original with { Revisions = [new(DocxRevisionKind.Insertion, "904", "Public Reviewer", null, "ins", null, [])] };
        var revisedOwner = source with { InlineTextBoxes = [inserted], InlineContent = TestAssert.NotNull(source.InlineContent)
            .Select(item => item is DocxInlineBox ? (DocxInlineItem)new DocxInlineBox(inserted) : item).ToArray() };
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(revisedOwner, null, new TrackingMeasurer(),
            11d, 10d, 16d, 1d, default, out _), "Revised box owners retain their explicit fallback.");
    }

    public static void DocxOwnedBoxParagraphsShareBodyAndCellAdmission()
    {
        var source = ReadOwnedBoxParagraph("inline-box-boxed-text-valid") with
            { SpacingBeforePoints = 0d, SpacingAfterPoints = 0d, Spacing = DocxParagraphSpacing.Empty };
        var following = DocxTests.CreateDocxLayoutParagraph("Following", 12d, 12d);
        var paragraphs = new[] { source, source with { }, following };
        foreach (bool cell in new[] { false, true })
        {
            var table = new DocxTable(null, [228d], [new DocxTableRow([
                new DocxTableCell(string.Empty, paragraphs, null, null, null, null, [], DocxTableCellMargins.Empty)], null)]);
            DocxDocument document = new(300d, 205d, 36d, 36d, 36d, 36d, DocxPageSettings.Empty, [], [], [],
                cell ? [new DocxTableElement(table)] : paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray(),
                [], cell ? [table] : []);
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default);
            TestAssert.Equal(1, layout.Pages.Count);
            var row = layout.Pages[0].Items.OfType<DocxTableRowLayout>().SingleOrDefault();
            var boxes = cell ? row!.Cells.Single().InlineTextBoxes : layout.Pages[0].Items.OfType<DocxInlineTextBoxLayout>().ToArray();
            TestAssert.Equal(2, boxes.Count);
            TestAssert.True(boxes.All(box => box.SourceLineIndex == 0), "Both containers publish measured source-line ownership.");
            TestAssert.True(Math.Abs(boxes[0].BoxTop - boxes[0].BoxHeight - boxes[1].BoxTop) < .001d,
                "Two complete60pt frames share the available page without duplicate struts.");
            TestAssert.True(boxes.All(box => box.BoxTop - box.BoxHeight >= 36d), "Admitted frame bounds stay above the page bottom.");
        }
    }

    public static void DocxOwnedBoxParagraphsMatchIndependentWordFramesAndFollowingText()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent owned-box geometry uses installed Windows fonts.");
        CheckInlineBoxMath(null, "inline-box-owned-paragraphs-office.json");
        CheckInlineBoxWordOrigins("inline-box-owned-paragraphs-office.json");
        CheckInlineBoxFollowingOrigins("inline-box-owned-paragraphs-office.json");
    }

    private static DocxParagraph ReadOwnedBoxParagraph(string id) =>
        ReadMutation("InlineBoxMath/" + id, OoxPdfDocxMarkupMode.Final, false, false)
            .Paragraphs.Single(paragraph => paragraph.InlineTextBoxes.Count != 0);
}
