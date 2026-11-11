using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxIgnoredInlineBoxesKeepMeasuredSiblingSourceIndices()
    {
        var source = ReadSiblingFrameSource(); var prepared = PrepareSiblingFrameMath(source);
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(source, prepared, new SiblingFrameMeasurer(),
            11d, 10d, 16d, 1d, default, out var atoms, includeInlineTextBoxes: false),
            "A container can measure supported items while retaining its ignored-box policy.");
        var items = source.InlineContent!.ToArray();
        TestAssert.True(atoms!.Fragments.All(fragment => fragment.Source is not DocxInlineBox && fragment.TextBox is null),
            "Ignored frames create no placement or source claim.");
        foreach (var fragment in atoms.Fragments.Where(fragment => fragment.Math is not null))
            TestAssert.Equal(Array.FindIndex(items, item => ReferenceEquals(item, fragment.Source)), fragment.SourceItemIndex);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryMeasureOrderedInlineAtoms(source, prepared,
            new SiblingFrameMeasurer(), 11d, 10d, 16d, 1d, cancellation.Token, out _, includeInlineTextBoxes: false));
    }

    public static void DocxInlineBoxFallbackRetainsInvalidFrameAndOwnerGuards()
    {
        var source = ReadSiblingFrameSource(); var prepared = PrepareSiblingFrameMath(source);
        var original = source.InlineTextBoxes.Single();
        foreach (var box in new[]
        {
            original with { ExtentCxValue = "0" }, original with { ExtentCyValue = "-1" },
            original with { ExtentCxValue = "not-an-extent" }, original with { BodyElements = [] },
            original with { TextBoxInsetLeftValue = original.ExtentCxValue }, original with { SourceRunIndex = -1 }
        })
        {
            var changed = source with { InlineTextBoxes = [box], InlineContent = source.InlineContent!.Select(item =>
                item is DocxInlineBox ? (DocxInlineItem)new DocxInlineBox(box) : item).ToArray() };
            TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(changed, prepared, new SiblingFrameMeasurer(),
                11d, 10d, 16d, 1d, default, out _), "A child fallback cannot admit an invalid or foreign frame.");
        }
    }

}
