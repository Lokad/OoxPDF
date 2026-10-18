using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMinimumMixedNoteFramesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent minimum-note frames require installed Windows fonts.");
        CheckMinimumMixedNoteFrames(null);
    }

    internal static void CheckMinimumMixedNoteFrames(IFontResolver? resolver)
    {
        foreach (string fixture in MinimumMixedNoteFixtures)
            CheckPictureNoteBatches(resolver, fixture);
    }

    private static readonly string[] MinimumMixedNoteFixtures =
        ["minimum-mixed-note-low-report", "minimum-mixed-note-large-report"];

    public static void DocxInlineCompositionMinimumNoteBatchesRequireMeasuredAuthoredFloors()
    {
        foreach (string fixture in MinimumMixedNoteFixtures)
        {
            var canonical = MeasureAtomicNoteBatchOwnershipControl(fixture);
            var line = canonical.TextLines.Single(candidate => candidate.InlineMath.Count != 0);
            double minimum = line.SourceParagraph!.LineSpacingPoints!.Value;
            TestAssert.True(line.InlineFrame is { } frame && frame.Height >= minimum &&
                DocxLayoutEngine.HasParagraphNoteFlowStory(canonical, default),
                "Authored minimums share the measured content-driven note frame.");
            var tooShort = line with { InlineFrame = new(minimum / 4d, minimum / 2d) };
            var inconsistent = canonical with { TextLines = canonical.TextLines.Select(candidate =>
                ReferenceEquals(candidate, line) ? tooShort : candidate).ToArray() };
            TestAssert.True(!DocxLayoutEngine.HasParagraphNoteFlowStory(inconsistent, default),
                "A finite frame below the authored minimum is insufficient source ownership.");

            var fixedStory = MeasureAtomicNoteBatchOwnershipControl(fixture, (xml, w, _) =>
                xml.Root!.Elements(w + "footnote").Single(note => (string?)note.Attribute(w + "id") == "2")
                    .Elements(w + "p").First().Element(w + "pPr")!.Element(w + "spacing")!
                    .SetAttributeValue(w + "lineRule", "exact"));
            TestAssert.True(!DocxLayoutEngine.HasParagraphNoteFlowStory(fixedStory, default),
                "Fixed boxes retain their clipping and placement fallback.");
            var legacyPictures = MeasureAtomicNoteBatchOwnershipControl(fixture, (xml, _, _) =>
            {
                XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
                xml.Descendants(math + "oMath").Remove();
            });
            TestAssert.True(!DocxLayoutEngine.HasParagraphNoteFlowStory(legacyPictures, default),
                "A legacy image extent does not establish a complete authored minimum line frame.");
        }
    }
}
