using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionSpanningNoteMathMatchesIndependentContentAndPageGeometry()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent compound page geometry uses installed Windows fonts.");
        CheckSpanningPageNotes(null, requireIndependentGeometry: true, referenceFile: "spanning-note-math-office.json");
    }

    public static void DocxInlineCompositionSpanningNoteMathIgnoresNeutralSourceTransformations()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native compound transformations use installed Windows fonts.");
        CheckSpanningNoteMathTransformations(null);
    }

    internal static void CheckSpanningNoteMathTransformations(IFontResolver? resolver)
    {
        foreach (string typography in new[] { "calibri-12-zero", "arial-11-zero" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = "CompositionProbes/spanning-note-math-" + typography;
            var source = ReadMutation(fixture, mode, false, false);
            var transformed = ReadMutation(fixture, mode, true, true);
            foreach (var document in new[] { source, transformed })
            {
                var paragraph = document.Paragraphs.Single();
                TestAssert.Equal(8, paragraph.Equations.Count);
                TestAssert.Equal(8, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
                TestAssert.Equal(4, paragraph.Images.Count);
                TestAssert.Equal(4, paragraph.InlineReferences.Count);
                TestAssert.True(paragraph.BookmarkAnchors.Count >= 4 && paragraph.InlineContent is not null,
                    "The compound source owns all bookmarks, objects and references in one ordered paragraph.");
            }
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderBlankPages(source, diagnostics.Add, default).ToArray();
            var after = renderer.RenderBlankPages(transformed, diagnostics.Add, default).ToArray();
            TestAssert.True(before.Length > 1, "The compound paragraph crosses a physical page boundary.");
            TestAssert.Equal(before.Length, after.Length);
            for (int index = 0; index < before.Length; index++)
            {
                TestAssert.Equal(DocxPageFootnoteNumberingTests.DecodeText(before[index]).Normalize(NormalizationForm.FormKC),
                    DocxPageFootnoteNumberingTests.DecodeText(after[index]).Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(before[index].Annotations.Count, after[index].Annotations.Count);
                for (int link = 0; link < before[index].Annotations.Count; link++)
                {
                    var original = before[index].Annotations[link];
                    var retained = after[index].Annotations[link];
                    TestAssert.Equal(original.Destination, retained.Destination);
                    TestAssert.True(Math.Abs(original.X - retained.X) < .001d && Math.Abs(original.Y - retained.Y) < .001d &&
                        Math.Abs(original.Width - retained.Width) < .001d && Math.Abs(original.Height - retained.Height) < .001d,
                        "Neutral source changes retain actual footnote hit rectangles and targets.");
                }
                var originalFonts = before[index].Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                    .Select(resource => resource.Font).ToHashSet();
                var retainedFonts = after[index].Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                    .Select(resource => resource.Font).ToHashSet();
                var originalMath = DocxCommentsTests.ReadEmbeddedGlyphTextShows(before[index], usePageCoordinates: true)
                    .Where(show => originalFonts.Contains(show.Font)).ToArray();
                var retainedMath = DocxCommentsTests.ReadEmbeddedGlyphTextShows(after[index], usePageCoordinates: true)
                    .Where(show => retainedFonts.Contains(show.Font)).ToArray();
                TestAssert.Equal(originalMath.Length, retainedMath.Length);
                for (int glyph = 0; glyph < originalMath.Length; glyph++)
                    TestAssert.True(originalMath[glyph].Text == retainedMath[glyph].Text &&
                        Math.Abs(originalMath[glyph].X - retainedMath[glyph].X) < .001d &&
                        Math.Abs(originalMath[glyph].Y - retainedMath[glyph].Y) < .001d,
                        "Bookmarks and unchanged-format run splitting retain placed mathematical glyphs.");
            }
            TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Neutral source transformations cannot turn supported mathematics into ignored content.");
        }
    }
}
