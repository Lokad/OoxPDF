using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingInsertedInlineBoxesKeepChildAndSiblingEquations()
    {
        foreach (var (fixture, part) in new (string, string?)[]
        {
            ("inline-box-body-equation-inline-box-valid", null), ("inline-sibling-mixed", null),
            ("inline-box-math-cell", null), ("inline-box-math-header", "word/inline-box-header.xml")
        })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var control = DocxInlineCompositionTests.ReadMutation("InlineBoxMath/" + fixture, mode, false, false, partOverride: part);
            var inserted = DocxInlineCompositionTests.ReadInsertedBoxSource(fixture, mode, part);
            var expected = renderer.RenderPagesAsync(control, null, default).ToBlockingEnumerable().Single();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(inserted, diagnostics.Add, default).ToBlockingEnumerable().Single();
            var before = MathShows(expected); var after = MathShows(actual);
            TestAssert.Equal(fixture == "inline-sibling-mixed" ? 8 : 4, after.Length);
            TestAssert.Equal(before.Length, after.Length);
            for (int i = 0; i < after.Length; i++)
            {
                TestAssert.True(after[i].Text == before[i].Text && after[i].Size == before[i].Size,
                    "Actual box insertion metadata preserves every prepared mathematical glyph.");
                if (mode == OoxPdfDocxMarkupMode.Final)
                    TestAssert.True(Math.Abs(after[i].X - before[i].X) < .001d && Math.Abs(after[i].Y - before[i].Y) < .001d,
                        "Resolving an insertion in Final preserves the original supported formula frame.");
            }
            string text = DocxPageFootnoteNumberingTests.DecodeText(actual).Normalize(NormalizationForm.FormKC);
            TestAssert.Contains("Before inline", text); TestAssert.Contains("after equation.", text);
            TestAssert.Equal(expected.Images.Count, actual.Images.Count);
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Owned insertions do not report painted equations as ignored.");
            PdfContentValidator.ValidatePage(actual, 0, default);
        }

        static (string Text, double Size, double X, double Y)[] MathShows(PdfPage page) =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true)
                .Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                    ReferenceEquals(font.Font, show.Font))).Select(show => (show.Text,show.Size,show.X,show.Y)).ToArray();
    }
}
