using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxCombinedMarkRevisionsRetainTheirIndependentSourceOwners()
    {
        foreach (string story in new[] { "body", "cell", "header", "footer", "plain" })
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var document = ReadMarkPropertyMath(story, mode, inserted: true);
            var paragraph = MarkPropertyParagraph(document);
            var mark = paragraph.ParagraphMarkRun!;
            TestAssert.Equal(2, mark.Revisions.Count);
            TestAssert.True(mark.Revisions.All(revision => paragraph.Revisions.Any(item => ReferenceEquals(item, revision))),
                "Both insertion and formatting retain their actual source objects.");
            TestAssert.True(DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
                "An insertion cannot make the one actual formatting snapshot ambiguous: " + story + "/" + mode);
            TestAssert.Equal(4, DocxBlockTraversal.EnumerateSourceParagraphs(document)
                .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance).Sum(p => DocxMathParagraphPolicy.GetPreparationEquations(p).Count));
            var duplicate = MarkPropertyParagraph(ReadMarkPropertyMath(story, mode, duplicate: true, inserted: true));
            TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(duplicate, default),
                "Insertion cannot make duplicate formatting histories eligible.");
            foreach (var invalid in new[]
            {
                paragraph with { ParagraphMarkRun = null },
                paragraph with { ParagraphMarkRun = mark with { SourceRunIndex = 0 } },
                paragraph with { ParagraphMarkRun = mark with { Revisions = mark.Revisions.Select(r => r with { }).ToArray() } }
            })
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(invalid, default),
                    "Missing, misplaced or copied combined metadata cannot invent ownership.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, canceled.Token));
        }
    }

    public static void DocxCombinedMarkRevisionsKeepMathPicturesAndNavigation()
    {
        foreach (string story in new[] { "body", "cell", "header", "footer", "plain" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMarkPropertyMath(story, mode, inserted: true);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var pages = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            var math = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                .Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))))
                .ToArray();
            TestAssert.Equal(18, math.Length);
            TestAssert.Equal("x+2ya+b=cx+2ya+b=c", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Contains("share the same input series.", string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)));
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Combining actual mark roles cannot remove supported math.");
            TestAssert.True(pages.Any(page => page.Images.Any(image => page.Content.Contains("/" + image.ResourceName + " Do", StringComparison.Ordinal))),
                "The combined paragraph retains its adjacent picture.");
            TestAssert.True(pages.SelectMany(page => page.Annotations).Any(), "The composed content retains a real source link/bookmark target.");
            for (int page = 0; page < pages.Length; page++) Lokad.OoxPdf.Pdf.PdfContentValidator.ValidatePage(pages[page], page, default);
        }
    }
}
