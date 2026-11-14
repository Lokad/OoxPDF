using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxConsumedMarkFormattingKeepsActualUniqueSourceOwners()
    {
        foreach (string story in new[] { "body", "cell", "header", "footer", "plain" })
        foreach (bool inserted in new[] { false, true })
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var document = ReadMarkPropertyMath(story, mode, inserted: inserted, deleted: true);
            var paragraph = MarkPropertyParagraph(document);
            bool joined = mode is OoxPdfDocxMarkupMode.Final or OoxPdfDocxMarkupMode.SimpleMarkup;
            TestAssert.Equal(joined ? 1 : 0, paragraph.ConsumedParagraphMarks.Count);
            TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(paragraph), "Actual marks own metadata in every view, independently of joining.");
            TestAssert.Equal(4, DocxBlockTraversal.EnumerateSourceParagraphs(document)
                .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance).Sum(p => DocxMathParagraphPolicy.GetPreparationEquations(p).Count));
            if (joined)
            {
                var mark = paragraph.ConsumedParagraphMarks.Single();
                TestAssert.Equal(inserted ? 3 : 2, mark.Revisions.Count);
                TestAssert.True(mark.Revisions.All(r => paragraph.Revisions.Any(owner => ReferenceEquals(owner, r))),
                    "The consumed deletion and one format snapshot retain actual shared source identities.");
                foreach (var invalid in new[]
                {
                    paragraph with { ConsumedParagraphMarks = [] },
                    paragraph with { ConsumedParagraphMarks = [mark with { SourceRunIndex = 0 }] },
                    paragraph with { ConsumedParagraphMarks = [mark with { Text = "visible" }] },
                    paragraph with { ConsumedParagraphMarks = [mark with { Revisions = mark.Revisions.Select(r =>
                        r.Kind == DocxRevisionKind.RunPropertiesChange ? r with { } : r).ToArray() }] }
                })
                    TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(invalid),
                        "Missing, misplaced, visible or copied consumed formatting cannot invent ownership.");
            }
            else
            {
                TestAssert.True(paragraph.HasDeletedParagraphMark,
                    "Original/Review retain the unconsumed boundary and all of its inline content.");
                var mark = paragraph.ParagraphMarkRun!;
                foreach (var invalid in new[]
                {
                    paragraph with { ParagraphMarkRun = null },
                    paragraph with { ParagraphMarkRun = mark with { SourceRunIndex = 0 } },
                    paragraph with { ParagraphMarkRun = mark with { Text = "visible" } },
                    paragraph with { ParagraphMarkRun = mark with { Revisions = mark.Revisions.Select(r => r with { }).ToArray() } }
                })
                    TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(invalid), "Unconsumed metadata also requires its actual source mark.");
            }
            var duplicate = MarkPropertyParagraph(ReadMarkPropertyMath(story, mode, duplicate: true, inserted: inserted, deleted: true));
            TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(duplicate), "Joining cannot make duplicate formatting history unique.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedSourceOwner(paragraph, token: canceled.Token));
        }
    }

    public static void DocxConsumedMarkFormattingPreservesMathPicturesAndNavigation()
    {
        foreach (string story in new[] { "body", "cell", "header", "footer", "plain" })
        foreach (bool inserted in new[] { false, true })
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = renderer.RenderPagesAsync(ReadMarkPropertyMath(story, mode, inserted: inserted, deleted: true), diagnostics.Add, default)
                .ToBlockingEnumerable().ToArray();
            var math = pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                .Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))))
                .ToArray();
            TestAssert.Equal(18, math.Length);
            TestAssert.Equal("x+2ya+b=cx+2ya+b=c",
                string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"), "Boundary revisions cannot discard adjacent supported equations.");
            TestAssert.Contains("share the same input series.", string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)));
            TestAssert.True(pages.Any(page => page.Images.Any(image => page.Content.Contains("/" + image.ResourceName + " Do", StringComparison.Ordinal))),
                "A consumed mark's formatting preserves its adjacent picture.");
            TestAssert.True(pages.SelectMany(page => page.Annotations).Any(), "The supported composition retains source navigation.");
            for (int page = 0; page < pages.Length; page++) Lokad.OoxPdf.Pdf.PdfContentValidator.ValidatePage(pages[page], page, default);
            if (mode != OoxPdfDocxMarkupMode.Final) continue;
            var control = renderer.RenderPagesAsync(ReadMarkPropertyMath(story, mode, baseline: true, inserted: inserted, deleted: true), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(control.Length, pages.Length);
            for (int page = 0; page < pages.Length; page++)
            {
                var expected = DocxCommentsTests.ReadEmbeddedGlyphTextShows(control[page], true).ToArray();
                var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[page], true).ToArray();
                TestAssert.True(expected.Select(s => (s.Text, s.X, s.Y, s.Size)).SequenceEqual(actual.Select(s => (s.Text, s.X, s.Y, s.Size))),
                    "Final source glyph text/geometry is unchanged by the consumed formatting history.");
                TestAssert.True(control[page].Annotations.Select(a => (a.Uri, a.Destination?.PageIndex))
                    .SequenceEqual(pages[page].Annotations.Select(a => (a.Uri, a.Destination?.PageIndex))), "Final navigation matches the complete causal control.");
            }
        }
    }
}
