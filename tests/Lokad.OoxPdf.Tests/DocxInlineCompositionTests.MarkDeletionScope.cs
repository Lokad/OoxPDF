using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxParagraphMarkDeletionKeepsEquationRevisionScope()
    {
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            const string fixture = "CompositionProbes/equation-revision-body-del";
            var baseline = ReadMutation(fixture, mode, false, false);
            var marked = ReadMutation(fixture, mode, false, false, (xml, w, _) =>
            {
                var paragraph = xml.Root!.Element(w + "body")!.Elements(w + "p").Last();
                var properties = paragraph.Element(w + "pPr");
                if (properties is null) { properties = new XElement(w + "pPr"); paragraph.AddFirst(properties); }
                var mark = properties.Element(w + "rPr");
                if (mark is null) { mark = new XElement(w + "rPr"); properties.Add(mark); }
                mark.Add(new XElement(w + "del", new XAttribute(w + "id", "990"), new XAttribute(w + "author", "Mark Reviewer")));
            });
            var last = marked.Paragraphs.Last();
            TestAssert.True(last.HasDeletedParagraphMark && last.ConsumedParagraphMarks.Count == 0 &&
                DocxInlineContent.HasSupportedSourceOwner(last), "An unjoined terminal mark owns only its boundary revision.");
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var diagnostics = new List<OoxPdfDiagnostic>();
            var expected = renderer.RenderPagesAsync(baseline, null, default).ToBlockingEnumerable().ToArray();
            var actual = renderer.RenderPagesAsync(marked, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            static string MathText(IEnumerable<Lokad.OoxPdf.Pdf.PdfPage> pages) => string.Concat(pages.SelectMany(page =>
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(show => page.Fonts.Any(font =>
                    font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))))
                .Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
            string text = MathText(expected);
            TestAssert.Equal(mode is OoxPdfDocxMarkupMode.Final or OoxPdfDocxMarkupMode.SimpleMarkup ? 14 : 18, text.Length);
            TestAssert.Equal(text, MathText(actual));
            TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"), "Boundary deletion cannot hide a separate supported equation.");
            TestAssert.True(actual.Any(page => page.Images.Any(image => page.Content.Contains("/" + image.ResourceName + " Do", StringComparison.Ordinal))),
                "Separating boundary/content revisions preserves the picture.");
            for (int page = 0; page < actual.Length; page++) Lokad.OoxPdf.Pdf.PdfContentValidator.ValidatePage(actual[page], page, default);
        }
    }
}
