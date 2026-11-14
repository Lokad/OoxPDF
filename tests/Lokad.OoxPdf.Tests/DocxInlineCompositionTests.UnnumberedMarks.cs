using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxUnnumberedInsertedMarksRetainActualRevisionOwners()
    {
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (string numbering in new[] { "absent", "0", "999" })
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var document = ReadUnnumberedInsertedMark(story, numbering, mode);
            var paragraphs = DocxBlockTraversal.EnumerateSourceParagraphs(document)
                .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance).ToArray();
            var first = paragraphs.First(p => p.Equations.Count != 0);
            TestAssert.True(first.ListLabel is null && first.ParagraphMarkRun is not null,
                "Disabling numbering preserves the captured paragraph mark: " + story + "/" + numbering);
            var insertion = first.Revisions.Single(r => r.Kind == DocxRevisionKind.Insertion);
            TestAssert.True(first.ParagraphMarkRun!.Revisions.Any(r => ReferenceEquals(r, insertion)) &&
                DocxInlineContent.HasSupportedSourceOwner(first), "The actual mark owns its source insertion without a list marker.");
            TestAssert.Equal(4, paragraphs.Sum(p => DocxMathParagraphPolicy.GetPreparationEquations(p).Count));
            TestAssert.True(paragraphs.Where(p => p.ListLabel is not null).Select(p => p.ListLabel!.Text)
                .SequenceEqual(new[] { "1.", "2." }), "An absent/disabled/unresolved first list reference does not advance the counter.");
        }
    }

    public static void DocxUnnumberedInsertedMarksKeepMathPicturesAndProse()
    {
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (string numbering in new[] { "absent", "0", "999" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadUnnumberedInsertedMark(story, numbering, mode);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            var math = pages.SelectMany(page =>
            {
                var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                    .Select(font => font.Font).ToHashSet();
                return DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true)
                    .Where(show => fonts.Contains(show.Font)).ToArray();
            }).ToArray();
            TestAssert.Equal(18, math.Length);
            string text = string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("x+2y", StringSplitOptions.None).Length - 1);
            TestAssert.Equal(2, text.Split("a+b=c", StringSplitOptions.None).Length - 1);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Removing only numbering cannot remove a supported fraction: " + story + "/" + numbering + "/" + mode);
            TestAssert.Contains("share the same input series.", string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)));
            TestAssert.True(pages.Any(page => page.Images.Any(image => page.Content.Contains("/" + image.ResourceName + " Do", StringComparison.Ordinal))),
                "The same paragraph composition preserves its adjacent picture.");
            for (int page = 0; page < pages.Length; page++) PdfContentValidator.ValidatePage(pages[page], page, default);
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(document, null, canceled.Token).ToBlockingEnumerable().ToArray());
        }
    }

    private static DocxDocument ReadUnnumberedInsertedMark(string story, string numbering, OoxPdfDocxMarkupMode mode) =>
        ReadMutation("CompositionProbes/numbering-mark-math-" + story, mode, false, false, (xml, w, _) =>
        {
            var properties = xml.Descendants(w + "p").First(p => p.Element(w + "pPr")?.Element(w + "numPr") is not null)
                .Element(w + "pPr")!;
            if (numbering == "absent") properties.Element(w + "numPr")!.Remove();
            else properties.Element(w + "numPr")!.Element(w + "numId")!.SetAttributeValue(w + "val", numbering);
        }, partOverride: story is "body" or "cell" ? "word/document.xml" : "word/" + story + "1.xml");
}
