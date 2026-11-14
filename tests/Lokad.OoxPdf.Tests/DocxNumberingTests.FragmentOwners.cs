using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxNumberingTests
{
    public static void DocxCellBreakFragmentsRetainOneNumberedSourceOwner()
    {
        foreach (string kind in new[] { "page", "page-column", "column-page", "column", "soft" })
        foreach (bool pure in new[] { false, true })
        foreach (int count in new[] { 1, 2 })
        foreach (bool inserted in new[] { false, true })
        foreach (OoxPdfDocxMarkupMode mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var document = ReadFirstMarkerBreak(kind, pure, count, "runs", "decimal", inserted, mode, inCell: true);
            var cell = document.BodyElements.OfType<DocxTableElement>().Single().Table.Rows.Single().Cells.Single();
            var paragraphs = cell.BodyElements.OfType<DocxParagraphElement>().Select(element => element.Paragraph).ToArray();
            var numbered = paragraphs.Where(paragraph => paragraph.ListLabel is not null).ToArray();
            TestAssert.True(numbered.Select(paragraph => paragraph.ListLabel!.Text).SequenceEqual(new[] { "1.", "2.", "3." }),
                "Cell breaks retain the first marker and advance the source counter once: " + kind + "/" + pure + "/" + mode);
            bool review = DocxMarkupContext.FromMode(mode).AppliesInlineRevisionStyle;
            TestAssert.Equal(inserted && review ? "1." : null, numbered[1].ListLabel?.NumberingChange?.PreviousText);
            TestAssert.Equal(inserted && review ? "2." : null, numbered[2].ListLabel?.NumberingChange?.PreviousText);
            int pageEvents = cell.BodyElements.OfType<DocxPageBreakElement>().Sum(element => element.BreakCount);
            TestAssert.Equal(pure && kind.Contains("page", StringComparison.Ordinal) ? count : 0, pageEvents);
            foreach (var continuation in paragraphs.SkipWhile(paragraph => !ReferenceEquals(paragraph, numbered[0]))
                .Skip(1).TakeWhile(paragraph => !ReferenceEquals(paragraph, numbered[1])))
            {
                TestAssert.True(continuation.Indent.FirstLinePoints is null && continuation.Indent.HangingPoints is null,
                    "A cell continuation does not repeat the source first-line offset.");
                TestAssert.Equal(36d, continuation.Indent.LeftPoints ?? 0d);
                TestAssert.True(DocxInlineContent.HasSupportedParagraphRevisionOwners(continuation, default),
                    "Actual numbered mark ownership survives suppressed continuation ink.");
            }
        }
    }

    public static void DocxCellBreakMarkersPaintWithoutLosingFollowingItems()
    {
        foreach (string kind in new[] { "page", "page-column", "column-page", "soft" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadFirstMarkerBreak(kind, true, 2, "control", "decimal", true, mode, inCell: true);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            var text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            foreach (string expected in new[] { "BEFORE", "1.", "2.", "3.", "NEXT", "LAST" }) TestAssert.Contains(expected, text);
            for (int page = 0; page < pages.Length; page++) PdfContentValidator.ValidatePage(pages[page], page, default);
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPagesAsync(document, null, cancelled.Token)
                .ToBlockingEnumerable().ToArray());
        }
    }

    public static void DocxNumberedBreakContinuationsKeepRevisedMathContent()
    {
        foreach (string kind in new[] { "page", "column" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
                "CompositionProbes", "numbering-mark-math-body.docx"));
            using var input = new MemoryStream();
            using (var source = ZipFile.OpenRead(path))
            using (var output = new ZipArchive(input, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var entry in source.Entries)
                {
                    using var original = entry.Open();
                    using var target = output.CreateEntry(entry.FullName).Open();
                    if (entry.FullName != "word/document.xml") { original.CopyTo(target); continue; }
                    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    var xml = XDocument.Load(original);
                    var first = xml.Descendants(w + "p").First(paragraph => paragraph.Element(w + "pPr")?.Element(w + "numPr") is not null);
                    first.Element(w + "pPr")!.AddAfterSelf(new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", kind))));
                    xml.Save(target);
                }
            }
            input.Position = 0;
            var document = new DocxReader().Read(OoxPackage.Open(input, default), null, default, mode);
            var continuation = document.Paragraphs.Single(paragraph => paragraph.Equations.Count != 0 && paragraph.ListLabel is null);
            TestAssert.True(continuation.ParagraphMarkRun is not null &&
                DocxInlineContent.HasSupportedSourceOwner(continuation), "The actual numbered source owns the continuation revision.");
            TestAssert.Equal("Font: Bold", DocxRenderer.BuildWordCompatibleRunFormattingPreview(continuation, default));
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            var math = pages.SelectMany(page =>
            {
                var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
                return DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true)
                    .Where(show => fonts.Contains(show.Font)).ToArray();
            }).ToArray();
            TestAssert.Equal(18, math.Length);
            string text = string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("x+2y", StringSplitOptions.None).Length - 1);
            TestAssert.Equal(2, text.Split("a+b=c", StringSplitOptions.None).Length - 1);
            TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Adding a break cannot omit supported equations in a revised numbered paragraph.");
            TestAssert.True(pages.Any(page => page.Images.Any(image => page.Content.Contains("/" + image.ResourceName + " Do", StringComparison.Ordinal))),
                "The same composition retains its adjacent picture.");
            for (int page = 0; page < pages.Length; page++) PdfContentValidator.ValidatePage(pages[page], page, default);
        }
    }
}
