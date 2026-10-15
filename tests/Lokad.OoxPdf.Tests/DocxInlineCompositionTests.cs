using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionSourceViewOrdersEquationsPicturesTabsAndBreaks()
    {
        DocxDocument document = ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false,
            (source, w, m) =>
            {
                var equation = new XElement(source.Descendants(m + "oMath").First());
                var drawing = new XElement(source.Descendants(w + "drawing").Single());
                source.Root!.Element(w + "body")!.ReplaceNodes(new XElement(w + "p",
                    new XElement(w + "bookmarkStart", new XAttribute(w + "id", "9"), new XAttribute(w + "name", "Mixed")),
                    equation,
                    new XElement(w + "r", new XElement(w + "t", "ABC"), drawing, new XElement(w + "tab"),
                        new XElement(w + "t", "DEF"), new XElement(w + "br")),
                    new XElement(equation), new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "9"))));
            });
        DocxParagraph paragraph = document.Paragraphs.Single();
        var items = TestAssert.NotNull(paragraph.InlineContent);
        TestAssert.Equal("DocxInlineEquation,DocxInlineText,DocxInlinePicture,DocxInlineTab,DocxInlineText,DocxInlineBreak,DocxInlineEquation",
            string.Join(',', items.Select(item => item.GetType().Name)));
        TestAssert.True(ReferenceEquals(paragraph.Equations[0], ((DocxInlineEquation)items[0]).Equation) &&
            ReferenceEquals(paragraph.Equations[1], ((DocxInlineEquation)items[^1]).Equation) &&
            ReferenceEquals(paragraph.Images.Single(), ((DocxInlinePicture)items[2]).Image),
            "Ordered items retain source identities instead of reparsing or cloning visible objects.");
        TestAssert.Equal(3, paragraph.Images.Single().SourceTextOffsetInRun);
        TestAssert.Equal(1, paragraph.BookmarkAnchors.Count);
        TestAssert.Equal("ABCDEF", string.Concat(items.OfType<DocxInlineText>().Select(text => text.Run.Text.Substring(text.Start, text.Length))));
        TestAssert.True(DocxInlineContent.TryGetEquationTextOffset(items, paragraph.Equations[1], out int offset) && offset == 8,
            "The second equation follows six text characters, one tab and one break; pictures have their own width.");
    }

    public static void DocxInlineCompositionSourceViewRetainsPrefixOrderAndCancellation()
    {
        DocxDocument document = ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false);
        foreach (DocxParagraph paragraph in document.Paragraphs.Where(p => p.Equations.Count != 0))
        {
            var items = TestAssert.NotNull(paragraph.InlineContent);
            TestAssert.True(items.OfType<DocxInlineEquation>().Select(item => item.Equation).SequenceEqual(paragraph.Equations),
                "Multiple equations retain their authored order.");
            TestAssert.Equal(string.Concat(paragraph.Runs.Select(run => run.Text)),
                string.Concat(items.OfType<DocxInlineText>().Select(text => text.Run.Text.Substring(text.Start, text.Length))));
        }
        var imageParagraph = document.Paragraphs.Single(p => p.Images.Count != 0);
        var ordered = TestAssert.NotNull(imageParagraph.InlineContent);
        TestAssert.True(ordered[0] is DocxInlineText && ordered[1] is DocxInlinePicture &&
            ordered[2] is DocxInlineText && ordered[3] is DocxInlineEquation,
            "A picture before the first equation joins the initial source prefix in place.");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.FromPrefix(imageParagraph.Runs, imageParagraph.Images, cancelled.Token));
    }

    public static void DocxInlineCompositionSourceViewCannotOutliveReplacedCollections()
    {
        var document = ReadMutation("docx-omml-isolated-inline-body-fraction-left", OoxPdfDocxMarkupMode.Final, false, false);
        var paragraph = document.Paragraphs.Single(p => p.Equations.Count != 0);
        TestAssert.NotNull(paragraph.InlineContent);
        TestAssert.NotNull((paragraph with { SpacingAfterPoints = 9d }).InlineContent);
        var changed = paragraph with { Runs = [new DocxTextRun("New visible prose", 11d, null, false, false, false, null, null)] };
        TestAssert.True(changed.InlineContent is null && DocxMathParagraphPolicy.GetStandaloneRow(changed) is null,
            "A legacy source replacement cannot reuse the old standalone-content eligibility.");
        TestAssert.True((paragraph with { Images = [] }).InlineContent is null &&
            (paragraph with { Equations = [] }).InlineContent is null,
            "Image and equation replacement also invalidate their borrowed source view.");
    }

    public static void DocxInlineCompositionBookmarkRetainsSupportedMathOwnership()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument source = ReadMutation("docx-omml-isolated-inline-body-fraction-left", mode, false, false);
            DocxDocument annotated = ReadMutation("docx-omml-isolated-inline-body-fraction-left", mode, true, false);
            DocxParagraph before = source.Paragraphs.Single(p => p.Equations.Count != 0);
            DocxParagraph after = annotated.Paragraphs.Single(p => p.Equations.Count != 0);
            TestAssert.Equal(1, after.BookmarkAnchors.Count);
            TestAssert.NotNull(DocxMathParagraphPolicy.GetStandaloneRow(before));
            TestAssert.NotNull(DocxMathParagraphPolicy.GetStandaloneRow(after),
                "A zero-width source annotation cannot remove supported mathematical content.");
            TestAssert.True(DocxMathLayout.TryGetRuns(before.Equations.Single(), default, out var beforeRuns) &&
                DocxMathLayout.TryGetRuns(after.Equations.Single(), default, out var afterRuns) &&
                string.Concat(beforeRuns.Select(r => r.Run.Text.Text)) == string.Concat(afterRuns.Select(r => r.Run.Text.Text)),
                "Source mathematical content survives the annotation unchanged.");
        }
    }

    public static void DocxInlineCompositionBookmarkAndRunSplittingPreserveRenderedContent()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the rendering comparison.");
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string fixture in new[] { "docx-omml-isolated-inline-body-fraction-left", "docx-omml-inline-fraction-middle12",
            "docx-omml-isolated-inline-cell-fraction-left", "docx-omml-static-header-fraction", "docx-omml-static-footer-fraction" })
        {
            var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var renderer = new DocxRenderer(null, mode, geometry);
            var before = renderer.RenderBlankPages(ReadMutation(fixture, mode, false, false), null, default).ToArray();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var after = renderer.RenderBlankPages(ReadMutation(fixture, mode, true, true), diagnostics.Add, default).ToArray();
            string Text(IEnumerable<Lokad.OoxPdf.Pdf.PdfPage> pages) =>
                string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(Text(before), Text(after));
            TestAssert.Equal(before.Length, after.Length);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "A bookmark or unchanged-format run split cannot turn painted math into ignored content.");
        }
    }

    private static DocxDocument ReadMutation(string fixture, OoxPdfDocxMarkupMode mode, bool bookmark, bool split,
        Action<XDocument, XNamespace, XNamespace>? change = null)
    {
        string source = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", fixture + ".docx"));
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        File.Copy(source, path);
        try
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                string part = fixture.Contains("static-header", StringComparison.Ordinal) ? "word/header1.xml" :
                    fixture.Contains("static-footer", StringComparison.Ordinal) ? "word/footer1.xml" : "word/document.xml";
                var entry = zip.GetEntry(part)!;
                XDocument document;
                using (var stream = entry.Open()) document = XDocument.Load(stream);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
                var equation = document.Descendants(m + "oMath").First();
                if (bookmark)
                {
                    var container = equation.Parent!.Name == m + "oMathPara" ? equation.Parent : equation;
                    container.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "41"), new XAttribute(w + "name", "PublicCalculation")));
                    container.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "41")));
                }
                if (split)
                {
                    foreach (var run in document.Descendants(w + "r").Where(r => r.Parent?.Name == w + "p").ToArray())
                    {
                        if (run.Element(w + "t") is not { Value.Length: > 2 } text || run.Elements().Any(e => e.Name != w + "t" && e.Name != w + "rPr")) continue;
                        int middle = text.Value.Length / 2;
                        var second = new XElement(run);
                        second.Element(w + "t")!.Value = text.Value[middle..];
                        text.Value = text.Value[..middle];
                        text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                        second.Element(w + "t")!.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                        run.AddAfterSelf(second);
                    }
                }
                change?.Invoke(document, w, m);
                entry.Delete();
                using var output = zip.CreateEntry(part).Open();
                document.Save(output);
            }
            return DocxTests.ReadDocx(path, mode);
        }
        finally { File.Delete(path); }
    }
}
