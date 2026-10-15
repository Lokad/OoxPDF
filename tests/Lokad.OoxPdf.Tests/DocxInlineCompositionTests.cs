using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxInlineCompositionTests
{
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

    private static DocxDocument ReadMutation(string fixture, OoxPdfDocxMarkupMode mode, bool bookmark, bool split)
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
                var equation = document.Descendants(m + "oMath").Single();
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
                entry.Delete();
                using var output = zip.CreateEntry(part).Open();
                document.Save(output);
            }
            return DocxTests.ReadDocx(path, mode);
        }
        finally { File.Delete(path); }
    }
}
