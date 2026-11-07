using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class DocxCellColumnBreakTests
{
    public static void DocxCellColumnsKeepCompleteParagraphSourcesBeforePreparation() =>
        AssertCompleteSources("column", ["before", "same-run", "leading", "trailing", "wrapped", "nested", "column-paragraph", "page"]);

    public static void DocxCellInlinePagesKeepCompleteParagraphSourcesBeforePreparation()
    {
        AssertCompleteSources("page", ["before", "same-run", "leading", "trailing", "wrapped", "nested", "page"]);
        var paragraph = ContentCell(Read("plain", breakType: "page"), "plain").Paragraphs.Single();
        TestAssert.Equal("LeftRight", string.Concat(paragraph.Runs.Select(run => run.Text)));
    }

    private static void AssertCompleteSources(string breakType, IEnumerable<string> variants)
    {
        foreach (string variant in variants)
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = Read(variant, mode, breakType);
            var cell = ContentCell(document, variant);
            var elements = DocxTableCellContent.GetBodyElements(cell);
            TestAssert.Equal(variant == "column-paragraph" ? 3 : 1, elements.Count);
            TestAssert.True(elements.All(element => element is DocxParagraphElement),
                "Inline cell flow tokens do not manufacture block boundaries: " + variant);
            var paragraph = elements.OfType<DocxParagraphElement>().Select(p => p.Paragraph).Single(p => p.Equations.Count != 0);
            var items = TestAssert.NotNull(paragraph.InlineContent);
            TestAssert.Equal(2, paragraph.Equations.Count);
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.Equal(2, DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, default).Sum(p => p.Equations.Count));
            TestAssert.True(items.OfType<DocxInlineEquation>().Select(item => item.Equation)
                .SequenceEqual(paragraph.Equations, ReferenceEqualityComparer.Instance), "Preparation uses the actual ordered equation owners.");
            TestAssert.True(paragraph.Images.Single().Bytes is { Length: > 0 } &&
                ReferenceEquals(paragraph.Images.Single(), items.OfType<DocxInlinePicture>().Single().Image),
                "The source picture remains owned and resolvable.");
            TestAssert.True(DocxInlineContent.HasValidHyperlinkSourceRanges(paragraph, default), "Source link boundaries survive the break.");
            TestAssert.True(TestAssert.NotNull(DocxInlineContent.GetHyperlinkedEquations(paragraph, default))
                .Contains(paragraph.Equations.Last()), "The last formula retains its internal link.");
            if (variant == "wrapped")
            {
                TestAssert.Equal(1, paragraph.NonRenderingRunOwners.Count);
                TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(paragraph with { NonRenderingRunOwners = [] }),
                    "Only the actual nonpainting cell run owns its insertion.");
            }
        }
    }

    public static void DocxCellColumnsRetainParagraphMarksAndRealPageBreaks()
    {
        var paragraphs = ContentCell(Read("column-paragraph"), "column-paragraph").Paragraphs;
        TestAssert.Equal(3, paragraphs.Count);
        TestAssert.True(paragraphs[1].Runs.All(run => run.Text.Length == 0), "A real source paragraph mark is not discarded with its column break.");
        var document = Read("column-paragraph", breakType: "page");
        var pageElements = ContentCell(document, "column-paragraph").BodyElements;
        TestAssert.Equal(1, pageElements.OfType<DocxPageBreakElement>().Count());
        TestAssert.True(!pageElements.OfType<DocxManualBreakElement>().Any(), "A standalone cell page paragraph retains its existing flow owner.");
        TestAssert.Equal(2, DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, default).Sum(p => p.Equations.Count));
        var unsupported = ContentCell(Read("unsupported-owner"), "unsupported-owner").Paragraphs.Single();
        TestAssert.Equal(0, unsupported.NonRenderingRunOwners.Count);
        TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(unsupported) && DocxMathParagraphPolicy.GetPreparationEquations(unsupported).Count == 0,
            "An unknown object beside the ignored break retains the explicit unsupported-owner fallback.");
    }

    public static void DocxCellColumnConversionKeepsFormulaPictureAndNavigationContent() =>
        AssertConversionContent("column", ["before", "same-run", "leading", "trailing", "nested", "column-paragraph", "wrapped", "page"]);

    public static void DocxCellPageConversionKeepsFormulaPictureAndNavigationContent() =>
        AssertConversionContent("page", ["before", "same-run", "leading", "trailing", "nested", "column-paragraph", "wrapped", "page"]);

    private static void AssertConversionContent(string breakType, IEnumerable<string> variants)
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            foreach (string variant in variants)
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                    OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
                var pages = renderer.RenderPagesAsync(Read(variant, mode, breakType), diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
                TestAssert.True(text.Contains("x+2", StringComparison.Ordinal) && text.Contains("u+v=w", StringComparison.Ordinal),
                    "Both body formulas survive: " + variant + "/" + mode);
                TestAssert.Equal(1, pages.Sum(page => page.Images.Count));
                TestAssert.True(pages.Sum(page => page.Annotations.Count(annotation => annotation.Destination is not null)) >= 1,
                    "The formula's destination survives: " + variant);
                TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"), "Known cell flow tokens do not make ordinary equations unsupported.");
                Convert(variant, mode, breakType);
            }
        }
    }

    private static void Convert(string variant, OoxPdfDocxMarkupMode mode, string breakType)
    {
        using var input = Input(variant, breakType); using var output = new MemoryStream();
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions {
            InputKind = OoxPdfInputKind.Docx, FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true),
            DocxMarkupMode = mode, FixedCreationDate = DateTimeOffset.UnixEpoch, DiagnosticSink = diagnostics.Add });
        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"), "Public conversion preserves supported math: " + variant);
    }

    internal static DocxDocument Read(string variant, OoxPdfDocxMarkupMode mode = OoxPdfDocxMarkupMode.Final, string breakType = "column")
    {
        using var input = Input(variant, breakType); using var package = OoxPackage.Open(input, default);
        return new DocxReader().Read(package, null, default, mode);
    }

    private static DocxTableCell ContentCell(DocxDocument document, string variant)
    {
        var cell = document.BodyElements.OfType<DocxTableElement>().Single().Table.Rows.Single().Cells.Single();
        return variant == "nested" ? cell.BodyElements.OfType<DocxTableElement>().Single().Table.Rows.Single().Cells.Single() : cell;
    }

    private static MemoryStream Input(string variant, string breakType = "column")
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes/body.docx"));
        using var source = ZipFile.OpenRead(path);
        var parts = source.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var stream = entry.Open(); using var copy = new MemoryStream(); stream.CopyTo(copy); return copy.ToArray(); });
        using var xmlStream = new MemoryStream(parts["word/document.xml"]); var xml = XDocument.Load(xmlStream);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var paragraph = xml.Root!.Element(w + "body")!.Element(w + "p")!;
        var math = paragraph.Elements(m + "oMath").Last(); math.Descendants(m + "t").Single().Value = "u+v=w";
        math.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "MixedObjects"), new XElement(math)));
        paragraph.AddFirst(Text("PublicPrefix ")); paragraph.Add(Text(" PublicSuffix"));
        var cell = new XElement(w + "tc");
        var content = new XElement(paragraph);
        var column = new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", breakType)));
        switch (variant)
        {
            case "before": case "nested":
                content.Element(w + "bookmarkStart")!.AddBeforeSelf(column); break;
            case "same-run":
                content.Elements(w + "r").First().ReplaceNodes(new XElement(w + "t", "Public"), column.Element(w + "br"),
                    new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), "Prefix ")); break;
            case "leading": content.AddFirst(column); break;
            case "trailing": content.Add(column); break;
            case "wrapped": case "unsupported-owner":
                if (variant == "unsupported-owner") column.Add(new XElement(w + "object"));
                content.Element(w + "bookmarkStart")!.AddBeforeSelf(new XElement(w + "ins", new XAttribute(w + "id", "7"),
                    new XAttribute(w + "author", "Public reviewer"), column)); break;
            case "page":
                column.Add(new XElement(w + "br", new XAttribute(w + "type", "page")));
                content.Element(w + "bookmarkStart")!.AddBeforeSelf(column); break;
            case "plain":
                content.ReplaceNodes(new XElement(w + "r", new XElement(w + "t", "Left"), column.Element(w + "br"), new XElement(w + "t", "Right"))); break;
            case "column-paragraph":
                var prefix = content.Elements(w + "r").First(); prefix.Remove(); cell.Add(new XElement(w + "p", prefix), new XElement(w + "p", column)); break;
        }
        cell.Add(content);
        var table = Table(cell);
        if (variant == "nested") table = Table(new XElement(w + "tc", table));
        paragraph.ReplaceWith(table);
        using var updated = new MemoryStream(); xml.Save(updated, SaveOptions.DisableFormatting); parts["word/document.xml"] = updated.ToArray();
        return TestFixtures.CreateZipPackage(parts);

        XElement Text(string value) => new(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value));
        XElement Table(XElement sourceCell) => new(w + "tbl", new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "8640"))), new XElement(w + "tr", sourceCell));
    }
}
