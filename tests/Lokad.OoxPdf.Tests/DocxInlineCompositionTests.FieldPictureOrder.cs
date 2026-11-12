using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxComputedFieldsKeepAtomicOrderBesideRunDrawings()
    {
        foreach (string kind in new[] { "PAGE", "NUMPAGES" })
        foreach (string cache in new[] { "", "7", "77" })
        foreach (string prefix in new[] { "", "Lead ", "A longer source prefix " })
        foreach (bool pictureFirst in new[] { false, true })
        {
            var paragraph = ReadFieldPictureBoundary(kind, cache, prefix, pictureFirst, joined: true,
                OoxPdfDocxMarkupMode.Final).Paragraphs.Single(p => p.Images.Count != 0);
            var items = TestAssert.NotNull(paragraph.InlineContent);
            var result = items.OfType<DocxInlineText>().Single(text => text.Run.FieldKind is not null);
            TestAssert.Equal("{" + kind + "}", result.Run.Text);
            TestAssert.Equal(0, result.Start);
            TestAssert.Equal(result.Run.Text.Length, result.Length);
            int fieldIndex = items.ToList().IndexOf(result);
            int pictureIndex = items.ToList().FindIndex(item => item is DocxInlinePicture);
            TestAssert.Equal(pictureFirst, pictureIndex < fieldIndex);
            TestAssert.Equal(prefix.Length, result.Run.SourceTextOffsetInRun);
            TestAssert.Equal(pictureFirst ? 1 : 0, result.Run.SourceDrawingIndexInRun);
            TestAssert.Equal(1, items.OfType<DocxInlineEquation>().Count());
        }
    }

    public static void DocxComputedFieldPictureRunBoundariesPreserveVisibleContent()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string kind in new[] { "PAGE", "NUMPAGES" })
        foreach (string cache in new[] { "", "7", "77" })
        foreach (bool pictureFirst in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var pages = new[] { false, true }.Select(joined => new DocxRenderer(resolver, mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderPagesAsync(
                        ReadFieldPictureBoundary(kind, cache, "Lead ", pictureFirst, joined, mode), null, default)
                    .ToBlockingEnumerable().ToArray()).ToArray();
            TestAssert.Equal(2, pages[0].Length);
            TestAssert.Equal(2, pages[1].Length);
            string Text(int variant) => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[variant][1], true)
                .Select(show => show.Text));
            TestAssert.Contains("Lead 2 tail ", Text(0));
            TestAssert.Equal(Text(0), Text(1));
            TestAssert.True(!Text(1).Contains("{" + kind + "}", StringComparison.Ordinal),
                "A picture cannot split the computed token into literal placeholder text.");
            TestAssert.Equal(1, pages[1][1].Images.Count);
            string Picture(int variant) => Regex.Match(pages[variant][1].Content,
                @"[-\d.]+ 0 0 [-\d.]+ [-\d.]+ [-\d.]+ cm\s+/Im\d+ Do").Value;
            TestAssert.True(Picture(0).Length != 0, "The control paints its source picture.");
            TestAssert.Equal(Picture(0), Picture(1));
        }
    }

    public static void DocxComposedEndnoteCountersKeepNeighborFractionsInBothViews()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument Source(bool counter) => ReadMutation("Composition/composed-static-cell-report", mode, false, false,
                (xml, w, _) =>
                {
                    var picture = xml.Descendants(w + "r").Single(run => run.Element(w + "drawing") is not null);
                    var result = new XElement(w + "r");
                    if (counter) result.Add(new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "begin")),
                        new XElement(w + "instrText", " PAGE "),
                        new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "separate")));
                    result.Add(new XElement(w + "t", counter ? "7" : "2"));
                    if (counter) result.Add(new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "end")));
                    picture.AddBeforeSelf(result);
                }, partOverride: "word/endnotes.xml");
            var renderer = new DocxRenderer(resolver, mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var literal = renderer.RenderPagesAsync(Source(false), null, default).ToBlockingEnumerable().ToArray();
            var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(Source(true), diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, literal.Length);
            TestAssert.Equal(literal.Length, actual.Length);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "A supported page result beside a note picture must not suppress either existing fraction.");
            for (int page = 0; page < literal.Length; page++)
            {
                string Text(Lokad.OoxPdf.Pdf.PdfPage value) => NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(value));
                TestAssert.Equal(Text(literal[page]), Text(actual[page]));
                TestAssert.Equal(literal[page].Images.Count, actual[page].Images.Count);
            }
        }
    }

    private static DocxDocument ReadFieldPictureBoundary(string kind, string cache, string prefix,
        bool pictureFirst, bool joined, OoxPdfDocxMarkupMode mode) =>
        ReadMutation("Composition/report", mode, false, false, (xml, w, math) =>
        {
            var body = xml.Root!.Element(w + "body")!;
            var section = new XElement(body.Element(w + "sectPr")!);
            var picture = new XElement(body.Descendants(w + "drawing").Single());
            var equation = new XElement(body.Descendants(math + "oMath").First());
            var prefixRun = new XElement(w + "r", new XElement(w + "t",
                new XAttribute(XNamespace.Xml + "space", "preserve"), prefix));
            var fieldRun = new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "begin")),
                new XElement(w + "instrText", " " + kind + " "),
                new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "separate")),
                new XElement(w + "t", cache), new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "end")));
            var pictureRun = new XElement(w + "r", picture);
            var suffixRun = new XElement(w + "r", new XElement(w + "t",
                new XAttribute(XNamespace.Xml + "space", "preserve"), " tail "));
            var runs = pictureFirst ? new[] { prefixRun, pictureRun, fieldRun, suffixRun } :
                new[] { prefixRun, fieldRun, pictureRun, suffixRun };
            var paragraph = new XElement(w + "p");
            if (joined) paragraph.Add(new XElement(w + "r", runs.SelectMany(run => run.Elements()).Select(node => new XElement(node))));
            else paragraph.Add(runs);
            paragraph.Add(equation, new XElement(w + "r", new XElement(w + "t", " end.")));
            body.RemoveNodes();
            body.Add(new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "Before field picture."))),
                new XElement(w + "p", new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page")))),
                paragraph, section);
        });
}
