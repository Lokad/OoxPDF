using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxSplitComputedFieldKeepsCachedPictureAfterHardBreak()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool complexField in new[] { false, true })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            PdfPage[] expected = Render(splitField: false);
            PdfPage[] actual = Render(splitField: true);
            TestAssert.Equal(2, actual.Length);
            TestAssert.Equal(0, actual[0].Images.Count);
            TestAssert.Equal(1, actual[1].Images.Count);
            TestAssert.True(expected.Select(page => page.Content).SequenceEqual(actual.Select(page => page.Content)),
                "The counter, continuation picture and prose retain the authored page placement.");

            PdfPage[] Render(bool splitField)
            {
                DocxDocument document = ReadMutation("CompositionProbes/inline-math-navigation", mode, false, false, (xml, w, _) =>
                {
                    XElement body = xml.Root!.Element(w + "body")!;
                    var picture = new XElement(body.Descendants(w + "drawing").Single());
                    var pageBreak = new XElement(w + "br", new XAttribute(w + "type", "page"));
                    var paragraph = new XElement(w + "p");
                    if (complexField)
                    {
                        paragraph.Add(Marker("begin"), new XElement(w + "r", new XElement(w + "instrText", " PAGE ")), Marker("separate"));
                        if (splitField)
                            paragraph.Add(new XElement(w + "r", pageBreak, picture, new XElement(w + "t", "32")), Marker("end"));
                        else
                            paragraph.Add(Marker("end"), new XElement(w + "r", pageBreak, picture));
                    }
                    else
                    {
                        var field = new XElement(w + "fldSimple", new XAttribute(w + "instr", "PAGE"));
                        paragraph.Add(field);
                        if (splitField)
                            field.Add(new XElement(w + "r", new XElement(w + "t", "31"), pageBreak, picture, new XElement(w + "t", "32")));
                        else
                            paragraph.Add(new XElement(w + "r", pageBreak, picture));
                    }
                    body.ReplaceNodes(paragraph, new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "AFTER PICTURE"))),
                        new XElement(w + "sectPr"));

                    XElement Marker(string kind) => new(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", kind)));
                });
                var counter = document.Paragraphs.SelectMany(paragraph => paragraph.FieldReferences).Single();
                TestAssert.Equal(DocxFieldKind.Page, counter.Kind);
                TestAssert.True(counter.UsesPlaceholder && counter.HasCachedResult == splitField,
                    "One instruction owns the counter and reports its authored cache.");
                return renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            }
        }
    }

    internal static void CheckComposedRunPageBreaks()
    {
        foreach (bool withNote in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = withNote ? "footnote-image-math" : "inline-math-navigation";
            DocxDocument Read(int pageBreaks) => ReadMutation("CompositionProbes/" + fixture, mode, false, false, (xml, w, m) =>
            {
                XElement body = xml.Root!.Element(w + "body")!;
                XElement paragraph = body.Elements(w + "p").Single(p => p.Elements(m + "oMath").Count() == 2);
                XElement[] equations = paragraph.Elements(m + "oMath").ToArray();
                foreach (var pair in equations.Zip(new[] { "flowAlpha", "flowBeta" }))
                {
                    XElement run = pair.First.Element(m + "f")!.Element(m + "num")!.Element(m + "r")!;
                    run.Element(m + "t")!.Value = pair.Second;
                    if (run.Element(m + "rPr") is not { } properties)
                        run.AddFirst(new XElement(m + "rPr", new XElement(m + "nor")));
                    else if (properties.Element(m + "nor") is null) properties.Add(new XElement(m + "nor"));
                }
                if (pageBreaks > 0)
                    equations[1].AddBeforeSelf(new XElement(w + "r", Enumerable.Range(0, pageBreaks)
                        .Select(_ => new XElement(w + "br", new XAttribute(w + "type", "page")))));
                if (!withNote)
                {
                    // Retain the existing raster relation and page properties;
                    // the note case below keeps the complete composed document.
                    body.Elements().Where(child => !ReferenceEquals(child, paragraph) && child.Name != w + "sectPr").Remove();
                    xml.Descendants(w + "headerReference").Remove();
                    xml.Descendants(w + "footerReference").Remove();
                    return;
                }

                var children = paragraph.Elements().ToArray();
                int first = Array.IndexOf(children, equations[0]), last = Array.IndexOf(children, equations[1]);
                var link = new XElement(w + "hyperlink", new XAttribute(w + "anchor", "BreakCalculation"));
                equations[0].AddBeforeSelf(link);
                foreach (XElement child in children[first..(last + 1)]) { child.Remove(); link.Add(child); }
                link.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "900"), new XAttribute(w + "name", "BreakCalculation")));
                link.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "900")));
                body.Elements(w + "p").First().Add(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "BreakCalculation"),
                    new XElement(w + "r", new XElement(w + "t", "Jump to calculation"))));
            });
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderPagesAsync(Read(0), null, default).ToBlockingEnumerable().ToArray();
            foreach (int pageBreaks in new[] { 1, 2 })
            {
                var pages = renderer.RenderPagesAsync(Read(pageBreaks), diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(pageBreaks + 1, pages.Length);
                // Painting order can interleave prose and formulas. The break must
                // preserve every non-whitespace glyph, independently of show order.
                static string Characters(IEnumerable<PdfPage> sequence) => string.Concat(sequence.SelectMany(page =>
                    DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true))
                    .SelectMany(show => show.Text).Where(c => !char.IsWhiteSpace(c)).Order());
                TestAssert.Equal(Characters(before), Characters(pages));
                string[] text = pages.Select(page => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                    usePageCoordinates: true).Select(show => show.Text))).ToArray();
                // A consumed wrapping space need not have its own painted glyph.
                // Check prose characters independently of line/show boundaries.
                string[] prose = text.Select(value => string.Concat(value.Where(c => !char.IsWhiteSpace(c)))).ToArray();
                TestAssert.Contains("flowAlpha", text[0]);
                TestAssert.True(!text[^1].Contains("flowAlpha", StringComparison.Ordinal), "The first formula stays before the break.");
                TestAssert.Contains("flowBeta", text[^1]);
                TestAssert.True(!text[0].Contains("flowBeta", StringComparison.Ordinal), "The second formula follows the break.");
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "A run break cannot discard supported formulas.");
                int Draws(PdfPage page) => page.Images.Sum(image => Regex.Matches(page.Content,
                    @"/" + Regex.Escape(image.ResourceName) + @"\s+Do\b").Count);
                TestAssert.Equal(1, Draws(pages[0]));
                TestAssert.Equal(0, pages.Skip(1).Sum(Draws));
                if (!withNote) continue;
                TestAssert.Contains("Publiccalculationnote.", prose[0]);
                TestAssert.True(!prose[^1].Contains("Publiccalculationnote.", StringComparison.Ordinal), "The note stays with its source marker.");
                TestAssert.Contains("Followingprose.", prose[^1]);
                var target = pages[^1].Annotations.Where(link => link.Destination.HasValue).Single().Destination!.Value;
                TestAssert.Equal(0, target.PageIndex);
                TestAssert.True(pages[0].Annotations.Count(link => link.Destination == target) >= 2,
                    "The leading navigation and first linked formula retain the bookmark destination.");
                TestAssert.True(pages[0].Annotations.Any(link => link.Destination is { PageIndex: 0 } note && note != target),
                    "The note reference retains its independent destination.");
            }
        }
    }
}
