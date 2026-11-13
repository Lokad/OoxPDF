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
                // This mixed cell reference's PAGE value remains approximate.
                // Compare every word/formula; exact reference values have their
                // separate qualified counter-owner tests.
                string WithoutNotePage(string text) => Regex.Replace(text, @"(?<=Calculation:)[0-9]+(?=and)", "#PAGE#");
                TestAssert.Equal(WithoutNotePage(Text(literal[page])), WithoutNotePage(Text(actual[page])));
                TestAssert.Equal(literal[page].Images.Count, actual[page].Images.Count);
            }
            string painted = NormalizeNoteMathFieldText(string.Concat(actual.Select(DocxPageFootnoteNumberingTests.DecodeText)));
            var result = Regex.Matches(painted, @"Calculation:(?<page>[0-9]+)and");
            TestAssert.Equal(1, result.Count);
            int value = int.Parse(result[0].Groups["page"].Value, System.Globalization.CultureInfo.InvariantCulture);
            TestAssert.True(value >= 1 && value <= actual.Length, "The note still paints one computed page result.");
        }
    }

    public static void DocxComputedSimpleFieldPicturesSurviveCacheRunSplitsAcrossStories()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (string kind in new[] { "PAGE", "NUMPAGES" })
        foreach (int pictures in new[] { 1, 2 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var before = renderer.RenderPagesAsync(ReadComputedSimpleFieldPictures(story, kind, mode, false, pictures),
                null, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(ReadComputedSimpleFieldPictures(story, kind, mode, true, pictures),
                null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(pictures, before.Sum(page => page.Images.Count));
            TestAssert.Equal(pictures, after.Sum(page => page.Images.Count));
            TestAssert.Equal(before.Length, after.Length);
            for (int index = 0; index < before.Length; index++)
            {
                TestAssert.Equal(before[index].Content, after[index].Content);
                TestAssert.Equal(before[index].Fonts.Count, after[index].Fonts.Count);
                TestAssert.True(before[index].Annotations.SequenceEqual(after[index].Annotations),
                    "Splitting a cached picture run preserves equation and navigation ownership.");
            }
        }
    }

    public static void DocxComputedSimpleFieldPicturesKeepVisibilityLocksAndRevisionOwners()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string revision in new[] { "none", "ins", "del" })
        foreach (bool hidden in new[] { false, true })
        foreach (bool locked in new[] { false, true })
        {
            var document = ReadComputedSimpleFieldPictures("body", "PAGE", mode, true, 2,
                hidden, locked, revision);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document).Single(p => p.FieldReferences.Count > 0 || p.Equations.Count > 0);
            bool included = revision == "none" || (revision == "ins" ? mode != OoxPdfDocxMarkupMode.Original :
                mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup);
            TestAssert.Equal(included ? hidden ? 1 : 2 : 0, paragraph.Images.Count);
            if (!included) continue;
            var field = paragraph.FieldReferences.Single();
            TestAssert.Equal(!locked, field.UsesPlaceholder);
            TestAssert.Equal(locked, field.RendersCachedResult);
            var items = TestAssert.NotNull(paragraph.InlineContent);
            TestAssert.Equal(paragraph.Images.Count, items.OfType<DocxInlinePicture>().Count());
            TestAssert.True(DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default),
                "Every retained cached picture keeps the computed result's actual revision owner.");
            if (!locked)
            {
                var result = paragraph.Runs.Single(run => run.FieldKind == DocxFieldKind.Page);
                TestAssert.True(paragraph.Images.All(image => image.SourceRunIndex == result.SourceRunIndex),
                    "Cache run count does not invent extra computed field source positions.");
            }
        }
    }

    private static DocxDocument ReadComputedSimpleFieldPictures(string story, string kind,
        OoxPdfDocxMarkupMode mode, bool split, int count, bool hiddenLast = false,
        bool locked = false, string revision = "none") =>
        ReadMutation("CompositionProbes/math-composition-unnumbered" + (story == "body" ? "" : "-" + story),
            mode, true, true, (xml, w, _) =>
            {
                var drawing = xml.Descendants(w + "drawing").Single();
                var source = drawing.Parent!;
                var field = new XElement(w + "fldSimple", new XAttribute(w + "instr", " " + kind + " "),
                    locked ? new XAttribute(w + "fldLock", "true") : null);
                var value = new XElement(w + "r", new XElement(w + "t", "77"));
                field.Add(value);
                for (int index = 0; index < count; index++)
                {
                    var picture = new XElement(drawing);
                    if (index > 0)
                        foreach (var extent in picture.Descendants().Where(element => element.Attribute("cx") is not null))
                            extent.SetAttributeValue("cx", (long)extent.Attribute("cx")! / 2);
                    if (split)
                        field.Add(new XElement(w + "r",
                            hiddenLast && index == count - 1 ? new XElement(w + "rPr", new XElement(w + "vanish")) : null,
                            picture));
                    else value.Add(picture);
                }
                source.ReplaceWith(revision == "none" ? field : new XElement(w + revision,
                    new XAttribute(w + "id", "907"), new XAttribute(w + "author", "Public reviewer"), field));
            }, partOverride: story is "header" or "footer" ? "word/" + story + "1.xml" : null);

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
