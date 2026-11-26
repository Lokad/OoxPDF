using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxFieldResultCodeSpellingKeepsLiteralTextAndDrawingOrder()
    {
        const string cache = " Cached dispatch ";
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string story in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var (instruction, locked) in new[] { ("DATE", false), ("PAGE", true), ("PAGE", false) })
        {
            DocxDocument Source(string cacheElement)
            {
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                DocxDocument document = ReadMutation("CompositionProbes/math-composition-unnumbered" + (story == "body" ? "" : "-cell"),
                    mode, true, false, (xml, w, _) =>
                    {
                        XElement drawing = xml.Descendants(w + "drawing").Single();
                        XElement Marker(string kind) => new(w + "fldChar", new XAttribute(w + "fldCharType", kind));
                        XElement TextElement(string kind, string value) =>
                            new(w + kind, new XAttribute(XNamespace.Xml + "space", "preserve"), value);
                        XElement begin = Marker("begin");
                        if (locked) begin.SetAttributeValue(w + "fldLock", "1");
                        XElement field = new(w + "r", begin, TextElement("instrText", " " + instruction + " "),
                            Marker("separate"), new XElement(drawing), TextElement(cacheElement, cache), new XElement(drawing),
                            Marker("begin"), TextElement("instrText", " REF Target "), Marker("separate"),
                            TextElement("t", "Nested value"), Marker("end"), Marker("end"), TextElement("t", " after field "));
                        drawing.Parent!.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), field));
                    }, diagnosticSink: diagnostics.Add);
                TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_COMPLEX_FIELD"),
                    "A literal code element inside a stored result is supported.");
                return document;
            }

            var control = Source("t");
            var candidate = Source("instrText");
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(candidate).Single(p => p.Images.Count != 0);
            TestAssert.True(paragraph.Images.Select(image => image.SourceTextOffsetInRun).SequenceEqual([0, cache.Length]),
                "Only the second picture follows the literal text in the same run.");
            TestAssert.True(paragraph.Images.Select(image => image.SourceOrderInRun).SequenceEqual([0, 1]),
                "Both pictures retain their distinct source order.");
            var nested = paragraph.FieldReferences.Single(field => field.NestingDepth == 1);
            TestAssert.Equal("REF Target", nested.Instruction!.Trim());
            TestAssert.Equal(1, nested.InstructionRunCount);

            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var expected = renderer.RenderPagesAsync(control, null, default).ToBlockingEnumerable().ToArray();
            var actual = renderer.RenderPagesAsync(candidate, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(expected.Length, actual.Length);
            string text = string.Concat(actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true))
                .Select(show => show.Text));
            bool computed = instruction == "PAGE" && !locked;
            TestAssert.Equal(!computed, text.Contains(cache.Trim(), StringComparison.Ordinal));
            TestAssert.Equal(!computed, text.Contains("Nested value", StringComparison.Ordinal));
            TestAssert.Contains("after field", text);
            TestAssert.True(actual.SelectMany(page => page.Annotations).Any(annotation => annotation.Destination.HasValue),
                "The cached field and neighboring content retain internal links.");
            for (int index = 0; index < actual.Length; index++)
            {
                TestAssert.Equal(expected[index].Content, actual[index].Content);
                TestAssert.True(expected[index].Annotations.SequenceEqual(actual[index].Annotations),
                    "The literal source spelling preserves picture and link placement.");
            }
        }
    }

    public static void DocxComputedComplexCachesKeepPicturesMathAndNavigationAcrossStories()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (string kind in new[] { "PAGE", "NUMPAGES" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument Source(string cache, string shape) => ReadMutation(
                "CompositionProbes/math-composition-unnumbered" + (story == "body" ? "" : "-" + story),
                mode, true, true, (xml, w, _) =>
                {
                    var drawing = xml.Descendants(w + "drawing").Single();
                    var value = new XElement(w + "r", new XElement(w + "t", cache));
                    object result = shape == "joined" ? new XElement(w + "r", new XElement(w + "t", cache), new XElement(drawing))
                        : new[] { value, new XElement(w + "r", new XElement(drawing)) };
                    var field = CachedMathField(w, result, complex: true, instruction: kind, locked: false);
                    drawing.Parent!.ReplaceWith(shape == "packed"
                        ? new XElement(w + "r", field.SelectMany(node => node.Name == w + "r" ? new[] { node } : node.Elements(w + "r"))
                            .SelectMany(run => run.Elements()).Select(child => new XElement(child)))
                        : field);
                }, partOverride: story is "header" or "footer" ? "word/" + story + "1.xml" : null);
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var expected = renderer.RenderPagesAsync(Source("7", "separate"), null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, expected.Sum(page => page.Images.Count));
            foreach (string cache in new[] { "stale", "7stale", "7", "Ⅶ", "" })
            foreach (string shape in new[] { "separate", "joined", "packed" })
            {
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var actual = renderer.RenderPagesAsync(Source(cache, shape), diagnostics.Add, default).ToBlockingEnumerable().ToArray();
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Cached counter spelling and run boundaries cannot discard neighboring math: " +
                    story + "/" + kind + "/" + cache + "/" + shape + "/" + mode);
                TestAssert.Equal(expected.Length, actual.Length);
                for (int page = 0; page < expected.Length; page++)
                {
                    TestAssert.Equal(expected[page].Content, actual[page].Content);
                    TestAssert.Equal(expected[page].Images.Count, actual[page].Images.Count);
                    TestAssert.True(expected[page].Annotations.SequenceEqual(actual[page].Annotations),
                        "Computed cache replacement retains picture and navigation source ownership.");
                }
            }
        }
    }

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

    public static void DocxSingleSectionPageNumberStartBindsComposedCounters()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
                mode, CachedMathGeometry(mode));
            var document = ReadSectionCounterReport(mode, start: 25, evenAndOdd: false);
            var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
            var pages = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            string first = SectionCounterPageText(pages[0]);
            TestAssert.Contains("onpage25(page25)", first);
            TestAssert.Contains("ionpage25", first);
            TestAssert.Contains("afterformula.Page25of2", first);
            TestAssert.Contains("afterformula.Page26of2", SectionCounterPageText(pages[1]));
            TestAssert.Contains("x+2y", first);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "The new field values preserve the adjacent formulas.");
        }
    }

    public static void DocxSingleSectionPageNumberStartSelectsStaticStoryParity()
    {
        var mode = OoxPdfDocxMarkupMode.Final;
        var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
            mode, CachedMathGeometry(mode));
        const string header = "Quarterlyoperationsreport";
        foreach (int start in new[] { 25, 26 })
        {
            var document = ReadSectionCounterReport(mode, start, evenAndOdd: true);
            var pages = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, pages.Length);
            int oddPage = start % 2 == 1 ? 0 : 1;
            string oddText = SectionCounterPageText(pages[oddPage]);
            TestAssert.Contains(header, HeaderText(pages[oddPage]));
            TestAssert.Contains($"afterformula.Page{start + oddPage}of2", oddText);
            // A missing even variant leaves the selected static frame empty.
            string evenText = SectionCounterPageText(pages[1 - oddPage]);
            TestAssert.DoesNotContain(header, HeaderText(pages[1 - oddPage]));
            TestAssert.DoesNotContain("afterformula.Page", evenText);
        }

        // The body repeats the title. Only the text above its top margin
        // witnesses the selected static header in this report.
        static string HeaderText(Lokad.OoxPdf.Pdf.PdfPage page) =>
            NormalizeNoteMathFieldText(string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Where(show => page.Height - show.Y < 60d).Select(show => show.Text)));
    }

    public static void DocxPageSeparatedSectionStartsBindSharedStoriesAndPhysicalDestinations()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true),
                mode, CachedMathGeometry(mode));
            var control = Render(null, null);
            int secondSectionPage = Array.FindIndex(control, page => SectionCounterPageText(page).Contains("Sectioncounter", StringComparison.Ordinal));
            TestAssert.True(secondSectionPage > 0, "The second section must begin on a later physical page.");
            var controlTargets = control.SelectMany(page => page.Annotations).Select(link => (link.Uri, link.Destination?.PageIndex)).ToArray();
            foreach (int? start in new int?[] { null, 40, 1 })
            {
                var pages = Render(25, start);
                TestAssert.True(control.Length == pages.Length,
                    $"A numbering start preserves physical pagination: {mode}/{start}, {control.Length} versus {pages.Length} pages.");
                var text = pages.Select(SectionCounterPageText).ToArray();
                TestAssert.Contains("onpage25(page25)", text[0]);
                TestAssert.Contains("ionpage25", text[0]);
                TestAssert.Contains("x+2y", text[0]);
                TestAssert.Contains($"Sectioncounter{start ?? 25 + secondSectionPage}", text[secondSectionPage]);
                for (int index = 0; index < pages.Length; index++)
                {
                    int expected = start is { } restart && index >= secondSectionPage
                        ? restart + index - secondSectionPage : 25 + index;
                    TestAssert.Contains($"afterformula.Page{expected}of{pages.Length}", text[index]);
                }
                TestAssert.True(pages.Select(page => page.Images.Count).SequenceEqual(control.Select(page => page.Images.Count)) &&
                    pages.Sum(page => page.Images.Count) > 0, "Restarted numbering preserves the pictures on their physical pages.");
                var targets = pages.SelectMany(page => page.Annotations).Select(link => (link.Uri, PageIndex: link.Destination?.PageIndex)).ToArray();
                TestAssert.True(targets.Any(target => target.PageIndex > 0) && targets.SequenceEqual(controlTargets),
                    "A displayed restart cannot renumber the PDF link targets.");
            }

            Lokad.OoxPdf.Pdf.PdfPage[] Render(int? firstStart, int? secondStart)
            {
                var doc = ReadMutation("Composition/composed-static-cell-report", mode, false, false, (xml, w, _) =>
                {
                    var body = xml.Root!.Element(w + "body")!;
                    var final = body.Element(w + "sectPr")!;
                    var first = new XElement(final);
                    final.Elements().Where(element => element.Name == w + "headerReference" ||
                        element.Name == w + "footerReference").Remove();
                    var boundary = Paragraph("Review section 6.");
                    var properties = boundary.Element(w + "pPr");
                    if (properties is null)
                    {
                        properties = new XElement(w + "pPr");
                        boundary.AddFirst(properties);
                    }
                    properties.Add(first);
                    Start(first, firstStart);
                    Start(final, secondStart);
                    Paragraph("Review section 7.").Add(
                        new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " Section counter ")),
                        new XElement(w + "fldSimple", new XAttribute(w + "instr", " PAGE "),
                            new XElement(w + "r", new XElement(w + "t", "900"))));

                    XElement Paragraph(string prefix) => body.Elements(w + "p").Single(paragraph =>
                        string.Concat(paragraph.Descendants(w + "t").Select(text => text.Value)).StartsWith(prefix, StringComparison.Ordinal));
                    void Start(XElement section, int? value)
                    {
                        if (value is { } number)
                            section.Add(new XElement(w + "pgNumType", new XAttribute(w + "start", number)));
                    }
                });
                return renderer.RenderPagesAsync(doc, null, default).ToBlockingEnumerable().ToArray();
            }
        }
    }

    private static DocxDocument ReadSectionCounterReport(OoxPdfDocxMarkupMode mode, int start, bool evenAndOdd) =>
        ReadMutation("Composition/composed-static-cell-report", mode, false, false, (xml, w, _) =>
        {
            XElement section = xml.Descendants(w + "sectPr").Single();
            section.Add(new XElement(w + "pgNumType", new XAttribute(w + "start", start)));
            if (evenAndOdd) section.Add(new XElement(w + "evenAndOddHeaders"));
        });

    private static string SectionCounterPageText(Lokad.OoxPdf.Pdf.PdfPage page) =>
        NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(page));

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
            var paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(document).ToArray();
            bool included = revision == "none" || (revision == "ins" ? mode != OoxPdfDocxMarkupMode.Original :
                mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup);
            TestAssert.Equal(included ? hidden ? 1 : 2 : 0, paragraphs.Sum(p => p.Images.Count));
            if (!included)
            {
                TestAssert.True(paragraphs.All(p => p.FieldReferences.Count == 0),
                    "An excluded revision cannot publish cached pictures or field metadata.");
                continue;
            }
            var paragraph = paragraphs.Single(p => p.FieldReferences.Count > 0);
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

    public static void DocxComputedFieldCacheContainersPreservePicturesAcrossStories()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (string kind in new[] { "PAGE", "NUMPAGES", "PAGE \\* MERGEFORMAT" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadComputedSimpleFieldPictures(story, kind, mode, true, 2),
                null, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, control.Sum(page => page.Images.Count));
            foreach (string container in new[] { "customXml", "smartTag", "sdt", "nested", "hyperlink" })
            {
                var wrapped = renderer.RenderPagesAsync(ReadComputedSimpleFieldPictures(story, kind, mode, true, 2,
                    container: container), null, default).ToBlockingEnumerable().ToArray();
                TestAssert.Equal(2, wrapped.Sum(page => page.Images.Count));
                TestAssert.Equal(control.Length, wrapped.Length);
                for (int index = 0; index < control.Length; index++)
                {
                    TestAssert.Equal(control[index].Content, wrapped[index].Content);
                    TestAssert.Equal(control[index].Fonts.Count, wrapped[index].Fonts.Count);
                    TestAssert.True(control[index].Annotations.SequenceEqual(wrapped[index].Annotations),
                        "Transparent cache wrappers preserve the existing equation and navigation owners.");
                }
            }
        }
    }

    public static void DocxComputedFieldCacheContainersKeepVisibilityAndOwnerBoundaries()
    {
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        foreach (string revision in new[] { "none", "ins", "del" })
        foreach (bool hidden in new[] { false, true })
        foreach (bool locked in new[] { false, true })
        {
            var control = ReadComputedSimpleFieldPictures("body", "PAGE", mode, true, 2, hidden, locked, revision);
            var expectedParagraphs = DocxBlockTraversal.EnumerateBodyParagraphs(control).ToArray();
            if (revision == "none")
                TestAssert.Equal(hidden ? 1 : 2, expectedParagraphs.Sum(paragraph => paragraph.Images.Count));
            foreach (string container in new[] { "customXml", "smartTag", "sdt", "nested", "hyperlink", "nested-field", "directional" })
            {
                var wrapped = ReadComputedSimpleFieldPictures("body", "PAGE", mode, true, 2, hidden, locked, revision, container);
                var paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(wrapped).ToArray();
                TestAssert.Equal(expectedParagraphs.Length, paragraphs.Length);
                for (int index = 0; index < paragraphs.Length; index++)
                {
                    var expected = expectedParagraphs[index];
                    var paragraph = paragraphs[index];
                    TestAssert.Equal(expected.Images.Count, paragraph.Images.Count);
                    int nestedFields = locked && container == "nested-field" && expected.FieldReferences.Count != 0 ? 1 : 0;
                    TestAssert.Equal(expected.FieldReferences.Count + nestedFields, paragraph.FieldReferences.Count);
                    TestAssert.True(paragraph.Images.Select(image => image.SourceTextOffsetInRun)
                        .SequenceEqual(expected.Images.Select(image => image.SourceTextOffsetInRun)),
                        "Cache containers retain each picture's source text offset.");
                    TestAssert.True(paragraph.Images.Select(image => image.SourceOrderInRun)
                        .SequenceEqual(expected.Images.Select(image => image.SourceOrderInRun)),
                        "Cache containers retain picture order at equal source offsets.");
                    TestAssert.Equal(DocxInlineContent.HasSupportedParagraphRevisionOwners(expected, default),
                        DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default));
                }
            }
        }
        foreach (string container in new[] { "unknown", "foreign", "misplaced-property" })
        {
            var document = ReadComputedSimpleFieldPictures("body", "PAGE", OoxPdfDocxMarkupMode.Final, true, 2,
                container: container);
            TestAssert.Equal(0, DocxBlockTraversal.EnumerateBodyParagraphs(document).Sum(paragraph => paragraph.Images.Count));
        }
    }

    private static DocxDocument ReadComputedSimpleFieldPictures(string story, string kind,
        OoxPdfDocxMarkupMode mode, bool split, int count, bool hiddenLast = false,
        bool locked = false, string revision = "none", string container = "none") =>
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
                if (container != "none")
                {
                    var children = field.Elements().Select(child => new XElement(child)).ToArray();
                    XElement Wrap(string name, params object[] content) => name == "sdt"
                        ? new XElement(w + name, new XElement(w + "sdtPr"), new XElement(w + "sdtContent", content))
                        : new XElement(w + name, new XElement(w + (name + "Pr")), content);
                    XElement[] wrapped = container switch
                    {
                        "hyperlink" => [children[0], new XElement(w + "hyperlink",
                            new XAttribute(w + "anchor", "PublicCalculation"), children.Skip(1))],
                        "nested" => [Wrap("sdt", Wrap("customXml", Wrap("smartTag", children)))],
                        "foreign" => [new XElement(XNamespace.Get("urn:ooxpdf:unknown") + "customXml", children)],
                        "nested-field" => [new XElement(w + "fldSimple", new XAttribute(w + "instr", "REF OtherOwner"), children)],
                        "directional" => [new XElement(w + "dir", new XAttribute(w + "val", "rtl"), children)],
                        "misplaced-property" => [new XElement(w + "customXmlPr", children)],
                        _ => [Wrap(container, children)]
                    };
                    field.ReplaceNodes(wrapped);
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
