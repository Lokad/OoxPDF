using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNoteMathFieldsIndependentStaticStoriesPreserveMixedFrames() => CheckIndependentStaticNoteFrames(false);

    public static void DocxNoteMathFieldsStaticCountersPreserveMixedFrames() => CheckIndependentStaticNoteFrames(true);

    public static void DocxNoteMathFieldsStaticPicturesAndCountersPreserveMixedFrames() => CheckIndependentStaticNoteFrames(true, true);

    public static void DocxNoteMathFieldsPictureOnlyStaticStoriesPreserveMixedFrames() => CheckIndependentStaticNoteFrames(false, true, true);

    private static void CheckIndependentStaticNoteFrames(bool staticCounters, bool staticPictures = false, bool pictureOnly = false)
    {
        foreach (string position in new[] { "docEnd", "sectEnd" })
        foreach (bool direct in new[] { false, true })
        foreach (bool footer in new[] { false, true })
        foreach (string instruction in new[] { "PAGE", "NUMPAGES" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            // Footer variants also move the rendered reference to page2.
            int referencePage = footer ? 2 : 1, pageCount = footer ? 3 : 2;
            string value = (instruction == "PAGE" ? referencePage : pageCount).ToString(System.Globalization.CultureInfo.InvariantCulture);
            DocxDocument Source(string? field, string cached, bool withStatic)
            {
                var source = ReadNoteMathField("endnote", mode, field, cached, direct);
                source = source with
                {
                    Settings = source.Settings with { EndnoteReferenceSettings =
                        source.Settings.EndnoteReferenceSettings with { PositionValue = position } },
                    BodyElements = footer
                        ? [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public earlier page", 12d, 14d)),
                            new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), .. source.BodyElements]
                        : source.BodyElements
                };
                if (!withStatic) return source;
                var paragraph = staticCounters ? ReadNoteMathStaticCounters(mode, complex: footer) : PlainNoteMathStaticParagraph();
                if (staticPictures) paragraph = NoteMathStaticPictureParagraph(source, paragraph, pictureOnly) with
                    { Alignment = footer ? DocxTextAlignment.Center : DocxTextAlignment.Left };
                return WithNoteMathStaticStory(source, footer, paragraph);
            }
            var renderer = PictureRangeRenderer(mode);
            var baseline = renderer.RenderPagesAsync(Source(null, value, false), null, default).ToBlockingEnumerable().ToArray();
            var literal = renderer.RenderPagesAsync(Source(null, value, true), null, default).ToBlockingEnumerable().ToArray();
            var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
            var source = Source(instruction, "9", true);
            var runs = NoteMathFieldParagraph(source).Runs.Select(run => (run.Text, run.FieldKind)).ToArray();
            var counter = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            string context = $"{position}/direct={direct}/footer={footer}/{instruction}/{mode}";
            TestAssert.Equal(pageCount, counter.Length); TestAssert.Equal(counter.Length, literal.Length);
            TestAssert.Equal(counter.Length, baseline.Length);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "An independent margin story cannot discard note math: " + context);
            TestAssert.True(runs.SequenceEqual(NoteMathFieldParagraph(source).Runs.Select(run => (run.Text, run.FieldKind))),
                "The measured counter does not replace cached source data.");
            var mathText = new System.Text.StringBuilder();
            for (int page = 0; page < counter.Length; page++)
            {
                TestAssert.Equal(NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(literal[page])),
                    NormalizeNoteMathFieldText(DocxPageFootnoteNumberingTests.DecodeText(counter[page])));
                var a = Math(baseline[page]); var b = Math(literal[page]); var c = Math(counter[page]);
                TestAssert.True(a.SequenceEqual(b) && a.SequenceEqual(c), "Independent static content preserves every note math frame: " + context);
                mathText.Append(string.Concat(c.Select(glyph => glyph.Text)));
                TestAssert.Equal(literal[page].Images.Count, counter[page].Images.Count);
                TestAssert.True(!staticPictures || counter[page].Images.Count > 0, "The selected static picture renders on every page.");
                TestAssert.True(baseline[page].Annotations.SequenceEqual(counter[page].Annotations), "Navigation follows the unchanged note geometry: " + context);
                PdfContentValidator.ValidatePage(counter[page], page, default);
            }
            TestAssert.Equal("ratetimeratetime", NormalizeNoteMathFieldText(mathText.ToString()));
            TestAssert.True(pictureOnly || string.Concat(counter.Select(DocxPageFootnoteNumberingTests.DecodeText)).Contains("Public margin story", StringComparison.Ordinal),
                "The independent story actually renders.");

            static (string Text, double X, double Y, double Size)[] Math(PdfPage page) =>
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
                    fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal))
                    .Select(glyph => (glyph.Text, glyph.X, glyph.Y, glyph.Size)).ToArray();
        }
    }

    public static void DocxNoteMathFieldsIntrusiveAndUnsupportedStaticFramesRetainFallback()
    {
        foreach (string position in new[] { "docEnd", "sectEnd" })
        foreach (bool footer in new[] { false, true })
        foreach (bool intrusive in new[] { false, true })
        {
            var source = ReadNoteMathField("endnote", OoxPdfDocxMarkupMode.Final, "PAGE", "9", false);
            source = source with { Settings = source.Settings with { EndnoteReferenceSettings =
                source.Settings.EndnoteReferenceSettings with { PositionValue = position } } };
            var paragraph = PlainNoteMathStaticParagraph();
            if (!intrusive) paragraph = paragraph with { LineSpacingPoints = 14d };
            source = WithNoteMathStaticStory(source, footer, paragraph, intrusive ? 120d : 36d);
            var layout = CreatePageCounterLayout(source);
            var notes = layout.Pages.SelectMany(page => page.PlacedRelatedStories).Where(note => note.StoryLayout.Story.Id == "37").ToArray();
            TestAssert.True(notes.Length > 0 && notes.All(note => !note.StoryLayout.HasResolvedCounterInputs && !note.UsesCompleteDocumentEndnoteFlow),
                $"An intrusive or unsupported static frame retains the existing fallback: {position}/footer={footer}/intrusive={intrusive}.");
        }
    }

    public static void DocxNoteMathFieldsStaticPicturesRespectSelectedUnsupportedAndIntrusiveOwners()
    {
        foreach (string position in new[] { "docEnd", "sectEnd" })
        foreach (bool footer in new[] { false, true })
        foreach (string variant in new[] { "rotated", "unowned", "tall", "earlier-rotated" })
        {
            var source = ReadNoteMathField("endnote", OoxPdfDocxMarkupMode.Final, "PAGE", "9", false);
            source = source with { Settings = source.Settings with { EndnoteReferenceSettings =
                source.Settings.EndnoteReferenceSettings with { PositionValue = position } } };
            var paragraph = NoteMathStaticPictureParagraph(source, PlainNoteMathStaticParagraph(), false);
            var image = paragraph.Images[0];
            image = variant switch
            {
                "rotated" or "earlier-rotated" => image with { RotationDegrees = 90d },
                "unowned" => image with { SourceRunIndex = -1 },
                _ => image with { HeightPoints = 200d }
            };
            paragraph = paragraph with { Images = [image], InlineContent = DocxInlineContent.FromPrefix(paragraph.Runs, [image], default) };
            source = WithNoteMathStaticStory(source, footer, paragraph);
            if (variant == "earlier-rotated") source = source with
            {
                HeaderParagraphs = [], FooterParagraphs = [], HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(),
                FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(),
                HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(),
                FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(),
                PageSettings = source.PageSettings with
                {
                    HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(),
                    FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(),
                    HeaderBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(),
                    FooterBodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>()
                }
            };
            var notes = CreatePageCounterLayout(source).Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(note => note.StoryLayout.Story.Id == "37").ToArray();
            TestAssert.True(notes.Length > 0 && notes.All(note => !note.StoryLayout.HasResolvedCounterInputs && !note.UsesCompleteDocumentEndnoteFlow),
                $"Selected unsupported/intrusive static pictures retain fallback independently of final document-level collections: {position}/{footer}/{variant}.");
        }
    }

    private static DocxParagraph NoteMathStaticPictureParagraph(DocxDocument source, DocxParagraph paragraph, bool pictureOnly)
    {
        var image = source.RelatedStories.Where(story => story.Id == "37").SelectMany(DocxBlockTraversal.EnumerateBodyParagraphs)
            .SelectMany(p => p.Images).First() with { WidthPoints = 18d, HeightPoints = 12d, SourceRunIndex = 0,
                SourceTextOffsetInRun = 0, ClickTarget = null, PictureFrame = null, Crop = default };
        DocxTextRun[] runs = pictureOnly ? [] : paragraph.Runs.Select((run, index) => run with { SourceRunIndex = index + 1 }).ToArray();
        return paragraph with { Runs = runs, Images = [image], InlineContent = DocxInlineContent.FromPrefix(runs, [image], default) };
    }

    public static void DocxNoteMathFieldsStaticCountersRetainUnknownAndInconsistentFallbacks()
    {
        foreach (string position in new[] { "docEnd", "sectEnd" })
        foreach (bool footer in new[] { false, true })
        foreach (bool unknown in new[] { false, true })
        {
            var source = ReadNoteMathField("endnote", OoxPdfDocxMarkupMode.Final, "PAGE", "9", false);
            source = source with { Settings = source.Settings with { EndnoteReferenceSettings =
                source.Settings.EndnoteReferenceSettings with { PositionValue = position } } };
            var paragraph = ReadNoteMathStaticCounters(OoxPdfDocxMarkupMode.Final, complex: footer);
            var fields = paragraph.FieldReferences.ToArray();
            fields[0] = unknown ? fields[0] with { Kind = DocxFieldKind.Other }
                : fields[0] with { Placeholder = "Unowned counter result" };
            source = WithNoteMathStaticStory(source, footer, paragraph with { FieldReferences = fields });
            var notes = CreatePageCounterLayout(source).Pages.SelectMany(page => page.PlacedRelatedStories)
                .Where(note => note.StoryLayout.Story.Id == "37").ToArray();
            TestAssert.True(notes.Length > 0 && notes.All(note => !note.StoryLayout.HasResolvedCounterInputs && !note.UsesCompleteDocumentEndnoteFlow),
                "Actual page inputs cannot admit unknown or inconsistent static field owners.");
        }
    }

    private static DocxParagraph ReadNoteMathStaticCounters(OoxPdfDocxMarkupMode mode, bool complex) =>
        ReadMutation("CompositionProbes/mixed-note-endnote-table", mode, false, false, (xml, w, _) =>
        {
            var paragraph = xml.Descendants(w + "body").Elements(w + "p").First();
            paragraph.RemoveNodes();
            paragraph.Add(new XElement(w + "pPr", new XElement(w + "spacing", new XAttribute(w + "after", "0"))));
            XElement Run(string text) => new(w + "r", new XElement(w + "rPr",
                new XElement(w + "rFonts", new XAttribute(w + "ascii", "Calibri"), new XAttribute(w + "hAnsi", "Calibri")),
                new XElement(w + "sz", new XAttribute(w + "val", "24"))),
                new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
            void Field(string instruction)
            {
                if (!complex) paragraph.Add(new XElement(w + "fldSimple", new XAttribute(w + "instr", instruction), Run("9")));
                else paragraph.Add(new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "begin"))),
                    new XElement(w + "r", new XElement(w + "instrText", " " + instruction + " ")),
                    new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "separate"))), Run("9"),
                    new XElement(w + "r", new XElement(w + "fldChar", new XAttribute(w + "fldCharType", "end"))));
            }
            paragraph.Add(Run("Public margin story ")); Field("PAGE"); paragraph.Add(Run(" of ")); Field("NUMPAGES");
        }).Paragraphs.First();

    private static DocxParagraph PlainNoteMathStaticParagraph() =>
        DocxTests.CreateDocxLayoutParagraph("Public margin story", 12d, 14d) with { LineSpacingPoints = null };

    private static DocxDocument WithNoteMathStaticStory(DocxDocument source, bool footer, DocxParagraph paragraph, double distance = 36d)
    {
        var paragraphs = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [paragraph] };
        var elements = new Dictionary<string, IReadOnlyList<DocxBodyElement>> { ["default"] = [new DocxParagraphElement(paragraph)] };
        DocxPageSettings Settings(DocxPageSettings settings) => footer
            ? settings with { FooterDistancePoints = distance, FooterParagraphsByType = paragraphs, FooterBodyElementsByType = elements }
            : settings with { HeaderDistancePoints = distance, HeaderParagraphsByType = paragraphs, HeaderBodyElementsByType = elements };
        return source with
        {
            HeaderParagraphs = footer ? source.HeaderParagraphs : [paragraph],
            FooterParagraphs = footer ? [paragraph] : source.FooterParagraphs,
            HeaderParagraphsByType = footer ? source.HeaderParagraphsByType : paragraphs,
            FooterParagraphsByType = footer ? paragraphs : source.FooterParagraphsByType,
            HeaderBodyElementsByType = footer ? source.HeaderBodyElementsByType : elements,
            FooterBodyElementsByType = footer ? elements : source.FooterBodyElementsByType,
            PageSettings = Settings(source.PageSettings),
            BodyElements = source.BodyElements.Select(element => element is DocxSectionBreakElement section
                ? section with { PageSettings = Settings(section.PageSettings) } : element).ToArray()
        };
    }
}
