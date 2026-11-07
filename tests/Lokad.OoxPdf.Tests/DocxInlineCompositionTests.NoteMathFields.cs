using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxNoteMathFieldsPrepareFromActualCounterOwners()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool direct in new[] { false, true })
        foreach (string instruction in new[] { "PAGE", "NUMPAGES" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadNoteMathField(kind, mode, instruction, "2", direct);
            var paragraph = NoteMathFieldParagraph(document);
            var owner = DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default)
                .Single(item => ReferenceEquals(item.Paragraph, paragraph));
            TestAssert.True(owner.HasPageInputs, "Normal note placement supplies actual counter inputs independently of math preparation.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph, owner.HasPageInputs).Count);
            TestAssert.True(paragraph.Runs.Any(run => run.FieldKind is DocxFieldKind.Page or DocxFieldKind.NumPages),
                "Preparing equation boxes must not materialize a guessed field value.");
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() =>
                DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, cancellation.Token).ToArray());
        }
    }

    public static void DocxNoteMathFieldsKeepLiteralContentAndGeometry()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool direct in new[] { false, true })
        foreach (string instruction in new[] { "PAGE", "NUMPAGES" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            // The source reference is on page1; the document has two pages.
            // Independent Word controls retain both formulas with PAGE1.
            string value = instruction == "PAGE" ? "1" : "2";
            var renderer = PictureRangeRenderer(mode);
            var literal = renderer.RenderPagesAsync(ReadNoteMathField(kind, mode, null, value, direct), null, default)
                .ToBlockingEnumerable().ToArray();
            var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
            var counter = renderer.RenderPagesAsync(ReadNoteMathField(kind, mode, instruction, "9", direct), diagnostics.Add, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(2, literal.Length); TestAssert.Equal(literal.Length, counter.Length);
            string Text(IEnumerable<PdfPage> pages) => NormalizeNoteMathFieldText(string.Concat(pages.SelectMany(page =>
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false)).Select(show => show.Text)));
            TestAssert.Equal(Text(literal), Text(counter));
            TestAssert.True(Text(counter).Contains("Counter" + value + ".", StringComparison.Ordinal),
                "The real page/count replaces the deliberately different cached9.");
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Known counter fields cannot discard supported math.");
            var a = literal.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
                fontResourceFilter: resource => resource.StartsWith("Eq", StringComparison.Ordinal))).ToArray();
            var b = counter.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
                fontResourceFilter: resource => resource.StartsWith("Eq", StringComparison.Ordinal))).ToArray();
            TestAssert.Equal("ratetimeratetime", NormalizeNoteMathFieldText(string.Concat(b.Select(show => show.Text))));
            TestAssert.Equal(a.Length, b.Length);
            for (int i = 0; i < a.Length; i++)
                TestAssert.True(a[i].Text == b[i].Text && Math.Abs(a[i].X - b[i].X) < .000001 &&
                    Math.Abs(a[i].Y - b[i].Y) < .000001 && a[i].Size == b[i].Size,
                    "A resolved source-equivalent field preserves each formula's placed glyph frame.");
            for (int i = 0; i < counter.Length; i++)
            {
                TestAssert.Equal(literal[i].Images.Count, counter[i].Images.Count);
                TestAssert.True(literal[i].Annotations.SequenceEqual(counter[i].Annotations), "Picture/link ownership survives counter binding.");
                PdfContentValidator.ValidatePage(counter[i], i, default);
            }
        }
    }

    public static void DocxNoteMathFieldsRetainUnsupportedAndCachedFallbacks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        {
            var cached = ReadNoteMathField(kind, OoxPdfDocxMarkupMode.Final, "AUTHOR", "1", false);
            var paragraph = NoteMathFieldParagraph(cached);
            TestAssert.True(DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(cached, default)
                .Any(item => ReferenceEquals(item.Paragraph, paragraph)), "An ordinary cached field keeps its existing literal fallback.");
            foreach (bool unknownStory in new[] { false, true })
            {
                var unsupported = ReadNoteMathField(kind, OoxPdfDocxMarkupMode.Final, unknownStory ? "PAGE" : "UNREGISTERED", "1", false,
                    includeResult: unknownStory, storyType: unknownStory ? "unregistered" : null);
                var source = NoteMathFieldParagraph(unsupported);
                TestAssert.Equal(2, source.Equations.Count);
                TestAssert.True(!DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(unsupported, default)
                    .Any(item => ReferenceEquals(item.Paragraph, source)), "Unknown instructions without a result and unregistered note owners keep their staged fallback.");
            }
        }
    }

    private static DocxParagraph NoteMathFieldParagraph(DocxDocument document) =>
        document.RelatedStories.Where(story => story.Id == "37").SelectMany(DocxBlockTraversal.EnumerateBodyParagraphs)
            .Single(paragraph => paragraph.Equations.Count != 0);

    private static string NormalizeNoteMathFieldText(string text) =>
        string.Concat(text.Normalize(NormalizationForm.FormKC).Where(character => !char.IsWhiteSpace(character)));

    private static DocxDocument ReadNoteMathField(string kind, OoxPdfDocxMarkupMode mode, string? instruction, string value,
        bool direct, bool includeResult = true, string? storyType = null) =>
        ReadMutation("CompositionProbes/mixed-note-" + kind + (direct ? "-paragraph" : "-table"), mode, false, false,
            (xml, w, _) =>
            {
                XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
                var note = xml.Root!.Elements(w + kind).Single(item => (string?)item.Attribute(w + "id") == "37");
                if (storyType is not null) note.SetAttributeValue(w + "type", storyType);
                var paragraph = note.Descendants(w + "p").First(item => item.Descendants(math + "oMath").Any());
                XElement Run(string text) => new(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
                paragraph.Add(Run(" Counter "));
                paragraph.Add(instruction is null ? Run(value) : new XElement(w + "fldSimple", new XAttribute(w + "instr", instruction),
                    includeResult ? Run(value) : null));
                paragraph.Add(Run("."));
            }, partOverride: "word/" + kind + "s.xml");
}
