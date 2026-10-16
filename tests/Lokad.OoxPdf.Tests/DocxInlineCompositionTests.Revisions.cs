using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionInsertedTextMathMatchesIndependentWord() =>
        AssertNumberedWordFrames("math-revisions-office-geometry.json");

    public static void DocxInlineCompositionInsertedTextMathPreservesSourceAndDeletionFallbacks()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native source controls use installed math fonts.");
        CheckTextRevisionMath(null);
    }

    internal static void CheckTextRevisionMath(IFontResolver? resolver)
    {
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (string kind in new[] { "ins", "del" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            bool supported = kind == "ins" || mode == OoxPdfDocxMarkupMode.Final;
            string? part = story == "header" ? "word/header1.xml" : story == "footer" ? "word/footer1.xml" : null;
            (double X, double Y, double Size)[]? original = null;
            foreach (bool metadata in new[] { false, true })
            {
                var source = ReadMutation("CompositionProbes/math-revision-" + story + "-" + kind, mode, metadata, metadata,
                    resolver is null ? null : (xml, w, m) =>
                    {
                        foreach (var run in xml.Descendants(m + "r"))
                        {
                            var props = run.Element(m + "rPr");
                            if (props is null) { props = new XElement(m + "rPr"); run.AddFirst(props); }
                            props.Elements(m + "nor").Remove(); props.Add(new XElement(m + "nor", new XAttribute(m + "val", "1")));
                        }
                    }, partOverride: part);
                var owner = story == "header" ? source.HeaderParagraphs.First(p => p.Equations.Count != 0) :
                    story == "footer" ? source.FooterParagraphs.First(p => p.Equations.Count != 0) :
                    DocxBlockTraversal.EnumerateBodyParagraphs(source.BodyElements).First(p => p.Equations.Count != 0);
                TestAssert.Equal(supported, owner.InlineContent is not null);
                TestAssert.True(owner.Equations.All(e => e.Revision is null), "Revising adjacent text cannot change equation ownership.");
                var page = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                    .RenderBlankPages(source, null, default).Single();
                var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
                var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
                TestAssert.Equal(supported ? 18 : 14, math.Length);
                string text = string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(supported ? 2 : 1, text.Split("x+2y", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(2, text.Split("a+b=c", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(1, page.Images.Count);
                TestAssert.Contains("share the same input series.", DocxPageFootnoteNumberingTests.DecodeText(page));
                if (original is not null)
                    for (int i = 0; i < math.Length; i++)
                        TestAssert.True(Math.Abs(original[i].X - math[i].X) < .01d && Math.Abs(original[i].Y - math[i].Y) < .01d && original[i].Size == math[i].Size,
                            "Bookmark/run-split metadata preserves inserted-text math geometry and the staged deletion fallback.");
                original = math.Select(show => (show.X, show.Y, show.Size)).ToArray();
            }
        }
    }

    public static void DocxInlineCompositionTextRevisionOwnershipKeepsNontextFallbacksAndCancellation()
    {
        var doc = ReadMutation("CompositionProbes/math-revision-body-ins", OoxPdfDocxMarkupMode.Final, false, false);
        var paragraph = doc.Paragraphs.First(p => p.Equations.Count != 0);
        var revision = paragraph.Revisions.Single();
        TestAssert.True(DocxInlineContent.HasSupportedTextRevisionOwners(paragraph.Revisions, paragraph.Runs, default),
            "The same parsed insertion object owns its ordinary resolved text run.");
        foreach (var unsupported in new[] { revision with { }, revision with { Kind = DocxRevisionKind.Deletion },
            revision with { Kind = DocxRevisionKind.ParagraphPropertiesChange }, revision with { Kind = DocxRevisionKind.MoveTo } })
            TestAssert.True(!DocxInlineContent.HasSupportedTextRevisionOwners([unsupported], paragraph.Runs, default),
                "Equal-looking metadata without an actual text owner cannot qualify a paragraph mark/property or another revision.");
        var equations = paragraph.Equations.Select(e => e with { Revision = revision }).ToArray();
        TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph with { Equations = equations }).Count);
        var unowned = paragraph with { Revisions = [revision with { }] };
        TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(unowned).Count);
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedTextRevisionOwners(paragraph.Revisions, paragraph.Runs, canceled.Token));
    }
}
