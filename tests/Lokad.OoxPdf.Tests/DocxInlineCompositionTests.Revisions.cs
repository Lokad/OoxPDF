using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionListLabelsKeepTextRevisionFormattingOwners()
    {
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (string kind in new[] { "ins", "del" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool annotated in new[] { false, true })
        {
            string? part = story == "header" ? "word/header1.xml" : story == "footer" ? "word/footer1.xml" : null;
            var document = ReadMutation("CompositionProbes/math-revision-" + story + "-" + kind,
                mode, annotated, annotated, partOverride: part);
            var paragraph = story == "header" ? document.HeaderParagraphs.First(p => p.Equations.Count != 0) :
                story == "footer" ? document.FooterParagraphs.First(p => p.Equations.Count != 0) :
                DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements).First(p => p.Equations.Count != 0);
            var text = paragraph.Runs[0];
            var label = DocxLayoutEngine.CreateListLabelRun(paragraph.ListLabel!, text, 11d);
            TestAssert.True(!label.Underline && !label.Strike && !label.DoubleStrike &&
                label.Revision is null && label.Revisions.Count == 0,
                "Unchanged labels do not acquire the first text run's revision decorations: " + story + "/" + kind + "/" + mode + "/" + annotated);
            TestAssert.Equal(text.FontFamily, label.FontFamily);
            if (mode == OoxPdfDocxMarkupMode.AllMarkup)
                TestAssert.True(text.Revision is not null && (kind == "ins" ? text.Underline : text.Strike),
                    "Restoring label inheritance must leave revised source text and its provenance intact.");
        }
    }

    public static void DocxInlineCompositionListLabelsRetainParagraphRevisionOwners()
    {
        foreach (bool paragraphMark in new[] { false, true })
        {
            var document = ReadMutation("CompositionProbes/math-revision-body-ins", OoxPdfDocxMarkupMode.AllMarkup,
                false, false, (xml, w, m) =>
                {
                    var paragraph = xml.Descendants(w + "p").First(p => p.Descendants(m + "oMath").Any());
                    if (paragraphMark)
                    {
                        var mark = new XElement(w + "rPr", new XElement(w + "ins", new XAttribute(w + "id", "93")));
                        paragraph.Element(w + "pPr")!.Add(mark);
                    }
                    else
                    {
                        var insertion = paragraph.Element(w + "ins")!;
                        insertion.ReplaceWith(insertion.Nodes().ToArray());
                        paragraph.ReplaceWith(new XElement(w + "ins", new XAttribute(w + "id", "94"), new XElement(paragraph)));
                    }
                });
            var paragraph = document.Paragraphs.First(p => p.Equations.Count != 0);
            var label = DocxLayoutEngine.CreateListLabelRun(paragraph.ListLabel!, paragraph.Runs[0], 11d);
            TestAssert.True(paragraph.ListLabel!.UnmarkedTextStyle is null && label.Revision is not null && label.Underline,
                "A paragraph/mark revision can own the label; retain that existing path rather than borrowing text-only ownership.");
        }
    }

    public static void DocxInlineCompositionListLabelOverridesKeepAuthoredFormattingAndScaledMetrics()
    {
        var document = ReadMutation("CompositionProbes/math-revision-body-ins", OoxPdfDocxMarkupMode.AllMarkup,
            false, false, (xml, w, m) =>
            {
                var run = xml.Descendants(w + "ins").First().Element(w + "r")!;
                var properties = run.Element(w + "rPr");
                if (properties is null) { properties = new XElement(w + "rPr"); run.AddFirst(properties); }
                properties.Add(new XElement(w + "color", new XAttribute(w + "val", "1964AE")),
                    new XElement(w + "u", new XAttribute(w + "val", "double")), new XElement(w + "b"));
            });
        var paragraph = document.Paragraphs.First(p => p.Equations.Count != 0);
        var text = paragraph.Runs[0] with { FontSize = 7d, CharacterSpacingPoints = .25d, PairKerningEnabled = false };
        var label = DocxLayoutEngine.CreateListLabelRun(paragraph.ListLabel!, text, 7d);
        TestAssert.Equal("1964AE", label.ColorHex);
        TestAssert.True(label.Underline && label.UnderlineValue == "double" && label.Bold && label.Revision is null,
            "Authored formatting survives independently of the added review decoration.");
        TestAssert.Equal(.25d, label.CharacterSpacingPoints);
        TestAssert.True(!label.PairKerningEnabled, "Placement-time spacing/kerning decisions remain on the actual base run.");
        var overridden = paragraph.ListLabel! with { Style = paragraph.ListLabel.Style with { ColorHex = "008844", Underline = false, UnderlineValue = "none", FontSize = 16d } };
        label = DocxLayoutEngine.CreateListLabelRun(overridden, text, 7d);
        TestAssert.Equal("008844", label.ColorHex);
        TestAssert.Equal(16d, label.FontSize);
        TestAssert.True(!label.Underline && label.UnderlineValue == "none", "Numbering-level formatting keeps precedence.");
    }

    public static void DocxInlineCompositionTextRevisionMathMatchesIndependentWord() =>
        AssertNumberedWordFrames("math-revisions-office-geometry.json");

    public static void DocxInlineCompositionTextRevisionMathPreservesSourceAndFallbacks()
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
                TestAssert.True(owner.InlineContent is not null, "Actual inline insertion/deletion owners preserve the shared source view.");
                TestAssert.True(owner.Equations.All(e => e.Revision is null), "Revising adjacent text cannot change equation ownership.");
                var page = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                    .RenderBlankPages(source, null, default).Single();
                var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
                var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
                TestAssert.Equal(18, math.Length);
                string text = string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(2, text.Split("x+2y", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(2, text.Split("a+b=c", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(1, page.Images.Count);
                TestAssert.Contains("share the same input series.", DocxPageFootnoteNumberingTests.DecodeText(page));
                if (original is not null)
                    for (int i = 0; i < math.Length; i++)
                        TestAssert.True(Math.Abs(original[i].X - math[i].X) < .01d && Math.Abs(original[i].Y - math[i].Y) < .01d && original[i].Size == math[i].Size,
                            "Bookmark/run-split metadata preserves inserted-text math geometry and deletion text geometry.");
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
        var deleted = ReadMutation("CompositionProbes/math-revision-body-del", OoxPdfDocxMarkupMode.AllMarkup, false, false)
            .Paragraphs.First(p => p.Equations.Count != 0);
        TestAssert.True(DocxInlineContent.HasSupportedTextRevisionOwners(deleted.Revisions, deleted.Runs, default),
            "Materialized deletion text owns its actual revision in the declared inline view.");
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
