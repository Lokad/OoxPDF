using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionRunFormattingRetainsAuthoredInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native controls use installed fonts.");
        CheckRunFormattingInk(null);
    }

    internal static void CheckRunFormattingInk(IFontResolver? resolver)
    {
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (bool authoredColor in new[] { false, true })
        {
            string? part = story == "header" ? "word/header1.xml" : story == "footer" ? "word/footer1.xml" : null;
            var document = ReadMutation("CompositionProbes/math-run-format-" + story + "-bold",
                OoxPdfDocxMarkupMode.AllMarkup, false, false, (xml, w, m) =>
                {
                    if (authoredColor)
                    {
                        var properties = xml.Descendants(w + "rPrChange").Single().Parent!;
                        properties.Add(new XElement(w + "color", new XAttribute(w + "val", "1964AE")),
                            new XElement(w + "u", new XAttribute(w + "val", "single")));
                    }
                    if (resolver is not null)
                        foreach (var run in xml.Descendants(m + "r"))
                        {
                            var properties = run.Element(m + "rPr");
                            if (properties is null) { properties = new XElement(m + "rPr"); run.AddFirst(properties); }
                            properties.Elements(m + "nor").Remove(); properties.Add(new XElement(m + "nor", new XAttribute(m + "val", "1")));
                        }
                }, partOverride: part);
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).Single();
            var targetText = new HashSet<string>(["1.", "Summary: "], StringComparer.Ordinal);
            var matched = new HashSet<string>(StringComparer.Ordinal);
            foreach (Match block in Regex.Matches(page.Content, @"(?m)^BT\r?\n.*?^ET\r?$", RegexOptions.Singleline))
            {
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(new PdfPage(page.Width, page.Height, block.Value, page.Fonts)).ToArray();
                if (shows.Length != 1 || !targetText.Contains(shows[0].Text)) continue;
                string expected = authoredColor ? "0.098 0.392 0.682 rg" : "0 g";
                TestAssert.Contains(expected, block.Value);
                matched.Add(shows[0].Text);
            }
            TestAssert.Equal(2, matched.Count);
            if (authoredColor)
                TestAssert.True(Regex.IsMatch(page.Content, @"0\.098 0\.392 0\.682 rg\s+[-\d.]+ [-\d.]+ [\d.]+ [\d.]+ re f"),
                    "Authored underline ink survives formatting metadata as well as glyph ink.");
            TestAssert.True(page.Content.Contains("0.949 g", StringComparison.Ordinal),
                "Retaining authored ink cannot remove the formatting balloon's review lane.");
        }
    }

    public static void DocxInlineCompositionRunFormattingMathMatchesIndependentWord() =>
        AssertNumberedWordFrames("math-run-format-office-geometry.json");

    public static void DocxInlineCompositionRunFormattingMathPreservesSourceAndMetadata()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native controls use installed math fonts.");
        CheckRunFormattingMath(null);
    }

    internal static void CheckRunFormattingMath(IFontResolver? resolver)
    {
        foreach (var (story, kind) in new[] { ("body", "bold"), ("cell", "bold"), ("header", "bold"), ("footer", "bold"), ("body", "void") })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string? part = story == "header" ? "word/header1.xml" : story == "footer" ? "word/footer1.xml" : null;
            (double X, double Y, double Size)[]? before = null;
            foreach (bool metadata in new[] { false, true })
            {
                var source = ReadMutation("CompositionProbes/math-run-format-" + story + "-" + kind,
                    mode, metadata, metadata, resolver is null ? null : (xml, w, m) =>
                    {
                        foreach (var run in xml.Descendants(m + "r"))
                        {
                            var properties = run.Element(m + "rPr");
                            if (properties is null) { properties = new XElement(m + "rPr"); run.AddFirst(properties); }
                            properties.Elements(m + "nor").Remove();
                            properties.Add(new XElement(m + "nor", new XAttribute(m + "val", "1")));
                        }
                    }, partOverride: part);
                var owner = story == "header" ? source.HeaderParagraphs.First(p => p.Equations.Count != 0) :
                    story == "footer" ? source.FooterParagraphs.First(p => p.Equations.Count != 0) :
                    DocxBlockTraversal.EnumerateBodyParagraphs(source.BodyElements).First(p => p.Equations.Count != 0);
                TestAssert.True(owner.InlineContent is not null && owner.Equations.All(e => e.Revision is null),
                    "Run formatting metadata cannot change adjacent equation source ownership.");
                TestAssert.True(kind == "void" ? owner.Revisions.Count == 0 : owner.Revisions.Count > 0,
                    "Void changes are discarded; split runs retain their actual formatting metadata.");
                if (kind != "void")
                {
                    TestAssert.True(owner.Runs[0].Bold && owner.Runs[0].Revisions.Any(r => ReferenceEquals(r, owner.Revisions[0])),
                        "Current run formatting and the actual change owner survive source publication.");
                }
                var page = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                    .RenderBlankPages(source, null, default).Single();
                var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
                var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
                TestAssert.Equal(18, math.Length);
                TestAssert.Equal("x+2ya+b=cx+2ya+b=c", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(1, page.Images.Count);
                TestAssert.Contains("share the same input series.", DocxPageFootnoteNumberingTests.DecodeText(page));
                if (before is not null)
                    for (int i = 0; i < math.Length; i++)
                        TestAssert.True(Math.Abs(before[i].X - math[i].X) < .01d && Math.Abs(before[i].Y - math[i].Y) < .01d && before[i].Size == math[i].Size,
                            "Bookmark/run splitting preserves formatting-revision equation geometry.");
                before = math.Select(show => (show.X, show.Y, show.Size)).ToArray();
            }
        }
    }

    public static void DocxInlineCompositionRunFormattingOwnershipKeepsNonRunFallbacks()
    {
        var paragraph = ReadMutation("CompositionProbes/math-run-format-body-bold", OoxPdfDocxMarkupMode.AllMarkup, false, false)
            .Paragraphs.First(p => p.Equations.Count != 0);
        var change = paragraph.Revisions.Single();
        TestAssert.True(DocxInlineContent.HasSupportedRunRevisionOwners([change], paragraph.Runs, default),
            "Actual resolved run-formatting owners qualify beside unchanged equations.");
        TestAssert.True(!DocxInlineContent.HasSupportedRunRevisionOwners([change with { }], paragraph.Runs, default),
            "A value-equal metadata copy is not an actual source owner.");
        foreach (var kind in new[] { DocxRevisionKind.ParagraphPropertiesChange, DocxRevisionKind.MoveTo })
        {
            var unsupported = change with { Kind = kind };
            var runs = paragraph.Runs.Select(run => run with { Revisions = [unsupported] }).ToArray();
            TestAssert.True(!DocxInlineContent.HasSupportedRunRevisionOwners([unsupported], runs, default),
                "Paragraph properties and moves retain their fallback even when attached to a text run.");
        }
        var empty = paragraph.Runs.Select(run => run with { Text = string.Empty }).ToArray();
        TestAssert.True(!DocxInlineContent.HasSupportedRunRevisionOwners([change], empty, default),
            "Empty/nontext owners cannot qualify a revision.");
        TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph with {
            Equations = paragraph.Equations.Select(e => e with { Revision = change }).ToArray() }).Count);
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedRunRevisionOwners([change], paragraph.Runs, canceled.Token));
    }
}
