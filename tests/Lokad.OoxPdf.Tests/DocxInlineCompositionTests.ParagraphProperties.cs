using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxParagraphPropertyMathKeepsResolvedSourceAcrossPlacementsAndViews()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var source = ReadParagraphPropertyMath(container, mode);
            var paragraph = CachedMathParagraph(source);
            var revision = TestAssert.NotNull(paragraph.PropertyRevision);
            TestAssert.Equal(DocxRevisionKind.ParagraphPropertiesChange, revision.Kind);
            TestAssert.Equal(DocxRevisionPropertyFamily.Paragraph, revision.PropertyChangeFamily!.Value);
            TestAssert.Equal("pPrChange", revision.SourceElement);
            TestAssert.True(paragraph.Revisions.Any(item => ReferenceEquals(item, revision)),
                "The already-parsed property change retains its actual owner in aggregate inspection metadata.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            var baseline = CachedMathParagraph(ReadParagraphPropertyMath(container, mode, baseline: true));
            TestAssert.Equal(baseline.Alignment, paragraph.Alignment);
            TestAssert.Equal(baseline.SpacingBeforePoints, paragraph.SpacingBeforePoints);
            TestAssert.Equal(baseline.SpacingAfterPoints, paragraph.SpacingAfterPoints);
            TestAssert.Equal(baseline.LineSpacingFactor, paragraph.LineSpacingFactor);
            TestAssert.Equal(baseline.Indent, paragraph.Indent);
            var unchanged = CachedMathParagraph(ReadParagraphPropertyMath(container, mode, noOp: true));
            TestAssert.True(unchanged.PropertyRevision is null && unchanged.Revisions.All(item => item.Id != "975"),
                "Existing semantic-void filtering remains at the reader boundary.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(unchanged).Count);
        }
    }

    public static void DocxParagraphPropertyMathRequiresItsActualUniquePropertyOwner()
    {
        var paragraph = CachedMathParagraph(ReadParagraphPropertyMath("body", OoxPdfDocxMarkupMode.Final));
        var revision = TestAssert.NotNull(paragraph.PropertyRevision);
        foreach (var invalid in new[]
        {
            paragraph with { PropertyRevision = null },
            paragraph with { PropertyRevision = revision with { } },
            paragraph with { Revisions = paragraph.Revisions.Select(item => ReferenceEquals(item, revision) ? revision with { } : item).ToArray() },
            paragraph with { Revisions = paragraph.Revisions.Append(revision with { Id = "976" }).ToArray() }
        })
            TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(invalid, default),
                "Missing, value-equal or extra aggregate metadata cannot invent a resolved property owner.");
        foreach (var invalid in new[] { revision with { SourceElement = "rPrChange" },
            revision with { PropertyChangeFamily = DocxRevisionPropertyFamily.Run }, revision with { Kind = DocxRevisionKind.MoveTo } })
            TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with
                { PropertyRevision = invalid, Revisions = [invalid] }, default),
                "The exact source owner still requires paragraph-property provenance and a supported revision kind.");
        var duplicate = CachedMathParagraph(ReadParagraphPropertyMath("body", OoxPdfDocxMarkupMode.Final, duplicate: true));
        TestAssert.True(duplicate.PropertyRevision is null, "Duplicate property changes retain their existing conservative fallback.");
        TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(duplicate).Count);
        TestAssert.True(!DocxInlineContent.HasSupportedRunRevisionOwners([revision], paragraph.Runs, default),
            "Paragraph formatting is not an ordinary run-owned revision.");
        TestAssert.True(!DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph),
            "This source correction does not silently broaden note, line-height or markup-lane correction guards.");
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, canceled.Token));
    }

    public static void DocxParagraphPropertyMathPreservesPaintAndNavigation()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var expected = renderer.RenderPagesAsync(ReadParagraphPropertyMath(container, mode, baseline: true, navigation: true), null, default)
                .ToBlockingEnumerable().ToArray();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(ReadParagraphPropertyMath(container, mode, navigation: true), diagnostics.Add, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(expected.Length, actual.Length);
            string MathText(IEnumerable<PdfPage> pages) => string.Concat(pages.SelectMany(page =>
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(show => page.Fonts.Any(font =>
                    font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))))
                .Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal("ratetimeratetime", MathText(actual));
            TestAssert.Equal(MathText(expected), MathText(actual));
            TestAssert.True(expected.SelectMany(page => page.Annotations).Any(annotation => annotation.Destination is not null),
                "The navigation control must contain an actual equation link.");
            TestAssert.True(expected.SelectMany(page => page.Annotations).Select(a => (a.Uri, a.Destination))
                .SequenceEqual(actual.SelectMany(page => page.Annotations).Select(a => (a.Uri, a.Destination))),
                "Formatting history must retain the neighboring equation link and bookmark destination.");
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Resolved paragraph-property history must not turn supported math into ignored content.");
        }
    }

    public static void DocxParagraphPropertyMathKeepsJoinedParagraphPropertyOwnership()
    {
        foreach (bool first in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.SimpleMarkup })
        {
            var paragraph = CachedMathParagraph(ReadParagraphPropertyMath("body", mode, joinMathFirst: first));
            var revision = paragraph.Revisions.Single(item => item.Id == "975");
            TestAssert.True(paragraph.ConsumedParagraphMarks.Any(mark => mark.Revisions.Any(item => item.Id == "976")),
                "The control must actually join paragraphs across a deleted source mark.");
            TestAssert.Equal(first, ReferenceEquals(revision, paragraph.PropertyRevision));
            TestAssert.Equal(first ? 2 : 0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            // The existing join retains the first paragraph's resolved properties.
            // A second paragraph's aggregate formatting history has no such owner.
        }
    }

    public static void DocxParagraphPropertyMathMatchesIndependentWordOrigins()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for independent paragraph-property Word origins.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "paragraph-property-math-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var pages = new DocxRenderer(null, mode, CachedMathGeometry(mode)).RenderPagesAsync(DocxTests.ReadDocx(input, mode), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(1, pages.Length);
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0], true).Where(show => pages[0].Fonts.Any(font =>
                font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))).ToArray();
            TestAssert.Equal("ratetimeratetime", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(16, math.Length);
            var anchors = row.GetProperty("MathOrigins").EnumerateArray().ToArray();
            for (int index = 0; index < anchors.Length; index++)
            {
                var actual = math[4 * index]; var expected = anchors[index];
                TestAssert.Equal(expected.GetProperty("Text").GetString(), actual.Text.Normalize(NormalizationForm.FormKC));
                TestAssert.True(Math.Abs(actual.X - expected.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(actual.Y - expected.GetProperty("Y").GetDouble()) < 1d,
                    "Resolved paragraph history retains the existing one-point Word math-origin gate: " + input + "/" + mode);
            }
        }
    }

    private static DocxDocument ReadParagraphPropertyMath(string container, OoxPdfDocxMarkupMode mode,
        bool baseline = false, bool noOp = false, bool duplicate = false, bool navigation = false, bool? joinMathFirst = null) =>
        ReadMutation("CompositionProbes/math-fields-simple-" + container, mode, navigation, navigation, (xml, w, m) =>
        {
            var equation = xml.Descendants(m + "oMath").First();
            var paragraph = equation.Parent!;
            if (!baseline)
            {
                var properties = paragraph.Element(w + "pPr");
                if (properties is null) { properties = new XElement(w + "pPr"); paragraph.AddFirst(properties); }
                var previous = new XElement(properties);
                if (!noOp)
                {
                    var spacing = previous.Element(w + "spacing");
                    if (spacing is null) { spacing = new XElement(w + "spacing"); previous.Add(spacing); }
                    spacing.SetAttributeValue(w + "after", "240");
                }
                var change = new XElement(w + "pPrChange", new XAttribute(w + "id", "975"), new XAttribute(w + "author", "Public"),
                    new XAttribute(w + "date", "2026-01-01T00:00:00Z"), previous);
                properties.Add(change);
                if (duplicate) properties.Add(new XElement(change));
            }
            if (joinMathFirst is bool first)
            {
                var head = first ? paragraph : paragraph.ElementsBeforeSelf(w + "p").Last();
                var properties = head.Element(w + "pPr");
                if (properties is null) { properties = new XElement(w + "pPr"); head.AddFirst(properties); }
                var mark = properties.Element(w + "rPr");
                if (mark is null) { mark = new XElement(w + "rPr"); properties.Add(mark); }
                mark.Add(new XElement(w + "del", new XAttribute(w + "id", "976"), new XAttribute(w + "author", "Public")));
            }
            if (navigation) equation.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), new XElement(equation)));
        }, partOverride: container is "header" or "footer" ? "word/math-fields-" + container + ".xml" : null);
}
