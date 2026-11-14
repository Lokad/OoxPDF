using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxParagraphMarkFormattingRetainsResolvedSourceAcrossViews()
    {
        foreach (string container in new[] { "body", "cell", "header", "footer", "plain" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var paragraph = MarkPropertyParagraph(ReadMarkPropertyMath(container, mode));
            var mark = TestAssert.NotNull(paragraph.ParagraphMarkRun);
            var revision = mark.Revisions.Single();
            TestAssert.Equal(-1, mark.SourceRunIndex);
            TestAssert.True(mark.Bold && mark.Text.Length == 0, "The resolved mark remains separate from inline text.");
            TestAssert.Equal(DocxRevisionKind.RunPropertiesChange, revision.Kind);
            TestAssert.Equal("rPrChange", revision.SourceElement);
            TestAssert.Equal(DocxRevisionPropertyFamily.Run, revision.PropertyChangeFamily!.Value);
            TestAssert.True(paragraph.Revisions.Any(item => ReferenceEquals(item, revision)), "Inspection retains the actual mark owner.");
            var baseline = MarkPropertyParagraph(ReadMarkPropertyMath(container, mode, baseline: true));
            TestAssert.Equal(baseline.Spacing, paragraph.Spacing);
            TestAssert.Equal(baseline.Indent, paragraph.Indent);
            TestAssert.Equal(baseline.ParagraphMarkFontSize, paragraph.ParagraphMarkFontSize);
            TestAssert.Equal(baseline.ListLabel?.Style, paragraph.ListLabel?.Style);
            TestAssert.True(baseline.Runs.Select(run => (run.Text, run.EffectiveProperties))
                .SequenceEqual(paragraph.Runs.Select(run => (run.Text, run.EffectiveProperties))), "History cannot rewrite the current source runs.");
            TestAssert.Equal(1, paragraph.Equations.Count);
            var noOp = MarkPropertyParagraph(ReadMarkPropertyMath(container, mode, noOp: true));
            TestAssert.True(noOp.Revisions.All(item => item.Id != "992"), "The existing reader removes semantically void history.");
        }
    }

    public static void DocxParagraphMarkFormattingRequiresActualUniqueOwnersAndCancellation()
    {
        foreach (string container in new[] { "body", "plain" })
        {
            var paragraph = MarkPropertyParagraph(ReadMarkPropertyMath(container, OoxPdfDocxMarkupMode.Final));
            var mark = paragraph.ParagraphMarkRun!;
            var revision = mark.Revisions.Single();
            TestAssert.True(DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default), "A resolved mark owns its formatting history.");
            foreach (var invalid in new[]
            {
                paragraph with { ParagraphMarkRun = null },
                paragraph with { ParagraphMarkRun = mark with { SourceRunIndex = 0 } },
                paragraph with { ParagraphMarkRun = mark with { Text = "unrelated" } },
                paragraph with { ParagraphMarkRun = mark with { Revisions = [revision with { }] } },
                paragraph with { Revisions = [revision with { }] },
                paragraph with { Revisions = [revision, revision with { Id = "993" }] }
            })
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(invalid, default), "Missing, misplaced or copied metadata cannot invent ownership.");
            foreach (var invalid in new[] { revision with { SourceElement = "pPrChange" },
                revision with { PropertyChangeFamily = DocxRevisionPropertyFamily.Paragraph }, revision with { Kind = DocxRevisionKind.MoveTo },
                revision with { Kind = DocxRevisionKind.Deletion, SourceElement = "del" } })
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with
                    { Revisions = [invalid], ParagraphMarkRun = mark with { Revisions = [invalid] } }, default), "Other mark roles keep their fallback.");
            var duplicate = MarkPropertyParagraph(ReadMarkPropertyMath(container, OoxPdfDocxMarkupMode.Final, duplicate: true));
            TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(duplicate).Count);
            if (container == "plain")
            {
                var insertion = revision with { Kind = DocxRevisionKind.Insertion, SourceElement = "ins" };
                TestAssert.True(DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with
                    { Revisions = [insertion], ParagraphMarkRun = mark with { Revisions = [insertion] } }, default),
                    "The actual empty mark owns its insertion independently of visible numbering.");
                var misplaced = insertion with { SourceElement = "pPrChange" };
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with
                    { Revisions = [misplaced], ParagraphMarkRun = mark with { Revisions = [misplaced] } }, default),
                    "An insertion with the wrong source element cannot manufacture mark ownership.");
            }
            TestAssert.True(!DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph), "This change does not broaden note or typography correction guards.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, canceled.Token));
        }
    }

    public static void DocxParagraphMarkFormattingPreservesComposedPaintAndNavigation()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string container in new[] { "body", "cell", "header", "footer", "plain" })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var document = ReadMarkPropertyMath(container, mode);
            var owner = MarkPropertyParagraph(document).ParagraphMarkRun;
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            var baseline = renderer.RenderPagesAsync(ReadMarkPropertyMath(container, mode, baseline: true), null, default).ToBlockingEnumerable().ToArray();
            string MathText(IEnumerable<Lokad.OoxPdf.Pdf.PdfPage> pages) => string.Concat(pages.SelectMany(page =>
                DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(show => page.Fonts.Any(font =>
                    font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))))
                .Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal("x+2ya+b=cx+2ya+b=c", MathText(actual));
            TestAssert.Equal(MathText(baseline), MathText(actual));
            TestAssert.Equal(baseline.Length, actual.Length);
            TestAssert.True(baseline.SelectMany(page => page.Annotations).Select(a => (a.Uri, a.Destination?.PageIndex))
                .SequenceEqual(actual.SelectMany(page => page.Annotations).Select(a => (a.Uri, a.Destination?.PageIndex))), "History retains source links and bookmark targets.");
            TestAssert.True(actual.SelectMany(page => page.Annotations).Any(), "The composite control must retain a real source link.");
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Formatting a paragraph mark cannot omit neighboring math.");
            TestAssert.True(ReferenceEquals(owner, MarkPropertyParagraph(document).ParagraphMarkRun), "Painting preserves the actual source mark.");
        }
    }

    public static void DocxParagraphMarkFormattingMatchesIndependentWordOrigins()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for independent paragraph-mark Word origins.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-mark-property-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var page = new DocxRenderer(null, mode, CachedMathGeometry(mode)).RenderPagesAsync(DocxTests.ReadDocx(input, mode), null, default)
                .ToBlockingEnumerable().Single();
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(show => page.Fonts.Any(font =>
                font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))).ToArray();
            var anchors = row.GetProperty("MathOrigins").EnumerateArray().ToArray();
            TestAssert.Equal(anchors.Length, math.Length);
            for (int i = 0; i < anchors.Length; i++)
            {
                TestAssert.Equal(anchors[i].GetProperty("Text").GetString(), math[i].Text.Normalize(NormalizationForm.FormKC));
                TestAssert.True(Math.Abs(math[i].X - anchors[i].GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(math[i].Y - anchors[i].GetProperty("Y").GetDouble()) < 1d,
                    "Mark formatting retains the unchanged one-point native origin gate: " + row.GetProperty("Container").GetString() + "/" + mode + "/" + i);
            }
        }
    }

    private static DocxParagraph MarkPropertyParagraph(DocxDocument document) => DocxBlockTraversal.EnumerateSourceParagraphs(document)
        .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance).Single(p => p.BookmarkAnchors.Any(anchor => anchor.Name == "Calculation"));

    private static DocxDocument ReadMarkPropertyMath(string container, OoxPdfDocxMarkupMode mode,
        bool baseline = false, bool noOp = false, bool duplicate = false) =>
        ReadMutation("CompositionProbes/math-mark-property-" + container, mode, false, false, (xml, w, m) =>
        {
            var change = xml.Descendants(w + "rPrChange").Single();
            if (baseline) change.Remove();
            else if (noOp) change.Descendants(w + "b").Single().SetAttributeValue(w + "val", null);
            else if (duplicate) change.AddAfterSelf(new XElement(change));
        }, partOverride: container == "header" ? "word/header1.xml" : container == "footer" ? "word/footer1.xml" : null);
}
