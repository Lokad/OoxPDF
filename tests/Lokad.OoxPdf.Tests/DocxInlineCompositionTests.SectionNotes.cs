using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionSectionNoteMathRetainsSourceOwners()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-section-notes-" + kind + "-" + container, mode, true, true);
            var paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements)
                .Where(paragraph => paragraph.Equations.Count != 0).ToArray();
            TestAssert.Equal(4, paragraphs.Length);
            TestAssert.Equal("3,4,vii,viii", string.Join(',', paragraphs.SelectMany(p => p.InlineReferences).Select(r => r.DisplayText)));
            foreach (var paragraph in paragraphs)
            {
                var reference = paragraph.InlineReferences.Single();
                TestAssert.Equal("eachSect", reference.NoteNumberRestartValue);
                TestAssert.True(DocxInlineContent.TryGetAutomaticNoteRun(reference, out var run) &&
                    paragraph.Runs.Any(owner => ReferenceEquals(owner, run)), "Section-resolved labels retain their actual marker owner.");
                TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
                foreach (string restart in new[] { "eachPage", "invalid" })
                    TestAssert.Equal(restart == "eachPage" && reference.Kind == DocxRelatedStoryKind.Footnote ? 2 : 0,
                        DocxMathParagraphPolicy.GetPreparationEquations(paragraph with
                            { InlineReferences = [reference with { NoteNumberRestartValue = restart }] }).Count);
            }
        }
    }

    public static void DocxInlineCompositionSectionNoteMathPreservesContentAndDestinations()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native composition controls use installed Windows fonts.");
        CheckSectionNoteMath(null);
    }

    public static void DocxInlineCompositionSectionNoteMathKeepsIndependentMarkerGeometry()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent marker controls use installed Windows fonts.");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "math-section-notes-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string path = Path.Combine(folder, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = DocxTests.ReadDocx(path, mode);
            var lines = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(document).Lines
                .Where(line => line.ContainerStoryKind == "Body" && line.Segments.Any(segment => segment.FontSize < 9d)).ToArray();
            TestAssert.Equal(4, lines.Length);
            foreach (var line in lines)
            {
                var marker = line.Segments.Single(segment => segment.FontSize < 9d);
                var prose = line.Segments.First(segment => segment.FontSize > 11d);
                TestAssert.True(Math.Abs(marker.BaselineY - prose.BaselineY - row.GetProperty("MarkerRaise").GetDouble()) < .1d,
                    "Section-resolved markers preserve the independent Word superscript raise.");
                TestAssert.True(Math.Abs(marker.PdfFontSize / prose.PdfFontSize - row.GetProperty("MarkerFontRatio").GetDouble()) < .002d,
                    "Section-resolved markers preserve the independent Word superscript font ratio.");
            }
        }
    }

    public static void DocxInlineCompositionSectionNoteMathKeepsIndependentEndnoteParagraphPitch()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent paragraph pitch uses installed Windows fonts.");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "math-section-notes-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            if (!row.TryGetProperty("NoteParagraphPitches", out var expected)) continue;
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = DocxTests.ReadDocx(Path.Combine(folder, row.GetProperty("Id").GetString() + ".docx"), mode);
            var lines = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(document).Lines
                .Where(line => line.ContainerStoryKind == "Endnote" && line.TextLength > 1).ToArray();
            TestAssert.Equal(expected.GetArrayLength() + 1, lines.Length);
            var baselines = row.GetProperty("NoteProseBaselines");
            for (int index = 0; index < lines.Length; index++)
                TestAssert.True(Math.Abs(lines[index].Segments.MaxBy(segment => segment.TextLength)!.BaselineY -
                    baselines[index].GetDouble()) < 1d, "Complete body inline depth preserves independent endnote placement.");
            for (int index = 0; index < expected.GetArrayLength(); index++)
            {
                double baseline = lines[index].Segments.MaxBy(segment => segment.TextLength)!.BaselineY;
                double next = lines[index + 1].Segments.MaxBy(segment => segment.TextLength)!.BaselineY;
                TestAssert.True(Math.Abs(baseline - next - expected[index].GetDouble()) < .03d,
                    "Adjacent complete endnote paragraphs share Word's measured pitch even beside body equations.");
            }
        }
    }

    public static void DocxInlineCompositionSectionNoteMathCountsExtraDepthOnceAndKeepsFallbacks()
    {
        var document = ReadMutation("CompositionProbes/math-section-notes-endnote-body", OoxPdfDocxMarkupMode.Final, false, false);
        TestAssert.True(DocxPlainBodyFrame.HasKnownBodyFrame(document), "Resolved mixed paragraphs can own a known body frame.");
        TestAssert.True(!DocxPlainBodyFrame.HasKnownBodyFrame(document with { FinalSectionBreak = null }) &&
            !DocxPlainBodyFrame.HasKnownBodyFrame(document with { Settings = document.Settings with { MirrorMargins = true } }),
            "Unknown and mirrored frames retain the existing fallback independently of paragraph contents.");
        var source = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default).Pages[0].Items.OfType<DocxTextLineLayout>().First();
        var line = source with { InlineFrame = new(15d, 21d) { TextBaselineInset = 11d }, SingleLineHeight = 14d };
        TestAssert.Equal(3d, DocxInlineLineGeometry.ExtraBelowTextFrame(line));
        TestAssert.Equal(3d, DocxInlineLineGeometry.ExtraBelowTextFrame(line with { BaselineY = line.BaselineY - 50d }));
        foreach (var fallback in new[] {
            line with { InlineFrame = null },
            line with { InlineFrame = new(15d, 21d) },
            line with { InlineFrame = new(15d, double.NaN) { TextBaselineInset = 11d } },
            line with { InlineFrame = new(15d, 14d) { TextBaselineInset = 11d } },
            line with { SingleLineHeight = null },
            line with { SingleLineHeight = 10d },
            line with { InlineFrame = new(11d, 14d) { TextBaselineInset = 11d } },
            line with { SourceParagraph = null },
            line with { SourceParagraph = line.SourceParagraph! with { LineSpacingPoints = 12d,
                Spacing = line.SourceParagraph.Spacing with { LineRuleValue = "exact" } } },
        }) TestAssert.Equal(0d, DocxInlineLineGeometry.ExtraBelowTextFrame(fallback));
    }

    internal static void CheckSectionNoteMath(IFontResolver? resolver)
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-section-notes-" + kind + "-" + container, mode, true, true);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(8, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            foreach (int index in new[] { 1, 2, 3, 4 }) TestAssert.Contains("Public destination " + index + ".", text);
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Restarted source counters cannot remove supported equations.");
            var links = pages.SelectMany(page => page.Annotations).Where(link => link.Destination is not null).ToArray();
            TestAssert.Equal(4, links.Length);
            TestAssert.Equal(4, links.Select(link => link.Destination).Distinct().Count());
            TestAssert.True(links.All(link => link.Width > 0 && link.Height > 0 &&
                link.Destination!.Value.PageIndex >= 0 && link.Destination.Value.PageIndex < pages.Length),
                "Every source marker retains a distinct, admitted note destination after section placement.");
        }
    }
}
