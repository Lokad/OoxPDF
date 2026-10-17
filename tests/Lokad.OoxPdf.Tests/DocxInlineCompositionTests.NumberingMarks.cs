using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNumberingMarksRequireActualSourceOwners()
    {
        var document = DocxTests.ReadDocx(NumberingMarkCase("list-decoration-mark-inserted.docx"), OoxPdfDocxMarkupMode.AllMarkup);
        var paragraph = document.FooterParagraphs.First(p => p.ListLabel?.Text == "1.");
        var mark = paragraph.ParagraphMarkRun;
        TestAssert.True(mark is not null, "A numbered paragraph retains its actual inserted mark independently of text runs.");
        var insertion = paragraph.Revisions.Single(r => r.Kind == DocxRevisionKind.Insertion);
        TestAssert.True(mark!.Revisions.Any(r => ReferenceEquals(r, insertion)), "The mark owns the original source revision object.");
        TestAssert.True(!DocxInlineContent.HasSupportedRunRevisionOwners(paragraph.Revisions, paragraph.Runs, default),
            "Ordinary text runs cannot stand in for an inserted paragraph mark.");
        TestAssert.True(DocxInlineContent.HasSupportedSourceOwner(paragraph), "Numbering gives the actual mark an explicit source owner.");
        TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(paragraph with { ParagraphMarkRun = null }),
            "A missing paragraph mark cannot own its insertion.");
        TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(paragraph with { ListLabel = null }),
            "Unnumbered mark revisions retain their existing fallback.");
        foreach (var invalid in new[]
        {
            mark with { Revisions = [insertion with { }] },
            mark with { SourceRunIndex = 0 },
            mark with { Text = "ordinary text" }
        })
            TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(paragraph with { ParagraphMarkRun = invalid }),
                "Value-equal or misplaced metadata cannot manufacture a numbering owner.");
        var deletion = insertion with { Kind = DocxRevisionKind.Deletion, SourceElement = "del" };
        TestAssert.True(!DocxInlineContent.HasSupportedSourceOwner(paragraph with
            { Revisions = [deletion], ParagraphMarkRun = mark with { Revisions = [deletion] } }),
            "Paragraph-mark deletion/merge behavior needs its separate qualification.");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.HasSupportedSourceOwner(paragraph, token: cancellation.Token));
    }

    public static void DocxInlineCompositionNumberingMarkOwnersKeepVisibleContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native comparison uses installed Word-compatible math fonts.");
        CheckNumberingMarkContent(null);
    }

    internal static void CheckNumberingMarkContent(IFontResolver? resolver)
    {
        string folder = Path.GetDirectoryName(NumberingMarkCase("numbering-marks.json"))!;
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "numbering-marks.json")));
        using var golden = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "numbering-marks-office.json")));
        foreach (var row in manifest.RootElement.EnumerateArray())
        foreach (bool final in new[] { true, false })
        {
            string id = row.GetProperty("Id").GetString()!;
            string file = Path.GetFileName(row.GetProperty("Input").GetString()!);
            string path = Path.Combine(folder, file);
            TestAssert.Equal(row.GetProperty("Sha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            string modeName = final ? "final-preserve" : "all-word-compatible";
            var expected = golden.RootElement.GetProperty("Rows").EnumerateArray()
                .Single(value => value.GetProperty("Id").GetString() == id && value.GetProperty("Mode").GetString() == modeName);
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = DocxTests.ReadDocx(path, mode);
            var paragraphs = DocxBlockTraversal.EnumerateSourceParagraphs(document).Where(p => p.ListLabel is not null)
                .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance).ToArray();
            var first = paragraphs.Single(p => p.ListLabel!.Text == "1.");
            var owners = paragraphs.SelectMany(p => p.Runs).ToArray();
            TestAssert.True(first.ParagraphMarkRun is not null && DocxInlineContent.HasSupportedSourceOwner(first),
                "Parsed numbered marks own their shared paragraph source: " + id);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true, includeTextTransforms: true).ToArray();
            foreach (string label in new[] { "1.", "2.", "3." })
                TestAssert.True(shows.Any(show => show.Text.Trim() == label && show.X >= 0d && show.X < page.Width && show.Y > 0d && show.Y < page.Height),
                    "A paragraph-mark insertion cannot discard its numbering content: " + id + "/" + modeName + "/" + label);
            if (resolver is null)
            {
                var firstPaint = shows.First(show => show.Text.Trim() == "1.");
                TestAssert.True(Math.Abs(firstPaint.X - expected.GetProperty("FirstMarkerX").GetDouble()) < 1d &&
                    Math.Abs(firstPaint.Y - expected.GetProperty("FirstMarkerY").GetDouble()) < 1d,
                    "Restored markers use the independent Word frame within the existing one-point tolerance: " + id);
            }
            var mathFonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
            var math = shows.Where(show => mathFonts.Contains(show.Font)).ToArray();
            bool hasMath = row.GetProperty("Kind").GetString() == "math";
            TestAssert.Equal(hasMath ? 18 : 0, math.Length);
            if (hasMath)
            {
                string text = string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(2, text.Split("x+2y", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(2, text.Split("a+b=c", StringSplitOptions.None).Length - 1);
                string story = row.GetProperty("Story").GetString()!;
                string part = story is "body" or "cell" ? document.MainPartName : "/word/" + story + "1.xml";
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION" || d.PartName != part),
                    "An actual numbering mark cannot remove supported selected-story equations.");
                TestAssert.True(page.Images.Any(image => page.Content.Contains("/" + image.ResourceName + " Do", StringComparison.Ordinal)),
                    "Numbering-mark ownership preserves the adjacent picture.");
            }
            TestAssert.Contains("share the same input series.", DocxPageFootnoteNumberingTests.DecodeText(page));
            TestAssert.True(owners.Zip(paragraphs.SelectMany(p => p.Runs)).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Admitting the mark owner preserves the original text objects.");
        }
    }

    private static string NumberingMarkCase(string file) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "CompositionProbes", file));
}
