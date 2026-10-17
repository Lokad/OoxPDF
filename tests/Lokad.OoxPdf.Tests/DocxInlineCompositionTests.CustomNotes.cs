using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionCustomNotesRetainExactSourceOwnersAndFallbacks()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (int count in new[] { 1, 2 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation("CompositionProbes/math-custom-note-" + kind + "-" + count, mode, true, true);
            var paragraph = source.Paragraphs.Single(p => p.Equations.Count != 0);
            var reference = paragraph.InlineReferences.Single();
            var owners = DocxCustomNoteMarkerOwners.IndexSourceRuns(paragraph.Runs, default);
            TestAssert.True(DocxCustomNoteMarkerOwners.TryGetSourceRun(reference, owners, out var run), "The custom mark owns one exact source run and offset.");
            TestAssert.True(paragraph.InlineContent!.OfType<DocxInlineText>().Any(item => ReferenceEquals(item.Run, run)), "The actual marker run survives in ordered content.");
            TestAssert.Equal(count, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            foreach (var invalid in new[] {
                reference with { TextOffsetInRun = reference.TextOffsetInRun + 1 },
                reference with { SourceRunIndex = int.MaxValue },
                reference with { CustomMarkText = "" }, reference with { CustomMarkText = " " },
                reference with { CustomMarkText = "**" }, reference with { CustomMarkText = "\ud83d" },
                reference with { Kind = DocxRelatedStoryKind.Comment },
            }) TestAssert.True(!DocxCustomNoteMarkerOwners.TryGetSourceRun(invalid, owners, out _), "Unowned or invalid marker metadata retains fallback.");
            var duplicate = DocxCustomNoteMarkerOwners.IndexSourceRuns([.. paragraph.Runs, run!], default);
            TestAssert.True(!DocxCustomNoteMarkerOwners.TryGetSourceRun(reference, duplicate, out _), "Duplicate source coordinates are ambiguous even for equal glyphs.");
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxCustomNoteMarkerOwners.IndexSourceRuns(paragraph.Runs, cancellation.Token));
        }
    }

    public static void DocxInlineCompositionCustomNotesMatchIndependentContentAndKeepFirstGlyphLinks()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent custom-note controls use installed Windows fonts.");
        CheckCustomNoteMath(null);
    }

    internal static void CheckCustomNoteMath(IFontResolver? resolver)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "math-custom-notes-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(folder, id + ".docx")))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            double? markerWidth = null;
            foreach (bool suffix in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/" + id, mode, true, true, (xml, w, _) =>
                {
                    if (!suffix) return;
                    var mark = xml.Descendants().Single(node => node.Attribute(w + "customMarkFollows")?.Value == "1");
                    mark.ElementsAfterSelf().First(node => node.Name == w + "t").Value = "*tail";
                });
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = new DocxRenderer(resolver, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                    : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, diagnostics.Add, default).ToArray();
                TestAssert.Equal(row.GetProperty("ExpectedPages").GetInt32(), pages.Length);
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(row.GetProperty("ExpectedEquations").GetInt32(), text.Split("ratetime", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(row.GetProperty("ExpectedCustomMarks").GetInt32(), text.Count(character => character == '*'));
                foreach (string prose in new[] { "Public destination Alpha.", "Public destination Delta." }) TestAssert.Contains(prose, text);
                if (suffix) TestAssert.Contains("*tail", text);
                var links = pages.SelectMany(page => page.Annotations).Where(link => link.Destination is not null).ToArray();
                TestAssert.Equal(4, links.Length);
                TestAssert.Equal(4, links.Select(link => link.Destination).Distinct().Count());
                if (!suffix) markerWidth = links[0].Width;
                else TestAssert.True(Math.Abs(markerWidth!.Value - links[0].Width) < .000001d, "Math and source-run splitting cannot extend a custom link beyond its first character.");
                TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "An owned custom mark cannot discard supported equations.");
            }
        }
    }
}
