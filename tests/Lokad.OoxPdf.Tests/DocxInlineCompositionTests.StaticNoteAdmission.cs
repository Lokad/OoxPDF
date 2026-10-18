using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionStaticNoteAdmissionMatchesIndependentWordPageCounts()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent static/note admission controls use installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("static-note-admission-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string path = StaticNoteAdmissionCasePath(row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = DocxTests.ReadDocx(path, mode);
            var renderer = new DocxRenderer(null, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = renderer.RenderBlankPages(source, diagnostics.Add, default).ToArray();
            TestAssert.Equal(row.GetProperty("Pages").GetInt32(), pages.Length);
            int mathematicalGlyphs = pages.Sum(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Count(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))));
            TestAssert.Equal(row.GetProperty("MathGlyphs").GetInt32(), mathematicalGlyphs);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION" || d.PartName != source.MainPartName),
                "Selected body formulas survive admission; unused header-part diagnostics remain explicit on header-removal controls.");
        }
    }

    public static void DocxInlineCompositionStaticNoteAdmissionSeparatesOwnedReferencesFromTargetLabels()
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("static-note-admission-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var source = DocxTests.ReadDocx(StaticNoteAdmissionCasePath(row.GetProperty("Input").GetString()!), OoxPdfDocxMarkupMode.Final);
            var note = source.RelatedStories.Single(story => story.Kind == DocxRelatedStoryKind.Footnote && story.IsNormalStoryType);
            var paragraph = note.Paragraphs.Single();
            TestAssert.True(paragraph.Runs.All(run => !run.IsAutomaticNoteStoryMarker) && paragraph.Runs[0].Text.StartsWith("All quantities", StringComparison.Ordinal),
                "The plain target has no authored number; admission must not manufacture one.");
            var equations = DocxBlockTraversal.EnumerateBodyParagraphs(source).Concat(
                    DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(source, default).Select(owner => owner.Paragraph))
                .SelectMany(owner => owner.Equations).Distinct().ToDictionary(equation => equation,
                    equation => new DocxPreparedMathParagraph(equation, new(18d, 9d, 4d, []), 12d, 16d));
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(source, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: equations);
            TestAssert.Equal(row.GetProperty("HasBodyNote").GetBoolean(), layout.UsesSourceOwnedFootnoteReserve);
            TestAssert.True(paragraph.Runs.All(run => !run.IsAutomaticNoteStoryMarker), "Physical admission leaves target numbering roles unchanged.");
        }
    }

    private static string StaticNoteAdmissionCasePath(string relative) =>
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", relative));
}
