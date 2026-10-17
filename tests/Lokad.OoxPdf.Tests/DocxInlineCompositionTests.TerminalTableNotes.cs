using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionTerminalTableNotesKeepIndependentParagraphFlow()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent terminal-table controls use installed Windows fonts.");
        CheckTerminalTableNotes(null, true);
    }

    internal static void CheckTerminalTableNotes(IFontResolver? resolver, bool compareGeometry)
    {
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, "math-terminal-table-notes-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string input = Path.Combine(folder, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = ReadMutation("CompositionProbes/" + id, mode, true, true);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(resolver, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(8, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            for (int index = 1; index <= 4; index++) TestAssert.Contains("Public destination " + index + ".", text);
            if (row.GetProperty("Shape").GetString() == "prose-zero") TestAssert.Contains("Public ending paragraph.", text);
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Terminal paragraph ownership retains supported math.");
            var links = pages.SelectMany(page => page.Annotations).Where(link => link.Destination is not null).ToArray();
            TestAssert.Equal(4, links.Length);
            TestAssert.Equal(4, links.Select(link => link.Destination).Distinct().Count());
            TestAssert.True(links.All(link => link.Width > 0 && link.Height > 0 && link.Destination!.Value.PageIndex < pages.Length),
                "Every marker owns a distinct admitted note destination.");
            if (!compareGeometry) continue;
            TestAssert.Equal(row.GetProperty("ExpectedPages").GetInt32(), pages.Length);
            var lines = renderer.InspectTextEmission(document).Lines.Where(line => line.ContainerStoryKind == "Endnote" && line.TextLength > 1).ToArray();
            var baselines = row.GetProperty("NoteProseBaselines");
            var pitches = row.GetProperty("NoteParagraphPitches");
            TestAssert.Equal(baselines.GetArrayLength(), lines.Length);
            for (int index = 0; index < lines.Length; index++)
            {
                double baseline = lines[index].Segments.MaxBy(segment => segment.TextLength)!.BaselineY;
                TestAssert.True(Math.Abs(baseline - baselines[index].GetDouble()) < 1d,
                    "Implicit and explicit terminal paragraphs preserve independent Word note positions.");
                if (index + 1 == lines.Length) continue;
                double next = lines[index + 1].Segments.MaxBy(segment => segment.TextLength)!.BaselineY;
                TestAssert.True(Math.Abs(baseline - next - pitches[index].GetDouble()) < .03d,
                    "An earlier table does not add a separator gap between complete note paragraphs.");
            }
        }
    }
}
