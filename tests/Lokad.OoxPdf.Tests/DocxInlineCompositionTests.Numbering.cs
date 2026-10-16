using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNumberedDisplayRowsMatchIndependentWord() =>
        AssertNumberedWordFrames("math-numbered-display-office-geometry.json");

    private static void AssertNumberedWordFrames(string referenceFile)
    {
        if (!OperatingSystem.IsWindows() || !File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc"))) TestAssert.Skip("Independent controls use installed Windows math/text metrics.");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder, referenceFile)));
        foreach (var row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            if (row.TryGetProperty("Supported", out var supported) && !supported.GetBoolean()) continue;
            string id = row.GetProperty("Id").GetString()!;
            string input = Path.Combine(folder, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var page = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(DocxTests.ReadDocx(input, mode), null, default).Single();
            var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
            var anchors = row.GetProperty("Anchors").EnumerateArray().ToArray();
            TestAssert.Equal(anchors.Length, math.Length);
            TestAssert.Equal(row.GetProperty("MathText").GetString(), string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            for (int i = 0; i < anchors.Length; i++)
            {
                TestAssert.True(Math.Abs(math[i].X - anchors[i].GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(math[i].Y - anchors[i].GetProperty("Y").GetDouble()) < 1d,
                    "Numbered formulas retain independent Word placement across row/container roles: " + id);
                TestAssert.True(Math.Abs(math[i].Size / 11d - anchors[i].GetProperty("FontSize").GetDouble() / row.GetProperty("ProseFontSize").GetDouble()) < .01d,
                    "Only the first authored row belongs to the list label; later isolated rows retain display sizing.");
            }
            foreach (string label in new[] { "1.", "2.", "3." })
                TestAssert.Equal(1, DocxPageFootnoteNumberingTests.DecodeText(page).Split(label, StringSplitOptions.None).Length - 1);
        }
    }

    public static void DocxInlineCompositionNumberedDisplayRowsPreserveSourceAndMetadata()
    {
        if (!OperatingSystem.IsWindows() || !File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc"))) TestAssert.Skip("Native controls use installed Windows math fonts.");
        CheckNumberedDisplayRows(null);
    }

    internal static void CheckNumberedDisplayRows(IFontResolver? resolver)
    {
        foreach (string story in new[] { "body", "cell", "header", "footer", "math-only", "math-break" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = "CompositionProbes/math-numbered-display-" + story;
            string? part = story == "header" ? "word/header1.xml" : story == "footer" ? "word/footer1.xml" : null;
            (double X, double Y, double Size)[]? before = null;
            foreach (bool changed in new[] { false, true })
            {
                var source = ReadMutation(fixture, mode, changed, changed, resolver is null ? null : (xml, w, m) =>
                {
                    foreach (var run in xml.Descendants(m + "r"))
                    {
                        var properties = run.Element(m + "rPr");
                        if (properties is null) { properties = new XElement(m + "rPr"); run.AddFirst(properties); }
                        properties.Elements(m + "nor").Remove(); properties.Add(new XElement(m + "nor", new XAttribute(m + "val", "1")));
                    }
                }, partOverride: part);
                var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
                var page = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                    .RenderBlankPages(source, diagnostics.Add, default).Single();
                var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
                var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
                TestAssert.Equal(story == "math-break" ? 22 : 18, math.Length);
                string text = string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(story == "math-break" ? 3 : 2, text.Split("x+2y", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(2, text.Split("a+b=c", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(1, page.Images.Count);
                string target = part is null ? source.MainPartName : "/" + part;
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION" || d.PartName != target),
                    "Explicit math zones cannot disappear beside a supported list label/prose/picture.");
                foreach (string label in new[] { "1.", "2.", "3." })
                    TestAssert.Equal(1, DocxPageFootnoteNumberingTests.DecodeText(page).Split(label, StringSplitOptions.None).Length - 1);
                if (before is not null)
                    for (int i = 0; i < before.Length; i++)
                        TestAssert.True(Math.Abs(before[i].X - math[i].X) < .01d && Math.Abs(before[i].Y - math[i].Y) < .01d && before[i].Size == math[i].Size,
                            "Bookmarks and unchanged-format run splitting preserve numbered mathematical frames.");
                before = math.Select(show => (show.X, show.Y, show.Size)).ToArray();
            }
        }
    }

    public static void DocxInlineCompositionListLabelOwnsOnlyTheFirstAuthoredMathRow()
    {
        foreach (string story in new[] { "math-only", "math-break" })
        {
            var document = ReadMutation("CompositionProbes/math-numbered-display-" + story, OoxPdfDocxMarkupMode.Final, false, false);
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements).First(p => p.ListLabel is not null && p.Equations.Count != 0);
            var mixed = DocxMathParagraphPolicy.GetMixedSourceEquations(paragraph, default);
            var display = DocxMathParagraphPolicy.GetDisplaySourceEquations(paragraph, default);
            TestAssert.True(mixed.Contains(paragraph.Equations[0]) && !display.Contains(paragraph.Equations[0]),
                "A visible list label owns the first authored row even when no prose accompanies its equation.");
            TestAssert.True(paragraph.Equations.All(e => e.Display), "Resolved measurement keeps authored source flags intact.");
            TestAssert.Equal(paragraph.Equations.Count, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            if (story == "math-break")
                TestAssert.True(display.Contains(paragraph.Equations[1]) && !mixed.Contains(paragraph.Equations[1]),
                    "An authored break ends label ownership before the independent math-only display row.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathParagraphPolicy.GetMixedSourceEquations(paragraph, canceled.Token));
        }
    }
}
