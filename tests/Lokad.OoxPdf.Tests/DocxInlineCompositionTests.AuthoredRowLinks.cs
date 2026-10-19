using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxAuthoredRowLinksRetainWordPlacementAcrossContainers()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent authored-row container origins use installed Windows fonts.");
        CheckAuthoredRowLinks(null, compareWord: true, "math-authored-row-link-containers-office.json", compareUnlinkedControls: false);
    }

    public static void DocxAuthoredRowLinksUseActualEquationOwnersAndObserveCancellation()
    {
        foreach (string variant in new[] { "unlinked", "linked-prose", "linked-math" })
        {
            var document = ReadMutation("CompositionProbes/math-authored-row-" + variant,
                OoxPdfDocxMarkupMode.Final, false, false);
            var paragraph = document.Paragraphs.Single(paragraph => paragraph.Equations.Count != 0);
            var equation = paragraph.Equations.Single();
            TestAssert.True(!equation.Display, "The source equation remains implicitly inline; authored-row placement owns promotion.");
            TestAssert.True(DocxInlineContent.HasValidHyperlinkSourceRanges(paragraph, default), "Every source link has owned inline ranges.");
            TestAssert.Equal(variant == "linked-math", DocxInlineContent.GetHyperlinkedEquations(paragraph, default)?.Contains(equation) == true);
            TestAssert.Equal(variant != "linked-math", DocxMathParagraphPolicy.GetDisplaySourceEquations(paragraph, default).Contains(equation));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathParagraphPolicy.GetDisplaySourceEquations(paragraph, cancellation.Token));
        }
    }

    public static void DocxAuthoredRowLinksPreserveIndependentWordGeometry()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent authored-row origins use installed Windows fonts.");
        CheckAuthoredRowLinks(null, compareWord: true);
    }

    internal static void CheckAuthoredRowLinks(IFontResolver? resolver, bool compareWord,
        string goldFile = "math-authored-row-links-office.json", bool compareUnlinkedControls = true)
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, goldFile)));
        var unlinked = new Dictionary<string, (double X, double Y, double Size)[]>();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string input = Path.Combine(cases, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            string preset = row.GetProperty("Mode").GetString()!;
            bool final = preset == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = DocxTests.ReadDocx(input, mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(resolver, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, diagnostics.Add, default).Single();
            var mathFonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(font => font.Font).ToHashSet();
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Where(show => mathFonts.Contains(show.Font) && show.Text.Trim().Length != 0).ToArray();
            TestAssert.Equal("x+2y", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(4, math.Length);
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Hyperlink ownership cannot remove the supported equation.");
            var positions = math.Select(show => (show.X, show.Y, show.Size)).ToArray();
            string variant = row.GetProperty("Variant").GetString()!;
            if (variant == "unlinked") unlinked.Add(preset, positions);
            if (variant == "linked-prose" && compareUnlinkedControls)
                TestAssert.True(positions.SequenceEqual(unlinked[preset]), "A link confined to the preceding text row preserves every mathematical origin and size.");
            if (variant != "unlinked")
            {
                var annotation = page.Annotations.Single(item => item.IsDestination);
                if (variant == "linked-prose")
                    TestAssert.True(math.All(show => show.Y < annotation.Y), "The preceding text link cannot cover a separate mathematical row.");
                else
                    TestAssert.True(math.All(show => annotation.X <= show.X + .01d && annotation.X + annotation.Width >= show.X - .01d &&
                        annotation.Y <= show.Y + .01d && annotation.Y + annotation.Height >= show.Y - .01d),
                        "The linked equation retains annotation bounds around every painted glyph.");
            }
            if (compareWord)
            {
                var expected = row.GetProperty("WordMathAnchors").EnumerateArray().ToArray();
                for (int index = 0; index < math.Length; index++)
                    TestAssert.True(Math.Abs(math[index].X - expected[index].GetProperty("X").GetDouble()) < 1d &&
                        Math.Abs(page.Height - math[index].Y - expected[index].GetProperty("Y").GetDouble()) < 1d &&
                        Math.Abs(math[index].Size - expected[index].GetProperty("Size").GetDouble()) < .1d,
                        id + ": original independent Word glyph origins and sizes remain within their thresholds.");
            }
        }
    }
}
