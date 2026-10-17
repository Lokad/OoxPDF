using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNestedMathKeepsPreparedOwnersThroughEveryMeasuredFrame()
    {
        using var manifest = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("mixed-nested-math.json")));
        foreach (var row in manifest.RootElement.EnumerateArray())
        {
            var document = ReadMutation("CompositionProbes/" + row.GetProperty("Id").GetString(),
                OoxPdfDocxMarkupMode.Final, false, false);
            var owners = DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, default)
                .Concat(DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default).Select(owner => owner.Paragraph))
                .Where(paragraph => paragraph.Equations.Count != 0).ToArray();
            TestAssert.Equal(row.GetProperty("ExpectedEquationCount").GetInt32(), owners.Sum(paragraph => paragraph.Equations.Count));
            var math = owners.SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 20d, 8d, []), 28d, 36d));
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
            var lines = layout.Pages.SelectMany(page => NestedMathRows(page.Items.OfType<DocxTableRowLayout>()
                .Concat(page.StaticTableRows).Concat(page.PlacedRelatedStories.SelectMany(story => story.TableRows))))
                .SelectMany(table => table.Cells).SelectMany(cell => cell.TextLines)
                .Where(line => line.InlineMath.Count != 0).ToArray();
            TestAssert.Equal(math.Count, lines.Sum(line => line.InlineMath.Count));
            TestAssert.True(lines.All(line => owners.Any(owner => ReferenceEquals(owner, line.SourceParagraph))) &&
                lines.SelectMany(line => line.InlineMath).All(placement => ReferenceEquals(math[placement.Math.Equation], placement.Math)),
                "Every nested measured/placed box keeps the prepared equation and actual paragraph owner.");
            foreach (var page in layout.Pages)
            foreach (var line in NestedMathRows(page.StaticTableRows).SelectMany(table => table.Cells).SelectMany(cell => cell.TextLines))
                TestAssert.True(DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double bottom, out double top) &&
                    bottom >= 0d && top <= page.Height,
                    "A table-only footer anchors the complete nested frame inside the physical page.");
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() =>
                DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, cancelled.Token).ToArray());
        }
    }

    public static void DocxInlineCompositionNestedMathKeepsContainingRowsAtomicAtPageBoundaries()
    {
        var document = ReadMutation("CompositionProbes/mixed-nested-cell-inside", OoxPdfDocxMarkupMode.Final, false, false);
        var math = DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, default)
            .SelectMany(paragraph => paragraph.Equations).ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 20d, 8d, []), 28d, 36d));
        document = document with { PageHeightPoints = 160d, MarginTopPoints = 10d, MarginBottomPoints = 10d,
            PageSettings = DocxPageSettings.Empty, FinalSectionBreak = null,
            BodyElements = [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Public leading content", 12d, 100d)), .. document.BodyElements] };
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
        var painted = layout.Pages.Select((page, index) => (Page: index, Rows: page.Items.OfType<DocxTableRowLayout>().ToArray()))
            .Where(item => item.Rows.Length != 0).ToArray();
        TestAssert.Equal(1, painted.Length);
        TestAssert.True(painted[0].Page > 0 && painted[0].Rows.All(row => row.FragmentCount == 1),
            "A partially filled page defers the containing row, preserving the atomic nested math fallback.");
        TestAssert.Equal(2, NestedMathRows(painted[0].Rows).SelectMany(row => row.Cells)
            .SelectMany(cell => cell.TextLines).Sum(line => line.InlineMath.Count));
    }

    public static void DocxInlineCompositionNestedMathKeepsIndependentContentAcrossContainersAndTransformations()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native nested-math controls use installed Windows fonts.");
        CheckNestedMath(null);
    }

    public static void DocxInlineCompositionNestedMathStaticFramesMatchIndependentVisibleWordOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent static nested origins use installed Windows fonts.");
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("mixed-nested-math-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray().Where(row => row.GetProperty("Container").GetString() == "static"))
        {
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderBlankPages(DocxTests.ReadDocx(NestedMathCasePath("mixed-nested-math-static.docx"), mode), null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))).ToArray();
            var words = row.GetProperty("Math").EnumerateArray().ToArray();
            TestAssert.Equal(words.Length * 4, shows.Length);
            for (int i = 0; i < words.Length; i++)
            {
                TestAssert.Equal(words[i].GetProperty("Text").GetString(), string.Concat(shows.Skip(i * 4).Take(4).Select(show => show.Text)));
                TestAssert.True(Math.Abs(shows[i * 4].X - words[i].GetProperty("X").GetDouble()) < .5d &&
                    Math.Abs(shows[i * 4].Y - words[i].GetProperty("Y").GetDouble()) < .5d,
                    $"The complete nested static frame retains Word's visible physical origin: {mode}/word={i}.");
            }
        }
    }

    internal static void CheckNestedMath(IFontResolver? resolver)
    {
        using var gold = JsonDocument.Parse(File.ReadAllText(NestedMathCasePath("mixed-nested-math-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(NestedMathCasePath(id + ".docx")))));
            string container = row.GetProperty("Container").GetString()!;
            string part = container switch { "footnote" => "word/footnotes.xml", "endnote" => "word/endnotes.xml",
                "static" => "word/nested-math-header.xml", _ => "word/document.xml" };
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            string? original = null;
            foreach (bool transformed in new[] { false, true })
            {
                var document = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed, partOverride: part);
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(row.GetProperty("ExpectedEquations").GetInt32(), text.Split("ratetime", StringSplitOptions.None).Length - 1);
                foreach (var prose in row.GetProperty("ExpectedText").EnumerateArray()) TestAssert.Contains(prose.GetString()!, text);
                TestAssert.Equal(row.GetProperty("ImagePlacements").GetInt32(), pages.Sum(page => page.Images.Sum(image =>
                    Regex.Matches(page.Content, "/" + Regex.Escape(image.ResourceName) + @"\s+Do\b").Count)));
                TestAssert.True(diagnostics.All(diagnostic => diagnostic.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "Prepared ordinary nested content must survive painting in every qualified container.");
                if (container is "footnote" or "endnote")
                {
                    var links = pages.SelectMany(page => page.Annotations).ToArray();
                    TestAssert.Equal(4, links.Length);
                    TestAssert.Equal(4, links.Select(link => link.Destination).Distinct().Count());
                    TestAssert.True(links.All(link => link.Destination is not null && link.Width > 0d && link.Height > 0d),
                        "Nested note content retains every independently authored source/destination.");
                }
                if (original is not null) TestAssert.Equal(Regex.Replace(original, @"\s", ""), Regex.Replace(text, @"\s", ""));
                original = text;
            }
        }
    }

    private static string NestedMathCasePath(string name) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "CompositionProbes", name));

    private static IEnumerable<DocxTableRowLayout> NestedMathRows(IEnumerable<DocxTableRowLayout> rows)
    {
        foreach (var row in rows)
        {
            yield return row;
            foreach (var child in NestedMathRows(row.Cells.SelectMany(cell => cell.NestedRows))) yield return child;
        }
    }
}
