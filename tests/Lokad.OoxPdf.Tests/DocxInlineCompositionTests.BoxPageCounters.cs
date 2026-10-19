using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineBoxPageCountersKeepIndependentWordMathAndFieldOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent inline-box origins require installed Windows fonts.");
        CheckBoxPageCounters(null, compareWord: true);
    }

    internal static void CheckBoxPageCounters(IFontResolver? resolver, bool compareWord,
        string goldFile = "inline-box-page-counter-office.json")
    {
        using var gold = ReadInlineBoxMathGold(goldFile);
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "InlineBoxMath", id + ".docx"));
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var geometry = final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var document = DocxTests.ReadDocx(input, mode);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(resolver, mode, geometry);
            var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            var page = pages[0];
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var mathFonts = page.Fonts.Where(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                .Select(font => font.Font).ToHashSet();
            var math = shows.Where(show => mathFonts.Contains(show.Font) && show.Text.Trim().Length != 0).ToArray();
            TestAssert.Equal("x+2y", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(4, math.Length);
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "A known PAGE value cannot remove supported child mathematics.");
            var field = shows.Single(show => show.Text.Trim() == "1");
            if (compareWord)
            {
                var anchors = row.GetProperty("WordMathAnchors").EnumerateArray().ToArray();
                for (int index = 0; index < math.Length; index++)
                    CheckOrigin(math[index].X, page.Height - math[index].Y, math[index].Size, anchors[index]);
                CheckOrigin(field.X, page.Height - field.Y, field.Size, row.GetProperty("WordFieldGlyphs")[0]);
            }
            // Admit the atomic box on page10 after measuring it on page9.
            // The eight completed leading pages reuse plain source prose only.
            var lead = document.Paragraphs[0];
            var elements = new List<DocxBodyElement>();
            for (int index = 0; index < 8; index++)
            {
                elements.Add(new DocxParagraphElement(lead));
                elements.Add(new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null));
            }
            double capacity = document.PageHeightPoints - document.MarginTopPoints - document.MarginBottomPoints;
            elements.Add(new DocxParagraphElement(lead with { SpacingAfterPoints = capacity - 30d }));
            elements.AddRange(document.BodyElements.Skip(1));
            diagnostics.Clear();
            var admitted = renderer.RenderBlankPages(document with { BodyElements = elements }, diagnostics.Add, default).ToArray();
            TestAssert.Equal(10, admitted.Length);
            var tenth = DocxCommentsTests.ReadEmbeddedGlyphTextShows(admitted[9], usePageCoordinates: true).ToArray();
            TestAssert.True(tenth.Count(show => show.Text.Trim() == "10") == 1,
                "The admitted page10 box must bind its field to10: " + string.Concat(tenth.Select(show => show.Text)));
            TestAssert.True(!tenth.Any(show => show.Text.Trim() == "9"), "A moved box must discard its unplaced page9 binding.");
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Page admission retains prepared child math.");
        }

        static void CheckOrigin(double x, double y, double size, JsonElement expected) =>
            TestAssert.True(Math.Abs(x - expected.GetProperty("X").GetDouble()) < 1d &&
                Math.Abs(y - expected.GetProperty("Y").GetDouble()) < 1d &&
                Math.Abs(size - expected.GetProperty("Size").GetDouble()) < .1d,
                "Child math and PAGE glyph retain their independent Word origins and sizes.");
    }

    public static void DocxInlineBoxPageCountersRebindOnlyUnplacedFramesAndKeepSourceOwners()
    {
        var (document, parent, equations) = BoxCounterSource();
        TestAssert.Equal(DocxCounterDependencies.PageNumber, DocxFieldText.GetDependencies(document.BodyElements));
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(parent, equations, new TrackingMeasurer(),
            12d, 9d, 16d, 1d, default, out var measured, pageNumber: 9), "PAGE owns a known page input.");
        var atoms = TestAssert.NotNull(measured);
        TestAssert.Equal(1, atoms.CounterFragmentIndices.Count);
        var before = TestAssert.NotNull(atoms.Fragments.Single().TextBox);
        TestAssert.Equal("9", CounterText(before));
        TestAssert.True(DocxLayoutEngine.TryRebindInlineCounters(atoms, 9, null, default, out var same) && ReferenceEquals(atoms, same),
            "An unchanged binding reuses its measured box.");
        TestAssert.True(DocxLayoutEngine.TryRebindInlineCounters(atoms, 10, null, default, out var moved), "An unplaced box can bind to the next page.");
        var after = TestAssert.NotNull(TestAssert.NotNull(moved).Fragments.Single().TextBox);
        TestAssert.Equal("10", CounterText(after));
        TestAssert.Equal("9", CounterText(before));
        TestAssert.True(!ReferenceEquals(before, after) && ReferenceEquals(before.TextBox, after.TextBox), "Rebinding preserves the source box and its previous immutable frame.");
        TestAssert.Equal(1, after.TextLines.Sum(line => line.InlineMath.Count));
        TestAssert.True(after.TextLines.SelectMany(line => line.InlineMath).All(math => ReferenceEquals(equations[math.Math.Equation], math.Math)),
            "Placement reuses prepared mathematical measurement.");
        TestAssert.True(DocxLayoutEngine.TryRebindInlineCounters(atoms, 10, null, default, out var consumed, firstUnplacedFragment: 1) &&
            ReferenceEquals(atoms, consumed), "Already consumed boxes keep their original binding without remeasurement.");
        TestAssert.True(!DocxLayoutEngine.TryRebindInlineCounters(atoms, null, null, default, out _), "Unknown page inputs retain explicit fallback.");
        var child = DocxBlockTraversal.EnumerateBodyParagraphs(before.TextBox.BodyElements).Single();
        TestAssert.Equal("{PAGE}", child.Runs.Single(run => run.FieldKind == DocxFieldKind.Page).Text);
    }

    public static void DocxInlineBoxPageCountersCancelDuringChildRebindingWithoutMutation()
    {
        var (_, parent, equations) = BoxCounterSource();
        var measurer = new BoxCounterMeasurer();
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(parent, equations, measurer,
            12d, 9d, 16d, 1d, default, out var measured, pageNumber: 9), "Initial child binding must measure.");
        var atoms = TestAssert.NotNull(measured);
        using var cancellation = new CancellationTokenSource();
        measurer.OnMeasure = cancellation.Cancel;
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryRebindInlineCounters(atoms, 10, null, cancellation.Token, out _));
        TestAssert.Equal("9", CounterText(TestAssert.NotNull(atoms.Fragments.Single().TextBox)));
    }

    public static void DocxInlineBoxPageCountersKeepUnknownInputsAndTotalCountFallbacks()
    {
        var (_, parent, equations) = BoxCounterSource();
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(parent, equations, new TrackingMeasurer(),
            12d, 9d, 16d, 1d, default, out _), "A page-dependent child cannot measure without its page input.");
        var source = ReadMutation("InlineBoxMath/inline-box-child-page-counter", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, _) => xml.Descendants(w + "fldSimple").Single().SetAttributeValue(w + "instr", "NUMPAGES"));
        var child = DocxMathParagraphPolicy.EnumerateInlineBoxPreparationCandidates(source, default).Single();
        TestAssert.True(!DocxMathParagraphPolicy.UsesOnlyPageNumberCounters(child) &&
            DocxMathParagraphPolicy.GetPreparationEquations(child).Count == 0, "Total counts keep their existing unavailable-input preparation fallback.");
    }

    private static (DocxDocument Document, DocxParagraph Parent, IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph> Equations) BoxCounterSource(
        string fixture = "inline-box-child-page-counter")
    {
        var document = ReadMutation("InlineBoxMath/" + fixture, OoxPdfDocxMarkupMode.Final, false, false);
        var parent = document.Paragraphs.Single(paragraph => paragraph.InlineTextBoxes.Count != 0);
        var child = DocxMathParagraphPolicy.EnumerateInlineBoxPreparationCandidates(document, default).Single();
        var equations = child.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
        return (document, parent, equations);
    }

    private static string CounterText(DocxInlineTextBoxLayout box, DocxFieldKind kind = DocxFieldKind.Page) =>
        string.Concat(box.TextLines.SelectMany(line => line.Segments)
            .Where(segment => segment.StyleRun.FieldKind == kind).Select(segment => segment.Text));

    private sealed class BoxCounterMeasurer : IDocxTextMeasurer
    {
        public Action? OnMeasure { get; set; }
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            OnMeasure?.Invoke();
            return text.Length * 5d;
        }
    }
}
