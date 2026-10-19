using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineBoxTotalCountersKeepIndependentWordMathAndFieldOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent inline-box origins require installed Windows fonts.");
        CheckBoxPageCounters(null, compareWord: true, "inline-box-total-counter-office.json");
    }

    public static void DocxInlineBoxTotalCountersBindOnlyTheirOwnedInputs()
    {
        var (document, parent, equations) = BoxCounterSource("inline-box-child-total-counter");
        TestAssert.Equal(DocxCounterDependencies.PageCount, DocxFieldText.GetDependencies(document.BodyElements));
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(parent, equations, new TrackingMeasurer(),
            12d, 9d, 16d, 1d, default, out var measured, pageCount: 9), "NUMPAGES requires its count without a page-number dependency.");
        var atoms = TestAssert.NotNull(measured);
        var before = TestAssert.NotNull(atoms.Fragments.Single().TextBox);
        TestAssert.Equal("9", CounterText(before, DocxFieldKind.NumPages));
        TestAssert.True(DocxLayoutEngine.TryRebindInlineCounters(atoms, 17, 9, default, out var same) && ReferenceEquals(atoms, same),
            "Changing an unneeded page number must reuse the measured frame.");
        TestAssert.True(DocxLayoutEngine.TryRebindInlineCounters(atoms, null, 10, default, out var rebound), "A known final count rebinds the child.");
        var after = TestAssert.NotNull(TestAssert.NotNull(rebound).Fragments.Single().TextBox);
        TestAssert.Equal("10", CounterText(after, DocxFieldKind.NumPages));
        TestAssert.Equal("9", CounterText(before, DocxFieldKind.NumPages));
        TestAssert.True(!ReferenceEquals(before, after) && ReferenceEquals(before.TextBox, after.TextBox),
            "Count rebinding retains the source owner and previous immutable frame.");
        TestAssert.Equal(1, after.TextLines.Sum(line => line.InlineMath.Count));
        TestAssert.True(after.TextLines.SelectMany(line => line.InlineMath).All(math => ReferenceEquals(equations[math.Math.Equation], math.Math)),
            "Count rebinding reuses prepared mathematics.");
        TestAssert.True(!DocxLayoutEngine.TryRebindInlineCounters(atoms, 17, null, default, out _), "An unrelated page number cannot replace a missing count.");
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineAtoms(parent, equations, new TrackingMeasurer(),
            12d, 9d, 16d, 1d, default, out _, pageNumber: 17), "Unknown counts keep an explicit measurement fallback.");
        var child = DocxBlockTraversal.EnumerateBodyParagraphs(before.TextBox.BodyElements).Single();
        TestAssert.Equal("{NUMPAGES}", child.Runs.Single(run => run.FieldKind == DocxFieldKind.NumPages).Text);
    }

    public static void DocxInlineBoxTotalCountersCombinePageAndCountDependencies()
    {
        var document = ReadMutation("InlineBoxMath/inline-box-child-total-counter", OoxPdfDocxMarkupMode.Final, false, false,
            (xml, w, _) => xml.Descendants(w + "txbxContent").Single().Element(w + "p")!.Add(
                new XElement(w + "fldSimple", new XAttribute(w + "instr", "PAGE"),
                    new XElement(w + "r", new XElement(w + "t", "1")))));
        var parent = document.Paragraphs.Single(paragraph => paragraph.InlineTextBoxes.Count != 0);
        var child = DocxMathParagraphPolicy.EnumerateInlineBoxPreparationCandidates(document, default).Single();
        var equations = child.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
        TestAssert.Equal(DocxCounterDependencies.PageNumber | DocxCounterDependencies.PageCount,
            DocxFieldText.GetDependencies(document.BodyElements));
        bool Measure(int? number, int? count, out DocxMeasuredInlineAtoms? measured) =>
            DocxLayoutEngine.TryMeasureOrderedInlineAtoms(parent, equations, new TrackingMeasurer(),
                12d, 9d, 16d, 1d, default, out measured, pageNumber: number, pageCount: count);
        TestAssert.True(!Measure(3, null, out _) && !Measure(null, 12, out _), "Mixed counters need both independent inputs.");
        TestAssert.True(Measure(3, 12, out var atoms), "Both actual inputs measure a mixed child.");
        var box = TestAssert.NotNull(TestAssert.NotNull(atoms).Fragments.Single().TextBox);
        TestAssert.Equal("3", CounterText(box));
        TestAssert.Equal("12", CounterText(box, DocxFieldKind.NumPages));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxFieldText.GetDependencies(document.BodyElements, cancellation.Token));
    }

    public static void DocxInlineBoxTotalCountersCancelDuringChildRebindingWithoutMutation()
    {
        var (_, parent, equations) = BoxCounterSource("inline-box-child-total-counter");
        var measurer = new BoxCounterMeasurer();
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(parent, equations, measurer,
            12d, 9d, 16d, 1d, default, out var measured, pageCount: 9), "Initial count binding must measure.");
        var atoms = TestAssert.NotNull(measured);
        using var cancellation = new CancellationTokenSource();
        measurer.OnMeasure = cancellation.Cancel;
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryRebindInlineCounters(atoms, null, 10, cancellation.Token, out _));
        TestAssert.Equal("9", CounterText(TestAssert.NotNull(atoms.Fragments.Single().TextBox), DocxFieldKind.NumPages));
    }

    public static void DocxInlineBoxTotalCountersResolvePlainChildLayoutWithoutMathematics()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native plain-child layout uses installed Windows fonts.");
        CheckPlainBoxTotalCount(null);
    }

    internal static void CheckPlainBoxTotalCount(IFontResolver? resolver)
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("InlineBoxMath/inline-box-child-total-counter", mode, false, false, (xml, w, m) =>
            {
                xml.Descendants(m + "oMath").Remove();
                XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                foreach (var extent in xml.Descendants(wp + "extent").Concat(xml.Descendants(a + "xfrm").Elements(a + "ext")))
                    extent.SetAttributeValue("cx", 190 * 12700);
            });
            var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var renderer = new DocxRenderer(resolver, mode, geometry);
            var box = renderer.InspectLayout(document).Pages.Single().Items.Single(item => item.Kind == "InlineTextBox");
            var lines = TestAssert.NotNull(box.TextLines);
            TestAssert.Equal(1, lines.Count);
            TestAssert.Equal("Before inline  after equation. 1".Length, lines.Sum(line => line.TextLength));
            var page = renderer.RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            TestAssert.Equal(1, shows.Count(show => show.Text.Trim() == "1"));
            TestAssert.True(page.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
                "Ordinary boxed count layout cannot depend on mathematical preparation.");
        }
    }
}
