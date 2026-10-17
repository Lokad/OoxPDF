using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionRowVisibilityKeepsIndependentWordViews()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent row visibility uses installed Windows fonts.");
        CheckEquationRevisions(null, false, "row-visibility-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionRowVisibilitySelectsRowsBeforeChildPreparation()
    {
        foreach (string shape in new[] { "sole", "middle", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.Original,
            OoxPdfDocxMarkupMode.SimpleMarkup, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = shape == "sole" ? "table-math-revision-row-del" : "row-visibility-" + shape + "-deleted";
            var document = ReadMutation("CompositionProbes/" + fixture, mode, false, false);
            var rows = document.EquationParts.SelectMany(part => part.BodyElements.OfType<DocxTableElement>())
                .SelectMany(element => element.Table.Rows).ToArray();
            bool visible = mode is OoxPdfDocxMarkupMode.Original or OoxPdfDocxMarkupMode.AllMarkup;
            TestAssert.Equal(shape == "middle" ? visible ? 3 : 2 : visible ? 1 : 0, rows.Length);
            var paragraphs = rows.SelectMany(row => row.Cells).SelectMany(cell => cell.Paragraphs).ToArray();
            TestAssert.Equal(visible ? 4 : 0, paragraphs.Sum(paragraph => paragraph.Equations.Count));
            TestAssert.True(rows.All(DocxMathParagraphPolicy.IsSupportedSimpleRow) && paragraphs.All(paragraph =>
                paragraph.Equations.All(equation => DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasPageInputs: true)
                    .Any(item => ReferenceEquals(item, equation)))),
                "Selection precedes conditional row layout and preparation; visible row marks do not replace actual child owners.");
        }
        var inserted = ReadMutation("CompositionProbes/table-math-revision-row-ins", OoxPdfDocxMarkupMode.Original, false, false);
        TestAssert.Equal(0, inserted.Tables.Count);
    }

    public static void DocxInlineCompositionRowVisibilityFinalPaintOmitsDeletedContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native final row visibility uses installed Windows fonts.");
        CheckFinalRowVisibility(null);
    }

    internal static void CheckFinalRowVisibility(IFontResolver? resolver)
    {
        foreach (string shape in new[] { "sole", "middle", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.SimpleMarkup })
        {
            string fixture = shape == "sole" ? "table-math-revision-row-del" : "row-visibility-" + shape + "-deleted";
            var document = ReadMutation("CompositionProbes/" + fixture, mode, false, false);
            var page = new DocxRenderer(resolver, mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(document, null, default).Single();
            string text = DocxPageFootnoteNumberingTests.DecodeText(page);
            TestAssert.DoesNotContain("Summary:", text);
            TestAssert.DoesNotContain("share the same input series.", text);
            TestAssert.Equal(0, page.Images.Count);
            if (shape == "sole") TestAssert.True(string.IsNullOrWhiteSpace(text), "The deleted sole row leaves no visible text.");
            else if (shape == "middle")
            {
                TestAssert.Contains("Retained row before.", text);
                TestAssert.Contains("Retained row after.", text);
            }
            else TestAssert.Contains("Composed story.", text);
        }
    }
}
