using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionCellStructureKeepsIndependentWordContentAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent cell structure uses installed Windows fonts.");
        CheckEquationRevisions(null, false, "cell-structure-revisions-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionCellStructureKeepsActualChildOwnersAndPrintedLane()
    {
        var context = DocxMarkupContext.FromMode(OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var formatted = ReadMutation("CompositionProbes/table-math-revision-cell-format", OoxPdfDocxMarkupMode.AllMarkup, false, false);
        double scale = DocxRenderer.ResolveWordCompatiblePrintScale(formatted, context);
        foreach (string story in new[] { "body", "header", "footer" })
        foreach (string kind in new[] { "ins", "del" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation(CellStructureFixture(story, kind), mode, false, false);
            var cell = document.EquationParts.SelectMany(part => part.BodyElements.OfType<DocxTableElement>())
                .SelectMany(element => element.Table.Rows).SelectMany(row => row.Cells).Single();
            var owner = cell.Revisions.Single();
            TestAssert.Equal(kind == "ins" ? DocxRevisionKind.TableCellInsertion : DocxRevisionKind.TableCellDeletion, owner.Kind);
            TestAssert.True(owner.PropertyChangeFamily is null && owner.Kind.IsCellStructureRevision() &&
                DocxMathParagraphPolicy.IsSupportedSimpleCell(cell) &&
                cell.Paragraphs.Sum(paragraph => paragraph.Equations.Count) == 4 &&
                cell.Paragraphs.All(paragraph => !paragraph.Revisions.Contains(owner) &&
                    paragraph.Runs.All(run => !run.Revisions.Contains(owner)) &&
                    paragraph.Equations.All(equation => !ReferenceEquals(equation.Revision, owner))),
                "Cell structure owns its metadata while every unchanged child retains its own source visibility and ink.");
            TestAssert.Equal(story == "body" ? scale : 1d, DocxRenderer.ResolveWordCompatiblePrintScale(document, context));
        }
    }

    public static void DocxInlineCompositionCellStructureKeepsIndependentWordLabels()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native cell labels use installed Windows fonts.");
        CheckCellStructureLabels(null);
    }

    internal static void CheckCellStructureLabels(IFontResolver? resolver)
    {
        foreach (string story in new[] { "body", "header", "footer" })
        foreach (string kind in new[] { "ins", "del" })
        {
            var document = ReadMutation(CellStructureFixture(story, kind), OoxPdfDocxMarkupMode.AllMarkup, false, false);
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).Single();
            string text = DocxPageFootnoteNumberingTests.DecodeText(page);
            string label = kind == "ins" ? "Inserted Cells" : "Deleted Cells";
            if (story == "body") TestAssert.Contains(label, text);
            else TestAssert.DoesNotContain(label, text);
            TestAssert.DoesNotContain("Table Editor", text);
        }
    }

    private static string CellStructureFixture(string story, string kind) => "CompositionProbes/" +
        (story == "body" ? "table-math-revision-cell-" + kind : "cell-structure-" + story + "-" + kind);
}
