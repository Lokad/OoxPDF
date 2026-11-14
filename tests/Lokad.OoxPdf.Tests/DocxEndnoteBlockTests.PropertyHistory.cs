using System.Reflection;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxEndnoteBlockTests
{
    public static void DocxEndnotePropertyHistoryPreservesUnknownAndAmbiguousFallbacks()
    {
        var baseline = PictureLayout(36, OoxPdfDocxMarkupMode.Final).RelatedStories.Single(s => s.Story.Id == "37");
        var inspect = typeof(DocxLayoutEngine).GetMethod("HasCompleteMeasuredNoteBlocks", BindingFlags.Static | BindingFlags.NonPublic)!
            .CreateDelegate<Func<DocxRelatedStoryLayout, CancellationToken, bool>>();
        TestAssert.True(inspect(baseline, default), "The control has a complete measured picture and source block.");
        foreach (bool cell in new[] { false, true })
        {
            var revision = new DocxRevisionInfo(cell ? DocxRevisionKind.TableCellPropertiesChange : DocxRevisionKind.TableRowPropertiesChange,
                "6801", "Property Reviewer", null, cell ? "tcPrChange" : "trPrChange",
                cell ? DocxRevisionPropertyFamily.Cell : DocxRevisionPropertyFamily.Row, [cell ? "shd" : "trHeight"]);
            DocxRelatedStoryLayout WithHistory(params DocxRevisionInfo[] revisions) => baseline with
            {
                Story = baseline.Story with { BodyElements = baseline.Story.BodyElements.Select(element =>
                    element is DocxTableElement table ? table with
                    {
                        Table = table.Table with { Rows = table.Table.Rows.Select(row => cell
                            ? row with { Cells = row.Cells.Select(source => source with { Revisions = revisions }).ToArray() }
                            : row with { Revisions = revisions }).ToArray() }
                    } : element).ToArray() }
            };
            var valid = WithHistory(revision);
            TestAssert.True(inspect(valid, default), "One resolved property snapshot retains its measured source frame.");
            foreach (var invalid in new[] {
                WithHistory(revision, revision with { Id = "6802" }),
                WithHistory(revision with { Kind = DocxRevisionKind.Insertion, SourceElement = "ins", PropertyChangeFamily = null }),
                WithHistory(revision with { Kind = DocxRevisionKind.MoveFrom }),
                WithHistory(revision with { SourceElement = "pPrChange" }),
                WithHistory(revision with { PropertyChangeFamily = DocxRevisionPropertyFamily.Paragraph }) })
                TestAssert.True(!inspect(invalid, default), "Duplicate, content, move and mismatched property owners keep their existing fallback.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Equal(canceled.Token, TestAssert.Throws<OperationCanceledException>(() => inspect(valid, canceled.Token)).CancellationToken);
        }
    }
}
