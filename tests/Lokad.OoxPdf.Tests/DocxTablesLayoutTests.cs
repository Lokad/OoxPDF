using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTablesLayoutTests
{
    public static void DocxTableLayoutStageKeepsParagraphWithFollowingTableFirstRow()
    {
        DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 60d);
        DocxParagraph heading = DocxTests.CreateDocxLayoutParagraph(
            "Heading",
            10d,
            20d,
            new DocxParagraphKeepRules(true, "1", null, null, null, null));
        var table = new DocxTable(
            null,
            [60d],
            [new DocxTableRow([new DocxTableCell("Cell", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d)]);
        var document = new DocxDocument(
            100d,
            100d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(filler), new DocxParagraphElement(heading), new DocxTableElement(table)],
            [filler, heading],
            [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal("Filler", layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal("Heading", layout.Pages[1].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTableRowLayout>().Count());
    }

    public static void DocxTableLayoutStageEstimatesKeptTableFirstRowWithResolvedPreferredWidth()
    {
        DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 40d);
        DocxParagraph heading = DocxTests.CreateDocxLayoutParagraph(
            "Heading",
            10d,
            10d,
            new DocxParagraphKeepRules(true, "1", null, null, null, null));
        var table = new DocxTable(
            null,
            [100d],
            [new DocxTableRow([new DocxTableCell("aaaa aaaa aaaa aaaa", [], null, null, null, null, [], DocxTableCellMargins.Empty)], null)]) with {PreferredWidthPoints = 25d };
        var document = new DocxDocument(
            100d,
            100d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxParagraphElement(filler), new DocxParagraphElement(heading), new DocxTableElement(table)],
            [filler, heading],
            [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout keptRow = layout.Pages[1].Items.OfType<DocxTableRowLayout>().Single();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal("Filler", layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal("Heading", layout.Pages[1].Items.OfType<DocxTextLineLayout>().Single().Text);
        TestAssert.Equal(25d, keptRow.Table.ResolvedTableWidth);
        TestAssert.True(keptRow.Height > 30d, "The first-row estimate must use the resolved preferred table width so wrapped cell text contributes to keep-with-next pagination.");
    }

    public static void DocxTableLayoutStageAppliesVerticalMergeGeometry()
    {
        var restart = new DocxTableCell(
            "Merged",
            [],
            null,
            null,
            null,
            null,
            [new DocxTableCellBorder("bottom", "single", "000000", "8")],
            DocxTableCellMargins.Empty) with {HasVerticalMerge = true,VerticalMergeValue = "restart" };
        var continuation = new DocxTableCell(
            "Continuation",
            [],
            null,
            null,
            null,
            null,
            [new DocxTableCellBorder("top", "single", "000000", "8")],
            DocxTableCellMargins.Empty) with {HasVerticalMerge = true };
        var table = new DocxTable(
            null,
            [60d],
            [
                new DocxTableRow([restart], 20d),
                new DocxTableRow([continuation], 30d)
            ]);
        var document = new DocxDocument(
            100d,
            100d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table]);

        DocxTableRowLayout[] rows = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .ToArray();

        TestAssert.Equal(2, rows.Length);
        TestAssert.True(Math.Abs(rows[0].Cells[0].Y - 38.08d) < 0.001d, $"Expected merged restart y near 38.08pt, got {rows[0].Cells[0].Y.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(Math.Abs(rows[0].Cells[0].Height - 51.92d) < 0.001d, $"Expected merged restart height near 51.92pt, got {rows[0].Cells[0].Height.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.Equal(DocxTableCellVisualOwnership.OwnCell, rows[0].Cells[0].VisualOwnership);
        TestAssert.True(ReferenceEquals(restart, rows[0].Cells[0].VisualCell), "Restart cells should keep themselves as visual cells.");
        TestAssert.True(rows[1].Cells[0].IsVerticalMergeContinuation, "Continuation cell should be layout-visible but skipped by rendering.");
        TestAssert.Equal(DocxTableCellVisualOwnership.VerticalMergeOwner, rows[1].Cells[0].VisualOwnership);
        TestAssert.Equal(0, rows[1].Cells[0].VerticalMergeOwner?.RowIndex ?? -1);
        TestAssert.Equal(0, rows[1].Cells[0].VerticalMergeOwner?.GridColumnIndex ?? -1);
        TestAssert.True(ReferenceEquals(restart, rows[1].Cells[0].VisualCell), "Continuation cells should expose the restart cell as their visual source.");
        TestAssert.True(Math.Abs(rows[1].Cells[0].Y - 38.08d) < 0.001d, $"Expected merge continuation y near 38.08pt, got {rows[1].Cells[0].Y.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(Math.Abs(rows[1].Cells[0].Height - 30.96d) < 0.001d, $"Expected merge continuation height near 30.96pt, got {rows[1].Cells[0].Height.ToString("0.###", CultureInfo.InvariantCulture)}.");
    }

    public static void DocxTableLayoutStageCarriesVerticalMergeOwnerAcrossPages()
    {
        var filler = new DocxTableCell("Filler", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var restart = new DocxTableCell(
            "Merged",
            [DocxTests.CreateDocxLayoutParagraph("Merged", 10d, 10d)],
            "D9EAD3",
            "clear",
            "auto",
            null,
            [new DocxTableCellBorder("left", "single", "000000", "8")],
            DocxTableCellMargins.Empty) with {HasVerticalMerge = true,VerticalMergeValue = "restart" };
        var continuation = new DocxTableCell(
            "Continuation",
            [],
            null,
            null,
            null,
            null,
            [],
            DocxTableCellMargins.Empty) with {HasVerticalMerge = true };
        var table = new DocxTable(
            null,
            [60d],
            [
                new DocxTableRow([filler], 60d),
                new DocxTableRow([restart], 20d),
                new DocxTableRow([continuation], 30d)
            ]);
        var document = new DocxDocument(
            100d,
            100d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout restartRow = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single(row => row.RowIndex == 1);
        DocxTableRowLayout continuationRow = layout.Pages[1].Items.OfType<DocxTableRowLayout>().Single(row => row.RowIndex == 2);
        DocxTableCellLayout continuationCell = continuationRow.Cells.Single();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.True(restartRow.Cells[0].Height > restartRow.Height, "The restart row should own the full merged span, even when it crosses the page boundary.");
        TestAssert.Equal(DocxTableCellVisualOwnership.OwnCell, restartRow.Cells[0].VisualOwnership);
        TestAssert.True(continuationCell.IsVerticalMergeContinuation, "The second-page row should remain marked as a merge continuation.");
        TestAssert.Equal(DocxTableCellVisualOwnership.VerticalMergeOwner, continuationCell.VisualOwnership);
        TestAssert.Equal(1, continuationCell.VerticalMergeOwner?.RowIndex ?? -1);
        TestAssert.Equal(0, continuationCell.VerticalMergeOwner?.GridColumnIndex ?? -1);
        TestAssert.True(ReferenceEquals(restart, continuationCell.VerticalMergeOwnerCell), "Continuation fragments should retain the restart cell as visual owner across pages.");
        TestAssert.True(ReferenceEquals(restart, continuationCell.VisualCell), "Continuation fragments should resolve their visual cell through layout ownership.");
        TestAssert.Equal("D9EAD3", continuationCell.VerticalMergeOwnerCell?.FillHex ?? string.Empty);
        TestAssert.Equal(60d, continuationCell.Y);
        TestAssert.Equal(30d, continuationCell.Height);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        DocxTableSnapshot tableSnapshot = snapshot.Tables.Single();
        TestAssert.True(tableSnapshot.HasVerticalMerge, "Table-level snapshots should expose vertical-merge presence.");
        TestAssert.Equal(2, tableSnapshot.AuthoredVerticalMergeCellCount);
        TestAssert.Equal(1, tableSnapshot.AuthoredVerticalMergeRestartCellCount);
        TestAssert.Equal(1, tableSnapshot.AuthoredVerticalMergeContinuationCellCount);
        TestAssert.Equal(1, tableSnapshot.LaidOutVerticalMergeContinuationCellCount);
        TestAssert.Equal(0, tableSnapshot.MissingVerticalMergeOwnerCellCount);

        DocxTableCellSnapshot continuationSnapshot = snapshot.Pages[1].TableRows.Single(row => row.RowIndex == 2).Cells.Single();
        TestAssert.Equal("VerticalMergeOwner", continuationSnapshot.VisualOwnership);
        TestAssert.Equal(1, continuationSnapshot.VerticalMergeOwnerRowIndex ?? -1);
        TestAssert.Equal(0, continuationSnapshot.VerticalMergeOwnerGridColumnIndex ?? -1);
        TestAssert.Equal(6, continuationSnapshot.TextLength);
        TestAssert.Equal(6, continuationSnapshot.VisualTextLength);
        TestAssert.Equal(1, continuationSnapshot.VisualParagraphCount);
        TestAssert.True(continuationCell.TextLines.Count != 0, "A page-crossing merge continuation should carry owner text lines for page-local clipped PDF emission.");
    }

    public static void DocxTableLayoutStageCarriesVerticalMergeOwnerAfterRepeatedHeader()
    {
        var header = new DocxTableRow(
            [new DocxTableCell("Header", [DocxTests.CreateDocxLayoutParagraph("Header", 10d, 10d)], null, null, null, null, [], DocxTableCellMargins.Empty)],
            10d) with {IsHeader = true };
        var filler = new DocxTableRow(
            [new DocxTableCell("Filler", [DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 10d)], null, null, null, null, [], DocxTableCellMargins.Empty)],
            50d);
        var restart = new DocxTableCell(
            "Merged",
            [DocxTests.CreateDocxLayoutParagraph("Merged", 10d, 10d)],
            "D9EAD3",
            "clear",
            "auto",
            null,
            [new DocxTableCellBorder("left", "single", "000000", "8")],
            DocxTableCellMargins.Empty) with {HasVerticalMerge = true,VerticalMergeValue = "restart" };
        var continuation = new DocxTableCell(
            "Continuation",
            [],
            null,
            null,
            null,
            null,
            [],
            DocxTableCellMargins.Empty) with {HasVerticalMerge = true };
        var table = new DocxTable(
            null,
            [60d],
            [
                header,
                filler,
                new DocxTableRow([restart], 20d),
                new DocxTableRow([continuation], 30d)
            ]);
        var document = new DocxDocument(
            100d,
            100d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableCellLayout continuationCell = secondPageRows.Single(row => row.RowIndex == 3).Cells.Single();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(2, secondPageRows.Length);
        TestAssert.Equal(0, secondPageRows[0].RowIndex);
        TestAssert.True(secondPageRows[0].IsHeader, "The second page should repeat the table header before the merge continuation.");
        TestAssert.Equal(3, secondPageRows[1].RowIndex);
        TestAssert.True(continuationCell.IsVerticalMergeContinuation, "The row after the repeated header should remain a merge continuation.");
        TestAssert.Equal(DocxTableCellVisualOwnership.VerticalMergeOwner, continuationCell.VisualOwnership);
        TestAssert.Equal(2, continuationCell.VerticalMergeOwner?.RowIndex ?? -1);
        TestAssert.Equal(0, continuationCell.VerticalMergeOwner?.GridColumnIndex ?? -1);
        TestAssert.True(ReferenceEquals(restart, continuationCell.VisualCell), "The repeated header must not become the visual owner of the merge continuation.");
        TestAssert.Equal("D9EAD3", continuationCell.VisualCell.FillHex ?? string.Empty);
        TestAssert.Equal(50d, continuationCell.Y);
        TestAssert.Equal(30d, continuationCell.Height);

        DocxTableCellSnapshot continuationSnapshot = DocxLayoutSnapshot.FromLayout(layout)
            .Pages[1]
            .TableRows
            .Single(row => row.RowIndex == 3)
            .Cells
            .Single();
        TestAssert.Equal("VerticalMergeOwner", continuationSnapshot.VisualOwnership);
        TestAssert.Equal(2, continuationSnapshot.VerticalMergeOwnerRowIndex ?? -1);
        TestAssert.Equal(6, continuationSnapshot.TextLength);
        TestAssert.Equal(6, continuationSnapshot.VisualTextLength);
        TestAssert.True(continuationCell.TextLines.Count != 0, "A merge continuation after a repeated header should carry owner text lines for page-local clipped PDF emission.");
    }

    public static void DocxTableLayoutStagePlacesCellsBeforePdfEmission()
    {
        var table = new DocxTable(
            null,
            [60d, 40d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded: null, cancellationToken: CancellationToken.None);

        DocxTableRowLayout row = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        TestAssert.Equal(170d, row.Y);
        TestAssert.Equal(20d, row.Height);
        TestAssert.Equal(10d, row.Cells[0].X);
        TestAssert.Equal(60d, row.Cells[0].Width);
        TestAssert.Equal(70d, row.Cells[1].X);
        TestAssert.Equal(40d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageScalesGridToPreferredWidth()
    {
        var table = new DocxTable(
            null,
            [60d, 60d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]) with {PreferredWidthPoints = 60d };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(30d, row.Cells[0].Width);
        TestAssert.Equal(40d, row.Cells[1].X);
        TestAssert.Equal(30d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageDoesNotClampExplicitDxaWidthToBody()
    {
        var table = new DocxTable(
            null,
            [110d, 110d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]) with {PreferredWidthPoints = 220d,PreferredWidthValue = "4400",PreferredWidthType = "dxa" };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(110d, row.Cells[0].Width);
        TestAssert.Equal(120d, row.Cells[1].X);
        TestAssert.Equal(110d, row.Cells[1].Width);
        TestAssert.Equal(220d, row.Table.ResolvedTableWidth);
    }

    public static void DocxTableLayoutStageScalesGridToPercentagePreferredWidth()
    {
        var table = new DocxTable(
            null,
            [60d, 60d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]) with {PreferredWidthValue = "2500",PreferredWidthType = "pct" };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(45d, row.Cells[0].Width);
        TestAssert.Equal(55d, row.Cells[1].X);
        TestAssert.Equal(45d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageDistributesMissingGridAcrossAvailableWidth()
    {
        var table = new DocxTable(
            null,
            [72d, 72d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]) with {HasExplicitGrid = false };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.True(!row.Table.HasExplicitGrid, "Layout context should preserve missing-grid provenance after width resolution.");
        TestAssert.Equal(90d, row.Cells[0].Width);
        TestAssert.Equal(100d, row.Cells[1].X);
        TestAssert.Equal(90d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageDistributesAutoLayoutByContentWidth()
    {
        // Office A/B (comment-table autofit probes, Word-COM rendered): tables without a
        // fixed layout distribute width by column content instead of the grid, so skewed
        // content yields skewed columns.
        var table = new DocxTable(
            null,
            [100d, 100d],
            [new DocxTableRow([
                new DocxTableCell(string.Empty, [DocxTests.CreateDocxLayoutParagraph("AA", 10d, 12d)], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell(string.Empty, [DocxTests.CreateDocxLayoutParagraph("AAAAAAAA", 10d, 12d)], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.True(Math.Abs(row.Cells[0].Width - 36d) < 1d, "Auto-layout columns should share width by content. Width0=" + row.Cells[0].Width.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(row.Cells[1].Width - 144d) < 1d, "Auto-layout columns should share width by content. Width1=" + row.Cells[1].Width.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxTableLayoutStageKeepsFixedLayoutGridWidths()
    {
        // Companion guard: explicit fixed layout keeps grid widths regardless of content.
        var table = new DocxTable(
            "fixed",
            [100d, 100d],
            [new DocxTableRow([
                new DocxTableCell(string.Empty, [DocxTests.CreateDocxLayoutParagraph("AA", 10d, 12d)], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell(string.Empty, [DocxTests.CreateDocxLayoutParagraph("AAAAAAAA", 10d, 12d)], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.True(Math.Abs(row.Cells[0].Width - 90d) < 1d, "Fixed-layout columns should keep grid widths. Width0=" + row.Cells[0].Width.ToString(CultureInfo.InvariantCulture));
        TestAssert.True(Math.Abs(row.Cells[1].Width - 90d) < 1d, "Fixed-layout columns should keep grid widths. Width1=" + row.Cells[1].Width.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxTableLayoutStageKeepsAutoPreferredCellWidths()
    {
        // Companion guard: explicit preferred cell widths win over measured content.
        DocxTableCell wideCell = new DocxTableCell(string.Empty, [DocxTests.CreateDocxLayoutParagraph("AA", 10d, 12d)], null, null, null, null, [], DocxTableCellMargins.Empty) with
        {
            PreferredWidthPoints = 150d
        };
        DocxTableCell narrowCell = new DocxTableCell(string.Empty, [DocxTests.CreateDocxLayoutParagraph("AAAAAAAA", 10d, 12d)], null, null, null, null, [], DocxTableCellMargins.Empty) with
        {
            PreferredWidthPoints = 30d
        };
        var table = new DocxTable(
            null,
            [100d, 100d],
            [new DocxTableRow([
                wideCell,
                narrowCell
            ], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.True(row.Cells[0].Width > row.Cells[1].Width, "Preferred cell widths should survive auto layout. Width0=" + row.Cells[0].Width.ToString(CultureInfo.InvariantCulture) + " Width1=" + row.Cells[1].Width.ToString(CultureInfo.InvariantCulture));
    }

    public static void DocxTableLayoutStageDistributesMissingGridWithSpansAcrossLogicalColumns()
    {
        var table = new DocxTable(
            null,
            [72d, 72d, 72d],
            [new DocxTableRow([
                new DocxTableCell("wide", [], null, null, null, null, [], DocxTableCellMargins.Empty) with {GridSpan = 2,GridSpanValue = "2" },
                new DocxTableCell("tail", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]) with {HasExplicitGrid = false };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(120d, row.Cells[0].Width);
        TestAssert.Equal(130d, row.Cells[1].X);
        TestAssert.Equal(60d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageInfersEmptyMissingGridFromRowSpans()
    {
        var table = new DocxTable(
            null,
            [],
            [new DocxTableRow([
                new DocxTableCell("wide", [], null, null, null, null, [], DocxTableCellMargins.Empty) with {GridSpan = 2,GridSpanValue = "2" },
                new DocxTableCell("tail", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]) with {HasExplicitGrid = false };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(120d, row.Cells[0].Width);
        TestAssert.Equal(130d, row.Cells[1].X);
        TestAssert.Equal(60d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageUsesFirstRowCellPreferredWidths()
    {
        var table = new DocxTable(
            null,
            [60d, 60d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty) with {PreferredWidthPoints = 40d },
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty) with {PreferredWidthPoints = 80d }
            ], 20d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(40d, row.Cells[0].Width);
        TestAssert.Equal(50d, row.Cells[1].X);
        TestAssert.Equal(80d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageUsesLaterRowCellPreferredWidths()
    {
        var table = new DocxTable(
            null,
            [60d, 60d],
            [
                new DocxTableRow([
                    new DocxTableCell("span", [], null, null, null, null, [], DocxTableCellMargins.Empty) with {GridSpan = 2,GridSpanValue = "2" }
                ], 20d),
                new DocxTableRow([
                    new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty) with {PreferredWidthPoints = 40d },
                    new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty) with {PreferredWidthPoints = 80d }
                ], 20d)
            ]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout[] rows = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .ToArray();

        TestAssert.Equal(2, rows.Length);
        TestAssert.Equal(120d, rows[0].Cells[0].Width);
        TestAssert.Equal(40d, rows[1].Cells[0].Width);
        TestAssert.Equal(50d, rows[1].Cells[1].X);
        TestAssert.Equal(80d, rows[1].Cells[1].Width);
    }

    public static void DocxTableLayoutStageUsesPercentageCellPreferredWidths()
    {
        var table = new DocxTable(
            null,
            [100d, 100d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty) with {PreferredWidthValue = "1250",PreferredWidthType = "pct" },
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty) with {PreferredWidthValue = "3750",PreferredWidthType = "pct" }
            ], 20d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(45d, row.Cells[0].Width);
        TestAssert.Equal(55d, row.Cells[1].X);
        TestAssert.Equal(135d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageAppliesGridSpanWidths()
    {
        var table = new DocxTable(
            null,
            [40d, 60d, 80d],
            [new DocxTableRow([
                new DocxTableCell("wide", [], null, null, null, null, [], DocxTableCellMargins.Empty) with {GridSpan = 2,GridSpanValue = "2" },
                new DocxTableCell("tail", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10d, row.Cells[0].X);
        TestAssert.Equal(100d, row.Cells[0].Width);
        TestAssert.Equal(110d, row.Cells[1].X);
        TestAssert.Equal(80d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageAppliesTableIndent()
    {
        var table = new DocxTable(
            null,
            [40d],
            [new DocxTableRow([
                new DocxTableCell("indented", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]) with {IndentPoints = 18d,IndentValue = "360",IndentType = "dxa" };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        // Office A/B (w66/w68 indent probes): tblInd pins TEXT at the indent origin, so the grid hangs left by max(borderHalf, margin) = 0.48 here (borderless, unset margins).
        TestAssert.Equal(27.52d, row.Cells[0].X);
        TestAssert.Equal(40d, row.Cells[0].Width);
    }

    public static void DocxTableLayoutStageAppliesCellSpacing()
    {
        var table = new DocxTable(
            null,
            [40d, 60d],
            [new DocxTableRow([
                new DocxTableCell("left", [], null, null, null, null, [], DocxTableCellMargins.Empty),
                new DocxTableCell("right", [], null, null, null, null, [], DocxTableCellMargins.Empty)
            ], 20d)]) with {CellSpacingPoints = 6d,CellSpacingValue = "120",CellSpacingType = "dxa" };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10d, row.Cells[0].X);
        TestAssert.Equal(40d, row.Cells[0].Width);
        TestAssert.Equal(56d, row.Cells[1].X);
        TestAssert.Equal(60d, row.Cells[1].Width);
    }

    public static void DocxTableLayoutStageBuildsParagraphTextLinesInsideCells()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        var firstParagraph = new DocxParagraph(
            [new DocxTextRun("Alpha", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var secondParagraph = new DocxParagraph(
            [
                new DocxTextRun("B", 14d, "336699", false, true, false, null, null),
                new DocxTextRun("G", 14d, "993333", true, false, false, null, null)
            ],
            [],
            null,
            DocxTextAlignment.Center,
            "center",
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell("Alpha BG", [firstParagraph, secondParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 44d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "Alpha BG".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded, CancellationToken.None);

        DocxTableCellLayout cellLayout = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
        TestAssert.Equal(2, cellLayout.TextLines.Count);
        TestAssert.Equal("Alpha", cellLayout.TextLines[0].Text);
        TestAssert.Equal(11d, cellLayout.TextLines[0].FontSize);
        TestAssert.Equal("BG", cellLayout.TextLines[1].Text);
        TestAssert.Equal(14d, cellLayout.TextLines[1].FontSize);
        TestAssert.Equal("336699", cellLayout.TextLines[1].StyleRun.ColorHex ?? string.Empty);
        TestAssert.Equal(2, cellLayout.TextLines[1].Segments.Count);
        TestAssert.Equal("B", cellLayout.TextLines[1].Segments[0].Text);
        TestAssert.Equal("336699", cellLayout.TextLines[1].Segments[0].StyleRun.ColorHex ?? string.Empty);
        TestAssert.Equal("G", cellLayout.TextLines[1].Segments[1].Text);
        TestAssert.Equal("993333", cellLayout.TextLines[1].Segments[1].StyleRun.ColorHex ?? string.Empty);
        TestAssert.True(cellLayout.TextLines[1].Segments[1].X > cellLayout.TextLines[1].Segments[0].X, "Second table-cell run segment should be positioned after the first segment.");
        TestAssert.True(cellLayout.TextLines[1].X > cellLayout.TextLines[0].X, "Centered table-cell paragraph text should be positioned from line width, not flattened at the left inset.");
        TestAssert.True(cellLayout.TextLines[1].BaselineY < cellLayout.TextLines[0].BaselineY, "Separate table-cell paragraphs should produce separate baselines.");
    }

    public static void DocxTableLayoutStageUsesCellBodyElementsForInlineImages()
    {
        var image = new DocxInlineImage(18d, 12d, "image/png", [1, 2, 3], "/word/media/image1.png");
        var paragraph = new DocxParagraph(
            [],
            [image],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxParagraphElement(paragraph)]
        };
        var table = new DocxTable(null, [60d], [new DocxTableRow([cell], 36d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded: null, cancellationToken: CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.Equal(1, cellLayout.InlineImages.Count);
        TestAssert.Equal(18d, cellLayout.InlineImages[0].Width);
        TestAssert.Equal(12d, cellLayout.InlineImages[0].Height);
    }

    public static void DocxTableLayoutStageLaysOutNestedTableCellBodies()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        DocxParagraph nestedParagraph = DocxTests.CreateDocxLayoutParagraph("Nested", 10d, 10d);
        var nestedCell = new DocxTableCell("Nested", [nestedParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var nestedTable = new DocxTable(null, [50d], [new DocxTableRow([nestedCell], null)]);
        var outerCell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxTableElement(nestedTable)]
        };
        var outerTable = new DocxTable(null, [80d], [new DocxTableRow([outerCell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(outerTable)], [outerTable, nestedTable]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font.Value.Font, "Nested".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded, CancellationToken.None);
        DocxTableRowLayout outerRow = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        DocxTableCellLayout outerCellLayout = outerRow.Cells.Single();
        DocxTableRowLayout nestedRow = outerCellLayout.NestedRows.Single();
        DocxTextLineLayout nestedLine = nestedRow.Cells.Single().TextLines.Single();
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);
        var renderer = new DocxRenderer(new DocxTests.SingleResolutionFontResolver(font.Value.Resolution), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxTextEmissionLineSnapshot[] emissionLines = renderer.InspectTextEmission(document).Lines
            .Where(line => !line.IsStaticStory)
            .ToArray();

        TestAssert.Equal("Nested", nestedLine.Text);
        TestAssert.True(outerRow.Height >= nestedRow.Height, "Outer table row height should include nested table content height.");
        TestAssert.True(nestedRow.Y >= outerCellLayout.Y && nestedRow.Y + nestedRow.Height <= outerCellLayout.Y + outerCellLayout.Height + 0.001d, "Nested table row should be placed inside the parent cell fragment.");
        TestAssert.Equal("Nested".Length, snapshot.Pages.Single().TableRows.Single().TextLength);
        TestAssert.True(emissionLines.Single().TextLength >= "Nested".Length, "Nested table-cell text should participate in renderer text-emission enumeration.");
    }

    public static void DocxTableLayoutStagePreservesMixedCellBodyOrderAroundNestedTables()
    {
        (FontFaceResolution Resolution, OpenTypeFont Font)? font = DocxTests.FindUsableInstalledFont();
        if (font is null)
        {
            TestAssert.Skip("Environmental precondition not met: (font is null)");
        }

        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d);
        DocxParagraph nestedParagraph = DocxTests.CreateDocxLayoutParagraph("Nested", 10d, 10d);
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 10d, 10d);
        var nestedCell = new DocxTableCell("Nested", [nestedParagraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var nestedTable = new DocxTable(null, [50d], [new DocxTableRow([nestedCell], null)]);
        var outerCell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxParagraphElement(before), new DocxTableElement(nestedTable), new DocxParagraphElement(after)]
        };
        var outerTable = new DocxTable(null, [90d], [new DocxTableRow([outerCell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(outerTable)], [outerTable, nestedTable]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font.Value.Font, "BeforeNestedAfter".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTableCellLayout outerCellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();
        DocxTextLineLayout[] outerLines = outerCellLayout.TextLines.ToArray();
        DocxTableRowLayout nestedRow = outerCellLayout.NestedRows.Single();

        TestAssert.Equal(2, outerLines.Length);
        TestAssert.Equal("Before", outerLines[0].Text);
        TestAssert.Equal("After", outerLines[1].Text);
        TestAssert.True(
            outerLines[0].BaselineY >= nestedRow.Y + nestedRow.Height,
            $"Paragraph before the nested table should remain above the nested table block. BeforeBaseline={outerLines[0].BaselineY}, NestedTop={nestedRow.Y + nestedRow.Height}, NestedBottom={nestedRow.Y}.");
        TestAssert.True(
            outerLines[1].BaselineY < nestedRow.Y,
            $"Paragraph after the nested table should be laid out below the nested table block. AfterBaseline={outerLines[1].BaselineY}, NestedTop={nestedRow.Y + nestedRow.Height}, NestedBottom={nestedRow.Y}.");
    }

    public static void DocxTableLayoutStageSplitsRowsAtCellPageBreakElements()
    {
        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d);
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 10d, 10d);
        // RV06 cellbreak probe: explicit run breaks never fragment cell rows (Office
        // flows through); only other provenances (like page-break-before) split here.
        var cell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(before),
                new DocxPageBreakElement(DocxBreakSourceKind.PageBreakBefore, "page", null),
                new DocxParagraphElement(after)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout[] rowFragments = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(2, rowFragments.Length);
        TestAssert.Equal(0, rowFragments[0].FragmentIndex);
        TestAssert.Equal(1, rowFragments[1].FragmentIndex);
        TestAssert.Equal(2, rowFragments[0].FragmentCount);
        TestAssert.Equal(2, rowFragments[1].FragmentCount);
        TestAssert.Equal("CellPageBreak", rowFragments[0].FragmentReason);
        TestAssert.Equal("CellPageBreak", rowFragments[1].FragmentReason);
        TestAssert.Equal("Before", rowFragments[0].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("After", rowFragments[1].Cells.Single().TextLines.Single().Text);
        DocxTableRowSnapshot[] rowSnapshots = DocxLayoutSnapshot.FromLayout(layout).Pages
            .SelectMany(page => page.TableRows)
            .OrderBy(row => row.FragmentIndex)
            .ToArray();
        TestAssert.Equal("CellPageBreak", rowSnapshots[0].FragmentReason);
        TestAssert.Equal("CellPageBreak", rowSnapshots[1].FragmentReason);
    }

    public static void DocxTableLayoutStageMovesCellPageBreakSplitToNextPageWhenFirstFragmentDoesNotFit()
    {
        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 20d);
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 10d, 10d);
        var fillerCell = new DocxTableCell("Filler", [DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 10d)], null, null, null, null, [], DocxTableCellMargins.Empty);
        var splitCell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(before),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(after)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([fillerCell], 70d), new DocxTableRow([splitCell], null)]);
        var document = new DocxDocument(
            100d,
            100d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout[] rowFragments = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();

        // RV06 cellbreak probe: run breaks flow through without fragmenting rows.
        TestAssert.True(rowFragments.Length >= 1, "Flowing rows should still lay out across natural overflow.");
        TestAssert.True(rowFragments.All(fragment => fragment.FragmentReason != "CellPageBreak"), "No fragment should come from an in-cell run break.");
        TestAssert.Equal("Filler", firstPageRows[0].Cells.Single().TextLines.Single().Text);
        string[] splitTexts = layout.Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .Where(row => row.RowIndex == 1)
            .SelectMany(row => row.Cells.Single().TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.Equal("Before| |After", string.Join("|", splitTexts));
    }

    public static void DocxTableLayoutStageSplitsOversizedCellPageBreakFragmentAcrossFreshPages()
    {
        DocxParagraph[] beforeParagraphs = Enumerable.Range(1, 10)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Before" + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 10d, 10d);
        var splitCell = new DocxTableCell(string.Empty, beforeParagraphs.Append(after).ToArray(), null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = beforeParagraphs
                .Select<DocxParagraph, DocxBodyElement>(paragraph => new DocxParagraphElement(paragraph))
                .Append(new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null))
                .Append(new DocxParagraphElement(after))
                .ToArray()
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([splitCell], null)]);
        var document = new DocxDocument(
            100d,
            100d,
            10d,
            10d,
            10d,
            10d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [new DocxTableElement(table)],
            [],
            [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout[] rowFragments = layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();

        // RV06 cellbreak probe: run breaks flow through; rows split only by natural
        // overflow, never for the break. All content stays present in order with the
        // spill row after the break.
        TestAssert.True(rowFragments.Length >= 1, "Flowing rows should still lay out across natural overflow.");
        TestAssert.True(rowFragments.All(fragment => fragment.FragmentReason != "CellPageBreak"), "No fragment should come from an in-cell run break.");
        string[] oversizedTexts = rowFragments
            .SelectMany(fragment => fragment.Cells.Single().TextLines)
            .Select(line => line.Text)
            .ToArray();
        TestAssert.Equal("Before1|Before2|Before3|Before4|Before5|Before6|Before7|Before8|Before9|Before10| |After", string.Join("|", oversizedTexts));
    }

    public static void DocxTableLayoutStageUsesEarliestCellPageBreakAsRowBoundary()
    {
        DocxParagraph earlyBefore = DocxTests.CreateDocxLayoutParagraph("EarlyBefore", 10d, 10d);
        DocxParagraph earlyAfter = DocxTests.CreateDocxLayoutParagraph("EarlyAfter", 10d, 10d);
        DocxParagraph laterFirst = DocxTests.CreateDocxLayoutParagraph("LaterFirst", 10d, 10d);
        DocxParagraph laterMiddle = DocxTests.CreateDocxLayoutParagraph("LaterMiddle", 10d, 10d);
        DocxParagraph laterAfter = DocxTests.CreateDocxLayoutParagraph("LaterAfter", 10d, 10d);
        var earlyCell = new DocxTableCell(string.Empty, [earlyBefore, earlyAfter], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(earlyBefore),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(earlyAfter)
            ]
        };
        var laterCell = new DocxTableCell(string.Empty, [laterFirst, laterMiddle, laterAfter], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(laterFirst),
                new DocxParagraphElement(laterMiddle),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(laterAfter)
            ]
        };
        DocxTable table = new(null, [60d, 60d], [new DocxTableRow([earlyCell, laterCell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout[] rowFragments = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .ToArray();

        // RV06 cellbreak probe: run breaks flow through without fragmenting rows.
        TestAssert.Equal(1, rowFragments.Length);
        TestAssert.Equal("EarlyBefore| |EarlyAfter", string.Join("|", rowFragments[0].Cells[0].TextLines.Select(line => line.Text)));
        TestAssert.Equal("LaterFirst|LaterMiddle| |LaterAfter", string.Join("|", rowFragments[0].Cells[1].TextLines.Select(line => line.Text)));
    }

    public static void DocxTableLayoutStageKeepsCellImagesOnAuthoredSideOfPageBreak()
    {
        var beforeImage = new DocxInlineImage(12d, 8d, "image/png", [1, 2, 3], "/word/media/before.png");
        var afterImage = new DocxInlineImage(10d, 6d, "image/png", [4, 5, 6], "/word/media/after.png");
        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 10d) with { Images = [beforeImage] };
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 10d) with { Images = [afterImage] };
        var cell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(before),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(after)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout[] rowFragments = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .ToArray();

        // RV06 cellbreak probe: run breaks flow through without fragmenting rows.
        TestAssert.Equal(1, rowFragments.Length);
        DocxInlineImageLayout[] images = rowFragments[0].Cells.Single().InlineImages.ToArray();
        TestAssert.Equal(2, images.Length);
        TestAssert.Equal(0, images[0].SourceParagraphIndex ?? -1);
        TestAssert.Equal(1, images[1].SourceParagraphIndex ?? -1);
        TestAssert.True(images[0].Image == beforeImage, "The single row should keep the image before the authored page break.");
        TestAssert.True(images[1].Image == afterImage, "The single row should keep the image after the authored page break.");
    }

    public static void DocxTableRowHeightCountsBreakSpillAfterMixedParagraph()
    {
        // RV06 cellbreak probe: the height estimate must reserve the spill row
        // whenever the previous paragraph emitted text lines, including a mixed
        // paragraph that also carries a block image. A run break reserves exactly
        // one laid-out line, like an empty paragraph does.
        static double MeasureCellRowHeight(params DocxBodyElement[] elements)
        {
            var cell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
            {
                BodyElements = elements
            };
            DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
            DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
            return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
                .Pages.Single().Items.OfType<DocxTableRowLayout>().Single().FullRowHeight;
        }

        var blockImage = new DocxInlineImage(10d, 6d, "image/png", [7, 8, 9], "/word/media/block.png");
        DocxParagraph mixed = DocxTests.CreateDocxLayoutParagraph("Mixed", 10d, 10d) with { Images = [blockImage] };
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 10d, 10d);
        DocxParagraph empty = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 10d);
        // The break para is style-less, so it carries the document-default 8pt after
        // (edge-t4, Word 16.0); the explicitly unspaced surroundings do not suppress it.
        TestAssert.Equal(
            50d,
            MeasureCellRowHeight(new DocxParagraphElement(mixed), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxParagraphElement(after)));
        TestAssert.Equal(
            MeasureCellRowHeight(new DocxParagraphElement(mixed), new DocxParagraphElement(empty), new DocxParagraphElement(after)),
            42d);
    }

    public static void DocxTableLayoutStagePlacesCellBlockImageBelowBreakSpillRow()
    {
        // RV06 cellbreak probe: the images walk mirrors the text walk at in-cell
        // breaks, so a block image after the break sits exactly where it would after
        // an empty default paragraph. The reference empty models the style-less break
        // para with the document-default 8pt after (edge-t4, Word 16.0).
        static double MeasureBlockImageY(params DocxBodyElement[] elements)
        {
            var cell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
            {
                BodyElements = elements
            };
            DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
            DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
            return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
                .Pages.Single().Items.OfType<DocxTableRowLayout>().Single().Cells.Single().InlineImages.Single().Y;
        }

        var afterImage = new DocxInlineImage(10d, 6d, "image/png", [4, 5, 6], "/word/media/after.png");
        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d);
        DocxParagraph imageOnly = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 10d) with { Images = [afterImage] };
        DocxParagraph breakModel = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 10d) with { SpacingAfterPoints = 8d };
        double emptyY = MeasureBlockImageY(new DocxParagraphElement(before), new DocxParagraphElement(breakModel), new DocxParagraphElement(imageOnly));
        double breakY = MeasureBlockImageY(new DocxParagraphElement(before), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxParagraphElement(imageOnly));
        TestAssert.True(
            Math.Abs(breakY - emptyY) < 1e-9d,
            $"The block image after an in-cell break should sit where it would after an empty line. breakY={breakY} emptyY={emptyY}.");
    }

    public static void DocxTableRowHeightCountsBreakSpillAfterEmptyParagraph()
    {
        // RV06 cellbreak probe (edge-t1, Word 16.0): an empty paragraph before an
        // in-cell break still yields its line, and the break still spills one row
        // after it, so the break reserves exactly what an empty paragraph does.
        static double MeasureCellRowHeight(params DocxBodyElement[] elements)
        {
            var cell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
            {
                BodyElements = elements
            };
            DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
            DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
            return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
                .Pages.Single().Items.OfType<DocxTableRowLayout>().Single().FullRowHeight;
        }

        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d);
        DocxParagraph empty = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 10d);
        // The break para is style-less, so the reference empty models it with the
        // document-default 8pt after (edge-t4, Word 16.0).
        DocxParagraph breakModel = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 10d) with { SpacingAfterPoints = 8d };
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 10d, 10d);
        TestAssert.Equal(
            MeasureCellRowHeight(new DocxParagraphElement(before), new DocxParagraphElement(empty), new DocxParagraphElement(breakModel), new DocxParagraphElement(after)),
            MeasureCellRowHeight(new DocxParagraphElement(before), new DocxParagraphElement(empty), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxParagraphElement(after)));
    }

    public static void DocxTableRowHeightFlowsSpacingAcrossInCellBreak()
    {
        // RV06 cellbreak probe (edge-t0, Word 16.0): style-less rows pitch 25pt =
        // 17pt line plus 8pt after, so an in-cell break behaves as an empty default
        // paragraph: pending after-spacing is consumed before the spill row and the
        // default after-spacing applies after it (resetting would pitch 17pt).
        static double MeasureCellRowHeight(params DocxBodyElement[] elements)
        {
            var cell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
            {
                BodyElements = elements
            };
            DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
            DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
            return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
                .Pages.Single().Items.OfType<DocxTableRowLayout>().Single().FullRowHeight;
        }

        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d) with { SpacingAfterPoints = 8d };
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 10d, 10d);
        TestAssert.Equal(
            46d,
            MeasureCellRowHeight(new DocxParagraphElement(before), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxParagraphElement(after)));
    }

    public static void DocxTableLayoutStagePlacesNestedTableBelowBreakSpillRow()
    {
        // RV06 cellbreak probe (edge-t3, Word 16.0): a nested table after an in-cell
        // break sits where it would after an empty paragraph line, below the previous
        // line plus one spill row. Pre-fix the nested walk skipped the break and the
        // table floated one line too high.
        static double MeasureNestedTop(DocxTable nested, params DocxBodyElement[] elements)
        {
            var cell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
            {
                BodyElements = elements
            };
            DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
            DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table, nested]);
            return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
                .Pages.Single().Items.OfType<DocxTableRowLayout>().Single().Cells.Single().NestedRows.Single().Y;
        }

        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d);
        // The break para is style-less, so the reference empty models it with the
        // document-default 8pt after (edge-t4, Word 16.0).
        DocxParagraph breakModel = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 10d) with { SpacingAfterPoints = 8d };
        DocxTable probe = DocxTests.CreateSingleCellTable("Inside", 12d);
        double emptyTop = MeasureNestedTop(probe, new DocxParagraphElement(before), new DocxParagraphElement(breakModel), new DocxTableElement(probe));
        double breakTop = MeasureNestedTop(probe, new DocxParagraphElement(before), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxTableElement(probe));
        TestAssert.True(
            Math.Abs(breakTop - emptyTop) < 1e-9d,
            $"The nested table after an in-cell break should sit where it would after an empty line. breakTop={breakTop} emptyTop={emptyTop}.");
        DocxParagraph spaced = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d) with { SpacingAfterPoints = 8d };
        double spacedEmptyTop = MeasureNestedTop(probe, new DocxParagraphElement(spaced), new DocxParagraphElement(breakModel), new DocxTableElement(probe));
        double spacedBreakTop = MeasureNestedTop(probe, new DocxParagraphElement(spaced), new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null), new DocxTableElement(probe));
        TestAssert.True(
            Math.Abs(spacedBreakTop - spacedEmptyTop) < 1e-9d,
            $"The nested table after an in-cell break should flow spacing like an empty line. breakTop={spacedBreakTop} emptyTop={spacedEmptyTop}.");
    }

    public static void DocxTableCellBreakSpillStaysLeftAlignedInCenteredCells()
    {
        // RV06 cellbreak probe (edge-t2, Word 16.0): Office left-aligns the spill
        // space even in centered cells; the spill uses the continuation text offset.
        // Characterization lock: centering the spill would fail this.
        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d) with { Alignment = DocxTextAlignment.Center };
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 10d, 10d) with { Alignment = DocxTextAlignment.Center };
        var cell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(before),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(after)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout laidCell = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages.Single().Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
        DocxTextLineLayout[] cellLines = laidCell.TextLines.ToArray();
        TestAssert.Equal(3, cellLines.Length);
        TestAssert.True(cellLines[0].X > laidCell.X, "The centered line should start past the cell edge.");
        TestAssert.True(
            cellLines[1].X < cellLines[0].X,
            $"The spill space should stay left-aligned in centered cells. spillX={cellLines[1].X} centeredX={cellLines[0].X}.");
    }

    public static void DocxTableSplitFragmentsFitWholeLinesOnly()
    {
        // RV06 pagination probe (edge-page-auto, Word 16.0): a split fragment keeps
        // only whole lines that fit its height (floor capacity); a line whose box
        // crosses the fragment bottom belongs to the continuation, even when its
        // baseline clears the edge. Pre-fix the packer kept baseline-clearing lines
        // and overfilled by one line on fractional remainders.
        DocxParagraph[] headParas = Enumerable.Range(1, 2)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Head " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        DocxParagraph[] tallParas = Enumerable.Range(1, 10)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Tall line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        var rows = new List<DocxTableRow>();
        rows.Add(new DocxTableRow([new DocxTableCell(string.Empty, [headParas[0]], null, null, null, null, [], DocxTableCellMargins.Empty)], 130.5d));
        rows.Add(new DocxTableRow([new DocxTableCell(string.Empty, [headParas[1]], null, null, null, null, [], DocxTableCellMargins.Empty)], null));
        rows.Add(new DocxTableRow([new DocxTableCell(string.Empty, tallParas, null, null, null, null, [], DocxTableCellMargins.Empty)], null));
        DocxTable table = new(null, [90d], rows);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        TestAssert.Equal(2, layout.Pages.Count);
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();
        TestAssert.Equal(3, firstPageRows.Length);
        TestAssert.Equal(3, firstPageRows[2].Cells[0].TextLines.Count);
        TestAssert.Equal(7, secondPageRows[0].Cells[0].TextLines.Count);
        TestAssert.Equal(10, firstPageRows[2].Cells[0].TextLines.Count + secondPageRows[0].Cells[0].TextLines.Count);
    }

    public static void DocxTableLayoutStageKeepsNestedTablesOnAuthoredSideOfPageBreak()
    {
        DocxTable beforeNestedTable = DocxTests.CreateSingleCellTable("Before", 12d);
        DocxTable afterNestedTable = DocxTests.CreateSingleCellTable("After", 12d);
        var cell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxTableElement(beforeNestedTable),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxTableElement(afterNestedTable)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table, beforeNestedTable, afterNestedTable]);

        DocxTableRowLayout[] rowFragments = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .ToArray();

        // RV06 cellbreak probe: run breaks flow through without fragmenting rows.
        TestAssert.Equal(1, rowFragments.Length);
        DocxTableRowLayout[] nestedRows = rowFragments[0].Cells.Single().NestedRows.ToArray();
        TestAssert.Equal(2, nestedRows.Length);
        TestAssert.Equal(0, nestedRows[0].Table.TableIndex);
        TestAssert.Equal(1, nestedRows[1].Table.TableIndex);
        TestAssert.Equal("Before", nestedRows[0].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("After", nestedRows[1].Cells.Single().TextLines.Single().Text);
    }

    public static void DocxTableLayoutStageKeepsTableCellColumnBreakInline()
    {
        DocxParagraph before = DocxTests.CreateDocxLayoutParagraph("Left", 10d, 10d);
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("Right", 10d, 10d);
        var cell = new DocxTableCell(string.Empty, [before, after], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(before),
                new DocxManualBreakElement(DocxBreakSourceKind.RunBreak, "column", null),
                new DocxParagraphElement(after)
            ]
        };
        DocxTable table = new(null, [90d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .Single()
            .Cells
            .Single();

        TestAssert.Equal(1, cellLayout.TextLines.Count);
        TestAssert.Equal("LeftRight", cellLayout.TextLines.Single().Text);
    }

    public static void DocxTableLayoutStagePartitionsNestedTablesAcrossCompetingCellPageBreaks()
    {
        DocxParagraph earlyBefore = DocxTests.CreateDocxLayoutParagraph("EarlyBefore", 8d, 12d);
        DocxParagraph earlyAfter = DocxTests.CreateDocxLayoutParagraph("EarlyAfter", 8d, 12d);
        DocxTable beforeNestedTable = DocxTests.CreateSingleCellTable("Before", 12d);
        DocxTable middleNestedTable = DocxTests.CreateSingleCellTable("Middle", 12d);
        DocxTable afterNestedTable = DocxTests.CreateSingleCellTable("After", 10d);
        var earlyCell = new DocxTableCell(string.Empty, [earlyBefore, earlyAfter], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(earlyBefore),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(earlyAfter)
            ]
        };
        var nestedCell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxTableElement(beforeNestedTable),
                new DocxTableElement(middleNestedTable),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxTableElement(afterNestedTable)
            ]
        };
        DocxTable table = new(null, [60d, 90d], [new DocxTableRow([earlyCell, nestedCell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table, beforeNestedTable, middleNestedTable, afterNestedTable]);

        DocxTableRowLayout[] rowFragments = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .ToArray();

        // RV06 cellbreak probe: run breaks flow through without fragmenting rows.
        TestAssert.Equal(1, rowFragments.Length);
        TestAssert.Equal("EarlyBefore| |EarlyAfter", string.Join("|", rowFragments[0].Cells[0].TextLines.Select(line => line.Text)));
        DocxTableRowLayout[] nestedRows = rowFragments[0].Cells[1].NestedRows.ToArray();
        TestAssert.Equal(3, nestedRows.Length);
        TestAssert.Equal("Before", nestedRows[0].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("Middle", nestedRows[1].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("After", nestedRows[2].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal(0, nestedRows[0].Table.TableIndex);
        TestAssert.Equal(1, nestedRows[1].Table.TableIndex);
        TestAssert.Equal(2, nestedRows[2].Table.TableIndex);
    }

    public static void DocxTableLayoutStageSplitsNestedTableRowsAcrossCompetingCellPageBreak()
    {
        DocxParagraph earlyBefore = DocxTests.CreateDocxLayoutParagraph("EarlyBefore", 10d, 20d);
        DocxParagraph earlyAfter = DocxTests.CreateDocxLayoutParagraph("EarlyAfter", 10d, 10d);
        DocxTableRow[] nestedRows = Enumerable.Range(1, 6)
            .Select(index => new DocxTableRow(
                [new DocxTableCell("Nested" + index.ToString(CultureInfo.InvariantCulture), [DocxTests.CreateDocxLayoutParagraph("Nested" + index.ToString(CultureInfo.InvariantCulture), 10d, 10d)], null, null, null, null, [], DocxTableCellMargins.Empty)],
                10d))
            .ToArray();
        DocxTable nestedTable = new(null, [60d], nestedRows);
        var earlyCell = new DocxTableCell(string.Empty, [earlyBefore, earlyAfter], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements =
            [
                new DocxParagraphElement(earlyBefore),
                new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", null),
                new DocxParagraphElement(earlyAfter)
            ]
        };
        var nestedCell = new DocxTableCell(string.Empty, [], null, null, null, null, [], DocxTableCellMargins.Empty)
        {
            BodyElements = [new DocxTableElement(nestedTable)]
        };
        DocxTable table = new(null, [60d, 90d], [new DocxTableRow([earlyCell, nestedCell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table, nestedTable]);

        DocxTableRowLayout[] rowFragments = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .ToArray();

        // RV06 cellbreak probe: run breaks flow through; nested rows split only by
        // natural overflow, never for the break. All content stays present in order.
        TestAssert.True(rowFragments.Length >= 1, "Flowing rows should still lay out across natural overflow.");
        TestAssert.True(rowFragments.All(fragment => fragment.FragmentReason != "CellPageBreak"), "No fragment should come from an in-cell run break.");
        TestAssert.Equal("EarlyBefore| |EarlyAfter", string.Join("|", rowFragments.SelectMany(fragment => fragment.Cells[0].TextLines).Select(line => line.Text)));
        string[] nestedTexts = rowFragments
            .SelectMany(fragment => fragment.Cells[1].NestedRows)
            .Select(row => row.Cells.Single().TextLines.Single().Text)
            .ToArray();
        TestAssert.Equal("Nested1|Nested2|Nested3|Nested4|Nested5|Nested6", string.Join("|", nestedTexts));
    }

    public static void DocxTableLayoutStageUsesCellMarginsForTextBox()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun("Inset", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var margins = new DocxTableCellMargins(3d, 8d, null, 12d, "60", "160", null, "240");
        var cell = new DocxTableCell("Inset", [paragraph], null, null, null, null, [], margins);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 30d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "Inset".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.Equal(cellLayout.X + 12d, cellLayout.TextLines[0].X);
        TestAssert.Equal(cellLayout.Y + cellLayout.Height - 11d * 0.94d - 3d, cellLayout.TextLines[0].BaselineY);
    }

    public static void DocxTableLayoutStageAppliesDefaultHorizontalCellMargin()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun("Flush", 12d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell("Flush", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 30d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "Flush".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        // Office A/B (w65 nil probe N): unset margins default to 0.48pt, so borderless text starts at grid + 0.48; explicit-0 restores the edge (w64).
        TestAssert.Equal(cellLayout.X + 0.48d, cellLayout.TextLines[0].X);
        TestAssert.Equal(cellLayout.Y + cellLayout.Height - 12d * 0.94d, cellLayout.TextLines[0].BaselineY);
    }

    public static void DocxTableLayoutStageStartsTextInsideVisibleCellBorder()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun("Bordered", 12d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var borders = new[]
        {
            new DocxTableCellBorder("left", "single", "auto", "4"),
            new DocxTableCellBorder("right", "single", "auto", "4")
        };
        var cell = new DocxTableCell("Bordered", [paragraph], null, null, null, null, borders, DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], 30d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "Bordered".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        // Office A/B (w63 sz4 probe): unset margins default to 0.48pt which dominates the 0.24pt border half, so text starts at grid + 0.48.
        TestAssert.Equal(cellLayout.X + 0.48d, cellLayout.TextLines[0].X);
    }

    public static void DocxTableLayoutStageDoesNotInventDefaultRowMinimumForAutoRows()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("A", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell("A", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10d, row.Height);
    }

    public static void DocxTableLayoutStageQuantizesAutoLineHeightToTwips()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("Alpha Beta", 9d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell("Alpha Beta", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [30d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTextLineLayout[] lines = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FractionalLineHeightTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single()
            .TextLines
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.Equal(12d, Math.Round(lines[0].BaselineY - lines[1].BaselineY, 6));
    }

    public static void DocxTableLayoutStageIncludesCollapsedHorizontalBorderAdvanceForContentRows()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("A", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var borders = new[]
        {
            new DocxTableCellBorder("bottom", "single", "auto", "4")
        };
        var cell = new DocxTableCell("A", [paragraph], null, null, null, null, borders, DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        // Table terminus (w7 doc-start probe): single-row bottom border hangs below
        // content, so height is content plus collapsed advance plus terminus width.
        TestAssert.Equal(10.96d, row.Height);
    }

    public static void DocxTableLayoutStageLetsTablePropertyExceptionRowsUseContentHeight()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("A", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell("A", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var rowExceptionMargins = new DocxTableCellMargins(0d, null, 0d, null, "0", null, "0", null);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], null) with {TablePropertyExceptionCellMargins = rowExceptionMargins }]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(10d, row.Height);
    }

    public static void DocxTableLayoutStageExpandsAtLeastRowsByAuthoredTopCellMargin()
    {
        var defaultCell = new DocxTableCell("A", [], null, null, null, null, [], DocxTableCellMargins.Empty);
        var topMarginCell = new DocxTableCell(
            "B",
            [],
            null,
            null,
            null,
            null,
            [],
            new DocxTableCellMargins(12d, null, null, null, "240", null, null, null));
        var table = new DocxTable(null, [40d, 40d], [new DocxTableRow([defaultCell, topMarginCell], 30d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(42d, row.Height);
    }

    public static void DocxTableLayoutStageIncludesParagraphBeforeSpacingInCellHeight()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("A", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            12d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell("A", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(22d, row.Height);
    }

    public static void DocxTableLayoutStageConsumesEmptyCellParagraphLineBox()
    {
        DocxParagraph first = DocxTests.CreateDocxLayoutParagraph("A", 10d, 10d);
        var empty = new DocxParagraph(
            [],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            10d,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        DocxParagraph second = DocxTests.CreateDocxLayoutParagraph("B", 10d, 10d);
        var cell = new DocxTableCell(string.Empty, [first, empty, second], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [80d], [new DocxTableRow([cell], null)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.Equal(30d, row.Height);
        TestAssert.Equal(2, row.Cells[0].TextLines.Count);
        TestAssert.Equal(20d, Math.Round(row.Cells[0].TextLines[0].BaselineY - row.Cells[0].TextLines[1].BaselineY, 3));
    }

    public static void DocxTableLayoutStageAppliesCellVerticalAlignment()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun("V", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var topCell = new DocxTableCell("V", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var centerCell = new DocxTableCell("V", [paragraph], null, null, null, "center", [], DocxTableCellMargins.Empty);
        var bottomCell = new DocxTableCell("V", [paragraph], null, null, null, "bottom", [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [40d, 40d, 40d], [new DocxTableRow([topCell, centerCell, bottomCell], 60d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "V".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        double topBaseline = row.Cells[0].TextLines[0].BaselineY;
        double centerBaseline = row.Cells[1].TextLines[0].BaselineY;
        double bottomBaseline = row.Cells[2].TextLines[0].BaselineY;
        TestAssert.True(centerBaseline < topBaseline, "Center-aligned cell text should move downward from the top baseline.");
        TestAssert.True(bottomBaseline < centerBaseline, "Bottom-aligned cell text should move below centered cell text.");
    }

    public static void DocxTableLayoutStageExpandsRowsToCellContent()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        var paragraph = new DocxParagraph(
            [new DocxTextRun("First Second", 11d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell("First Second", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [34d], [new DocxTableRow([cell], 10d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "First Second".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTableRowLayout row = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single();

        TestAssert.True(row.Height > 10d, "DOCX table rows should expand beyond a too-small declared height when cell text wraps.");
        TestAssert.True(row.Cells[0].TextLines.Count >= 2, "Expected the narrow cell to wrap content into multiple layout-owned text lines.");
    }

    public static void DocxTableLayoutStageHonorsCellNoWrap()
    {
        const string text = "First Second Third";
        var paragraph = new DocxParagraph(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell(text, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty) with {NoWrap = true };
        var table = new DocxTable(
            null,
            [34d],
            [new DocxTableRow([cell], 10d)]) with {PreferredWidthPoints = 34d };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        DocxTextLineLayout line = cellLayout.TextLines.Single();
        TestAssert.Equal(text, line.Text);
        TestAssert.True(line.Width > cellLayout.Width + 0.001d, "No-wrap cell text should remain on one overwide line instead of splitting to fit the cell frame.");
    }

    public static void DocxTableLayoutStageFitsCellTextToWidth()
    {
        const string text = "First Second Third";
        var paragraph = new DocxParagraph(
            [new DocxTextRun(text, 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell(text, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty) with {FitText = true };
        var table = new DocxTable(
            null,
            [34d],
            [new DocxTableRow([cell], 10d)]) with {PreferredWidthPoints = 34d };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        DocxTextLineLayout line = cellLayout.TextLines.Single();
        DocxTextSegmentLayout segment = line.Segments.Single();
        TestAssert.Equal(text, line.Text);
        // Office default insets (w68) narrow the text extents below the cell width; fit-text still targets the extents, now read off the record pads.
        TestAssert.True(Math.Abs(line.Width - (cellLayout.Width - cellLayout.ContentPaddingLeft - cellLayout.ContentPaddingRight)) < 0.001d, "Fit-text cell lines should target the cell text extents instead of keeping the natural overwide advance.");
        TestAssert.True(segment.PdfCharacterSpacing < 0d, "Fitting overwide text should reduce inter-character spacing.");
        TestAssert.Equal(DocxTextStateCharacterSpacingSource.AdvanceTarget, segment.PdfCharacterSpacingSource);
        TestAssert.True(!segment.CompensatePdfCharacterSpacing, "Fit-text advance targets should add a text-state spacing delta instead of compensating it away.");
    }

    public static void DocxTableLayoutStageSplitsOverwideCellTokensAtSafeCharacterBoundaries()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("ABCDEFGHIJ", 10d, null, false, false, false, null, null)],
            [],
            null,
            DocxTextAlignment.Left,
            null,
            0d,
            0d,
            1d,
            null,
            DocxParagraphSpacing.Empty,
            DocxParagraphKeepRules.Empty,
            null);
        var cell = new DocxTableCell("ABCDEFGHIJ", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(
            null,
            [16d],
            [new DocxTableRow([cell], 10d)]) with {PreferredWidthPoints = 16d };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.True(cellLayout.TextLines.Count >= 3, "Expected an overwide table-cell token to split into fitting line fragments.");
        TestAssert.Equal("ABC", cellLayout.TextLines[0].Text);
        TestAssert.Equal("DEF", cellLayout.TextLines[1].Text);
        TestAssert.Equal("GHI", cellLayout.TextLines[2].Text);
        TestAssert.True(cellLayout.TextLines[0].EndsWithIntraTokenBreak, "The layout model should preserve that this line ended inside an overwide token.");
        TestAssert.True(cellLayout.TextLines[1].EndsWithIntraTokenBreak, "Subsequent split fragments should keep their intra-token break reason.");
        TestAssert.True(cellLayout.TextLines[2].EndsWithIntraTokenBreak, "Every non-final split fragment should keep its intra-token break reason.");
        TestAssert.True(!cellLayout.TextLines.Last().EndsWithIntraTokenBreak, "The final token fragment should not be marked as a line-end intra-token break.");
        TestAssert.True(cellLayout.TextLines.All(line => line.Width <= cellLayout.Width + 0.001d), "Split token fragments should stay inside the cell frame.");
    }
}
