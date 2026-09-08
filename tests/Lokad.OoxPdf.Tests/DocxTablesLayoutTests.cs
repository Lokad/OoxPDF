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

        TestAssert.Equal(28d, row.Cells[0].X);
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
            return;
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "Alpha BG".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

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
            return;
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
            return;
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
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] thirdPageRows = layout.Pages[2].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(3, layout.Pages.Count);
        TestAssert.Equal(3, rowFragments.Length);
        TestAssert.Equal(1, firstPageRows.Length);
        TestAssert.Equal(1, secondPageRows.Length);
        TestAssert.Equal(1, thirdPageRows.Length);
        TestAssert.Equal(0, firstPageRows[0].RowIndex);
        TestAssert.Equal(1, secondPageRows[0].RowIndex);
        TestAssert.Equal(0, secondPageRows[0].FragmentIndex);
        TestAssert.Equal("CellPageBreak", secondPageRows[0].FragmentReason);
        TestAssert.True(secondPageRows[0].Y - secondPageRows[0].Height >= 10d, "The first cell-page-break fragment must fit inside the new page content frame.");
        TestAssert.Equal(1, thirdPageRows[0].RowIndex);
        TestAssert.Equal(1, thirdPageRows[0].FragmentIndex);
        TestAssert.Equal("CellPageBreak", thirdPageRows[0].FragmentReason);
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

        TestAssert.Equal(3, layout.Pages.Count);
        TestAssert.Equal(3, rowFragments.Length);
        for (int fragmentIndex = 0; fragmentIndex < rowFragments.Length; fragmentIndex++)
        {
            TestAssert.Equal(fragmentIndex, rowFragments[fragmentIndex].FragmentIndex);
            TestAssert.Equal(3, rowFragments[fragmentIndex].FragmentCount);
            TestAssert.Equal("CellPageBreak", rowFragments[fragmentIndex].FragmentReason);
        }

        TestAssert.Equal(80d, rowFragments[0].Height);
        TestAssert.Equal(20d, rowFragments[1].Height);
        TestAssert.Equal(10d, rowFragments[2].Height);
        TestAssert.Equal(8, rowFragments[0].Cells.Single().TextLines.Count);
        TestAssert.Equal("Before1|Before2|Before3|Before4|Before5|Before6|Before7|Before8", string.Join("|", rowFragments[0].Cells.Single().TextLines.Select(line => line.Text)));
        TestAssert.Equal("Before9|Before10", string.Join("|", rowFragments[1].Cells.Single().TextLines.Select(line => line.Text)));
        TestAssert.Equal("After", rowFragments[2].Cells.Single().TextLines.Single().Text);

        DocxTableRowSnapshot[] rowSnapshots = DocxLayoutSnapshot.FromLayout(layout).Pages
            .SelectMany(page => page.TableRows)
            .OrderBy(row => row.FragmentIndex)
            .ToArray();
        TestAssert.Equal(0d, rowSnapshots[0].FragmentOffsetFromRowTop);
        TestAssert.Equal(80d, rowSnapshots[1].FragmentOffsetFromRowTop);
        TestAssert.Equal(100d, rowSnapshots[2].FragmentOffsetFromRowTop);
        TestAssert.Equal(110d, rowSnapshots[0].FullRowHeight);
        TestAssert.Equal(110d, rowSnapshots[1].FullRowHeight);
        TestAssert.Equal(110d, rowSnapshots[2].FullRowHeight);
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

        TestAssert.Equal(3, rowFragments.Length);
        TestAssert.Equal("EarlyBefore", rowFragments[0].Cells[0].TextLines.Single().Text);
        TestAssert.Equal("EarlyAfter", rowFragments[1].Cells[0].TextLines.Single().Text);
        TestAssert.Equal("LaterFirst", rowFragments[0].Cells[1].TextLines.Single().Text);
        TestAssert.Equal("LaterMiddle", rowFragments[1].Cells[1].TextLines.Single().Text);
        TestAssert.Equal(0, rowFragments[2].Cells[0].TextLines.Count);
        TestAssert.Equal("LaterAfter", rowFragments[2].Cells[1].TextLines.Single().Text);
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

        TestAssert.Equal(2, rowFragments.Length);
        DocxInlineImageLayout firstImage = rowFragments[0].Cells.Single().InlineImages.Single();
        DocxInlineImageLayout secondImage = rowFragments[1].Cells.Single().InlineImages.Single();
        TestAssert.Equal(0, firstImage.SourceParagraphIndex ?? -1);
        TestAssert.Equal(1, secondImage.SourceParagraphIndex ?? -1);
        TestAssert.True(firstImage.Image == beforeImage, "The first row fragment should keep the image before the authored page break.");
        TestAssert.True(secondImage.Image == afterImage, "The second row fragment should keep the image after the authored page break.");
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

        TestAssert.Equal(2, rowFragments.Length);
        DocxTableRowLayout firstNestedRow = rowFragments[0].Cells.Single().NestedRows.Single();
        DocxTableRowLayout secondNestedRow = rowFragments[1].Cells.Single().NestedRows.Single();
        TestAssert.Equal(0, firstNestedRow.Table.TableIndex);
        TestAssert.Equal(1, secondNestedRow.Table.TableIndex);
        TestAssert.Equal("Before", firstNestedRow.Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("After", secondNestedRow.Cells.Single().TextLines.Single().Text);
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

        TestAssert.Equal(3, rowFragments.Length);
        TestAssert.Equal("Before", rowFragments[0].Cells[1].NestedRows.Single().Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("Middle", rowFragments[1].Cells[1].NestedRows.Single().Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("After", rowFragments[2].Cells[1].NestedRows.Single().Cells.Single().TextLines.Single().Text);
        TestAssert.Equal(0, rowFragments[0].Cells[1].NestedRows.Single().Table.TableIndex);
        TestAssert.Equal(1, rowFragments[1].Cells[1].NestedRows.Single().Table.TableIndex);
        TestAssert.Equal(2, rowFragments[2].Cells[1].NestedRows.Single().Table.TableIndex);
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

        TestAssert.Equal(2, rowFragments.Length);
        TestAssert.Equal(0, rowFragments[0].FragmentIndex);
        TestAssert.Equal(1, rowFragments[1].FragmentIndex);
        TestAssert.Equal("CellPageBreak", rowFragments[0].FragmentReason);
        TestAssert.Equal("CellPageBreak", rowFragments[1].FragmentReason);

        DocxTableRowLayout[] firstFragmentNestedRows = rowFragments[0].Cells[1].NestedRows.ToArray();
        DocxTableRowLayout[] secondFragmentNestedRows = rowFragments[1].Cells[1].NestedRows.ToArray();
        TestAssert.Equal(2, firstFragmentNestedRows.Length);
        TestAssert.Equal(4, secondFragmentNestedRows.Length);
        TestAssert.Equal("Nested1", firstFragmentNestedRows[0].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("Nested2", firstFragmentNestedRows[1].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("Nested3", secondFragmentNestedRows[0].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal("Nested6", secondFragmentNestedRows[3].Cells.Single().TextLines.Single().Text);
        TestAssert.Equal(0, firstFragmentNestedRows[0].Table.TableIndex);
        TestAssert.Equal(0, secondFragmentNestedRows[0].Table.TableIndex);
    }

    public static void DocxTableLayoutStageUsesCellMarginsForTextBox()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "Inset".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

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

    public static void DocxTableLayoutStageDoesNotInventHorizontalCellPadding()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "Flush".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.Equal(cellLayout.X, cellLayout.TextLines[0].X);
        TestAssert.Equal(cellLayout.Y + cellLayout.Height - 12d * 0.94d, cellLayout.TextLines[0].BaselineY);
    }

    public static void DocxTableLayoutStageStartsTextInsideVisibleCellBorder()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "Bordered".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

        DocxTableCellLayout cellLayout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        TestAssert.Equal(cellLayout.X + 0.24d, cellLayout.TextLines[0].X);
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

        TestAssert.Equal(10.48d, row.Height);
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
            return;
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "V".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

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
            return;
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
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(arial), "First Second".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);

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
        TestAssert.True(Math.Abs(line.Width - cellLayout.Width) < 0.001d, "Fit-text cell lines should target the cell text extents instead of keeping the natural overwide advance.");
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
