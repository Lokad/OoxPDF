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

internal static class DocxTableCellsTests
{
    public static void DocxTableUniformThreeRowTopReusesQualifiedSlotsMasksAndFlow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (int columns in new[] { 1, 2 })
        foreach (string? rule in new string?[] { null, "auto", "exact", "atLeast" })
        foreach (double minimum in new[] { 12d, 14d, 24d })
        foreach ((double top, double bottom) in new[] { (0d, 0d), (6d, 12d), (12d, 6d) })
        {
            bool automatic = rule is null or "auto";
            var margins = DocxTableCellMargins.Empty with { TopPoints = top, BottomPoints = bottom };
            DocxLayout Layout(int count) => CreateParagraphMinimumLayout("min-single", full,
                rowHeight: rule is null ? null : rule == "auto" ? 320d : 180d, rowRule: rule,
                verticalAlignment: "top", cellMargins: margins, rowCount: count, columnCount: columns,
                minimumSpacing: minimum, borderSize: "8");
            var profile = Layout(2).Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            var actual = Layout(3);
            var rows = actual.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            TestAssert.Equal(3, rows.Length);
            for (int i = 0; i < rows.Length; i++)
            {
                var expected = profile[i == 2 ? 1 : 0]; var row = rows[i];
                TestAssert.True(Math.Abs(row.Height - expected.Height) < .000001d,
                    $"Uniform three-row advance retains its qualified profile: columns={columns}, rule={rule}, minimum={minimum}, margins={top}/{bottom}, row={i}, expected={expected.Height}, actual={row.Height}.");
                TestAssert.True(!row.Table.AutomaticMinimumRowPlacement,
                    "Top masks preserve the existing paragraph baseline slots.");
                for (int ci = 0; ci < columns; ci++)
                {
                    var cell = row.Cells[ci]; var before = expected.Cells[ci];
                    TestAssert.True(Math.Abs((cell.TextLines.Single().BaselineY - cell.Y - cell.Height) -
                        (before.TextLines.Single().BaselineY - before.Y - before.Height)) < .000001d,
                        "Every row retains its baseline relative to its slot top.");
                    var clip = DocxRenderer.ResolveTableCellContentClip(row, cell);
                    double terminal = i == 2 ? .96d : 0d;
                    TestAssert.True(clip.Admitted && Math.Abs(clip.Y - cell.Y - bottom - terminal) < .000001d &&
                        Math.Abs(clip.Height - cell.Height + top + .96d + bottom + terminal) < .000001d,
                        "Interior and terminal masks match the independent Word model.");
                    TestAssert.Equal(automatic ? 0d : bottom, row.Table.AlignedMinimumRowBottomMargins[i]);
                }
            }
            var two = Layout(2);
            double followingAdvance = two.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY -
                actual.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
            TestAssert.True(Math.Abs(followingAdvance - rows[1].Height) < .000001d,
                "Adding an interior row advances following body text exactly once.");
        }
    }

    public static void DocxTableUniformThreeRowTopRetainsExcludedFallbacks()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (int columns in new[] { 1, 2 })
        foreach (string? rule in new string?[] { null, "auto", "exact", "atLeast" })
        foreach (string variant in new[] { "header", "rich", "four-rows", "three-columns", "wide-border", "frame", "metrics", "default-rule", "unknown-rule", "nonpositive", "alignment" })
        {
            string? actualRule = variant == "default-rule" ? null : variant == "unknown-rule" ? "other" : rule;
            double? height = variant == "default-rule" ? 180d : variant == "nonpositive" ? 0d :
                actualRule is null ? null : actualRule == "auto" ? 320d : 180d;
            var layout = CreateParagraphMinimumLayout("min-single", variant == "metrics" ? new MinimumMetricsFallback(full) : full,
                variant is "header" or "rich" ? variant : "", rowHeight: height, rowRule: actualRule,
                verticalAlignment: variant == "alignment" ? "other" : "top", rowCount: variant == "four-rows" ? 4 : 3,
                columnCount: variant == "three-columns" ? 3 : columns, pageHeight: variant == "frame" ? 100d : 2400d,
                borderSize: variant == "wide-border" ? "32" : "8");
            // An explicitly automatic zero height still has natural row semantics.
            if (variant == "nonpositive" && rule is null or "auto") continue;
            foreach (var row in layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()))
            foreach (var cell in row.Cells)
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                    $"Excluded uniform three-row layout retains fallback: variant={variant}, rule={rule}.");
        }
    }

    public static void DocxTableUniformThreeRowTopContentDominanceRemainsPerRow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (int columns in new[] { 1, 2 })
        foreach (int dominatedIndex in Enumerable.Range(0, 3))
        {
            var heights = Enumerable.Range(0, 3).Select(i => ((double?)(i == dominatedIndex ? 10d : 180d), (string?)"atLeast")).ToArray();
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var layout = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top", cellMargins: margins,
                rowCount: 3, columnCount: columns, borderSize: "8", rowHeightRules: heights);
            var rows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            TestAssert.Equal(3, rows.Length);
            foreach (var row in rows)
            {
                bool admitted = row.RowIndex != dominatedIndex;
                TestAssert.Equal(admitted ? 12d : 0d, row.Table.AlignedMinimumRowBottomMargins[row.RowIndex]);
                foreach (var cell in row.Cells)
                    TestAssert.True(DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted == admitted,
                        "A content-dominated minimum row retains fallback independently of adjacent declared rows.");
            }
        }
    }

    public static void DocxTableThreeRowMixedTopReusesQualifiedSlotsMasksAndFlow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (int columns in new[] { 1, 2 })
        foreach (int automaticMask in Enumerable.Range(1, 6))
        foreach (string? automaticRule in new string?[] { null, "auto" })
        foreach (string declaredRule in new[] { "exact", "atLeast", "alternating" })
        foreach (double minimum in new[] { 12d, 14d, 24d })
        foreach ((double top, double bottom) in new[] { (0d, 0d), (6d, 12d), (12d, 6d) })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = top, BottomPoints = bottom };
            var heights = Enumerable.Range(0, 3).Select(i => (automaticMask & (1 << i)) != 0 ?
                ((double?)(automaticRule is null ? null : 320d), automaticRule) :
                ((double?)180d, declaredRule == "alternating" ? i % 2 == 0 ? "exact" : "atLeast" : declaredRule)).ToArray();
            DocxLayout Profile(string? rule) => CreateParagraphMinimumLayout("min-single", full,
                rowHeight: rule is null ? null : 180d, rowRule: rule, verticalAlignment: "top", cellMargins: margins,
                rowCount: 2, columnCount: columns, minimumSpacing: minimum, borderSize: "8");
            var automatic = Profile(null).Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            var exact = Profile("exact").Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            var atLeast = Profile("atLeast").Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            var mixed = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top", cellMargins: margins,
                rowCount: 3, columnCount: columns, minimumSpacing: minimum, borderSize: "8",
                rowHeightRules: heights);
            var rows = mixed.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            TestAssert.Equal(3, rows.Length);
            for (int i = 0; i < rows.Length; i++)
            {
                bool isAutomatic = (automaticMask & (1 << i)) != 0;
                var expected = (isAutomatic ? automatic : heights[i].Item2 == "exact" ? exact : atLeast)[i == 2 ? 1 : 0];
                var row = rows[i];
                TestAssert.True(Math.Abs(row.Height - expected.Height) < .000001d,
                    $"A third mixed row uses the independently qualified advance: columns={columns}, mask={automaticMask}, auto={automaticRule}, declared={declaredRule}, minimum={minimum}, margins={top}/{bottom}, row={i}, expected={expected.Height}, actual={row.Height}.");
                TestAssert.True(!row.Table.AutomaticMinimumRowPlacement,
                    "Top composition preserves its existing first-baseline slots.");
                for (int ci = 0; ci < columns; ci++)
                {
                    var cell = row.Cells[ci]; var before = expected.Cells[ci];
                    TestAssert.True(Math.Abs((cell.TextLines.Single().BaselineY - cell.Y - cell.Height) -
                        (before.TextLines.Single().BaselineY - before.Y - before.Height)) < .000001d,
                        "Each paragraph keeps its qualified baseline relative to the top of its slot.");
                    var clip = DocxRenderer.ResolveTableCellContentClip(row, cell);
                    double terminal = i == 2 ? .96d : 0d;
                    TestAssert.True(clip.Admitted && Math.Abs(clip.Y - cell.Y - bottom - terminal) < .000001d &&
                        Math.Abs(clip.Height - cell.Height + top + .96d + bottom + terminal) < .000001d,
                        "Interior and terminal rows retain their source-measured masks.");
                }
            }
            var allAutomatic = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top", cellMargins: margins,
                rowCount: 3, columnCount: columns, minimumSpacing: minimum, borderSize: "8");
            double rowDelta = rows.Sum(r => r.Height) - allAutomatic.Pages[0].Items.OfType<DocxTableRowLayout>().Sum(r => r.Height);
            double followingDelta = allAutomatic.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY -
                mixed.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
            TestAssert.True(Math.Abs(rowDelta - followingDelta) < .000001d,
                "Following body flow includes the composed advances exactly once.");
        }
    }

    public static void DocxTableThreeRowMixedTopRetainsExcludedFallbacks()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (int columns in new[] { 1, 2 })
        foreach (string variant in new[] { "header", "rich", "four-rows", "three-columns", "wide-border", "frame", "metrics", "default-rule", "unknown-rule", "nonpositive" })
        {
            int count = variant == "four-rows" ? 4 : 3;
            var heights = Enumerable.Range(0, count).Select(i => i == 1 ?
                ((double?)(variant == "nonpositive" ? 0d : 180d), variant == "default-rule" ? null : variant == "unknown-rule" ? "other" : "exact") :
                ((double?)null, (string?)null)).ToArray();
            var layout = CreateParagraphMinimumLayout("min-single", variant == "metrics" ? new MinimumMetricsFallback(full) : full,
                variant is "header" or "rich" ? variant : "", verticalAlignment: "top", rowCount: count,
                columnCount: variant == "three-columns" ? 3 : columns, pageHeight: variant == "frame" ? 160d : 2400d,
                rowHeightRules: heights, borderSize: variant == "wide-border" ? "32" : null);
            foreach (var row in layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()))
            foreach (var cell in row.Cells)
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                    "Unqualified three-row compositions retain fallback: " + variant);
        }
    }

    public static void DocxTableThreeRowMixedTopContentDominanceRemainsPerRow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (int columns in new[] { 1, 2 })
        foreach (int automaticIndex in Enumerable.Range(0, 3))
        foreach (int exactIndex in Enumerable.Range(0, 3).Where(i => i != automaticIndex))
        {
            var heights = Enumerable.Range(0, 3).Select(i => i == automaticIndex ? ((double?)null, (string?)null) :
                i == exactIndex ? ((double?)180d, (string?)"exact") : ((double?)10d, (string?)"atLeast")).ToArray();
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var layout = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top", cellMargins: margins,
                rowCount: 3, columnCount: columns, borderSize: "8", rowHeightRules: heights);
            var rows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            foreach (var row in rows)
            {
                bool admitted = row.RowIndex == automaticIndex || row.RowIndex == exactIndex;
                TestAssert.Equal(row.RowIndex == exactIndex ? 12d : 0d, row.Table.AlignedMinimumRowBottomMargins[row.RowIndex]);
                foreach (var cell in row.Cells)
                    TestAssert.True(DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted == admitted,
                        "A content-dominated minimum row cannot borrow another row's qualified mask or margin advance.");
            }
        }
    }

    public static void DocxTableOneColumnMixedTopRowsReuseQualifiedGeometryAndFollowingFlow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string? automaticRule in new string?[] { null, "auto" })
        foreach (string declaredRule in new[] { "exact", "atLeast" })
        foreach (int automaticIndex in new[] { 0, 1 })
        foreach (double minimum in new[] { 12d, 14d, 24d })
        foreach ((double top, double bottom) in new[] { (0d, 0d), (6d, 12d), (12d, 6d) })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = top, BottomPoints = bottom };
            var heights = new (double? Height, string? Rule)[2];
            heights[automaticIndex] = (automaticRule is null ? null : 320d, automaticRule);
            heights[1 - automaticIndex] = (180d, declaredRule);
            var automatic = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top",
                cellMargins: margins, rowCount: 2, columnCount: 1, minimumSpacing: minimum, borderSize: "8");
            var declared = CreateParagraphMinimumLayout("min-single", full, rowHeight: 180d, rowRule: declaredRule,
                verticalAlignment: "top", cellMargins: margins, rowCount: 2, columnCount: 1,
                minimumSpacing: minimum, borderSize: "8");
            var mixed = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top",
                cellMargins: margins, rowCount: 2, columnCount: 1, minimumSpacing: minimum,
                borderSize: "8", rowHeightRules: heights);
            var mixedRows = mixed.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            for (int i = 0; i < 2; i++)
            {
                var expected = (i == automaticIndex ? automatic : declared).Pages[0].Items.OfType<DocxTableRowLayout>().ElementAt(i);
                var actual = mixedRows[i];
                TestAssert.Equal(expected.Height, actual.Height);
                TestAssert.True(!actual.Table.AutomaticMinimumRowPlacement,
                    "Mixed top rows must retain their first-baseline slot placement.");
                for (int c = 0; c < 1; c++)
                {
                    var before = expected.Cells[c]; var after = actual.Cells[c];
                    TestAssert.True(Math.Abs((before.TextLines.Single().BaselineY - before.Y - before.Height) -
                        (after.TextLines.Single().BaselineY - after.Y - after.Height)) < .000001d,
                        "Each row keeps its accepted paragraph baseline relative to the top of its content slot.");
                    var clip = DocxRenderer.ResolveTableCellContentClip(actual, after);
                    double terminalBand = i == 1 ? .96d : 0d;
                    TestAssert.True(clip.Admitted && Math.Abs(clip.Y - after.Y - bottom - terminalBand) < .000001d &&
                        Math.Abs(clip.Height - after.Height + top + .96d + bottom + terminalBand) < .000001d,
                        "Both row kinds use the independently measured margin and border insets.");
                }
            }
            double extraHeight = mixedRows.Sum(r => r.Height) - automatic.Pages[0].Items.OfType<DocxTableRowLayout>().Sum(r => r.Height);
            double followingDelta = automatic.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY -
                mixed.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
            TestAssert.True(Math.Abs(extraHeight - followingDelta) < .000001d,
                "Following body flow includes exactly the admitted declared-row advance.");
        }
    }

    public static void DocxTableOneColumnMixedTopRetainsExcludedShapesRulesAndFrames()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string variant in new[] { "header", "rich", "three-columns", "four-rows", "frame", "metrics", "default-rule", "unknown-rule", "nonpositive" })
        {
            int count = variant == "four-rows" ? 4 : 2;
            var heights = Enumerable.Range(0, count).Select(i => i == 1 ?
                ((double?)(variant == "nonpositive" ? 0d : 180d), variant == "default-rule" ? null : variant == "unknown-rule" ? "other" : "exact") :
                ((double?)null, (string?)null)).ToArray();
            var layout = CreateParagraphMinimumLayout("min-single", variant == "metrics" ? new MinimumMetricsFallback(full) : full,
                variant is "header" or "rich" ? variant : "", verticalAlignment: "top", rowCount: count,
                columnCount: variant == "three-columns" ? 3 : 1,
                pageHeight: variant == "frame" ? 160d : 2400d, rowHeightRules: heights);
            foreach (var row in layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()))
            foreach (var cell in row.Cells)
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                    "Unqualified mixed top rows retain their masks: " + variant);
        }
    }

    public static void DocxTableOneColumnMixedTopContentDominanceIsAdmittedPerRow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (int automaticIndex in new[] { 0, 1 })
        {
            var heights = new (double? Height, string? Rule)[2];
            heights[automaticIndex] = (null, null); heights[1 - automaticIndex] = (10d, "atLeast");
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var layout = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top",
                cellMargins: margins, rowCount: 2, columnCount: 1, borderSize: "8", rowHeightRules: heights);
            var rows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            foreach (var cell in rows[automaticIndex].Cells)
                TestAssert.True(DocxRenderer.ResolveTableCellContentClip(rows[automaticIndex], cell).Admitted,
                    "The independently qualified automatic mask remains admitted beside a content-dominated minimum row.");
            var fallback = rows[1 - automaticIndex];
            TestAssert.Equal(0d, fallback.Table.AlignedMinimumRowBottomMargins[fallback.RowIndex]);
            foreach (var cell in fallback.Cells)
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(fallback, cell).Admitted,
                    "Content-dominated minimum rows retain their recorded margin and mask fallback.");
        }
    }

    public static void DocxTableMixedTopRowsReuseQualifiedGeometryAndFollowingFlow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string? automaticRule in new string?[] { null, "auto" })
        foreach (string declaredRule in new[] { "exact", "atLeast" })
        foreach (int automaticIndex in new[] { 0, 1 })
        foreach (double minimum in new[] { 12d, 14d, 24d })
        foreach ((double top, double bottom) in new[] { (0d, 0d), (6d, 12d), (12d, 6d) })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = top, BottomPoints = bottom };
            var heights = new (double? Height, string? Rule)[2];
            heights[automaticIndex] = (automaticRule is null ? null : 320d, automaticRule);
            heights[1 - automaticIndex] = (180d, declaredRule);
            var automatic = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top",
                cellMargins: margins, rowCount: 2, columnCount: 2, minimumSpacing: minimum, borderSize: "8");
            var declared = CreateParagraphMinimumLayout("min-single", full, rowHeight: 180d, rowRule: declaredRule,
                verticalAlignment: "top", cellMargins: margins, rowCount: 2, columnCount: 2,
                minimumSpacing: minimum, borderSize: "8");
            var mixed = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top",
                cellMargins: margins, rowCount: 2, columnCount: 2, minimumSpacing: minimum,
                borderSize: "8", rowHeightRules: heights);
            var mixedRows = mixed.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            for (int i = 0; i < 2; i++)
            {
                var expected = (i == automaticIndex ? automatic : declared).Pages[0].Items.OfType<DocxTableRowLayout>().ElementAt(i);
                var actual = mixedRows[i];
                TestAssert.Equal(expected.Height, actual.Height);
                TestAssert.True(!actual.Table.AutomaticMinimumRowPlacement,
                    "Mixed top rows must retain their first-baseline slot placement.");
                for (int c = 0; c < 2; c++)
                {
                    var before = expected.Cells[c]; var after = actual.Cells[c];
                    TestAssert.True(Math.Abs((before.TextLines.Single().BaselineY - before.Y - before.Height) -
                        (after.TextLines.Single().BaselineY - after.Y - after.Height)) < .000001d,
                        "Each row keeps its accepted paragraph baseline relative to the top of its content slot.");
                    var clip = DocxRenderer.ResolveTableCellContentClip(actual, after);
                    double terminalBand = i == 1 ? .96d : 0d;
                    TestAssert.True(clip.Admitted && Math.Abs(clip.Y - after.Y - bottom - terminalBand) < .000001d &&
                        Math.Abs(clip.Height - after.Height + top + .96d + bottom + terminalBand) < .000001d,
                        "Both row kinds use the independently measured margin and border insets.");
                }
            }
            double extraHeight = mixedRows.Sum(r => r.Height) - automatic.Pages[0].Items.OfType<DocxTableRowLayout>().Sum(r => r.Height);
            double followingDelta = automatic.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY -
                mixed.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
            TestAssert.True(Math.Abs(extraHeight - followingDelta) < .000001d,
                "Following body flow includes exactly the admitted declared-row advance.");
        }
    }

    public static void DocxTableMixedTopRetainsExcludedShapesRulesAndFrames()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string variant in new[] { "header", "rich", "four-columns", "three-columns", "four-rows", "frame", "metrics", "default-rule", "unknown-rule", "nonpositive" })
        {
            int count = variant == "four-rows" ? 4 : 2;
            var heights = Enumerable.Range(0, count).Select(i => i == 1 ?
                ((double?)(variant == "nonpositive" ? 0d : 180d), variant == "default-rule" ? null : variant == "unknown-rule" ? "other" : "exact") :
                ((double?)null, (string?)null)).ToArray();
            var layout = CreateParagraphMinimumLayout("min-single", variant == "metrics" ? new MinimumMetricsFallback(full) : full,
                variant is "header" or "rich" ? variant : "", verticalAlignment: "top", rowCount: count,
                columnCount: variant == "four-columns" ? 4 : variant == "three-columns" ? 3 : 2,
                pageHeight: variant == "frame" ? 160d : 2400d, rowHeightRules: heights);
            foreach (var row in layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()))
            foreach (var cell in row.Cells)
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                    "Unqualified mixed top rows retain their masks: " + variant);
        }
    }

    public static void DocxTableMixedTopContentDominanceIsAdmittedPerRow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (int automaticIndex in new[] { 0, 1 })
        {
            var heights = new (double? Height, string? Rule)[2];
            heights[automaticIndex] = (null, null); heights[1 - automaticIndex] = (10d, "atLeast");
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var layout = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top",
                cellMargins: margins, rowCount: 2, columnCount: 2, borderSize: "8", rowHeightRules: heights);
            var rows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            foreach (var cell in rows[automaticIndex].Cells)
                TestAssert.True(DocxRenderer.ResolveTableCellContentClip(rows[automaticIndex], cell).Admitted,
                    "The independently qualified automatic mask remains admitted beside a content-dominated minimum row.");
            var fallback = rows[1 - automaticIndex];
            TestAssert.Equal(0d, fallback.Table.AlignedMinimumRowBottomMargins[fallback.RowIndex]);
            foreach (var cell in fallback.Cells)
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(fallback, cell).Admitted,
                    "Content-dominated minimum rows retain their recorded margin and mask fallback.");
        }
    }

    public static void DocxTableDeclaredTopMarginAdvanceAndMasksMatchWord()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach ((int rows, int columns) in new[] { (1, 1), (1, 2), (2, 1), (2, 2) })
        foreach (string rule in new[] { "exact", "atLeast" })
        foreach (double minimum in new[] { 12d, 14d, 24d })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var layout = CreateParagraphMinimumLayout("min-single", full, rowHeight: 180d,
                rowRule: rule, verticalAlignment: "top", cellMargins: margins, rowCount: rows,
                columnCount: columns, minimumSpacing: minimum, borderSize: "8");
            foreach (var row in layout.Pages[0].Items.OfType<DocxTableRowLayout>())
            {
                double terminalBand = row.RowIndex == rows - 1 ? .96d : 0d;
                double expectedAdvance = 180d + 12d + terminalBand + (rule == "atLeast" ? 6.96d : 0d);
                TestAssert.True(Math.Abs(row.Height - expectedAdvance) < .000001d,
                    "Word advances declared top rows by the bottom margin outside the paragraph slot.");
                foreach (var cell in row.Cells)
                {
                    var clip = DocxRenderer.ResolveTableCellContentClip(row, cell);
                    TestAssert.True(clip.Admitted &&
                        Math.Abs(clip.X - cell.X - .48d) < .000001d &&
                        Math.Abs(clip.Y - cell.Y - 12d - terminalBand) < .000001d &&
                        Math.Abs(clip.Width - cell.Width + .96d) < .000001d &&
                        Math.Abs(clip.Height - cell.Height + 6.96d + 12d + terminalBand) < .000001d,
                        "The inner mask excludes the new margin advance, top inset and terminal border.");
                    TestAssert.True(!row.Table.AutomaticMinimumRowPlacement,
                        "Declared top rows retain their paragraph-slot placement.");
                }
            }
        }
    }

    public static void DocxTableDeclaredTopRetainsHeightFrameAndComplexFallbacks()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string variant in new[] { "mixed-auto-three-columns", "default-rule", "header", "rich", "frame", "metrics", "four-rows", "three-columns", "content-dominated" })
        {
            IDocxTextMeasurer measurer = variant == "metrics" ? new MinimumMetricsFallback(full) : full;
            var heights = variant == "mixed-auto-three-columns" ? new (double? Height, string? Rule)[] { (180d, "exact"), (320d, "auto") } : null;
            var layout = CreateParagraphMinimumLayout(variant == "content-dominated" ? "min,min,min" : "min-single", measurer,
                variant is "header" or "rich" ? variant : "", rowHeight: variant == "content-dominated" ? 10d : 180d,
                rowRule: variant == "default-rule" ? null : variant == "content-dominated" ? "atLeast" : "exact",
                verticalAlignment: "top", rowCount: variant == "four-rows" ? 4 : 2,
                columnCount: variant is "three-columns" or "mixed-auto-three-columns" ? 3 : 2, rowHeightRules: heights,
                pageHeight: variant == "frame" ? 160d : 2400d);
            foreach (var row in layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()))
            foreach (var cell in row.Cells)
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                    "Unqualified declared top rows retain their content masks: " + variant);
        }
    }

    public static void DocxTableDeclaredTopCompoundBorderRetainsOriginalMaskExtent()
    {
        DocxDocument source = ReadMarginClipFixture("center");
        DocxTable table = source.BodyElements.OfType<DocxTableElement>().Single().Table;
        table = table with { Rows = table.Rows.Select(row => row with {
            Cells = row.Cells.Select(cell => cell with {
                VerticalAlignmentValue = "top",
                Borders = cell.Borders.Select(border => border with { Value = "triple" }).ToArray()
            }).ToArray()
        }).ToArray() };
        DocxDocument document = source with {
            BodyElements = source.BodyElements.Select(e => e is DocxTableElement ? (DocxBodyElement)new DocxTableElement(table) : e).ToArray(),
            FallbackTables = [table]
        };
        var renderer = MarginClipRenderer();
        var snapshot = renderer.InspectLayout(document);
        TestAssert.True(Math.Abs(snapshot.Pages[0].TableRows.Single().Height - 180.96d) < .001d,
            "Unqualified compound borders must retain the original full-cell extent, without an added margin that could expose clipped glyphs.");
    }

    public static void DocxTableTopAutoClipMatchesIndependentWordInsets()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach ((int rows, int columns) in new[] { (1, 1), (1, 2), (2, 1), (2, 2) })
        foreach (string? rule in new string?[] { null, "auto" })
        foreach (double minimum in new[] { 12d, 14d, 24d })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var layout = CreateParagraphMinimumLayout("min-single", full, rowHeight: rule is null ? null : 320d,
                rowRule: rule, verticalAlignment: "top", cellMargins: margins, rowCount: rows,
                columnCount: columns, minimumSpacing: minimum, borderSize: "8");
            foreach (var row in layout.Pages[0].Items.OfType<DocxTableRowLayout>())
            foreach (var cell in row.Cells)
            {
                var clip = DocxRenderer.ResolveTableCellContentClip(row, cell);
                double bottom = 12d + (row.RowIndex == rows - 1 ? .96d : 0d);
                TestAssert.True(clip.Admitted, "Qualified top-aligned automatic rows expose their inner content masks.");
                TestAssert.True(Math.Abs(clip.X - cell.X - .48d) < .000001d &&
                    Math.Abs(clip.Y - cell.Y - bottom) < .000001d &&
                    Math.Abs(clip.Width - cell.Width + .96d) < .000001d &&
                    Math.Abs(clip.Height - cell.Height + 6.96d + bottom) < .000001d,
                    "Top masks use the independently measured single band, margins and terminal bottom border.");
                TestAssert.True(!row.Table.AutomaticMinimumRowPlacement,
                    "Content-mask admission does not enable centered/bottom minimum-slot alignment in top rows.");
            }
        }
    }

    public static void DocxTableTopAutoClipRetainsFrameAndComplexFallbacks()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string variant in new[] { "header", "rich", "frame", "four-rows", "three-columns", "metrics" })
        {
            IDocxTextMeasurer measurer = variant == "metrics" ? new MinimumMetricsFallback(full) : full;
            var layout = CreateParagraphMinimumLayout("min,min,min", measurer,
                variant is "header" or "rich" ? variant : "", verticalAlignment: "top",
                rowCount: variant == "four-rows" ? 4 : 2, columnCount: variant == "three-columns" ? 3 : 2,
                pageHeight: variant == "frame" ? 160d : 2400d);
            foreach (var row in layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()))
            foreach (var cell in row.Cells)
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                    "Unqualified top-mask shapes, metrics and complex/frame-boundary paths retain full-cell fallback: " + variant);
        }
        var admitted = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top")
            .Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        foreach (var row in new[] { admitted with { FragmentCount = 2 }, admitted with { ReviewBorderPrintScale = .95d } })
        foreach (var cell in row.Cells)
            TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                "Fragmented/scaled top rows retain the full-cell mask.");
    }

    public static void DocxTableTopAutoClipRetainsUnqualifiedMixedHeightFallbacks()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string? rule in new string?[] { null, "exact", "atLeast" })
        {
            var heights = new (double? Height, string? Rule)[] { (180d, rule), (null, null) };
            var layout = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "top",
                rowCount: 2, columnCount: rule is null ? 2 : 3, rowHeightRules: heights);
            foreach (var row in layout.Pages[0].Items.OfType<DocxTableRowLayout>())
            foreach (var cell in row.Cells)
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                    "Default declared rules and larger mixed top tables retain their unqualified masks.");
        }
    }

    public static void DocxTableMixedHeightGridAutomaticSlotsMatchIndependentWord()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach (string? automaticRule in new string?[] { null, "auto" })
        foreach (string declaredRule in new[] { "exact", "atLeast" })
        foreach (int automaticIndex in new[] { 0, 1 })
        foreach (string pattern in new[] { "min-single", "min", "auto,min-single,auto" })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var heights = new (double? Height, string? Rule)[2];
            heights[automaticIndex] = (automaticRule is null ? null : 320d, automaticRule);
            heights[1 - automaticIndex] = (180d, declaredRule);
            var uniform = CreateParagraphMinimumLayout(pattern, full, rowRule: automaticRule,
                verticalAlignment: alignment, cellMargins: margins, rowCount: 2, columnCount: 2, borderSize: "8");
            var mixed = CreateParagraphMinimumLayout(pattern, full, verticalAlignment: alignment,
                cellMargins: margins, rowCount: 2, columnCount: 2, borderSize: "8", rowHeightRules: heights);
            var expected = uniform.Pages[0].Items.OfType<DocxTableRowLayout>().ElementAt(automaticIndex);
            var actual = mixed.Pages[0].Items.OfType<DocxTableRowLayout>().ElementAt(automaticIndex);
            TestAssert.Equal(expected.Height, actual.Height);
            for (int c = 0; c < 2; c++)
            for (int i = 0; i < actual.Cells[c].TextLines.Count; i++)
                TestAssert.True(Math.Abs((expected.Cells[c].TextLines[i].BaselineY - expected.Y) -
                    (actual.Cells[c].TextLines[i].BaselineY - actual.Y)) < .000001d,
                    "An automatic row keeps its independently qualified minimum slots beside a declared-height row.");
        }
    }

    public static void DocxTableMixedHeightGridRetainsDeclaredSlotsAndFollowingFlow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach (string declaredRule in new[] { "exact", "atLeast" })
        foreach (int automaticIndex in new[] { 0, 1 })
        foreach (double minimum in new[] { 12d, 14d, 24d })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var heights = new (double? Height, string? Rule)[2];
            heights[automaticIndex] = (null, null); heights[1 - automaticIndex] = (180d, declaredRule);
            var uniform = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: alignment,
                cellMargins: margins, rowCount: 2, columnCount: 2, minimumSpacing: minimum, borderSize: "8");
            var declared = CreateParagraphMinimumLayout("min-single", full, rowHeight: 180d, rowRule: declaredRule,
                verticalAlignment: alignment, cellMargins: margins, rowCount: 2, columnCount: 2,
                minimumSpacing: minimum, borderSize: "8");
            var mixed = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: alignment,
                cellMargins: margins, rowCount: 2, columnCount: 2, minimumSpacing: minimum,
                borderSize: "8", rowHeightRules: heights);
            var before = declared.Pages[0].Items.OfType<DocxTableRowLayout>().ElementAt(1 - automaticIndex);
            var after = mixed.Pages[0].Items.OfType<DocxTableRowLayout>().ElementAt(1 - automaticIndex);
            TestAssert.Equal(before.Height, after.Height);
            for (int c = 0; c < 2; c++)
            {
                var first = before.Cells[c].TextLines.Single(); var last = after.Cells[c].TextLines.Single();
                TestAssert.True(Math.Abs((first.BaselineY - before.Y) - (last.BaselineY - after.Y)) < .000001d,
                    "Declared rows retain their existing aligned minimum and natural-floor slots.");
                TestAssert.Equal(first.BodyLineBoxBaselineInsetPoints, last.BodyLineBoxBaselineInsetPoints);
            }
            double extraHeight = mixed.Pages[0].Items.OfType<DocxTableRowLayout>().Sum(r => r.Height) -
                uniform.Pages[0].Items.OfType<DocxTableRowLayout>().Sum(r => r.Height);
            double followingDelta = uniform.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY -
                mixed.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
            TestAssert.True(Math.Abs(extraHeight - followingDelta) < .000001d,
                "Following body flow advances only by the measured mixed-row height difference.");
        }
    }

    public static void DocxTableMixedHeightGridContentMasksMatchIndependentWord()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (int automaticIndex in new[] { 0, 1 })
        {
            var heights = new (double? Height, string? Rule)[2];
            heights[automaticIndex] = (null, null); heights[1 - automaticIndex] = (180d, "exact");
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var row = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "center",
                cellMargins: margins, rowCount: 2, columnCount: 2, borderSize: "8", rowHeightRules: heights)
                .Pages[0].Items.OfType<DocxTableRowLayout>().ElementAt(automaticIndex);
            foreach (var cell in row.Cells)
            {
                var clip = DocxRenderer.ResolveTableCellContentClip(row, cell);
                double bottom = 12d + (automaticIndex == 1 ? .96d : 0d);
                TestAssert.True(clip.Admitted, "Qualified automatic rows in a mixed grid expose their content masks.");
                TestAssert.True(Math.Abs(clip.X - cell.X - .48d) < .000001d &&
                    Math.Abs(clip.Y - cell.Y - bottom) < .000001d &&
                    Math.Abs(clip.Width - cell.Width + .96d) < .000001d &&
                    Math.Abs(clip.Height - cell.Height + 6.96d + bottom) < .000001d,
                    "Each automatic mask uses row margins and only the terminal row's bottom border.");
            }
        }
    }

    public static void DocxTableMixedHeightGridRetainsExcludedTableFallbacks()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string variant in new[] { "header", "rich", "three-rows", "one-column", "frame" })
        {
            int count = variant == "three-rows" ? 3 : 2;
            var heights = Enumerable.Range(0, count).Select(i => i == 1 ?
                ((double?)180d, (string?)"exact") : ((double?)null, (string?)null)).ToArray();
            var layout = CreateParagraphMinimumLayout("min-single", full,
                variant is "header" or "rich" ? variant : "", verticalAlignment: "center",
                rowCount: count, columnCount: variant == "one-column" ? 1 : 2,
                pageHeight: variant == "frame" ? 160d : 700d, rowHeightRules: heights);
            foreach (var row in layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>())
                .Where(r => r.DeclaredHeightPoints is null && r.HeightRuleValue is null))
            foreach (var cell in row.Cells)
            {
                TestAssert.True(cell.TextLines.All(l => l.BodyLineBoxBaselineInsetPoints is null),
                    "Unqualified mixed-grid shapes and complex/frame-boundary paths retain placement fallback: " + variant);
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                    "Unqualified mixed-grid masks retain full-cell fallback.");
            }
        }
    }

    public static void DocxTableAutomaticGridRetainsFractionalSingleCellPlacement()
    {
        var full = new FractionalCellBoxTextMeasurer();
        foreach (string alignment in new[] { "center", "bottom" })
        foreach (string? rule in new string?[] { null, "auto" })
        foreach (double minimum in new[] { 12d, 14d, 24d })
        foreach (string pattern in new[] { "min", "exact-wrap,min" })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var legacy = CreateParagraphMinimumLayout(pattern, full, "header", rowRule: rule,
                verticalAlignment: alignment, cellMargins: margins, minimumSpacing: minimum, exactSpacing: 4d, borderSize: "32");
            var actual = CreateParagraphMinimumLayout(pattern, full, rowRule: rule,
                verticalAlignment: alignment, cellMargins: margins, minimumSpacing: minimum, exactSpacing: 4d, borderSize: "32");
            var before = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            var after = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.Equal(before.Y, after.Y);
            TestAssert.Equal(before.Height, after.Height);
            var previous = before.Cells.Single().TextLines;
            var current = after.Cells.Single().TextLines;
            TestAssert.Equal(previous.Count, current.Count);
            for (int i = 0; i < current.Count; i++)
            {
                double excess = current[i].LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing &&
                    minimum > 14.6484375d ? (current[i].LineHeight ?? 0d) - 14.6484375d : 0d;
                TestAssert.True(Math.Abs(previous[i].BaselineY - current[i].BaselineY - excess) < .000001d,
                    "Automatic single-cell tables retain their accepted alignment with fractional metrics and exact/minimum overflow.");
            }
            TestAssert.Equal(legacy.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY,
                actual.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY);
        }
    }

    private sealed class FractionalCellBoxTextMeasurer : IDocxTextMeasurer, IDocxLineMetricsProvider,
        IDocxStaticTextMetricsProvider, IDocxHheaDescenderProvider, IDocxHheaLineGapProvider
    {
        private readonly CellBoxTextMeasurer inner = new(new DocxTests.FamilyWidthTextMeasurer());
        public double MeasureText(DocxTextRun? run, string text, double size) => inner.MeasureText(run, text, size);
        public double MeasureSingleLineHeight(DocxTextRun? run, double size) => 14.6484375d;
        public double MeasureHheaLineHeight(DocxTextRun? run, double size) => 14.6484375d;
        public double MeasureHheaAscender(DocxTextRun? run, double size) => inner.MeasureHheaAscender(run, size);
        public double MeasureHheaDescender(DocxTextRun? run, double size) => inner.MeasureHheaDescender(run, size);
        public double MeasureHheaLineGap(DocxTextRun? run, double size) => inner.MeasureHheaLineGap(run, size);
        public double MeasureWindowsAscender(DocxTextRun? run, double size) => inner.MeasureWindowsAscender(run, size);
        public double MeasureWindowsDescender(DocxTextRun? run, double size) => inner.MeasureWindowsDescender(run, size);
    }

    public static void DocxTableAutomaticGridMinimumSlotsMatchIndependentWord()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach ((int rows, int columns) in new[] { (1, 2), (2, 1), (2, 2) })
        foreach (string alignment in new[] { "center", "bottom" })
        foreach ((double? height, string? rule) in new (double?, string?)[] { (null, null), (320d, "auto") })
        foreach (string pattern in new[] { "min-single", "min", "auto,min-single,auto" })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var legacy = CreateParagraphMinimumLayout(pattern, full, "header", height, rule,
                verticalAlignment: alignment, cellMargins: margins, rowCount: rows, columnCount: columns);
            var actual = CreateParagraphMinimumLayout(pattern, full, rowHeight: height, rowRule: rule,
                verticalAlignment: alignment, cellMargins: margins, rowCount: rows, columnCount: columns);
            var before = legacy.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells).SelectMany(c => c.TextLines).ToArray();
            var after = actual.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells).SelectMany(c => c.TextLines).ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int i = 0; i < after.Length; i++)
                TestAssert.True(Math.Abs(before[i].BaselineY - after[i].BaselineY -
                    (after[i].LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing ? 10d : 0d)) < .000001d,
                    "Word moves only the minimum-height paragraph slot in each plain automatic grid cell.");
        }
    }

    public static void DocxTableAutomaticGridAlignsUnequalAndNaturalHeightSlots()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach (double shortMinimum in new[] { 12d, 14d, 24d })
        {
            var legacy = CreateParagraphMinimumLayout("min", full, "header", verticalAlignment: alignment,
                rowCount: 2, columnCount: 2, shortNeighborMinimum: shortMinimum);
            var actual = CreateParagraphMinimumLayout("min", full, verticalAlignment: alignment,
                rowCount: 2, columnCount: 2, shortNeighborMinimum: shortMinimum);
            var before = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            var after = actual.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            for (int r = 0; r < 2; r++)
            {
                // The synthetic face has a 10.44pt first-baseline inset. Word
                // centers/bottom-aligns the slot without subtracting that inset.
                double expected = (shortMinimum > 14d ? 10d : 0d) + (alignment == "center" ? 5.22d : 10.44d);
                TestAssert.True(Math.Abs(before[r].Cells[1].TextLines[0].BaselineY -
                    after[r].Cells[1].TextLines[0].BaselineY - expected) < .000001d,
                    "Unequal minimum/natural-height cells align their full paragraph slot beside taller cells.");
                TestAssert.Equal(before[r].Height, after[r].Height);
            }
            TestAssert.Equal(legacy.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY,
                actual.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY);
        }
        foreach (string alignment in new[] { "center", "bottom" })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            var neighborMargins = DocxTableCellMargins.Empty with { TopPoints = 12d, BottomPoints = 6d };
            var legacy = CreateParagraphMinimumLayout("min-single", full, "header", verticalAlignment: alignment,
                cellMargins: margins, columnCount: 2, shortNeighborMinimum: 24d, shortNeighborMargins: neighborMargins);
            var actual = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: alignment,
                cellMargins: margins, columnCount: 2, shortNeighborMinimum: 24d, shortNeighborMargins: neighborMargins);
            var before = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells[1].TextLines[0];
            var after = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells[1].TextLines[0];
            TestAssert.True(Math.Abs(before.BaselineY - after.BaselineY - 10d) < .000001d,
                "Unequal margins align against the row's maximum bottom margin without adding unused cell padding.");
        }
    }

    public static void DocxTableAutomaticGridRetainsNaturalFloorsAndFollowingFlow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach ((int rows, int columns) in new[] { (1, 2), (2, 1), (2, 2) })
        foreach (double minimum in new[] { 12d, 14d, 24d })
        {
            var legacy = CreateParagraphMinimumLayout("auto,min-single,auto", full, "header", rowRule: "auto",
                verticalAlignment: "center", rowCount: rows, columnCount: columns, minimumSpacing: minimum);
            var actual = CreateParagraphMinimumLayout("auto,min-single,auto", full, rowRule: "auto",
                verticalAlignment: "center", rowCount: rows, columnCount: columns, minimumSpacing: minimum);
            var before = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            var after = actual.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int r = 0; r < after.Length; r++)
            {
                TestAssert.Equal(before[r].Y, after[r].Y); TestAssert.Equal(before[r].Height, after[r].Height);
                for (int c = 0; c < columns; c++)
                {
                    TestAssert.Equal(before[r].Cells[c].X, after[r].Cells[c].X);
                    TestAssert.Equal(before[r].Cells[c].Width, after[r].Cells[c].Width);
                    if (minimum > 14d) continue;
                    for (int i = 0; i < after[r].Cells[c].TextLines.Count; i++)
                        TestAssert.Equal(before[r].Cells[c].TextLines[i].BaselineY, after[r].Cells[c].TextLines[i].BaselineY);
                }
            }
            TestAssert.Equal(legacy.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY,
                actual.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY);
        }
    }

    public static void DocxTableAutomaticGridClipsUseIndependentWordInsets()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        var rows = CreateParagraphMinimumLayout("min-single", full, rowRule: "auto", verticalAlignment: "center",
            rowCount: 2, columnCount: 2).Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        for (int ri = 0; ri < 2; ri++)
        {
            double y = ri == 0 ? 597.12d : 471.36d;
            double height = ri == 0 ? 122.88d : 125.76d;
            var cells = rows[ri].Cells.Select((cell, ci) => cell with {
                X = 72d + ci * 234d, Y = y, Width = 234d, Height = height,
                ContentPaddingTop = 6d, ContentPaddingBottom = 12d,
                Cell = cell.Cell with { Borders = new[] { "top", "left", "bottom", "right" }
                    .Select(edge => new DocxTableCellBorder(edge, "double", "000000", "8")).ToArray() }
            }).ToArray();
            // Read the qualified layout's border projection, rather than setting its marker in the test.
            DocxTable source = new(null, [234d, 234d], [new(cells.Select(c => c.Cell).ToArray(), null)]);
            DocxTable projected = DocxTableBorderGeometry.ResolveUniformDoubleBodyTable(source, 1d, 468d);
            cells = cells.Select((cell, ci) => cell with { Cell = projected.Rows[0].Cells[ci] }).ToArray();
            var row = rows[ri] with { Cells = cells };
            for (int ci = 0; ci < 2; ci++)
            {
                var clip = DocxRenderer.ResolveTableCellContentClip(row, cells[ci]);
                TestAssert.True(clip.Admitted, "Every qualified automatic grid row must expose its content mask.");
                TestAssert.True(Math.Abs(clip.X - (73.44d + ci * 234d)) < .000001d &&
                    Math.Abs(clip.Y - (ri == 0 ? 609.12d : 486.24d)) < .000001d &&
                    Math.Abs(clip.Width - 231.12d) < .000001d && Math.Abs(clip.Height - 102d) < .000001d,
                    "Grid masks use each row's margins and only the terminal row's bottom border advance.");
            }
        }
    }

    public static void DocxTableAutomaticGridRetainsComplexAndFragmentFallbacks()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string variant in new[] { "header", "rich", "mixed", "frame" })
        {
            var layout = CreateParagraphMinimumLayout("min,min,min", full, variant == "frame" ? "" : variant,
                verticalAlignment: "center", rowCount: 2, columnCount: 2, pageHeight: variant == "frame" ? 260d : 700d);
            foreach (var row in layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()))
            foreach (var cell in row.Cells)
            {
                TestAssert.True(cell.TextLines.All(line => line.BodyLineBoxBaselineInsetPoints is null),
                    "Complex grids and tables that do not fit their complete current frame retain placement fallback.");
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                    "Unqualified automatic grid masks retain full-cell clipping.");
            }
        }
        var admitted = CreateParagraphMinimumLayout("min-single", full, verticalAlignment: "bottom", rowCount: 2, columnCount: 2)
            .Pages[0].Items.OfType<DocxTableRowLayout>().First();
        foreach (var row in new[] { admitted with { FragmentCount = 2 }, admitted with { ReviewBorderPrintScale = .95d } })
            foreach (var cell in row.Cells)
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                    "Fragmented and scaled automatic grid masks retain fallback.");
    }

    public static void DocxTableDoubleBorderUsesIndependentWordBandWidths()
    {
        foreach ((string? size, double nominal, double printed) in new (string?, double, double)[] {
            ("4", 1.5d, 1.44d), ("8", 3d, 2.88d), ("24", 9d, 8.64d) })
        {
            var document = ReadDoubleBandFixture(size);
            var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document,
                new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
            var borders = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single().Cell.Borders;
            foreach (var border in borders)
            {
                TestAssert.Equal(nominal, DocxTableBorderGeometry.ResolveNominalWidth(border));
                TestAssert.True(Math.Abs(printed - DocxTableBorderGeometry.ResolveVisibleWidth(border)) < .000001d,
                    "Word double bands have two authored-width strokes and one equal gap.");
            }
        }
        foreach (string style in new[] { "single", "triple", "thinThickSmallGap", "dotted" })
            TestAssert.Equal(.96d, DocxTableBorderGeometry.ResolveVisibleWidth(new DocxTableCellBorder("top", style, "000000", "8")));
        foreach (string style in new[] { "nil", "none" })
            TestAssert.Equal(0d, DocxTableBorderGeometry.ResolveVisibleWidth(new DocxTableCellBorder("top", style, "000000", "8")));
    }

    public static void DocxTableDoubleBorderClipMatchesIndependentWordBounds()
    {
        DocxDocument document = ReadDoubleBandFixture("8");
        var renderer = MarginClipRenderer();
        TestAssert.Contains("73.44 540 465.12 171.12 re W n",
            renderer.RenderBlankPages(document, null, CancellationToken.None).First().Content);
        using var json = System.Text.Json.JsonDocument.Parse(System.Text.Json.JsonSerializer.Serialize(renderer.InspectLayout(document)));
        var clip = json.RootElement.GetProperty("Pages")[0].GetProperty("TableRows")[0].GetProperty("Cells")[0].GetProperty("ContentClip");
        TestAssert.True(clip.ValueKind == System.Text.Json.JsonValueKind.Object && clip.GetProperty("Admitted").GetBoolean(),
            "Qualified double-border body masks must expose the same effective snapshot bounds.");
    }

    public static void DocxTableDoubleBorderRetainsMixedAndRawWidths()
    {
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(ReadDoubleBandFixture("8", mixed: true),
            new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        var row = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        var cell = row.Cells.Single();
        TestAssert.Equal(1d, DocxTableBorderGeometry.ResolveNominalWidth(cell.Cell.Borders.Single(border => border.Edge == "top")));
        TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
            "Mixed double/single styles must retain the prior band widths and complete cell mask.");
        // Unprojected borders also retain their raw widths in static/related stories.
        TestAssert.Equal(.96d, DocxTableBorderGeometry.ResolveVisibleWidth(new DocxTableCellBorder("top", "double", "000000", "8")));
        foreach (string? size in new string?[] { null, "invalid", "0" })
        {
            var fallback = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(ReadDoubleBandFixture(size),
                new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
            var fallbackRow = fallback.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            var fallbackCell = fallbackRow.Cells.Single();
            TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(fallbackRow, fallbackCell).Admitted,
                "Omitted/invalid/nonpositive authored widths must retain the prior mask.");
            TestAssert.Equal(size == "0" ? .25d : .75d, DocxTableBorderGeometry.ResolveNominalWidth(fallbackCell.Cell.Borders[0]));
        }
    }

    private static DocxDocument ReadDoubleBandFixture(string? size, bool mixed = false)
    {
        DocxDocument source = ReadMarginClipFixture("center");
        DocxTable original = source.BodyElements.OfType<DocxTableElement>().Single().Table;
        DocxTable table = original with { Rows = original.Rows.Select(row => row with { Cells = row.Cells.Select(cell =>
            cell with { Borders = cell.Borders.Select(border => border with {
                Value = mixed && border.Edge != "top" ? "single" : "double", SizeValue = size
            }).ToArray() }).ToArray() }).ToArray() };
        return source with {
            BodyElements = source.BodyElements.Select(e => e is DocxTableElement ? (DocxBodyElement)new DocxTableElement(table) : e).ToArray(),
            FallbackTables = [table] };
    }

    public static void DocxTableDoubleBorderRetainsComplexAndFrameClipFallbacks()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string variant in new[] { "header", "rich", "mixed" })
        {
            var row = CreateParagraphMinimumLayout("min,min,min", full, variant, verticalAlignment: "center")
                .Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            var cell = row.Cells.Single() with { Cell = row.Cells.Single().Cell with {
                Borders = [new DocxTableCellBorder("top", "double", "000000", "8")] } };
            row = row with { Cells = [cell] };
            TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                "Double width correction must not broaden the body/frame admission.");
        }
        var automatic = CreateParagraphMinimumLayout("min", full, verticalAlignment: "center")
            .Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        var automaticCell = automatic.Cells.Single() with { Cell = automatic.Cells.Single().Cell with {
            Borders = [new DocxTableCellBorder("bottom", "double", "000000", "8")] } };
        foreach (var row in new[] { automatic with { Cells = [automaticCell], FragmentCount = 2 },
            automatic with { Cells = [automaticCell], ReviewBorderPrintScale = .95d } })
            TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, automaticCell).Admitted,
                "Fragmented and scaled masks retain fallback after the double width fix.");
    }

    public static void DocxTableAutoClipMatchesIndependentWordInsets()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach ((double? height, string? rule) in new (double?, string?)[] { (null, null), (320d, "auto") })
        foreach ((string style, string size, double x, double y, double width, double clipHeight) in new[] {
            ("single", "8", 72.48d, 659.04d, 467.04d, 54d),
            ("nil", "8", 72d, 658.08d, 468d, 55.92d),
            ("single", "32", 73.92d, 661.92d, 464.16d, 48.24d) })
        {
            DocxTableRowLayout row = CreateParagraphMinimumLayout("min", full, rowHeight: height,
                rowRule: rule, verticalAlignment: alignment).Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            // Independent Word controls establish the margin and border insets.
            // Keep the existing 0.96 border quantization and a fixed source frame.
            DocxTableCellLayout cell = row.Cells.Single() with {
                X = 72d, Y = 646.08d, Width = 468d, Height = 73.92d,
                ContentPaddingTop = 6d, ContentPaddingBottom = 12d,
                Cell = row.Cells.Single().Cell with { Borders = new[] { "top", "left", "bottom", "right" }
                    .Select(edge => new DocxTableCellBorder(edge, style, "000000", size)).ToArray() } };
            row = row with { Cells = [cell] };
            var clip = DocxRenderer.ResolveTableCellContentClip(row, cell);
            TestAssert.True(clip.Admitted, "Eligible automatic cells must use the independent Word content bounds.");
            TestAssert.True(Math.Abs(clip.X - x) < .000001d && Math.Abs(clip.Y - y) < .000001d &&
                Math.Abs(clip.Width - width) < .000001d && Math.Abs(clip.Height - clipHeight) < .000001d,
                "Automatic masks must exclude margins and full vertical/half horizontal border widths.");
        }
    }

    public static void DocxTableAutoClipSnapshotMatchesEmittedMask()
    {
        foreach (string alignment in new[] { "center", "bottom" })
        {
            DocxDocument document = DocxTests.ReadDocx(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
                "..", "..", "..", "Cases", "table-automatic-margin-minimum-line-spacing-" + alignment + ".docx")),
                OoxPdfDocxMarkupMode.AllMarkup);
            var renderer = MarginClipRenderer();
            var snapshot = renderer.InspectLayout(document);
            TestAssert.Equal(1, snapshot.Pages[0].TableRows.Count);
            using var json = System.Text.Json.JsonDocument.Parse(System.Text.Json.JsonSerializer.Serialize(snapshot));
            var cell = json.RootElement.GetProperty("Pages")[0].GetProperty("TableRows")[0].GetProperty("Cells")[0];
            TestAssert.True(cell.TryGetProperty("ContentClip", out var clip) && clip.ValueKind == System.Text.Json.JsonValueKind.Object,
                "The source snapshot must expose the admitted automatic content mask.");
            string rectangle = string.Join(" ", new[] { "X", "Y", "Width", "Height" }
                .Select(name => PdfDocumentWriter.FormatNumber(clip.GetProperty(name).GetDouble())));
            string content = renderer.RenderBlankPages(document, null, CancellationToken.None).First().Content;
            TestAssert.Contains(rectangle + " re W n", content);
            TestAssert.True(clip.GetProperty("Admitted").GetBoolean(), "Snapshot must record automatic mask admission.");
        }
    }

    public static void DocxTableAutoClipRetainsUnsupportedAndCollapsedBounds()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach ((string variant, int rows, int columns, double pageHeight) in new[] {
            ("header", 1, 1, 700d), ("rich", 1, 1, 700d), ("mixed", 1, 1, 700d),
            ("header", 2, 1, 700d), ("header", 1, 2, 700d), ("", 1, 1, 160d) })
        {
            var layout = CreateParagraphMinimumLayout("min,min,min", full, variant, verticalAlignment: "center",
                pageHeight: pageHeight, rowCount: rows, columnCount: columns);
            foreach (var row in layout.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()))
            foreach (var cell in row.Cells)
                TestAssert.True(!DocxRenderer.ResolveTableCellContentClip(row, cell).Admitted,
                    "Complex and frame-boundary automatic cells must retain the full mask.");
        }
        DocxTableRowLayout admitted = CreateParagraphMinimumLayout("min", full, verticalAlignment: "center")
            .Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        foreach (string variant in new[] { "compound", "fragment", "scaled", "width", "height" })
        {
            var cell = admitted.Cells.Single();
            var row = admitted;
            if (variant == "compound") cell = cell with { Cell = cell.Cell with {
                Borders = [new DocxTableCellBorder("bottom", "triple", "000000", "8")] } };
            if (variant == "fragment") row = row with { FragmentCount = 2 };
            if (variant == "scaled") row = row with { ReviewBorderPrintScale = .95d };
            if (variant == "width") cell = cell with { Width = 0d };
            if (variant == "height") cell = cell with { Height = 1d, ContentPaddingTop = 6d, ContentPaddingBottom = 12d };
            row = row with { Cells = [cell] };
            var clip = DocxRenderer.ResolveTableCellContentClip(row, cell);
            TestAssert.True(!clip.Admitted && clip.X == cell.X && clip.Y == cell.Y && clip.Width == cell.Width && clip.Height == cell.Height,
                "Unsupported or collapsed masks retain their exact full-cell bounds: " + variant);
        }
    }

    public static void DocxTableAutoMinimumPlacementMatchesIndependentWordSlots()
    {
        var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach ((double? height, string? rule) in new (double?, string?)[] { (null, null), (320d, "auto") })
        foreach ((double top, double bottom) in new[] { (0d, 0d), (6d, 12d), (12d, 6d) })
        foreach (string pattern in new[] { "min", "min-single", "auto,min,auto", "auto,min-single,auto", "min,auto", "auto,min", "exact,min,auto", "auto,min,exact" })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = top, BottomPoints = bottom };
            DocxTableCellLayout legacy = CreateParagraphMinimumLayout(pattern, measurer, "header", height, rule,
                verticalAlignment: alignment, cellMargins: margins).Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            DocxTableCellLayout actual = CreateParagraphMinimumLayout(pattern, measurer, rowHeight: height, rowRule: rule,
                verticalAlignment: alignment, cellMargins: margins).Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            TestAssert.Equal(legacy.TextLines.Count, actual.TextLines.Count);
            for (int i = 0; i < actual.TextLines.Count; i++)
            {
                double expected = actual.TextLines[i].LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing ? 10d : 0d;
                TestAssert.True(Math.Abs(legacy.TextLines[i].BaselineY - actual.TextLines[i].BaselineY - expected) < .000001d,
                    "Word moves only the minimum-height paragraph slot in a bounded natural-height aligned cell.");
            }
        }
    }

    public static void DocxTableAutoMinimumPlacementRetainsNaturalFloors()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach (double minimum in new[] { 12d, 14d })
        foreach (string pattern in new[] { "min", "auto,min,auto" })
        {
            DocxLayout legacy = CreateParagraphMinimumLayout(pattern, full, "header", rowRule: "auto",
                verticalAlignment: alignment, minimumSpacing: minimum);
            DocxLayout actual = CreateParagraphMinimumLayout(pattern, full, rowRule: "auto",
                verticalAlignment: alignment, minimumSpacing: minimum);
            DocxTextLineLayout[] before = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single().TextLines.ToArray();
            DocxTextLineLayout[] after = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single().TextLines.ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int i = 0; i < after.Length; i++)
            {
                TestAssert.Equal(before[i].BaselineY, after[i].BaselineY);
                TestAssert.Equal(before[i].LineHeight, after[i].LineHeight);
                TestAssert.True(after[i].BodyLineBoxBaselineInsetPoints is null,
                    "A minimum at or below the natural font height must retain the original emission branch without a correction.");
            }
        }
    }

    public static void DocxTableAutoMinimumPlacementPreservesRowsAndFollowingFlow()
    {
        var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach ((double? height, string? rule) in new (double?, string?)[] { (null, null), (320d, "auto") })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
            DocxLayout legacy = CreateParagraphMinimumLayout("auto,min,auto", measurer, "header", height, rule,
                verticalAlignment: alignment, cellMargins: margins);
            DocxLayout actual = CreateParagraphMinimumLayout("auto,min,auto", measurer, rowHeight: height, rowRule: rule,
                verticalAlignment: alignment, cellMargins: margins);
            DocxTableRowLayout before = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            DocxTableRowLayout after = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.Equal(before.Y, after.Y); TestAssert.Equal(before.Height, after.Height);
            TestAssert.Equal(before.Cells[0].X, after.Cells[0].X); TestAssert.Equal(before.Cells[0].Width, after.Cells[0].Width);
            TestAssert.Equal(legacy.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY,
                actual.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY);
        }
    }

    public static void DocxTableAutoMinimumPlacementRetainsComplexGridsAndFrameBoundaries()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
        foreach (string alignment in new[] { "center", "bottom" })
        foreach ((string variant, int rows, int columns, double pageHeight, double? height, string? rule) in new (string, int, int, double, double?, string?)[] {
            ("header", 1, 1, 700d, null, null), ("rich", 1, 1, 700d, null, null), ("mixed", 1, 1, 700d, 320d, "auto"),
            ("header", 2, 1, 700d, null, null), ("header", 1, 2, 700d, null, null), ("", 1, 1, 160d, null, null),
            ("", 1, 1, 700d, 80d, "atLeast"), ("", 1, 1, 700d, 320d, "unknown") })
        {
            DocxLayout layout = CreateParagraphMinimumLayout("min,min,min", full, variant, height, rule,
                verticalAlignment: alignment, cellMargins: margins, pageHeight: pageHeight, rowCount: rows, columnCount: columns);
            TestAssert.True(layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells)
                .SelectMany(c => c.TextLines).All(l => l.BodyLineBoxBaselineInsetPoints is null),
                "Automatic-row placement requires an eligible plain body table that fits its complete current frame.");
        }
        DocxLayout fallback = CreateParagraphMinimumLayout("auto,min,auto", new MinimumMetricsFallback(full), rowRule: "auto",
            verticalAlignment: "center", cellMargins: margins);
        TestAssert.True(fallback.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells)
            .SelectMany(c => c.TextLines).All(l => l.BodyLineBoxBaselineInsetPoints is null), "Partial font metrics must retain fallback.");
    }

    public static void DocxTableMarginClipUsesWordContentBounds()
    {
        DocxDocument document = ReadMarginClipFixture("center");
        var renderer = MarginClipRenderer();
        string content = renderer.RenderBlankPages(document, null, CancellationToken.None).First().Content;
        // Word's 1pt borders quantize to 0.96pt in our existing border geometry.
        TestAssert.Contains("72.48 540 467.04 173.04 re W n", content);
        TestAssert.True(!content.Contains("72 527.04 468 192.96 re W n", StringComparison.Ordinal),
            "The content mask must exclude the positive margins inside the unchanged cell border.");
        using var json = System.Text.Json.JsonDocument.Parse(System.Text.Json.JsonSerializer.Serialize(renderer.InspectLayout(document)));
        var cell = json.RootElement.GetProperty("Pages")[0].GetProperty("TableRows")[0].GetProperty("Cells")[0];
        TestAssert.True(cell.TryGetProperty("ContentClip", out var clip) && clip.ValueKind == System.Text.Json.JsonValueKind.Object,
            "The source snapshot must expose the content rectangle used for PDF clipping.");
        TestAssert.True(clip.GetProperty("Admitted").GetBoolean(), "The bounded body cell should be admitted.");
        TestAssert.True(Math.Abs(clip.GetProperty("Y").GetDouble() - 540d) < .001d,
            "Snapshot and emitted content clip must share the Word-qualified bottom bound.");
    }

    public static void DocxTableMarginClipUsesRowMarginsAndTerminalBorder()
    {
        DocxDocument document = ReadMarginClipFixture("bottom");
        var renderer = MarginClipRenderer();
        string content = renderer.RenderBlankPages(document, null, CancellationToken.None).First().Content;
        foreach (string rectangle in new[] {
            "72.48 540 233.04 173.04", "306.48 540 233.04 173.04",
            "72.48 348 233.04 173.04", "306.48 348 233.04 173.04" })
            TestAssert.Contains(rectangle + " re W n", content);
        TestAssert.Equal(4, Regex.Matches(content, " re W n").Count);
        TestAssert.Equal(2, renderer.InspectLayout(document).Pages[0].TableRows.Count);
    }

    public static void DocxTableMarginClipRetainsUnsupportedGeometry()
    {
        DocxDocument source = ReadMarginClipFixture("center");
        DocxTable original = source.BodyElements.OfType<DocxTableElement>().Single().Table;
        foreach (string variant in new[] { "auto-rich", "header", "unknown-top", "rich", "compound", "frame" })
        {
            DocxTableRow row = original.Rows.Single();
            if (variant == "auto-rich") row = row with { HeightRuleValue = "auto" };
            if (variant == "header") row = row with { IsHeader = true };
            if (variant == "frame") row = row with { HeightPoints = 720d };
            if (variant is "unknown-top" or "rich" or "auto-rich" or "compound") row = row with { Cells = row.Cells.Select(cell =>
            {
                if (variant == "unknown-top") return cell with { VerticalAlignmentValue = "unrecognized" };
                if (variant == "compound") return cell with { Borders = cell.Borders.Select(border => border with { Value = "triple" }).ToArray() };
                DocxParagraph[] paragraphs = cell.Paragraphs.Select(paragraph => paragraph with {
                    Runs = paragraph.Runs.Select(run => run with { Text = run.Text + "\n" }).ToArray() }).ToArray();
                return cell with { Paragraphs = paragraphs,
                    BodyElements = paragraphs.Select(paragraph => (DocxBodyElement)new DocxParagraphElement(paragraph)).ToArray() };
            }).ToArray() };
            DocxTable table = original with { Rows = [row] };
            DocxDocument document = source with {
                BodyElements = source.BodyElements.Select(e => e is DocxTableElement ? (DocxBodyElement)new DocxTableElement(table) : e).ToArray(),
                FallbackTables = [table] };
            var renderer = MarginClipRenderer();
            var snapshot = renderer.InspectLayout(document);
            string content = string.Concat(renderer.RenderBlankPages(document, null, CancellationToken.None).Select(page => page.Content));
            TestAssert.True(snapshot.Pages.Sum(page => page.TableRows.Count) > 0, "Fallback controls must contain rendered table cells.");
            foreach (var cell in snapshot.Pages.SelectMany(page => page.TableRows).SelectMany(row => row.Cells))
            {
                string rectangle = string.Join(" ", new[] { cell.X, cell.Y, cell.Width, cell.Height }.Select(PdfDocumentWriter.FormatNumber));
                TestAssert.True(content.Contains(rectangle + " re W n", StringComparison.Ordinal),
                    "Expected full clip for " + variant + ": " + rectangle + "; emitted: " + string.Join("; ", Regex.Matches(content, @"[^\r\n]+ re W n").Select(match => match.Value)));
                using var json = System.Text.Json.JsonDocument.Parse(System.Text.Json.JsonSerializer.Serialize(cell));
                TestAssert.True(!json.RootElement.TryGetProperty("ContentClip", out var clip) || clip.ValueKind == System.Text.Json.JsonValueKind.Null,
                    "Unsupported cell geometry must retain its complete cell clip: " + variant);
            }
        }
    }

    private static DocxRenderer MarginClipRenderer() => new(new TestFaceFontResolver(),
        OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);

    private static DocxDocument ReadMarginClipFixture(string alignment) => DocxTests.ReadDocx(
        Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "table-aligned-margin-minimum-line-spacing-" + alignment + ".docx")), OoxPdfDocxMarkupMode.AllMarkup);

    public static void DocxTableMarginMinimumPlacementMatchesIndependentWordSlots()
    {
        var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach ((double top, double bottom) in new[] { (6d, 0d), (0d, 12d), (12d, 6d), (12d, 12d) })
        foreach (string pattern in new[] { "min,auto", "min-single,auto", "auto,min,auto", "auto,min-single,auto", "auto-wrap,min", "auto-wrap,min-single" })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = top, BottomPoints = bottom };
            DocxTableCellLayout legacy = CreateParagraphMinimumLayout(pattern, measurer, "header", 320d, "exact", verticalAlignment: alignment, cellMargins: margins)
                .Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            DocxTableCellLayout actual = CreateParagraphMinimumLayout(pattern, measurer, rowHeight: 320d, rowRule: "exact", verticalAlignment: alignment, cellMargins: margins)
                .Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            TestAssert.Equal(legacy.TextLines.Count, actual.TextLines.Count);
            double sharedShift = alignment == "center" ? 5.22d + bottom / 2d : 10.44d + bottom;
            for (int i = 0; i < actual.TextLines.Count; i++)
            {
                double expected = sharedShift + (actual.TextLines[i].LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing ? 10d : 0d);
                TestAssert.True(Math.Abs(legacy.TextLines[i].BaselineY - actual.TextLines[i].BaselineY - expected) < .000001d,
                    "Independent Word margin probes align the paragraph-slot extent without reserving the bottom margin; the top margin still sets its origin.");
            }
        }
    }

    public static void DocxTableMarginMinimumPlacementUsesWordRowAdvanceAndFollowingFlow()
    {
        var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach (string? rule in new string?[] { "exact", "atLeast", null })
        foreach ((double top, double bottom) in new[] { (6d, 0d), (0d, 12d), (12d, 6d), (12d, 12d) })
        {
            var margins = DocxTableCellMargins.Empty with { TopPoints = top, BottomPoints = bottom };
            DocxLayout legacy = CreateParagraphMinimumLayout("auto,min,auto", measurer, "header", 320d, rule, verticalAlignment: alignment, cellMargins: margins);
            DocxLayout actual = CreateParagraphMinimumLayout("auto,min,auto", measurer, rowHeight: 320d, rowRule: rule, verticalAlignment: alignment, cellMargins: margins);
            DocxTableRowLayout before = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            DocxTableRowLayout after = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.Equal(legacy.Pages.Count, actual.Pages.Count);
            TestAssert.True(Math.Abs(before.Height + bottom - after.Height) < .000001d,
                "Word includes the maximum bottom cell margin in declared-row advancement.");
            TestAssert.True(Math.Abs(before.Y - bottom - after.Y) < .000001d,
                "The bottom margin extends the row downward from its stable top.");
            TestAssert.Equal(legacy.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY - bottom,
                actual.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY);
        }
    }

    public static void DocxTableMarginMinimumPlacementRetainsUnqualifiedAndComplexCells()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
        foreach (string alignment in new[] { "center", "bottom" })
        foreach ((string variant, double? height, string? rule) in new (string, double?, string?)[]
        {
            ("rich", null, null), ("mixed", 320d, "auto"), ("", 0d, "exact"), ("", 320d, "unknown"),
            ("rich", 320d, "exact"), ("mixed", 320d, "atLeast"), ("header", 320d, "exact")
        })
        {
            DocxLayout layout = CreateParagraphMinimumLayout("min,min,min", full, variant, height, rule, verticalAlignment: alignment, cellMargins: margins);
            TestAssert.True(layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells)
                .SelectMany(c => c.TextLines).All(l => l.BodyLineBoxBaselineInsetPoints is null),
                "Positive margins do not admit invalid rows, rich automatic cells or repeated headers.");
        }
        DocxLayout fallback = CreateParagraphMinimumLayout("min,min,min", new MinimumMetricsFallback(full), rowHeight: 320d, rowRule: "exact", verticalAlignment: "bottom", cellMargins: margins);
        TestAssert.True(fallback.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single().TextLines.All(l => l.BodyLineBoxBaselineInsetPoints is null),
            "A measurer without full vertical metrics retains the prior margin model.");
    }

    public static void DocxTableMarginMinimumPlacementRetainsContentDominatedAndPageBoundaryGeometry()
    {
        var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        var margins = DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d };
        foreach ((double height, string rule, double pageHeight) in new[] { (80d, "atLeast", 700d), (320d, "exact", 260d) })
        {
            DocxLayout layout = CreateParagraphMinimumLayout("min,min,min", measurer, rowHeight: height, rowRule: rule,
                verticalAlignment: "bottom", cellMargins: margins, pageHeight: pageHeight);
            TestAssert.True(layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells)
                .SelectMany(c => c.TextLines).All(l => l.BodyLineBoxBaselineInsetPoints is null),
                "Content-dominated minimum rows and tables that do not fit the body frame retain the earlier margin geometry.");
        }
        double naturalHeight = CreateParagraphMinimumLayout("min,min,min", measurer, rowHeight: 80d, rowRule: "atLeast",
            verticalAlignment: "bottom", cellMargins: margins).Pages[0].Items.OfType<DocxTableRowLayout>().Single().Height;
        DocxLayout coincidentHeight = CreateParagraphMinimumLayout("min,min,min", measurer,
            rowHeight: naturalHeight - 6d - 12d, rowRule: "atLeast", verticalAlignment: "bottom", cellMargins: margins);
        TestAssert.True(coincidentHeight.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single().TextLines
            .All(l => l.BodyLineBoxBaselineInsetPoints is null),
            "An equal final height does not admit a margin advance that row measurement rejected as content-dominated.");
    }

    public static void DocxTableAlignedMinimumPlacementMatchesIndependentWordSlots()
    {
        var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach (string pattern in new[] { "min,auto", "min-single,auto", "auto,min,auto", "auto,min-single,auto", "auto-wrap,min", "auto-wrap,min-single" })
        {
            DocxTableCellLayout legacy = CreateParagraphMinimumLayout(pattern, measurer, "header", 320d, "exact", verticalAlignment: alignment)
                .Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            DocxTableCellLayout actual = CreateParagraphMinimumLayout(pattern, measurer, rowHeight: 320d, rowRule: "exact", verticalAlignment: alignment)
                .Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            TestAssert.Equal(legacy.TextLines.Count, actual.TextLines.Count);
            double sharedShift = alignment == "center" ? 5.22d : 10.44d;
            for (int i = 0; i < actual.TextLines.Count; i++)
            {
                double expected = sharedShift + (actual.TextLines[i].LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing ? 10d : 0d);
                TestAssert.True(Math.Abs(legacy.TextLines[i].BaselineY - actual.TextLines[i].BaselineY - expected) < .000001d,
                    "Word aligns the full paragraph slot extent and puts excess minimum height above its baselines; ordinary in-cell paragraphs share only the alignment correction.");
            }
        }
    }

    public static void DocxTableAlignedMinimumPlacementPreservesRowsAndFollowingFlow()
    {
        var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach (string? rule in new string?[] { "exact", "atLeast", null })
        {
            DocxLayout legacy = CreateParagraphMinimumLayout("auto,min,auto", measurer, "header", 320d, rule, verticalAlignment: alignment);
            DocxLayout actual = CreateParagraphMinimumLayout("auto,min,auto", measurer, rowHeight: 320d, rowRule: rule, verticalAlignment: alignment);
            DocxTableRowLayout before = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            DocxTableRowLayout after = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.Equal(legacy.Pages.Count, actual.Pages.Count);
            TestAssert.Equal(before.Y, after.Y);
            TestAssert.Equal(before.Height, after.Height);
            TestAssert.Equal(legacy.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY,
                actual.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY);
        }
    }

    public static void DocxTableAlignedMinimumPlacementRetainsUnqualifiedAndComplexCells()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "center", "bottom" })
        foreach ((string variant, double? height, string? rule) in new (string, double?, string?)[]
        {
            ("rich", null, null), ("mixed", 320d, "auto"), ("", 0d, "exact"), ("", 320d, "unknown"),
            ("rich", 320d, "exact"), ("mixed", 320d, "atLeast"), ("header", 320d, "exact"), ("center-rich", 320d, "auto")
        })
        {
            DocxLayout layout = CreateParagraphMinimumLayout("min,min,min", full, variant, height, rule, verticalAlignment: alignment);
            TestAssert.True(layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells)
                .SelectMany(c => c.TextLines).All(l => l.BodyLineBoxBaselineInsetPoints is null),
                "Rich automatic cells, invalid rules and repeated headers retain their earlier aligned placement.");
        }
    }

    public static void DocxTableMixedMinimumPlacementUsesEachPlainParagraphSlot()
    {
        var measurer = new ParagraphMinimumTextMeasurer();
        foreach (var fonts in new (string Family, double Size, bool Bold, bool Italic)[][]
        {
            [("LoAsc", 12d, false, false), ("HiAsc", 12d, false, false), ("LoAsc", 12d, false, false)],
            [("LoAsc", 10d, false, false), ("LoAsc", 16d, false, false), ("LoAsc", 12d, false, false)],
            [("HiAsc", 16d, false, false), ("LoAsc", 10d, false, false), ("HiAsc", 12d, false, false)],
            [("LoAsc", 12d, false, false), ("LoAsc", 12d, true, true), ("LoAsc", 12d, false, false)]
        })
        foreach (string pattern in new[] { "min,auto", "min-single,auto", "auto,min,auto", "auto,min-single,auto", "auto-wrap,min", "auto-wrap,min-single", "min,min,min" })
        {
            DocxTableCellLayout legacy = CreateParagraphMinimumLayout(pattern, measurer, "header", paragraphFonts: fonts)
                .Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            DocxTableCellLayout actual = CreateParagraphMinimumLayout(pattern, measurer, paragraphFonts: fonts)
                .Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            TestAssert.Equal(legacy.TextLines.Count, actual.TextLines.Count);
            for (int i = 0; i < actual.TextLines.Count; i++)
            {
                DocxTextLineLayout line = actual.TextLines[i];
                DocxTextRun run = line.SourceParagraph!.Runs[0];
                double naturalHeight = run.FontFamily == "HiAsc" ? run.FontSize * 1.5d : run.FontSize * 1.2d;
                double expectedShift = line.LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing ? Math.Max(0d, 24d - naturalHeight) : 0d;
                TestAssert.True(Math.Abs(legacy.TextLines[i].BaselineY - line.BaselineY - expectedShift) < .000001d,
                    "Independent Word controls put each paragraph's excess minimum height above its baseline, across family, size and style changes between paragraphs.");
            }
        }
    }

    public static void DocxTableMixedMinimumPlacementPreservesRowsAndOrdinaryNeighbors()
    {
        var measurer = new ParagraphMinimumTextMeasurer();
        (string Family, double Size, bool Bold, bool Italic)[] fonts = [("LoAsc", 10d, false, false), ("HiAsc", 16d, true, false), ("LoAsc", 12d, false, true)];
        foreach (string? rule in new string?[] { null, "exact", "atLeast", "auto" })
        {
            DocxLayout legacy = CreateParagraphMinimumLayout("auto,min,auto", measurer, "header", 320d, rule, fonts);
            DocxLayout actual = CreateParagraphMinimumLayout("auto,min,auto", measurer, rowHeight: 320d, rowRule: rule, paragraphFonts: fonts);
            DocxTableRowLayout before = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            DocxTableRowLayout after = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.Equal(legacy.Pages.Count, actual.Pages.Count);
            TestAssert.Equal(before.Y, after.Y);
            TestAssert.Equal(before.Height, after.Height);
            DocxTextLineLayout[] ordinaryBefore = before.Cells.Single().TextLines.Where(l => l.LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing).ToArray();
            DocxTextLineLayout[] ordinaryAfter = after.Cells.Single().TextLines.Where(l => l.LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing).ToArray();
            TestAssert.Equal(ordinaryBefore.Length, ordinaryAfter.Length);
            for (int i = 0; i < ordinaryBefore.Length; i++)
            {
                TestAssert.Equal(ordinaryBefore[i].BaselineY, ordinaryAfter[i].BaselineY);
                TestAssert.Equal(ordinaryBefore[i].X, ordinaryAfter[i].X);
                TestAssert.Equal(ordinaryBefore[i].FontSize, ordinaryAfter[i].FontSize);
                TestAssert.Equal(ordinaryBefore[i].LineHeight, ordinaryAfter[i].LineHeight);
                TestAssert.Equal(ordinaryBefore[i].Text, ordinaryAfter[i].Text);
            }
            TestAssert.Equal(legacy.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY,
                actual.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY);
        }
    }

    public static void DocxTableMixedMinimumPlacementRetainsRichRunsAndComplexCells()
    {
        var measurer = new ParagraphMinimumTextMeasurer();
        (string Family, double Size, bool Bold, bool Italic)[] fonts = [("LoAsc", 10d, false, false), ("HiAsc", 16d, false, false), ("LoAsc", 12d, false, false)];
        foreach (string variant in new[] { "mixed", "header", "center-rich", "rich" })
        {
            DocxLayout layout = CreateParagraphMinimumLayout("min,min,min", measurer, variant, paragraphFonts: fonts);
            TestAssert.True(layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells)
                .SelectMany(c => c.TextLines).All(l => l.BodyLineBoxBaselineInsetPoints is null),
                "Mixed runs inside a paragraph and complex cells retain their qualified fallback.");
        }
    }

    private sealed class ParagraphMinimumTextMeasurer : IDocxTextMeasurer, IDocxLineMetricsProvider,
        IDocxStaticTextMetricsProvider, IDocxHheaDescenderProvider, IDocxHheaLineGapProvider
    {
        private readonly CellBoxTextMeasurer metrics = new(new DocxTests.FamilyWidthTextMeasurer());
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => metrics.MeasureText(run, text, fontSize);
        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => run?.FontFamily == "HiAsc" ? fontSize * 1.5d : fontSize * 1.2d;
        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => MeasureSingleLineHeight(run, fontSize);
        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => metrics.MeasureHheaAscender(run, fontSize);
        public double MeasureHheaDescender(DocxTextRun? run, double fontSize) => metrics.MeasureHheaDescender(run, fontSize);
        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => metrics.MeasureHheaLineGap(run, fontSize);
        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => metrics.MeasureWindowsAscender(run, fontSize);
        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => metrics.MeasureWindowsDescender(run, fontSize);
    }

    public static void DocxTableDeclaredMinimumPlacementMatchesIndependentWordSlots()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string? rule in new string?[] { "exact", "atLeast", null, "auto" })
        foreach (string pattern in new[] { "min,auto", "min-single,auto", "auto,min,auto", "auto,min-single,auto", "auto-wrap,min", "auto-wrap,min-single" })
        {
            DocxTableCellLayout legacy = CreateParagraphMinimumLayout(pattern, full, "header",
                rowHeight: 320d, rowRule: rule).Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            DocxTableCellLayout actual = CreateParagraphMinimumLayout(pattern, full,
                rowHeight: 320d, rowRule: rule).Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            TestAssert.Equal(legacy.TextLines.Count, actual.TextLines.Count);
            for (int i = 0; i < actual.TextLines.Count; i++)
            {
                DocxTextLineLayout line = actual.TextLines[i];
                TestAssert.True(Math.Abs(legacy.TextLines[i].BaselineY - line.BaselineY -
                    (line.LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing ? 10d : 0d)) < .000001d,
                    "Word uses the same 24-minus-14pt excess above minimum paragraph baselines in declared exact/minimum/automatic rows.");
            }
        }
    }

    public static void DocxTableDeclaredMinimumPlacementPreservesRowsAndFollowingFlow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string? rule in new string?[] { "exact", "atLeast", null, "auto" })
        {
            DocxLayout legacy = CreateParagraphMinimumLayout("auto,min,auto", full, "header", rowHeight: 320d, rowRule: rule);
            DocxLayout actual = CreateParagraphMinimumLayout("auto,min,auto", full, rowHeight: 320d, rowRule: rule);
            TestAssert.Equal(legacy.Pages.Count, actual.Pages.Count);
            DocxTableRowLayout before = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            DocxTableRowLayout after = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.Equal(before.Y, after.Y);
            TestAssert.Equal(before.Height, after.Height);
            TestAssert.Equal(legacy.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY,
                actual.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY);
        }
    }

    public static void DocxTableDeclaredMinimumPlacementRetainsInvalidAndComplexRows()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach ((string variant, double height, string? rule) in new (string, double, string?)[]
        {
            ("", 0d, "exact"), ("", 0d, "atLeast"), ("", -1d, "exact"), ("", 320d, "unknown"),
            ("header", 320d, "exact"), ("center-rich", 320d, "atLeast"), ("rich", 320d, null), ("mixed", 320d, "exact")
        })
        {
            DocxLayout layout = CreateParagraphMinimumLayout("min,min,min", full, variant, rowHeight: height, rowRule: rule);
            TestAssert.True(layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells)
                .SelectMany(c => c.TextLines).All(l => l.BodyLineBoxBaselineInsetPoints is null),
                "Nonpositive exact/minimum rows, unrecognized height rules, headers and complex cells retain their earlier baseline model.");
        }
    }

    public static void DocxTableParagraphMinimumSingleLinesFollowIndependentWordSlots()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string pattern in new[] { "min-single", "min-single,auto", "auto,min-single", "auto-wrap,min-single", "min-single,exact", "auto,min-single,auto" })
        {
            DocxLayout actual = CreateParagraphMinimumLayout(pattern, full);
            DocxLayout legacy = CreateParagraphMinimumLayout(pattern, full, "legacy-header");
            DocxTableRowLayout expected = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            DocxTableRowLayout corrected = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.Equal(expected.Y, corrected.Y);
            TestAssert.Equal(expected.Height, corrected.Height);
            DocxTextLineLayout[] before = expected.Cells.Single().TextLines.ToArray();
            DocxTextLineLayout[] after = corrected.Cells.Single().TextLines.ToArray();
            for (int i = 0; i < after.Length; i++)
                TestAssert.True(Math.Abs(before[i].BaselineY - after[i].BaselineY -
                    (after[i].LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing ? 10d : 0d)) < .000001d,
                    "Independent Word single lines use the same excess-above-baseline slot and retain their ordinary/exact neighbors.");
        }
    }

    public static void DocxTableParagraphMinimumPlacementMatchesWordPositionsAndTransitions()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string pattern in new[] { "min,auto", "min,exact", "min,min", "auto-wrap,min", "auto,min,auto", "min,min,min", "min,exact-wrap,min", "min,auto-wrap,min" })
        {
            DocxTableCellLayout legacy = CreateParagraphMinimumLayout(pattern, full, "legacy-header").Pages[0]
                .Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            DocxTableCellLayout actual = CreateParagraphMinimumLayout(pattern, full).Pages[0]
                .Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            TestAssert.Equal(legacy.TextLines.Count, actual.TextLines.Count);
            for (int i = 0; i < actual.TextLines.Count; i++)
            {
                DocxTextLineLayout line = actual.TextLines[i];
                bool minimum = line.LineHeightSource == DocxLineHeightSource.AtLeastLineSpacing &&
                    actual.TextLines.Count(l => ReferenceEquals(l.SourceParagraph, line.SourceParagraph)) > 1;
                TestAssert.True(Math.Abs(legacy.TextLines[i].BaselineY - line.BaselineY - (minimum ? 10d : 0d)) < .000001d,
                    "Word puts the 24-minus-14pt excess above each wrapped minimum paragraph, preserving automatic/exact neighbors.");
            }
        }
    }

    public static void DocxTableParagraphMinimumPlacementPreservesRowsAndFollowingFlow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string pattern in new[] { "min,auto", "min,exact", "min,min", "auto,min,auto", "min,exact-wrap,min" })
        {
            DocxLayout actual = CreateParagraphMinimumLayout(pattern, full);
            DocxLayout legacy = CreateParagraphMinimumLayout(pattern, full, "legacy-header");
            DocxTableRowLayout expectedRow = legacy.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            DocxTableRowLayout actualRow = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.Equal(expectedRow.Y, actualRow.Y);
            TestAssert.Equal(expectedRow.Height, actualRow.Height);
            TestAssert.Equal(legacy.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY,
                actual.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY);
        }
    }

    public static void DocxTableParagraphMinimumPlacementRetainsExcludedCellsAndRows()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string variant in new[] { "header", "center-rich", "rich", "mixed" })
        {
            DocxLayout actual = CreateParagraphMinimumLayout("min,min,min", full, variant);
            TestAssert.True(actual.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells)
                .SelectMany(c => c.TextLines).All(l => l.BodyLineBoxBaselineInsetPoints is null),
                "The broader paragraph-position rule retains header rows and complex cells.");
        }
    }

    public static void DocxTableMinimumPicturesMoveWithTheirTextAndLineFrames()
    {
        var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string alignment in new[] { "top", "center", "bottom" })
        foreach (double minimum in new[] { 12d, 18d, 24d })
        foreach (double pictureHeight in new[] { 4d, 18d, 48d })
        foreach (string? rule in new string?[] { null, "exact", "atLeast", "auto" })
        {
            double? height = rule is "exact" or "atLeast" ? 180d : null;
            DocxLayout fallback = CreateParagraphMinimumLayout("min-single", measurer, "header", height, rule,
                verticalAlignment: alignment, minimumSpacing: minimum, pictureHeight: pictureHeight);
            DocxLayout actual = CreateParagraphMinimumLayout("min-single", measurer, rowHeight: height, rowRule: rule,
                verticalAlignment: alignment, minimumSpacing: minimum, pictureHeight: pictureHeight);
            DocxTableRowLayout oldRow = fallback.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            DocxTableRowLayout newRow = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.True(Math.Abs(oldRow.Y + oldRow.Height - newRow.Y - newRow.Height) < .000001d,
                "Minimum picture measurement preserves the row's top origin.");
            if (rule is "exact" or "atLeast" || pictureHeight == 4d)
                TestAssert.Equal(oldRow.Height, newRow.Height);
            else
                TestAssert.True(newRow.Height <= oldRow.Height,
                    "Automatic rows count overlapping minimum and picture headroom once.");
            var oldCell = oldRow.Cells.Single();
            var newCell = newRow.Cells.Single();
            var oldLine = oldCell.TextLines.Single();
            var newLine = newCell.TextLines.Single();
            var oldImage = oldCell.InlineImages.Single();
            var newImage = newCell.InlineImages.Single();
            if (minimum > 14d)
                TestAssert.True(newLine.BodyLineBoxBaselineInsetPoints is not null,
                    "A resolved picture paragraph receives the independently qualified minimum-slot correction.");
            if (alignment == "top" && minimum > 14d && pictureHeight is 4d or 48d)
                TestAssert.True(Math.Abs(oldLine.BaselineY - newLine.BaselineY -
                    (pictureHeight == 4d ? minimum - 14d : 0d)) < .000001d,
                    "Small pictures retain minimum headroom; tall pictures already occupy it.");
            TestAssert.True(Math.Abs(newImage.Y - newLine.BaselineY) < .000001d,
                "A minimum-leading inline picture stays on its surrounding text baseline.");
            TestAssert.True(Math.Abs((oldImage.Y - newImage.Y) - (oldLine.BaselineY - newLine.BaselineY)) < .000001d,
                "Text and its source-affined picture move together.");
            TestAssert.Equal(oldImage.X, newImage.X);
            TestAssert.Equal(oldImage.Width, newImage.Width);
            TestAssert.Equal(oldImage.Height, newImage.Height);
            TestAssert.True(oldImage.LineFrame is { } oldFrame && newImage.LineFrame is { } newFrame &&
                Math.Abs((oldFrame.Bottom - newFrame.Bottom) - (oldImage.Y - newImage.Y)) < .000001d &&
                oldFrame.Height == newFrame.Height,
                "The picture's drawable line frame follows the same translation.");
            TestAssert.True(Math.Abs(actual.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY -
                fallback.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY -
                (oldRow.Height - newRow.Height)) < .000001d,
                "Following body flow tracks the actual row-height reduction.");
        }
    }

    public static void DocxTableMinimumPictureRowsUseTheOccupiedInlineExtent()
    {
        var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string? rule in new string?[] { null, "auto", "atLeast" })
        {
            var layout = CreateParagraphMinimumLayout("min-single", measurer,
                rowHeight: rule == "atLeast" ? 20d : null, rowRule: rule,
                minimumSpacing: 24d, pictureHeight: 48d);
            var row = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
            // The source picture occupies48pt above its baseline. The qualified
            // face contributes14pt natural height,10.44pt ascent, plus6pt after.
            TestAssert.True(Math.Abs(row.Height - 57.56d) < .000001d,
                "An automatic/content-dominated row includes the picture, descender and paragraph after-gap once.");
            TestAssert.True(Math.Abs(row.Cells.Single().InlineImages.Single().Y -
                row.Cells.Single().TextLines.Single().BaselineY) < .000001d,
                "Measured row extent retains text/picture baseline ownership.");
        }
    }

    public static void DocxKeptMinimumPictureRowUsesTheAdmittedOccupiedFrame()
    {
        foreach ((double available, int page) in new[] { (68d, 2), (76d, 1), (90d, 1) })
        {
            var layout = CreateKeptMinimumPictureLayout(available);
            int headingPage = Array.FindIndex(layout.Pages.ToArray(), p => p.Items.OfType<DocxTextLineLayout>()
                .Any(line => line.Text == "Keep")) + 1;
            int rowPage = Array.FindIndex(layout.Pages.ToArray(), p => p.Items.OfType<DocxTableRowLayout>().Any()) + 1;
            TestAssert.Equal(page, headingPage);
            TestAssert.Equal(page, rowPage);
        }
    }

    public static void DocxKeptMinimumPictureRowRetainsIncompleteFrameFallbacks()
    {
        foreach (string variant in new[] { "header", "rich", "two-rows" })
        {
            var layout = CreateKeptMinimumPictureLayout(76d, variant);
            TestAssert.True(!layout.Pages[0].Items.OfType<DocxTableRowLayout>().Any(),
                "Unqualified or incomplete table frames preserve the existing kept-row fallback: " + variant);
            TestAssert.True(!layout.Pages[0].Items.OfType<DocxTextLineLayout>().Any(line => line.Text == "Keep"),
                "The heading stays with the first row under fallback admission.");
        }
    }

    private static DocxLayout CreateKeptMinimumPictureLayout(double available, string variant = "")
    {
        var metrics = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Intro", 12d, 12d) with
        {
            SpacingBeforePoints = 0d, SpacingAfterPoints = 180d - 12d - available,
            Spacing = DocxParagraphSpacing.Empty with { LineValue = "240", LineRuleValue = "exact" }
        };
        DocxParagraph heading = DocxTests.CreateDocxLayoutParagraph("Keep", 12d, 14d) with
        {
            LineSpacingPoints = null, SpacingBeforePoints = 0d, SpacingAfterPoints = 0d,
            KeepRules = DocxParagraphKeepRules.Empty with { KeepNext = true }
        };
        DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Short" + (variant == "rich" ? "\n" : ""), 12d, 24d) with
        {
            SpacingBeforePoints = 0d, SpacingAfterPoints = 6d,
            Spacing = DocxParagraphSpacing.Empty with { LineValue = "480", LineRuleValue = "atLeast" },
            Images = [new DocxInlineImage(18d, 48d, "image/png", [], null) { SourceRunIndex = 1 }]
        };
        var cell = new DocxTableCell("", [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty)
            { BodyElements = [new DocxParagraphElement(paragraph)] };
        var row = new DocxTableRow([cell], null) { IsHeader = variant == "header" };
        var table = new DocxTable(null, [120d], variant == "two-rows" ? [row, row] : [row]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument(
            [new DocxParagraphElement(filler), new DocxParagraphElement(heading), new DocxTableElement(table)], [table]);
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, metrics, CancellationToken.None);
    }

    public static void DocxTableMinimumPicturesRetainIncompleteAndComplexOwnershipFallbacks()
    {
        var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string variant in new[] { "header", "rich", "mixed", "block-picture" })
        {
            var layout = CreateParagraphMinimumLayout("min-single,min-single", measurer, variant,
                rowHeight: 180d, rowRule: "exact", verticalAlignment: "bottom", pictureHeight: 18d);
            var cell = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
            TestAssert.True(cell.TextLines.All(l => l.BodyLineBoxBaselineInsetPoints is null),
                "Unqualified cell content retains the existing minimum placement fallback: " + variant);
            TestAssert.Equal(2, cell.InlineImages.Count);
        }
    }

    public static void DocxTableMinimumPicturesPreserveWrappedLineAffinity()
    {
        var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        var layout = CreateParagraphMinimumLayout("min", measurer,
            rowHeight: 180d, rowRule: "exact", minimumSpacing: 24d, pictureHeight: 18d);
        var cell = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single();
        var image = cell.InlineImages.Single();
        TestAssert.True(cell.TextLines.Count > 1, "The control must wrap.");
        TestAssert.True(cell.TextLines.Any(line => Math.Abs(line.BaselineY - image.Y) < .000001d),
            "The picture stays on its own wrapped line after minimum placement.");
        for (int i = 1; i < cell.TextLines.Count; i++)
            TestAssert.True(cell.TextLines[i - 1].BaselineY - cell.TextLines[i].BaselineY >= 24d - .000001d,
                "The authored minimum pitch remains a floor across picture-driven wrapped lines.");
    }

    private static DocxLayout CreateParagraphMinimumLayout(string pattern, IDocxTextMeasurer measurer, string variant = "",
        double? rowHeight = null, string? rowRule = null,
        IReadOnlyList<(string Family, double Size, bool Bold, bool Italic)>? paragraphFonts = null,
        string? verticalAlignment = null, DocxTableCellMargins? cellMargins = null, double pageHeight = 700d, int rowCount = 1, int columnCount = 1, double minimumSpacing = 24d,
        double? shortNeighborMinimum = null, DocxTableCellMargins? shortNeighborMargins = null, double exactSpacing = 24d, string? borderSize = null,
        IReadOnlyList<(double? Height, string? Rule)>? rowHeightRules = null, double? pictureHeight = null)
    {
        DocxParagraph[] paragraphs = pattern.Split(',').Select((token, index) =>
        {
            bool wrapped = token == "min" || token.EndsWith("-wrap", StringComparison.Ordinal);
            string rule = token.StartsWith("min", StringComparison.Ordinal) ? "atLeast" : token.StartsWith("exact", StringComparison.Ordinal) ? "exact" : "auto";
            double slot = rule == "atLeast" ? minimumSpacing : rule == "exact" ? exactSpacing : 24d;
            DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph(wrapped ? "Wrapped word Wrapped word Wrapped word Wrapped word" : "Short", 12d, 24d) with
            {
                LineSpacingPoints = rule == "auto" ? null : slot,
                SpacingBeforePoints = index == 0 ? 0d : 6d, SpacingAfterPoints = 6d,
                Spacing = new DocxParagraphSpacing(null, null, null, null, null, null, rule == "auto" ? "240" : ((int)(slot * 20d)).ToString(CultureInfo.InvariantCulture), rule, null)
            };
            var fonts = paragraphFonts?[index];
            paragraph = paragraph with { Runs = paragraph.Runs.Select(r => r with
            {
                FontFamily = fonts?.Family ?? (variant == "mixed" && index == 1 ? "HiAsc" : "LoAsc"),
                FontSize = fonts?.Size ?? 12d, Bold = fonts?.Bold ?? false, Italic = fonts?.Italic ?? false,
                Text = r.Text + (variant is "rich" or "center-rich" ? "\n" : string.Empty)
            }).ToArray() };
            if (variant == "mixed" && index == 1)
                paragraph = paragraph with { Runs = [.. paragraph.Runs, paragraph.Runs[0] with { FontFamily = "LoAsc", Text = " Mixed run" }] };
            if (pictureHeight is { } imageHeight)
                paragraph = paragraph with
                {
                    Runs = paragraph.Runs.Select(r => r with { SourceRunIndex = 0 }).ToArray(),
                    Images = [new DocxInlineImage(18d, imageHeight, "image/png", [], null)
                        { SourceRunIndex = variant == "block-picture" ? -1 : 1 }]
                };
            return paragraph;
        }).ToArray();
        var cell = new DocxTableCell(string.Empty, paragraphs, null, null, null, verticalAlignment ?? (variant is "center" or "center-rich" ? "center" : null), [],
            cellMargins ?? (variant == "margins" ? DocxTableCellMargins.Empty with { TopPoints = 6d, BottomPoints = 12d } : DocxTableCellMargins.Empty))
            { BodyElements = paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p)).ToArray() };
        if (borderSize is not null) cell = cell with { Borders = new[] { "top", "left", "bottom", "right" }
            .Select(edge => new DocxTableCellBorder(edge, "single", "000000", borderSize)).ToArray() };
        var cells = Enumerable.Repeat(cell, columnCount).ToArray();
        if (shortNeighborMinimum is { } shortMinimum && columnCount > 1)
        {
            var shortParagraph = paragraphs[0] with { LineSpacingPoints = shortMinimum,
                Spacing = paragraphs[0].Spacing with { LineValue = ((int)(shortMinimum * 20d)).ToString(CultureInfo.InvariantCulture) },
                Runs = paragraphs[0].Runs.Select(run => run with { Text = "Short" }).ToArray() };
            cells[1] = cell with { Paragraphs = [shortParagraph], BodyElements = [new DocxParagraphElement(shortParagraph)],
                Margins = shortNeighborMargins ?? cell.Margins };
        }
        DocxTableRow row = new(cells, rowHeight) { HeightRuleValue = rowRule };
        if (variant is "header" or "legacy-header") row = row with { IsHeader = true };
        var rows = Enumerable.Range(0, rowCount).Select(i => rowHeightRules is null ? row :
            row with { HeightPoints = rowHeightRules[i].Height, HeightRuleValue = rowHeightRules[i].Rule }).ToArray();
        DocxTable table = new(null, Enumerable.Repeat(120d / columnCount, columnCount).ToArray(), rows);
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 12d, 14d) with { LineSpacingPoints = null };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table), new DocxParagraphElement(after)], [table]) with
            { PageWidthPoints = 220d, PageHeightPoints = pageHeight };
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, CancellationToken.None);
    }

    public static void DocxTableGridMinimumLinePlacementMatchesIndependentWord()
    {
        foreach ((int rowCount, int columnCount) in new[] { (1, 2), (2, 1), (2, 2) })
        foreach (bool predecessor in new[] { false, true })
        foreach (double minimum in new[] { 18d, 24d, 36d })
        {
            var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
            DocxTableRowLayout[] automatic = CreateMinimumTableLayout(null, predecessor, measurer,
                rowCount: rowCount, columnCount: columnCount).Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            DocxTableRowLayout[] actual = CreateMinimumTableLayout(minimum, predecessor, measurer,
                rowCount: rowCount, columnCount: columnCount).Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            TestAssert.Equal(rowCount, actual.Length);
            for (int r = 0; r < rowCount; r++)
            for (int c = 0; c < columnCount; c++)
            {
                DocxTextLineLayout firstAutomatic = automatic[r].Cells[c].TextLines.First(l => l.Text.StartsWith("Current", StringComparison.Ordinal));
                DocxTextLineLayout[] lines = actual[r].Cells[c].TextLines.Where(l => l.SourceParagraph?.LineSpacingPoints == minimum).ToArray();
                double automaticInset = automatic[r].Y + automatic[r].Height - firstAutomatic.BaselineY;
                double actualInset = actual[r].Y + actual[r].Height - lines[0].BaselineY;
                TestAssert.True(Math.Abs(actualInset - automaticInset - (minimum - 14d)) < .000001d,
                    "Independent Word grids put excess minimum height above each eligible first baseline, relative to its row top.");
                for (int i = 1; i < lines.Length; i++)
                    TestAssert.True(Math.Abs(lines[i - 1].BaselineY - lines[i].BaselineY - minimum) < .001d,
                        "The correction retains minimum-height advances in every eligible cell.");
            }
        }
    }

    public static void DocxTableGridMinimumLinePlacementPreservesRowsAndFlow()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach ((int rowCount, int columnCount) in new[] { (1, 2), (2, 1), (2, 2) })
        {
            DocxLayout corrected = CreateMinimumTableLayout(24d, true, full, rowCount: rowCount, columnCount: columnCount);
            DocxLayout fallback = CreateMinimumTableLayout(24d, true, new MinimumMetricsFallback(full), rowCount: rowCount, columnCount: columnCount);
            TestAssert.Equal(fallback.Pages.Count, corrected.Pages.Count);
            DocxTableRowLayout[] expected = fallback.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            DocxTableRowLayout[] actual = corrected.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
            for (int i = 0; i < expected.Length; i++)
            {
                TestAssert.Equal(expected[i].Y, actual[i].Y);
                TestAssert.Equal(expected[i].Height, actual[i].Height);
            }
            TestAssert.Equal(fallback.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY,
                corrected.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY);
        }
    }

    public static void DocxTableGridMinimumLinePlacementRetainsComplexAndUnrecognizedRows()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string variant in new[] { "center-rich", "rich", "mixed", "unknown-height", "header" })
        {
            DocxLayout layout = CreateMinimumTableLayout(24d, true, full, variant, rowCount: 2, columnCount: 2);
            TestAssert.True(layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells)
                .SelectMany(c => c.TextLines).All(l => l.BodyLineBoxBaselineInsetPoints is null),
                "Complex cells, unrecognized row height rules and header rows retain their previous baseline model.");
        }
    }

    public static void DocxTableMinimumLinePlacementMatchesIndependentWordExcess()
    {
        foreach (bool predecessor in new[] { false, true })
        foreach (double minimum in new[] { 12d, 14d, 18d, 24d, 36d })
        {
            var measurer = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
            DocxLayout automatic = CreateMinimumTableLayout(null, predecessor, measurer);
            DocxLayout actual = CreateMinimumTableLayout(minimum, predecessor, measurer);
            DocxTextLineLayout[] baseline = automatic.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single().TextLines
                .Where(l => l.Text.StartsWith("Current", StringComparison.Ordinal)).ToArray();
            DocxTextLineLayout[] lines = actual.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single().TextLines
                .Where(l => l.SourceParagraph?.LineSpacingPoints == minimum).ToArray();
            TestAssert.True(lines.Length > 1, "The independent Word admission covers wrapped paragraphs.");
            TestAssert.True(Math.Abs(baseline[0].BaselineY - lines[0].BaselineY - Math.Max(0d, minimum - 14d)) < .000001d,
                "Word places excess minimum-height space above the first baseline, in first and following cell paragraphs.");
            for (int i = 1; i < lines.Length; i++)
                TestAssert.True(Math.Abs(lines[i - 1].BaselineY - lines[i].BaselineY - Math.Max(minimum, 14d)) < .001d,
                    "Later advances retain the effective minimum or natural line height.");
        }
    }

    public static void DocxTableMinimumLinePlacementPreservesRowHeightAndFollowingFlow()
    {
        foreach (double minimum in new[] { 18d, 24d, 36d })
        {
            var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
            DocxLayout corrected = CreateMinimumTableLayout(minimum, true, full);
            DocxLayout fallback = CreateMinimumTableLayout(minimum, true, new MinimumMetricsFallback(full));
            TestAssert.Equal(fallback.Pages.Count, corrected.Pages.Count);
            TestAssert.Equal(fallback.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Height,
                corrected.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Height);
            TestAssert.Equal(fallback.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY,
                corrected.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY);
        }
    }

    public static void DocxTableMinimumLinePlacementRetainsUnsupportedCellsAndMetrics()
    {
        var full = new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer());
        foreach (string variant in new[] { "center-rich", "rich", "mixed" })
        {
            DocxLayout layout = CreateMinimumTableLayout(24d, true, full, variant);
            TestAssert.True(layout.Pages.SelectMany(p => p.Items.OfType<DocxTableRowLayout>()).SelectMany(r => r.Cells)
                .SelectMany(c => c.TextLines).All(l => l.BodyLineBoxBaselineInsetPoints is null),
                "Rich aligned and mixed text retain the original cell baseline model.");
        }
        DocxLayout providerless = CreateMinimumTableLayout(24d, true, new MinimumMetricsFallback(full));
        TestAssert.True(providerless.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single().TextLines
            .All(l => l.BodyLineBoxBaselineInsetPoints is null), "Partial metric providers keep the legacy characterization.");
    }

    private static DocxLayout CreateMinimumTableLayout(double? minimum, bool predecessor,
        IDocxTextMeasurer measurer, string variant = "", int rowCount = 1, int columnCount = 1)
    {
        DocxParagraph current = DocxTests.CreateDocxLayoutParagraph("Current word Current word Current word Current word", 12d, minimum ?? 14d) with
        {
            LineSpacingPoints = minimum,
            SpacingBeforePoints = 6d, SpacingAfterPoints = 24d,
            Spacing = new DocxParagraphSpacing(null, null, null, null, null, null, minimum is null ? "240" : ((int)(minimum * 20d)).ToString(CultureInfo.InvariantCulture),
                minimum is null ? "auto" : "atLeast", null)
        };
        current = current with { Runs = current.Runs.Select(r => r with { FontFamily = "LoAsc" }).ToArray() };
        if (variant is "rich" or "center-rich") current = current with { Runs = current.Runs.Select(r => r with { Text = r.Text + "\n" }).ToArray() };
        DocxParagraph previous = DocxTests.CreateDocxLayoutParagraph("Before", 12d, 14d) with { LineSpacingPoints = null, SpacingAfterPoints = 0d };
        previous = previous with { Runs = previous.Runs.Select(r => r with { FontFamily = variant == "mixed" ? "HiAsc" : "LoAsc" }).ToArray() };
        if (variant == "mixed")
            previous = previous with { Runs = [.. previous.Runs, previous.Runs[0] with { FontFamily = "LoAsc", Text = " Mixed run" }] };
        DocxParagraph[] paragraphs = predecessor ? [previous, current] : [current];
        var cell = new DocxTableCell(string.Empty, paragraphs, null, null, null, variant is "center" or "center-rich" ? "center" : null, [], DocxTableCellMargins.Empty)
        { BodyElements = paragraphs.Select(p => (DocxBodyElement)new DocxParagraphElement(p)).ToArray() };
        DocxTableRow row = new(Enumerable.Repeat(cell, columnCount).ToArray(), variant == "unknown-height" ? 200d : null);
        if (variant == "unknown-height") row = row with { HeightRuleValue = "unknown" };
        if (variant == "header") row = row with { IsHeader = true };
        DocxTable table = new(null, Enumerable.Repeat(120d, columnCount).ToArray(), Enumerable.Repeat(row, rowCount).ToArray());
        DocxParagraph after = DocxTests.CreateDocxLayoutParagraph("After", 12d, 14d) with { LineSpacingPoints = null };
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table), new DocxParagraphElement(after)], [table]) with
            { PageWidthPoints = columnCount * 120d + 100d, PageHeightPoints = 700d };
        return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, measurer, CancellationToken.None);
    }

    private sealed class MinimumMetricsFallback(CellBoxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double size) => inner.MeasureText(run, text, size);
        public double MeasureSingleLineHeight(DocxTextRun? run, double size) => inner.MeasureSingleLineHeight(run, size);
        public double MeasureHheaLineHeight(DocxTextRun? run, double size) => inner.MeasureHheaLineHeight(run, size);
        public double MeasureHheaAscender(DocxTextRun? run, double size) => inner.MeasureHheaAscender(run, size);
    }

    public static void DocxSyntheticThreeDTableBorderStylesRenderWithoutDiagnostic()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="threeDEmboss" w:color="336699" w:sz="16"/>
                              <w:bottom w:val="threeDEngrave" w:color="336699" w:sz="16"/>
                            </w:tcBorders>
                          </w:tcPr>
                          <w:p/>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_TABLE_BORDER_STYLE", ids);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Split(" re f", StringSplitOptions.None).Length - 1 >= 4, "3D table borders should render light/dark relief strips instead of flattening to one solid fill.");
        TestAssert.True(pdf.Split(" rg", StringSplitOptions.None).Length - 1 >= 4, "3D table borders should emit separate light and dark fill colors.");

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(2, table.VisibleBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
    }

    public static void DocxSyntheticWaveTableBorderStylesRenderWithoutDiagnostic()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="wave" w:color="336699" w:sz="16"/>
                              <w:bottom w:val="doubleWave" w:color="336699" w:sz="16"/>
                            </w:tcBorders>
                          </w:tcPr>
                          <w:p/>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_TABLE_BORDER_STYLE", ids);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(DocxTests.CountOccurrences(pdf, " l S") >= 8, "Wave table borders should render as repeated stroked wave segments.");

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(2, table.VisibleBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
    }

    public static void DocxSyntheticTripleTableBorderStyleRendersWithoutDiagnostic()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="triple" w:color="993333" w:sz="18"/>
                            </w:tcBorders>
                          </w:tcPr>
                          <w:p/>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_TABLE_BORDER_STYLE", ids);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.6 0.2 0.2 rg", pdf);
        TestAssert.True(pdf.Split(" re f", StringSplitOptions.None).Length - 1 >= 3, "Triple borders should render as three filled strips.");

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(1, table.VisibleBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
    }

    public static void DocxSyntheticCompoundTableBorderStylesRenderWithoutDiagnostic()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="thinThickSmallGap" w:color="225588" w:sz="24"/>
                              <w:bottom w:val="thickThinSmallGap" w:color="225588" w:sz="24"/>
                              <w:left w:val="thinThickThinSmallGap" w:color="225588" w:sz="24"/>
                            </w:tcBorders>
                          </w:tcPr>
                          <w:p/>
                        </w:tc>
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="thinThickMediumGap" w:color="225588" w:sz="24"/>
                              <w:bottom w:val="thickThinMediumGap" w:color="225588" w:sz="24"/>
                              <w:left w:val="thinThickThinMediumGap" w:color="225588" w:sz="24"/>
                            </w:tcBorders>
                          </w:tcPr>
                          <w:p/>
                        </w:tc>
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="thinThickLargeGap" w:color="225588" w:sz="24"/>
                              <w:bottom w:val="thickThinLargeGap" w:color="225588" w:sz="24"/>
                              <w:right w:val="thinThickThinLargeGap" w:color="225588" w:sz="24"/>
                            </w:tcBorders>
                          </w:tcPr>
                          <w:p/>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string ids = string.Join("|", diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal));
        TestAssert.DoesNotContain("DOCX_TABLE_BORDER_STYLE", ids);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Split(" re f", StringSplitOptions.None).Length - 1 >= 18, "Compound table borders should render as separate thin/thick filled strips.");

        using FileStream stream = File.OpenRead(input);
        DocxStructureTableSnapshot table = DocxStructureSnapshot.FromDocument(new DocxReader().Read(OoxPackage.Open(stream, CancellationToken.None), null, CancellationToken.None, OoxPdfDocxMarkupMode.Final)).Tables.Single();
        TestAssert.Equal(9, table.VisibleBorderCount);
        TestAssert.Equal(0, table.OtherBorderStyleCount);
    }

    public static void DocxReaderTablePreservesHeaderRowToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:trPr><w:tblHeader w:val="0"/></w:trPr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.True(document.Tables[0].Rows[0].IsHeader, "Expected implicit tblHeader to mark the row as repeating.");
        TestAssert.True(document.Tables[0].Rows[0].HeaderValue is null, "Expected implicit tblHeader to preserve a null source token.");
        TestAssert.True(!document.Tables[0].Rows[1].IsHeader, "Expected tblHeader w:val=\"0\" to disable repeating.");
        TestAssert.Equal("0", document.Tables[0].Rows[1].HeaderValue ?? string.Empty);
    }

    public static void DocxReaderTableCellPreservesVerticalAlignmentToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:vAlign w:val="top"/></w:tcPr><w:p><w:r><w:t>Top</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>Center</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Default</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxTableCell> cells = document.Tables[0].Rows[0].Cells;

        TestAssert.Equal("top", cells[0].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("center", cells[1].VerticalAlignmentValue ?? string.Empty);
        TestAssert.True(cells[2].VerticalAlignmentValue is null, "Expected missing cell vertical alignment to keep a null source token.");
    }

    public static void DocxReaderTableCellPreservesTextDirectionToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:textDirection w:val="tbRl"/></w:tcPr><w:p><w:r><w:t>Rotated</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Default</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxTableCell> cells = document.Tables[0].Rows[0].Cells;

        TestAssert.Equal("tbRl", cells[0].TextDirectionValue ?? string.Empty);
        TestAssert.True(cells[1].TextDirectionValue is null, "Expected missing cell text direction to keep a null source token.");
    }

    public static void DocxReaderTableCellPreservesNoWrapToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:noWrap/></w:tcPr><w:p><w:r><w:t>Kept</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:noWrap w:val="0"/></w:tcPr><w:p><w:r><w:t>Wrapped</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxTableCell> cells = document.Tables[0].Rows[0].Cells;

        TestAssert.True(cells[0].NoWrap, "Expected implicit w:noWrap to disable automatic cell text wrapping.");
        TestAssert.True(cells[0].NoWrapValue is null, "Expected implicit w:noWrap to preserve a null source token.");
        TestAssert.True(!cells[1].NoWrap, "Expected w:noWrap w:val=\"0\" to allow automatic cell text wrapping.");
        TestAssert.Equal("0", cells[1].NoWrapValue ?? string.Empty);
    }

    public static void DocxReaderTableCellPreservesFitTextToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:tcFitText/></w:tcPr><w:p><w:r><w:t>Fit</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:tcFitText w:val="0"/></w:tcPr><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxTableCell> cells = document.Tables[0].Rows[0].Cells;

        TestAssert.True(cells[0].FitText, "Expected implicit w:tcFitText to fit cell text to its extents.");
        TestAssert.True(cells[0].FitTextValue is null, "Expected implicit w:tcFitText to preserve a null source token.");
        TestAssert.True(!cells[1].FitText, "Expected w:tcFitText w:val=\"0\" to disable cell text fitting.");
        TestAssert.Equal("0", cells[1].FitTextValue ?? string.Empty);
    }

    public static void DocxReaderTableCellPreservesMarginTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc>
                        <w:tcPr>
                          <w:tcMar>
                            <w:top w:w="120" w:type="dxa"/>
                            <w:right w:w="180" w:type="dxa"/>
                            <w:bottom w:w="240" w:type="dxa"/>
                            <w:left w:w="300" w:type="dxa"/>
                          </w:tcMar>
                        </w:tcPr>
                        <w:p><w:r><w:t>Margins</w:t></w:r></w:p>
                      </w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxTableCellMargins margins = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Tables[0].Rows[0].Cells[0].Margins;

        TestAssert.Equal("120", margins.TopValue ?? string.Empty);
        TestAssert.Equal("180", margins.RightValue ?? string.Empty);
        TestAssert.Equal("240", margins.BottomValue ?? string.Empty);
        TestAssert.Equal("300", margins.LeftValue ?? string.Empty);
        TestAssert.Equal(6d, margins.TopPoints ?? 0d);
        TestAssert.Equal(15d, margins.LeftPoints ?? 0d);
    }

    public static void DocxReaderTableCellPreservesGridSpanToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="720"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Span</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTableCell cell = document.Tables[0].Rows[0].Cells[0];
        TestAssert.Equal(2, cell.GridSpan);
        TestAssert.Equal("2", cell.GridSpanValue ?? string.Empty);
    }

    public static void DocxReaderInfersMissingTableGridFromLogicalGridSpans()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tr>
                        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Span</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Tail</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxTable table = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Tables[0];

        TestAssert.True(!table.HasExplicitGrid, "Missing tblGrid should remain distinguishable from authored grid geometry.");
        TestAssert.Equal(3, table.ColumnWidthsPoints.Count);
    }

    public static void DocxReaderTableCellPreservesVerticalMergeToken()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="720"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Top</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>Bottom</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTableCell restart = document.Tables[0].Rows[0].Cells[0];
        DocxTableCell continuation = document.Tables[0].Rows[1].Cells[0];
        TestAssert.True(restart.HasVerticalMerge, "Reader should preserve vMerge restart presence.");
        TestAssert.Equal("restart", restart.VerticalMergeValue ?? string.Empty);
        TestAssert.True(continuation.HasVerticalMerge, "Reader should preserve val-less vMerge continuation presence.");
        TestAssert.True(continuation.VerticalMergeValue is null, "Val-less vMerge continuation should keep its missing value distinct from restart.");
    }

    public static void DocxReaderTableCellPreservesBorderTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc>
                          <w:tcPr>
                            <w:tcBorders>
                              <w:top w:val="single" w:color="112233" w:sz="12"/>
                              <w:bottom w:val="nil"/>
                              <w:start w:val="dashed" w:color="445566" w:sz="8"/>
                            </w:tcBorders>
                          </w:tcPr>
                          <w:p><w:r><w:t>Bordered</w:t></w:r></w:p>
                        </w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxTableCellBorder> borders = document.Tables[0].Rows[0].Cells[0].Borders;

        TestAssert.Equal(3, borders.Count);
        TestAssert.Equal("top", borders[0].Edge);
        TestAssert.Equal("single", borders[0].Value ?? string.Empty);
        TestAssert.Equal("112233", borders[0].Color ?? string.Empty);
        TestAssert.Equal("12", borders[0].SizeValue ?? string.Empty);
        TestAssert.Equal("bottom", borders[1].Edge);
        TestAssert.Equal("nil", borders[1].Value ?? string.Empty);
        TestAssert.Equal("start", borders[2].Edge);
        TestAssert.Equal("dashed", borders[2].Value ?? string.Empty);
    }

    public static void DocxReaderTableBordersApplyOuterAndInsideEdges()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblBorders>
                          <w:top w:val="single" w:color="111111" w:sz="8"/>
                          <w:bottom w:val="single" w:color="222222" w:sz="10"/>
                          <w:left w:val="single" w:color="333333" w:sz="12"/>
                          <w:right w:val="single" w:color="444444" w:sz="14"/>
                          <w:insideH w:val="single" w:color="555555" w:sz="16"/>
                          <w:insideV w:val="single" w:color="666666" w:sz="18"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        IReadOnlyList<DocxTableCellBorder> first = document.Tables[0].Rows[0].Cells[0].Borders;
        IReadOnlyList<DocxTableCellBorder> inner = document.Tables[0].Rows[1].Cells[1].Borders;
        TestAssert.Equal("111111", first.Single(border => border.Edge == "top").Color ?? string.Empty);
        TestAssert.Equal("333333", first.Single(border => border.Edge == "left").Color ?? string.Empty);
        TestAssert.Equal("555555", first.Single(border => border.Edge == "bottom").Color ?? string.Empty);
        TestAssert.Equal("666666", first.Single(border => border.Edge == "right").Color ?? string.Empty);
        TestAssert.Equal("222222", inner.Single(border => border.Edge == "bottom").Color ?? string.Empty);
        TestAssert.Equal("444444", inner.Single(border => border.Edge == "right").Color ?? string.Empty);
    }

    public static void DocxReaderTableStyleBordersApplyOuterAndInsideEdges()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="table" w:styleId="BorderedTable">
                    <w:tblPr>
                      <w:tblBorders>
                        <w:top w:val="single" w:color="111111" w:sz="8"/>
                        <w:bottom w:val="single" w:color="222222" w:sz="10"/>
                        <w:left w:val="single" w:color="333333" w:sz="12"/>
                        <w:right w:val="single" w:color="444444" w:sz="14"/>
                        <w:insideH w:val="single" w:color="555555" w:sz="16"/>
                        <w:insideV w:val="single" w:color="666666" w:sz="18"/>
                      </w:tblBorders>
                    </w:tblPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr><w:tblStyle w:val="BorderedTable"/></w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        IReadOnlyList<DocxTableCellBorder> first = document.Tables[0].Rows[0].Cells[0].Borders;
        IReadOnlyList<DocxTableCellBorder> inner = document.Tables[0].Rows[1].Cells[1].Borders;
        TestAssert.Equal("111111", first.Single(border => border.Edge == "top").Color ?? string.Empty);
        TestAssert.Equal("333333", first.Single(border => border.Edge == "left").Color ?? string.Empty);
        TestAssert.Equal("555555", first.Single(border => border.Edge == "bottom").Color ?? string.Empty);
        TestAssert.Equal("666666", first.Single(border => border.Edge == "right").Color ?? string.Empty);
        TestAssert.Equal("222222", inner.Single(border => border.Edge == "bottom").Color ?? string.Empty);
        TestAssert.Equal("444444", inner.Single(border => border.Edge == "right").Color ?? string.Empty);
    }

    public static void DocxReaderTableLogicalBordersApplyOuterEdgesInLeftToRightLayout()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblBorders>
                          <w:start w:val="single" w:color="123456" w:sz="8"/>
                          <w:end w:val="single" w:color="654321" w:sz="8"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxTableCellBorder> borders = document.Tables[0].Rows[0].Cells[0].Borders;

        TestAssert.Equal("123456", borders.Single(border => border.Edge == "left").Color ?? string.Empty);
        TestAssert.Equal("654321", borders.Single(border => border.Edge == "right").Color ?? string.Empty);
    }

    public static void DocxReaderDirectTableBordersOverrideTableStyleCellBordersPerEdge()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="table" w:styleId="StyledTable">
                    <w:tblPr>
                      <w:tblBorders>
                        <w:left w:val="single" w:color="333333" w:sz="12"/>
                        <w:bottom w:val="single" w:color="222222" w:sz="10"/>
                      </w:tblBorders>
                    </w:tblPr>
                    <w:tblStylePr w:type="firstRow">
                      <w:tcPr>
                        <w:tcBorders>
                          <w:top w:val="single" w:color="FF0000" w:sz="18"/>
                          <w:right w:val="single" w:color="AA0000" w:sz="18"/>
                        </w:tcBorders>
                      </w:tcPr>
                    </w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="StyledTable"/>
                        <w:tblBorders>
                          <w:top w:val="nil"/>
                          <w:right w:val="single" w:color="0000FF" w:sz="6"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        IReadOnlyList<DocxTableCellBorder> borders = document.Tables[0].Rows[0].Cells[0].Borders;
        TestAssert.Equal("nil", borders.Single(border => border.Edge == "top").Value ?? string.Empty);
        TestAssert.Equal("0000FF", borders.Single(border => border.Edge == "right").Color ?? string.Empty);
        TestAssert.Equal("333333", borders.Single(border => border.Edge == "left").Color ?? string.Empty);
        TestAssert.Equal("222222", borders.Single(border => border.Edge == "bottom").Color ?? string.Empty);
    }

    public static void DocxReaderTableStyleAppliesConditionalCellVerticalAlignment()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="table" w:styleId="AlignedTable">
                    <w:tcPr><w:vAlign w:val="center"/></w:tcPr>
                    <w:tblStylePr w:type="firstRow">
                      <w:tcPr><w:vAlign w:val="bottom"/></w:tcPr>
                    </w:tblStylePr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblStyle w:val="AlignedTable"/>
                        <w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Inherited</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:vAlign w:val="top"/></w:tcPr><w:p><w:r><w:t>Direct</w:t></w:r></w:p></w:tc>
                      </w:tr>
                      <w:tr>
                        <w:tc><w:p><w:r><w:t>Whole</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Whole2</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        IReadOnlyList<DocxTableRow> rows = document.Tables[0].Rows;
        TestAssert.Equal("bottom", rows[0].Cells[0].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("top", rows[0].Cells[1].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("center", rows[1].Cells[0].VerticalAlignmentValue ?? string.Empty);
        TestAssert.Equal("center", rows[1].Cells[1].VerticalAlignmentValue ?? string.Empty);
    }

    public static void DocxReaderTableCellPreservesShadingTokens()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:shd w:val="pct20" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>Shaded</w:t></w:r></w:p></w:tc>
                        <w:tc><w:p><w:r><w:t>Default</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        IReadOnlyList<DocxTableCell> cells = document.Tables[0].Rows[0].Cells;

        TestAssert.Equal("D9EAD3", cells[0].FillHex ?? string.Empty);
        TestAssert.Equal("pct20", cells[0].ShadingValue ?? string.Empty);
        TestAssert.Equal("112233", cells[0].ShadingColor ?? string.Empty);
        TestAssert.True(cells[1].ShadingValue is null, "Expected missing cell shading to keep a null source token.");
    }

    public static void DocxSyntheticTableCellPatternShadingUsesPdfTilingPattern()
    {
        string[] shadingValues =
        [
            "horzStripe",
            "thinHorzStripe",
            "vertStripe",
            "thinVertStripe",
            "diagStripe",
            "thinDiagStripe",
            "reverseDiagStripe",
            "thinReverseDiagStripe"
        ];
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid>
                        <w:gridCol w:w="720"/><w:gridCol w:w="720"/><w:gridCol w:w="720"/><w:gridCol w:w="720"/>
                        <w:gridCol w:w="720"/><w:gridCol w:w="720"/><w:gridCol w:w="720"/><w:gridCol w:w="720"/>
                      </w:tblGrid>
                      <w:tr>
                        <w:tc><w:tcPr><w:shd w:val="horzStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="thinHorzStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="vertStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>3</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="thinVertStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>4</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="diagStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>5</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="thinDiagStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>6</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="reverseDiagStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>7</w:t></w:r></w:p></w:tc>
                        <w:tc><w:tcPr><w:shd w:val="thinReverseDiagStripe" w:color="112233" w:fill="D9EAD3"/></w:tcPr><w:p><w:r><w:t>8</w:t></w:r></w:p></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/PatternType 1", pdf);
        TestAssert.Contains("/Pattern cs", pdf);
        TestAssert.Contains("/ImPattern", pdf);
        TestAssert.True(DocxTests.CountOccurrences(pdf, "/PatternType 1") >= shadingValues.Length, "Each supported DOCX stripe family should have a distinct tiling pattern resource.");
    }

    public static void DocxReaderTableCellPreservesParagraphModel()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:style w:type="paragraph" w:styleId="CellHeading">
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:rPr><w:sz w:val="32"/></w:rPr>
                  </w:style>
                  <w:style w:type="character" w:styleId="CellEmphasis">
                    <w:rPr><w:color w:val="336699"/><w:i/></w:rPr>
                  </w:style>
                </w:styles>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc>
                        <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
                        <w:p><w:pPr><w:pStyle w:val="CellHeading"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CellEmphasis"/><w:b/></w:rPr><w:t>Beta</w:t></w:r></w:p>
                      </w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        DocxTableCell cell = document.Tables[0].Rows[0].Cells[0];
        TestAssert.Equal("Alpha Beta", cell.Text);
        TestAssert.Equal(2, cell.Paragraphs.Count);
        TestAssert.Equal("Alpha", cell.Paragraphs[0].Runs[0].Text);
        TestAssert.Equal(DocxTextAlignment.Center, cell.Paragraphs[1].Alignment);
        TestAssert.Equal(16d, cell.Paragraphs[1].Runs[0].FontSize);
        TestAssert.Equal("336699", cell.Paragraphs[1].Runs[0].ColorHex ?? string.Empty);
        TestAssert.True(cell.Paragraphs[1].Runs[0].Bold, "Expected table-cell paragraph runs to preserve direct run properties.");
        TestAssert.True(cell.Paragraphs[1].Runs[0].Italic, "Expected table-cell paragraph runs to preserve character styles.");
    }

    public static void DocxReaderTableCellPreservesBodyFlowBreakElements()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr><w:tc>
                        <w:p><w:r><w:t>Left</w:t><w:br w:type="column"/><w:t>Right</w:t></w:r></w:p>
                      </w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        var diagnostics = new List<OoxPdfDiagnostic>();

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, diagnostics.Add, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.True(!diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_MANUAL_BREAK"), "Visible table-cell column breaks should be preserved structurally without a stale unsupported diagnostic.");

        DocxTableCell cell = document.Tables[0].Rows[0].Cells[0];
        TestAssert.Equal("Left Right", cell.Text);
        TestAssert.Equal(3, cell.BodyElements.Count);
        TestAssert.Equal(2, cell.Paragraphs.Count);
        TestAssert.Equal("Left", cell.Paragraphs[0].Runs.Single().Text);
        TestAssert.True(cell.BodyElements[1] is DocxManualBreakElement, "The table-cell body stream should preserve the typed column break.");
        TestAssert.Equal("Right", cell.Paragraphs[1].Runs.Single().Text);

        DocxStructureTableCellSnapshot cellSnapshot = DocxStructureSnapshot.FromDocument(document).Tables.Single().Rows.Single().Cells.Single();
        TestAssert.Equal(3, cellSnapshot.BodyElementCount);
        TestAssert.Equal(1, cellSnapshot.ManualBreakElementCount);
        TestAssert.Equal(0, cellSnapshot.PageBreakElementCount);
    }

    public static void DocxReaderTableInventoryIncludesNestedTableCellBodies()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr><w:tc>
                        <w:p><w:r><w:t>Outer</w:t></w:r></w:p>
                        <w:tbl>
                          <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                          <w:tr><w:tc><w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:tc></w:tr>
                        </w:tbl>
                      </w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);

        TestAssert.Equal(1, document.BodyElements.OfType<DocxTableElement>().Count());
        TestAssert.Equal(2, document.Tables.Count);
        TestAssert.Equal("Outer", document.Tables[0].Rows[0].Cells[0].Paragraphs[0].Runs.Single().Text);
        TestAssert.Equal("Nested", document.Tables[1].Rows[0].Cells[0].Paragraphs[0].Runs.Single().Text);
        TestAssert.True(document.Tables[0].Rows[0].Cells[0].BodyElements[1] is DocxTableElement, "The top-level body stream should not flatten nested table structure.");
    }

    public static void DocxReaderTableCellPreservesNumberedParagraphs()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
                """,
            ["word/numbering.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="7">
                    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
                </w:numbering>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>
                      <w:tr><w:tc>
                        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>
                      </w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxDocument document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        DocxTableCell cell = document.Tables[0].Rows[0].Cells[0];

        TestAssert.Equal("1.", cell.Paragraphs[0].ListLabel?.Text ?? string.Empty);
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(OpenTypeFont.Load(File.ReadAllBytes(arial)), "1. Item".EnumerateRunes().Select(rune => rune.Value), CancellationToken.None);
        DocxTextLineLayout line = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, embedded, CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells[0]
            .TextLines[0];
        TestAssert.Equal("1.\tItem", line.Text);
        TestAssert.Equal(3, line.Segments.Count);
        TestAssert.Equal("1.", line.Segments[0].Text);
        TestAssert.Equal(" ", line.Segments[1].Text);
        TestAssert.Equal("Item", line.Segments[2].Text);
        TestAssert.True(line.Segments[2].X > line.Segments[0].X, "Numbered table-cell text should be segmented after the list label.");
    }

    public static void DocxSyntheticTableKeepsBodyOrder()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:r><w:t>Before</w:t></w:r></w:p>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblBorders>
                          <w:top w:val="single" w:color="000000" w:sz="4"/>
                          <w:left w:val="single" w:color="000000" w:sz="4"/>
                          <w:bottom w:val="single" w:color="000000" w:sz="4"/>
                          <w:right w:val="single" w:color="000000" w:sz="4"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr><w:tc><w:p/></w:tc></w:tr>
                    </w:tbl>
                    <w:p><w:r><w:t>After</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int firstText = DocxTests.FirstPdfTextShowIndex(pdf);
        int tableGrid = pdf.IndexOf(" re f", StringComparison.Ordinal);
        int lastText = DocxTests.LastPdfTextShowIndex(pdf);
        TestAssert.True(firstText >= 0 && tableGrid > firstText && lastText > tableGrid, "DOCX tables should render in body order between surrounding paragraphs.");
    }

    public static void DocxSyntheticTableUsesRowHeights()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblBorders>
                          <w:top w:val="single" w:color="000000" w:sz="4"/>
                          <w:left w:val="single" w:color="000000" w:sz="4"/>
                          <w:bottom w:val="single" w:color="000000" w:sz="4"/>
                          <w:right w:val="single" w:color="000000" w:sz="4"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr><w:trPr><w:trHeight w:val="720"/></w:trPr><w:tc><w:p/></w:tc></w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        // Table terminus (w7 doc-start probe): the last-row bottom border hangs below
        // content, so the single-row bottom strip sits one 0.48pt width lower than the
        // legacy bottom-inside convention placed it.
        // w63 centered-vertical law: the strip starts at grid + half border width (72.24); the end keeps the butt-overlap convention.
        TestAssert.Contains("72.24 683.04 143.76 0.48 re f", pdf);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxTableRow row = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Tables[0].Rows[0];
        TestAssert.Equal("720", row.HeightValue ?? string.Empty);
        TestAssert.True(row.HeightRuleValue is null, "Missing hRule should stay distinct from exact/auto.");
    }

    public static void DocxReaderTableRowPreservesPropertyExceptionCellMargins()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr>
                        <w:tblPrEx>
                          <w:tblCellMar>
                            <w:top w:w="0" w:type="dxa"/>
                            <w:bottom w:w="0" w:type="dxa"/>
                          </w:tblCellMar>
                        </w:tblPrEx>
                        <w:tc><w:p/></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxTableRow row = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Tables[0].Rows[0];

        TestAssert.True(row.TablePropertyExceptionCellMargins is not null, "Expected row-level tblPrEx cell margins to stay distinct from cell margins.");
        TestAssert.Equal("0", row.TablePropertyExceptionCellMargins!.TopValue ?? string.Empty);
        TestAssert.Equal("0", row.TablePropertyExceptionCellMargins.BottomValue ?? string.Empty);
    }

    public static void DocxReaderTableRowPreservesCantSplit()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
                      <w:tr>
                        <w:trPr><w:cantSplit w:val="1"/></w:trPr>
                        <w:tc><w:p/></w:tc>
                      </w:tr>
                    </w:tbl>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        DocxTableRow row = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final).Tables[0].Rows[0];

        TestAssert.True(row.CantSplit, "Expected w:cantSplit to be preserved for row-fragment pagination.");
        TestAssert.Equal("1", row.CantSplitValue ?? string.Empty);
    }

    public static void DocxSyntheticTableRowsBreakAcrossPages()
    {
        string rows = string.Concat(Enumerable.Range(0, 8).Select(i => "<w:tr><w:tc><w:p/></w:tc></w:tr>"));
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>{{rows}}</w:tbl>
                    <w:sectPr><w:pgSz w:w="2880" w:h="2880"/><w:pgMar w:top="360" w:right="360" w:bottom="360" w:left="360"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Type /Pages /Count 2", pdf);
    }

    public static void DocxTableLayoutStageKeepsManualPageBreakBoundary()
    {
        DocxTable first = DocxTests.CreateSingleCellTable("first", rowHeight: 20d);
        DocxTable second = DocxTests.CreateSingleCellTable("second", rowHeight: 20d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([
            new DocxTableElement(first),
            new DocxPageBreakElement(DocxBreakSourceKind.PageBreakBefore, null, null),
            new DocxTableElement(second)
        ], [first, second]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded: null, cancellationToken: CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTableRowLayout>().Count());
    }

    public static void DocxTableLayoutStageManualColumnBreakAdvancesPageInSingleColumn()
    {
        DocxTable first = DocxTests.CreateSingleCellTable("first", rowHeight: 20d);
        DocxTable second = DocxTests.CreateSingleCellTable("second", rowHeight: 20d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([
            new DocxTableElement(first),
            new DocxManualBreakElement(DocxBreakSourceKind.RunBreak, "column", null),
            new DocxTableElement(second)
        ], [first, second]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded: null, cancellationToken: CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
        TestAssert.Equal(1, layout.Pages[1].Items.OfType<DocxTableRowLayout>().Count());
    }

    public static void DocxTableLayoutStageContinuesRowsInActiveColumnFrame()
    {
        var firstRow = new DocxTableRow([new DocxTableCell("first", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 100d);
        var secondRow = new DocxTableRow([new DocxTableCell("second", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 100d);
        var table = new DocxTable(null, [60d], [firstRow, secondRow]);
        DocxPageSettings sectionSettings = DocxPageSettings.Empty with
        {
            WidthValue = "4000",
            HeightValue = "4000",
            MarginLeftValue = "360",
            MarginRightValue = "360",
            MarginTopValue = "360",
            MarginBottomValue = "360"
        };
        var document = new DocxDocument(
            300d,
            300d,
            72d,
            72d,
            72d,
            72d,
            DocxPageSettings.Empty,
            [],
            [],
            [],
            [
                new DocxTableElement(table),
                new DocxSectionBreakElement(sectionSettings, DocxSectionBreakType.NextPage, "2", "1", "360", [])
            ],
            [],
            [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout[] rows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout);

        TestAssert.Equal(1, layout.Pages.Count);
        TestAssert.Equal(2, rows.Length);
        TestAssert.Equal(18d, rows[0].Table.TableX);
        TestAssert.Equal(109d, rows[1].Table.TableX);
        TestAssert.Equal(0, snapshot.Pages[0].Items[0].ColumnIndex ?? -1);
        TestAssert.Equal(1, snapshot.Pages[0].Items[1].ColumnIndex ?? -1);
    }

    public static void DocxTableLayoutStageRunPageBreakParagraphConsumesLineBox()
    {
        DocxTable first = DocxTests.CreateSingleCellTable("first", rowHeight: 150d);
        DocxTable second = DocxTests.CreateSingleCellTable("second", rowHeight: 20d);
        DocxParagraph marker = DocxTests.CreateDocxLayoutParagraph("marker", 10d, 20d);
        DocxParagraph breakParagraph = DocxTests.CreateDocxLayoutParagraph("", 10d, 20d);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([
            new DocxTableElement(first),
            new DocxParagraphElement(marker),
            new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, "page", breakParagraph),
            new DocxTableElement(second)
        ], [first, second]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded: null, cancellationToken: CancellationToken.None);

        TestAssert.Equal(3, layout.Pages.Count);
        TestAssert.Equal(1, layout.Pages[0].Items.OfType<DocxTableRowLayout>().Count());
        TestAssert.Equal(0, layout.Pages[1].Items.Count);
        TestAssert.Equal(1, layout.Pages[2].Items.OfType<DocxTableRowLayout>().Count());
    }

    public static void DocxTableLayoutStageRepeatsHeaderRowsAfterPageBreak()
    {
        var header = new DocxTableRow([new DocxTableCell("Header", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d) with {IsHeader = true };
        var first = new DocxTableRow([new DocxTableCell("First", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 50d);
        var second = new DocxTableRow([new DocxTableCell("Second", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 50d);
        var table = new DocxTable(null, [60d], [header, first, second]);
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded: null, cancellationToken: CancellationToken.None);

        TestAssert.Equal(2, layout.Pages.Count);
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();
        TestAssert.Equal(2, firstPageRows.Length);
        TestAssert.Equal(2, secondPageRows.Length);
        TestAssert.Equal("Header", secondPageRows[0].Cells[0].Cell.Text);
        TestAssert.Equal("Second", secondPageRows[1].Cells[0].Cell.Text);
    }

    public static void DocxTableRowHeightIncludesListLabelFirstLineExtraLeading()
    {
        DocxListLabel tallLabel = new DocxListLabel("*", "bullet", "*", "tab", "1", 0, DocxNumberingIndent.Empty, new DocxTextRunStyle(10d, null, false, false, false, null, "Label Metrics", new DocxRunFonts("Label Metrics", null, null, null, null, null, null, null)));
        DocxListLabel flatLabel = new DocxListLabel("*", "bullet", "*", "tab", "1", 0, DocxNumberingIndent.Empty, new DocxTextRunStyle(10d, null, false, false, false, null, null, new DocxRunFonts(null, null, null, null, null, null, null, null)));
        DocxParagraphSpacing autoSpacing = new DocxParagraphSpacing(null, null, null, null, null, null, null, "auto", null);
        DocxParagraph tallItem = new DocxParagraph([new DocxTextRun("Tall", 10d, null, false, false, false, null, null)], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1.15d, null, autoSpacing, DocxParagraphKeepRules.Empty, tallLabel);
        DocxParagraph flatItem = new DocxParagraph([new DocxTextRun("Flat", 10d, null, false, false, false, null, null)], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1.15d, null, autoSpacing, DocxParagraphKeepRules.Empty, flatLabel);
        var tallRow = new DocxTableRow([new DocxTableCell("Tall", [tallItem], null, null, null, null, [], DocxTableCellMargins.Empty)], null);
        var flatRow = new DocxTableRow([new DocxTableCell("Flat", [flatItem], null, null, null, null, [], DocxTableCellMargins.Empty)], null);
        var table = new DocxTable(null, [400d], [tallRow, flatRow]);
        var document = new DocxDocument(400d, 400d, 10d, 10d, 10d, 10d, DocxPageSettings.Empty, [], [], [], [new DocxTableElement(table)], [], [table]);
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
        DocxTableRowLayout[] rows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] tallRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        TestAssert.Equal(2, tallRows.Length);
        // RV06 line-box maximum: row heights compose the Windows-box line height (tall-minus-flat keeps the 1.0 label-ascender excess).
        TestAssert.Equal(14.92d, Math.Round(tallRows[0].Height, 4));
        TestAssert.Equal(13.92d, Math.Round(tallRows[1].Height, 4));
    }


    public static void DocxTableCellFirstBaselineFollowsBoxMinusDescender()
    {
        // RV06 cell box rule (Word 16.0, Harlow/Arial/Calibri size matrix plus Palatino/Harlow mixed orders): in-cell first baselines sit one full line box above the box bottom: per-run max(Windows box, hhea box plus hhea gap) minus max descender, maximized across runs with no 0.94em floor.
        double hi = LayoutCellBoxFirst("HiAsc", "HiAsc");
        double lo = LayoutCellBoxFirst("LoAsc", "LoAsc");
        double mixHi = LayoutCellBoxFirst("HiAsc", "LoAsc");
        double mixLo = LayoutCellBoxFirst("LoAsc", "HiAsc");
        TestAssert.True(Math.Abs((hi - lo) - -2.16d) < 0.02d, "Uniform cell first gap must carry the box spread; observed=" + (hi - lo).ToString(CultureInfo.InvariantCulture) + ".");
        TestAssert.True(Math.Abs((mixHi - lo) - -2.16d) < 0.02d, "Mixed cell first must sit at the max-box level.");
        TestAssert.True(Math.Abs(mixHi - mixLo) < 0.02d, "Mixed cell first must ignore run order.");
    }

    private static double LayoutCellBoxFirst(string firstFamily, string secondFamily)
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>""",
            ["_rels/.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
            ["word/_rels/document.xml.rels"] = """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""",
            ["word/document.xml"] = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="F1" w:hAnsi="F1"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Cell box line one</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="F2" w:hAnsi="F2"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">Cell box line two</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""".Replace("F1", firstFamily).Replace("F2", secondFamily)
        });
        DocxDocument document;
        using (FileStream stream = File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            document = new DocxReader().Read(package, null, CancellationToken.None, OoxPdfDocxMarkupMode.Final);
        }
        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new CellBoxTextMeasurer(new DocxTests.FamilyWidthTextMeasurer()), CancellationToken.None);
        DocxTextLineLayout[] lines = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single().Cells.Single().TextLines.Where(line => line.Text.StartsWith("Cell box", StringComparison.Ordinal)).Take(1).ToArray();
        TestAssert.Equal(1, lines.Length);
        TestAssert.Equal(firstFamily, lines[0].StyleRun.FontFamily);
        return lines[0].BaselineY;
    }

    private sealed class CellBoxTextMeasurer(IDocxTextMeasurer inner) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider, IDocxHheaDescenderProvider, IDocxHheaLineGapProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => inner.MeasureText(run, text, fontSize);

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize) => 14d;

        public double MeasureHheaLineHeight(DocxTextRun? run, double fontSize) => 14d;

        public double MeasureHheaAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "HiAsc", StringComparison.Ordinal) ? fontSize * 1.05d : fontSize * 0.5d;

        public double MeasureHheaDescender(DocxTextRun? run, double fontSize) => fontSize * 0.1d;

        public double MeasureHheaLineGap(DocxTextRun? run, double fontSize) => 0d;

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize) => string.Equals(run?.FontFamily, "HiAsc", StringComparison.Ordinal) ? fontSize * 1.05d : fontSize * 0.87d;

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize) => fontSize * 0.1d;
    }

    public static void DocxTableCellPageFieldsEvaluatePerFragmentPage()
    {
        // W04: PAGE fields must evaluate against each fragment page. The static
        // cell memo omits page args only for cells without dynamic field runs, so
        // the two fragments below must resolve different page numbers.
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph[] secondParagraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph(
                index == 4 ? "Page {PAGE} mark" : "Line " + index.ToString(CultureInfo.InvariantCulture),
                10d,
                10d))
            .ToArray();
        var first = new DocxTableRow([new DocxTableCell("First", [firstParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var second = new DocxTableRow([new DocxTableCell("Second", secondParagraphs, null, null, null, null, [], DocxTableCellMargins.Empty)], 80d);
        var table = new DocxTable(null, [60d], [first, second]);
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

        TestAssert.Equal(2, layout.Pages.Count);
        string firstFragmentText = string.Concat(layout.Pages[0].Items.OfType<DocxTableRowLayout>().Where(row => row.RowIndex == 1).SelectMany(row => row.Cells[0].TextLines).Select(line => line.Text));
        string secondFragmentText = string.Concat(layout.Pages[1].Items.OfType<DocxTableRowLayout>().Where(row => row.RowIndex == 1).SelectMany(row => row.Cells[0].TextLines).Select(line => line.Text));
        bool firstHasOwn = firstFragmentText.Contains("Page 1 mark", StringComparison.Ordinal);
        bool secondHasOwn = secondFragmentText.Contains("Page 2 mark", StringComparison.Ordinal);
        bool firstHasForeign = firstFragmentText.Contains("Page 2 mark", StringComparison.Ordinal);
        bool secondHasForeign = secondFragmentText.Contains("Page 1 mark", StringComparison.Ordinal);
        TestAssert.True(
            (firstHasOwn && !firstHasForeign && !secondHasOwn && !secondHasForeign) ||
            (!firstHasOwn && !firstHasForeign && secondHasOwn && !secondHasForeign),
            "Exactly one split fragment must resolve PAGE against its own page. First: <" + firstFragmentText + "> Second: <" + secondFragmentText + ">");
        TestAssert.True(!firstFragmentText.Contains("{PAGE}", StringComparison.Ordinal) && !secondFragmentText.Contains("{PAGE}", StringComparison.Ordinal), "Split fragments must not leak unresolved PAGE placeholders.");
    }

    public static void DocxTableCellTextLinesMemoSharesSplitCells()
    {
        // Observe the actual cell-layout path through one local memo. A split
        // changes the placement height/origin, while its complete measured lines
        // remain reusable. Compare with uncached layout at the second origin.
        DocxParagraph[] paragraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        var cell = new DocxTableCell("Second", paragraphs, null, null, null, null, [], DocxTableCellMargins.Empty);
        var memo = new DocxLayoutEngine.DocxTableCellTextLinesMemo();
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        var method = typeof(DocxLayoutEngine).GetMethod("LayoutTableCellTextLines",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected live cell-layout method.");

        (IReadOnlyList<DocxTextLineLayout> Lines, IReadOnlyList<DocxInlineImageLayout> Images, double Height)
            Place(double y, double height, DocxLayoutEngine.DocxTableCellTextLinesMemo? cache) =>
            ((IReadOnlyList<DocxTextLineLayout>, IReadOnlyList<DocxInlineImageLayout>, double))method.Invoke(null,
                [cell, 10d, y, 60d, height, 0d, measurer, 36d, 1, 2, 1d, cache, 0, null, CancellationToken.None])!;

        var first = Place(70d, 20d, memo);
        var replay = Place(10d, 60d, memo);
        var fresh = Place(10d, 60d, null);
        TestAssert.Equal(1, memo.Hits);
        TestAssert.Equal(1, memo.Misses);
        TestAssert.Equal(8, first.Lines.Count);
        TestAssert.Equal(fresh.Height, replay.Height);
        TestAssert.Equal(fresh.Lines.Count, replay.Lines.Count);
        for (int index = 0; index < fresh.Lines.Count; index++)
        {
            DocxTextLineLayout expected = fresh.Lines[index];
            DocxTextLineLayout actual = replay.Lines[index];
            TestAssert.Equal(expected.Text, actual.Text);
            TestAssert.True(Math.Abs(expected.X - actual.X) < 1e-9 && Math.Abs(expected.BaselineY - actual.BaselineY) < 1e-9,
                "Cached split-cell placement must equal fresh geometry at its new origin.");
            TestAssert.True(expected.Segments.SequenceEqual(actual.Segments), "Cached split-cell segments must equal fresh segments.");
        }
    }

    public static void DocxTableLayoutStageSplitsTallRowsAcrossPagesByDefault()
    {
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph[] secondParagraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        var first = new DocxTableRow([new DocxTableCell("First", [firstParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var second = new DocxTableRow([new DocxTableCell("Second", secondParagraphs, null, null, null, null, [], DocxTableCellMargins.Empty)], 80d);
        var table = new DocxTable(null, [60d], [first, second]);
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
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(2, firstPageRows.Length);
        TestAssert.Equal(1, secondPageRows.Length);
        TestAssert.Equal(1, firstPageRows[1].RowIndex);
        TestAssert.Equal(0, firstPageRows[1].FragmentIndex);
        TestAssert.Equal(2, firstPageRows[1].FragmentCount);
        TestAssert.Equal("PageBoundary", firstPageRows[1].FragmentReason);
        TestAssert.Equal(80d, firstPageRows[1].FullRowHeight);
        TestAssert.Equal(0d, firstPageRows[1].FragmentOffsetFromRowTop);
        TestAssert.Equal(20d, firstPageRows[1].Height);
        TestAssert.Equal(1, secondPageRows[0].RowIndex);
        TestAssert.Equal(1, secondPageRows[0].FragmentIndex);
        TestAssert.Equal(2, secondPageRows[0].FragmentCount);
        TestAssert.Equal("PageBoundary", secondPageRows[0].FragmentReason);
        TestAssert.Equal(80d, secondPageRows[0].FullRowHeight);
        TestAssert.Equal(20d, secondPageRows[0].FragmentOffsetFromRowTop);
        TestAssert.Equal(60d, secondPageRows[0].Height);
        TestAssert.True(firstPageRows[1].Cells[0].TextLines.Count > 0, "The first split fragment should own its visible row text.");
        TestAssert.True(secondPageRows[0].Cells[0].TextLines.Count > 0, "The continuation split fragment should own its visible row text.");
        TestAssert.Equal(8, firstPageRows[1].Cells[0].TextLines.Count + secondPageRows[0].Cells[0].TextLines.Count);

        DocxTableSnapshot snapshot = DocxLayoutSnapshot.FromLayout(layout).Tables.Single();
        TestAssert.Equal(1, snapshot.FragmentedRowCount);
        TestAssert.Equal(2, snapshot.FragmentedRowLayoutCount);
        TestAssert.Equal(2, snapshot.MaxRowFragmentCount);
        DocxTableRowSnapshot[] splitRowSnapshots = DocxLayoutSnapshot.FromLayout(layout).Pages
            .SelectMany(page => page.TableRows)
            .Where(row => row.RowIndex == 1)
            .OrderBy(row => row.FragmentIndex)
            .ToArray();
        TestAssert.Equal(80d, splitRowSnapshots[0].FullRowHeight);
        TestAssert.Equal("PageBoundary", splitRowSnapshots[0].FragmentReason);
        TestAssert.Equal(0d, splitRowSnapshots[0].FragmentOffsetFromRowTop);
        TestAssert.Equal(80d, splitRowSnapshots[1].FullRowHeight);
        TestAssert.Equal("PageBoundary", splitRowSnapshots[1].FragmentReason);
        TestAssert.Equal(20d, splitRowSnapshots[1].FragmentOffsetFromRowTop);
    }

    public static void DocxTableLayoutStageKeepsCellKeepLinesParagraphWholeAtPageBoundary()
    {
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph keptCellParagraph = DocxTests.CreateDocxLayoutParagraph(
            string.Join('\n', Enumerable.Range(1, 8).Select(index => "Line " + index.ToString(CultureInfo.InvariantCulture))),
            10d,
            10d,
            keepRules: DocxParagraphKeepRules.Empty with { KeepLines = true });
        var first = new DocxTableRow([new DocxTableCell("First", [firstParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var second = new DocxTableRow([new DocxTableCell("Second", [keptCellParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 80d);
        var table = new DocxTable(null, [60d], [first, second]);
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
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, firstPageRows.Length);
        TestAssert.Equal(1, secondPageRows.Length);
        TestAssert.Equal(1, secondPageRows[0].RowIndex);
        TestAssert.Equal(0, secondPageRows[0].FragmentIndex);
        TestAssert.Equal(1, secondPageRows[0].FragmentCount);
        TestAssert.Equal("None", secondPageRows[0].FragmentReason);
        TestAssert.Equal(8, secondPageRows[0].Cells[0].TextLines.Count);
        TestAssert.Equal(0, DocxLayoutSnapshot.FromLayout(layout).Tables.Single().FragmentedRowCount);
    }

    public static void DocxTableLayoutStageKeepsCellWidowControlledParagraphWholeAtPageBoundary()
    {
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph widowControlledParagraph = DocxTests.CreateDocxLayoutParagraph("One\nTwo\nThree", 10d, 10d);
        var first = new DocxTableRow([new DocxTableCell("First", [firstParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var second = new DocxTableRow([new DocxTableCell("Second", [widowControlledParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 30d);
        var table = new DocxTable(null, [60d], [first, second]);
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
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, firstPageRows.Length);
        TestAssert.Equal(1, secondPageRows.Length);
        TestAssert.Equal(1, secondPageRows[0].RowIndex);
        TestAssert.Equal(1, secondPageRows[0].FragmentCount);
        TestAssert.Equal("None", secondPageRows[0].FragmentReason);
        TestAssert.Equal(3, secondPageRows[0].Cells[0].TextLines.Count);
    }

    public static void DocxTableLayoutStageKeepsCellKeepNextParagraphWithFollowingParagraphAtPageBoundary()
    {
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph keepNextParagraph = DocxTests.CreateDocxLayoutParagraph(
            "One\nTwo",
            10d,
            10d,
            keepRules: DocxParagraphKeepRules.Empty with { KeepNext = true });
        DocxParagraph followingParagraph = DocxTests.CreateDocxLayoutParagraph("Three", 10d, 10d);
        var first = new DocxTableRow([new DocxTableCell("First", [firstParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var second = new DocxTableRow([new DocxTableCell("Second", [keepNextParagraph, followingParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 30d);
        var table = new DocxTable(null, [60d], [first, second]);
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
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(1, firstPageRows.Length);
        TestAssert.Equal(1, secondPageRows.Length);
        TestAssert.Equal(1, secondPageRows[0].RowIndex);
        TestAssert.Equal(1, secondPageRows[0].FragmentCount);
        TestAssert.Equal("None", secondPageRows[0].FragmentReason);
        TestAssert.Equal(3, secondPageRows[0].Cells[0].TextLines.Count);
    }

    public static void DocxSharedHorizontalTableBorderHangsBelowBoundary()
    {
        // Office A/B (w9 exact/atLeast bordered row pairs, Word-COM rendered plus
        // PdfInspect, re-rendered digit-identical): interior horizontal border bands
        // hang below the row boundary by the full nominal width (atLeast mid band at
        // 683.02 for boundary 683.52, exact mid band at 683.50 for boundary 684.0),
        // while the layout centered them on the boundary (atLeast 683.28, exact
        // 683.76). Text baselines are identical both sides, so the shift is
        // pitch-neutral paint. The band start (72.24 here versus Office 72.26) is now Office-exact per the w63 centered-vertical law; the end keeps the butt-overlap convention (239.76 here versus Office 239.54).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:tbl>
                      <w:tblPr>
                        <w:tblW w:w="0" w:type="auto"/>
                        <w:tblBorders>
                          <w:top w:val="single" w:color="000000" w:sz="4"/>
                          <w:left w:val="single" w:color="000000" w:sz="4"/>
                          <w:bottom w:val="single" w:color="000000" w:sz="4"/>
                          <w:right w:val="single" w:color="000000" w:sz="4"/>
                          <w:insideH w:val="single" w:color="000000" w:sz="4"/>
                          <w:insideV w:val="single" w:color="000000" w:sz="4"/>
                        </w:tblBorders>
                      </w:tblPr>
                      <w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid>
                      <w:tr><w:trPr><w:trHeight w:val="720" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>Row one</w:t></w:r></w:p></w:tc></w:tr>
                      <w:tr><w:trPr><w:trHeight w:val="720" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>Row two</w:t></w:r></w:p></w:tc></w:tr>
                    </w:tbl>
                    <w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr><w:t>After pair</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72.24 683.02 239.76 0.48 re f", pdf);
    }

    private static string WriteW68TablePackage(string tblPrExtra)
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders>
                """ + tblPrExtra + """
                </w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>S4 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>S4 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>
                """
        });
    }

    public static void DocxTableCellTextStartsAtMaxBorderMarginInset()
    {
        // Office A/B (w63/w65 border-size probes, Word-COM rendered plus PdfInspect):
        // first-column text starts at grid + max(borderHalf, margin) with unset
        // margins defaulting to 0.48pt: sz4 text at 72.504 for grid 72.024 while
        // the layout stacked margin + borderHalf, placing sz4 text at 72.24 here.
        string input = WriteW68TablePackage("");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 72.48 709.202 Tm", pdf);
    }

    public static void DocxTableVerticalBordersCenterOnGridLines()
    {
        // Office A/B (w63 border-size probe, Word-COM rendered plus PdfInspect):
        // vertical border bands center on grid lines: sz4 outer-left at 71.784
        // for grid 72.024 with top bands butting from grid + half width, while
        // the renderer hung every band right of its edge (72.0 here).
        string input = WriteW68TablePackage("");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("71.76 704.387 0.48 15.613 re f", pdf);
        TestAssert.DoesNotContain("72 704.387 0.48 15.613 re f", pdf);
        TestAssert.Contains("72.24 719.52 119.76 0.48 re f", pdf);
    }

    public static void DocxIndentedTablePinsTextAtIndentOrigin()
    {
        // Office A/B (w62/w66/w68 indent probes, Word-COM rendered plus PdfInspect):
        // tblInd pins first-column TEXT at margin + indent for every border size
        // (w66 B4/D12/C24 all at 108.02), so the grid hangs left by
        // max(borderHalf, margin). The layout added indent without shifting,
        // placing sz4 indented text at 108.24 here.
        string input = WriteW68TablePackage("<w:tblInd w:w=\"720\" w:type=\"dxa\"/>");

        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 108 709.202 Tm", pdf);
    }

    public static void DocxCompatGridAlignsOuterBorderAtMargin()
    {
        // Office A/B (m13 bare+compat15 probe, Word-COM rendered plus PdfInspect):
        // a present compatibilityMode aligns the outer border edge at the margin,
        // so the sz4 grid sits at 72.24 and text at 72.72 (modern grid would give
        // band 71.76 and text 72.48).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>C15 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>C15 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>

                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 72.72 709.202 Tm", pdf);
        TestAssert.Contains("72 704.387 0.48 15.613 re f", pdf);
    }

    public static void DocxCompatIndentAlignsOuterBorderAtIndentedMargin()
    {
        // Office A/B (m14 compat+indent probe, Word-COM rendered plus PdfInspect):
        // compatibilityMode aligns the outer border edge at margin+indent with text
        // at grid+max (Office col1 text 108.74, band [108.02, 108.50], top strip
        // butting at 108.50; pipeline unsnapped: grid 108.24, text 108.72).
        // Pre-change the unshifted grid stacked margin+border, giving 108.24 here.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders><w:tblInd w:w="720" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>M14 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>M14 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>

                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 108.72 709.202 Tm", pdf);
        TestAssert.Contains("108 704.387 0.48 15.613 re f", pdf);
    }
    public static void DocxStyleMarginsPinTextAtMarginWithoutCompat()
    {
        // Office A/B (w72 bare+styles probe, Word-COM rendered plus PdfInspect):
        // without compat, table-style margins pin text at the margin (grid hangs
        // left by the style margin: TN-108dxa grid at 66.6, text at 72.0).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>P72 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>P72 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 72 709.202 Tm", pdf);
        TestAssert.Contains("66.36 704.387 0.48 15.613 re f", pdf);
    }

    public static void DocxTableNormalMissingCellMarginsFallBackToBuiltIn()
    {
        // Office A/B (m1 ladder mutant, Word-COM rendered plus PdfInspect): a present
        // TableNormal style without stored cell margins still contributes the built-in
        // 108dxa margins (Word col1 text at 77.664 for grid 72.264 under compat;
        // pipeline emits 77.64 = 72 + 0.24 + 5.4, i.e. the same law on our unsnapped
        // base: Word export shifts every X by a global +0.024 snap, inside gate tolerance).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>F6 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>F6 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/></w:tblPr></w:style></w:styles>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 77.64 709.202 Tm", pdf);
    }

    public static void DocxDirectMarginsReplaceStyleMarginsForTextOffset()
    {
        // Office A/B (w74 ladder+direct-20dxa probe, Word-COM rendered plus PdfInspect):
        // direct margins cascade-replace style margins for the text offset (col1 at
        // 73.224 for grid 72.264, i.e. grid + max(borderHalf, direct);
        // pipeline emits 73.24 = 72 + 0.24 + max(0.24, 1.0) on our unsnapped base
        // (same +0.024 Word export snap; residual direct-margin 600dpi px-quantization
        // mapped over 8 sweep points stays parked per precedent).
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:color="000000" w:sz="4"/><w:left w:val="single" w:color="000000" w:sz="4"/><w:bottom w:val="single" w:color="000000" w:sz="4"/><w:right w:val="single" w:color="000000" w:sz="4"/><w:insideH w:val="single" w:color="000000" w:sz="4"/><w:insideV w:val="single" w:color="000000" w:sz="4"/></w:tblBorders><w:tblCellMar><w:left w:w="20" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>G74 left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>G74 right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>
                """,
            ["word/_rels/document.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>
                """,
            ["word/styles.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>
                """,
            ["word/settings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 73.24 709.202 Tm", pdf);
    }

    public static void DocxTableRendererDoesNotDrawRowEdgeBordersAtSplitFragmentBoundaries()
    {
        DocxTableCellBorder[] borders =
        [
            new("top", "single", "000000", "4"),
            new("bottom", "single", "000000", "4"),
            new("left", "single", "000000", "4"),
            new("right", "single", "000000", "4")
        ];
        DocxParagraph firstParagraph = DocxTests.CreateDocxLayoutParagraph("First", 10d, 10d);
        DocxParagraph[] splitParagraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        var first = new DocxTableRow([new DocxTableCell("First", [firstParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var split = new DocxTableRow([new DocxTableCell("Split", splitParagraphs, null, null, null, null, borders, DocxTableCellMargins.Empty)], 80d);
        var table = new DocxTable(null, [60d], [first, split]);
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

        List<PdfPage> pages = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, CancellationToken.None).ToList();

        TestAssert.Equal(2, pages.Count);
        TestAssert.DoesNotContain("10.48 10 59.52 0.48 re f", pages[0].Content);
        TestAssert.DoesNotContain("10.48 89.52 59.52 0.48 re f", pages[1].Content);
        // w63 centered-vertical law: the outer-left band centers on the grid line (9.76 for grid 10).
        TestAssert.Contains("9.76 10 0.48 20 re f", pages[0].Content);
        // Table terminus (w7 doc-start probe): the split row is the last row, so its
        // continuation fragment carries the extra bottom width and its bottom strip
        // sits 0.48pt lower; fragment-edge suppression at the split itself is unchanged.
        // w63 centered-vertical law: the bottom strip starts at grid + half border width (10.24).
        TestAssert.Contains("10.24 29.04 59.76 0.48 re f", pages[1].Content);
    }

    public static void DocxTableLayoutStageRepeatsHeaderRowsBeforeSplitRowContinuations()
    {
        DocxParagraph headerParagraph = DocxTests.CreateDocxLayoutParagraph("Header", 10d, 10d);
        DocxParagraph fillerParagraph = DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 10d);
        DocxParagraph[] splitParagraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        var header = new DocxTableRow([new DocxTableCell("Header", [headerParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 10d) with {IsHeader = true };
        var filler = new DocxTableRow([new DocxTableCell("Filler", [fillerParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 50d);
        var split = new DocxTableRow([new DocxTableCell("Split", splitParagraphs, null, null, null, null, [], DocxTableCellMargins.Empty)], 80d);
        var table = new DocxTable(null, [60d], [header, filler, split]);
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
        DocxTableRowLayout[] firstPageRows = layout.Pages[0].Items.OfType<DocxTableRowLayout>().ToArray();
        DocxTableRowLayout[] secondPageRows = layout.Pages[1].Items.OfType<DocxTableRowLayout>().ToArray();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(3, firstPageRows.Length);
        TestAssert.Equal(2, secondPageRows.Length);
        TestAssert.Equal(0, firstPageRows[0].RowIndex);
        TestAssert.True(firstPageRows[0].IsHeader, "The first page should keep the table header before body rows.");
        TestAssert.Equal(2, firstPageRows[2].RowIndex);
        TestAssert.Equal(0, firstPageRows[2].FragmentIndex);
        TestAssert.Equal(2, firstPageRows[2].FragmentCount);
        TestAssert.Equal(0, secondPageRows[0].RowIndex);
        TestAssert.True(secondPageRows[0].IsHeader, "Split-row continuations should repeat table headers before the carried fragment.");
        TestAssert.Equal(2, secondPageRows[1].RowIndex);
        TestAssert.Equal(1, secondPageRows[1].FragmentIndex);
        TestAssert.Equal(2, secondPageRows[1].FragmentCount);
        TestAssert.Equal(60d, secondPageRows[1].Height);
    }

    public static void DocxTableLayoutStageKeepsInlineImagesInSplitRowFragments()
    {
        var image = new DocxInlineImage(20d, 20d, "image/png", [0x89, 0x50, 0x4E, 0x47], "word/media/image1.png");
        DocxParagraph fillerParagraph = DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 10d);
        DocxParagraph[] splitParagraphs = Enumerable.Range(1, 7)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .Append(new DocxParagraph(
                [],
                [image],
                null,
                DocxTextAlignment.Left,
                null,
                0d,
                0d,
                1d,
                10d,
                DocxParagraphSpacing.Empty,
                DocxParagraphKeepRules.Empty,
                null))
            .ToArray();
        var filler = new DocxTableRow([new DocxTableCell("Filler", [fillerParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var split = new DocxTableRow([new DocxTableCell("Split", splitParagraphs, null, null, null, null, [], DocxTableCellMargins.Empty)], 100d);
        var table = new DocxTable(null, [60d], [filler, split]);
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
        DocxTableRowLayout[] splitFragments = layout.Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .Where(row => row.RowIndex == 1)
            .ToArray();

        TestAssert.Equal(2, splitFragments.Length);
        TestAssert.True(splitFragments.All(fragment => fragment.FragmentCount == 2), "The image-owning body row should split into two fragments.");
        TestAssert.True(splitFragments.Sum(fragment => fragment.Cells.Sum(cell => cell.InlineImages.Count)) > 0, "Split row fragments should keep overlapping inline image layouts for the renderer clip path.");
    }

    public static void DocxTableLayoutStageClipsVerticalMergeRestartToSplitRowFragments()
    {
        DocxParagraph fillerParagraph = DocxTests.CreateDocxLayoutParagraph("Filler", 10d, 10d);
        DocxParagraph[] splitParagraphs = Enumerable.Range(1, 8)
            .Select(index => DocxTests.CreateDocxLayoutParagraph("Line " + index.ToString(CultureInfo.InvariantCulture), 10d, 10d))
            .ToArray();
        var filler = new DocxTableRow([new DocxTableCell("Filler", [fillerParagraph], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var restart = new DocxTableRow([
            new DocxTableCell(
                "Merged",
                splitParagraphs,
                null,
                null,
                null,
                null,
                [],
                DocxTableCellMargins.Empty) with {HasVerticalMerge = true,VerticalMergeValue = "restart" }
        ], 100d);
        var continuation = new DocxTableRow([
            new DocxTableCell(
                "Continuation",
                [],
                null,
                null,
                null,
                null,
                [],
                DocxTableCellMargins.Empty) with {HasVerticalMerge = true }
        ], 20d);
        var table = new DocxTable(null, [60d], [filler, restart, continuation]);
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

        DocxTableRowLayout[] splitFragments = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages
            .SelectMany(page => page.Items.OfType<DocxTableRowLayout>())
            .Where(row => row.RowIndex == 1)
            .ToArray();

        TestAssert.Equal(2, splitFragments.Length);
        TestAssert.True(splitFragments.All(fragment => fragment.FragmentCount == 2), "The merged restart row should still split into physical row fragments.");
        TestAssert.True(splitFragments.All(fragment => fragment.Cells[0].Y == fragment.Y), "A split merged restart cell should use the fragment top as its visible clip.");
        TestAssert.True(splitFragments.All(fragment => fragment.Cells[0].Height == fragment.Height), "A split merged restart cell should clip to the fragment height, not the full cross-row merged span.");
        TestAssert.True(splitFragments.All(fragment => fragment.Cells[0].TextLines.Count != 0), "Full merged-cell text coordinates should remain available for fragment clipping.");
    }

    public static void DocxTableLayoutStageNormalizesPlainCellTextThroughSharedParagraphs()
    {
        var table = new DocxTable(
            null,
            [60d],
            [new DocxTableRow([new DocxTableCell("Plain", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table)], [table]);

        DocxTableCellLayout cell = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None)
            .Pages[0]
            .Items
            .OfType<DocxTableRowLayout>()
            .Single()
            .Cells
            .Single();

        DocxTextLineLayout line = cell.TextLines.Single();
        TestAssert.Equal("Plain", line.Text);
        TestAssert.Equal(11d, line.FontSize);
    }

    public static void DocxTableLayoutStageHonorsCantSplitRowsAtPageBoundary()
    {
        var first = new DocxTableRow([new DocxTableCell("First", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 60d);
        var second = new DocxTableRow([new DocxTableCell("Second", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 80d) with {CantSplit = true,CantSplitValue = "1" };
        var table = new DocxTable(null, [60d], [first, second]);
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

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, embedded: null, cancellationToken: CancellationToken.None);
        DocxTableRowLayout firstPageRow = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        DocxTableRowLayout secondPageRow = layout.Pages[1].Items.OfType<DocxTableRowLayout>().Single();

        TestAssert.Equal(2, layout.Pages.Count);
        TestAssert.Equal(0, firstPageRow.RowIndex);
        TestAssert.Equal(1, secondPageRow.RowIndex);
        TestAssert.Equal(0, secondPageRow.FragmentIndex);
        TestAssert.Equal(1, secondPageRow.FragmentCount);
        TestAssert.True(secondPageRow.CantSplit, "w:cantSplit rows should move whole instead of creating fragments.");
    }

    public static void DocxTableLayoutStageKeepsFollowingParagraphAdjacent()
    {
        var paragraph = new DocxParagraph(
            [new DocxTextRun("After", 10d, null, false, false, false, null, null)],
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
        var table = new DocxTable(
            null,
            [60d],
            [new DocxTableRow([new DocxTableCell("Cell", [], null, null, null, null, [], DocxTableCellMargins.Empty)], 20d)]);
        DocxDocument document = DocxTests.CreateLayoutTestDocument([new DocxTableElement(table), new DocxParagraphElement(paragraph)], [table]);

        DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);

        DocxTableRowLayout row = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        DocxTextLineLayout following = layout.Pages[0].Items.OfType<DocxTextLineLayout>().Single();
        double expectedBaselineY = row.Y - DocxLineMetrics.ResolveBodyBaselineOffset(10d, 10d, hasExplicitLineSpacing: true);
        TestAssert.Equal(Math.Round(expectedBaselineY, 3), Math.Round(following.BaselineY, 3));
    }
}
