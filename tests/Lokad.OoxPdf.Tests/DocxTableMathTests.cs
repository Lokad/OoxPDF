using System.Text.Json;
using System.Security.Cryptography;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using System.Text;

namespace Lokad.OoxPdf.Tests;

internal static class DocxTableMathTests
{
    public static void DocxTableMathOrderedFramesMatchIndependentBorderClearance()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent bordered controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-note-border-office-baselines.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            if (!id.EndsWith("-baseline", StringComparison.Ordinal)) continue;
            string input = Path.Combine(cases, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var renderer = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            (double Math, double Plain) Gaps(string path)
            {
                var table = renderer.InspectLayout(DocxTests.ReadDocx(path, mode)).Pages.Single().Items.Single(item => item.Kind == "TableRow");
                double top = table.Y + table.Height;
                return (top - table.TextLines!.Single(line => line.TextLength == 1).Y,
                    top - table.TextLines!.Single(line => line.TextLength > 1).Y);
            }
            var gap = Gaps(input);
            TestAssert.True(Math.Abs(gap.Math - row.GetProperty("MarkerBelowTop").GetDouble()) < .25d,
                $"The ordered math frame must retain Word's border clearance: {id}/{mode}/gap={gap.Math}.");
            TestAssert.True(Math.Abs(gap.Plain - row.GetProperty("PlainPeerBelowTop").GetDouble()) < .25d,
                "The ordinary peer retains its independently qualified nominal-inset conversion.");
            var superscript = Gaps(Path.Combine(cases, id.Replace("-baseline", "-superscript", StringComparison.Ordinal) + ".docx"));
            TestAssert.True(Math.Abs(superscript.Math - gap.Math) < .000001d,
                "Changing only the note script must preserve the fraction's baseline frame.");
        }
    }

    public static void DocxTableMathNaturalParagraphFramesMatchIndependentTableOrigins()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent mixed-flow table controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "table-flow-frame-office-baselines.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string source = Path.Combine(cases, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(source))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var renderer = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var page = renderer.InspectLayout(DocxTests.ReadDocx(source, mode)).Pages.Single();
            double leading = page.Items.Single(item => item.Kind == "TextLine" && item.SourceBlockIndex == 0).Y;
            var table = page.TableRows.Single();
            double delta = leading - table.Y - table.Height - row.GetProperty("LeadingToTableTop").GetDouble();
            TestAssert.True(Math.Abs(delta) < .4d,
                $"Adding natural cell paragraphs must preserve the admitted table origin: {id}/{mode} delta={delta}.");
        }
    }

    public static void DocxTableMathStandaloneMatchesIndependentCellBaselineTransitions()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent cell baseline controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-cell-frame-office-baselines.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string source = Path.Combine(cases, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(source))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var renderer = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var page = renderer.InspectLayout(DocxTests.ReadDocx(source, mode)).Pages.Single();
            double leading = page.Items.Single(item => item.Kind == "TextLine" && item.SourceBlockIndex == 0).Y;
            double rightCellX = page.TableRows.Single().Cells[1].X;
            double equation = page.Items.Single(item => item.Kind == "TableRow").TextLines!
                .Single(line => line.X > rightCellX && line.SourceParagraphIndex == row.GetProperty("ParagraphIndex").GetInt32()).Y;
            double delta = leading - equation - row.GetProperty("LeadingToMathBaseline").GetDouble();
            TestAssert.True(Math.Abs(delta) < .4d,
                $"An isolated equation must share the cell's baseline transition: {id}/{mode} delta={delta}.");
        }
    }

    public static void DocxTableMathCandidatesShareStandalonePolicyAndKeepUnsupportedOwners()
    {
        var (paragraph, _) = Prepared(20d);
        var cell = new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var row = new DocxTableRow([cell], null);
        DocxDocument Document(DocxTableRow candidate) => TableDocument(candidate);
        foreach (DocxTableRow supported in new[] { row,
            row with { HeightPoints = 12d, HeightRuleValue = "exact" },
            row with { HeightPoints = 50d, HeightRuleValue = "atLeast" },
            row with { HeightPoints = 50d, HeightValue = "1000" },
            row with { HeightPoints = 0d, HeightRuleValue = "exact" } })
            TestAssert.True(ReferenceEquals(paragraph, DocxMathParagraphPolicy.EnumeratePreparationCandidates(Document(supported), CancellationToken.None)
                .Single(source => source.Equations.Count != 0)),
                "Resolved cell frames retain the same source paragraph identity as prepared math.");
        foreach (DocxTableRow unsupported in new[]
        {
            row with { HeightPoints = 12d, HeightRuleValue = "unknown" },
            row with { HeightPoints = -12d, HeightRuleValue = "exact" },
            row with { HeightPoints = double.PositiveInfinity, HeightRuleValue = "exact" },
            row with { HeightValue = "bad", HeightRuleValue = "exact" },
            row with { HeightRuleValue = "exact" },
            row with { Cells = [cell with { HasVerticalMerge = true }] },
            row with { Cells = [cell with { FitText = true }] },
            row with { Cells = [cell with { TextDirectionValue = "tbRl" }] },
            row with { Cells = [cell with { BodyElements = [new DocxPageBreakElement(DocxBreakSourceKind.RunBreak, null, null)] }] }
        })
            TestAssert.Equal(0, DocxMathParagraphPolicy.EnumeratePreparationCandidates(Document(unsupported), CancellationToken.None)
                .Count(source => source.Equations.Count != 0));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() =>
            DocxMathParagraphPolicy.EnumeratePreparationCandidates(Document(row), cancelled.Token).ToArray());
    }

    public static void DocxTableMathIsolatedInlineSharesStandaloneAdvanceAndSourceIdentity()
    {
        var (display, prepared) = Prepared(24d);
        DocxEquation equation = prepared.Equation with { Display = false };
        DocxParagraph paragraph = display with { Equations = [equation] };
        TestAssert.True(ReferenceEquals(equation, DocxMathParagraphPolicy.GetStandaloneRow(paragraph)) &&
            DocxMathParagraphPolicy.GetInlineRow(paragraph) is null, "A sole equation has one standalone owner.");
        foreach (DocxTextAlignment alignment in new[] { DocxTextAlignment.Left, DocxTextAlignment.Center, DocxTextAlignment.Right })
        {
            DocxParagraph source = paragraph with { Alignment = alignment };
            var cell = new DocxTableCell(string.Empty, [source], null, null, null, null, [], DocxTableCellMargins.Empty);
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(TableDocument(new DocxTableRow([cell], null)), new DocxTests.FamilyWidthTextMeasurer(), default,
                    mathEquations: new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
                    { [source.Equations.Single()] = prepared with { Equation = equation } });
            var row = layout.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
            var line = row.Cells.Single().TextLines.Single();
            TestAssert.True(line.Math is { } math && ReferenceEquals(math.Equation, equation) && !math.Equation.Display &&
                !line.IsInlineMath && line.Segments.Count == 0 && line.X > row.Cells[0].X + 20d,
                "Office centers isolated inline math in the existing standalone frame, independent of paragraph alignment.");
            TestAssert.True(row.Height >= 24d, "The measured cell reserves the complete standalone advance.");
        }
        TestAssert.True(DocxMathParagraphPolicy.GetStandaloneRow(paragraph with { Equations = [equation, equation] }) is null &&
            DocxMathParagraphPolicy.GetStandaloneRow(paragraph with { Runs = [new DocxTextRun("Public", 10d, null, false, false, false, null, null)] }) is null,
            "Multiple equations and visible prose retain their separate semantic ownership.");
    }

    public static void DocxTableZeroExactHeightRetainsTheNaturalContentFrame()
    {
        foreach (bool bordered in new[] { false, true })
        {
            DocxParagraph paragraph = DocxTests.CreateDocxLayoutParagraph("Public cell.", 10d, 12d);
            var cell = new DocxTableCell("Public cell.", [paragraph], null, null, null, null,
                bordered ? [new DocxTableCellBorder("bottom", "single", "243042", "4")] : [], DocxTableCellMargins.Empty);
            DocxLayout Layout(string rule)
            {
                DocxDocument source = TableDocument(new DocxTableRow([cell], 0d) { HeightRuleValue = rule });
                return new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .Create(source, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None);
            }
            DocxLayout exact = Layout("exact"), minimum = Layout("atLeast");
            var exactRow = exact.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
            var naturalRow = minimum.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
            TestAssert.True(exactRow.Height >= 12d && Math.Abs(exactRow.Height - naturalRow.Height) < .000001,
                "Zero exact height uses natural content, including its visible border, rather than a one-point clip.");
            double afterExact = exact.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
            double afterNatural = minimum.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
            TestAssert.True(Math.Abs(afterExact - afterNatural) < .000001,
                "Following flow must use the same natural zero-height frame.");
        }
    }

    public static void DocxTableMathMeasurementMemoAndFollowingFlowUseOneBoxAdvance()
    {
        var (paragraph, math) = Prepared(20d);
        DocxParagraph following = DocxTests.CreateDocxLayoutParagraph("Following cell.", 10d, 10d);
        var cell = new DocxTableCell(string.Empty, [paragraph, following], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxDocument document = TableDocument(new DocxTableRow([cell], null));
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        DocxLayout Layout(double height) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, measurer, CancellationToken.None, mathEquations:
                new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
                { [paragraph.Equations.Single()] = math with { LineHeight = height } });
        DocxLayout first = Layout(20d), expanded = Layout(30d);
        var a = first.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
        var b = expanded.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
        TestAssert.True(Math.Abs(b.Height - a.Height - 10d) < .000001,
            "The row must reserve the same extra math advance that cell placement consumes.");
        var originalLines = a.Cells.Single().TextLines;
        var expandedLines = b.Cells.Single().TextLines;
        TestAssert.True(originalLines[0].Math is { } prepared && ReferenceEquals(prepared.Box, math.Box) &&
            prepared.LineHeight == 20d && originalLines[0].Segments.Count == 0,
            "A math line carries the prepared box, not a duplicate prose encoding.");
        TestAssert.True(Math.Abs(originalLines[0].BaselineY - expandedLines[0].BaselineY) < .000001 &&
            Math.Abs(originalLines[1].BaselineY - expandedLines[1].BaselineY - 10d) < .000001,
            "Changing only occupied math height preserves its baseline and shifts following cell prose.");
        double afterA = first.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
        double afterB = expanded.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
        TestAssert.True(Math.Abs(afterA - afterB - 10d) < .000001,
            "Following body flow must observe the same measured row extent, including a fresh memo per layout.");
    }

    public static void DocxTableMathKeepEstimateUsesThePreparedFirstRowExtent()
    {
        var (paragraph, math) = Prepared(15d);
        var cell = new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxDocument source = TableDocument(new DocxTableRow([cell], null));
        DocxParagraph filler = DocxTests.CreateDocxLayoutParagraph("Fill", 10d, 30d);
        DocxParagraph keep = DocxTests.CreateDocxLayoutParagraph("Keep", 10d, 10d,
            new DocxParagraphKeepRules(true, null, null, null, null, null));
        var table = source.BodyElements.OfType<DocxTableElement>().Single();
        DocxDocument document = source with { PageHeightPoints = 80d,
            BodyElements = [new DocxParagraphElement(filler), new DocxParagraphElement(keep), table] };
        foreach ((double height, int pages) in new[] { (15d, 1), (25d, 2) })
        {
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), CancellationToken.None,
                    mathEquations: new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
                    { [paragraph.Equations.Single()] = math with { LineHeight = height } });
            TestAssert.Equal(pages, layout.Pages.Count);
            var tablePage = layout.Pages.Single(page => page.Items.OfType<DocxTableRowLayout>().Any());
            TestAssert.True(tablePage.Items.OfType<DocxTextLineLayout>().Any(line => line.Text == "Keep"),
                "Keep-with-next and math row placement must move as the same measured block.");
        }
    }

    public static void DocxTableMathInlineMeasurementAndWrappingRetainSourceOwnership()
    {
        var (display, prepared) = Prepared(24d);
        DocxTextRun before = new("Before ", 10d, null, false, false, false, null, "TestFont");
        DocxTextRun after = before with { Text = " after public equation with more prose." };
        DocxEquation equation = prepared.Equation with { Display = false, SourceRunIndex = 1 };
        DocxParagraph paragraph = display with { Runs = [before, after], Equations = [equation] };
        var math = prepared with { Equation = equation };
        var cell = new DocxTableCell(string.Concat(paragraph.Runs.Select(run => run.Text)), [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var document = TableDocument(new DocxTableRow([cell], null));
        TestAssert.True(ReferenceEquals(paragraph, DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, default)
            .Single(source => source.Equations.Count != 0)), "Inline cells share the body semantic owner policy.");
        DocxLayout Layout(double height) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations:
                new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
                { [paragraph.Equations.Single()] = math with { LineHeight = height } });
        DocxLayout initial = Layout(24d), expanded = Layout(34d);
        var a = initial.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
        var b = expanded.Pages.Single().Items.OfType<DocxTableRowLayout>().Single();
        var lines = a.Cells.Single().TextLines;
        TestAssert.True(lines.Count > 1 && lines.Count(line => line.Math is not null) == 1 &&
            lines[0].IsInlineMath && lines[0].InlineMathOffsetX > 0d,
            "A complete box stays at its source boundary while its following prose wraps.");
        string prose = string.Concat(lines.Select(line => line.Text));
        TestAssert.True(!prose.Contains('\u200B') && new[] { "Before", "after", "public", "equation", "with", "more", "prose." }
            .All(word => prose.Contains(word, StringComparison.Ordinal)), "Every public prose word survives without layout markers.");
        TestAssert.True(Math.Abs(b.Height - a.Height - 10d) < .000001,
            "Cell measurement reserves exactly the extra descent consumed by inline placement.");
        double afterA = initial.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
        double afterB = expanded.Pages[0].Items.OfType<DocxTextLineLayout>().Single().BaselineY;
        TestAssert.True(Math.Abs(afterA - afterB - 10d) < .000001, "Following body flow shares the measured row advance.");
        TestAssert.Equal(before.Text + after.Text, string.Concat(paragraph.Runs.Select(run => run.Text)));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), cancelled.Token,
                mathEquations: new Dictionary<DocxEquation, DocxPreparedMathParagraph> { [paragraph.Equations.Single()] = math }));
    }

    public static void DocxTableMathInlineUnfitBoxesPreserveOrdinaryRowSplitting()
    {
        var (display, prepared) = Prepared(24d);
        DocxTextRun run = new(string.Join(" ", Enumerable.Repeat("Public prose", 20)), 10d, null, false, false, false, null, "TestFont");
        DocxEquation equation = prepared.Equation with { Display = false, SourceRunIndex = 1 };
        DocxParagraph paragraph = display with { Runs = [run], Equations = [equation] };
        var cell = new DocxTableCell(string.Concat(paragraph.Runs.Select(run => run.Text)), [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        DocxDocument source = TableDocument(new DocxTableRow([cell], 150d) { HeightRuleValue = "exact" }) with { PageHeightPoints = 100d };
        var engine = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        DocxLayout plain = engine.Create(source, measurer, default);
        DocxLayout fallback = engine.Create(source, measurer, default, mathEquations:
            new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
            { [paragraph.Equations.Single()] = prepared with { Equation = equation, Box = prepared.Box with { Width = 1000d } } });
        TestAssert.Equal(plain.Pages.Count, fallback.Pages.Count);
        var ordinaryRows = plain.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
        var fallbackRows = fallback.Pages.SelectMany(page => page.Items.OfType<DocxTableRowLayout>()).ToArray();
        TestAssert.True(ordinaryRows.Length > 1, "This control must exercise actual row fragments across pages.");
        TestAssert.Equal(ordinaryRows.Length, fallbackRows.Length);
        for (int index = 0; index < ordinaryRows.Length; index++)
        {
            TestAssert.True(Math.Abs(ordinaryRows[index].Height - fallbackRows[index].Height) < .000001,
                "An unpainted prepared box cannot make an otherwise splittable prose row atomic.");
            TestAssert.True(fallbackRows[index].Cells.SelectMany(c => c.TextLines).All(line => line.Math is null),
                "An unfit equation retains the whole fallback.");
        }
    }

    public static void DocxStaticMathFramesSharePreparedAdvanceAndFooterBlockTranslation()
    {
        var (paragraph, math) = Prepared(20d);
        DocxParagraph following = DocxTests.CreateDocxLayoutParagraph("Following static prose.", 10d, 10d);
        var sources = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [paragraph, following] };
        foreach (bool header in new[] { true, false })
        {
            DocxPageSettings settings = DocxPageSettings.Empty with
            {
                HeaderDistancePoints = 12d, FooterDistancePoints = 12d,
                HeaderParagraphsByType = header ? sources : new Dictionary<string, IReadOnlyList<DocxParagraph>>(),
                FooterParagraphsByType = header ? new Dictionary<string, IReadOnlyList<DocxParagraph>>() : sources
            };
            DocxDocument document = DocxTests.CreateLayoutTestDocument(
                [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body.", 10d, 10d))], []) with { PageSettings = settings };
            DocxLayout Layout(double height) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations:
                    new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance) { [paragraph.Equations.Single()] = math with { LineHeight = height } });
            DocxLayout initial = Layout(20d), expanded = Layout(30d);
            var a = initial.Pages.Single().StaticTextLines; var z = expanded.Pages.Single().StaticTextLines;
            TestAssert.Equal(2, a.Count); TestAssert.Equal(2, z.Count);
            TestAssert.True(a[0].Math is { } prepared && ReferenceEquals(prepared.Equation, math.Equation) && a[0].Story is not null,
                "Static math retains the original prepared source and selected story carrier.");
            if (header)
                TestAssert.True(Math.Abs(a[1].BaselineY - z[1].BaselineY - 10d) < .000001 &&
                    Math.Abs(initial.HeaderContentBottomByPage[0] - expanded.HeaderContentBottomByPage[0] - 10d) < .000001,
                    "Header math reserves the same extra advance in following prose and the body keep-out edge.");
            else
                TestAssert.True(Math.Abs(a[1].BaselineY - z[1].BaselineY) < .000001 &&
                    Math.Abs(z[0].BaselineY - a[0].BaselineY - 10d) < .000001 &&
                    Math.Abs(expanded.FooterContentTopByPage[0] - initial.FooterContentTopByPage[0] - 10d) < .000001,
                    "Footer translation anchors its following prose while growing math and the body keep-out edge upward.");
        }
    }

    public static void DocxStaticFrameMappingKeepsFooterAnchoredAndUsesOneHeaderAdvance()
    {
        const double scale = .75d;
        var (paragraph, prepared) = Prepared(20d);
        DocxParagraph following = DocxTests.CreateDocxLayoutParagraph("Following static prose.", 10d, 10d) with { LineSpacingPoints = null };
        var paragraphs = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [paragraph, following] };
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        foreach (bool header in new[] { true, false })
        {
            DocxPageSettings settings = DocxPageSettings.Empty with
            {
                HeaderDistancePoints = 12d, FooterDistancePoints = 12d,
                HeaderParagraphsByType = header ? paragraphs : new Dictionary<string, IReadOnlyList<DocxParagraph>>(),
                FooterParagraphsByType = header ? new Dictionary<string, IReadOnlyList<DocxParagraph>>() : paragraphs
            };
            DocxDocument document = DocxTests.CreateLayoutTestDocument(
                [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body.", 10d, 10d))], []) with { PageSettings = settings };
            DocxLayout Layout(double height) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, scale)
                .Create(document, new DocxTests.ScaledLayoutTextMeasurer(measurer, scale, scale), default, measurer,
                    mathEquations: new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance)
                    { [paragraph.Equations.Single()] = prepared with { LineHeight = height } });
            DocxLayout a = Layout(20d), z = Layout(30d);
            var first = a.Pages.Single().StaticTextLines; var expanded = z.Pages.Single().StaticTextLines;
            TestAssert.True(first.Count == 2 && first.All(line => line.StaticStoryUsesPhysicalY) &&
                first[0].Math is { } math && ReferenceEquals(math.Equation, prepared.Equation),
                "A complete natural static source maps one original prepared equation and its prose.");
            if (header)
                TestAssert.True(Math.Abs(first[1].BaselineY - expanded[1].BaselineY - 7.5d) < .000001 &&
                    Math.Abs(a.HeaderContentBottomByPage[0] - z.HeaderContentBottomByPage[0] - 7.5d) < .000001,
                    "The same seven-and-a-half-point growth moves following prose and the body keep-out edge.");
            else
                TestAssert.True(Math.Abs(first[1].BaselineY - expanded[1].BaselineY) < .000001 &&
                    Math.Abs(expanded[0].BaselineY - first[0].BaselineY - 7.5d) < .000001 &&
                    Math.Abs(z.FooterContentTopByPage[0] - a.FooterContentTopByPage[0] - 7.5d) < .000001,
                    "Footer growth moves upward, preserving its trailing prose and source-bound frame edge.");
        }
    }

    public static void DocxStaticFrameMappingRetainsIncompleteAndDeclaredSourceFallbacks()
    {
        var (paragraph, prepared) = Prepared(20d);
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        foreach (string guard in new[] { "field", "declared", "table-first", "missing-metrics", "unprepared-math" })
        {
            DocxParagraph footer = DocxTests.CreateDocxLayoutParagraph(guard == "field" ? "{PAGE}" : "Public footer.", 10d, 10d);
            if (guard != "declared") footer = footer with { LineSpacingPoints = null };
            DocxPageSettings settings = DocxPageSettings.Empty with
            {
                HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [paragraph] },
                FooterParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [footer] }
            };
            var cell = new DocxTableCell("Public cell.", [footer], null, null, null, null, [], DocxTableCellMargins.Empty);
            DocxDocument document = guard == "table-first" ? TableDocument(new DocxTableRow([cell], null)) :
                DocxTests.CreateLayoutTestDocument([new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body.", 10d, 10d))], []);
            document = document with { PageSettings = settings };
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, .75d)
                .Create(document, new DocxTests.ScaledLayoutTextMeasurer(measurer, .75d, .75d), default,
                    guard == "missing-metrics" ? null : measurer,
                    mathEquations: guard == "unprepared-math" ? null :
                        new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance) { [paragraph.Equations.Single()] = prepared });
            TestAssert.True(layout.Pages.SelectMany(page => page.StaticTextLines).All(line => !line.StaticStoryUsesPhysicalY),
                "Fields, declared frames, table-first bodies and unresolved metrics/math keep their existing coordinate owner.");
        }
    }

    public static void DocxStaticFrameMappingSeparatesAlignmentFromPlainContentOwnership()
    {
        DocxParagraph source = DocxTests.CreateDocxLayoutParagraph("Public static prose.", 10d, 10d) with { LineSpacingPoints = null };
        var measurer = new DocxTests.FamilyWidthTextMeasurer();
        var positions = new List<double>();
        foreach (DocxTextAlignment alignment in new[] { DocxTextAlignment.Left, DocxTextAlignment.Center, DocxTextAlignment.Right, DocxTextAlignment.Justified })
        {
            DocxParagraph paragraph = source with { Alignment = alignment };
            TestAssert.True(DocxPlainBodyFrame.HasPlainTextContent(paragraph), "Alignment does not change plain content ownership.");
            TestAssert.Equal(alignment == DocxTextAlignment.Left, DocxPlainBodyFrame.IsPlainParagraph(paragraph));
            DocxPageSettings settings = DocxPageSettings.Empty with
            { HeaderParagraphsByType = new Dictionary<string, IReadOnlyList<DocxParagraph>> { ["default"] = [paragraph] } };
            DocxDocument document = DocxTests.CreateLayoutTestDocument(
                [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body.", 10d, 10d))], []) with { PageSettings = settings };
            DocxLayout layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup, .75d)
                .Create(document, new DocxTests.ScaledLayoutTextMeasurer(measurer, .75d, .75d), default, measurer);
            var line = layout.Pages.Single().StaticTextLines.Single();
            TestAssert.Equal(alignment != DocxTextAlignment.Justified, line.StaticStoryUsesPhysicalY);
            if (alignment != DocxTextAlignment.Justified) positions.Add(line.X);
            TestAssert.Equal(alignment, paragraph.Alignment);
        }
        TestAssert.True(positions[0] < positions[1] && positions[1] < positions[2],
            "Already supported left/center/right placement shares one source rule while body/note eligibility remains left-only.");
    }

    public static void DocxStaticTableMathAndPlainFramesMatchIndependentWordAnchors()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent anchors use installed Windows font metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-static-cell-office-geometry.json")));
        foreach (var group in reference.RootElement.GetProperty("Rows").EnumerateArray()
            .GroupBy(row => (Id: row.GetProperty("Id").GetString()!, Mode: row.GetProperty("Mode").GetString())))
        {
            string input = Path.Combine(cases, group.Key.Id + ".docx");
            TestAssert.Equal(group.First().GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = group.Key.Mode == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var pages = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderBlankPages(DocxTests.ReadDocx(input, mode), null, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages.Single()).ToArray();
            foreach (var row in group)
            {
                string prefix = row.GetProperty("Container").GetString() + ":";
                var show = shows.Single(item => item.Text.StartsWith(prefix, StringComparison.Ordinal));
                TestAssert.True(Math.Abs(show.X - row.GetProperty("X").GetDouble()) < .2d &&
                    Math.Abs(show.Y - row.GetProperty("BaselineY").GetDouble()) < .2d,
                    "Source-backed body/cell/static placement retains the independent Word anchor: " + group.Key + "/" + prefix);
            }
        }
    }

    public static void DocxStaticTableMathPreparationSharesCellEligibilityAndSourceIdentity()
    {
        var (paragraph, _) = Prepared(20d);
        var cell = new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var row = new DocxTableRow([cell], null);
        foreach (bool header in new[] { true, false })
        foreach (string guard in new[] { "plain", "exact", "merge", "fit", "vertical", "nested", "height", "note" })
        {
            DocxTableCell source = guard switch
            {
                "merge" => cell with { HasVerticalMerge = true },
                "fit" => cell with { FitText = true },
                "vertical" => cell with { TextDirectionValue = "tbRl" },
                "nested" => cell with { BodyElements = [new DocxTableElement(new DocxTable(null, [100d], [row]))] },
                _ => cell
            };
            var table = new DocxTable(null, [100d], [row with
            {
                Cells = [source], HeightPoints = guard == "exact" ? 50d : null,
                HeightRuleValue = guard == "height" ? "unknown" : guard == "exact" ? "exact" : null
            }]);
            IReadOnlyList<DocxBodyElement> elements = [new DocxTableElement(table)];
            var stories = new Dictionary<string, IReadOnlyList<DocxBodyElement>> { ["default"] = elements, ["first"] = elements };
            var settings = DocxPageSettings.Empty with
            {
                HeaderBodyElementsByType = header ? stories : new Dictionary<string, IReadOnlyList<DocxBodyElement>>(),
                FooterBodyElementsByType = header ? new Dictionary<string, IReadOnlyList<DocxBodyElement>>() : stories
            };
            var document = DocxTests.CreateLayoutTestDocument([], []) with
            {
                PageSettings = guard == "note" ? DocxPageSettings.Empty : settings,
                EquationParts = [new DocxEquationPartSource(guard == "note" ? "/word/footnotes.xml" : "/word/static.xml", 1, elements)]
            };
            var candidates = DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default).ToArray();
            TestAssert.Equal(guard is "plain" or "exact" ? 1 : 0, candidates.Length);
            if (candidates.Length != 0)
                TestAssert.True(ReferenceEquals(paragraph, candidates[0].Paragraph) && candidates[0].HasPageInputs,
                    "Shared variants prepare one original source owner with actual page inputs.");
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() =>
                DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, cancelled.Token).ToArray());
        }
    }

    public static void DocxStaticTableMathOccupiedHeightUsesPreparedCellFrames()
    {
        var (paragraph, math) = Prepared(20d);
        var cell = new DocxTableCell(string.Empty, [paragraph], null, null, null, null, [], DocxTableCellMargins.Empty);
        var table = new DocxTable(null, [100d], [new DocxTableRow([cell], null)]) { PreferredWidthPoints = 100d };
        var stories = new Dictionary<string, IReadOnlyList<DocxBodyElement>> { ["default"] = [new DocxTableElement(table)] };
        foreach (bool header in new[] { true, false })
        {
            var document = DocxTests.CreateLayoutTestDocument(
                [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Body.", 10d, 10d))], []) with
            {
                PageSettings = DocxPageSettings.Empty with
                {
                    HeaderBodyElementsByType = header ? stories : new Dictionary<string, IReadOnlyList<DocxBodyElement>>(),
                    FooterBodyElementsByType = header ? new Dictionary<string, IReadOnlyList<DocxBodyElement>>() : stories
                }
            };
            DocxLayout Layout(double height) => new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations:
                    new Dictionary<DocxEquation, DocxPreparedMathParagraph> { [math.Equation] = math with { LineHeight = height } });
            var a = Layout(20d); var b = Layout(40d);
            var first = a.Pages.Single().StaticTableRows.Single();
            var expanded = b.Pages.Single().StaticTableRows.Single();
            TestAssert.True(first.Cells.Single().TextLines.Single().Math is { } placed &&
                ReferenceEquals(placed.Equation, math.Equation), "Static cells consume their original prepared box.");
            TestAssert.True(Math.Abs(expanded.Height - first.Height - 20d) < .000001,
                "Grid measurement and cell painting reserve the same prepared height.");
            if (!header)
            {
                TestAssert.True(Math.Abs(a.FooterContentTopByPage[0] - first.Y - first.Height) < .000001 &&
                    Math.Abs(b.FooterContentTopByPage[0] - expanded.Y - expanded.Height) < .000001,
                    "The footer keep-out is the row's top edge; its stored Y is the bottom.");
            }
        }
    }

    private static (DocxParagraph Paragraph, DocxPreparedMathParagraph Math) Prepared(double height)
    {
        var run = new DocxTextRun("A+B", 10d, null, false, false, false, null, "TestFont");
        var row = new DocxMathRow([new DocxMathRun(run, "p", null)]);
        var equation = new DocxEquation(row, 0, 0, true, null, null, null) { ParentFontSize = 10d };
        var font = new PdfFontResource("Eq1", PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(),
            run.Text.EnumerateRunes().Select(rune => rune.Value), CancellationToken.None));
        TestAssert.True(DocxMathRowLayout.TryMeasure(row, _ => font, CancellationToken.None, out var box, out _), "Expected portable math glyphs.");
        var paragraph = DocxTests.CreateDocxLayoutParagraph(string.Empty, 10d, 10d) with
            { Runs = [], LineSpacingPoints = null, Equations = [equation] };
        return (paragraph, new(equation, TestAssert.NotNull(box), 12d, height));
    }

    private static DocxDocument TableDocument(DocxTableRow row)
    {
        var table = new DocxTable(null, [100d], [row]) { PreferredWidthPoints = 100d };
        return DocxTests.CreateLayoutTestDocument([new DocxTableElement(table),
            new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Following body.", 10d, 10d))], [table]);
    }
}
