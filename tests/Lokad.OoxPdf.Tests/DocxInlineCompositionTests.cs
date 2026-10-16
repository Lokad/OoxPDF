using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionOrderedPictureTranslationKeepsClipsAndFallbacks()
    {
        var source = ReadMutation("CompositionProbes/math-picture-first-table-single-cell", OoxPdfDocxMarkupMode.Final, false, false);
        var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(source.BodyElements).Single(p => p.Equations.Count != 0);
        var equations = paragraph.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 8d, 2d, []), 10d, 16d) { UsesDisplayStyle = false });
        var measurer = new TrackingMeasurer();
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(source, measurer, default, mathEquations: equations);
        var row = layout.Pages[0].Items.OfType<DocxTableRowLayout>().Single();
        var cell = row.Cells.Single();
        var picture = cell.InlineImages.Single();
        TestAssert.True(picture.SourceLineIndex == cell.TextLines.Single().SourceLineIndex,
            "A picture retains the measured source-line owner without copying its geometry.");
        picture = picture with { ClipBounds = new(picture.X, picture.Y, picture.Width, picture.Height),
            LineFrame = new(picture.Y - 2d, picture.Height + 2d) };
        row = row with { Cells = [cell with { InlineImages = [picture] }] };
        var items = new List<DocxLayoutItem> { row };
        double cursor = 0d;
        DocxLayoutEngine.ApplyReviewTableBaselineGeometry(items, 0, 12d, measurer, .75d, ref cursor, default);
        var moved = ((DocxTableRowLayout)items[0]).Cells.Single();
        var image = moved.InlineImages.Single();
        TestAssert.True(Math.Abs(image.Y - picture.Y + 3d) < .000001d &&
            Math.Abs(image.ClipBounds!.Value.Y - picture.ClipBounds!.Value.Y + 3d) < .000001d &&
            Math.Abs(image.LineFrame!.Value.Bottom - picture.LineFrame!.Value.Bottom + 3d) < .000001d &&
            Math.Abs(moved.TextLines.Single().BaselineY - cell.TextLines.Single().BaselineY + 3d) < .000001d,
            "Line, picture, crop clip and optional picture frame translate together exactly once.");
        foreach (int? missingOwner in new int?[] { null, 999 })
        {
            var invalid = row with { RowIndex = 1, Cells = [cell with
                { InlineImages = [picture with { SourceLineIndex = missingOwner }] }] };
            var fallback = new List<DocxLayoutItem> { row, invalid };
            cursor = 0d;
            DocxLayoutEngine.ApplyReviewTableBaselineGeometry(fallback, 0, 12d, measurer, .75d, ref cursor, default);
            TestAssert.True(ReferenceEquals(row, fallback[0]) && ReferenceEquals(invalid, fallback[1]) && cursor == 0d,
                "An unowned picture retains the whole table fallback before any earlier row is mutated.");
        }
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() =>
            DocxLayoutEngine.ApplyReviewTableBaselineGeometry(items, 0, 12d, measurer, .75d, ref cursor, canceled.Token));
    }

    public static void DocxInlineCompositionDisplayImageRowsRetainIndependentOfficeFractionFrames()
        => AssertPictureWordFrames("math-display-images-office-geometry.json");

    public static void DocxInlineCompositionPictureFirstTablesUseTheirOwnedTextAnchor()
    {
        AssertPictureWordFrames("math-picture-first-tables-office-geometry.json");
        if (OperatingSystem.IsWindows()) CheckPictureFirstTables(null);
    }

    internal static void CheckPictureFirstTables(Lokad.OoxPdf.Fonts.IFontResolver? resolver)
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            double? baseline = null;
            foreach (string shape in new[] { "two-cells", "single-cell" })
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation("CompositionProbes/math-picture-first-table-" + shape, mode, transformed, transformed);
                var owner = DocxBlockTraversal.EnumerateBodyParagraphs(source.BodyElements).Single(p => p.Equations.Count != 0);
                TestAssert.True(owner.Runs.Count == 0 && owner.ParagraphMarkRun is null && owner.ParagraphMarkFontSize == 12d,
                    "Unpainted paragraph-mark metrics remain available for the object-only source owner.");
                var page = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                    ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                    .RenderBlankPages(source, null, default).Single();
                TestAssert.Equal(1, page.Images.Count);
                var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
                var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
                TestAssert.Equal(16, math.Length);
                TestAssert.Equal("ratetimeratetime", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                baseline ??= math[0].Y;
                TestAssert.True(Math.Abs(math[0].Y - baseline.Value) < .01d,
                    "Removing a neighboring prose cell or adding source metadata cannot change the picture/math baseline.");
                TestAssert.Contains("Following prose.", DocxPageFootnoteNumberingTests.DecodeText(page));
            }
        }
    }

    private static void AssertPictureWordFrames(string referenceFile)
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent controls use installed Arial metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, referenceFile)));
        foreach (var row in reference.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var page = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(DocxTests.ReadDocx(input, mode), null, default).Single();
            TestAssert.Equal(1, page.Images.Count);
            var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
            TestAssert.Equal(row.GetProperty("MathText").GetString(), string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(16, math.Length);
            var anchors = row.GetProperty("Anchors").EnumerateArray().ToArray();
            double origin = anchors[0].GetProperty("X").GetDouble();
            foreach (var anchor in anchors)
            {
                int index = anchor.GetProperty("Occurrence").GetInt32() * 8 + (anchor.GetProperty("Text").GetString() == "rate" ? 0 : 4);
                TestAssert.True(Math.Abs(math[index].X - anchor.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(math[index].X - math[0].X - (anchor.GetProperty("X").GetDouble() - origin)) < .2d,
                    "Word's inline fraction geometry and separation survive adjacent source equation boundaries.");
                // Body, cell and table-first placements share the measured text strut.
                // Whole-page visual thresholds remain independent and unchanged.
                TestAssert.True(Math.Abs(math[index].Y - anchor.GetProperty("Y").GetDouble()) < .2d,
                    "Owned picture/math baselines retain independent Word placement.");
                TestAssert.True(Math.Abs(math[index].Size / 12d - anchor.GetProperty("FontSize").GetDouble() / row.GetProperty("ProseFontSize").GetDouble()) < .01d,
                    "A visible image selects independently observed inline argument sizing.");
            }
        }
    }

    public static void DocxInlineCompositionDisplayImageRowsPreserveSourceAndFractionSeparation()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native controls use installed math fonts.");
        CheckDisplayImageRows(null);
    }

    internal static void CheckDisplayImageRows(Lokad.OoxPdf.Fonts.IFontResolver? resolver)
    {
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var combined = renderer.RenderBlankPages(ReadMutation($"CompositionProbes/math-display-image-{container}-combined", mode, false, false), null, default).Single();
            var combinedFonts = combined.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
            var expected = DocxCommentsTests.ReadEmbeddedGlyphTextShows(combined, usePageCoordinates: true).Where(show => combinedFonts.Contains(show.Font)).ToArray();
            foreach (string position in new[] { "before", "between", "combined" })
            foreach (bool transformed in new[] { false, true })
            {
                var source = ReadMutation($"CompositionProbes/math-display-image-{container}-{position}", mode, transformed, transformed);
                var owner = DocxBlockTraversal.EnumerateBodyParagraphs(source.BodyElements).Single(paragraph => paragraph.Equations.Count != 0);
                TestAssert.Equal(position == "combined" ? 1 : 2, owner.Equations.Count);
                TestAssert.True(owner.Equations.All(equation => equation.Display), "Source display flags remain attached to the original formulas.");
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = renderer.RenderBlankPages(source, diagnostics.Add, default).Single();
                TestAssert.Equal(1, page.Images.Count);
                var picture = System.Text.RegularExpressions.Regex.Match(page.Content,
                    @"(?<width>-?[\d.]+) 0 0 (?<height>-?[\d.]+) (?<x>-?[\d.]+) (?<y>-?[\d.]+) cm\s*/" + page.Images.Single().ResourceName + @" Do");
                TestAssert.True(picture.Success, "The real image must reach PDF painting with its own placement.");
                var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
                var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
                TestAssert.Equal("ratetimeratetime", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(16, math.Length);
                if (position != "between")
                {
                    TestAssert.Equal(16, expected.Length);
                    for (int index = 0; index < math.Length; index++)
                        TestAssert.True(Math.Abs(math[index].X - expected[index].X) < .01d && Math.Abs(math[index].Y - expected[index].Y) < .01d,
                            "Merging adjacent source equations must preserve fraction spacing and placement: " + container + "/" + mode + "/" + position + "/" + index + ": " + math[index].X + "," + math[index].Y + " versus " + expected[index].X + "," + expected[index].Y + ".");
                }
                else TestAssert.True(math[8].X - math[0].X > owner.Images.Single().WidthPoints * (mode == OoxPdfDocxMarkupMode.Final ? 1d : .7d),
                    "The image retains its source position between both fractions.");
                TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "A visible picture cannot discard supported equations.");
                TestAssert.Contains("Following prose.", DocxPageFootnoteNumberingTests.DecodeText(page));
            }
        }
    }

    public static void DocxInlineCompositionMixedDisplaySourceRolesFollowAuthoredRows()
    {
        var document = ReadMutation("CompositionProbes/math-mixed-display-literal-all", OoxPdfDocxMarkupMode.Final, false, false);
        var source = document.Paragraphs.Single(paragraph => paragraph.Equations.Count == 2);
        var prose = source.Runs.First(run => run.Text.Contains("Before formula:", StringComparison.Ordinal));
        var paragraph = source with { InlineContent = [new DocxInlineText(prose, 0, prose.Text.Length),
            new DocxInlineEquation(source.Equations[0]), new DocxInlineBreak(prose, 0), new DocxInlineEquation(source.Equations[1])] };
        var mixed = DocxMathParagraphPolicy.GetMixedSourceEquations(paragraph, default);
        var display = DocxMathParagraphPolicy.GetDisplaySourceEquations(paragraph, default);
        TestAssert.Equal(1, mixed.Count);
        TestAssert.Equal(1, display.Count);
        TestAssert.True(mixed.Contains(source.Equations[0]) && display.Contains(source.Equations[1]),
            "Authored row boundaries keep the prose formula inline and the following math-only row in display style.");
        TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathParagraphPolicy.GetMixedSourceEquations(paragraph, canceled.Token));
    }

    public static void DocxInlineCompositionMixedDisplayRowsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent controls use installed math fonts.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-mixed-display-office-geometry.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var page = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(DocxTests.ReadDocx(input, mode), null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
            var math = shows.Where(show => fonts.Contains(show.Font)).ToArray();
            TestAssert.Equal(row.GetProperty("MathText").GetString(), string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(row.GetProperty("Id").GetString()!.Contains("fraction", StringComparison.Ordinal) ? 16 : 10, math.Length);
            var anchors = row.GetProperty("Anchors").EnumerateArray().ToArray();
            double origin = anchors[0].GetProperty("X").GetDouble();
            foreach (var anchor in anchors)
            {
                string text = anchor.GetProperty("Text").GetString()!;
                int occurrence = anchor.GetProperty("Occurrence").GetInt32();
                var glyph = text is "rate" or "time" ? math[occurrence * 8 + (text == "rate" ? 0 : 4)]
                    : math.Where(show => show.Text.Normalize(NormalizationForm.FormKC) == text).ElementAt(occurrence);
                // Retain the existing table-origin/footer-baseline residuals.
                TestAssert.True(Math.Abs(glyph.X - anchor.GetProperty("X").GetDouble()) < 1.25d &&
                    Math.Abs(glyph.Y - anchor.GetProperty("Y").GetDouble()) < 1d &&
                    Math.Abs(glyph.X - math[0].X - (anchor.GetProperty("X").GetDouble() - origin)) < .75d,
                    "Mixed explicit formulas retain independent Word geometry: " + row.GetProperty("Id").GetString() + ".");
                // These pinned sources use 12pt text/math. Printed cell prose
                // scales its Tf while math uses cm; the reader keeps raw Tf.
                TestAssert.True(Math.Abs(glyph.Size / 12d - anchor.GetProperty("FontSize").GetDouble() / row.GetProperty("ProseFontSize").GetDouble()) < .01d,
                    "Prose-row ownership selects independently observed inline argument sizing.");
            }
        }
    }

    public static void DocxInlineCompositionMixedDisplayRowsRetainSourceAndResolvedInlineStyle()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native controls use installed math fonts.");
        CheckMixedDisplaySourceRows(null);
    }

    public static void DocxInlineCompositionMixedDisplayTotalPageCountersRetainSupportedMath()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native controls use installed math fonts.");
        CheckTotalPageMath(null, explicitSource: true);
    }

    internal static void CheckMixedDisplaySourceRows(Lokad.OoxPdf.Fonts.IFontResolver? resolver)
    {
        // The pinned portable font has upright Greek glyphs rather than the
        // native fixtures' mathematical italic alphabet. Both sides receive
        // the same explicit test-only styling before the source transformation.
        Action<XDocument, XNamespace, XNamespace>? portableStyle = resolver is null ? null : (xml, w, m) =>
        {
            foreach (var run in xml.Descendants(m + "r"))
            {
                var properties = run.Element(m + "rPr");
                if (properties is null) { properties = new XElement(m + "rPr"); run.AddFirst(properties); }
                properties.Elements(m + "sty").Remove();
                properties.Add(new XElement(m + "sty", new XAttribute(m + "val", "p")));
            }
        };
        foreach (string suffix in new[] { "literal-first", "literal-all", "fraction-first", "fraction-all", "cell", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = "CompositionProbes/math-mixed-display-" + suffix;
            string? part = suffix == "footer" ? "word/math-tabs-footer.xml" : null;
            var source = ReadMutation(fixture, mode, false, false, portableStyle, partOverride: part);
            var inline = ReadMutation(fixture, mode, false, false, (xml, w, m) =>
            {
                portableStyle?.Invoke(xml, w, m);
                foreach (var group in xml.Descendants(m + "oMathPara").ToArray())
                    group.ReplaceWith(group.Elements(m + "oMath").Select(equation => new XElement(equation)));
            }, partOverride: part);
            var owner = suffix == "footer" ? source.FooterParagraphs.Single(p => p.Equations.Count == 2)
                : DocxBlockTraversal.EnumerateBodyParagraphs(source.BodyElements).Single(p => p.Equations.Count == 2);
            var originalFlags = owner.Equations.Select(equation => equation.Display).ToArray();
            TestAssert.True(originalFlags.Any(value => value), "The source retains explicit display containers.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetMixedSourceEquations(owner, default).Count);
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.RenderBlankPages(inline, null, default).Single();
            var baselineFonts = before.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
            var expected = DocxCommentsTests.ReadEmbeddedGlyphTextShows(before, usePageCoordinates: true).Where(show => baselineFonts.Contains(show.Font)).ToArray();
            foreach (bool transformed in new[] { false, true })
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = renderer.RenderBlankPages(transformed ? ReadMutation(fixture, mode, true, true, portableStyle, partOverride: part) : source,
                    diagnostics.Add, default).Single();
                var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
                var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
                TestAssert.Equal(suffix.StartsWith("fraction", StringComparison.Ordinal) ? 16 : 10, actual.Length);
                TestAssert.Equal(expected.Length, actual.Length);
                for (int i = 0; i < actual.Length; i++)
                    TestAssert.True(actual[i].Text == expected[i].Text && Math.Abs(actual[i].X - expected[i].X) < .001d &&
                        Math.Abs(actual[i].Y - expected[i].Y) < .001d && Math.Abs(actual[i].Size - expected[i].Size) < .001d,
                        "Explicit containers beside prose keep the equivalent inline geometry across source transformations.");
                TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "A source display flag cannot discard otherwise supported prose-row content.");
            }
            TestAssert.True(originalFlags.SequenceEqual(owner.Equations.Select(equation => equation.Display)), "Resolved style cannot mutate source containers.");
        }
    }

    public static void DocxInlineCompositionTotalPageMathRetainsIndependentWordGeometry()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent controls use installed Arial metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-total-pages-office-geometry.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(DocxTests.ReadDocx(input, mode), null, default).ToArray();
            TestAssert.Equal(2, pages.Length);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0], usePageCoordinates: true).ToArray();
            var fonts = pages[0].Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
            var math = shows.Where(show => fonts.Contains(show.Font)).ToArray();
            TestAssert.Equal(row.GetProperty("MathText").GetString(), string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(16, math.Length);
            var anchors = row.GetProperty("Anchors").EnumerateArray().ToArray();
            double originX = anchors[0].GetProperty("X").GetDouble();
            // The existing cell-origin and simple-field Calibri/Arial advance
            // residuals accumulate to about two points; retain the failed visual
            // gates and check fraction geometry independently of that offset.
            TestAssert.True(Math.Abs(math[0].X - originX) < 2.1d, "Retain the measured Office field/cell-origin residual.");
            foreach (var anchor in anchors)
            {
                int index = anchor.GetProperty("Occurrence").GetInt32() * 8 + (anchor.GetProperty("Word").GetString() == "rate" ? 0 : 4);
                TestAssert.True(Math.Abs(math[index].Y - anchor.GetProperty("Y").GetDouble()) < .2d &&
                    Math.Abs(math[index].X - math[0].X - (anchor.GetProperty("X").GetDouble() - originX)) < .75d,
                    "Resolved total-page fields retain Word's fraction baselines and local geometry.");
            }
            var counters = shows.Where(show => show.Text.Trim() == "2").ToArray();
            TestAssert.Equal(2, counters.Length);
            foreach (var counter in counters)
                TestAssert.True(Math.Abs(counter.Y - row.GetProperty("Start").GetProperty("Y").GetDouble()) < .2d,
                    "Actual count digits retain the prose baseline.");
        }
    }

    public static void DocxInlineCompositionTotalPageMathKeepsCanonicalCountersAndSourceTransformations()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent native controls use installed Arial metrics.");
        CheckTotalPageMath(null);
    }

    internal static void CheckTotalPageMath(Lokad.OoxPdf.Fonts.IFontResolver? resolver, bool explicitSource = false)
    {
        foreach (string form in new[] { "simple", "packed" })
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool transformed in new[] { false, true })
        {
            var document = ReadMutation($"CompositionProbes/math-total-pages-{form}-{container}", mode, transformed, transformed,
                explicitSource ? (xml, w, m) =>
                {
                    foreach (var equation in xml.Descendants(m + "oMath").ToArray())
                        equation.ReplaceWith(new XElement(m + "oMathPara", new XElement(equation)));
                } : null);
            var owner = DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements).Single(p => p.Equations.Count != 0);
            var counters = owner.Runs.Where(run => run.FieldKind == DocxFieldKind.NumPages).ToArray();
            if (explicitSource) TestAssert.True(owner.Equations.All(equation => equation.Display), "Counter composition retains explicit source flags.");
            TestAssert.Equal(2, counters.Length);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(resolver, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(2, pages.Length);
            string text = DocxPageFootnoteNumberingTests.DecodeText(pages[0]);
            TestAssert.True(!text.Contains("777", StringComparison.Ordinal) && !text.Contains("NUMPAGES", StringComparison.Ordinal),
                "Actual total pages must replace neither current-page identity nor the stale cached result.");
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[0], usePageCoordinates: true).ToArray();
            TestAssert.Equal(2, shows.Count(show => show.Text.Trim() == "2"));
            var fonts = pages[0].Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
            string math = string.Concat(shows.Where(show => fonts.Contains(show.Font)).Select(show => show.Text)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal("ratetimeratetime", math);
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Resolvable total-page fields cannot discard supported equations.");
            TestAssert.Contains("Following prose.", DocxPageFootnoteNumberingTests.DecodeText(pages[1]));
            TestAssert.True(counters.All(run => run.Text == "{NUMPAGES}"), "The retry must preserve canonical source run identity.");
            if (transformed) TestAssert.True(owner.BookmarkAnchors.Count != 0,
                "Bookmarks remain attached to the supported equation's source owner.");
        }
    }

    public static void DocxInlineCompositionTotalPageRetryRetainsUnstableCountFallback()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("The native control uses installed Arial metrics.");
        CheckUnstableTotalPageMath(null);
    }

    internal static void CheckUnstableTotalPageMath(Lokad.OoxPdf.Fonts.IFontResolver? resolver)
    {
        var renderer = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
        DocxDocument Boundary(string instruction)
        {
            var document = ReadMutation("CompositionProbes/math-total-pages-simple-body", OoxPdfDocxMarkupMode.Final,
                false, false, (xml, w, m) =>
                {
                    foreach (var field in xml.Descendants(w + "fldSimple")) field.SetAttributeValue(w + "instr", " " + instruction + " ");
                });
            var math = document.Paragraphs.Single(paragraph => paragraph.Equations.Count != 0);
            return document with { BodyElements = [new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph("Spacer", 12d, 633d)),
                new DocxParagraphElement(math)] };
        }
        // The plain field line fits after the spacer; restoring its fractions
        // moves it to page two. A PAGE control proves that actual placement.
        TestAssert.Equal(2, renderer.RenderBlankPages(Boundary("PAGE"), null, default).Count());
        var diagnostics = new List<OoxPdfDiagnostic>();
        var source = Boundary("NUMPAGES");
        var pages = renderer.RenderBlankPages(source, diagnostics.Add, default).ToArray();
        TestAssert.Equal(1, pages.Length);
        TestAssert.True(diagnostics.Any(item => item.Id == "DOCX_UNSUPPORTED_EQUATION"),
            "A changed total cannot publish math measured with a stale count or start an unbounded fixed-point loop.");
        TestAssert.True(pages[0].Fonts.All(resource => !resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
            "Rejected retry subsets must not leak into the original fallback.");
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => renderer.RenderBlankPages(source, null, canceled.Token).ToArray());
    }

    public static void DocxInlineCompositionNoteImageDimensionsAndVerticalFrameMatchOffice()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent controls use installed Arial metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-note-images-office-geometry.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = DocxTests.ReadDocx(input, mode);
            var page = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, default).Single();
            TestAssert.Equal(1, page.Images.Count);
            var image = System.Text.RegularExpressions.Regex.Match(page.Content,
                @"(?<width>-?[\d.]+) 0 0 (?<height>-?[\d.]+) (?<x>-?[\d.]+) (?<y>-?[\d.]+) cm\s*/" + page.Images.Single().ResourceName + @" Do");
            TestAssert.True(image.Success, "The note image must reach PDF painting.");
            double Value(string name) => double.Parse(image.Groups[name].Value, System.Globalization.CultureInfo.InvariantCulture);
            var bounds = row.GetProperty("ImageBounds");
            // Each independent raster edge is sampled at half-point intervals;
            // a dimension spans two edges and therefore has one-point uncertainty.
            // Horizontal prose-advance residuals remain in the visual report.
            TestAssert.True(Math.Abs(Value("width") - bounds.GetProperty("Width").GetDouble()) <= 1.001d &&
                Math.Abs(Value("height") - bounds.GetProperty("Height").GetDouble()) <= 1.001d,
                $"{input}/{mode}: image={Value("width")}/{Value("height")}, raster={bounds.GetProperty("Width").GetDouble()}/{bounds.GetProperty("Height").GetDouble()}.");
            TestAssert.True(Math.Abs(Value("y") + Value("height") - bounds.GetProperty("Top").GetDouble()) < .75d,
                $"The occupied note line owns the picture's vertical frame: {input}/{mode}.");
        }
    }

    public static void DocxInlineCompositionNoteImagesAndMathSurviveSourceSplitting()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native controls use installed math fonts.");
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string suffix in new[] { "", "-wrapped" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/" + kind + "-image-math" + suffix, mode, true, true,
                partOverride: "word/" + kind + "s.xml");
            var note = document.RelatedStories.Single(story => story.IsNormalStoryType && story.Id == "1" &&
                story.Kind == (kind == "footnote" ? DocxRelatedStoryKind.Footnote : DocxRelatedStoryKind.Endnote));
            TestAssert.Equal(1, note.Paragraphs.Single().InlineContent!.OfType<DocxInlinePicture>().Count());
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = System.Text.RegularExpressions.Regex.Replace(string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText))
                .Normalize(NormalizationForm.FormKC), @"\s", "");
            TestAssert.Equal(4, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.Contains("Endofsupportednote.", text.Replace("ratetime", "", StringComparison.Ordinal));
            TestAssert.Equal(1, pages.Single().Images.Count);
            TestAssert.Equal(1, pages.Single().Annotations.Count(annotation => annotation.Destination is not null));
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "A picture or unchanged-format run split cannot remove supported note math.");
        }
    }

    public static void DocxInlineCompositionNormalNoteBaselinesMatchIndependentOffice()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent controls use installed Arial metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-note-content-office-geometry.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = DocxTests.ReadDocx(input, mode);
            var lines = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(document).Lines;
            string kind = input.Contains("-footnote-", StringComparison.Ordinal) ? "Footnote" : "Endnote";
            var prose = lines.First(line => line.ContainerStoryKind == kind && line.TextLength > 1)
                .Segments.First(segment => segment.TextLength > 1);
            TestAssert.True(Math.Abs(prose.BaselineY - row.GetProperty("FirstProseBaseline").GetDouble()) < .5d,
                $"Normal-note math retains the independent printed story frame: {input}/{mode}/baseline={prose.BaselineY}.");
            TestAssert.True(Math.Abs(prose.X - row.GetProperty("FirstProseX").GetDouble()) < .3d,
                "Math and adjacent note prose share the printed horizontal frame.");
        }
    }

    public static void DocxInlineCompositionNormalNotesMatchIndependentPrintedWrapping()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent controls use installed Arial metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-note-content-office-geometry.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (bool removeMath in new[] { false, true })
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            string fixture = "CompositionProbes/" + row.GetProperty("Id").GetString();
            string noteKind = input.Contains("-footnote-", StringComparison.Ordinal) ? "footnote" : "endnote";
            var document = ReadMutation(fixture, mode, false, false,
                (xml, w, m) => { if (removeMath) xml.Descendants(m + "oMath").Remove(); }, "word/" + noteKind + "s.xml");
            var lines = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(document).Lines;
            string kind = input.Contains("-footnote-", StringComparison.Ordinal) ? "Footnote" : "Endnote";
            var noteLines = lines.Where(line => line.ContainerStoryKind == kind && line.TextLength > 1).ToArray();
            TestAssert.True(row.GetProperty("ProseLineCount").GetInt32() == noteLines.Length,
                $"The normal note inherits its owning section's design width before print scaling: {input}/{mode}.");
        }
    }

    public static void DocxInlineCompositionNormalNoteSourcesRetainEquationOwners()
    {
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (string flow in new[] { "inline", "wrapped" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-note-content-" + kind + "-" + flow,
                mode, true, true, partOverride: "word/" + kind + "s.xml");
            var note = document.RelatedStories.Single(story => story.IsNormalStoryType && story.Id == "1" &&
                story.Kind == (kind == "footnote" ? DocxRelatedStoryKind.Footnote : DocxRelatedStoryKind.Endnote));
            var part = document.EquationParts.Single(part => part.PartName == note.PartName);
            TestAssert.Equal(2, part.SourceCount);
            var paragraph = note.Paragraphs.Single();
            TestAssert.True(part.BodyElements.OfType<DocxParagraphElement>().Any(element => ReferenceEquals(element.Paragraph, paragraph)),
                "Math diagnostics and preparation retain the parsed note paragraph identity.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.Equal(2, DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default)
                .Where(owner => ReferenceEquals(owner.Paragraph, paragraph)).SelectMany(owner => owner.Paragraph.Equations).Count());
        }
    }

    public static void DocxInlineCompositionNormalNoteMathSurvivesRunSplittingAndContinuation()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native controls use installed math fonts.");
        foreach (string kind in new[] { "footnote", "endnote" })
        foreach (bool continuation in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-note-content-" + kind + "-wrapped", mode, true, true,
                (xml, w, m) =>
                {
                    if (!continuation) return;
                    var paragraph = xml.Root!.Elements(w + kind).Single(note => (string?)note.Attribute(w + "id") == "1").Element(w + "p")!;
                    paragraph.Add(new XElement(w + "r", new XElement(w + "t", string.Concat(Enumerable.Repeat(" Continued public note content remains present.", 180)) + " Tail.")));
                }, "word/" + kind + "s.xml");
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            string text = System.Text.RegularExpressions.Regex.Replace(string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText))
                .Normalize(NormalizationForm.FormKC), @"\s", "");
            TestAssert.Equal(4, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            // PDF painting emits each line's prose before its mathematical glyphs.
            // Content checks do not impose a reading-order contract on that stream.
            TestAssert.Contains("Endofsupportednote.", text.Replace("ratetime", "", StringComparison.Ordinal));
            if (continuation)
            {
                TestAssert.True(pages.Length > 1, "The long note must exercise real continuation admission.");
                TestAssert.Contains("Tail.", text);
                TestAssert.Equal(180, text.Split("Continued", StringSplitOptions.None).Length - 1);
                TestAssert.Equal(180, text.Split("remainspresent.", StringSplitOptions.None).Length - 1);
            }
            TestAssert.Equal(1, pages.Sum(page => page.Annotations.Count(annotation => annotation.Destination is not null)));
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Note fractions remain placed and diagnosed from their real part.");
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_APPROXIMATED_EQUATION" && d.PartName == "/word/" + kind + "s.xml"),
                "Painted note math retains its source-part diagnostic.");
        }
    }

    public static void DocxInlineCompositionEndnotesShareOwnedAutomaticReferenceRules()
    {
        foreach (string position in new[] { "before", "after" })
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadPlacedComposition(mode, "math-endnote-" + position + "-" + container);
            var paragraph = document.Paragraphs.Concat(document.Tables.SelectMany(table => table.Rows).SelectMany(row => row.Cells)
                .SelectMany(cell => cell.Paragraphs)).First(paragraph => paragraph.Equations.Count != 0);
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            var reference = paragraph.InlineReferences.Single();
            TestAssert.Equal(DocxRelatedStoryKind.Endnote, reference.Kind);
            TestAssert.Equal("i", reference.DisplayText);
            var marker = TestAssert.NotNull(reference.DisplayRun);
            TestAssert.True(marker.IsAutomaticNoteReferenceMarker && paragraph.InlineContent!.OfType<DocxInlineText>()
                .Any(item => ReferenceEquals(item.Run, marker)), "Endnote markers retain their actual owned source run.");
            foreach (string restart in new[] { "eachPage", "eachSect" })
                TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph with
                    { InlineReferences = [reference with { NoteNumberRestartValue = restart }] }).Count);
            foreach (DocxTextRun? other in new DocxTextRun?[] { null, marker with { } })
                TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph with
                    { InlineReferences = [reference with { DisplayRun = other }] }).Count);
        }
    }

    public static void DocxInlineCompositionEndnoteRunSplittingRetainsSingleAndRepeatedMath()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the native endnote composition control.");
        foreach (int start in new[] { 1, 12 })
        foreach (bool single in new[] { false, true })
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-endnote-before-" + container, mode, true, true, (xml, w, m) =>
            {
                if (single) xml.Descendants(m + "oMath").Last().Remove();
                var reference = xml.Descendants(w + "endnoteReference").Single();
                reference.AddBeforeSelf(new XElement(w + "t", "mark"));
                reference.AddAfterSelf(new XElement(w + "t", "edge"));
                xml.Descendants(w + "endnotePr").Single().Add(new XElement(w + "numStart", new XAttribute(w + "val", start)));
            });
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(single ? 1 : 2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.Contains("mark" + (start == 1 ? "i" : "xii") + "edge", System.Text.RegularExpressions.Regex.Replace(text, @"\s", ""));
            TestAssert.Contains("Public calculation note.", text);
            TestAssert.Equal(1, pages.Sum(page => page.Annotations.Count(annotation => annotation.Destination is not null)));
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Owned endnotes cannot remove supported math.");
        }
    }

    public static void DocxInlineCompositionStaticPageFieldsMatchIndependentWordFrames()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent page-field controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-page-fields-office-offsets.json")));
        foreach (var group in references.RootElement.GetProperty("Rows").EnumerateArray()
            .GroupBy(row => (Id: row.GetProperty("Id").GetString()!, Mode: row.GetProperty("Mode").GetString()!)))
        {
            string input = Path.Combine(cases, group.Key.Id + ".docx");
            TestAssert.Equal(group.First().GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = group.Key.Mode == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var pages = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(DocxTests.ReadDocx(input, mode), null, default).ToArray();
            TestAssert.Equal(12, pages.Length);
            foreach (var row in group)
            {
                int number = row.GetProperty("Page").GetInt32();
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[number - 1]).ToArray();
                var start = shows.Single(show => show.Text.TrimStart().StartsWith("Sta", StringComparison.Ordinal));
                TestAssert.True(Math.Abs(start.Y - row.GetProperty("BaselineY").GetDouble()) < .5d,
                    $"Page-owned field/math baseline retains its independent Word frame: {group.Key}/{number}/{start.Y}.");
                var fields = shows.Where(show => Math.Abs(show.Y - start.Y) < .1d &&
                    (show.Text.Trim() == number.ToString(System.Globalization.CultureInfo.InvariantCulture) || show.Text.Trim() == "12"))
                    .OrderBy(show => show.X).ToArray();
                TestAssert.Equal(2, fields.Length);
                TestAssert.Equal(number.ToString(System.Globalization.CultureInfo.InvariantCulture), fields[0].Text.Trim());
                TestAssert.Equal("12", fields[1].Text.Trim());
                TestAssert.True(Math.Abs(fields[0].X - start.X - row.GetProperty("PageOffset").GetDouble()) < .2d,
                    "The leading field keeps its independent Word source position.");
                // Word selects Calibri for these simple-field digits; authored
                // cached runs retain Arial here. Keep this measured residual
                // explicit, separately from the unchanged whole-page gates.
                TestAssert.True(Math.Abs(fields[1].X - start.X - row.GetProperty("CountOffset").GetDouble()) < 1.25d,
                    $"Total-page field advance remains within the qualified font-owner residual: {group.Key}/{number}.");
            }
        }
    }

    public static void DocxInlineCompositionPageFieldCacheEndsWithItsOwner()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-page-fields-packed-header", mode, false, false, (xml, w, _) =>
            {
                foreach (var run in xml.Descendants(w + "r").Where(run => run.Element(w + "instrText") is not null).ToArray())
                    run.ElementsAfterSelf(w + "r").First().Element(w + "t")!.Value = " 2026 ";
            }, partOverride: "word/math-fields-header.xml");
            var paragraph = DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default).Single().Paragraph;
            TestAssert.Equal(2, paragraph.Runs.Count(run => run.Text == " 2026 "));
            TestAssert.Equal(2, paragraph.InlineContent!.OfType<DocxInlineText>().Count(item => item.Run.Text == " 2026 "));
            TestAssert.True(paragraph.FieldReferences.All(field => !field.RendersCachedResult && field.UsesPlaceholder),
                "Authored text after the end marker does not become a page-field cache.");
        }
    }

    public static void DocxInlineCompositionBodyPageCountersMatchIndependentWordFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent controls use installed Arial metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        string input = Path.Combine(cases, "math-body-page-counter-boundary.docx");
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-body-page-counter-office-geometry.json")));
        TestAssert.Equal(reference.RootElement.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        foreach (var group in reference.RootElement.GetProperty("Rows").EnumerateArray().GroupBy(row => row.GetProperty("Mode").GetString()))
        {
            bool review = group.Key == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = DocxTests.ReadDocx(input, mode);
            var pages = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, default).ToArray();
            TestAssert.Equal(10, pages.Length);
            foreach (var row in group)
            {
                int number = row.GetProperty("Page").GetInt32();
                string anchor = row.GetProperty("Anchor").GetString()!;
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[number - 1]).ToArray();
                var prose = shows.Single(show => show.Text.Trim() == anchor);
                TestAssert.True(Math.Abs(prose.Y - row.GetProperty("BaselineY").GetDouble()) < .15d,
                    "A source line admitted to the next page retains its independent Word baseline.");
                var field = shows.Single(show => show.Text.Trim() == number.ToString(System.Globalization.CultureInfo.InvariantCulture));
                TestAssert.True(Math.Abs(field.X - VisibleTextStartX(prose) - row.GetProperty("FieldOffset").GetDouble()) < .3d,
                    "Actual digits retain their advance relative to the same prose owner.");
            }
            TestAssert.True(document.Paragraphs.SelectMany(paragraph => paragraph.Runs).Where(run => run.FieldKind == DocxFieldKind.Page)
                .All(run => run.Text == "{PAGE}"), "Pagination must not replace canonical source runs.");
        }
    }

    public static void DocxInlineCompositionCellPageCountersMatchIndependentWordFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent controls use installed Arial metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        string input = Path.Combine(cases, "math-cell-page-counter-boundary.docx");
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-cell-page-counter-office-geometry.json")));
        TestAssert.Equal(reference.RootElement.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        foreach (var group in reference.RootElement.GetProperty("Rows").EnumerateArray().GroupBy(row => row.GetProperty("Mode").GetString()))
        {
            bool review = group.Key == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var pages = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(DocxTests.ReadDocx(input, mode), null, default).ToArray();
            TestAssert.Equal(10, pages.Length);
            foreach (var row in group)
            {
                int number = row.GetProperty("Page").GetInt32();
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages[number - 1]).ToArray();
                var start = shows.Single(show => show.Text.TrimStart().StartsWith("Sta", StringComparison.Ordinal));
                TestAssert.True(Math.Abs(start.Y - row.GetProperty("BaselineY").GetDouble()) < .15d, "Cell PAGE and math share the independently measured baseline.");
                var fields = shows.Where(show => show.Text.Trim() == number.ToString(System.Globalization.CultureInfo.InvariantCulture)).OrderBy(show => show.X).ToArray();
                TestAssert.Equal(2, fields.Length);
                var offsets = row.GetProperty("FieldOffsets").EnumerateArray().Select(value => value.GetDouble()).ToArray();
                // Word selects Calibri for simple-field digits while the
                // canonical source run keeps Arial, as in the static controls.
                // Retain that measured residual; whole-page gates stay failed.
                for (int index = 0; index < fields.Length; index++)
                    TestAssert.True(Math.Abs(fields[index].X - start.X - offsets[index]) < 1.25d, "Counter advances retain the measured Office font-owner residual across the digit boundary.");
            }
        }
    }

    public static void DocxInlineCompositionCellPageCountersPreserveSupportedMath() => AssertPageCounterMath("cell");

    public static void DocxInlineCompositionBodyPageCountersPreserveSupportedMath() => AssertPageCounterMath("body");

    public static void DocxInlineCompositionBodyCursorSkipsUnfitContinuationColumns()
    {
        var source = ReadMutation("CompositionProbes/math-fields-simple-body", OoxPdfDocxMarkupMode.Final,
            false, false, (xml, w, m) =>
            {
                foreach (var field in xml.Descendants(w + "fldSimple")) field.SetAttributeValue(w + "instr", " PAGE ");
            });
        var original = source.Paragraphs.Single(paragraph => paragraph.Equations.Count != 0);
        var paragraph = original with { InlineContent = original.InlineContent!.SelectMany(item =>
            new DocxInlineItem[] { item, new DocxInlineBreak(original.Runs[0], 0) }).ToArray() };
        var equations = paragraph.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
        var document = new DocxDocument(120d, 60d, 10d, 10d, 10d, 10d, DocxPageSettings.Empty,
            [], [], [], [new DocxParagraphElement(paragraph)], [], [])
        {
            FinalSectionBreak = new(DocxPageSettings.Empty, null, "2", "0", "200", [new("1600", "200"), new("200", null)])
        };
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new TrackingMeasurer(), default, mathEquations: equations);
        TestAssert.True(layout.Pages.Count > 1 && layout.Pages.Count < 10, "An unfit column must retain bounded source progress.");
        var mathLines = layout.Pages.SelectMany(page => page.Items.OfType<DocxTextLineLayout>()).Where(line => line.InlineMath.Count != 0).ToArray();
        TestAssert.Equal(2, mathLines.Sum(line => line.InlineMath.Count));
        TestAssert.True(mathLines.All(line => line.X < 90d), "Atomic math must move to the fitting wide column.");
        foreach (var page in layout.Pages.Select((value, index) => (Value: value, Number: index + 1)))
            foreach (var segment in page.Value.Items.OfType<DocxTextLineLayout>().SelectMany(line => line.Segments)
                .Where(segment => segment.StyleRun.FieldKind == DocxFieldKind.Page))
                TestAssert.Equal(page.Number.ToString(System.Globalization.CultureInfo.InvariantCulture), segment.Text);
    }

    private static void AssertPageCounterMath(string container)
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native controls use installed math fonts.");
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string form in new[] { "simple", "packed" })
        {
            var document = ReadMutation("CompositionProbes/math-fields-" + form + "-" + container, mode, true, true,
                (xml, w, m) =>
                {
                    foreach (var field in xml.Descendants(w + "fldSimple")) field.SetAttributeValue(w + "instr", " PAGE ");
                    foreach (var instruction in xml.Descendants(w + "instrText")) instruction.Value = " PAGE ";
                    foreach (var text in xml.Descendants(w + "t").Where(text => text.Value is "North" or "South")) text.Value = "1";
                });
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            foreach (string value in new[] { "Start", "middle", "End.", "rate", "time" })
                TestAssert.Contains(value, string.Concat(text.Where(character => !char.IsWhiteSpace(character))));
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "PAGE cannot hide supported " + container + " equations.");
            TestAssert.Equal(1, pages.Length);
            TestAssert.True(!text.Contains("{PAGE}", StringComparison.Ordinal), "Actual placement must resolve canonical counters.");
        }
    }

    public static void DocxInlineCompositionSourceCursorConsumesTrailingWhitespace()
    {
        var document = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "math-tabs-default18-body");
        var original = document.Paragraphs.Single(paragraph => paragraph.Equations.Count != 0);
        var whitespace = new DocxTextRun("   ", 12d, null, false, false, false, null, null);
        var paragraph = original with { Runs = original.Runs.Append(whitespace).ToArray() };
        paragraph = paragraph with { InlineContent = original.InlineContent!.Concat(new DocxInlineItem[]
            { new DocxInlineBreak(original.Runs[0], 0), new DocxInlineText(whitespace, 0, 3) }).ToArray() };
        var equations = paragraph.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new TrackingMeasurer(),
            12d, 9d, 16d, 1d, default, out var measured, defaultTabStopPoints: 18d), "Trailing source whitespace must measure.");
        var atoms = TestAssert.NotNull(measured);
        TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 300d, 300d, default, out var whole), "The control must wrap.");
        var lines = TestAssert.NotNull(whole);
        TestAssert.Equal(2, lines.Count);
        TestAssert.Equal(0, lines[^1].Fragments.Count);
        var cursor = new DocxInlineSourceCursor();
        foreach (var expected in lines)
        {
            TestAssert.True(DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, cursor, 300d, 300d, default, out var next), "The expected line must remain.");
            cursor = TestAssert.NotNull(next).EndCursor;
        }
        TestAssert.True(cursor.IsComplete(atoms), "The trailing authored line and paragraph mark must be consumed without an extra incremental line.");
    }

    public static void DocxInlineCompositionAuthoredWhitespaceFramesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent controls use installed Calibri/Arial metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        string input = Path.Combine(cases, "math-authored-inline-whitespace.docx");
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-authored-inline-whitespace-office-geometry.json")));
        TestAssert.Equal(reference.RootElement.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        foreach (var group in reference.RootElement.GetProperty("Rows").EnumerateArray().GroupBy(row => row.GetProperty("Mode").GetString()))
        {
            bool review = group.Key == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup :
                OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(DocxTests.ReadDocx(input, mode), diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(pages.Single()).ToArray();
            foreach (var container in group.GroupBy(row => row.GetProperty("Container").GetString()))
            {
                var anchors = shows.Where(show => show.Text.TrimStart() == container.Key ||
                    show.Text.TrimStart().StartsWith(container.Key + " ", StringComparison.Ordinal)).OrderByDescending(show => show.Y).ToArray();
                var expected = container.ToArray();
                TestAssert.Equal(3, anchors.Length);
                for (int index = 0; index < expected.Length; index++)
                    TestAssert.True(Math.Abs(VisibleTextStartX(anchors[index]) - expected[index].GetProperty("X").GetDouble()) < .2d &&
                        Math.Abs(anchors[index].Y - expected[index].GetProperty("BaselineY").GetDouble()) < .2d,
                        "Authored spaces and closing line struts retain independently measured " + container.Key + " geometry.");
            }
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Whitespace cannot remove supported mixed math.");
        }
    }

    public static void DocxInlineCompositionAuthoredWhitespaceRetainsSourceLineOwnership()
    {
        var document = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "math-authored-inline-whitespace");
        var paragraphs = document.Paragraphs.Where(paragraph => paragraph.Equations.Count != 0).ToArray();
        TestAssert.Equal(3, paragraphs.Length);
        foreach (var paragraph in paragraphs)
        {
            var equations = paragraph.Equations.ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new TrackingMeasurer(),
                12d, 9d, 16d, 1d, default, out var measured), "Whitespace is authored inline source.");
            var atoms = TestAssert.NotNull(measured);
            foreach (double width in new[] { 500d, 90d })
            {
                TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, width, width, default, out var whole), "Mixed whitespace must wrap.");
                var lines = TestAssert.NotNull(whole);
                var cursor = new DocxInlineSourceCursor();
                foreach (var expected in lines)
                {
                    TestAssert.True(DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, cursor, width, width, default, out var next), "The source cursor retains the same whitespace decisions.");
                    var line = TestAssert.NotNull(next);
                    TestAssert.Equal(expected.EndCursor, line.EndCursor);
                    TestAssert.True(Math.Abs(expected.Advance - line.Advance) < .001d &&
                        Math.Abs(expected.OccupiedHeight - line.OccupiedHeight) < .001d, "The line frame is independent of incremental wrapping.");
                    cursor = line.EndCursor;
                }
                TestAssert.True(cursor.IsComplete(atoms), "The paragraph mark must be consumed exactly once.");
                string Text(DocxMeasuredInlineLine line) => string.Concat(line.Fragments.SelectMany(fragment => fragment.Text).Select(span => span.Text));
                if (paragraph.Runs[0].Text.StartsWith("  ", StringComparison.Ordinal)) TestAssert.True(Text(lines[0]).StartsWith("  ", StringComparison.Ordinal), "Paragraph-start spaces remain visible source.");
                if (width == 500d && paragraph.Runs.Any(run => run.Text.Contains("Hard:", StringComparison.Ordinal)))
                    TestAssert.True(Text(lines[1]).StartsWith(" and ", StringComparison.Ordinal), "Authored-break leading spaces retain their advance.");
                if (paragraph.Runs.Any(run => run.Text.Contains("Trailing:", StringComparison.Ordinal)))
                    TestAssert.True(lines[^1].Fragments.Count == 0 && lines[^1].OccupiedHeight == 16d, "The closing paragraph mark retains a normal text strut.");
            }
        }
    }

    public static void DocxInlineCompositionSourceCursorMatchesWholeParagraphWrapping()
    {
        var document = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "math-page-fields-simple-header");
        var paragraph = DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default).Single().Paragraph;
        var equations = paragraph.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new TrackingMeasurer(),
            12d, 9d, 16d, 1d, default, out var measured, pageNumber: 9, pageCount: 123456), "Counter atoms must measure.");
        var atoms = TestAssert.NotNull(measured);
        foreach (double width in new[] { 500d, 90d, 20d })
        {
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, width, width, default, out var all), "The control must wrap.");
            var expected = TestAssert.NotNull(all);
            var cursor = new DocxInlineSourceCursor();
            for (int index = 0; index < expected.Count; index++)
            {
                TestAssert.True(DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, cursor, width, width, default, out var next),
                    "A source cursor must reproduce each qualified line.");
                var actual = TestAssert.NotNull(next);
                string Text(DocxMeasuredInlineLine line) => string.Concat(line.Fragments.SelectMany(fragment => fragment.Text).Select(span => span.Text));
                TestAssert.Equal(Text(expected[index]), Text(actual));
                TestAssert.True(Math.Abs(expected[index].Advance - actual.Advance) < .001d &&
                    Math.Abs(expected[index].BaselineInset - actual.BaselineInset) < .001d &&
                    Math.Abs(expected[index].OccupiedHeight - actual.OccupiedHeight) < .001d,
                    "Incremental wrapping retains the measured line frame.");
                TestAssert.Equal(expected[index].EndCursor, actual.EndCursor);
                TestAssert.True(actual.EndCursor.LineIndex == cursor.LineIndex + 1 &&
                    (actual.EndCursor.AtomIndex > cursor.AtomIndex || actual.EndCursor.TextOffset > cursor.TextOffset),
                    "This control must consume source content on every line.");
                cursor = actual.EndCursor;
            }
            TestAssert.True(cursor.IsComplete(atoms) &&
                !DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, cursor, width, width, default, out _),
                "Completion cannot duplicate the final source item.");
        }
    }

    public static void DocxInlineCompositionSourceCursorRebindsUnplacedPageCounters()
    {
        var document = ReadMutation("CompositionProbes/math-page-fields-simple-header", OoxPdfDocxMarkupMode.Final,
            false, false, (xml, w, m) =>
            {
                foreach (var field in xml.Descendants(w + "fldSimple")) field.SetAttributeValue(w + "instr", " PAGE ");
            }, "word/math-fields-header.xml");
        var paragraph = DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default).Single().Paragraph;
        var equations = paragraph.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new TrackingMeasurer(),
            12d, 9d, 16d, 1d, default, out var measured, pageNumber: 9), "PAGE does not depend on the total count.");
        var atoms = TestAssert.NotNull(measured);
        var cursor = new DocxInlineSourceCursor();
        var values = new List<string>();
        int steps = 0, math = 0;
        while (!cursor.IsComplete(atoms))
        {
            TestAssert.True(++steps < 20, "The source cursor must make bounded progress.");
            int number = cursor.LineIndex == 0 ? 9 : 10;
            TestAssert.True(DocxLayoutEngine.TryRebindInlineCounters(atoms, number, null, default, out var bound), "The current page owns unplaced counters.");
            TestAssert.True(DocxLayoutEngine.TryGetNextOrderedInlineLine(TestAssert.NotNull(bound), cursor, 90d, 90d, default, out var next), "The next line must wrap.");
            var line = TestAssert.NotNull(next);
            values.AddRange(line.Fragments.Where(fragment => fragment.Source is DocxInlineText { Run.FieldKind: DocxFieldKind.Page })
                .SelectMany(fragment => fragment.Text).Select(span => span.Text));
            math += line.Fragments.Count(fragment => fragment.Math is not null);
            cursor = line.EndCursor;
        }
        TestAssert.Equal("9|10", string.Join("|", values));
        TestAssert.Equal(2, math);
        TestAssert.True(paragraph.Runs.Where(run => run.FieldKind == DocxFieldKind.Page).All(run => run.Text == "{PAGE}"),
            "Page-specific values never overwrite canonical source runs.");
    }

    public static void DocxInlineCompositionSourceCursorRetainsBreakTabsAndFirstLineLabel()
    {
        foreach (string fixture in new[] { "math-tabs-default18-body", "math-composition-numbered" })
        {
            var document = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, fixture);
            var source = document.Paragraphs.First(paragraph => paragraph.Equations.Count != 0);
            var paragraph = source with { InlineContent = source.InlineContent!.Take(1)
                .Concat(new DocxInlineItem[] { new DocxInlineBreak(source.Runs[0], 0) })
                .Concat(source.InlineContent!.Skip(1)).ToArray() };
            var equations = paragraph.Equations.ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new TrackingMeasurer(),
                12d, 9d, 16d, 1d, default, out var measured, defaultTabStopPoints: 18d), "The mixed source must measure.");
            var atoms = TestAssert.NotNull(measured);
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 1d, 90d, default, out var whole), "A narrow first line must wrap.");
            var expected = TestAssert.NotNull(whole);
            var cursor = new DocxInlineSourceCursor();
            for (int index = 0; index < expected.Count; index++)
            {
                double width = cursor.LineIndex == 0 ? 1d : 90d;
                TestAssert.True(DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, cursor, width, 90d, default, out var next), "The cursor must retain authored controls.");
                var line = TestAssert.NotNull(next);
                TestAssert.Equal(expected[index].EndCursor, line.EndCursor);
                TestAssert.True(Math.Abs(expected[index].Advance - line.Advance) < .001d, "Control advances must match.");
                TestAssert.Equal(expected[index].ListLabelSegments.Count, line.ListLabelSegments.Count);
                if (index != 0) TestAssert.Equal(0, line.ListLabelSegments.Count);
                cursor = line.EndCursor;
            }
            TestAssert.True(cursor.IsComplete(atoms), "Empty first lines and authored breaks must retain progress.");
            foreach (var invalid in new[] { new DocxInlineSourceCursor(-1), new(atoms.Fragments.Count + 1), new(0, -1), new(0, LineIndex: -1) })
                TestAssert.True(!DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, invalid, 90d, 90d, default, out _), "Invalid cursors retain explicit fallback.");
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, default,
                90d, 90d, cancellation.Token, out _));
        }
    }

    public static void DocxInlineCompositionSourceCursorKeepsSplitCounterValue()
    {
        var document = ReadMutation("CompositionProbes/math-page-fields-simple-header", OoxPdfDocxMarkupMode.Final,
            false, false, (xml, w, m) =>
            {
                foreach (var field in xml.Descendants(w + "fldSimple")) field.SetAttributeValue(w + "instr", " PAGE ");
            }, "word/math-fields-header.xml");
        var paragraph = DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default).Single().Paragraph;
        var equations = paragraph.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new TrackingMeasurer(),
            12d, 9d, 16d, 1d, default, out var measured, pageNumber: 123456789), "Counter digits may exceed the source placeholder.");
        var atoms = TestAssert.NotNull(measured);
        var cursor = new DocxInlineSourceCursor();
        var values = new List<string>();
        int steps = 0;
        while (!cursor.IsComplete(atoms))
        {
            TestAssert.True(++steps < 40, "Partial counters must make bounded progress.");
            int page = values.Count == 0 ? 123456789 : 10;
            TestAssert.True(DocxLayoutEngine.TryRebindInlineCounters(atoms, page, null, default, out var rebound), "Rebinding must succeed.");
            TestAssert.True(DocxLayoutEngine.TryGetNextOrderedInlineLine(TestAssert.NotNull(rebound), cursor, 20d, 20d,
                default, out var next), "A counter can wrap beyond its placeholder length.");
            var line = TestAssert.NotNull(next);
            values.AddRange(line.Fragments.Where(fragment => fragment.Source is DocxInlineText { Run.FieldKind: DocxFieldKind.Page })
                .SelectMany(fragment => fragment.Text).Select(span => span.Text));
            cursor = line.EndCursor;
        }
        TestAssert.Equal("12345678910", string.Concat(values));
        TestAssert.True(paragraph.Runs.Where(run => run.FieldKind == DocxFieldKind.Page).All(run => run.Text == "{PAGE}"),
            "Resolved slices retain their canonical source owner.");
    }

    public static void DocxInlineCompositionMeasuredAtomsRebindOnlyCounters()
    {
        var document = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "math-page-fields-simple-header");
        var paragraph = DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default).Single().Paragraph;
        var equations = paragraph.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
        var measurer = new CountingMeasurer();
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, measurer, 12d, 9d, 16d, 1d,
            default, out var measured, pageNumber: 9, pageCount: 12), "The initial values own a reusable measured sequence.");
        var atoms = TestAssert.NotNull(measured);
        int startCalls = measurer.Calls.GetValueOrDefault("Start");
        int middleCalls = measurer.Calls.GetValueOrDefault("middle");
        TestAssert.True(startCalls > 0 && middleCalls > 0, "The fixed-content measurement control must be exercised.");
        TestAssert.True(DocxLayoutEngine.TryRebindInlineCounters(atoms, 9, 12, default, out var unchanged) &&
            ReferenceEquals(atoms, unchanged), "Unchanged inputs reuse the measured sequence.");
        TestAssert.True(DocxLayoutEngine.TryRebindInlineCounters(atoms, 10, 12, default, out var rebound),
            "A page-number change can refresh its source-owned field.");
        var changed = TestAssert.NotNull(rebound);
        TestAssert.Equal(atoms.Fragments.Count, changed.Fragments.Count);
        for (int index = 0; index < atoms.Fragments.Count; index++)
        {
            var before = atoms.Fragments[index];
            var after = changed.Fragments[index];
            if (before.Source is DocxInlineText { Run.FieldKind: DocxFieldKind.Page })
            {
                TestAssert.Equal("9", before.Text.Single().Text);
                TestAssert.Equal("10", after.Text.Single().Text);
                TestAssert.True(ReferenceEquals(before.Source, after.Source), "Counter binding keeps the source item.");
                TestAssert.True(Math.Abs(after.Advance - before.Advance - 5d) < .001d, "Binding measures the extra digit.");
            }
            else TestAssert.True(ReferenceEquals(before, after), "Fixed text, images, math and unchanged counters retain their measurements.");
        }
        TestAssert.Equal(startCalls, measurer.Calls.GetValueOrDefault("Start"));
        TestAssert.Equal(middleCalls, measurer.Calls.GetValueOrDefault("middle"));
        foreach (double width in new[] { 500d, 90d })
        {
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(changed, width, width, default, out var lines),
                "The measured sequence can wrap under different container constraints.");
            string text = System.Text.RegularExpressions.Regex.Replace(string.Concat(TestAssert.NotNull(lines)
                .SelectMany(line => line.Fragments).SelectMany(fragment => fragment.Text).Select(span => span.Text)), @"\s", "");
            TestAssert.Equal("Start10middle12End.", text);
            TestAssert.Equal(2, lines!.SelectMany(line => line.Fragments).Count(fragment => fragment.Math is not null));
        }
        TestAssert.True(!DocxLayoutEngine.TryRebindInlineCounters(atoms, null, 12, default, out _) &&
            !DocxLayoutEngine.TryRebindInlineCounters(atoms, 10, null, default, out _),
            "A failed rebind publishes no partial sequence and retains missing-input fallback.");
        TestAssert.Equal("9", atoms.Fragments.Single(fragment => fragment.Source is DocxInlineText
            { Run.FieldKind: DocxFieldKind.Page }).Text.Single().Text);
    }

    public static void DocxFieldTextCounterInputsRemainIndependent()
    {
        var template = DocxTests.CreateDocxLayoutParagraph("Source", 12d, 14d).Runs[0];
        foreach (var kind in new[] { DocxFieldKind.Page, DocxFieldKind.NumPages, DocxFieldKind.Other })
        foreach (int? number in new int?[] { null, 7 })
        foreach (int? count in new int?[] { null, 12 })
        {
            string source = "Prefix " + (DocxFieldText.Placeholder(kind) ?? "literal") + " suffix";
            var run = template with { Text = source, FieldKind = kind };
            int? expected = kind == DocxFieldKind.Page ? number : kind == DocxFieldKind.NumPages ? count : null;
            TestAssert.Equal(expected is not null, DocxFieldText.CanResolve(kind, number, count));
            TestAssert.Equal(expected is not null, DocxFieldText.TryResolve(run, source, number, count, out string resolved));
            string value = expected is { } actual ? "Prefix " + actual + " suffix" : source;
            TestAssert.Equal(value, resolved);
            TestAssert.Equal(value, DocxFieldText.Resolve(run, source, number, count));
            TestAssert.Equal(source, run.Text);
        }
    }

    public static void DocxInlineCompositionCounterAtomsMeasureOnlyTheirOwnedInputs()
    {
        foreach (string form in new[] { "simple", "packed" })
        foreach (string container in new[] { "body", "cell" })
        foreach (string instruction in new[] { "PAGE", "NUMPAGES" })
        {
            var document = ReadMutation("CompositionProbes/math-fields-" + form + "-" + container,
                OoxPdfDocxMarkupMode.Final, false, false, (xml, w, m) =>
                {
                    foreach (var field in xml.Descendants(w + "fldSimple"))
                    {
                        field.SetAttributeValue(w + "instr", " " + instruction + " ");
                        foreach (var cache in field.Descendants(w + "t")) cache.Value = "1";
                    }
                    foreach (var field in xml.Descendants(w + "instrText"))
                    {
                        field.Value = " " + instruction + " ";
                        foreach (var cache in field.Parent!.Elements(w + "t")) cache.Value = "1";
                    }
                });
            var paragraph = DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements).Single(p => p.Equations.Count == 2);
            var equations = paragraph.Equations.ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
            bool Measure(int? number, int? count, out IReadOnlyList<DocxMeasuredInlineLine>? lines) =>
                DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, equations, new TrackingMeasurer(), 12d,
                    500d, 500d, 9d, 16d, 1d, default, out lines, pageNumber: number, pageCount: count);
            TestAssert.True(!Measure(null, null, out _), "A counter cannot measure placeholder letters when its actual value is unavailable.");
            TestAssert.True(instruction == "PAGE" ? !Measure(null, 12, out _) : !Measure(12, null, out _),
                "Supplying the unrelated counter input cannot resolve an owned field.");
            double? firstWidth = null;
            string canonical = string.Concat(paragraph.Runs.Select(run => run.Text));
            foreach (int value in new[] { 9, 10, 12 })
            {
                bool measured = instruction == "PAGE" ? Measure(value, null, out var lines) : Measure(null, value, out lines);
                TestAssert.True(measured, $"{form}/{container}/{instruction}/{value}: the counter requires only its owned actual input.");
                var line = TestAssert.NotNull(lines).Single();
                string text = string.Concat(line.Fragments.SelectMany(fragment => fragment.Text).Select(span => span.Text));
                TestAssert.Equal("Start " + value + "  middle " + value + "  End.", text);
                var fields = line.Fragments.Where(fragment => fragment.Source is DocxInlineText source && source.Run.FieldKind is not null).ToArray();
                TestAssert.Equal(2, fields.Length);
                TestAssert.True(fields.All(fragment => fragment.Text.Count == 1 && ReferenceEquals(fragment.Text[0].StyleRun,
                    ((DocxInlineText)fragment.Source).Run)), "Resolved field atoms retain their canonical run owner and stay atomic.");
                TestAssert.Equal(canonical, string.Concat(paragraph.Runs.Select(run => run.Text)));
                if (firstWidth is null) firstWidth = line.Advance;
                else TestAssert.True(Math.Abs(line.Advance - firstWidth.Value - 10d) < .001d,
                    "Digit-boundary advances use the resolved values, independently of the unavailable counter.");
            }
            TestAssert.True(Measure(12, 12, out _), "Complete static-story field inputs remain supported.");
        }
    }

    public static void DocxInlineCompositionStaticPageFieldsRequireActualPageInputs()
    {
        foreach (string form in new[] { "simple", "packed" })
        foreach (string story in new[] { "header", "footer" })
        {
            var document = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "math-page-fields-" + form + "-" + story);
            var owner = DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default).Single();
            TestAssert.True(owner.HasPageInputs, "The referenced header/footer owns per-page field inputs.");
            var paragraph = owner.Paragraph;
            TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasPageInputs: true).Count);
            string ordered = string.Concat(TestAssert.NotNull(paragraph.InlineContent).OfType<DocxInlineText>()
                .Select(item => item.Run.Text.Substring(item.Start, item.Length)));
            TestAssert.Equal(string.Concat(paragraph.Runs.Select(run => run.Text)), ordered);
            var equations = paragraph.Equations.ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 10d, 16d));
            bool Measure(int? number, int? count, out IReadOnlyList<DocxMeasuredInlineLine>? lines) =>
                DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, equations, new TrackingMeasurer(), 12d,
                    500d, 500d, 9d, 16d, 1d, default, out lines, pageNumber: number, pageCount: count);
            TestAssert.True(!Measure(null, null, out _) && !Measure(1, null, out _),
                "Unavailable page inputs retain an explicit fallback rather than measuring placeholder letters.");
            foreach (int number in new[] { 1, 9, 10, 12 })
            {
                TestAssert.True(Measure(number, 12, out var lines), "Page-owned fields must measure with actual values.");
                string text = string.Concat(TestAssert.NotNull(lines).Single().Fragments.SelectMany(fragment => fragment.Text).Select(span => span.Text));
                TestAssert.Equal("Start " + number + "  middle 12  End.", text);
                TestAssert.Contains("{PAGE}", ordered);
                TestAssert.Contains("{NUMPAGES}", ordered);
            }
        }
    }

    public static void DocxInlineCompositionCachedFieldsMatchIndependentWordOffsets()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent cached-field controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-fields-office-offsets.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = DocxTests.ReadDocx(input, mode);
            var page = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).ToArray();
            var start = shows.Single(show => show.Text.TrimStart().StartsWith("Sta", StringComparison.Ordinal));
            TestAssert.True(Math.Abs(start.Y - row.GetProperty("BaselineY").GetDouble()) < .5d,
                $"Cached results preserve the independently printed paragraph frame: {input}/{mode}/{start.Y}.");
            foreach (string value in new[] { "North", "South" })
            {
                var field = shows.Single(show => show.Text.Trim() == value && Math.Abs(show.Y - start.Y) < .1d);
                double offset = field.X - start.X;
                TestAssert.True(Math.Abs(offset - row.GetProperty(value + "Offset").GetDouble()) < .2d,
                    $"Materialized cached fields retain independent Word advances: {input}/{mode}/{value}/{offset}.");
            }
        }
    }

    public static void DocxInlineCompositionCachedFieldsPublishResolvedSourceItems()
    {
        foreach (string form in new[] { "simple", "packed" })
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadPlacedComposition(mode, "math-fields-" + form + "-" + container);
            var paragraph = document.Paragraphs.Concat(document.Tables.SelectMany(table => table.Rows)
                .SelectMany(row => row.Cells).SelectMany(cell => cell.Paragraphs))
                .Concat(document.EquationParts.SelectMany(part => part.BodyElements.OfType<DocxParagraphElement>()).Select(item => item.Paragraph))
                .First(paragraph => paragraph.Equations.Count != 0);
            TestAssert.Equal(2, paragraph.FieldReferences.Count);
            TestAssert.True(paragraph.FieldReferences.All(field => field.HasCachedResult && field.RendersCachedResult && !field.UsesPlaceholder),
                "Both field forms expose their already materialized cached results.");
            string prose = string.Concat(paragraph.Runs.Select(run => run.Text));
            string ordered = string.Concat(TestAssert.NotNull(paragraph.InlineContent).OfType<DocxInlineText>().Select(item => item.Run.Text.Substring(item.Start, item.Length)));
            TestAssert.True(prose == ordered, "Cached field results before and after a formula must enter the ordered source view.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
        }
    }

    public static void DocxInlineCompositionCachedFieldsRetainRunSplittingAndFallbacks()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the native cached-field composition controls.");
        foreach (string form in new[] { "simple", "packed" })
        foreach (string container in new[] { "body", "cell" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = "CompositionProbes/math-fields-" + form + "-" + container;
            var source = ReadMutation(fixture, mode, false, false);
            var split = ReadMutation(fixture, mode, true, true, (xml, w, _) =>
            {
                foreach (var run in xml.Descendants(w + "r").Where(run => run.Element(w + "fldChar") is not null).ToArray())
                {
                    var properties = run.Element(w + "rPr");
                    run.ReplaceWith(run.Elements().Where(child => child.Name != w + "rPr")
                        .Select(child => new XElement(w + "r", properties is null ? null : new XElement(properties), new XElement(child))).ToArray());
                }
            });
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var before = renderer.RenderBlankPages(source, diagnostics.Add, default).ToArray();
            var after = renderer.RenderBlankPages(split, diagnostics.Add, default).ToArray();
            string Text(IEnumerable<Lokad.OoxPdf.Pdf.PdfPage> pages) => string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(System.Text.RegularExpressions.Regex.Replace(Text(before), @"\s", ""),
                System.Text.RegularExpressions.Regex.Replace(Text(after), @"\s", ""));
            TestAssert.Equal(2, Text(before).Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Cached field metadata cannot suppress supported equations.");
            var paragraph = source.Paragraphs.Concat(source.Tables.SelectMany(table => table.Rows).SelectMany(row => row.Cells)
                .SelectMany(cell => cell.Paragraphs)).First(paragraph => paragraph.Equations.Count != 0);
            foreach (var field in paragraph.FieldReferences)
            {
                TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph with
                    { FieldReferences = [field with { UsesPlaceholder = true }] }).Count);
                TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph with
                    { FieldReferences = [field with { RendersCachedResult = false }] }).Count);
            }
            var dynamic = ReadMutation(fixture, mode, false, false, (xml, w, _) =>
            {
                foreach (var field in xml.Descendants(w + "fldSimple")) field.SetAttributeValue(w + "instr", " PAGE ");
                foreach (var instruction in xml.Descendants(w + "instrText")) instruction.Value = " NUMPAGES ";
            });
            var dynamicParagraph = dynamic.Paragraphs.Concat(dynamic.Tables.SelectMany(table => table.Rows).SelectMany(row => row.Cells)
                .SelectMany(cell => cell.Paragraphs)).First(paragraph => paragraph.Equations.Count != 0);
            TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(dynamicParagraph).Count);
        }
    }

    public static void DocxInlineCompositionAutomaticFootnotesRetainOwnedMarkerRuns()
    {
        foreach (string position in new[] { "before", "after" })
        foreach (string suffix in new[] { "", "-plain" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var paragraph = ReadPlacedComposition(mode, "math-footnote-" + position + "-body" + suffix).Paragraphs.Single(p => p.Equations.Count != 0);
            var reference = paragraph.InlineReferences.Single();
            TestAssert.True(paragraph.InlineContent is not null && reference.DisplayRun is { } run &&
                paragraph.Runs.Any(item => ReferenceEquals(item, run)) && run.Text == reference.DisplayText,
                "Automatic footnote glyphs must have an explicit resolved source owner in the ordered content.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            foreach (string restart in new[] { "eachPage", "eachSect", "invalid" })
                TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph with
                    { InlineReferences = [reference with { NoteNumberRestartValue = restart }] }).Count);
            TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph with
                { InlineReferences = [reference with { DisplayRun = null }] }).Count);
        }
    }

    public static void DocxInlineCompositionMathCannotEnterPlainTextCorrections()
    {
        var document = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "math-footnote-before-body-plain");
        var paragraph = document.Paragraphs.Single(p => p.Equations.Count != 0);
        TestAssert.True(!DocxPlainBodyFrame.HasPlainTextContent(paragraph),
            "Plain-text correction ownership must exclude visible measured objects such as equations.");
        TestAssert.True(!DocxPlainBodyFrame.IsEligible(document), "A paragraph's math cannot disappear from whole-document eligibility.");
    }

    public static void DocxInlineCompositionFootnoteOwnershipSurvivesSharedRunsAndNumberStarts()
    {
        foreach (int start in new[] { 1, 9, 12, 103 })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-footnote-before-body-plain", mode, true, true, (xml, w, _) =>
            {
                var marker = xml.Descendants(w + "footnoteReference").Single();
                marker.AddBeforeSelf(new XElement(w + "t", "mark"));
                marker.AddAfterSelf(new XElement(w + "t", "edge"));
                xml.Descendants(w + "footnotePr").Single().Add(new XElement(w + "numStart", new XAttribute(w + "val", start)));
            });
            var paragraph = document.Paragraphs.Single(p => p.Equations.Count != 0);
            var reference = paragraph.InlineReferences.Single();
            TestAssert.Equal(start.ToString(System.Globalization.CultureInfo.InvariantCulture), reference.DisplayText);
            TestAssert.True(reference.DisplayRun is { } owner && paragraph.Runs.Count(run => run.SourceRunIndex == reference.SourceRunIndex) == 3 &&
                paragraph.InlineContent!.OfType<DocxInlineText>().Any(item => ReferenceEquals(item.Run, owner)),
                "The marker's identity must survive adjoining text in the same authored run and wider automatic labels.");
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
        }
    }

    public static void DocxInlineCompositionFootnoteLinksExcludeAdjacentIdenticalAuthoredDigits()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the rendered marker-owner control.");
        foreach (bool includeMath in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-footnote-before-body-plain", mode, false, false, (xml, w, m) =>
            {
                xml.Descendants(w + "footnoteReference").Single().AddAfterSelf(new XElement(w + "t", "1"));
                if (!includeMath) xml.Descendants(m + "oMath").Remove();
            });
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).ToArray();
            TestAssert.Equal(1, pages.Single().Annotations.Count(annotation => annotation.Destination is not null));
            TestAssert.Equal(includeMath ? 2 : 0, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC)
                .Split("ratetime", StringSplitOptions.None).Length - 1);
        }
    }

    public static void DocxInlineCompositionAutomaticFootnotesPreserveContentAndDestinations()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the footnote composition controls.");
        foreach (string position in new[] { "before", "after" })
        foreach (string container in new[] { "body", "cell" })
        foreach (string suffix in new[] { "", "-plain" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadPlacedComposition(mode, "math-footnote-" + position + "-" + container + suffix);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.Contains("Public calculation note.", text);
            var link = pages.Single().Annotations.Single(annotation => annotation.Destination is not null);
            TestAssert.True(link.Width > 0d && link.Height > 0d && link.Destination!.Value.PageIndex == 0,
                "The measured marker must keep its link to the visible note.");
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "A footnote cannot discard the supported formulas.");
        }
    }

    public static void DocxInlineCompositionFootnoteMarkersMatchIndependentPrintedOffsets() =>
        CheckAutomaticNoteMarkerGeometry("math-footnotes-office-markers.json");

    public static void DocxInlineCompositionEndnoteMarkersMatchIndependentPrintedOffsets() =>
        CheckAutomaticNoteMarkerGeometry("math-endnotes-office-markers.json");

    public static void DocxInlineCompositionEndnoteStartMatchesIndependentParagraphBoundary()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent controls use installed Arial metrics.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-endnotes-office-markers.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = DocxTests.ReadDocx(input, mode);
            var lines = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout).InspectTextEmission(document).Lines;
            var body = lines.Last(line => line.ContainerStoryKind == "Body");
            var note = lines.Single(line => line.ContainerStoryKind == "Endnote" && line.TextLength > 1);
            double gap = body.Segments[0].BaselineY - note.Segments.MaxBy(segment => segment.TextLength)!.BaselineY;
            TestAssert.True(Math.Abs(gap - row.GetProperty("StoryBaselineGap").GetDouble()) < .25d,
                $"The trailing paragraph owns the printed endnote start, including earlier tables and print scaling: {input}/{mode}/gap={gap}.");
        }
    }

    private static void CheckAutomaticNoteMarkerGeometry(string referenceFile)
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent note marker controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, referenceFile)));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = DocxTests.ReadDocx(input, mode);
            var renderer = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var line = renderer.InspectTextEmission(document).Lines.Single(line =>
                line.Segments.Any(segment => segment.TextLength == 1 &&
                    segment.FontSize < line.Segments.Max(s => s.FontSize) - .01d));
            var segments = line.Segments.Where(segment => !segment.IsTerminalLineSpace).ToArray();
            var marker = segments.Single(segment => segment.FontSize == segments.Min(s => s.FontSize));
            var prose = segments.First(segment => segment.FontSize == segments.Max(s => s.FontSize));
            double raise = marker.BaselineY - prose.BaselineY;
            TestAssert.True(Math.Abs(raise - row.GetProperty("MarkerRaise").GetDouble()) < .1d,
                $"Source vertical offsets must scale with the printed container: {input}/{mode}/raise={raise}.");
            TestAssert.True(Math.Abs(marker.PdfFontSize / prose.PdfFontSize - row.GetProperty("MarkerFontRatio").GetDouble()) < .002d,
                "The reference marker must keep its independently measured superscript font ratio.");
        }
    }

    public static void DocxInlineCompositionCommentsRetainOrderedRangeBoundaries()
    {
        foreach (string anchor in new[] { "prose", "between", "formula", "mixed" })
        foreach (bool split in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-comment-" + anchor + "-body", mode, false, split);
            var paragraph = document.Paragraphs.Single(p => p.Equations.Count != 0);
            var items = TestAssert.NotNull(paragraph.InlineContent);
            var range = paragraph.CommentRanges.Single();
            TestAssert.True(range.StartInlineItemIndex is >= 0 && range.EndInlineItemIndex > range.StartInlineItemIndex,
                "Comment boundaries must retain ordered source identity, including unchanged-format run splits.");
            int start = range.StartInlineItemIndex!.Value, end = range.EndInlineItemIndex!.Value;
            TestAssert.Equal(anchor is "prose" or "between" ? 0 : anchor == "formula" ? 1 : 2,
                items.Skip(start).Take(end - start).OfType<DocxInlineEquation>().Count());
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            var note = new DocxInlineReference(DocxRelatedStoryKind.Footnote, "1", null, "1", 0, 0, 0);
            TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph with { InlineReferences = [note] }).Count);
        }
    }

    public static void DocxInlineCompositionCommentsPreserveMathematicsAndCaptions()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the native comment composition controls.");
        foreach (string anchor in new[] { "prose", "between", "formula", "mixed" })
        foreach (string container in new[] { "body", "cell" })
        foreach (bool split in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-comment-" + anchor + "-" + container, mode, true, split);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(1, pages.Length);
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            foreach (string word in new[] { "Start", "middle", "End.", "Following prose." }) TestAssert.Contains(word, text);
            if (mode == OoxPdfDocxMarkupMode.AllMarkup) TestAssert.Contains("Check composed calculation.", text);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Comments and bookmarks cannot remove otherwise supported math.");
        }
    }

    public static void DocxInlineCompositionCommentHighlightsMatchIndependentSelectionBounds()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent comment geometry controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-comments-office-bounds.json")));
        var heights = new Dictionary<string, double>();
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var document = DocxTests.ReadDocx(input, OoxPdfDocxMarkupMode.AllMarkup);
            var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderBlankPages(document, null, default).Single();
            var fills = System.Text.RegularExpressions.Regex.Matches(page.Content,
                @"0\.973 0\.863 0\.867 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f")
                .Cast<System.Text.RegularExpressions.Match>().Select(match => (
                    X: double.Parse(match.Groups["x"].Value, System.Globalization.CultureInfo.InvariantCulture),
                    Width: double.Parse(match.Groups["w"].Value, System.Globalization.CultureInfo.InvariantCulture),
                    Height: double.Parse(match.Groups["h"].Value, System.Globalization.CultureInfo.InvariantCulture)))
                .Where(fill => fill.X < 400d).ToArray();
            TestAssert.Equal(1, fills.Length);
            TestAssert.True(Math.Abs(fills[0].Height - row.GetProperty("Height").GetDouble()) < .25d,
                $"Comment selection must retain the independently measured Word line height: {input}/{fills[0].Height}.");
            string container = Path.GetFileNameWithoutExtension(input).Split('-')[^1];
            if (heights.TryGetValue(container, out double height))
                TestAssert.True(Math.Abs(height - fills[0].Height) < .001d,
                    "Independent Word controls retain one line selection height when only the annotated source range changes.");
            else heights.Add(container, fills[0].Height);
            // PDF text-show boundaries may join a word with its authored space.
            // The independent reference is anchored to visible source text.
            var startShow = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)
                .Single(show => show.Text.TrimStart().StartsWith("Start", StringComparison.Ordinal));
            double start = VisibleTextStartX(startShow);
            double baseline = startShow.Y;
            double connectorOffset = renderer.InspectMarkupBalloons(document).Single().AnchorY - baseline;
            TestAssert.True(Math.Abs(connectorOffset - row.GetProperty("ConnectorBaselineOffset").GetDouble()) < .3d,
                $"Comment connector must consume the independent selected line frame: {input}/{connectorOffset}.");
            double titleOffset = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)
                .Single(show => show.Text.StartsWith("Commented ", StringComparison.Ordinal)).Y - baseline;
            // The accepted balloon baseline remains approximate; connector depth
            // must not make its independent Office residual worse.
            TestAssert.True(Math.Abs(titleOffset - row.GetProperty("BalloonTitleBaselineOffset").GetDouble()) < 3d,
                $"Selection depth cannot displace the balloon title: {input}/{titleOffset}.");
            foreach (var (value, property) in new[] { (fills[0].X - start, "StartOffset"), (fills[0].X + fills[0].Width - start, "EndOffset") })
                TestAssert.True(Math.Abs(value - row.GetProperty(property).GetDouble()) < .4d,
                    $"Comment fill must follow the independent selected source range: {input}/{property}={value}.");
        }
    }

    public static void DocxInlineCompositionCommentedSoleFormulaRetainsDisplayOwnership()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-comment-formula-body", mode, false, false, (xml, w, m) =>
            {
                var equation = xml.Descendants(m + "oMath").First();
                equation.Parent!.ReplaceNodes(new XElement(w + "commentRangeStart", new XAttribute(w + "id", "0")),
                    new XElement(equation), new XElement(w + "commentRangeEnd", new XAttribute(w + "id", "0")),
                    new XElement(w + "r", new XElement(w + "commentReference", new XAttribute(w + "id", "0"))));
            });
            var paragraph = document.Paragraphs.Single(p => p.Equations.Count != 0);
            TestAssert.Equal(1, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.True(DocxMathParagraphPolicy.GetDisplaySourceEquations(paragraph, default).Contains(paragraph.Equations.Single()),
                "Annotation metadata cannot turn a sole visible formula into inline mathematical styling.");
        }
    }

    public static void DocxInlineCompositionLeftTabsMatchIndependentFieldOffsets()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent tab controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-tabs-office-offsets.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = DocxTests.ReadDocx(input, mode);
            var renderer = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("ratetime", StringSplitOptions.None).Length - 1);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Tabs cannot discard the supported formulas.");
            var shows = pages.SelectMany(DocxCommentsTests.ReadEmbeddedGlyphTextShows).ToArray();
            double start = VisibleTextStartX(shows.Single(show => show.Text.TrimStart().StartsWith("Start", StringComparison.Ordinal)));
            foreach (var (textPrefix, property) in new[] { ("Middle", "MiddleOffset"), ("End.", "EndOffset") })
            {
                double actual = VisibleTextStartX(shows.Single(show => show.Text.TrimStart().StartsWith(textPrefix, StringComparison.Ordinal))) - start;
                TestAssert.True(Math.Abs(actual - row.GetProperty(property).GetDouble()) < .2d,
                    $"Measured tabs must retain independent Word field offsets: {input}/{mode}/{property}={actual}.");
            }
        }
    }

    public static void DocxInlineCompositionLeftTabMeasurementRetainsWrapAndFallbackContracts()
    {
        var paragraph = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "math-tabs-default18-body").Paragraphs.Single(p => p.Equations.Count != 0);
        var equations = paragraph.Equations.ToDictionary(e => e, e => new DocxPreparedMathParagraph(e, new(18d, 10d, 4d, []), 11d, 16d));
        var sole = paragraph with { Equations = [paragraph.Equations[0]],
            InlineContent = [new DocxInlineEquation(paragraph.Equations[0]), paragraph.InlineContent!.OfType<DocxInlineTab>().First()] };
        TestAssert.True(DocxMathParagraphPolicy.GetStandaloneRow(sole) is null,
            "A visible tab is another inline owner; a lone formula must not promote to a centered display row.");
        bool Measure(DocxParagraph source, double width, double scale, out IReadOnlyList<DocxMeasuredInlineLine>? lines) =>
            DocxLayoutEngine.TryMeasureOrderedInlineContent(source, equations, new TrackingMeasurer(), 12d,
                width, width, 9d, 16d, scale, default, out lines, defaultTabStopPoints: 18d);
        TestAssert.True(Measure(paragraph, 300d, 1d, out var full), "Default tabs must measure.");
        TestAssert.Equal(2, TestAssert.NotNull(full).Single().Fragments.Count(fragment => fragment.Source is DocxInlineTab));
        foreach (double scale in new[] { 1d, .5d })
        {
            var explicitStops = paragraph with { TabStops = [new(108d, null, "left", null), new(216d, null, "left", null)] };
            TestAssert.True(Measure(explicitStops, 300d, scale, out var positioned), "Explicit left stops must measure in scaled coordinates.");
            double advance = 0d;
            int stopIndex = 0;
            foreach (var fragment in TestAssert.NotNull(positioned).Single().Fragments)
            {
                advance += fragment.LeadingAdvance + fragment.Advance;
                if (fragment.Source is DocxInlineTab) TestAssert.Equal(++stopIndex * 108d * scale, advance);
            }
        }
        foreach (double width in new[] { 110d, 80d })
        {
            TestAssert.True(Measure(paragraph, width, 1d, out var wrapped), "Default tabs must retain wrapping.");
            TestAssert.True(TestAssert.NotNull(wrapped).Count > 1 && wrapped.All(line => line.Advance <= width), "Each wrapped row must fit.");
            TestAssert.Equal(2, wrapped.SelectMany(line => line.Fragments).Count(fragment => fragment.Math is not null));
        }
        foreach (var stop in new[] { new DocxTabStop(72d, null, "right", null), new(72d, null, "center", null),
            new(72d, null, "decimal", null), new(72d, null, "bar", null), new(72d, null, "left", "dot") })
            TestAssert.True(!Measure(paragraph with { TabStops = [stop] }, 300d, 1d, out _), "Overlapping aligned fields, decimal/bar stops and leaders retain explicit fallback.");
        TestAssert.True(!Measure(paragraph with { TabStops = [new(400d, null, "left", null)] }, 100d, 1d, out _), "An out-of-frame stop retains fallback.");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph,
            equations, new TrackingMeasurer(), 12d, 300d, 300d, 9d, 16d, 1d, cancellation.Token, out _));
    }

    public static void DocxInlineCompositionAlignedTabsMeasureCompleteFieldsAndKeepFallbacks()
    {
        var source = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "math-tabs-default18-body")
            .Paragraphs.Single(paragraph => paragraph.Equations.Count != 0);
        var equations = source.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(18d, 10d, 4d, []), 11d, 16d));
        foreach (string alignment in new[] { "right", "center" })
        foreach (double scale in new[] { 1d, .5d })
        {
            var paragraph = source with { TabStops = [new(216d, null, alignment, null), new(432d, null, alignment, null)] };
            bool Measure(double width, out IReadOnlyList<DocxMeasuredInlineLine>? lines) =>
                DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, equations, new TrackingMeasurer(),
                    12d, width, width, 9d, 16d, scale, default, out lines, defaultTabStopPoints: 18d);
            TestAssert.True(Measure(600d, out var lines), "Complete right/center fields must measure.");
            var line = TestAssert.NotNull(lines).Single();
            TestAssert.Equal(2, line.Fragments.Count(fragment => fragment.Math is not null));
            double offset = 0d;
            int tab = 0;
            foreach (var fragment in line.Fragments)
            {
                offset += fragment.LeadingAdvance + fragment.Advance;
                if (fragment.Source is not DocxInlineTab) continue;
                double width = tab == 0 ? 35d + 18d * scale : 20d;
                double expected = (++tab * 216d * scale) - width / (alignment == "right" ? 1d : 2d);
                TestAssert.True(Math.Abs(offset - expected) < .000001d,
                    "The tab positions the complete field, including following mathematics.");
            }
            TestAssert.True(!Measure(100d, out lines) && lines is null,
                "An aligned field that exceeds the frame keeps the whole-owner fallback.");
            var overlap = paragraph with { TabStops = [new((35d + 18d * scale) / scale, null, alignment, null)] };
            TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineContent(overlap, equations,
                new TrackingMeasurer(), 12d, 600d, 600d, 9d, 16d, scale, default, out _),
                "An aligned field cannot overlap already placed inline content.");
        }
    }

    public static void DocxInlineCompositionAlignedTabsMatchIndependentWordAndSourceTransforms()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed math fonts are required for aligned mixed fields.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-tabs-aligned-office-geometry.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (bool annotated in new[] { false, true })
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(cases, id + ".docx")))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = ReadMutation("CompositionProbes/" + id, mode, annotated, annotated, (xml, w, m) =>
            {
                if (!annotated) return;
                foreach (var equation in xml.Descendants(m + "oMath").ToArray())
                    equation.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), new XElement(equation)));
            }, partOverride: id.EndsWith("-footer", StringComparison.Ordinal) ? "word/math-tabs-footer.xml" : null);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var start = shows.Single(show => show.Text.TrimStart().StartsWith("Sta", StringComparison.Ordinal));
            double startX = VisibleTextStartX(start);
            TestAssert.True(Math.Abs(start.Y - row.GetProperty("StartY").GetDouble()) < .5d,
                "Aligned fields retain the independent body/cell/static baseline within the existing frame residual.");
            foreach (var name in new[] { "Middle", "End." })
            {
                string prefix = name == "Middle" ? "Mid" : "En";
                var text = shows.Single(show => show.Text.TrimStart().StartsWith(prefix, StringComparison.Ordinal));
                double expected = row.GetProperty(name.TrimEnd('.') + "OffsetX").GetDouble();
                TestAssert.True(Math.Abs(VisibleTextStartX(text) - startX - expected) < .2d,
                    "Text consumes the complete independently aligned field: " + id + "/" + name + ".");
            }
            var mathFonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                .Select(resource => resource.Font).ToHashSet();
            var math = shows.Where(show => mathFonts.Contains(show.Font)).ToArray();
            TestAssert.Equal("ratetimeratetime", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(16, math.Length);
            var anchors = row.GetProperty("MathAnchors").EnumerateArray().ToArray();
            for (int anchor = 0; anchor < anchors.Length; anchor++)
            {
                var glyph = math[anchor * 4];
                TestAssert.True(Math.Abs(glyph.X - startX - anchors[anchor].GetProperty("OffsetX").GetDouble()) < .25d &&
                    Math.Abs(glyph.Y - start.Y - anchors[anchor].GetProperty("OffsetY").GetDouble()) < .3d,
                    "Mathematical numerator/denominator anchors follow the independent tab field: " + id + ".");
            }
            if (annotated) AssertTwoFormulaLinksCoverGlyphs(page, math, id + "/" + mode);
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Supported aligned fields cannot discard their formulas.");
        }
    }

    public static void DocxInlineCompositionLeftTabMathLinksCoverCellAndStaticGlyphs()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed math fonts are required for mathematical link frames.");
        foreach (string container in new[] { "cell", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = "CompositionProbes/math-tabs-default18-" + container;
            var document = ReadMutation(fixture, mode, true, false, (xml, w, m) =>
            {
                foreach (var equation in xml.Descendants(m + "oMath").ToArray())
                    equation.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), new XElement(equation)));
            }, partOverride: container == "footer" ? "word/math-tabs-footer.xml" : null);
            var page = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, null, default).Single();
            var mathFonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                .Select(resource => resource.Font).ToHashSet();
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Where(show => mathFonts.Contains(show.Font)).ToArray();
            TestAssert.Equal("ratetimeratetime", string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            AssertTwoFormulaLinksCoverGlyphs(page, math, fixture + "/" + mode);
        }
    }

    private static void AssertTwoFormulaLinksCoverGlyphs(PdfPage page,
        IEnumerable<(PdfEmbeddedFont Font, string Text, double Size, double X, double Y)> glyphs, string context)
    {
        var links = page.Annotations.Where(annotation => annotation.IsDestination).ToArray();
        TestAssert.Equal(2, links.Length);
        foreach (var glyph in glyphs)
            TestAssert.True(links.Any(link => link.X <= glyph.X + .01d && link.X + link.Width >= glyph.X - .01d &&
                link.Y <= glyph.Y + .01d && link.Y + link.Height >= glyph.Y - .01d),
                "Source-backed links cover painted mathematical glyphs: " + context + " at " + glyph.X + "," + glyph.Y + "; " +
                string.Join(";", links.Select(link => $"{link.X},{link.Y},{link.Width},{link.Height}")));
    }

    public static void DocxInlineCompositionContainedInkRetainsPreparedOccupiedFrames()
    {
        var paragraph = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final,
            "math-printed-frame-script-anchor12-multiple").Paragraphs.Single(p => p.Equations.Count != 0);
        TestAssert.Equal(2, paragraph.Equations.Count);
        var equations = paragraph.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(10d, 6d, 1d, []), 14d, 18d));
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, equations,
            new TrackingMeasurer(), 12d, 500d, 500d, 12d, 13d, .75d, default, out var lines,
            physicalTextBaselineInset: 9d), "Contained ink must retain a complete measured line.");
        var line = TestAssert.NotNull(lines).Single();
        TestAssert.Equal(10.5d, line.BaselineInset);
        TestAssert.Equal(14.5d, line.OccupiedHeight);
        TestAssert.Equal(9d, line.TextBaselineInset);
        TestAssert.True(line.Fragments.Where(fragment => fragment.Math is not null).All(fragment =>
            fragment.BaselineInset == 10.5d && fragment.OccupiedHeight == 13.5d),
            "Prepared leading contributes once to the shared frame and remains owned by each formula.");
    }

    public static void DocxInlineCompositionScriptFramesMatchIndependentPrintedGaps()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent printed script controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-printed-frame-script-office-geometry.json")));
        foreach (var group in references.RootElement.GetProperty("Rows").EnumerateArray()
            .GroupBy(row => (Id: row.GetProperty("Id").GetString()!, Mode: row.GetProperty("Mode").GetString())))
        {
            string input = Path.Combine(cases, group.Key.Id + ".docx");
            TestAssert.Equal(group.First().GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            bool review = group.Key.Mode == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var document = DocxTests.ReadDocx(input, mode);
            var renderer = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(renderer.RenderBlankPages(document, null, default).Single());
            double leading = shows.Single(show => show.Text.StartsWith("Leading paragraph.", StringComparison.Ordinal)).Y;
            foreach (var row in group)
            {
                string prefix = row.GetProperty("Prefix").GetString()!;
                double gap = leading - shows.Single(show => show.Text.StartsWith(prefix, StringComparison.Ordinal)).Y;
                TestAssert.True(Math.Abs(gap - row.GetProperty("LeadingToBaseline").GetDouble()) < .15d,
                    $"Prepared script frames retain independent Word baseline gaps: {group.Key}/{prefix}/{gap}.");
            }
        }
    }

    public static void DocxInlineCompositionPrintedMathMatchesIndependentBaselineGaps()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent printed-frame controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-printed-frame-office-baselines.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string source = Path.Combine(cases, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(source))));
            bool review = row.GetProperty("Mode").GetString() == "all-word-compatible";
            var mode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final;
            var renderer = new DocxRenderer(null, mode, review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup
                : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
            var items = renderer.InspectLayout(DocxTests.ReadDocx(source, mode)).Pages.Single().Items;
            double Baseline(int block) => items.Single(item => item.Kind == "TextLine" && item.SourceBlockIndex == block).Y;
            foreach (var (block, property) in new[] { (1, "LeadingToMath"), (2, "LeadingToFollowing") })
            {
                double delta = Baseline(0) - Baseline(block) - row.GetProperty(property).GetDouble();
                TestAssert.True(Math.Abs(delta) < .4d,
                    $"Printed placement and advance must retain independent Word gaps: {id}/{mode}/{property} delta={delta}.");
            }
        }
    }

    public static void DocxInlineCompositionBrokenLiteralsSurviveCellAndStaticPlacement()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the broken-literal placement control.");
        foreach (string story in new[] { "cell", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-literal-inline-break-default-" + story, mode, true, true,
                partOverride: story == "cell" ? null : "word/" + story + "-math-break.xml");
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            foreach (string glyph in new[] { "α", "+", "β", "=", "γ" })
                TestAssert.Equal(1, text.Split(glyph, StringSplitOptions.None).Length - 1);
            TestAssert.Contains("Following prose.", text);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Placement and source annotations cannot remove the broken literal.");
        }
    }

    public static void DocxInlineCompositionGroupedLiteralsMatchIndependentOperatorRows()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent grouped literal controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-grouped-literal-office-geometry.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string input = Path.Combine(cases, row.GetProperty("Id").GetString() + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = renderer.RenderBlankPages(DocxTests.ReadDocx(input, mode), diagnostics.Add, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var mathFonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                .Select(resource => resource.Font).ToHashSet();
            var lines = shows.Where(show => mathFonts.Contains(show.Font)).GroupBy(show => show.Y)
                .OrderByDescending(group => group.Key).Select(group => group.ToArray()).ToArray();
            var expected = row.GetProperty("MathRows").EnumerateArray().ToArray();
            TestAssert.True(expected.Length == lines.Length, input + "/" + mode + ": " +
                string.Join("; ", lines.Select(line => line[0].Y + ":" + string.Concat(line.Select(show => show.Text)))));
            for (int line = 0; line < expected.Length; line++)
            {
                TestAssert.Equal(expected[line].GetProperty("Text").GetString(),
                    string.Concat(lines[line].Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.True(Math.Abs(lines[line][0].Y - expected[line].GetProperty("BaselineY").GetDouble()) < .3d,
                    "Grouped rows retain independent Word baselines.");
                var offsets = expected[line].GetProperty("GlyphOffsets").EnumerateArray().ToArray();
                TestAssert.Equal(offsets.Length, lines[line].Length);
                for (int glyph = 0; glyph < offsets.Length; glyph++)
                    TestAssert.True(Math.Abs(lines[line][glyph].X - lines[line][0].X - offsets[glyph].GetDouble()) < .15d,
                        "Operator breaks retain independently measured local mathematical glyph geometry.");
            }
            TestAssert.True(Math.Abs(shows.Single(show => show.Text.StartsWith("Following prose.", StringComparison.Ordinal)).Y -
                row.GetProperty("FollowingBaselineY").GetDouble()) < .3d, "Following prose must consume the complete grouped frame.");
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Balanced literal groups cannot lose their formula.");
        }
    }

    public static void DocxInlineCompositionGroupedLiteralsRetainBookmarksAndRunSplitting()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for grouped source transformations.");
        foreach (string fixture in new[] { "atomic-math-parenthesized", "math-grouped-nested" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            (string Text, double X, double Y)[]? baseline = null;
            foreach (var (bookmark, split, link) in new[] { (false, false, false), (true, false, false), (true, true, false), (true, false, true) })
            {
                var document = ReadMutation("CompositionProbes/" + fixture, mode, bookmark, split, (xml, w, m) =>
                {
                    if (!link) return;
                    var equation = xml.Descendants(m + "oMath").First();
                    equation.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), new XElement(equation)));
                });
                var page = renderer.RenderBlankPages(document, null, default).Single();
                var mathFonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                    .Select(resource => resource.Font).ToHashSet();
                var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => mathFonts.Contains(show.Font))
                    .Select(show => (show.Text, show.X, show.Y)).ToArray();
                TestAssert.True(math.Length != 0, "The source formula must remain visible.");
                if (link)
                {
                    var annotations = page.Annotations.Where(annotation => annotation.IsDestination).ToArray();
                    var rows = math.GroupBy(show => show.Y).ToArray();
                    TestAssert.Equal(rows.Length, annotations.Length);
                    foreach (var row in rows)
                        TestAssert.True(annotations.Any(annotation => annotation.Y <= row.Key && annotation.Y + annotation.Height >= row.Key &&
                            annotation.X <= row.Min(show => show.X) + .01d && annotation.X + annotation.Width >= row.Max(show => show.X) - .01d),
                            "A formula hyperlink must cover every painted mathematical row in page coordinates.");
                }
                if (baseline is null) baseline = math;
                else
                {
                    TestAssert.Equal(baseline.Length, math.Length);
                    foreach (var (before, after) in baseline.Zip(math))
                    {
                        TestAssert.True(before.Text == after.Text && before.Y == after.Y,
                            "Bookmarks and unchanged-format run splitting retain every glyph and mathematical row baseline.");
                        TestAssert.True(Math.Abs(before.X - after.X) < .001d,
                            "Unchanged-format source transformations preserve absolute grouped glyph positions: " +
                            before.X + " versus " + after.X + ".");
                    }
                }
            }
        }
    }

    public static void DocxInlineCompositionGroupedHyperlinksFollowPrintedReviewRows()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for printed grouped links.");
        var document = ReadMutation("CompositionProbes/math-printed-frame-script-anchor12-multiple",
            OoxPdfDocxMarkupMode.AllMarkup, true, false, (xml, w, m) =>
            {
                var equation = xml.Descendants(m + "oMath").First();
                var paragraph = equation.Ancestors(w + "p").First();
                var properties = paragraph.Element(w + "pPr");
                if (properties is null) { properties = new XElement(w + "pPr"); paragraph.AddFirst(properties); }
                properties.Add(new XElement(w + "ind", new XAttribute(w + "right", "7160")));
                equation.ReplaceNodes(new XElement(m + "r", new XElement(w + "rPr", new XElement(w + "rFonts",
                    new XAttribute(w + "ascii", "Cambria Math"), new XAttribute(w + "hAnsi", "Cambria Math"))),
                    new XElement(m + "t", "(α+β=γ)")));
                equation.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), new XElement(equation)));
            });
        var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup,
            OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
        var page = renderer.RenderBlankPages(document, null, default).Single();
        var mathFonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
            .Select(resource => resource.Font).ToHashSet();
        var glyphs = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
            .Where(show => mathFonts.Contains(show.Font) && "(α+β=γ)".Contains(show.Text.Normalize(NormalizationForm.FormKC), StringComparison.Ordinal)).ToArray();
        TestAssert.Equal("(α+β=γ)", string.Concat(glyphs.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
        var rows = glyphs.GroupBy(show => show.Y).ToArray();
        TestAssert.Equal(2, rows.Length);
        var annotations = page.Annotations.Where(annotation => annotation.IsDestination).ToArray();
        TestAssert.Equal(rows.Length, annotations.Length);
        foreach (var row in rows)
            TestAssert.True(annotations.Any(annotation => annotation.Y <= row.Key && annotation.Y + annotation.Height >= row.Key &&
                annotation.X <= row.Min(show => show.X) + .01d && annotation.X + annotation.Width >= row.Max(show => show.X) - .01d),
                "Review transforms must move hyperlink bounds with their painted mathematical rows.");
    }

    public static void DocxInlineCompositionRunSplittingPreservesMeasurementAndSourceOwners()
    {
        var source = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "atomic-math-unary")
            .Paragraphs.Single(paragraph => paragraph.Equations.Count != 0);
        var equation = source.Equations.Single();
        foreach (double spacing in new[] { 0d, .75d })
        {
            var whole = source.Runs[0] with { Text = "B F ", SourceRunIndex = 0,
                SourceTextOffsetInRun = 0, CharacterSpacingPoints = spacing };
            var before = whole with { Text = "B" };
            var after = whole with { Text = " F ", SourceRunIndex = 1 };
            DocxMeasuredInlineLine Measure(params DocxTextRun[] runs)
            {
                var paragraph = source with { Runs = runs, InlineContent = runs
                    .Select(run => (DocxInlineItem)new DocxInlineText(run, 0, run.Text.Length))
                    .Append(new DocxInlineEquation(equation)).ToArray() };
                TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph,
                    new Dictionary<DocxEquation, DocxPreparedMathParagraph> { [equation] = new(equation,
                        new(10d, 6d, 1d, []), 9d, 16d) }, new TrackingMeasurer(), 12d,
                    200d, 200d, 9d, 16d, 1d, default, out var lines), "The supported source paragraph must measure.");
                return TestAssert.NotNull(lines).Single();
            }
            var original = Measure(whole);
            var split = Measure(before, after);
            TestAssert.True(Math.Abs(original.Advance - split.Advance) < .000001d,
                "Unchanged formatting preserves the total advance, including boundary kerning and spacing.");
            TestAssert.True(Math.Abs(split.Advance - (26d + 3d * spacing)) < .000001d, "The expected pair and character-spacing advances survive.");
            var owners = split.Fragments.SelectMany(fragment => fragment.Text).ToArray();
            TestAssert.True(owners.Any(span => ReferenceEquals(span.StyleRun, before) && span.SourceTextRunIndex == 0) &&
                owners.Any(span => ReferenceEquals(span.StyleRun, after) && span.SourceTextRunIndex == 1),
                "Shared measurement state never merges annotation/source ownership.");
            var changed = Measure(before, after with { Bold = !after.Bold });
            TestAssert.True(Math.Abs(changed.Advance - split.Advance - 2d) < .000001d,
                "An actual text-state change retains its existing boundary fallback.");
        }
    }

    public static void DocxInlineCompositionAtomicPlacementRetainsSourceBoundAndOverwidthFallbacks()
    {
        var paragraph = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "atomic-math-unary")
            .Paragraphs.Single(value => value.Equations.Count != 0);
        var equation = paragraph.Equations.Single();
        bool Measure(bool relocation, double width, out IReadOnlyList<DocxMeasuredInlineLine>? lines) =>
            DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph,
                new Dictionary<DocxEquation, DocxPreparedMathParagraph> { [equation] = new(equation,
                    new(width, 6d, 1d, []) { AllowsInlineRelocation = relocation }, 9d, 16d) },
                new TrackingMeasurer(), 12d, 96d, 96d, 9d, 16d, 1d, default, out lines);
        TestAssert.True(!Measure(false, 30d, out var rejected) && rejected is null,
            "A source-bound box cannot move to the following line or publish a partial result.");
        TestAssert.True(Measure(true, 30d, out var moved) && TestAssert.NotNull(moved).Count > 1,
            "A complete supported box can move without changing the text owner.");
        TestAssert.Equal(1, TestAssert.NotNull(moved).SelectMany(line => line.Fragments).Count(fragment => fragment.Math is not null));
        TestAssert.True(!Measure(true, 130d, out rejected) && rejected is null,
            "Relocation never permits a box wider than every available line.");
    }

    public static void DocxInlineCompositionAtomicUnaryMatchesIndependentWordPlacement()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for independent atomic placement.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var reference = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-atomic-literal-office-geometry.json")));
        var expected = reference.RootElement;
        string input = Path.Combine(cases, "atomic-math-unary.docx");
        TestAssert.Equal(expected.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = renderer.RenderBlankPages(DocxTests.ReadDocx(input, mode), diagnostics.Add, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var mathFonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                .Select(resource => resource.Font).ToHashSet();
            var math = shows.Where(show => mathFonts.Contains(show.Font)).ToArray();
            TestAssert.Equal(expected.GetProperty("MathText").GetString(), string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.True(math.All(show => Math.Abs(show.Y - expected.GetProperty("MathBaselineY").GetDouble()) < .15d) &&
                Math.Abs(math[0].X - expected.GetProperty("MathOriginX").GetDouble()) < .15d,
                "The complete unary row must move to Word's independently measured origin and baseline.");
            var after = shows.Single(show => show.Text.TrimStart().StartsWith("after", StringComparison.Ordinal));
            TestAssert.True(Math.Abs(VisibleTextStartX(after) - expected.GetProperty("AfterX").GetDouble()) < .2d &&
                Math.Abs(after.Y - expected.GetProperty("AfterBaselineY").GetDouble()) < .15d &&
                Math.Abs(shows.Single(show => show.Text.StartsWith("Following prose.", StringComparison.Ordinal)).Y -
                    expected.GetProperty("FollowingBaselineY").GetDouble()) < .15d,
                "Following prose retains independently measured placement after atomic relocation.");
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Qualified relocation cannot omit the unary formula.");
        }
    }

    public static void DocxInlineCompositionAtomicFormulaHyperlinksCoverPaintedGlyphs()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for atomic formula hyperlinks.");
        foreach (string fixture in new[] { "docx-omml-inline-fraction-middle12", "docx-omml-inline-script-middle12", "CompositionProbes/atomic-math-unary" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation(fixture, mode, true, false, (xml, w, m) =>
            {
                var equation = xml.Descendants(m + "oMath").First();
                equation.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), new XElement(equation)));
            });
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var page = renderer.RenderBlankPages(document, null, default).Single();
            var mathFonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal))
                .Select(resource => resource.Font).ToHashSet();
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true)
                .Where(show => mathFonts.Contains(show.Font)).ToArray();
            TestAssert.True(math.Length != 0, "The complete atomic formula must remain visible.");
            var annotation = page.Annotations.Single(item => item.IsDestination);
            TestAssert.True(math.All(show => annotation.X <= show.X + .01d && annotation.X + annotation.Width >= show.X - .01d &&
                annotation.Y <= show.Y + .01d && annotation.Y + annotation.Height >= show.Y - .01d),
                "A complete atomic formula's hyperlink must include every painted glyph in both presets.");
        }
    }

    public static void DocxInlineCompositionLiteralBreakKeepsMalformedFallbacks()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the literal fallback control.");
        foreach (string literal in new[] { "(α+β=γ", "([α+β)=γ]" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("docx-omml-inline-overfull12", mode, false, false, (xml, w, m) =>
            {
                var texts = xml.Descendants(m + "t").ToArray();
                texts[0].Value = literal;
                foreach (var text in texts.Skip(1)) text.Value = string.Empty;
            });
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.True(diagnostics.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"),
                "Malformed grouping retains its source-line fallback.");
            TestAssert.True(pages.All(page => page.Fonts.All(font => !font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))),
                "No partial or unused equation resource may escape the fallback.");
        }
    }

    public static void DocxInlineCompositionLiteralOperatorBreakRetainsOfficeInlineSettingBehavior()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the inline operator control.");
        foreach (string setting in new[] { "default", "before", "after", "repeat" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-literal-inline-break-" + setting, mode, true, true);
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            foreach (string glyph in new[] { "α", "+", "β", "=", "γ" })
                TestAssert.Equal(1, text.Split(glyph, StringSplitOptions.None).Length - 1);
            TestAssert.Contains("Following prose.", text);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "A bookmark and run splitting cannot remove the broken literal.");
            TestAssert.Equal(2, renderer.InspectLayout(document).Pages.Single().Items.Count(item => item.Kind == "TextLine" && item.SourceBlockIndex == 0));
        }
    }

    public static void DocxInlineCompositionNumberedMathPreservesFourStoryContent()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the numbered composition control.");
        foreach (string story in new[] { "", "-cell", "-header", "-footer" })
        foreach (bool numbered in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-composition-" + (numbered ? "numbered" : "unnumbered") + story, mode, false, true);
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(2, text.Split("x+2y", StringSplitOptions.None).Length - 1);
            TestAssert.Equal(2, text.Split("a+b=c", StringSplitOptions.None).Length - 1);
            foreach (string label in new[] { "1.", "2.", "3." })
                TestAssert.Equal(numbered ? 1 : 0, text.Split(label, StringSplitOptions.None).Length - 1);
            TestAssert.Contains("share the same input series.", text);
            string part = story == "-header" ? "/word/header1.xml" : story == "-footer" ? "/word/footer1.xml" : document.MainPartName;
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION" || d.PartName != part),
                "Ordinary numbering and unchanged-format run splitting cannot remove supported equations.");
        }
    }

    public static void DocxInlineCompositionMathTableTailAdmissionMatchesWordBoundaryControls()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the page-boundary control.");
        foreach (string suffix in new[] { "0", "6", "24", "6-cant-split", "24-tight" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/math-table-boundary-after" + suffix, mode, false, false);
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(document, null, default).ToArray();
            int expectedPage = suffix.Contains("cant-split", StringComparison.Ordinal) || suffix.Contains("tight", StringComparison.Ordinal) ? 1 : 0;
            TestAssert.Contains("x+2y", DocxPageFootnoteNumberingTests.DecodeText(pages[expectedPage]).Normalize(NormalizationForm.FormKC));
            var snapshot = renderer.InspectLayout(document);
            var row = snapshot.Pages[expectedPage].TableRows.Single(row => row.RowIndex == 1);
            TestAssert.Equal(1, row.FragmentCount);
            TestAssert.True(row.Cells.All(cell => cell.Y >= snapshot.Pages[expectedPage].MarginBottom),
                "Discarding unpainted spacing cannot admit the row below its body frame.");
        }
    }

    public static void DocxInlineCompositionAutofitMathRespectsAdditionalInlineContent()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the intrinsic-width control.");
        const string fixture = "CompositionProbes/math-cell-spacing";
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var source = ReadMutation(fixture, mode, false, false);
            double originalWidth = renderer.InspectLayout(source).Pages.Single().TableRows[0].Cells[1].Width;
            foreach (bool duplicate in new[] { false, true })
            {
                var document = ReadMutation(fixture, mode, false, false, (part, w, m) =>
                {
                    foreach (var math in part.Descendants(m + "oMath").ToArray())
                    {
                        if (duplicate) math.AddAfterSelf(new XElement(math));
                        else
                        {
                            math.AddBeforeSelf(new XElement(w + "r", new XElement(w + "t", "Ratio ")));
                            math.AddAfterSelf(new XElement(w + "r", new XElement(w + "t", " remains visible.")));
                        }
                    }
                });
                double width = renderer.InspectLayout(document).Pages.Single().TableRows[0].Cells[1].Width;
                TestAssert.True(width > originalWidth, "Additional measured inline content must contribute to autofit width.");
                var diagnostics = new List<OoxPdfDiagnostic>();
                var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
                string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
                TestAssert.Equal(duplicate ? 8 : 4, text.Split("x+2y", StringSplitOptions.None).Length - 1);
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION" || d.PartName != document.MainPartName),
                    "Adding prose or another formula cannot discard supported cell equations.");
            }
        }
    }

    public static void DocxInlineCompositionSingleMathLinksTargetMeasuredFormula()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the formula target control.");
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "CompositionProbes", "single-math-annotations.docx"));
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(path, mode);
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(document, null, default).ToArray();
            var targets = pages.SelectMany(page => page.Annotations).Where(annotation => annotation.Destination is not null)
                .Select(annotation => annotation.Destination!.Value).ToArray();
            var prefixes = renderer.InspectTextEmission(document).Lines.SelectMany(line => line.Segments)
                .Where(segment => !segment.IsTerminalLineSpace && segment.TextLength == "Before formula: ".Length).ToArray();
            TestAssert.Equal(2, prefixes.Length);
            TestAssert.Equal(2, targets.Length);
            for (int index = 0; index < targets.Length; index++)
                TestAssert.True(Math.Abs(targets[index].Left!.Value - prefixes[index].X - prefixes[index].Width + 3d) < 0.01d,
                    "The bookmark must target the formula after its measured prefix, rather than following prose.");
            TestAssert.Equal(2, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC)
                .Split("a+bc", StringSplitOptions.None).Length - 1);
        }
    }

    public static void DocxInlineCompositionHyperlinkPreservesSingleMathFrame()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the annotation composition control.");
        foreach (string fixture in new[] { "docx-omml-inline-middle12", "docx-omml-inline-fraction-middle12", "docx-omml-inline-script-middle12" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation(fixture, mode, true, false);
            var linked = ReadMutation(fixture, mode, true, false, (document, w, m) =>
            {
                var run = document.Descendants(w + "p").First(paragraph => paragraph.Elements(m + "oMath").Any())
                    .Elements(w + "r").First(run => run.Element(w + "t") is not null);
                run.ReplaceWith(new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), new XElement(run)));
            });
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.InspectLayout(source).Pages.SelectMany(page => page.Items).Where(item => item.Kind == "TextLine").ToArray();
            var after = renderer.InspectLayout(linked).Pages.SelectMany(page => page.Items).Where(item => item.Kind == "TextLine").ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int index = 0; index < before.Length; index++)
                TestAssert.True(Math.Abs(before[index].X - after[index].X) < 0.000001d &&
                    Math.Abs(before[index].Y - after[index].Y) < 0.000001d &&
                    Math.Abs(before[index].Width - after[index].Width) < 0.000001d &&
                    before[index].LineHeightPoints == after[index].LineHeightPoints,
                    "An unchanged-format hyperlink cannot switch a mathematical paragraph's measurement or placement.");
        }
    }

    public static void DocxInlineCompositionHiddenRunPreservesMathGeometry()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the visible geometry control.");
        const string fixture = "docx-omml-isolated-inline-body-fraction-left";
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var source = ReadMutation(fixture, mode, false, false);
            var hidden = ReadMutation(fixture, mode, false, false, (document, w, m) =>
                document.Descendants(m + "oMath").Single().AddBeforeSelf(new XElement(w + "r",
                    new XElement(w + "rPr", new XElement(w + "vanish")), new XElement(w + "t", "Invisible prose"))));
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var before = renderer.InspectLayout(source).Pages.Single().Items.Single(item => item.Kind == "Equation");
            var after = renderer.InspectLayout(hidden).Pages.Single().Items.Single(item => item.Kind == "Equation");
            TestAssert.True(before.X == after.X && before.Y == after.Y && before.Width == after.Width && before.Height == after.Height,
                "Invisible source prose must not alter mathematical style, alignment or occupied geometry.");
            var pages = renderer.RenderBlankPages(hidden, null, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText));
            TestAssert.True(!text.Contains("Invisible", StringComparison.Ordinal), "Hidden prose cannot paint.");
            TestAssert.Equal(1, text.Normalize(NormalizationForm.FormKC).Split("x+2y", StringSplitOptions.None).Length - 1);
        }
    }

    public static void DocxInlineCompositionExplicitDisplaysRetainSourceRowsAndFallbacks()
    {
        var paragraph = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "math-explicit-display-same-group")
            .Paragraphs.Single(value => value.Equations.Count == 2);
        TestAssert.True(paragraph.Equations.All(equation => equation.Display), "The probe uses explicit display source containers.");
        TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
        var math = paragraph.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(10d, 6d, 1d, []), 9d, 16d) { UsesDisplayStyle = true });
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, math, new TrackingMeasurer(),
            12d, 100d, 100d, 9d, 16d, 1d, default, out var lines), "Explicit display rows share ordered measurement.");
        TestAssert.True(TestAssert.NotNull(lines).Single().IsDisplayGroup, "The complete math-only row retains display alignment.");
        var prose = new DocxTextRun("Visible prose", 12d, null, false, false, false, null, null);
        var mixed = paragraph with { Runs = [prose], InlineContent = [new DocxInlineText(prose, 0, prose.Text.Length), .. paragraph.InlineContent!] };
        TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(mixed).Count);
        foreach (string alignment in new[] { "left", "right" })
        {
            var equations = paragraph.Equations.Select(equation => equation with { AlignmentValue = alignment }).ToArray();
            var aligned = paragraph with { Equations = equations, InlineContent = equations.Select(equation => (DocxInlineItem)new DocxInlineEquation(equation)).ToArray() };
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(aligned).Count);
        }
        var conflicting = paragraph.Equations.Select((equation, index) => equation with { AlignmentValue = index == 0 ? "left" : "right" }).ToArray();
        TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(paragraph with { Equations = conflicting,
            InlineContent = conflicting.Select(equation => (DocxInlineItem)new DocxInlineEquation(equation)).ToArray() }).Count);
        TestAssert.True(!DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, math, new TrackingMeasurer(),
            12d, 8d, 8d, 9d, 16d, 1d, default, out _), "An individual box wider than every line retains fallback.");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathParagraphPolicy.GetPreparationEquations(paragraph, false, cancellation.Token));
    }

    public static void DocxInlineCompositionExplicitDisplaysMatchWordAndSourceTransforms()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for explicit display groups.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-explicit-display-office-geometry.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (bool transformed in new[] { false, true })
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(cases, id + ".docx")))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(ReadMutation("CompositionProbes/" + id, mode, transformed, transformed), diagnostics.Add, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).ToArray();
            var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
            var math = shows.Where(show => fonts.Contains(show.Font)).ToArray();
            TestAssert.Equal(row.GetProperty("MathText").GetString(), string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(10, math.Length);
            TestAssert.True(math.All(show => Math.Abs(show.Y - row.GetProperty("BaselineY").GetDouble()) < .3d) &&
                Math.Abs(math[0].X - row.GetProperty("OriginX").GetDouble()) < .3d, "The complete display row retains independent Word placement.");
            var operators = math.Where(show => show.Text is "+" or "=").ToArray();
            var expected = row.GetProperty("Operators").EnumerateArray().ToArray();
            TestAssert.Equal(expected.Length, operators.Length);
            for (int index = 0; index < operators.Length; index++)
                TestAssert.True(operators[index].Text == expected[index].GetProperty("Text").GetString() &&
                    Math.Abs(operators[index].X - expected[index].GetProperty("X").GetDouble()) < .35d,
                    "Display-group operator advances follow independent Word geometry.");
            var following = shows.Single(show => show.Text.StartsWith("Foll", StringComparison.Ordinal));
            TestAssert.True(Math.Abs(following.Y - row.GetProperty("FollowingBaselineY").GetDouble()) < .75d,
                "Following prose retains the complete display frame; the remaining sub-point pitch residual stays explicit.");
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "Explicit display containers cannot discard either supported formula.");
        }
    }

    public static void DocxInlineCompositionExplicitDisplayLinksRetainSourceAndResolvedStyleOwnership()
    {
        foreach (string kind in new[] { "body", "cell", "footer", "fraction" })
        {
            var document = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "math-explicit-display-linked-" + kind);
            var paragraph = DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, default)
                .Concat(DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default).Select(owner => owner.Paragraph))
                .Single(value => value.Equations.Count != 0);
            var linked = TestAssert.NotNull(DocxInlineContent.GetHyperlinkedEquations(paragraph, default));
            TestAssert.True(paragraph.Equations.All(equation => equation.Display && linked.Contains(equation)),
                "Hyperlink ownership retains the actual explicit source equations.");
            TestAssert.Equal(paragraph.Equations.Count, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.Equal(0, DocxMathParagraphPolicy.GetDisplaySourceEquations(paragraph, default).Count);
            var overlapping = paragraph with { Hyperlinks = paragraph.Hyperlinks.Concat(paragraph.Hyperlinks).ToArray() };
            TestAssert.Equal(linked.Count, TestAssert.NotNull(DocxInlineContent.GetHyperlinkedEquations(overlapping, default)).Count);
            var math = paragraph.Equations.ToDictionary(equation => equation,
                equation => new DocxPreparedMathParagraph(equation, new(10d, 6d, 1d, []), 9d, 16d) { UsesDisplayStyle = false });
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, math, new TrackingMeasurer(),
                12d, 600d, 600d, 9d, 16d, 1d, default, out var lines), "Linked display source shares inline measurement.");
            var measured = TestAssert.NotNull(lines);
            TestAssert.True(measured.All(line => !line.IsDisplayGroup) &&
                measured.SelectMany(line => line.Fragments).Count(fragment => fragment.Math is not null) == linked.Count,
                "Resolved inline style controls placement without losing either source formula.");
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.GetHyperlinkedEquations(paragraph, cancellation.Token));
        }
    }

    public static void DocxInlineCompositionExplicitDisplayLinksMatchIndependentWordAcrossContainers()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for linked display-source controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-explicit-display-linked-office-geometry.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            string input = Path.Combine(cases, id + ".docx");
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(DocxTests.ReadDocx(input, mode), diagnostics.Add, default).Single();
            var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
            string expectedText = row.GetProperty("MathText").GetString()!;
            TestAssert.Equal(expectedText, string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(expectedText.Length, math.Length);
            var anchors = row.GetProperty("Anchors").EnumerateArray().ToArray();
            var origin = anchors.First(anchor => anchor.GetProperty("Text").GetString() == math[0].Text.Normalize(NormalizationForm.FormKC));
            TestAssert.True(Math.Abs(math[0].X - origin.GetProperty("X").GetDouble()) < 1d,
                "Linked math retains source-container placement; the existing cell-origin residual remains explicit.");
            foreach (var anchor in anchors)
            {
                string text = anchor.GetProperty("Text").GetString()!;
                var glyph = math.Where(show => show.Text.Normalize(NormalizationForm.FormKC) == text)
                    .ElementAt(anchor.GetProperty("Occurrence").GetInt32());
                TestAssert.True(Math.Abs((glyph.X - math[0].X) - (anchor.GetProperty("X").GetDouble() - origin.GetProperty("X").GetDouble())) < .5d &&
                    Math.Abs(glyph.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                    "Linked mathematical anchors follow independent Word geometry: " + id + "/" + mode + ".");
                if (id.EndsWith("-fraction", StringComparison.Ordinal))
                    TestAssert.True(Math.Abs(glyph.Size - anchor.GetProperty("FontSize").GetDouble()) < .1d,
                        "A linked explicit fraction uses independently observed inline argument sizing.");
            }
            var link = page.Annotations.Single(annotation => annotation.IsDestination);
            foreach (var glyph in math)
                TestAssert.True(link.X <= glyph.X + .01d && link.X + link.Width >= glyph.X - .01d &&
                    link.Y <= glyph.Y + .01d && link.Y + link.Height >= glyph.Y - .01d,
                    "The source hyperlink covers every final mathematical glyph across containers.");
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "A supported source hyperlink cannot omit explicit formulas.");
        }
    }

    public static void DocxInlineCompositionAuthoredDisplayAlignmentOwnsMathRowsAndKeepsFallbacks()
    {
        var paragraph = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "math-explicit-display-right-body")
            .Paragraphs.Single(value => value.Equations.Count == 2);
        TestAssert.True(paragraph.Alignment == DocxTextAlignment.Justified &&
            DocxMathParagraphPolicy.ResolveDisplayAlignment(paragraph.Equations, default) == DocxTextAlignment.Right,
            "Authored math alignment is a separate owner from the surrounding paragraph alignment.");
        TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
        var math = paragraph.Equations.ToDictionary(equation => equation,
            equation => new DocxPreparedMathParagraph(equation, new(10d, 6d, 1d, []), 9d, 16d));
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, math, new TrackingMeasurer(),
            12d, 100d, 100d, 9d, 16d, 1d, default, out var lines), "A math-owned paragraph alignment must measure.");
        TestAssert.True(TestAssert.NotNull(lines).Single().DisplayAlignment == DocxTextAlignment.Right,
            "The measured line retains one display alignment instead of a centering flag.");
        var prose = new DocxTextRun("Visible prose", 12d, null, false, false, false, null, null);
        var lineBreak = prose with { Text = "\n" };
        var mixed = paragraph with { Runs = [lineBreak, prose], InlineContent = [.. paragraph.InlineContent!,
            new DocxInlineBreak(lineBreak, 0), new DocxInlineText(prose, 0, prose.Text.Length)] };
        TestAssert.Equal(0, DocxMathParagraphPolicy.GetPreparationEquations(mixed).Count);
        var unknown = paragraph.Equations.Select(equation => equation with { AlignmentValue = "unknown" }).ToArray();
        TestAssert.True(DocxMathParagraphPolicy.ResolveDisplayAlignment(unknown, default) is null,
            "An unknown alignment retains an explicit fallback.");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxMathParagraphPolicy.ResolveDisplayAlignment(paragraph.Equations, cancellation.Token));
    }

    public static void DocxInlineCompositionAuthoredDisplayAlignmentMatchesIndependentWord()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for authored display alignment.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-explicit-display-alignment-office-geometry.json")));
        foreach (var row in references.RootElement.GetProperty("Rows").EnumerateArray())
        foreach (bool transformed in new[] { false, true })
        {
            string id = row.GetProperty("Id").GetString()!;
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(cases, id + ".docx")))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var document = ReadMutation("CompositionProbes/" + id, mode, transformed, transformed,
                partOverride: id.EndsWith("-footer", StringComparison.Ordinal) ? "word/math-tabs-footer.xml" : null);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
                .RenderBlankPages(document, diagnostics.Add, default).Single();
            var fonts = page.Fonts.Where(resource => resource.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(resource => resource.Font).ToHashSet();
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true).Where(show => fonts.Contains(show.Font)).ToArray();
            TestAssert.Equal(row.GetProperty("MathText").GetString(), string.Concat(math.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(10, math.Length);
            var anchors = row.GetProperty("Anchors").EnumerateArray().ToArray();
            var origin = anchors.Single(anchor => anchor.GetProperty("Text").GetString() == math[0].Text.Normalize(NormalizationForm.FormKC));
            TestAssert.True(Math.Abs(math[0].X - origin.GetProperty("X").GetDouble()) < 1d,
                "Authored display alignment follows independent Word placement, including the retained cell-origin residual.");
            foreach (var anchor in anchors)
            {
                var glyph = math.Where(show => show.Text.Normalize(NormalizationForm.FormKC) == anchor.GetProperty("Text").GetString())
                    .ElementAt(anchor.GetProperty("Occurrence").GetInt32());
                TestAssert.True(Math.Abs((glyph.X - math[0].X) - (anchor.GetProperty("X").GetDouble() - origin.GetProperty("X").GetDouble())) < .5d &&
                    Math.Abs(glyph.Y - anchor.GetProperty("Y").GetDouble()) < 1d,
                    "Authored display glyphs retain independent local geometry and static/cell baselines: " + id + ".");
            }
            TestAssert.True(diagnostics.All(item => item.Id != "DOCX_UNSUPPORTED_EQUATION"), "A supported authored alignment cannot omit the formulas.");
        }
    }

    public static void DocxInlineCompositionMathOnlyGroupsCenterAcrossParagraphAlignment()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the display-group control.");
        string root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases"));
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            double singleWidth = renderer.InspectLayout(DocxTests.ReadDocx(Path.Combine(root,
                "CompositionProbes/math-navigation.docx"), mode)).Pages.Single().Items.Single(item => item.Kind == "Equation").Width;
            var document = DocxTests.ReadDocx(Path.Combine(root, "CompositionProbes/math-display-groups.docx"), mode);
            var page = renderer.InspectLayout(document).Pages.Single();
            double scale = DocxRenderer.ResolveWordCompatiblePrintScale(document, DocxMarkupContext.FromMode(mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup));
            // The pinned 11pt default contributes three math units between Inner atoms.
            double fractionGap = 11d / 6d * scale;
            var groups = page.Items.Where(item => item.Kind == "TextLine" && item.TextLength == 0).ToArray();
            TestAssert.Equal(3, groups.Length);
            var column = page.ColumnFrames.Single();
            for (int index = 0; index < groups.Length; index++)
            {
                int count = index == 1 ? 3 : 2;
                TestAssert.True(Math.Abs(groups[index].Width - singleWidth * count - fractionGap * (count - 1)) < 0.000001d,
                    "Adjacent fractions retain each box's measurement plus their mathematical spacing: " + groups[index].Width + "/" + singleWidth + "/" + count + "/" + fractionGap);
                TestAssert.True(Math.Abs(groups[index].X + groups[index].Width / 2d - column.X - column.Width / 2d) < 0.000001d,
                    "Word's math-only groups center independently of surrounding paragraph alignment.");
            }
            var diagnostics = new List<OoxPdfDiagnostic>();
            var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
            TestAssert.Equal(11, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC)
                .Split("x+2y", StringSplitOptions.None).Length - 1);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "All fitting group and mixed-row formulas must survive.");
        }
    }

    public static void DocxInlineCompositionMixedMathBookmarksRetainOrderedSourceTargets()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "CompositionProbes", "inline-math-navigation.docx"));
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = DocxTests.ReadDocx(path, mode);
            var paragraphs = DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, default)
                .Concat(DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default).Select(owner => owner.Paragraph))
                .Where(paragraph => paragraph.Equations.Count == 2).ToArray();
            TestAssert.Equal(4, paragraphs.Length);
            foreach (var paragraph in paragraphs)
            {
                TestAssert.Equal(2, paragraph.BookmarkAnchors.Count);
                for (int index = 0; index < 2; index++)
                {
                    var bookmark = paragraph.BookmarkAnchors[index];
                    TestAssert.True(bookmark.InlineItemIndex is >= 0 &&
                        paragraph.InlineContent![bookmark.InlineItemIndex.Value] is DocxInlineEquation item &&
                        ReferenceEquals(item.Equation, paragraph.Equations[index]),
                        "Each bookmark must retain its own formula even when run-relative positions coincide.");
                }
            }
        }
    }

    public static void DocxInlineCompositionCompoundMathBookmarksRetainFourStoryDestinations()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the compound navigation control.");
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "CompositionProbes", "math-navigation.docx"));
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(DocxTests.ReadDocx(path, mode), null, default).ToArray();
            var destinations = pages.SelectMany(page => page.Annotations).Where(a => a.Destination is not null)
                .Select(a => a.Destination!.Value).ToArray();
            TestAssert.Equal(4, destinations.Length);
            TestAssert.True(destinations.All(d => d.PageIndex == 0 && d.Left is >= 0d and < 612d && d.Top is > 0d and < 792d),
                "Each mathematical story target must resolve inside its emitted page frame.");
            TestAssert.Equal(4, string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC)
                .Split("x+2y", StringSplitOptions.None).Length - 1);
        }
    }

    public static void DocxInlineCompositionEquationOnlyBookmarkResolvesInternalNavigation()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the navigation control.");
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string fixture in new[] { "docx-omml-isolated-inline-body-fraction-left", "docx-omml-isolated-inline-cell-fraction-left",
            "docx-omml-static-header-fraction", "docx-omml-static-footer-fraction" })
        {
            var source = ReadMutation(fixture, mode, true, false);
            var link = DocxTests.CreateDocxLayoutParagraph("Jump to calculation", 11d, 14d) with
            {
                Hyperlinks = [new DocxHyperlinkSpan(null, "PublicCalculation", null, null, null, null, null, 0, 1, 0, 1, 19)]
            };
            var document = source with { BodyElements = [new DocxParagraphElement(link), .. source.BodyElements] };
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(document, null, default).ToArray();
            var annotation = pages.SelectMany(page => page.Annotations).SingleOrDefault(a => a.Destination is not null);
            TestAssert.True(annotation.Destination is { PageIndex: >= 0, Top: > 0d },
                "A painted equation-only bookmark must retain its internal link destination: " + fixture + "/" + mode);
            TestAssert.True(string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC)
                .Contains("x+2y", StringComparison.Ordinal), "Navigation must not remove the mathematical content.");
        }
    }

    public static void DocxInlineCompositionReviewStaticStoriesPreserveSharedBodyCapacity()
    {
        foreach (string kind in new[] { "body", "header", "footer" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
                "ReviewBodyFrames", kind + ".docx"));
            var document = DocxTests.ReadDocx(path, mode);
            var renderer = new DocxRenderer(new TestFaceFontResolver(), mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(document, null, default).ToArray();
            TestAssert.Equal(2, pages.Length);
            TestAssert.Contains("Following body", DocxPageFootnoteNumberingTests.DecodeText(pages[1]));
        }
    }

    public static void DocxInlineCompositionReviewPaginationRetainsSourcePageCapacity()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Office fonts are required for this independent Word control.");
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(ReadMutation("Composition/report", mode, false, false), null, default).ToArray();
            // Independent Word exports keep these sections on the same source
            // pages in both modes, including the header, table, image and note.
            TestAssert.Equal(2, pages.Length);
            TestAssert.Contains("Review section 8.", DocxPageFootnoteNumberingTests.DecodeText(pages[0]));
            TestAssert.Contains("Review section 9.", DocxPageFootnoteNumberingTests.DecodeText(pages[1]));
        }
    }

    public static void DocxInlineCompositionMeasurementPreservesSourceTrackingKerningAndHiddenRuns()
    {
        var document = ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false, (source, w, m) =>
        {
            var equations = source.Descendants(m + "oMath").Take(2).Select(e => new XElement(e)).ToArray();
            var drawing = new XElement(source.Descendants(w + "drawing").Single().Parent!);
            XElement Tracked(string value) => new(w + "r", new XElement(w + "rPr", new XElement(w + "spacing", new XAttribute(w + "val", "20"))), new XElement(w + "t", value));
            source.Root!.Element(w + "body")!.ReplaceNodes(new XElement(w + "p", Tracked("AB CD"),
                new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "vanish")), new XElement(w + "t", "Hidden")),
                drawing, Tracked("EF GH"), equations));
        });
        var paragraph = document.Paragraphs.Single();
        var math = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var equation in paragraph.Equations)
            math.Add(equation, new(equation, new DocxMathBox(18d, 10d, 4d, []), 11d, 16d));
        TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, math, new TrackingMeasurer(),
            11d, 200d, 200d, 9d, 12d, 1d, default, out var measured), "The supported mixed paragraph must measure.");
        var line = TestAssert.NotNull(measured).Single();
        TestAssert.Equal(126d, line.Advance); // Two 27pt text runs (including pair kerning), a 36pt image and two 18pt equations.
        TestAssert.Equal("AB CDEF GH", string.Concat(line.Fragments.SelectMany(fragment => fragment.Text).Select(span => span.Text)));
        foreach (var fragment in line.Fragments.Where(fragment => fragment.Source is DocxInlineText))
        {
            var text = (DocxInlineText)fragment.Source;
            var span = fragment.Text.Single();
            TestAssert.Equal(text.Run.SourceRunIndex, span.SourceTextRunIndex);
            TestAssert.Equal(text.Run.SourceTextOffsetInRun + text.Start, span.SourceTextOffsetInRun);
        }
    }

    public static void DocxInlineCompositionKeepNextReservesTheMeasuredMixedFrame()
    {
        var source = ReadPlacedComposition(OoxPdfDocxMarkupMode.Final, "body");
        var paragraph = source.Paragraphs.Single() with { KeepRules = source.Paragraphs.Single().KeepRules with { KeepNext = true } };
        var before = DocxTests.CreateDocxLayoutParagraph("Before", 10d, 10d) with { SpacingAfterPoints = 24d };
        var after = DocxTests.CreateDocxLayoutParagraph("Following", 10d, 10d);
        var document = DocxTests.CreateLayoutTestDocument([
            new DocxParagraphElement(before), new DocxParagraphElement(paragraph), new DocxParagraphElement(after)], [])
            with { PageHeightPoints = 100d };
        var math = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var equation in paragraph.Equations)
            math.Add(equation, new(equation, new DocxMathBox(18d, 30d, 10d, []), 30d, 40d));
        var layout = new DocxLayoutEngine(OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .Create(document, new DocxTests.FamilyWidthTextMeasurer(), default, mathEquations: math);
        int mathPage = Enumerable.Range(0, layout.Pages.Count).Single(index =>
            layout.Pages[index].Items.OfType<DocxTextLineLayout>().Any(line => line.InlineMath.Count != 0));
        int followingPage = Enumerable.Range(0, layout.Pages.Count).Single(index =>
            layout.Pages[index].Items.OfType<DocxTextLineLayout>().Any(line => line.SourceBlockIndex == 2));
        TestAssert.True(mathPage == followingPage && mathPage == 1,
            "Keep-next reserves the complete mixed frame and the following line, rather than only the picture or prose.");
    }

    public static void DocxInlineCompositionUnsupportedEquationRetainsAccuratePartialPartDiagnostics()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for this rendering control.");
        var document = ReadMutation("docx-omml-static-guard-partial", OoxPdfDocxMarkupMode.Final, false, false,
            (source, w, m) => source.Root!.Add(new XElement(w + "p", new XElement(m + "oMath", new XElement(m + "future")))),
            "word/header1.xml");
        var diagnostics = new List<OoxPdfDiagnostic>();
        var output = new DocxRenderer(null, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
            .RenderBlankPages(document, diagnostics.Add, default).ToArray();
        TestAssert.True(output.Length != 0 && diagnostics.Count(d => d.Id == "DOCX_APPROXIMATED_EQUATION" && d.PartName == "/word/header1.xml") == 1 &&
            diagnostics.Count(d => d.Id == "DOCX_UNSUPPORTED_EQUATION" && d.PartName == "/word/header1.xml") == 1,
            "A genuinely unsupported third equation retains the part warning while supported neighbors still paint.");
    }

    public static void DocxInlineCompositionWordWrappingSharesBreakRulesAcrossRunSplits()
    {
        string word = string.Concat(Enumerable.Repeat("abcdefghij", 12));
        int? expectedLines = null;
        foreach (bool split in new[] { false, true })
        {
            var document = ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false, (source, w, m) =>
            {
                var equations = source.Descendants(m + "oMath").Take(2).Select(e => new XElement(e)).ToArray();
                var paragraph = new XElement(w + "p");
                foreach (string part in split ? new[] { word[..37], word[37..81], word[81..] } : new[] { word })
                    paragraph.Add(new XElement(w + "r", new XElement(w + "t", part)));
                paragraph.Add(equations);
                source.Root!.Element(w + "body")!.ReplaceNodes(paragraph);
            });
            var paragraph = document.Paragraphs.Single();
            var math = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance);
            foreach (var equation in paragraph.Equations)
                math.Add(equation, new(equation, new DocxMathBox(18d, 10d, 4d, []), 11d, 16d));
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, math,
                new DocxTests.FamilyWidthTextMeasurer(), 11d, 50d, 50d, 9d, 12d, 1d, default, out var measured),
                "Supported character-level wrapping must remain available beside multiple equations.");
            var lines = TestAssert.NotNull(measured);
            TestAssert.Equal(word, string.Concat(lines.SelectMany(line => line.Fragments).SelectMany(fragment => fragment.Text).Select(span => span.Text)));
            TestAssert.Equal(2, lines.SelectMany(line => line.Fragments).Count(fragment => fragment.Math is not null));
            if (expectedLines is { } count) TestAssert.Equal(count, lines.Count);
            expectedLines = lines.Count;
        }
    }

    public static void DocxInlineCompositionMixedObjectsSurviveBodyCellHeaderAndFooterPlacement()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for this rendering control.");
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string placement in new[] { "body", "cell", "header", "footer", "body-review", "cell-review", "header-review", "footer-review" })
        {
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(ReadPlacedComposition(mode, placement), diagnostics.Add, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.True(text.Contains("x+2y", StringComparison.Ordinal) && text.Contains("a+b=c", StringComparison.Ordinal),
                "Both supported mathematical source objects reach the PDF in " + placement + "/" + mode);
            TestAssert.Equal(1, pages.Sum(page => page.Images.Count));
            TestAssert.Equal(1, pages.Sum(page => page.Images.Sum(image =>
                System.Text.RegularExpressions.Regex.Matches(page.Content,
                    "/" + System.Text.RegularExpressions.Regex.Escape(image.ResourceName) + @"\s+Do(?:\s|$)").Count)));
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Moving supported content into a container cannot silently drop an equation: " + placement + "/" + mode);
        }
    }

    public static void DocxInlineCompositionReportPreservesEverySupportedEquationAndInlinePicture()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for this rendering control.");
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("Composition/report", mode, false, false);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var renderer = new DocxRenderer(null, mode, mode == OoxPdfDocxMarkupMode.Final
                ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var pages = renderer.RenderBlankPages(document, diagnostics.Add, default).ToArray();
            string text = string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.True(text.Split("x+2y", StringSplitOptions.None).Length - 1 == 3,
                "All three source fractions must reach the PDF in " + mode);
            TestAssert.True(text.Split("a+b=c", StringSplitOptions.None).Length - 1 >= 3,
                "Both body literal equations and the repeated header equation must remain visible.");
            TestAssert.True(pages.Sum(page => page.Images.Count) == 1 &&
                diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "The inline picture and all six source equations survive ordinary feature composition.");
            TestAssert.True(pages.SelectMany(page => page.Fonts).Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal)),
                "Mathematical content reaches PDF resources and painting.");
        }
    }

    public static void DocxInlineCompositionMeasurementKeepsMultipleEquationsAndPicturesInSourceOrder()
    {
        var document = ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false);
        foreach (var paragraph in document.Paragraphs.Where(p => p.Equations.Count > 1 || p.Images.Count != 0))
        {
            var math = new Dictionary<DocxEquation, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance);
            foreach (var equation in paragraph.Equations)
                math.Add(equation, new(equation, new DocxMathBox(18d, 10d, 4d, []), 11d, 16d));
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, math,
                new DocxTests.FamilyWidthTextMeasurer(), 11d, 160d, 160d, 9d, 12d, 1d, default, out var lines),
                "Fitting source objects share one measured paragraph rather than competing for a single math slot.");
            var measured = TestAssert.NotNull(lines);
            var fragments = measured.SelectMany(line => line.Fragments).ToArray();
            TestAssert.True(fragments.Where(fragment => fragment.Math is not null).Select(fragment => fragment.Math!.Equation)
                .SequenceEqual(paragraph.Equations), "All mathematical source identities survive wrapping in order.");
            TestAssert.Equal(paragraph.Images.Count, fragments.Count(fragment => fragment.Source is DocxInlinePicture));
            TestAssert.True(measured.All(line => line.Advance <= 160d && line.OccupiedHeight >= 12d &&
                line.BaselineInset >= 9d), "Each line owns finite fitting advance and its complete occupied frame.");
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph, math,
                new DocxTests.FamilyWidthTextMeasurer(), 11d, 160d, 160d, 9d, 12d, 1d, cancelled.Token, out _));
        }
    }

    public static void DocxInlineCompositionSourceViewOrdersEquationsPicturesTabsAndBreaks()
    {
        DocxDocument document = ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false,
            (source, w, m) =>
            {
                var equation = new XElement(source.Descendants(m + "oMath").First());
                var drawing = new XElement(source.Descendants(w + "drawing").Single());
                source.Root!.Element(w + "body")!.ReplaceNodes(new XElement(w + "p",
                    new XElement(w + "bookmarkStart", new XAttribute(w + "id", "9"), new XAttribute(w + "name", "Mixed")),
                    equation,
                    new XElement(w + "r", new XElement(w + "t", "ABC"), drawing, new XElement(w + "tab"),
                        new XElement(w + "t", "DEF"), new XElement(w + "br")),
                    new XElement(equation), new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "9"))));
            });
        DocxParagraph paragraph = document.Paragraphs.Single();
        var items = TestAssert.NotNull(paragraph.InlineContent);
        TestAssert.Equal("DocxInlineEquation,DocxInlineText,DocxInlinePicture,DocxInlineTab,DocxInlineText,DocxInlineBreak,DocxInlineEquation",
            string.Join(',', items.Select(item => item.GetType().Name)));
        TestAssert.True(ReferenceEquals(paragraph.Equations[0], ((DocxInlineEquation)items[0]).Equation) &&
            ReferenceEquals(paragraph.Equations[1], ((DocxInlineEquation)items[^1]).Equation) &&
            ReferenceEquals(paragraph.Images.Single(), ((DocxInlinePicture)items[2]).Image),
            "Ordered items retain source identities instead of reparsing or cloning visible objects.");
        TestAssert.Equal(3, paragraph.Images.Single().SourceTextOffsetInRun);
        TestAssert.Equal(1, paragraph.BookmarkAnchors.Count);
        TestAssert.Equal("ABCDEF", string.Concat(items.OfType<DocxInlineText>().Select(text => text.Run.Text.Substring(text.Start, text.Length))));
        TestAssert.True(DocxInlineContent.TryGetEquationTextOffset(items, paragraph.Equations[1], out int offset) && offset == 8,
            "The second equation follows six text characters, one tab and one break; pictures have their own width.");
    }

    public static void DocxInlineCompositionSourceViewRetainsPrefixOrderAndCancellation()
    {
        DocxDocument document = ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false);
        foreach (DocxParagraph paragraph in document.Paragraphs.Where(p => p.Equations.Count != 0))
        {
            var items = TestAssert.NotNull(paragraph.InlineContent);
            TestAssert.True(items.OfType<DocxInlineEquation>().Select(item => item.Equation).SequenceEqual(paragraph.Equations),
                "Multiple equations retain their authored order.");
            TestAssert.Equal(string.Concat(paragraph.Runs.Select(run => run.Text)),
                string.Concat(items.OfType<DocxInlineText>().Select(text => text.Run.Text.Substring(text.Start, text.Length))));
        }
        var imageParagraph = document.Paragraphs.Single(p => p.Images.Count != 0);
        var ordered = TestAssert.NotNull(imageParagraph.InlineContent);
        TestAssert.True(ordered[0] is DocxInlineText && ordered[1] is DocxInlinePicture &&
            ordered[2] is DocxInlineText && ordered[3] is DocxInlineEquation,
            "A picture before the first equation joins the initial source prefix in place.");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.FromPrefix(imageParagraph.Runs, imageParagraph.Images, cancelled.Token));
    }

    public static void DocxInlineCompositionSourceViewCannotOutliveReplacedCollections()
    {
        var document = ReadMutation("docx-omml-isolated-inline-body-fraction-left", OoxPdfDocxMarkupMode.Final, false, false);
        var paragraph = document.Paragraphs.Single(p => p.Equations.Count != 0);
        TestAssert.NotNull(paragraph.InlineContent);
        TestAssert.NotNull((paragraph with { SpacingAfterPoints = 9d }).InlineContent);
        var changed = paragraph with { Runs = [new DocxTextRun("New visible prose", 11d, null, false, false, false, null, null)] };
        TestAssert.True(changed.InlineContent is null && DocxMathParagraphPolicy.GetStandaloneRow(changed) is null,
            "A legacy source replacement cannot reuse the old standalone-content eligibility.");
        TestAssert.True((paragraph with { Images = [] }).InlineContent is null &&
            (paragraph with { Equations = [] }).InlineContent is null,
            "Image and equation replacement also invalidate their borrowed source view.");
    }

    public static void DocxInlineCompositionBookmarkRetainsSupportedMathOwnership()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument source = ReadMutation("docx-omml-isolated-inline-body-fraction-left", mode, false, false);
            DocxDocument annotated = ReadMutation("docx-omml-isolated-inline-body-fraction-left", mode, true, false);
            DocxParagraph before = source.Paragraphs.Single(p => p.Equations.Count != 0);
            DocxParagraph after = annotated.Paragraphs.Single(p => p.Equations.Count != 0);
            TestAssert.Equal(1, after.BookmarkAnchors.Count);
            TestAssert.NotNull(DocxMathParagraphPolicy.GetStandaloneRow(before));
            TestAssert.NotNull(DocxMathParagraphPolicy.GetStandaloneRow(after),
                "A zero-width source annotation cannot remove supported mathematical content.");
            TestAssert.True(DocxMathLayout.TryGetRuns(before.Equations.Single(), default, out var beforeRuns) &&
                DocxMathLayout.TryGetRuns(after.Equations.Single(), default, out var afterRuns) &&
                string.Concat(beforeRuns.Select(r => r.Run.Text.Text)) == string.Concat(afterRuns.Select(r => r.Run.Text.Text)),
                "Source mathematical content survives the annotation unchanged.");
        }
    }

    public static void DocxInlineCompositionBookmarkAndRunSplittingPreserveRenderedContent()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the rendering comparison.");
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string fixture in new[] { "docx-omml-isolated-inline-body-fraction-left", "docx-omml-inline-fraction-middle12",
            "docx-omml-isolated-inline-cell-fraction-left", "docx-omml-static-header-fraction", "docx-omml-static-footer-fraction" })
        {
            var geometry = mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
            var renderer = new DocxRenderer(null, mode, geometry);
            var before = renderer.RenderBlankPages(ReadMutation(fixture, mode, false, false), null, default).ToArray();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var after = renderer.RenderBlankPages(ReadMutation(fixture, mode, true, true), diagnostics.Add, default).ToArray();
            string Text(IEnumerable<Lokad.OoxPdf.Pdf.PdfPage> pages) =>
                string.Concat(pages.Select(DocxPageFootnoteNumberingTests.DecodeText)).Normalize(NormalizationForm.FormKC);
            TestAssert.Equal(Text(before), Text(after));
            TestAssert.Equal(before.Length, after.Length);
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "A bookmark or unchanged-format run split cannot turn painted math into ignored content.");
        }
    }

    private static DocxDocument ReadPlacedComposition(OoxPdfDocxMarkupMode mode, string placement) =>
        DocxTests.ReadDocx(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "CompositionProbes", placement + ".docx")), mode);

    // Text-show boundaries are an encoding detail. Independent anchors describe
    // the first visible glyph, including when leading spaces share its show.
    private static double VisibleTextStartX((PdfEmbeddedFont Font, string Text, double Size, double X, double Y) show)
    {
        int leading = 0;
        while (leading < show.Text.Length && char.IsWhiteSpace(show.Text[leading])) { leading++; }
        return show.X + show.Font.MeasureTextPoints(show.Text[..leading], show.Size, kerningEnabled: false);
    }

    private sealed class CountingMeasurer : IDocxTextMeasurer
    {
        public Dictionary<string, int> Calls { get; } = new(StringComparer.Ordinal);
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            Calls[text] = Calls.GetValueOrDefault(text) + 1;
            return DocxTextSpacing.AddCharacterSpacing(text.Length * 5d, run, text);
        }
    }

    private sealed class TrackingMeasurer : IDocxTextMeasurer
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) =>
            DocxTextSpacing.AddCharacterSpacing(text.Length * 5d, run, text) -
            (text.Contains("B ", StringComparison.Ordinal) ? 2d : 0d) -
            (text.Contains("F ", StringComparison.Ordinal) ? 2d : 0d);
    }

    private static DocxDocument ReadMutation(string fixture, OoxPdfDocxMarkupMode mode, bool bookmark, bool split,
        Action<XDocument, XNamespace, XNamespace>? change = null, string? partOverride = null)
    {
        string source = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", fixture + ".docx"));
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        File.Copy(source, path);
        new FileInfo(path).IsReadOnly = false; // Only the temporary mutation copy is writable.
        try
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                string part = partOverride ?? (fixture.Contains("static-header", StringComparison.Ordinal) ? "word/header1.xml" :
                    fixture.Contains("static-footer", StringComparison.Ordinal) ? "word/footer1.xml" : "word/document.xml");
                var entry = zip.GetEntry(part)!;
                XDocument document;
                using (var stream = entry.Open()) document = XDocument.Load(stream);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
                var equation = document.Descendants(m + "oMath").FirstOrDefault();
                if (bookmark)
                {
                    if (equation is null) throw new InvalidOperationException("The bookmark mutation requires an existing equation.");
                    var container = equation.Parent!.Name == m + "oMathPara" ? equation.Parent : equation;
                    container.AddBeforeSelf(new XElement(w + "bookmarkStart", new XAttribute(w + "id", "41"), new XAttribute(w + "name", "PublicCalculation")));
                    container.AddAfterSelf(new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "41")));
                }
                if (split)
                {
                    foreach (var run in document.Descendants(w + "r").Where(r => r.Parent?.Name == w + "p").ToArray())
                    {
                        if (run.Element(w + "t") is not { Value.Length: > 2 } text || run.Elements().Any(e => e.Name != w + "t" && e.Name != w + "rPr")) continue;
                        int middle = text.Value.Length / 2;
                        var second = new XElement(run);
                        second.Element(w + "t")!.Value = text.Value[middle..];
                        text.Value = text.Value[..middle];
                        text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                        second.Element(w + "t")!.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                        run.AddAfterSelf(second);
                    }
                }
                change?.Invoke(document, w, m);
                entry.Delete();
                using (var output = zip.CreateEntry(part).Open()) document.Save(output);
            }
            return DocxTests.ReadDocx(path, mode);
        }
        finally { File.Delete(path); }
    }
}
