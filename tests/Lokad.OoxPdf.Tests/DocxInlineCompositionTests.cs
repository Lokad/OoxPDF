using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static class DocxInlineCompositionTests
{
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

    public static void DocxInlineCompositionFootnoteMarkersMatchIndependentPrintedOffsets()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the independent footnote marker controls.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var references = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "math-footnotes-office-markers.json")));
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
                line.Segments.Any(segment => segment.TextLength == 1 && segment.CharacterProfile.DigitCount == 1 &&
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
            string container = Path.GetFileNameWithoutExtension(input).Split('-')[^1];
            if (heights.TryGetValue(container, out double height))
                TestAssert.True(Math.Abs(height - fills[0].Height) < .001d,
                    "Independent Word controls retain one line selection height when only the annotated source range changes.");
            else heights.Add(container, fills[0].Height);
            var line = renderer.InspectTextEmission(document).Lines.Single(line =>
                line.Segments.Count(segment => !segment.IsTerminalLineSpace && segment.TextLength == 5) == 1 &&
                line.Segments.Count(segment => !segment.IsTerminalLineSpace && segment.TextLength == 6) == 1);
            double start = line.Segments.Single(segment => !segment.IsTerminalLineSpace && segment.TextLength == 5).X;
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
            var fields = renderer.InspectTextEmission(document).Lines.Single(line =>
                new[] { 5, 6, 4 }.All(length => line.Segments.Count(segment => !segment.IsTerminalLineSpace && segment.TextLength == length) == 1)).Segments;
            double start = fields.Single(segment => !segment.IsTerminalLineSpace && segment.TextLength == 5).X;
            foreach (var (length, property) in new[] { (6, "MiddleOffset"), (4, "EndOffset") })
            {
                double actual = fields.Single(segment => !segment.IsTerminalLineSpace && segment.TextLength == length).X - start;
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
            TestAssert.True(!Measure(paragraph with { TabStops = [stop] }, 300d, 1d, out _), "Unmigrated alignment and leaders retain explicit fallback.");
        TestAssert.True(!Measure(paragraph with { TabStops = [new(400d, null, "left", null)] }, 100d, 1d, out _), "An out-of-frame stop retains fallback.");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryMeasureOrderedInlineContent(paragraph,
            equations, new TrackingMeasurer(), 12d, 300d, 300d, 9d, 16d, 1d, cancellation.Token, out _));
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

    public static void DocxInlineCompositionLiteralBreakKeepsParenthesizedAndUnaryFallbacks()
    {
        if (!File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Installed Cambria Math is required for the literal fallback control.");
        foreach (string literal in new[] { "(α+β=γ)", "+αβγαβγ" })
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
                "Nested or unary-only breaking retains its qualified whole-formula fallback.");
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
            var groups = page.Items.Where(item => item.Kind == "TextLine" && item.TextLength == 0).ToArray();
            TestAssert.Equal(3, groups.Length);
            var column = page.ColumnFrames.Single();
            for (int index = 0; index < groups.Length; index++)
            {
                TestAssert.True(Math.Abs(groups[index].Width - singleWidth * (index == 1 ? 3d : 2d)) < 0.000001d,
                    "Adding adjacent formulas must preserve each display box's measurement.");
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
                .Concat(DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, default))
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
