using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineHiddenBreaksKeepFittingWordsAndComposedObjects()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string container in new[] { "body", "cell", "header", "footer" })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = Render(null);
            foreach (char marker in new[] { '\u00ad', '\u200b' })
            {
                var actual = Render(marker);
                TestAssert.Equal(control.Length, actual.Length);
                TestAssert.Equal(Text(control), Text(actual));
                TestAssert.True(!Text(actual).Contains(marker), "Unused break characters must not acquire painted glyphs beside equations.");
                TestAssert.True(MathShows(control).SequenceEqual(MathShows(actual)),
                    "An unused break cannot move either supported equation.");
                TestAssert.True(Pictures(control).SequenceEqual(Pictures(actual)),
                    "The following picture keeps its painted frame.");
            }
            PdfPage[] Render(char? marker) => renderer.RenderPagesAsync(ReadInlineBreakWord(container, mode, marker), null, default)
                .ToBlockingEnumerable().ToArray();
        }
        static string Text(PdfPage[] pages) => string.Concat(pages.SelectMany(page =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false)).Select(show => show.Text));
        static IEnumerable<(string Text, double X, double Y)> MathShows(PdfPage[] pages) => pages.SelectMany(page =>
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false,
                fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal)))
            .Select(show => (show.Text, show.X, show.Y));
        static IEnumerable<string> Pictures(PdfPage[] pages) => pages.SelectMany(page => Regex.Matches(page.Content,
            @"[-\d.]+ 0 0 [-\d.]+ [-\d.]+ [-\d.]+ cm\s+/Im\d+ Do").Select(match => match.Value));
    }

    public static void DocxInlineHiddenBreaksRetainWrappingAndSourceProgress()
    {
        foreach (char marker in new[] { '\u00ad', '\u200b' })
        {
            var (paragraph, equations) = InlineBreakMeasurement(["ABC" + marker + "DEFG"]);
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations,
                new DocxTests.FamilyWidthTextMeasurer(), 10d, 8d, 12d, 1d, default, out var measured), "Measure supported break opportunities once.");
            var atoms = TestAssert.NotNull(measured);
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 25d, 25d, default, out var wrapped), "The existing preferred break remains available.");
            var lines = TestAssert.NotNull(wrapped);
            TestAssert.Equal(3, lines.Count);
            TestAssert.Equal(marker == '\u00ad' ? "ABC\u00ad" : "ABC", Paint(lines[0]));
            TestAssert.Equal("DEFG", Paint(lines[1]));
            var cursor = default(DocxInlineSourceCursor);
            foreach (var expected in lines)
            {
                TestAssert.True(DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, cursor, 25d, 25d, default, out var next), "Resume source after the admitted row.");
                var actual = TestAssert.NotNull(next);
                TestAssert.Equal(Paint(expected), Paint(actual));
                TestAssert.Equal(expected.EndCursor, actual.EndCursor);
                TestAssert.True(actual.EndCursor != cursor, "Every resumed row consumes source, including nonprinting boundaries.");
                cursor = actual.EndCursor;
            }
            TestAssert.True(cursor.IsComplete(atoms), "All source text and the equation are consumed exactly once.");
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms,
                default, 25d, 25d, cancelled.Token, out _));
        }
        static string Paint(DocxMeasuredInlineLine line) => string.Concat(line.Fragments.SelectMany(fragment => fragment.TextSegments).Select(segment => segment.Text));
    }

    public static void DocxInlineNonprintingRunsKeepVisibleBoundaryKerning()
    {
        foreach (char marker in new[] { '\u00ad', '\u200b' })
        {
            var (paragraph, equations) = InlineBreakMeasurement(["A", marker.ToString(), "V"]);
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations, new BreakPairMeasurer(),
                10d, 8d, 12d, 1d, default, out var measured), "A nonprinting source run retains its owner.");
            var atoms = TestAssert.NotNull(measured);
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 27d, 27d, default, out var wrapped), "The visible pair and equation fit exactly.");
            var line = TestAssert.NotNull(wrapped).Single();
            TestAssert.Equal(27d, line.Advance);
            TestAssert.Equal(4, line.Fragments.Count);
            TestAssert.True(line.Fragments[1].Source is DocxInlineText && line.Fragments[1].TextSegments.Count == 0,
                "The unused marker keeps its source slot without a painted segment.");
        }
    }

    public static void DocxInlineLeadingHiddenBreaksKeepEveryWrappedSourceCharacter()
    {
        const string word = "ABCDEFGHIJKLMN";
        foreach (string[] values in new[] { new[] { "\u200b" + word }, new[] { "\u200b", word } })
        {
            var (paragraph, equations) = InlineBreakMeasurement(values);
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, equations,
                new DocxTests.FamilyWidthTextMeasurer(), 10d, 8d, 12d, 1d, default, out var measured), "Measure the original source once.");
            var atoms = TestAssert.NotNull(measured);
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 25d, 25d, default, out var wrapped), "The complete wrap retains every letter.");
            var lines = TestAssert.NotNull(wrapped);
            var cursor = default(DocxInlineSourceCursor);
            var actualText = new System.Text.StringBuilder();
            foreach (var expected in lines)
            {
                TestAssert.True(DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, cursor, 25d, 25d, default, out var next),
                    "An empty painted row cannot consume the rest of its source word.");
                var actual = TestAssert.NotNull(next);
                TestAssert.Equal(expected.EndCursor, actual.EndCursor);
                actualText.Append(string.Concat(actual.Fragments.SelectMany(fragment => fragment.TextSegments).Select(segment => segment.Text)));
                TestAssert.True(actual.EndCursor != cursor, "Source advances even when the row has no painted spans.");
                cursor = actual.EndCursor;
            }
            TestAssert.Equal(word, actualText.ToString());
            TestAssert.True(cursor.IsComplete(atoms), "All text and the following equation are consumed once.");
        }
    }

    private static DocxDocument ReadInlineBreakWord(string container, OoxPdfDocxMarkupMode mode, char? marker) =>
        ReadMutation("CompositionProbes/inline-math-navigation" + (container == "body" ? "" : "-" + container), mode, false, false,
            (xml, w, m) =>
            {
                var paragraph = xml.Descendants(w + "p").Single(p => p.Descendants(m + "oMath").Count() == 2 && p.Descendants(w + "drawing").Count() == 1);
                var first = paragraph.Elements().First(element => element.Name != w + "pPr");
                first.AddBeforeSelf(new XElement(w + "r", new XElement(w + "t", new XAttribute(XNamespace.Xml + "space", "preserve"),
                    marker is null ? "Lemonade " : "Lem" + marker + "onade ")));
            }, partOverride: container is "header" or "footer" ? "word/" + container + "1.xml" : null);

    private static (DocxParagraph Paragraph, Dictionary<DocxEquation, DocxPreparedMathParagraph> Equations)
        InlineBreakMeasurement(string[] values)
    {
        var source = ReadMutation("CompositionProbes/math-minimum-line-72", OoxPdfDocxMarkupMode.Final, false, false)
            .Paragraphs.First(paragraph => paragraph.Equations.Count != 0);
        var runs = values.Select((value, index) => new DocxTextRun(value, 10d, null, false, false, false, null, null)
            { SourceRunIndex = index }).ToArray();
        var equation = source.Equations.Single();
        var paragraph = source with { Runs = runs, Images = [], LineSpacingPoints = null, LineSpacingFactor = 1d,
            Spacing = DocxParagraphSpacing.Empty, InlineContent = runs.Select(run => (DocxInlineItem)new DocxInlineText(run, 0, run.Text.Length))
                .Append(new DocxInlineEquation(equation)).ToArray() };
        return (paragraph, new(ReferenceEqualityComparer.Instance) { [equation] = new(equation, new(18d, 8d, 2d, []), 8d, 12d) });
    }

    private sealed class BreakPairMeasurer : IDocxTextMeasurer
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize) => text.Length * 5d - (text.Contains("AV", StringComparison.Ordinal) ? 1d : 0d);
    }
}
