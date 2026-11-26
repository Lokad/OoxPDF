using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionJustifiedMathMatchesIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent justification anchors require installed Windows fonts.");
        CheckJustifiedMath(null);
    }

    internal static void CheckJustifiedMath(IFontResolver? resolver) =>
        CheckComposedMathReport("justified-math-office.json", resolver, exactLine: false);

    public static void DocxInlineCompositionJustificationOwnsWhitespaceAndContinuation()
    {
        var document = ReadMutation("Composition/report", OoxPdfDocxMarkupMode.Final, false, false, (xml, w, m) =>
        {
            var equations = xml.Descendants(m + "oMath").Take(2).Select(e => new System.Xml.Linq.XElement(e)).ToArray();
            var drawing = new System.Xml.Linq.XElement(xml.Descendants(w + "drawing").Single().Parent!);
            System.Xml.Linq.XElement Text(string text) => new(w + "r", new System.Xml.Linq.XElement(w + "t", text));
            xml.Root!.Element(w + "body")!.ReplaceNodes(new System.Xml.Linq.XElement(w + "p",
                Text("Alpha beta "), equations[0], Text(" gamma delta "), drawing, Text(" epsilon "), equations[1],
                Text(" zeta eta theta iota kappa lambda mu.")));
        });
        var source = document.Paragraphs.Single();
        var prepared = source.Equations.ToDictionary(e => e,
            e => new DocxPreparedMathParagraph(e, new(18d, 10d, 4d, []), 11d, 16d) { UsesDisplayStyle = false });
        foreach (double scale in new[] { 1d, .75d })
        {
            var paragraph = source with { Alignment = DocxTextAlignment.Justified };
            TestAssert.Equal(2, DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(paragraph, prepared, new TrackingMeasurer(),
                12d, 9d, 16d, scale, default, out var measured), "Mixed justification shares the ordinary atom measurements.");
            var atoms = TestAssert.NotNull(measured);
            TestAssert.True(DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 150d, 170d, default, out var wrapped), "The composed paragraph wraps.");
            var lines = TestAssert.NotNull(wrapped);
            TestAssert.True(lines.Count > 2, "Exercise continuation and the final paragraph line.");
            TestAssert.Equal(2, lines.SelectMany(line => line.Fragments).Count(f => f.Math is not null));
            TestAssert.Equal(1, lines.SelectMany(line => line.Fragments).Count(f => f.Source is DocxInlinePicture));
            TestAssert.True(lines[^1].Advance < 170d, "The final paragraph line retains its natural width.");
            var cursor = new DocxInlineSourceCursor();
            for (int index = 0; index < lines.Count; index++)
            {
                var expected = lines[index];
                TestAssert.True(DocxLayoutEngine.TryGetNextOrderedInlineLine(atoms, cursor,
                    index == 0 ? 150d : 170d, 170d, default, out var next), "A resumed container obtains the same source line.");
                var actual = TestAssert.NotNull(next);
                TestAssert.Equal(expected.Advance, actual.Advance);
                TestAssert.Equal(expected.EndCursor, actual.EndCursor);
                TestAssert.True(expected.Fragments.Select(f => (f.Advance, f.LeadingAdvance, f.Source, f.Math))
                    .SequenceEqual(actual.Fragments.Select(f => (f.Advance, f.LeadingAdvance, f.Source, f.Math))),
                    "Resuming placement retains every source owner, stretched gap and intrinsic box.");
                if (index + 1 < lines.Count && expected.Fragments.SelectMany(f => f.Text).Any(s => s.Text.Contains(' ')))
                    TestAssert.Equal(index == 0 ? 150d : 170d, expected.Advance);
                foreach (var fragment in actual.Fragments.Where(f => f.Math is not null || f.Source is DocxInlinePicture))
                {
                    var original = atoms.Fragments.Single(f => ReferenceEquals(f.Source, fragment.Source));
                    TestAssert.Equal(original.Advance, fragment.Advance);
                    TestAssert.Equal(original.BaselineInset, fragment.BaselineInset);
                    TestAssert.Equal(original.OccupiedHeight, fragment.OccupiedHeight);
                }
                cursor = actual.EndCursor;
            }
            TestAssert.True(cursor.IsComplete(atoms), "The closing paragraph mark is consumed once.");
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxLayoutEngine.TryWrapOrderedInlineAtoms(atoms, 150d, 170d, cancelled.Token, out _));
        }
        var legacy = ReadMutation("docx-omml-isolated-inline-body-fraction-left", OoxPdfDocxMarkupMode.Final, false, false)
            .Paragraphs.Single(p => p.Equations.Count != 0) with { Alignment = DocxTextAlignment.Justified };
        var equation = legacy.Equations.Single();
        var legacyPrepared = new Dictionary<DocxEquation, DocxPreparedMathParagraph> { [equation] =
            new(equation, new(18d, 10d, 4d, []), 11d, 16d) { UsesDisplayStyle = false } };
        TestAssert.True(!DocxMathParagraphPolicy.NeedsOrderedInlineLayout(legacy, legacyPrepared),
            "Qualified legacy single-formula justification keeps its existing placement owner.");
    }
}
