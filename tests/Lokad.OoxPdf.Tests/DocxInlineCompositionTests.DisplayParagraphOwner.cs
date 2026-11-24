using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    internal static void CheckDisplayParagraphOwner()
    {
        foreach ((string container, string? first, bool implicitPrefix) in new (string, string?, bool)[]
        {
            ("body", "left", false), ("body", "right", false), ("body", null, false),
            ("cell", "left", false), ("cell", "right", false), ("cell", null, false),
            ("footer", "left", false), ("footer", "right", false), ("footer", null, false),
            ("body", "right", true)
        })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string fixture = container == "footer" ? "math-explicit-display-right-footer" :
                "math-explicit-display-left-" + container;
            string part = container == "footer" ? "word/math-tabs-footer.xml" : "word/document.xml";
            string later = first == "right" ? "left" : "right";
            DocxDocument Read(bool different) => ReadMutation("CompositionProbes/" + fixture, mode, false, false,
                (xml, w, m) =>
                {
                    var group = xml.Descendants(m + "oMathPara").Single();
                    var equations = group.Elements(m + "oMath").ToArray();
                    TestAssert.Equal(2, equations.Length);
                    group.Element(m + "oMathParaPr")?.Remove();
                    string? initial = implicitPrefix && !different ? null : first;
                    if (initial is not null)
                        group.AddFirst(new XElement(m + "oMathParaPr", new XElement(m + "jc", new XAttribute(m + "val", initial))));
                    if (implicitPrefix)
                        group.AddBeforeSelf(new XElement(equations[0]), new XElement(w + "r", new XElement(w + "br")));
                    equations[1].Remove();
                    var next = new XElement(m + "oMathPara");
                    string? value = different ? later : implicitPrefix ? null : first;
                    if (value is not null)
                        next.Add(new XElement(m + "oMathParaPr", new XElement(m + "jc", new XAttribute(m + "val", value))));
                    next.Add(equations[1]);
                    group.AddAfterSelf(new XElement(w + "r", new XElement(w + "br")), next);
                }, partOverride: part);

            var baseline = Read(false);
            var document = Read(true);
            var owner = DocxBlockTraversal.EnumerateBodyParagraphs(document.EquationParts
                .Single(source => source.PartName == "/" + part).BodyElements).Single(p => p.Equations.Count == (implicitPrefix ? 3 : 2));
            int displayStart = implicitPrefix ? 1 : 0;
            TestAssert.Equal(first, owner.Equations[displayStart].AlignmentValue);
            TestAssert.Equal(later, owner.Equations[displayStart + 1].AlignmentValue);
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var expected = renderer.RenderPagesAsync(baseline, null, default).ToBlockingEnumerable().ToArray();
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(expected.Length, actual.Length);
            for (int page = 0; page < expected.Length; page++)
            {
                var before = DocxCommentsTests.ReadEmbeddedGlyphTextShows(expected[page], usePageCoordinates: true)
                    .Select(show => (show.Text, show.X, show.Y, show.Size)).ToArray();
                var after = DocxCommentsTests.ReadEmbeddedGlyphTextShows(actual[page], usePageCoordinates: true)
                    .Select(show => (show.Text, show.X, show.Y, show.Size)).ToArray();
                TestAssert.True(before.Length != 0 && before.SequenceEqual(after),
                    "Later display justification cannot remove or reposition source glyphs: " + container + "/" + first + "/" + implicitPrefix + "/" + mode);
            }
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                "Differently justified display groups retain supported formulas.");
            TestAssert.Equal(first, owner.Equations[displayStart].AlignmentValue);
            TestAssert.Equal(later, owner.Equations[displayStart + 1].AlignmentValue);
        }
    }

    internal static void CheckDisplayMathRunBreak()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument Read(bool insideEquation, bool mathRun, string? type = null, int breakCount = 1, bool proofing = false,
                bool? bookmarkStartsAfterBreaks = null, string? paginationHint = null, bool packBreaks = false) => ReadMutation(
                "CompositionProbes/math-explicit-display-same-group", mode, false, false, (xml, w, m) =>
                {
                    var group = xml.Descendants(m + "oMathPara").Single();
                    var equations = group.Elements(m + "oMath").ToArray();
                    XElement? packedRun = null;
                    for (int index = 0; index < breakCount; index++)
                    {
                        string? breakType = index == 0 ? type : "textWrapping";
                        var lineBreak = new XElement(w + "br", breakType is null ? null : new XAttribute(w + "type", breakType));
                        if (packedRun is not null)
                        {
                            packedRun.Add(lineBreak);
                            continue;
                        }
                        var run = new XElement((mathRun ? m : w) + "r", lineBreak);
                        if (packBreaks) packedRun = run;
                        if (index == 0 && paginationHint is ("inside" or "malformed"))
                            run.Add(new XElement(w + "lastRenderedPageBreak",
                                paginationHint == "malformed" ? new XAttribute(w + "future", "1") : null));
                        if (insideEquation) equations[0].Add(run);
                        else equations[1].AddBeforeSelf(run);
                        if (index == 0 && paginationHint == "between")
                            run.AddAfterSelf(new XElement(m + "r", new XElement(w + "lastRenderedPageBreak")));
                    }
                    if (proofing)
                    {
                        equations[0].AddFirst(new XElement(w + "proofErr", new XAttribute(w + "type", "spellStart")));
                        equations[0].Add(new XElement(w + "proofErr", new XAttribute(w + "type", "spellEnd")));
                    }
                    if (bookmarkStartsAfterBreaks is { } startsAfterBreaks)
                    {
                        var start = new XElement(w + "bookmarkStart", new XAttribute(w + "id", 55),
                            new XAttribute(w + "name", "MathBreakAnchor"));
                        var end = new XElement(w + "bookmarkEnd", new XAttribute(w + "id", 55));
                        XElement afterBreaks = startsAfterBreaks ? start : end;
                        if (insideEquation) equations[0].Add(afterBreaks);
                        else equations[1].AddBeforeSelf(afterBreaks);
                        if (startsAfterBreaks) equations[1].AddAfterSelf(end);
                        else equations[0].AddFirst(start);

                        group.Parent!.AddAfterSelf(new XElement(w + "p",
                            new XElement(w + "hyperlink", new XAttribute(w + "anchor", "MathBreakAnchor"),
                                new XElement(w + "r", new XElement(w + "t", "Go to formula")))));
                    }
                });

            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                    OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var between = renderer.RenderPagesAsync(Read(insideEquation: false, mathRun: false), null, default)
                .ToBlockingEnumerable().Single();
            var expected = MathShows(between);
            TestAssert.Equal("α+β=γδ+ε=ζ", string.Concat(expected.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(2, expected.Select(show => show.Y).Distinct().Count());
            foreach ((bool insideEquation, string? type) in new[] { (true, (string?)null), (true, "textWrapping"), (false, null) })
            {
                CheckMathBreaks(Read(insideEquation, mathRun: true, type), expected);
            }

            var betweenTwo = renderer.RenderPagesAsync(Read(false, false, breakCount: 2), null, default)
                .ToBlockingEnumerable().Single();
            var expectedTwo = MathShows(betweenTwo);
            TestAssert.Equal("α+β=γδ+ε=ζ", string.Concat(expectedTwo.Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.True(expectedTwo[0].Y == expected[0].Y &&
                Math.Abs(expectedTwo[^1].Y - expectedTwo[0].Y) > Math.Abs(expected[^1].Y - expected[0].Y),
                "The second authored break must retain an additional line of space.");
            CheckMathBreaks(Read(insideEquation: true, mathRun: true, breakCount: 2), expectedTwo);
            CheckMathBreaks(Read(insideEquation: true, mathRun: true, breakCount: 2, proofing: true), expectedTwo);
            foreach (bool startsAfterBreaks in new[] { false, true })
            {
                var control = CheckMathBreaks(Read(false, false, breakCount: 2, bookmarkStartsAfterBreaks: startsAfterBreaks), expectedTwo);
                var nested = CheckMathBreaks(Read(true, true, breakCount: 2, bookmarkStartsAfterBreaks: startsAfterBreaks), expectedTwo);
                var target = control.Annotations.Single(link => link.IsDestination);
                TestAssert.Equal(target, nested.Annotations.Single(link => link.IsDestination));
                if (startsAfterBreaks)
                {
                    foreach (string hint in new[] { "inside", "between" })
                    {
                        var withHint = CheckMathBreaks(Read(true, true, breakCount: 2,
                            bookmarkStartsAfterBreaks: true, paginationHint: hint), expectedTwo);
                        TestAssert.Equal(target, withHint.Annotations.Single(link => link.IsDestination));
                    }
                    foreach (bool insideEquation in new[] { false, true })
                    {
                        var packed = CheckMathBreaks(Read(insideEquation, mathRun: true, breakCount: 2,
                            bookmarkStartsAfterBreaks: true, paginationHint: insideEquation ? "inside" : null,
                            packBreaks: true), expectedTwo);
                        TestAssert.Equal(target, packed.Annotations.Single(link => link.IsDestination));
                    }
                }
            }

            foreach (var document in new[] { Read(true, true, type: "page"), Read(true, true, paginationHint: "malformed"),
                Read(true, true, type: "page", breakCount: 2, packBreaks: true) })
            {
                var unsupported = new List<OoxPdfDiagnostic>();
                var page = renderer.RenderPagesAsync(document, unsupported.Add, default).ToBlockingEnumerable().Single();
                TestAssert.Equal("δ+ε=ζ", string.Concat(MathShows(page).Select(show => show.Text)).Normalize(NormalizationForm.FormKC));
                TestAssert.True(unsupported.Any(d => d.Id == "DOCX_UNSUPPORTED_EQUATION"),
                    "An authored page break or malformed hint keeps its diagnosed fallback.");
            }

            PdfPage CheckMathBreaks(DocxDocument document, (string Text, double X, double Y, double Size)[] control)
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var page = renderer.RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
                TestAssert.True(control.SequenceEqual(MathShows(page)),
                    "Math-run line breaks must preserve both formulas at their ordinary line-break positions.");
                TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"),
                    "The display breaks must preserve their equation.");
                return page;
            }
        }
    }
}
