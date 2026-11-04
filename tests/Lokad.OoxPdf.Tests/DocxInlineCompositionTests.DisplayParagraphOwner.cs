using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

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
            string?[] properties = owner.Equations.Select(equation => equation.PropertiesXml).ToArray();
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
            TestAssert.True(properties.SequenceEqual(owner.Equations.Select(equation => equation.PropertiesXml)) &&
                owner.Equations[displayStart].AlignmentValue == first && owner.Equations[displayStart + 1].AlignmentValue == later,
                "Effective paragraph ownership preserves authored group properties.");
        }
    }
}
