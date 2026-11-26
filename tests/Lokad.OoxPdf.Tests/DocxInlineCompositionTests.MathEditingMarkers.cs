using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxDisabledMathEffectsKeepControlAndArgumentPaint()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string owner in new[] { "cell", "header", "argument" })
        {
            var renderer = PictureRangeRenderer(mode);
            var baseline = Render(null, null);
            int glyphs = CountMathGlyphs(baseline.Page);
            TestAssert.True(glyphs > 0, "The control must paint equations.");
            TestAssert.True(owner == "header" || baseline.Page.Annotations.Any(annotation => annotation.Destination.HasValue),
                "The body fixture must emit formula destinations.");

            (string Property, string Value)[] effects = owner switch
            {
                "cell" => [("shadow", "off"), ("outline", "false"), ("emboss", "0"), ("imprint", "false"),
                    ("effect", "none"), ("em", "none"), ("bdr", "none"), ("bdr", "nil")],
                "header" => [("shadow", "off"), ("effect", "none"), ("bdr", "none")],
                _ => [("shadow", "off"), ("effect", "none"), ("u", "none"), ("highlight", "none")]
            };
            foreach (var (property, value) in effects)
            {
                var observed = Render(property, value);
                TestAssert.Equal(baseline.Page.Content, observed.Page.Content);
                TestAssert.True(baseline.Page.Annotations.SequenceEqual(observed.Page.Annotations),
                    "A disabled visual effect preserves the control's destinations.");
                TestAssert.True(baseline.Diagnostics.SequenceEqual(observed.Diagnostics),
                    "A disabled visual effect preserves the control's diagnostic outcome.");
            }

            (string Property, string? Value)[] fallbacks = owner switch
            {
                "cell" => [("shadow", null), ("shadow", "unrecognized"), ("effect", "lights"), ("em", "dot"), ("bdr", "single")],
                "header" => [("shadow", null)],
                _ => [("shadow", null), ("u", "single"), ("highlight", "yellow")]
            };
            foreach (var (property, value) in fallbacks)
            {
                var fallback = Render(property, value);
                TestAssert.True(fallback.Diagnostics.Contains("DOCX_UNSUPPORTED_EQUATION") && CountMathGlyphs(fallback.Page) < glyphs,
                    $"The unsupported {property} declaration retains the equation fallback.");
            }

            (PdfPage Page, List<string> Diagnostics) Render(string? property, string? value)
            {
                var diagnostics = new List<string>();
                void Record(Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic item) => diagnostics.Add(item.Id);
                const string navigation = "CompositionProbes/inline-math-navigation-cell";
                const string headerTable = "CompositionProbes/table-math-revision-header-row-ins";
                var source = ReadMutation(owner == "header" ? headerTable : navigation, mode, false, false,
                    (xml, w, m) =>
                    {
                        if (property is null) return;
                        var fraction = xml.Descendants(w + "tbl").First().Descendants(m + "f").First();
                        var holder = owner == "argument" ? fraction.Element(m + "num")! :
                            fraction.Element(m + "fPr") ?? Add(fraction, new(m + "fPr"), first: true);
                        var control = holder.Element(m + "ctrlPr") ?? Add(holder, new(m + "ctrlPr"));
                        var run = control.Element(w + "rPr") ?? Add(control, new(w + "rPr"));
                        var flag = new XElement(w + property, value is null ? null : new XAttribute(w + "val", value));
                        if (property is "outline" or "shadow" or "emboss" or "imprint" && run.Element(w + "sz") is { } size)
                            size.AddBeforeSelf(flag);
                        else run.Add(flag);
                    }, partOverride: owner == "header" ? "word/header1.xml" : "word/document.xml", diagnosticSink: Record);
                return (renderer.RenderPagesAsync(source, Record, default).ToBlockingEnumerable().Single(), diagnostics);
            }
        }

        static XElement Add(XElement parent, XElement child, bool first = false)
        {
            if (first) parent.AddFirst(child);
            else parent.Add(child);
            return child;
        }

        static int CountMathGlyphs(PdfPage page) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
            usePageCoordinates: true, honorActualText: false, fontResourceFilter: font => font.StartsWith("Eq", StringComparison.Ordinal))
            .Sum(show => show.Text.Length);
    }

    public static void DocxMathEditingMarkersKeepPortablePaintAcrossContainers()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var source in new[]
        {
            (Fixture: "docx-omml-fraction-inline", Part: "word/document.xml", Page: 0,
                Numerator: "a+b", Denominator: "c", Companion: "", Count: 1),
            (Fixture: "CompositionProbes/math-picture-first-table-single-cell", Part: "word/document.xml", Page: 0,
                Numerator: "rate", Denominator: "time", Companion: "", Count: 2),
            (Fixture: "CompositionProbes/equation-revision-header-ins", Part: "word/header1.xml", Page: 0,
                Numerator: "x+2", Denominator: "y", Companion: "a+b=c", Count: 2),
            (Fixture: "CompositionProbes/mixed-note-endnote-table", Part: "word/endnotes.xml", Page: 1,
                Numerator: "rate", Denominator: "time", Companion: "", Count: 2)
        })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode,
                mode == OoxPdfDocxMarkupMode.Final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
                    : OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var direct = DocxTests.ReadDocx(TestFixtures.PublicCasePath(source.Fixture + ".docx"), mode);
            var marked = ReadMutation(source.Fixture, mode, false, false, (xml, w, m) =>
            {
                var arguments = xml.Descendants(m + "num").ToArray();
                TestAssert.Equal(source.Count, arguments.Length);
                int markerId = 12;
                foreach (var argument in arguments)
                {
                    argument.AddFirst(new XElement(w + "proofErr", new XAttribute(w + "type", "spellStart")),
                        new XElement(w + "permStart", new XAttribute(w + "id", markerId),
                            new XAttribute(w + "edGrp", "everyone")));
                    argument.Add(new XElement(w + "permEnd", new XAttribute(w + "id", markerId)),
                        new XElement(w + "proofErr", new XAttribute(w + "type", "spellEnd")));
                    markerId++;
                }
            }, partOverride: source.Part);
            var before = renderer.RenderPagesAsync(direct, null, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(marked, null, default).ToBlockingEnumerable().ToArray();
            CheckEditingMarkerFractionPaint(before[source.Page], source.Numerator, source.Denominator, source.Companion, source.Count);
            CheckEditingMarkerFractionPaint(after[source.Page], source.Numerator, source.Denominator, source.Companion, source.Count);
            TestAssert.Equal(before.Length, after.Length);
            for (int page = 0; page < before.Length; page++)
            {
                TestAssert.Equal(before[page].Content, after[page].Content);
                TestAssert.Equal(before[page].Fonts.Count, after[page].Fonts.Count);
                TestAssert.Equal(before[page].Images.Count, after[page].Images.Count);
                TestAssert.Equal(before[page].Annotations.Count, after[page].Annotations.Count);
            }
        }
    }

    private static void CheckEditingMarkerFractionPaint(PdfPage page, string numerator, string denominator, string companion, int count)
    {
        string equation = numerator + denominator + companion;
        string expected = string.Concat(Enumerable.Repeat(equation, count));
        var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(page, expected, mathOnly: true);
        TestAssert.Equal(expected.Length, glyphs.Count);
        var rectangles = ReadVisibleFilledRectangles(page).ToArray();
        for (int i = 0; i < count; i++)
        {
            int start = i * equation.Length;
            int denominatorStart = start + numerator.Length;
            ReadVisibleFractionRule(rectangles, glyphs[start].Origin, glyphs[denominatorStart - 1].Origin,
                glyphs[denominatorStart].Origin, glyphs[denominatorStart + denominator.Length - 1].Origin);
        }
    }
}
