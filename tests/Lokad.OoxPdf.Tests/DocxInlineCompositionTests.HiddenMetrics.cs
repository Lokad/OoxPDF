using System.Xml.Linq;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxHiddenTextMetricsShareBodyAndCellAscenderOwnership()
    {
        var control = ReadHiddenMetricSource("picture", OoxPdfDocxMarkupMode.Final, "direct", addRun: false, uniformSize: true).Paragraphs[0];
        var document = ReadHiddenMetricSource("picture", OoxPdfDocxMarkupMode.Final, "direct", addRun: true, uniformSize: true);
        var source = document.Paragraphs[0];
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        var plan = DocxFontPlan.Create(document, resolver, default);
        var measurer = DocxFontPlanTextMeasurer.CreateAsync(plan, resolver.Resolve(new FontRequest("DejaVu Sans", false, false)),
            default, resolver).GetAwaiter().GetResult();
        TestAssert.True(DocxLineMetrics.TableCellFontSizesUniform([control]) && DocxLineMetrics.TableCellFontSizesUniform([source]),
            "A hidden large run cannot revoke the visible cell's uniform-size metric contract.");
        foreach (bool maximum in new[] { true, false })
            TestAssert.Equal(DocxLineMetrics.ResolveHheaAscenderPoints(control, measurer, maximum),
                DocxLineMetrics.ResolveHheaAscenderPoints(source, measurer, maximum));
        TestAssert.Equal(DocxLineMetrics.ResolveBodyTierAMaxPoints(control, measurer), DocxLineMetrics.ResolveBodyTierAMaxPoints(source, measurer));
        TestAssert.Equal(DocxLineMetrics.ResolveTableCellFirstBaselinePoints([control], control, 11d, measurer),
            DocxLineMetrics.ResolveTableCellFirstBaselinePoints([source], source, 11d, measurer));
        TestAssert.Equal(DocxLineMetrics.ResolveTableCellFirstBaselineInset([control], measurer),
            DocxLineMetrics.ResolveTableCellFirstBaselineInset([source], measurer));
    }

    public static void DocxHiddenTextMetricsPreserveComposedPaginationAndPaint()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string content in new[] { "picture", "math", "mixed" })
        foreach (string owner in new[] { "direct", "character" })
        {
            var control = ReadHiddenMetricSource(content, mode, owner, addRun: false);
            var source = ReadHiddenMetricSource(content, mode, owner, addRun: true);
            var original = source.Paragraphs[0].Runs.ToArray();
            TestAssert.True(original.Single(run => run.Text == "Unprinted").EffectiveProperties.Hidden,
                "The parser retains the hidden source and its effective visibility.");
            var actual = RenderTextRowMetricDocument(source, mode);
            TestAssert.Equal(1, actual.Length);
            AssertMetricPaint(RenderTextRowMetricDocument(control, mode), actual);
            TestAssert.True(original.Zip(source.Paragraphs[0].Runs).All(pair => ReferenceEquals(pair.First, pair.Second)),
                "Metric filtering cannot replace source runs.");
        }
    }

    public static void DocxHiddenTextMetricsRespectInheritedAndExplicitVisibility()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string owner in new[] { "paragraph", "defaults" })
        {
            var control = ReadHiddenMetricSource("picture", mode, owner, addRun: false);
            var hidden = ReadHiddenMetricSource("picture", mode, owner, addRun: true);
            AssertMetricPaint(RenderTextRowMetricDocument(control, mode), RenderTextRowMetricDocument(hidden, mode));
        }
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var visible = ReadHiddenMetricSource("picture", mode, "direct", addRun: true, visible: true);
            var overridden = ReadHiddenMetricSource("picture", mode, "character", addRun: true, visible: true);
            var run = overridden.Paragraphs[0].Runs.Single(run => run.Text == "Unprinted");
            TestAssert.True(!run.EffectiveProperties.Hidden && run.HiddenValue == "false",
                "Explicit false overrides the hidden character style without losing source metadata.");
            var actual = RenderTextRowMetricDocument(overridden, mode);
            AssertMetricPaint(RenderTextRowMetricDocument(visible, mode), actual);
            string text = string.Concat(actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)).Select(show => show.Text));
            TestAssert.True(string.Concat(text.Where(character => !char.IsWhiteSpace(character))).Contains("Unprinted", StringComparison.Ordinal),
                "A visible override keeps the authored large text instead of treating it as hidden.");
        }
    }

    public static void DocxHiddenTextMetricsKeepParagraphMarksAndDeclaredSlots()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string slot in new[] { "exact", "atLeast" })
        {
            var control = ReadHiddenMetricSource("picture", mode, "direct", addRun: false, slot: slot);
            var hidden = ReadHiddenMetricSource("picture", mode, "direct", addRun: true, slot: slot);
            AssertMetricPaint(RenderTextRowMetricDocument(control, mode), RenderTextRowMetricDocument(hidden, mode));
        }
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            DocxDocument Source(int hiddenSize) => ReadTextRowMetrics("plain", "ordinary", mode, (xml, w, _) =>
            {
                var paragraph = xml.Descendants(w + "body").Single().Element(w + "p")!;
                paragraph.Elements(w + "r").Remove();
                paragraph.Element(w + "pPr")!.Add(new XElement(w + "rPr", new XElement(w + "sz", new XAttribute(w + "val", "52"))));
                paragraph.Add(new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "vanish"),
                    new XElement(w + "sz", new XAttribute(w + "val", hiddenSize * 2))), new XElement(w + "t", "Unprinted")));
                paragraph.AddAfterSelf(new XElement(w + "p", new XElement(w + "pPr", new XElement(w + "spacing",
                    new XAttribute(w + "before", "0"), new XAttribute(w + "after", "0"))),
                    new XElement(w + "r", new XElement(w + "t", "Afterword"))));
            });
            // Keep the same hidden owner and authored paragraph mark. A parsed
            // empty paragraph has a separate existing implicit-mark layout path.
            var ordinaryHidden = Source(11); var allHidden = Source(120);
            TestAssert.Equal(26d, allHidden.Paragraphs[0].ParagraphMarkFontSize!.Value);
            var control = RenderTextRowMetricDocument(ordinaryHidden, mode); var actual = RenderTextRowMetricDocument(allHidden, mode);
            TestAssert.Equal(control.Length, actual.Length);
            var expected = DocxCommentsTests.ReadEmbeddedGlyphTextShows(control[0]).Where(show => show.Text.Trim().Length != 0).ToArray();
            var observed = DocxCommentsTests.ReadEmbeddedGlyphTextShows(actual[0]).Where(show => show.Text.Trim().Length != 0).ToArray();
            TestAssert.Equal(string.Concat(expected.Select(show => show.Text)), string.Concat(observed.Select(show => show.Text)));
            TestAssert.True(expected.Zip(observed).All(pair => Math.Abs(pair.First.Y - pair.Second.Y) < .001d),
                "Changing only hidden font size retains the authored mark slot and following baseline.");
        }
    }

    private static void AssertMetricPaint(PdfPage[] expected, PdfPage[] actual)
    {
        TestAssert.Equal(expected.Length, actual.Length);
        for (int index = 0; index < expected.Length; index++)
        {
            // Retained hidden source can change subset CIDs without changing
            // painted glyphs. Check decoded physical glyphs and original font
            // programs, then keep every other emitted operator exact.
            var before = DocxCommentsTests.ReadEmbeddedGlyphTextShows(expected[index], true, honorActualText: false).ToArray();
            var after = DocxCommentsTests.ReadEmbeddedGlyphTextShows(actual[index], true, honorActualText: false).ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int glyph = 0; glyph < before.Length; glyph++)
            {
                TestAssert.Equal(before[glyph].Text, after[glyph].Text);
                TestAssert.Equal(before[glyph].Size, after[glyph].Size);
                TestAssert.Equal(before[glyph].X, after[glyph].X);
                TestAssert.Equal(before[glyph].Y, after[glyph].Y);
            }
            TestAssert.Equal(expected[index].Fonts.Count, actual[index].Fonts.Count);
            foreach (var resource in expected[index].Fonts)
                TestAssert.True(resource.Font.Font.Bytes.Span.SequenceEqual(actual[index].Fonts.Single(
                    font => font.ResourceName == resource.ResourceName).Font.Font.Bytes.Span), "Paint uses the same source font program.");
            TestAssert.Equal(Regex.Replace(expected[index].Content, @"<[0-9A-Fa-f]+>", "<>"),
                Regex.Replace(actual[index].Content, @"<[0-9A-Fa-f]+>", "<>"));
            TestAssert.Equal(expected[index].Images.Count, actual[index].Images.Count);
            TestAssert.True(expected[index].Annotations.SequenceEqual(actual[index].Annotations),
                "Hidden metric owners cannot move navigation.");
        }
    }

    private static DocxDocument ReadHiddenMetricSource(string content, OoxPdfDocxMarkupMode mode,
        string owner, bool addRun, bool visible = false, string? slot = null, bool uniformSize = false)
    {
        return ReadTextRowMetrics(content, "ordinary", mode, (xml, w, _) =>
        {
            var paragraph = xml.Descendants(w + "body").Single().Element(w + "p")!;
            if (uniformSize)
                foreach (var size in paragraph.Descendants(w + "sz")) size.SetAttributeValue(w + "val", "22");
            if (slot is not null)
            {
                var spacing = paragraph.Element(w + "pPr")!.Element(w + "spacing")!;
                spacing.SetAttributeValue(w + "lineRule", slot); spacing.SetAttributeValue(w + "line", "400");
            }
            if (owner is "paragraph" or "defaults")
                foreach (var run in paragraph.Elements(w + "r"))
                {
                    var properties = run.Element(w + "rPr");
                    if (properties is null) { properties = new XElement(w + "rPr"); run.AddFirst(properties); }
                    properties.Add(new XElement(w + "vanish", new XAttribute(w + "val", "false")));
                }
            if (owner == "paragraph") paragraph.Element(w + "pPr")!.Add(new XElement(w + "pStyle", new XAttribute(w + "val", "HiddenMetric")));
            if (!addRun) return;
            var hiddenProperties = new XElement(w + "rPr", new XElement(w + "sz", new XAttribute(w + "val", "240")));
            if (owner == "character") hiddenProperties.Add(new XElement(w + "rStyle", new XAttribute(w + "val", "HiddenMetric")));
            if (owner == "direct" || visible) hiddenProperties.Add(new XElement(w + "vanish", new XAttribute(w + "val", visible ? "false" : "true")));
            paragraph.Element(w + "pPr")!.AddAfterSelf(new XElement(w + "r", hiddenProperties, new XElement(w + "t", "Unprinted")));
        }, packageChange: zip =>
        {
            if (owner == "direct") return;
            var entry = zip.GetEntry("word/styles.xml")!;
            XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            if (owner == "defaults") xml.Root!.Element(w + "docDefaults")!.Element(w + "rPrDefault")!.Element(w + "rPr")!.Add(new XElement(w + "vanish"));
            else
            {
                string kind = owner == "paragraph" ? "paragraph" : "character";
                xml.Root!.Add(new XElement(w + "style", new XAttribute(w + "type", kind), new XAttribute(w + "styleId", "HiddenMetricBase"),
                    new XElement(w + "rPr", new XElement(w + "vanish"))),
                    new XElement(w + "style", new XAttribute(w + "type", kind), new XAttribute(w + "styleId", "HiddenMetric"),
                    new XElement(w + "basedOn", new XAttribute(w + "val", "HiddenMetricBase"))));
            }
            entry.Delete(); using var output = zip.CreateEntry("word/styles.xml").Open(); xml.Save(output);
        });
    }
}
