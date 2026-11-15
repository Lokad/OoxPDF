using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxListLabelCasingSharesMeasuredAndPaintedTextAcrossContainers()
    {
        foreach (var (container, suffix) in new[] { ("body", "space"), ("cell", "nothing"), ("header", "tab"), ("footer", "space") })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadListCasing(container, mode, true, false, false, suffix, upperPrefix: true),
                null, default).ToBlockingEnumerable().ToArray();
            var source = ReadListCasing(container, mode, true, false, false, suffix);
            var paragraphs = NumberedCounterParagraphs(source);
            TestAssert.Equal(3, paragraphs.Length);
            var rawLabels = paragraphs.Select(p => p.ListLabel!.DisplayText).ToArray();
            TestAssert.True(rawLabels.All(text => text.Contains("step ", StringComparison.Ordinal)),
                "Formatted counters retain their source spelling independently of display casing.");
            var diagnostics = new List<OoxPdfDiagnostic>();
            var actual = renderer.RenderPagesAsync(source, diagnostics.Add, default).ToBlockingEnumerable().ToArray();
            AssertMetricPaint(control, actual);
            var labels = actual.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false))
                .Where(show => show.Text.StartsWith("STEP ", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(paragraphs.Sum(p => p.ListLabel!.NumberingChange is null ? 1 : 2), labels.Length);
            TestAssert.True(labels.All(show => show.Text == show.Text.ToUpperInvariant()), "Each current and previous label paints resolved glyphs.");
            TestAssert.True(actual.Any(page => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal))),
                "The composed source retains its supported mathematics.");
            TestAssert.True(diagnostics.All(d => d.Id != "DOCX_UNSUPPORTED_EQUATION"), "Label casing cannot exclude neighboring equations.");
            TestAssert.Equal(string.Join('|', rawLabels), string.Join('|', paragraphs.Select(p => p.ListLabel!.DisplayText)));
        }
    }

    public static void DocxListLabelCasingBelongsToLevelAndParagraphMark()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (bool? level in new bool?[] { null, false, true })
        foreach (bool? mark in new bool?[] { null, false, true })
        foreach (bool body in new[] { false, true })
        {
            var source = ReadListCasing("body", mode, level, mark, body, "space");
            foreach (var paragraph in NumberedCounterParagraphs(source))
            {
                var label = paragraph.ListLabel!;
                bool caps = level ?? mark ?? false;
                var run = DocxLayoutEngine.CreateListLabelRun(label, paragraph.Runs.First(), 11d);
                TestAssert.Equal(caps, run.AllCaps);
                TestAssert.Equal(caps ? label.Text.ToUpperInvariant() : label.Text, run.Text);
                string previous = label.NumberingChange?.PreviousText ?? "step 0.";
                var previousRun = DocxLayoutEngine.CreatePreviousListLabelRun(run, previous);
                TestAssert.Equal(caps ? previous.ToUpperInvariant() : previous, previousRun.Text);
                TestAssert.True(previousRun.Strike && !previousRun.Underline, "Previous counters keep their independent strike decoration.");
                TestAssert.Equal(body, paragraph.Runs.First().AllCaps);
                TestAssert.True(label.Text.StartsWith("step ", StringComparison.Ordinal), "Display resolution cannot rewrite counter state.");
            }
        }
    }

    private static DocxDocument ReadListCasing(string container, OoxPdfDocxMarkupMode mode,
        bool? levelCaps, bool? markCaps, bool bodyCaps, string suffix, bool upperPrefix = false,
        string? labelPrefix = null, Action<XDocument>? paragraphChange = null)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        void SetCaps(XElement owner, bool? value)
        {
            var properties = owner.Element(w + "rPr");
            if (properties is null)
            {
                properties = new XElement(w + "rPr");
                if (owner.Name == w + "r") owner.AddFirst(properties); else owner.Add(properties);
            }
            properties.Elements(w + "caps").Remove();
            if (value is { } caps) properties.Add(new XElement(w + "caps", new XAttribute(w + "val", caps ? "true" : "false")));
        }
        return ReadMutation("CompositionProbes/numbering-mark-math-" + container, mode, false, false,
            (xml, _, _) =>
            {
                foreach (var paragraph in xml.Descendants(w + "p").Where(p => p.Element(w + "pPr")?.Element(w + "numPr") is not null))
                {
                    SetCaps(paragraph.Element(w + "pPr")!, markCaps);
                    SetCaps(paragraph.Elements(w + "r").First(), bodyCaps);
                }
                paragraphChange?.Invoke(xml);
            }, packageChange: archive =>
            {
                var entry = archive.GetEntry("word/numbering.xml")!; XDocument xml;
                using (var stream = entry.Open()) xml = XDocument.Load(stream);
                foreach (var level in xml.Descendants(w + "lvl"))
                {
                    SetCaps(level, levelCaps);
                    level.Element(w + "lvlText")!.SetAttributeValue(w + "val", (labelPrefix ?? (upperPrefix ? "STEP " : "step ")) + "%1.");
                    level.Elements(w + "suff").Remove();
                    level.Add(new XElement(w + "suff", new XAttribute(w + "val", suffix)));
                }
                entry.Delete(); using var output = archive.CreateEntry("word/numbering.xml").Open(); xml.Save(output);
            }, partOverride: container is "header" or "footer" ? "word/" + container + "1.xml" : null);
    }
}
