using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxStaticParagraphFramesShareBodyInsetsAndWrapping()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string frame in new[] { "first", "hanging", "right" })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadStaticInsetFrame("body", mode, frame), null, default).ToBlockingEnumerable().ToArray();
            var expected = StaticInsetRows(control);
            TestAssert.True(expected.Length >= 2, "The source exercises first and continuation constraints.");
            foreach (string container in new[] { "header", "footer" })
            {
                var source = ReadStaticInsetFrame(container, mode, frame);
                var actual = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
                var rows = StaticInsetRows(actual);
                TestAssert.Equal(expected.Length, rows.Length);
                for (int index = 0; index < rows.Length; index++)
                {
                    TestAssert.Equal(expected[index].Text, rows[index].Text);
                    TestAssert.True(Math.Abs(expected[index].X - rows[index].X) < .001d,
                        "Static paragraphs use the same first/continuation/right frame as body content: " + container + "/" + frame + "/" + mode);
                }
                string text = string.Join(' ', rows.Select(row => row.Text));
                foreach (string token in StaticInsetText().Split(' '))
                    TestAssert.True(text.Contains(token, StringComparison.Ordinal), "Every authored frame word remains visible.");
            }
        }
    }

    public static void DocxStaticParagraphFramesPreserveNeutralInsetPaintAndOwners()
    {
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (string container in new[] { "header", "footer" })
        {
            var renderer = new DocxRenderer(new PortableRenderingTests.PinnedResolver(), mode, CachedMathGeometry(mode));
            var control = renderer.RenderPagesAsync(ReadStaticInsetFrame(container, mode, "none", longText: false), null, default).ToBlockingEnumerable().ToArray();
            var source = ReadStaticInsetFrame(container, mode, "zero", longText: false);
            var paragraph = DocxBlockTraversal.EnumerateSourceParagraphs(source)
                .Distinct<DocxParagraph>(ReferenceEqualityComparer.Instance)
                .Single(p => p.Runs.Any(run => run.Text.StartsWith("FrameAlpha", StringComparison.Ordinal)));
            var indent = paragraph.EffectiveProperties.Indent;
            var actual = renderer.RenderPagesAsync(source, null, default).ToBlockingEnumerable().ToArray();
            AssertMetricPaint(control, actual);
            TestAssert.Equal(indent, paragraph.EffectiveProperties.Indent);
            TestAssert.True(paragraph.ListLabel is null && paragraph.Equations.Count == 0 && paragraph.Images.Count == 0,
                "Authored constraints require no synthetic label, equation or picture owner.");
        }
    }

    private static (string Text, double X)[] StaticInsetRows(PdfPage[] pages) =>
        pages.SelectMany(page => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true,
            includeTextTransforms: true, honorActualText: false))
            .Where(show => show.Text.Contains("Frame", StringComparison.Ordinal)).GroupBy(show => show.Y)
            .OrderBy(group => group.Key).Select(group =>
                (Regex.Replace(string.Join(' ', group.Select(show => show.Text.Trim())), @"\s+", " ").Trim(), group.Min(show => show.X))).ToArray();

    private static string StaticInsetText() => "FrameAlpha " + string.Join(' ', Enumerable.Range(1, 10).Select(index => $"Frame{index:00}")) + " FrameOmega.";

    private static DocxDocument ReadStaticInsetFrame(string container, OoxPdfDocxMarkupMode mode,
        string frame, bool longText = true) => ReadPlainTabFields(container, mode, "none", paragraphChange: (paragraph, w) =>
        {
            var properties = paragraph.Element(w + "pPr")!; properties.Elements(w + "ind").Remove();
            if (frame != "none")
            {
                var indent = new XElement(w + "ind", new XAttribute(w + "left", frame is "first" or "hanging" ? "720" : "0"),
                    new XAttribute(w + "right", frame == "right" ? "1440" : frame == "zero" ? "0" : "360"));
                if (frame is "first" or "zero") indent.Add(new XAttribute(w + "firstLine", frame == "zero" ? "0" : "240"));
                if (frame == "hanging") indent.Add(new XAttribute(w + "hanging", "240"));
                properties.Add(indent);
            }
            var textRuns = paragraph.Elements(w + "r").Where(run => run.Element(w + "t") is not null).ToArray();
            var style = textRuns[0].Element(w + "rPr");
            textRuns[0].AddBeforeSelf(new XElement(w + "r", style is null ? null : new XElement(style),
                new XElement(w + "t", longText ? StaticInsetText() : "FrameAlpha FrameOmega.")));
            foreach (var run in textRuns) run.Remove();
        });
}
