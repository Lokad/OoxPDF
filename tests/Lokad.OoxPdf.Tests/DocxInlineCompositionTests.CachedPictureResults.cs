using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxCachedPictureResultsKeepSourceAcrossPlacementsAndFormats()
    {
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (bool vml in new[] { false, true })
        foreach (bool complex in new[] { false, true })
        foreach (var mode in Enum.GetValues<OoxPdfDocxMarkupMode>())
        {
            var before = PictureResultParagraph(ReadPictureResult(story, vml, complex, mode, baseline: true));
            var paragraph = PictureResultParagraph(ReadPictureResult(story, vml, complex, mode));
            var field = paragraph.FieldReferences.Single(f => f.Instruction!.Trim() == "REF StoredPicture");
            TestAssert.True(field.HasCachedResult && field.RendersCachedResult && !field.UsesPlaceholder,
                "A supported picture is a stored field result independently of ordinary text and image format.");
            TestAssert.Equal(0, field.ResultRunCount);
            TestAssert.Equal(before.Images.Count, paragraph.Images.Count);
            TestAssert.Equal(before.Equations.Count, paragraph.Equations.Count);
            TestAssert.True(before.Equations.Count > 0 && paragraph.InlineContent is not null, "Preserve the picture's supported equation neighbors.");
            TestAssert.Equal(DocxMathParagraphPolicy.GetPreparationEquations(before).Count,
                DocxMathParagraphPolicy.GetPreparationEquations(paragraph).Count);
        }
    }

    public static void DocxCachedPictureResultsPreservePaintAndNavigation()
    {
        var resolver = new PortableRenderingTests.PinnedResolver(mathFont: true);
        foreach (string story in new[] { "body", "cell", "header", "footer" })
        foreach (bool vml in new[] { false, true })
        foreach (bool complex in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var renderer = new DocxRenderer(resolver, mode, CachedMathGeometry(mode));
            var beforeDiagnostics = new List<OoxPdfDiagnostic>(); var afterDiagnostics = new List<OoxPdfDiagnostic>();
            var before = renderer.RenderPagesAsync(ReadPictureResult(story, vml, complex, mode, baseline: true),
                beforeDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            var after = renderer.RenderPagesAsync(ReadPictureResult(story, vml, complex, mode),
                afterDiagnostics.Add, default).ToBlockingEnumerable().ToArray();
            TestAssert.Equal(before.Length, after.Length);
            for (int index = 0; index < before.Length; index++)
            {
                TestAssert.Equal(before[index].Content, after[index].Content);
                TestAssert.Equal(before[index].Images.Count, after[index].Images.Count);
                TestAssert.Equal(before[index].Fonts.Count, after[index].Fonts.Count);
                TestAssert.True(before[index].Annotations.SequenceEqual(after[index].Annotations), "Field metadata cannot displace image, equation or navigation geometry.");
            }
            TestAssert.True(beforeDiagnostics.Select(d => d.Id).SequenceEqual(afterDiagnostics.Select(d => d.Id)),
                "The wrapper adds no unsupported-field or equation-omission warning.");
        }
    }

    public static void DocxCachedPictureResultsKeepNestedComputedAndEmptyOwners()
    {
        foreach (bool vml in new[] { false, true })
        foreach (bool complex in new[] { false, true })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var nested = ReadPictureResult("body", vml, complex, mode, nested: true);
            var paragraph = PictureResultParagraph(nested);
            foreach (string instruction in new[] { "REF StoredPicture", "NUMPAGES" })
                TestAssert.True(paragraph.FieldReferences.Where(f => f.Instruction!.Trim() == instruction)
                    .All(f => f.HasCachedResult && f.RendersCachedResult && !f.UsesPlaceholder), "Nested stored results retain their independent locks.");
            var pages = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode))
                .RenderPagesAsync(nested, null, default).ToBlockingEnumerable().ToArray();
            var baseline = new DocxRenderer(new PortableRenderingTests.PinnedResolver(mathFont: true), mode, CachedMathGeometry(mode))
                .RenderPagesAsync(ReadPictureResult("body", vml, complex, mode, baseline: true, nested: true), null, default)
                .ToBlockingEnumerable().ToArray();
            TestAssert.Equal(baseline.Length, pages.Length);
            for (int index = 0; index < pages.Length; index++)
                TestAssert.True(baseline[index].Annotations.SequenceEqual(pages[index].Annotations),
                    "Picture-only source links retain the independently scoped navigation fallback.");
            TestAssert.True(paragraph.Hyperlinks.Any(link => link.Anchor == "PublicCalculation"), "The source link remains inspectable.");
            var computed = PictureResultParagraph(ReadPictureResult("body", vml, complex, mode, computed: true));
            TestAssert.True(computed.FieldReferences.Any(f => f.Kind == DocxFieldKind.Page && f.UsesPlaceholder && !f.RendersCachedResult),
                "Computed counters retain replacement ownership.");
        }
        foreach (bool complex in new[] { false, true })
        {
            var empty = ReadMutation("CompositionProbes/math-picture-first-table-single-cell", OoxPdfDocxMarkupMode.Final, false, false,
                (xml, w, _) => xml.Descendants(w + "drawing").Single().Parent!.ReplaceWith(
                    CachedMathField(w, new XElement(w + "r", new XElement(w + "drawing")), complex, "REF StoredPicture")));
            var field = DocxBlockTraversal.EnumerateBodyParagraphs(empty.BodyElements).SelectMany(p => p.FieldReferences).Single();
            TestAssert.True(!field.HasCachedResult && !field.RendersCachedResult,
                "An empty drawing shell must not become a stored picture result.");
        }
    }

    public static void DocxCachedPictureResultsMatchIndependentWordContentAndOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent picture/math origins use installed Windows fonts.");
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "CompositionProbes"));
        using var gold = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "cached-picture-results-office.json")));
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(cases, row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            var mode = row.GetProperty("Mode").GetString() == "final-preserve" ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var page = new DocxRenderer(null, mode, CachedMathGeometry(mode)).RenderPagesAsync(DocxTests.ReadDocx(input, mode), null, default)
                .ToBlockingEnumerable().Single();
            TestAssert.Equal(1, page.Images.Count);
            var fonts = page.Fonts.Where(f => f.ResourceName.StartsWith("Eq", StringComparison.Ordinal)).Select(f => f.Font).ToHashSet();
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(show => fonts.Contains(show.Font)).ToArray();
            TestAssert.Equal("ratetimeratetime", string.Concat(math.Select(s => s.Text)).Normalize(NormalizationForm.FormKC));
            TestAssert.Equal(16, math.Length);
            int index = 0;
            foreach (var anchor in row.GetProperty("MathOrigins").EnumerateArray())
            {
                var actual = math[4 * index++];
                TestAssert.True(Math.Abs(actual.X - anchor.GetProperty("X").GetDouble()) < 1d &&
                    Math.Abs(actual.Y - anchor.GetProperty("Y").GetDouble()) < .2d,
                    "Stored pictures preserve the existing independent1pt X/.2pt Y fraction-origin gate: " + input + "/" + mode);
            }
        }
    }

    private static DocxDocument ReadPictureResult(string story, bool vml, bool complex, OoxPdfDocxMarkupMode mode,
        bool baseline = false, bool nested = false, bool computed = false) =>
        ReadMutation("CompositionProbes/math-composition-unnumbered" + (story == "body" ? "" : "-" + story), mode, true, true,
            (xml, w, _) =>
            {
                var drawing = xml.Descendants(w + "drawing").Single();
                var run = drawing.Parent!;
                if (vml)
                {
                    XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                    XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                    XNamespace v = "urn:schemas-microsoft-com:vml";
                    XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    var extent = drawing.Descendants(wp + "extent").Single();
                    string Size(string axis) => ((double)extent.Attribute(axis)! / 12700d).ToString(CultureInfo.InvariantCulture) + "pt";
                    drawing.ReplaceWith(new XElement(w + "pict", new XElement(v + "shape", new XAttribute("id", "PublicStoredPicture"),
                        new XAttribute("style", "width:" + Size("cx") + ";height:" + Size("cy")),
                        new XElement(v + "imagedata", new XAttribute(r + "id", drawing.Descendants(a + "blip").Single().Attribute(r + "embed")!.Value)))));
                }
                if (baseline && !nested) return;
                var content = baseline ? [new XElement(run)]
                    : CachedMathField(w, new XElement(run), complex, computed ? "PAGE" : "REF StoredPicture", locked: !computed);
                if (nested)
                {
                    var link = new XElement(w + "hyperlink", new XAttribute(w + "anchor", "PublicCalculation"), content);
                    content = baseline ? [link] : CachedMathField(w, link, !complex, "NUMPAGES");
                }
                run.ReplaceWith(content);
            }, partOverride: story is "header" or "footer" ? "word/" + story + "1.xml" : null);

    private static DocxParagraph PictureResultParagraph(DocxDocument document) =>
        DocxBlockTraversal.EnumerateBodyParagraphs(document.BodyElements).Concat(document.EquationParts.SelectMany(part =>
            DocxBlockTraversal.EnumerateBodyParagraphs(part.BodyElements))).First(p => p.Images.Count != 0);
}
