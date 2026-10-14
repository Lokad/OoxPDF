using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static void ReportEquationDiagnostics(DocxDocument document, IEnumerable<DocxEquationLayout> placed,
        Action<OoxPdfDiagnostic>? sink, CancellationToken token)
    {
        if (sink is null) { return; }
        var painted = new HashSet<DocxEquation>(ReferenceEqualityComparer.Instance);
        foreach (DocxEquationLayout item in placed)
        { token.ThrowIfCancellationRequested(); painted.Add(item.Math.Equation); }
        foreach (OoxPdfDiagnostic source in document.EquationDiagnostics)
        {
            token.ThrowIfCancellationRequested();
            if (source.PartName == document.MainPartName && painted.Count > 0 && painted.Count == document.EquationSourceCount) { continue; }
            sink(source);
        }
        if (painted.Count != 0)
        {
            sink(new OoxPdfDiagnostic("DOCX_APPROXIMATED_EQUATION", OoxPdfSeverity.Warning,
                "Native DOCX display equations were rendered with approximate mathematical layout.", document.MainPartName,
                SlideIndex: null, PageIndex: null, Feature: "equation", Fallback: "Approximated"));
        }
    }

    private static IReadOnlyDictionary<DocxParagraph, DocxPreparedMathParagraph> PrepareMathParagraphs(
        DocxDocument document, IFontResolver resolver, Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        List<PdfFontResource> resources, CancellationToken token)
    {
        var candidates = new List<(DocxParagraph Paragraph, DocxEquation Equation)>();
        var groups = new Dictionary<(string StableId, int FaceIndex), (OpenTypeFont Font, HashSet<int> Points)>();
        var runKeys = new Dictionary<DocxMathRun, (string StableId, int FaceIndex)>();
        foreach (DocxParagraphElement element in document.BodyElements.OfType<DocxParagraphElement>())
        {
            token.ThrowIfCancellationRequested();
            DocxParagraph paragraph = element.Paragraph;
            DocxEquation? equation = DocxMathParagraphPolicy.GetDisplayRow(paragraph);
            if (equation is null || !DocxMathRowLayout.TryGetRenderCodePoints(equation.Content, token, out _, out _)) { continue; }
            bool valid = true;
            foreach (DocxMathRun run in equation.Content.Children.OfType<DocxMathRun>())
            {
                token.ThrowIfCancellationRequested();
                string family = run.ExplicitFontFamily ?? document.Settings.MathFontFamily ?? "Cambria Math";
                FontFaceResolution resolution = resolver.Resolve(new FontRequest(family));
                OpenTypeFont? font = LoadFont(resolution, fontCache, token);
                if (font is null || !font.HasTrueTypeOutlines) { valid = false; break; }
                var key = (resolution.Source.StableId, resolution.FontFaceIndex);
                if (!groups.TryGetValue(key, out var group)) { group = (font, new HashSet<int>()); groups.Add(key, group); }
                DocxMathRowLayout.TryGetRenderCodePoints(new DocxMathRow([run]), token, out IReadOnlyList<int> points, out _);
                group.Points.UnionWith(points);
                runKeys[run] = key;
            }
            if (valid) { candidates.Add((paragraph, equation)); }
        }
        var fonts = new Dictionary<(string StableId, int FaceIndex), PdfFontResource>();
        foreach (var group in groups)
        {
            token.ThrowIfCancellationRequested();
            var resource = new PdfFontResource("Eq" + (fonts.Count + 1), PdfEmbeddedFont.Create(group.Value.Font, group.Value.Points, token));
            fonts.Add(group.Key, resource);
            resources.Add(resource);
        }
        var layouts = new Dictionary<DocxParagraph, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var candidate in candidates)
        {
            token.ThrowIfCancellationRequested();
            if (!DocxMathRowLayout.TryMeasure(candidate.Equation.Content,
                run => runKeys.TryGetValue(run, out var key) ? fonts[key] : null, token, out DocxMathBox? box, out _) ||
                box is null || box.Glyphs.Count == 0) { continue; }
            double above = box.AboveBaseline, below = box.BelowBaseline;
            foreach (DocxMathGlyphPlacement glyph in box.Glyphs)
            {
                var font = glyph.Font.Font.Font;
                above = Math.Max(above, DocxLineMetrics.MeasureTypographicBaselineInset(font, glyph.FontSize));
                below = Math.Max(below, -font.Os2.TypographicDescender * glyph.FontSize / font.UnitsPerEm);
            }
            double lineHeight = above + below;
            if (!double.IsFinite(above) || !double.IsFinite(lineHeight) || lineHeight <= 0d) { continue; }
            layouts.Add(candidate.Paragraph, new(candidate.Equation, box, above, lineHeight));
        }
        return layouts;
    }
}
