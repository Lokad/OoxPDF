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
                "Native DOCX equations were rendered with approximate mathematical layout.", document.MainPartName,
                SlideIndex: null, PageIndex: null, Feature: "equation", Fallback: "Approximated"));
        }
    }

    private static IReadOnlyDictionary<DocxParagraph, DocxPreparedMathParagraph> PrepareMathParagraphs(
        DocxDocument document, IFontResolver resolver, Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        List<PdfFontResource> resources, CancellationToken token)
    {
        var candidates = new List<(DocxParagraph Paragraph, DocxEquation Equation)>();
        var groups = new Dictionary<(string StableId, int FaceIndex), (OpenTypeFont Font, Dictionary<ushort, int> Glyphs)>();
        var runKeys = new Dictionary<DocxMathRun, (string StableId, int FaceIndex)>();
        foreach (DocxParagraphElement element in document.BodyElements.OfType<DocxParagraphElement>())
        {
            token.ThrowIfCancellationRequested();
            DocxParagraph paragraph = element.Paragraph;
            DocxEquation? equation = DocxMathParagraphPolicy.GetDisplayRow(paragraph) ?? DocxMathParagraphPolicy.GetInlineRow(paragraph);
            if (equation is null || !DocxMathLayout.TryGetRuns(equation, token, out IReadOnlyList<DocxMathRunSource> mathRuns)) { continue; }
            bool valid = true;
            foreach (DocxMathRunSource source in mathRuns)
            {
                token.ThrowIfCancellationRequested();
                DocxMathRun run = source.Run;
                string family = run.ExplicitFontFamily ?? document.Settings.MathFontFamily ?? "Cambria Math";
                FontFaceResolution resolution = resolver.Resolve(new FontRequest(family));
                OpenTypeFont? font = LoadFont(resolution, fontCache, token);
                if (font is null || !font.HasTrueTypeOutlines) { valid = false; break; }
                var key = (resolution.Source.StableId, resolution.FontFaceIndex);
                if (!groups.TryGetValue(key, out var group)) { group = (font, new Dictionary<ushort, int>()); groups.Add(key, group); }
                DocxMathRowLayout.TryGetRenderCodePoints(new DocxMathRow([run]), token, out IReadOnlyList<int> points, out _);
                foreach (int point in points)
                {
                    ushort glyph = font.MapCodePoint(point);
                    if (glyph == 0 || source.ScriptLevel != 0 && !font.TryGetMathScriptGlyph(glyph, source.ScriptLevel, token, out glyph))
                    { valid = false; break; }
                    if (!group.Glyphs.TryGetValue(glyph, out int existing) || existing != 0x20 && point == 0x20) { group.Glyphs[glyph] = point; }
                }
                if (!valid) { break; }
                runKeys[run] = key;
            }
            if (valid) { candidates.Add((paragraph, equation)); }
        }
        var fonts = new Dictionary<(string StableId, int FaceIndex), PdfFontResource>();
        foreach (var group in groups)
        {
            token.ThrowIfCancellationRequested();
            var resource = new PdfFontResource("Eq" + (fonts.Count + 1), PdfEmbeddedFont.CreateForGlyphMappings(group.Value.Font,
                group.Value.Glyphs.Select(pair => (pair.Key, pair.Value)), token));
            fonts.Add(group.Key, resource);
            resources.Add(resource);
        }
        var layouts = new Dictionary<DocxParagraph, DocxPreparedMathParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var candidate in candidates)
        {
            token.ThrowIfCancellationRequested();
            if (!DocxMathLayout.TryMeasure(candidate.Equation,
                run => runKeys.TryGetValue(run, out var key) ? fonts[key] : null, token, out DocxMathBox? box, out _) ||
                box is null || box.Glyphs.Count == 0) { continue; }
            if (!DocxMathLayout.TryGetLineMetrics(box, token, out double above, out double lineHeight)) { continue; }
            layouts.Add(candidate.Paragraph, new(candidate.Equation, box, above, lineHeight));
        }
        return layouts;
    }
}
