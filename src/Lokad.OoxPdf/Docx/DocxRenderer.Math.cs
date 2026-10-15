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
        var candidates = new List<(DocxParagraph Paragraph, DocxEquation Equation, OpenTypeFont? ParentMetrics)>();
        var groups = new Dictionary<(string StableId, int FaceIndex), (OpenTypeFont Font, Dictionary<ushort, int> Glyphs)>();
        var runKeys = new Dictionary<DocxMathRun, (string StableId, int FaceIndex)>();
        foreach (DocxParagraphElement element in document.BodyElements.OfType<DocxParagraphElement>())
        {
            token.ThrowIfCancellationRequested();
            DocxParagraph paragraph = element.Paragraph;
            DocxEquation? equation = DocxMathParagraphPolicy.GetDisplayRow(paragraph) ?? DocxMathParagraphPolicy.GetInlineRow(paragraph);
            if (equation is null || !DocxMathLayout.TryGetRuns(equation, token, out IReadOnlyList<DocxMathRunSource> mathRuns)) { continue; }
            OpenTypeFont? parentMetrics = null;
            if (DocxMathLayout.NeedsParentFractionMetrics(equation, mathRuns, token))
            {
                FontFaceResolution parent = resolver.Resolve(new FontRequest(document.Settings.MathFontFamily ?? "Cambria Math"));
                parentMetrics = LoadFont(parent, fontCache, token);
                if (parentMetrics is null || !parentMetrics.TryGetMathConstants(out _)) { continue; }
            }
            bool valid = true;
            foreach (DocxMathRunSource source in mathRuns)
            {
                token.ThrowIfCancellationRequested();
                DocxMathRun run = source.Run;
                string family = run.NormalText
                    ? run.ExplicitFontFamily ?? run.Text.FontFamily ?? document.Settings.MathFontFamily ?? "Cambria Math"
                    : (run.NormalTextValue is false ? null : run.ExplicitFontFamily) ?? document.Settings.MathFontFamily ?? "Cambria Math";
                FontRequest request = run.NormalText ? new(family, run.Text.Bold, run.Text.Italic) : new(family);
                FontFaceResolution resolution = resolver.Resolve(request);
                // A normal-text face owns its authored weight/slant. Missing
                // styles need separate synthetic-emission qualification.
                if (run.NormalText && !resolution.HasMathTable &&
                    (request.Bold != resolution.Bold || request.Italic != resolution.Italic))
                    { valid = false; break; }
                OpenTypeFont? font = LoadFont(resolution, fontCache, token);
                if (font is null || !font.HasTrueTypeOutlines) { valid = false; break; }
                var key = (resolution.Source.StableId, resolution.FontFaceIndex);
                if (!groups.TryGetValue(key, out var group)) { group = (font, new Dictionary<ushort, int>()); groups.Add(key, group); }
                DocxMathRowLayout.TryGetRenderCodePoints(new DocxMathRow([run]), token, out IReadOnlyList<int> points, out _);
                foreach (int point in points)
                {
                    ushort glyph = font.MapCodePoint(point);
                    if (glyph == 0 || source.ScriptLevel != 0 && !run.NormalText && !font.TryGetMathScriptGlyph(glyph, source.ScriptLevel, token, out glyph))
                    { valid = false; break; }
                    if (!group.Glyphs.TryGetValue(glyph, out int existing) || existing != 0x20 && point == 0x20) { group.Glyphs[glyph] = point; }
                }
                if (!valid) { break; }
                runKeys[run] = key;
            }
            if (valid) { candidates.Add((paragraph, equation, parentMetrics)); }
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
                run => runKeys.TryGetValue(run, out var key) ? fonts[key] : null, token, out DocxMathBox? box, out _, candidate.ParentMetrics) ||
                box is null || box.Glyphs.Count == 0) { continue; }
            double lineSpacingFactor = 1d;
            if (candidate.Equation.Display && !DocxMathParagraphPolicy.TryGetDisplayLineSpacingFactor(candidate.Paragraph, out lineSpacingFactor))
                { continue; }
            if (!DocxMathLayout.TryGetLineMetrics(box, token, out double above, out double lineHeight, lineSpacingFactor)) { continue; }
            layouts.Add(candidate.Paragraph, new(candidate.Equation, box, above, lineHeight));
        }
        return layouts;
    }
}
