using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static void ReportEquationDiagnostics(DocxDocument document, IEnumerable<DocxEquation> placed,
        Action<OoxPdfDiagnostic>? sink, CancellationToken token)
    {
        if (sink is null) { return; }
        var painted = new HashSet<DocxEquation>(ReferenceEqualityComparer.Instance);
        foreach (DocxEquation equation in placed)
        { token.ThrowIfCancellationRequested(); painted.Add(equation); }
        var paintedByPart = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (DocxEquationPartSource part in document.EquationParts)
        {
            token.ThrowIfCancellationRequested();
            var seen = new HashSet<DocxEquation>(ReferenceEqualityComparer.Instance);
            foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(part.BodyElements))
            foreach (DocxEquation equation in paragraph.Equations)
            {
                token.ThrowIfCancellationRequested();
                if (painted.Contains(equation)) { seen.Add(equation); }
            }
            if (seen.Count > 0) { paintedByPart[part.PartName] = seen.Count; }
        }
        var sourcesByPart = document.EquationParts.ToDictionary(part => part.PartName, StringComparer.Ordinal);
        foreach (OoxPdfDiagnostic source in document.EquationDiagnostics)
        {
            token.ThrowIfCancellationRequested();
            if (source.PartName is { } sourcePart && sourcesByPart.TryGetValue(sourcePart, out var part) &&
                paintedByPart.TryGetValue(sourcePart, out int count) && count == part.SourceCount) { continue; }
            sink(source);
        }
        var paintedParts = new HashSet<string>(paintedByPart.Keys, StringComparer.Ordinal);
        if (paintedByPart.Values.Sum() < painted.Count) { paintedParts.Add(document.MainPartName); }
        foreach (string part in paintedParts)
        {
            token.ThrowIfCancellationRequested();
            sink(new OoxPdfDiagnostic("DOCX_APPROXIMATED_EQUATION", OoxPdfSeverity.Warning,
                "Native DOCX equations were rendered with approximate mathematical layout.", part,
                SlideIndex: null, PageIndex: null, Feature: "equation", Fallback: "Approximated"));
        }
    }

    private static IEnumerable<DocxEquation> EnumeratePaintedEquations(IEnumerable<DocxLayoutItem> items, CancellationToken token)
    {
        var rows = new Stack<DocxTableRowLayout>();
        foreach (DocxLayoutItem item in items)
        {
            token.ThrowIfCancellationRequested();
            if (item is DocxEquationLayout equation) { yield return equation.Math.Equation; }
            else if (item is DocxTextLineLayout { Math: { } math }) { yield return math.Equation; }
            else if (item is DocxTableRowLayout row) { rows.Push(row); }
        }
        while (rows.TryPop(out var row))
        {
            token.ThrowIfCancellationRequested();
            foreach (DocxTableCellLayout cell in row.Cells)
            {
                foreach (DocxTextLineLayout line in cell.TextLines)
                {
                    token.ThrowIfCancellationRequested();
                    if (line.Math is { } math) { yield return math.Equation; }
                }
                foreach (DocxTableRowLayout nested in cell.NestedRows) { token.ThrowIfCancellationRequested(); rows.Push(nested); }
            }
        }
    }

    private static void PaintMath(DocxPreparedMathParagraph math, double x, double baseline, double scale,
        PdfGraphicsBuilder graphics, DocxMarkupContext context, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        graphics.SaveState();
        graphics.Transform(scale, 0d, 0d, scale, x + ResolveTextEmissionXOffset(context),
            baseline - ResolveTextEmissionBaselineOffset(context));
        DocxMathBoxPainter.Paint(math.Box, graphics, 0d, 0d, token);
        graphics.RestoreState();
    }

    private static IReadOnlyDictionary<DocxParagraph, DocxPreparedMathParagraph> PrepareMathParagraphs(
        DocxDocument document, IFontResolver resolver, Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        List<PdfFontResource> resources, CancellationToken token)
    {
        var candidates = new List<(DocxParagraph Paragraph, DocxEquation Equation, DocxEquation SourceEquation, OpenTypeFont? ParentMetrics, IReadOnlyList<OpenTypeFont> ControlFaces)>();
        var groups = new Dictionary<(string StableId, int FaceIndex), (OpenTypeFont Font, Dictionary<ushort, int> Glyphs)>();
        var runKeys = new Dictionary<DocxMathRun, (string StableId, int FaceIndex)>();
        var controlFaces = new Dictionary<OpenTypeFont, (Dictionary<ushort, int> Glyphs, Dictionary<ushort, int> Operators, HashSet<int> Characters)>(ReferenceEqualityComparer.Instance);
        var controlKeys = new Dictionary<IDocxMathControl, (string StableId, int FaceIndex)>(ReferenceEqualityComparer.Instance);
        var unavailableControlFaces = new HashSet<OpenTypeFont>(ReferenceEqualityComparer.Instance);
        foreach (DocxParagraph paragraph in DocxMathParagraphPolicy.EnumeratePreparationCandidates(document, token)
            .Concat(DocxMathParagraphPolicy.EnumerateStaticPreparationCandidates(document, token)))
        {
            token.ThrowIfCancellationRequested();
            DocxEquation? standalone = DocxMathParagraphPolicy.GetStandaloneRow(paragraph);
            DocxEquation? sourceEquation = standalone ?? DocxMathParagraphPolicy.GetInlineRow(paragraph);
            if (sourceEquation is null) { continue; }
            DocxEquation equation = standalone is { Display: false } isolated
                ? isolated with { Display = true } : sourceEquation;
            if (!DocxMathLayout.TryGetRuns(equation, token, out IReadOnlyList<DocxMathRunSource> mathRuns)) { continue; }
            OpenTypeFont? parentMetrics = null;
            List<OpenTypeFont>? controlOwners = null;
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
                foreach (IDocxMathControl control in source.ControlOwners)
                {
                    token.ThrowIfCancellationRequested();
                    IReadOnlyList<int> controlPoints = DocxMathLayout.GetControlCodePoints(control);
                    if (!controlKeys.TryGetValue(control, out var controlKey))
                    {
                        FontFaceResolution controlFace = resolver.Resolve(new FontRequest(control.ControlFontFamily ?? document.Settings.MathFontFamily ?? "Cambria Math"));
                        OpenTypeFont? controlFont = LoadFont(controlFace, fontCache, token);
                        if (controlFont is null || (controlPoints.Count != 0 && !controlFont.HasTrueTypeOutlines) || !controlFont.TryGetMathConstants(out _))
                        { valid = false; break; }
                        controlKey = (controlFace.Source.StableId, controlFace.FontFaceIndex);
                        // Metric-only controls retain the conversion-owned font
                        // without preparing an unused PDF subset or glyph map.
                        if (controlPoints.Count != 0)
                        {
                            if (!groups.TryGetValue(controlKey, out var controlGroup))
                            { controlGroup = (controlFont, new Dictionary<ushort, int>()); groups.Add(controlKey, controlGroup); }
                            if (!controlFaces.TryGetValue(controlFont, out var face))
                            {
                                face = (controlGroup.Glyphs, new Dictionary<ushort, int>(), new HashSet<int>());
                                controlFaces.Add(controlFont, face);
                            }
                            foreach (int point in controlPoints)
                            {
                                if (control is DocxMathDelimiter { GrowValue: false })
                                {
                                    ushort canonical = controlFont.MapCodePoint(point);
                                    if (canonical == 0) { valid = false; break; }
                                    if (!TryRegisterOperatorGlyph(face.Operators, canonical, point))
                                        unavailableControlFaces.Add(controlFont);
                                    // This registers only the authored base glyph;
                                    // a later growing control still prepares variants.
                                    continue;
                                }
                                if (face.Characters.Contains(point)) { continue; }
                                ushort glyph = controlFont.MapCodePoint(point);
                                double height = 0d;
                                bool found = false;
                                while (controlFont.TryGetMathVerticalVariant(glyph, height, token, out var variant))
                                {
                                    found = true;
                                    if (!TryRegisterOperatorGlyph(face.Operators, variant.Glyph, point))
                                        unavailableControlFaces.Add(controlFont);
                                    height = variant.Advance + 1d;
                                }
                                if (!found) { valid = false; break; }
                                face.Characters.Add(point);
                            }
                            if (!valid) { break; }
                        }
                        controlKeys.Add(control, controlKey);
                    }
                    if (controlPoints.Count != 0) { (controlOwners ??= []).Add(groups[controlKey].Font); }
                }
                if (!valid) { break; }
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
            if (valid) { candidates.Add((paragraph, equation, sourceEquation, parentMetrics, controlOwners ?? [])); }
        }
        foreach (var face in controlFaces)
        {
            token.ThrowIfCancellationRequested();
            // Publish implicit operators after all authored glyphs are known.
            // A source alias needs a separate encoding owner; keep control fallback
            // rather than changing that source character's Unicode identity.
            if (face.Value.Operators.Any(v => face.Value.Glyphs.TryGetValue(v.Key, out int point) && point != v.Value))
            { unavailableControlFaces.Add(face.Key); continue; }
            if (unavailableControlFaces.Contains(face.Key)) { continue; }
            foreach (var glyph in face.Value.Operators) { face.Value.Glyphs[glyph.Key] = glyph.Value; }
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
            if (candidate.ControlFaces.Any(unavailableControlFaces.Contains)) { continue; }
            if (!DocxMathLayout.TryMeasure(candidate.Equation,
                run => runKeys.TryGetValue(run, out var key) ? fonts[key] : null, token, out DocxMathBox? box, out _, candidate.ParentMetrics,
                control => controlKeys.TryGetValue(control, out var controlKey) && fonts.TryGetValue(controlKey, out var resource) ? resource : null,
                control => controlKeys.TryGetValue(control, out var controlKey) && fontCache.TryGetValue(controlKey, out var metricFont) ? metricFont : null) ||
                box is null || box.Glyphs.Count == 0) { continue; }
            double lineSpacingFactor = 1d;
            if (candidate.Equation.Display && !DocxMathParagraphPolicy.TryGetDisplayLineSpacingFactor(candidate.Paragraph, out lineSpacingFactor))
                { continue; }
            if (!DocxMathLayout.TryGetLineMetrics(box, token, out double above, out double lineHeight, lineSpacingFactor)) { continue; }
            layouts.Add(candidate.Paragraph, new(candidate.SourceEquation, box, above, lineHeight));
        }
        return layouts;
    }

    private static bool TryRegisterOperatorGlyph(Dictionary<ushort, int> glyphs, ushort glyph, int point)
    {
        if (glyphs.TryGetValue(glyph, out int previous) && previous != point) { return false; }
        glyphs[glyph] = point;
        return true;
    }
}
