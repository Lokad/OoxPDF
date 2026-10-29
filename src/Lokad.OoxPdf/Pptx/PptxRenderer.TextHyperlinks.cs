using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Carries the slide hyperlink state into text emission so run-level clicks in
    // shape and table spans resolve against the owning part (S08).
    private sealed record PptxTextHyperlinkScope(
        PptxRenderContext Context,
        string? SourcePartName,
        List<PdfLinkAnnotation> Annotations,
        HashSet<string> ReportedIds)
    {
        public void CollectEmissionSpan(PptxPositionedTextSpan emissionSpan, PdfEmbeddedFont? embedded)
        {
            AddEmissionSpanHyperlinkAnnotation(emissionSpan, embedded, this);
        }
    }

    private static void AddEmissionSpanHyperlinkAnnotation(
        PptxPositionedTextSpan emissionSpan,
        PdfEmbeddedFont? embedded,
        PptxTextHyperlinkScope scope)
    {
        TextRun run = emissionSpan.Run;
        if (run.Text.Length == 0 || run.Width <= 0d)
        {
            return;
        }

        string? relationshipId = emissionSpan.SourceRun?.Style.HyperlinkClickId;
        if (string.IsNullOrEmpty(relationshipId))
        {
            if (emissionSpan.SourceRun?.Style.HasHyperlinkClick == true &&
                !string.IsNullOrEmpty(emissionSpan.SourceRun?.Style.HyperlinkClickAction) &&
                scope.ReportedIds.Add("action:" + emissionSpan.SourceRun?.Style.HyperlinkClickAction))
            {
                EmitHyperlinkDiagnostic(
                    scope.Context,
                    scope.SourcePartName ?? scope.Context.SlidePartName,
                    "PPTX_UNSUPPORTED_HYPERLINK_ACTION",
                    $"Hyperlink action '{emissionSpan.SourceRun?.Style.HyperlinkClickAction}' is not supported and was ignored.");
            }

            return;
        }

        PptxHyperlinkResolution resolution = ResolveHyperlink(scope.Context, scope.SourcePartName, relationshipId);
        if (!resolution.Resolved)
        {
            if (scope.ReportedIds.Add(resolution.FailureId + "\0" + relationshipId))
            {
                EmitHyperlinkDiagnostic(
                    scope.Context,
                    scope.SourcePartName ?? scope.Context.SlidePartName,
                    resolution.FailureId,
                    resolution.FailureMessage);
            }

            return;
        }

        if (!TryGetSpanLinkRect(run, embedded, out double x, out double y, out double width, out double height))
        {
            return;
        }

        scope.Annotations.Add(resolution.Uri is not null
            ? PdfLinkAnnotation.ToUri(x, y, width, height, resolution.Uri)
            : PdfLinkAnnotation.ToDestination(x, y, width, height, resolution.Destination!.Value));
    }

    // Baseline-anchored box from the emitted font metrics (null means the standard
    // PDF fallback), mirroring the DOCX
    // hyperlink contract; rotated/flipped runs reuse the emission matrix so the
    // link area covers the painted glyphs.
    private static bool TryGetSpanLinkRect(
        TextRun run,
        PdfEmbeddedFont? embedded,
        out double x,
        out double y,
        out double width,
        out double height)
    {
        x = 0d;
        y = 0d;
        width = 0d;
        height = 0d;
        double pdfFontSize = OfficePdfTextEmissionProfile.FontSize(run.FontSize);
        double unitsPerEm = embedded?.Font.UnitsPerEm ?? PdfFallbackFont.UnitsPerEm;
        if (unitsPerEm <= 0d || pdfFontSize <= 0d)
        {
            return false;
        }

        double ascender = embedded is null ? PdfFallbackFont.AscentEm * pdfFontSize
            : embedded.Font.Os2.WindowsAscender * pdfFontSize / unitsPerEm;
        double descender = embedded is null ? PdfFallbackFont.DescentEm * pdfFontSize
            : embedded.Font.Os2.WindowsDescender * pdfFontSize / unitsPerEm;
        double baselineY = run.Y + run.BaselineOffset;
        if (embedded is null && !BaselineIntersectsClip(run, baselineY))
        {
            return false;
        }
        double x0 = run.X;
        double y0 = baselineY - descender;
        double x1 = run.X + run.Width;
        double y1 = baselineY + ascender;
        if (Math.Abs(run.RotationDegrees) > PptxTextMetricRules.TextStateTolerance || run.FlipHorizontal || run.FlipVertical)
        {
            (double a, double b, double c, double d, double e, double f) = TextTransformMatrix(run);
            double tx0 = Math.Min(Math.Min(a * x0 + c * y0 + e, a * x1 + c * y0 + e), Math.Min(a * x0 + c * y1 + e, a * x1 + c * y1 + e));
            double tx1 = Math.Max(Math.Max(a * x0 + c * y0 + e, a * x1 + c * y0 + e), Math.Max(a * x0 + c * y1 + e, a * x1 + c * y1 + e));
            double ty0 = Math.Min(Math.Min(b * x0 + d * y0 + f, b * x1 + d * y0 + f), Math.Min(b * x0 + d * y1 + f, b * x1 + d * y1 + f));
            double ty1 = Math.Max(Math.Max(b * x0 + d * y0 + f, b * x1 + d * y0 + f), Math.Max(b * x0 + d * y1 + f, b * x1 + d * y1 + f));
            x0 = tx0;
            x1 = tx1;
            y0 = ty0;
            y1 = ty1;
        }

        if (x1 - x0 <= 0d || y1 - y0 <= 0d)
        {
            return false;
        }

        x = x0;
        y = y0;
        width = x1 - x0;
        height = y1 - y0;
        return true;
    }
}
