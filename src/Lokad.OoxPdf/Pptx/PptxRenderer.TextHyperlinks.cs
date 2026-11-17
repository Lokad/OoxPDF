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

        if (!TryGetSpanLinkRect(run, embedded, scope.Context.Document, out double x, out double y, out double width, out double height))
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
        PptxDocument document,
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
            PdfRectangle frame = TransformedRectangleBounds(new PdfRectangle(x0, y0, x1 - x0, y1 - y0), TextTransformMatrix(run));
            x0 = frame.X;
            x1 = frame.X + frame.Width;
            y0 = frame.Y;
            y1 = frame.Y + frame.Height;
        }

        // Paint clips precede rotation, while unrotated flips transform the
        // local clip. Navigation uses that same physical frame and slide crop.
        var clip = new PdfRectangle(run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
        if (!TextClipUsesPhysicalCoordinates(run) && HasTextTransform(run))
            clip = TransformedRectangleBounds(clip, TextTransformMatrix(run));
        x0 = Math.Max(x0, clip.X);
        y0 = Math.Max(y0, clip.Y);
        x1 = Math.Min(x1, clip.X + clip.Width);
        y1 = Math.Min(y1, clip.Y + clip.Height);
        return TryIntersectWithSlideBounds(x0, y0, x1 - x0, y1 - y0, document, out x, out y, out width, out height);
    }
}
