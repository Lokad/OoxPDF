using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfGraphicsBuilder
{
    internal const double SyntheticItalicShear = 0.213d;

    private readonly StringBuilder builder = new();
    private readonly List<PdfExtGStateResource> extGStates = [];
    private readonly List<PdfShadingResource> shadings = [];
    private readonly List<PdfTilingPatternResource> patterns = [];
    private int stateDepth;

    public IReadOnlyList<PdfExtGStateResource> ExtGStates => extGStates;

    public IReadOnlyList<PdfShadingResource> Shadings => shadings;

    public IReadOnlyList<PdfTilingPatternResource> Patterns => patterns;

    public int StateDepth => stateDepth;

    public void SetFillRgb(byte red, byte green, byte blue)
    {
        if (TryAppendFillGray(red, green, blue))
        {
            return;
        }

        builder.Append(PdfDocumentWriter.FormatColor(red)).Append(' ').Append(PdfDocumentWriter.FormatColor(green)).Append(' ').Append(PdfDocumentWriter.FormatColor(blue)).AppendLine(" rg");
    }

    public void SetStrokeRgb(byte red, byte green, byte blue)
    {
        if (TryAppendStrokeGray(red, green, blue))
        {
            return;
        }

        builder.Append(PdfDocumentWriter.FormatColor(red)).Append(' ').Append(PdfDocumentWriter.FormatColor(green)).Append(' ').Append(PdfDocumentWriter.FormatColor(blue)).AppendLine(" RG");
    }

    public void SetLineWidth(double width)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(width)).AppendLine(" w");
    }

    public void SetLineDash(double dashLength, double gapLength)
    {
        builder.Append('[').Append(PdfDocumentWriter.FormatNumber(dashLength)).Append(' ').Append(PdfDocumentWriter.FormatNumber(gapLength)).AppendLine("] 0 d");
    }

    public void SetLineDash(IReadOnlyList<double> lengths)
    {
        builder.Append('[');
        foreach (double length in lengths)
        {
            builder.Append(PdfDocumentWriter.FormatNumber(length)).Append(' ');
        }

        builder.AppendLine("] 0 d");
    }

    public void ClearLineDash()
    {
        builder.AppendLine("[] 0 d");
    }

    public void SetLineCap(int lineCap)
    {
        builder.Append(lineCap.ToString(CultureInfo.InvariantCulture)).AppendLine(" J");
    }

    public void SetLineJoin(int lineJoin)
    {
        builder.Append(lineJoin.ToString(CultureInfo.InvariantCulture)).AppendLine(" j");
    }

    public void SaveState()
    {
        builder.AppendLine("q");
        stateDepth++;
    }

    public void RestoreState()
    {
        builder.AppendLine("Q");
        if (stateDepth > 0)
        {
            stateDepth--;
        }
    }

    public void RestoreToStateDepth(int targetDepth)
    {
        targetDepth = Math.Max(0, targetDepth);
        while (stateDepth > targetDepth)
        {
            RestoreState();
        }
    }

    public void SetAlpha(double fillAlpha, double strokeAlpha)
    {
        fillAlpha = Math.Clamp(fillAlpha, 0d, 1d);
        strokeAlpha = Math.Clamp(strokeAlpha, 0d, 1d);
        string resourceName = "GS" +
            ((int)Math.Round(fillAlpha * 100000d, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture) +
            "F" +
            ((int)Math.Round(strokeAlpha * 100000d, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture) +
            "S";
        if (!extGStates.Any(state => state.ResourceName.Equals(resourceName, StringComparison.Ordinal)))
        {
            extGStates.Add(new PdfExtGStateResource(resourceName, fillAlpha, strokeAlpha, null));
        }

        builder.Append('/').Append(resourceName).AppendLine(" gs");
    }

    public void SetLuminositySoftMask(PdfLuminositySoftMask mask, double fillAlpha, double strokeAlpha)
    {
        fillAlpha = Math.Clamp(fillAlpha, 0d, 1d);
        strokeAlpha = Math.Clamp(strokeAlpha, 0d, 1d);
        string resourceName = "GSM" + (extGStates.Count + 1).ToString(CultureInfo.InvariantCulture);
        foreach (PdfExtGStateResource state in extGStates)
        {
            if (state.SoftMask is not null &&
                state.SoftMask.ResourceKey.Equals(mask.ResourceKey, StringComparison.Ordinal) &&
                Math.Abs(state.FillAlpha - fillAlpha) < 0.000001d &&
                Math.Abs(state.StrokeAlpha - strokeAlpha) < 0.000001d)
            {
                builder.Append('/').Append(state.ResourceName).AppendLine(" gs");
                return;
            }
        }

        extGStates.Add(new PdfExtGStateResource(resourceName, fillAlpha, strokeAlpha, mask));
        builder.Append('/').Append(resourceName).AppendLine(" gs");
    }

    public void Transform(double a, double b, double c, double d, double e, double f)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(a)).Append(' ').Append(PdfDocumentWriter.FormatNumber(b)).Append(' ');
        builder.Append(PdfDocumentWriter.FormatNumber(c)).Append(' ').Append(PdfDocumentWriter.FormatNumber(d)).Append(' ');
        builder.Append(PdfDocumentWriter.FormatNumber(e)).Append(' ').Append(PdfDocumentWriter.FormatNumber(f)).AppendLine(" cm");
    }

    public void FillRectangle(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re f");
    }

    public void FillRectangleEvenOdd(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re f*");
    }

    public void FillRectangleWithTilingPattern(double x, double y, double width, double height, PdfTilingPattern pattern)
    {
        string resourceName = "P" + (patterns.Count + 1).ToString(CultureInfo.InvariantCulture);
        PdfTilingPatternResource? existing = patterns.FirstOrDefault(resource => resource.Pattern.ResourceKey == pattern.ResourceKey);
        if (existing is null)
        {
            patterns.Add(new PdfTilingPatternResource(resourceName, pattern));
        }
        else
        {
            resourceName = existing.ResourceName;
        }

        builder.Append("/Pattern cs /").Append(PdfEmbeddedFont.SanitizeName(resourceName)).AppendLine(" scn");
        FillRectangle(x, y, width, height);
    }

    public void PaintAxialShading(double x0, double y0, double x1, double y1, byte startRed, byte startGreen, byte startBlue, byte endRed, byte endGreen, byte endBlue)
    {
        PaintAxialShading(x0, y0, x1, y1, [new PdfShadingStop(0d, startRed, startGreen, startBlue), new PdfShadingStop(1d, endRed, endGreen, endBlue)]);
    }

    public void PaintAxialShading(double x0, double y0, double x1, double y1, IReadOnlyList<PdfShadingStop> stops)
    {
        var shading = new PdfAxialShading(x0, y0, x1, y1, stops);
        string resourceName = "Sh" + (shadings.Count + 1).ToString(CultureInfo.InvariantCulture);
        PdfShadingResource? existing = shadings.FirstOrDefault(resource => resource.Shading.ResourceKey == shading.ResourceKey);
        if (existing is null)
        {
            shadings.Add(new PdfShadingResource(resourceName, shading));
        }
        else
        {
            resourceName = existing.ResourceName;
        }

        builder.Append('/').Append(resourceName).AppendLine(" sh");
    }

    public void StrokeRectangle(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re S");
    }

    public void FillStrokeRectangleEvenOdd(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re B*");
    }

    public void FillRoundedRectangle(double x, double y, double width, double height, double radius)
    {
        AppendRoundedRectanglePath(x, y, width, height, radius);
        builder.AppendLine("f");
    }

    public void FillRoundedRectangleEvenOdd(double x, double y, double width, double height, double radius)
    {
        AppendRoundedRectanglePath(x, y, width, height, radius);
        builder.AppendLine("f*");
    }

    public void StrokeRoundedRectangle(double x, double y, double width, double height, double radius)
    {
        AppendRoundedRectanglePath(x, y, width, height, radius);
        builder.AppendLine("S");
    }

    public void ClipRectangle(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re W n");
    }

    public void ClipRectangleEvenOdd(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re W* n");
    }

    public void ClipOpenRectangleEvenOdd(double x, double y, double width, double height)
    {
        AppendOpenRectanglePath(x, y, width, height);
        builder.AppendLine("W* n");
    }

    public void ClipEllipse(double x, double y, double width, double height)
    {
        AppendEllipsePath(x, y, width, height);
        builder.AppendLine("W n");
    }

    public void ClipRoundedRectangle(double x, double y, double width, double height, double radius)
    {
        AppendRoundedRectanglePath(x, y, width, height, radius);
        builder.AppendLine("W n");
    }

    public void ClipPolygon((double X, double Y)[] points)
    {
        AppendPolygonPath(points);
        builder.AppendLine("W n");
    }

    public void ClipCurrentPath()
    {
        builder.AppendLine("W n");
    }

    public void StrokeLine(double x1, double y1, double x2, double y2)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x1)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y1)).Append(" m ");
        builder.Append(PdfDocumentWriter.FormatNumber(x2)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y2)).AppendLine(" l S");
    }

    public void FillPolygon((double X, double Y)[] points)
    {
        AppendPolygonPath(points);
        builder.AppendLine("f");
    }

    public void FillPolygonEvenOdd((double X, double Y)[] points)
    {
        AppendPolygonPath(points);
        builder.AppendLine("f*");
    }

    public void MoveTo(double x, double y)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).AppendLine(" m");
    }

    public void LineTo(double x, double y)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).AppendLine(" l");
    }

    public void CurveTo(double x1, double y1, double x2, double y2, double x3, double y3)
    {
        Curve(x1, y1, x2, y2, x3, y3);
    }

    public void ClosePath()
    {
        builder.AppendLine("h");
    }

    public void FillCurrentPath()
    {
        builder.AppendLine("f");
    }

    public void FillCurrentPathEvenOdd()
    {
        builder.AppendLine("f*");
    }

    public void StrokeCurrentPath()
    {
        builder.AppendLine("S");
    }

    public void StrokePolygon((double X, double Y)[] points)
    {
        AppendPolygonPath(points);
        builder.AppendLine("S");
    }

    public void FillEllipse(double x, double y, double width, double height)
    {
        AppendEllipsePath(x, y, width, height);
        builder.AppendLine("f");
    }

    public void FillEllipseEvenOdd(double x, double y, double width, double height)
    {
        AppendEllipsePath(x, y, width, height);
        builder.AppendLine("f*");
    }

    public void StrokeEllipse(double x, double y, double width, double height)
    {
        AppendEllipsePath(x, y, width, height);
        builder.AppendLine("S");
    }

    public void DrawGlyphText(
        string fontResourceName,
        double fontSize,
        double x,
        double y,
        byte red,
        byte green,
        byte blue,
        string glyphHex,
        bool italic,
        double characterSpacing,
        int textRenderingMode,
        byte strokeRed,
        byte strokeGreen,
        byte strokeBlue,
        double strokeWidth)
    {
        DrawGlyphTextOperator(fontResourceName, fontSize, x, y, red, green, blue, '<' + glyphHex + "> Tj", italic, characterSpacing, textRenderingMode, strokeRed, strokeGreen, strokeBlue, strokeWidth);
    }

    public void DrawGlyphPositionedText(
        string fontResourceName,
        double fontSize,
        double x,
        double y,
        byte red,
        byte green,
        byte blue,
        string glyphPositioningArray,
        bool italic,
        double characterSpacing,
        int textRenderingMode,
        byte strokeRed,
        byte strokeGreen,
        byte strokeBlue,
        double strokeWidth)
    {
        DrawGlyphTextOperator(fontResourceName, fontSize, x, y, red, green, blue, glyphPositioningArray + " TJ", italic, characterSpacing, textRenderingMode, strokeRed, strokeGreen, strokeBlue, strokeWidth);
    }

    private void DrawGlyphTextOperator(
        string fontResourceName,
        double fontSize,
        double x,
        double y,
        byte red,
        byte green,
        byte blue,
        string textOperator,
        bool italic,
        double characterSpacing,
        int textRenderingMode,
        byte strokeRed,
        byte strokeGreen,
        byte strokeBlue,
        double strokeWidth)
    {
        builder.AppendLine("BT");
        if (!TryAppendFillGray(red, green, blue))
        {
            builder.Append(PdfDocumentWriter.FormatColor(red)).Append(' ').Append(PdfDocumentWriter.FormatColor(green)).Append(' ').Append(PdfDocumentWriter.FormatColor(blue)).AppendLine(" rg");
        }

        if (textRenderingMode is 1 or 2)
        {
            if (!TryAppendStrokeGray(strokeRed, strokeGreen, strokeBlue))
            {
                builder.Append(PdfDocumentWriter.FormatColor(strokeRed)).Append(' ').Append(PdfDocumentWriter.FormatColor(strokeGreen)).Append(' ').Append(PdfDocumentWriter.FormatColor(strokeBlue)).AppendLine(" RG");
            }

            builder.Append(PdfDocumentWriter.FormatNumber(strokeWidth)).AppendLine(" w");
            builder.Append(textRenderingMode.ToString(CultureInfo.InvariantCulture)).AppendLine(" Tr");
        }

        builder.Append('/').Append(PdfEmbeddedFont.SanitizeName(fontResourceName)).Append(' ').Append(PdfDocumentWriter.FormatNumber(fontSize)).AppendLine(" Tf");
        builder.Append(Math.Abs(characterSpacing) > 0.001d ? PdfDocumentWriter.FormatNumber(characterSpacing) : "0").AppendLine(" Tc");

        double shear = italic ? SyntheticItalicShear : 0d;
        builder.Append("1 0 ").Append(PdfDocumentWriter.FormatNumber(shear)).Append(" 1 ").Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).AppendLine(" Tm");
        builder.AppendLine(textOperator);
        if (textRenderingMode is 1 or 2)
        {
            builder.AppendLine("0 Tr");
        }

        builder.AppendLine("ET");
    }

    public void DrawImage(string imageResourceName, double x, double y, double width, double height)
    {
        builder.AppendLine("q");
        builder.Append(PdfDocumentWriter.FormatNumber(width)).Append(" 0 0 ").Append(PdfDocumentWriter.FormatNumber(height)).Append(' ').Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).AppendLine(" cm");
        builder.Append('/').Append(PdfEmbeddedFont.SanitizeName(imageResourceName)).AppendLine(" Do");
        builder.AppendLine("Q");
    }

    public void DrawImageCropped(string imageResourceName, double x, double y, double width, double height, double cropLeft, double cropTop, double cropRight, double cropBottom)
    {
        double visibleWidth = Math.Max(0.001d, 1d - cropLeft - cropRight);
        double visibleHeight = Math.Max(0.001d, 1d - cropTop - cropBottom);
        double scaledWidth = width / visibleWidth;
        double scaledHeight = height / visibleHeight;
        double imageX = x - cropLeft * scaledWidth;
        double imageY = y - cropBottom * scaledHeight;

        builder.AppendLine("q");
        builder.Append(PdfDocumentWriter.FormatNumber(scaledWidth)).Append(" 0 0 ").Append(PdfDocumentWriter.FormatNumber(scaledHeight)).Append(' ').Append(PdfDocumentWriter.FormatNumber(imageX)).Append(' ').Append(PdfDocumentWriter.FormatNumber(imageY)).AppendLine(" cm");
        builder.Append('/').Append(PdfEmbeddedFont.SanitizeName(imageResourceName)).AppendLine(" Do");
        builder.AppendLine("Q");
    }

    // Operation/resource boundary for node-level transactional recovery (S09).
    // Snapshots capture the append-only builder state so a failed node rewinds its
    // paint without disturbing earlier nodes. Font/image caches are intentionally
    // outside the boundary: orphan entries are inert, while index rollback could
    // dangle references held by surviving content.
    public readonly record struct ContentMark(int ContentLength, int ExtGStateCount, int ShadingCount, int PatternCount, int StateDepth);

    public ContentMark MarkContent()
    {
        return new ContentMark(builder.Length, extGStates.Count, shadings.Count, patterns.Count, stateDepth);
    }

    public void TruncateContent(ContentMark mark)
    {
        if (mark.ContentLength < builder.Length)
        {
            builder.Length = Math.Max(0, mark.ContentLength);
        }

        while (extGStates.Count > mark.ExtGStateCount)
        {
            extGStates.RemoveAt(extGStates.Count - 1);
        }

        while (shadings.Count > mark.ShadingCount)
        {
            shadings.RemoveAt(shadings.Count - 1);
        }

        while (patterns.Count > mark.PatternCount)
        {
            patterns.RemoveAt(patterns.Count - 1);
        }

        stateDepth = Math.Max(0, mark.StateDepth);
    }

    public override string ToString()
    {
        return builder.ToString();
    }

    private void AppendEllipsePath(double x, double y, double width, double height)
    {
        const double kappa = 0.5522847498307936d;
        double rx = width / 2d;
        double ry = height / 2d;
        double cx = x + rx;
        double cy = y + ry;
        double ox = rx * kappa;
        double oy = ry * kappa;

        builder.Append(PdfDocumentWriter.FormatNumber(cx + rx)).Append(' ').Append(PdfDocumentWriter.FormatNumber(cy)).AppendLine(" m");
        Curve(cx + rx, cy + oy, cx + ox, cy + ry, cx, cy + ry);
        Curve(cx - ox, cy + ry, cx - rx, cy + oy, cx - rx, cy);
        Curve(cx - rx, cy - oy, cx - ox, cy - ry, cx, cy - ry);
        Curve(cx + ox, cy - ry, cx + rx, cy - oy, cx + rx, cy);
        builder.AppendLine("h");
    }

    private void AppendRoundedRectanglePath(double x, double y, double width, double height, double radius)
    {
        const double kappa = 0.5522847498307936d;
        double r = Math.Clamp(radius, 0d, Math.Min(width, height) / 2d);
        double ox = r * kappa;

        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y + height - r)).AppendLine(" m");
        Curve(x, y + height - r + ox, x + r - ox, y + height, x + r, y + height);
        builder.Append(PdfDocumentWriter.FormatNumber(x + width - r)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y + height)).AppendLine(" l");
        Curve(x + width - r + ox, y + height, x + width, y + height - r + ox, x + width, y + height - r);
        builder.Append(PdfDocumentWriter.FormatNumber(x + width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y + r)).AppendLine(" l");
        Curve(x + width, y + r - ox, x + width - r + ox, y, x + width - r, y);
        builder.Append(PdfDocumentWriter.FormatNumber(x + r)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).AppendLine(" l");
        Curve(x + r - ox, y, x, y + r - ox, x, y + r);
        builder.AppendLine("h");
    }

    private void AppendOpenRectanglePath(double x, double y, double width, double height)
    {
        double right = x + width;
        double top = y + height;
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(top)).AppendLine(" m");
        builder.Append(PdfDocumentWriter.FormatNumber(right)).Append(' ').Append(PdfDocumentWriter.FormatNumber(top)).AppendLine(" l");
        builder.Append(PdfDocumentWriter.FormatNumber(right)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).AppendLine(" l");
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).AppendLine(" l");
    }

    private void AppendPolygonPath((double X, double Y)[] points)
    {
        if (points.Length == 0)
        {
            return;
        }

        builder.Append(PdfDocumentWriter.FormatNumber(points[0].X)).Append(' ').Append(PdfDocumentWriter.FormatNumber(points[0].Y)).AppendLine(" m");
        for (int i = 1; i < points.Length; i++)
        {
            builder.Append(PdfDocumentWriter.FormatNumber(points[i].X)).Append(' ').Append(PdfDocumentWriter.FormatNumber(points[i].Y)).AppendLine(" l");
        }

        builder.AppendLine("h");
    }

    private void Curve(double x1, double y1, double x2, double y2, double x3, double y3)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x1)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y1)).Append(' ');
        builder.Append(PdfDocumentWriter.FormatNumber(x2)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y2)).Append(' ');
        builder.Append(PdfDocumentWriter.FormatNumber(x3)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y3)).AppendLine(" c");
    }

    private bool TryAppendFillGray(byte red, byte green, byte blue)
    {
        if (red != green || red != blue)
        {
            return false;
        }

        builder.Append(PdfDocumentWriter.FormatColor(red)).AppendLine(" g");
        return true;
    }

    private bool TryAppendStrokeGray(byte red, byte green, byte blue)
    {
        if (red != green || red != blue)
        {
            return false;
        }

        builder.Append(PdfDocumentWriter.FormatColor(red)).AppendLine(" G");
        return true;
    }
}
