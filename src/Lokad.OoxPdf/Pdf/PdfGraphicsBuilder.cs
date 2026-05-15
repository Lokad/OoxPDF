using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfGraphicsBuilder
{
    private readonly StringBuilder builder = new();

    public void SetFillRgb(byte red, byte green, byte blue)
    {
        builder.Append(C(red)).Append(' ').Append(C(green)).Append(' ').Append(C(blue)).AppendLine(" rg");
    }

    public void SetStrokeRgb(byte red, byte green, byte blue)
    {
        builder.Append(C(red)).Append(' ').Append(C(green)).Append(' ').Append(C(blue)).AppendLine(" RG");
    }

    public void SetLineWidth(double width)
    {
        builder.Append(N(width)).AppendLine(" w");
    }

    public void SetLineDash(double dashLength, double gapLength)
    {
        builder.Append('[').Append(N(dashLength)).Append(' ').Append(N(gapLength)).AppendLine("] 0 d");
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
    }

    public void RestoreState()
    {
        builder.AppendLine("Q");
    }

    public void Transform(double a, double b, double c, double d, double e, double f)
    {
        builder.Append(N(a)).Append(' ').Append(N(b)).Append(' ');
        builder.Append(N(c)).Append(' ').Append(N(d)).Append(' ');
        builder.Append(N(e)).Append(' ').Append(N(f)).AppendLine(" cm");
    }

    public void FillRectangle(double x, double y, double width, double height)
    {
        builder.Append(N(x)).Append(' ').Append(N(y)).Append(' ').Append(N(width)).Append(' ').Append(N(height)).AppendLine(" re f");
    }

    public void StrokeRectangle(double x, double y, double width, double height)
    {
        builder.Append(N(x)).Append(' ').Append(N(y)).Append(' ').Append(N(width)).Append(' ').Append(N(height)).AppendLine(" re S");
    }

    public void FillRoundedRectangle(double x, double y, double width, double height, double radius)
    {
        AppendRoundedRectanglePath(x, y, width, height, radius);
        builder.AppendLine("f");
    }

    public void StrokeRoundedRectangle(double x, double y, double width, double height, double radius)
    {
        AppendRoundedRectanglePath(x, y, width, height, radius);
        builder.AppendLine("S");
    }

    public void ClipRectangle(double x, double y, double width, double height)
    {
        builder.Append(N(x)).Append(' ').Append(N(y)).Append(' ').Append(N(width)).Append(' ').Append(N(height)).AppendLine(" re W n");
    }

    public void StrokeLine(double x1, double y1, double x2, double y2)
    {
        builder.Append(N(x1)).Append(' ').Append(N(y1)).Append(" m ");
        builder.Append(N(x2)).Append(' ').Append(N(y2)).AppendLine(" l S");
    }

    public void FillPolygon((double X, double Y)[] points)
    {
        AppendPolygonPath(points);
        builder.AppendLine("f");
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

    public void StrokeEllipse(double x, double y, double width, double height)
    {
        AppendEllipsePath(x, y, width, height);
        builder.AppendLine("S");
    }

    public void DrawGlyphText(string fontResourceName, double fontSize, double x, double y, byte red, byte green, byte blue, string glyphHex, bool italic = false, double characterSpacing = 0d)
    {
        DrawGlyphTextOperator(fontResourceName, fontSize, x, y, red, green, blue, '<' + glyphHex + "> Tj", italic, characterSpacing);
    }

    public void DrawGlyphPositionedText(string fontResourceName, double fontSize, double x, double y, byte red, byte green, byte blue, string glyphPositioningArray, bool italic = false, double characterSpacing = 0d)
    {
        DrawGlyphTextOperator(fontResourceName, fontSize, x, y, red, green, blue, glyphPositioningArray + " TJ", italic, characterSpacing);
    }

    private void DrawGlyphTextOperator(string fontResourceName, double fontSize, double x, double y, byte red, byte green, byte blue, string textOperator, bool italic, double characterSpacing)
    {
        builder.AppendLine("BT");
        builder.Append(C(red)).Append(' ').Append(C(green)).Append(' ').Append(C(blue)).AppendLine(" rg");
        builder.Append('/').Append(PdfEmbeddedFont.SanitizeName(fontResourceName)).Append(' ').Append(N(fontSize)).AppendLine(" Tf");
        if (Math.Abs(characterSpacing) > 0.001d)
        {
            builder.Append(N(characterSpacing)).AppendLine(" Tc");
        }

        double shear = italic ? 0.213d : 0d;
        builder.Append("1 0 ").Append(N(shear)).Append(" 1 ").Append(N(x)).Append(' ').Append(N(y)).AppendLine(" Tm");
        builder.AppendLine(textOperator);
        builder.AppendLine("ET");
    }

    public void DrawImage(string imageResourceName, double x, double y, double width, double height)
    {
        builder.AppendLine("q");
        builder.Append(N(width)).Append(" 0 0 ").Append(N(height)).Append(' ').Append(N(x)).Append(' ').Append(N(y)).AppendLine(" cm");
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
        builder.Append(N(x)).Append(' ').Append(N(y)).Append(' ').Append(N(width)).Append(' ').Append(N(height)).AppendLine(" re W n");
        builder.Append(N(scaledWidth)).Append(" 0 0 ").Append(N(scaledHeight)).Append(' ').Append(N(imageX)).Append(' ').Append(N(imageY)).AppendLine(" cm");
        builder.Append('/').Append(PdfEmbeddedFont.SanitizeName(imageResourceName)).AppendLine(" Do");
        builder.AppendLine("Q");
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

        builder.Append(N(cx + rx)).Append(' ').Append(N(cy)).AppendLine(" m");
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

        builder.Append(N(x + r)).Append(' ').Append(N(y)).AppendLine(" m");
        builder.Append(N(x + width - r)).Append(' ').Append(N(y)).AppendLine(" l");
        Curve(x + width - r + ox, y, x + width, y + r - ox, x + width, y + r);
        builder.Append(N(x + width)).Append(' ').Append(N(y + height - r)).AppendLine(" l");
        Curve(x + width, y + height - r + ox, x + width - r + ox, y + height, x + width - r, y + height);
        builder.Append(N(x + r)).Append(' ').Append(N(y + height)).AppendLine(" l");
        Curve(x + r - ox, y + height, x, y + height - r + ox, x, y + height - r);
        builder.Append(N(x)).Append(' ').Append(N(y + r)).AppendLine(" l");
        Curve(x, y + r - ox, x + r - ox, y, x + r, y);
        builder.AppendLine("h");
    }

    private void AppendPolygonPath((double X, double Y)[] points)
    {
        if (points.Length == 0)
        {
            return;
        }

        builder.Append(N(points[0].X)).Append(' ').Append(N(points[0].Y)).AppendLine(" m");
        for (int i = 1; i < points.Length; i++)
        {
            builder.Append(N(points[i].X)).Append(' ').Append(N(points[i].Y)).AppendLine(" l");
        }

        builder.AppendLine("h");
    }

    private void Curve(double x1, double y1, double x2, double y2, double x3, double y3)
    {
        builder.Append(N(x1)).Append(' ').Append(N(y1)).Append(' ');
        builder.Append(N(x2)).Append(' ').Append(N(y2)).Append(' ');
        builder.Append(N(x3)).Append(' ').Append(N(y3)).AppendLine(" c");
    }

    private static string C(byte value)
    {
        return (value / 255d).ToString("0.###", CultureInfo.InvariantCulture);
    }

    private static string N(double value)
    {
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }
}
