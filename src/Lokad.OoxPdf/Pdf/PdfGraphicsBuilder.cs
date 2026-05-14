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

    public void StrokeLine(double x1, double y1, double x2, double y2)
    {
        builder.Append(N(x1)).Append(' ').Append(N(y1)).Append(" m ");
        builder.Append(N(x2)).Append(' ').Append(N(y2)).AppendLine(" l S");
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

    public void DrawGlyphText(string fontResourceName, double fontSize, double x, double y, byte red, byte green, byte blue, string glyphHex)
    {
        builder.AppendLine("BT");
        builder.Append(C(red)).Append(' ').Append(C(green)).Append(' ').Append(C(blue)).AppendLine(" rg");
        builder.Append('/').Append(PdfEmbeddedFont.SanitizeName(fontResourceName)).Append(' ').Append(N(fontSize)).AppendLine(" Tf");
        builder.Append("1 0 0 1 ").Append(N(x)).Append(' ').Append(N(y)).AppendLine(" Tm");
        builder.Append('<').Append(glyphHex).AppendLine("> Tj");
        builder.AppendLine("ET");
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
