using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Pdf;

internal static class PdfGlyphOutlinePath
{
    public static bool TryAppendGlyphPath(
        PdfGraphicsBuilder graphics,
        OpenTypeFont font,
        ushort glyphId,
        double x,
        double y,
        double fontSize)
    {
        if (!font.TryReadGlyphOutline(glyphId, out var outline) || outline.Contours.Count == 0)
        {
            return false;
        }

        AppendGlyphPath(graphics, font, outline, x, y, fontSize);
        return true;
    }

    public static void AppendGlyphPath(
        PdfGraphicsBuilder graphics,
        OpenTypeFont font,
        OpenTypeFont.OpenTypeGlyphOutline outline,
        double x,
        double y,
        double fontSize)
    {
        double scale = font.UnitsPerEm == 0 ? 0d : fontSize / font.UnitsPerEm;
        foreach (OpenTypeFont.OpenTypeGlyphContour contour in outline.Contours)
        {
            AppendContourPath(graphics, contour, x, y, scale);
        }
    }

    private static void AppendContourPath(
        PdfGraphicsBuilder graphics,
        OpenTypeFont.OpenTypeGlyphContour contour,
        double x,
        double y,
        double scale)
    {
        IReadOnlyList<OpenTypeFont.OpenTypeGlyphPoint> source = contour.Points;
        if (source.Count == 0)
        {
            return;
        }

        DPoint[] points = BuildPdfContourPoints(source);
        if (points.Length == 0 || !points[0].IsOnCurve)
        {
            return;
        }

        DPoint start = points[0];
        graphics.MoveTo(x + start.X * scale, y + start.Y * scale);
        DPoint current = start;

        for (int i = 1; i < points.Length;)
        {
            DPoint point = points[i];
            if (point.IsOnCurve)
            {
                graphics.LineTo(x + point.X * scale, y + point.Y * scale);
                current = point;
                i++;
                continue;
            }

            DPoint end = i + 1 < points.Length ? points[i + 1] : start;
            if (!end.IsOnCurve)
            {
                break;
            }

            DPoint control1 = new(
                current.X + (point.X - current.X) * 2d / 3d,
                current.Y + (point.Y - current.Y) * 2d / 3d,
                IsOnCurve: false);
            DPoint control2 = new(
                end.X + (point.X - end.X) * 2d / 3d,
                end.Y + (point.Y - end.Y) * 2d / 3d,
                IsOnCurve: false);

            graphics.CurveTo(
                x + control1.X * scale,
                y + control1.Y * scale,
                x + control2.X * scale,
                y + control2.Y * scale,
                x + end.X * scale,
                y + end.Y * scale);

            current = end;
            i += 2;
        }

        graphics.ClosePath();
    }

    private static DPoint[] BuildPdfContourPoints(IReadOnlyList<OpenTypeFont.OpenTypeGlyphPoint> source)
    {
        var normalized = new List<DPoint>(source.Count + 1);
        OpenTypeFont.OpenTypeGlyphPoint first = source[0];
        OpenTypeFont.OpenTypeGlyphPoint last = source[^1];
        if (first.IsOnCurve)
        {
            AppendSourcePoints(normalized, source, 0, source.Count);
        }
        else if (last.IsOnCurve)
        {
            normalized.Add(ToDPoint(last));
            AppendSourcePoints(normalized, source, 0, source.Count - 1);
        }
        else
        {
            normalized.Add(Midpoint(ToDPoint(last), ToDPoint(first)));
            AppendSourcePoints(normalized, source, 0, source.Count);
        }

        var expanded = new List<DPoint>(normalized.Count * 2);
        for (int i = 0; i < normalized.Count; i++)
        {
            DPoint current = normalized[i];
            DPoint next = normalized[(i + 1) % normalized.Count];
            expanded.Add(current);
            if (!current.IsOnCurve && !next.IsOnCurve)
            {
                expanded.Add(Midpoint(current, next));
            }
        }

        return expanded.ToArray();
    }

    private static void AppendSourcePoints(
        List<DPoint> points,
        IReadOnlyList<OpenTypeFont.OpenTypeGlyphPoint> source,
        int start,
        int count)
    {
        for (int i = 0; i < count; i++)
        {
            points.Add(ToDPoint(source[start + i]));
        }
    }

    private static DPoint ToDPoint(OpenTypeFont.OpenTypeGlyphPoint point)
    {
        return new DPoint(point.X, point.Y, point.IsOnCurve);
    }

    private static DPoint Midpoint(DPoint left, DPoint right)
    {
        return new DPoint((left.X + right.X) / 2d, (left.Y + right.Y) / 2d, IsOnCurve: true);
    }

    private readonly record struct DPoint(double X, double Y, bool IsOnCurve);
}
