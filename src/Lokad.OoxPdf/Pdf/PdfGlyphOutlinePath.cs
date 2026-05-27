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
        double fontSize,
        double shear = 0d)
    {
        if (!font.TryReadGlyphOutline(glyphId, out var outline) || outline.Contours.Count == 0)
        {
            return false;
        }

        AppendGlyphPath(graphics, font, outline, x, y, fontSize, shear);
        return true;
    }

    public static void AppendGlyphPath(
        PdfGraphicsBuilder graphics,
        OpenTypeFont font,
        OpenTypeFont.OpenTypeGlyphOutline outline,
        double x,
        double y,
        double fontSize,
        double shear = 0d)
    {
        double scale = font.UnitsPerEm == 0 ? 0d : fontSize / font.UnitsPerEm;
        foreach (OpenTypeFont.OpenTypeGlyphContour contour in outline.Contours)
        {
            AppendContourPath(graphics, contour, x, y, scale, shear);
        }
    }

    private static void AppendContourPath(
        PdfGraphicsBuilder graphics,
        OpenTypeFont.OpenTypeGlyphContour contour,
        double x,
        double y,
        double scale,
        double shear)
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
        PdfPoint pdfStart = ToPdfPoint(start, x, y, scale, shear);
        graphics.MoveTo(pdfStart.X, pdfStart.Y);
        DPoint current = start;

        for (int i = 1; i < points.Length;)
        {
            DPoint point = points[i];
            if (point.IsOnCurve)
            {
                PdfPoint pdfPoint = ToPdfPoint(point, x, y, scale, shear);
                graphics.LineTo(pdfPoint.X, pdfPoint.Y);
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

            PdfPoint pdfControl1 = ToPdfPoint(control1, x, y, scale, shear);
            PdfPoint pdfControl2 = ToPdfPoint(control2, x, y, scale, shear);
            PdfPoint pdfEnd = ToPdfPoint(end, x, y, scale, shear);
            graphics.CurveTo(pdfControl1.X, pdfControl1.Y, pdfControl2.X, pdfControl2.Y, pdfEnd.X, pdfEnd.Y);

            current = end;
            i += 2;
        }

        graphics.ClosePath();
    }

    private static PdfPoint ToPdfPoint(DPoint point, double x, double y, double scale, double shear)
    {
        double scaledY = point.Y * scale;
        return new PdfPoint(x + point.X * scale + scaledY * shear, y + scaledY);
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

    private readonly record struct PdfPoint(double X, double Y);
}
