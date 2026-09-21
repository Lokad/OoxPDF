using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal readonly record struct ShapeBounds(
        long X,
        long Y,
        long Width,
        long Height,
        double RotationDegrees,
        bool FlipHorizontal,
        bool FlipVertical);

    private readonly record struct GroupTransform(
        long OffsetX,
        long OffsetY,
        long Width,
        long Height,
        long ChildOffsetX,
        long ChildOffsetY,
        double ScaleX,
        double ScaleY,
        double RotationDegrees,
        bool FlipHorizontal,
        bool FlipVertical)
    {
        public static GroupTransform Identity { get; } = new(0, 0, 0, 0, 0, 0, 1d, 1d, 0d, FlipHorizontal: false, FlipVertical: false);

        public ShapeBounds Apply(ShapeBounds bounds)
        {
            long width = (long)Math.Round(bounds.Width * ScaleX);
            long height = (long)Math.Round(bounds.Height * ScaleY);
            long localX = (long)Math.Round((bounds.X - ChildOffsetX) * ScaleX);
            long localY = (long)Math.Round((bounds.Y - ChildOffsetY) * ScaleY);
            long x = FlipHorizontal
                ? OffsetX + Width - localX - width
                : OffsetX + localX;
            long y = FlipVertical
                ? OffsetY + Height - localY - height
                : OffsetY + localY;
            double rotationDegrees = NormalizeRotationDegrees(bounds.RotationDegrees + RotationDegrees);
            if (Math.Abs(RotationDegrees) > PptxTextMetricRules.TextStateTolerance && Width > 0 && Height > 0)
            {
                double groupCenterX = OffsetX + Width / 2d;
                double groupCenterY = OffsetY + Height / 2d;
                double boundsCenterX = x + width / 2d;
                double boundsCenterY = y + height / 2d;
                double radians = RotationDegrees * Math.PI / 180d;
                double cos = Math.Cos(radians);
                double sin = Math.Sin(radians);
                double dx = boundsCenterX - groupCenterX;
                double dy = boundsCenterY - groupCenterY;
                double rotatedCenterX = groupCenterX + dx * cos - dy * sin;
                double rotatedCenterY = groupCenterY + dx * sin + dy * cos;
                x = (long)Math.Round(rotatedCenterX - width / 2d);
                y = (long)Math.Round(rotatedCenterY - height / 2d);
            }

            return new ShapeBounds(
                x,
                y,
                width,
                height,
                rotationDegrees,
                bounds.FlipHorizontal ^ FlipHorizontal,
                bounds.FlipVertical ^ FlipVertical);
        }

        public GroupTransform Combine(GroupTransform child)
        {
            ShapeBounds childBounds = Apply(new ShapeBounds(
                child.OffsetX,
                child.OffsetY,
                child.Width,
                child.Height,
                child.RotationDegrees,
                FlipHorizontal: false,
                FlipVertical: false));
            return new GroupTransform(
                childBounds.X,
                childBounds.Y,
                childBounds.Width,
                childBounds.Height,
                child.ChildOffsetX,
                child.ChildOffsetY,
                ScaleX * child.ScaleX,
                ScaleY * child.ScaleY,
                childBounds.RotationDegrees,
                FlipHorizontal ^ child.FlipHorizontal,
                FlipVertical ^ child.FlipVertical);
        }
    }

    private readonly record struct BezierSegment(
        double StartX,
        double StartY,
        double Control1X,
        double Control1Y,
        double Control2X,
        double Control2Y,
        double EndX,
        double EndY);

    private readonly record struct CurveSample(
        double X,
        double Y,
        double TangentX,
        double TangentY);

    private readonly record struct CurvedConnectorFillPath(
        IReadOnlyList<(double X, double Y)> Points,
        IReadOnlyList<(double X, double Y)>? TailSubpath,
        double TipX,
        double TipY,
        double DirectionX,
        double DirectionY,
        double NormalX,
        double NormalY);
}
