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

    // R17: internal so geometry boundary tests pin the finite/range contract directly.
    internal readonly record struct GroupTransform(
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
            // R17: one finite/range contract at the transform boundary. Hostile EMU
            // inputs and non-finite scales/rotations fail as malformed geometry (the
            // per-node renderer treats InvalidDataException as a skipped node with a
            // diagnostic) instead of overflowing long math, wrapping offsets, or
            // carrying NaN into layout. Long differences widen to double before
            // scaling so the subtraction itself cannot wrap; every narrowing cast is
            // range-checked and every long sum runs checked. In-range inputs follow
            // the exact legacy rounding order, so rendered output is unchanged.
            if (!double.IsFinite(RotationDegrees) || !double.IsFinite(ScaleX) || !double.IsFinite(ScaleY))
            {
                throw new InvalidDataException("Group transform scale or rotation is not finite.");
            }

            double rotationSum = bounds.RotationDegrees + RotationDegrees;
            if (!double.IsFinite(rotationSum))
            {
                throw new InvalidDataException("Group transform rotation is not finite.");
            }

            try
            {
                checked
                {
                    long width = ToEmu(bounds.Width * ScaleX, "width");
                    long height = ToEmu(bounds.Height * ScaleY, "height");
                    long localX = ToEmu(((double)bounds.X - ChildOffsetX) * ScaleX, "x");
                    long localY = ToEmu(((double)bounds.Y - ChildOffsetY) * ScaleY, "y");
                    long x = FlipHorizontal
                        ? OffsetX + Width - localX - width
                        : OffsetX + localX;
                    long y = FlipVertical
                        ? OffsetY + Height - localY - height
                        : OffsetY + localY;
                    double rotationDegrees = NormalizeRotationDegrees(rotationSum);
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
                        x = ToEmu(rotatedCenterX - width / 2d, "x");
                        y = ToEmu(rotatedCenterY - height / 2d, "y");
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
            }
            catch (OverflowException ex)
            {
                throw new InvalidDataException("Group transform coordinates overflow.", ex);
            }
        }

        private static long ToEmu(double value, string what)
        {
            // R17: the upper bound is exactly 2^63, not (double)long.MaxValue (which
            // rounds up to 2^63): doubles between long.MaxValue and 2^63 pass a naive
            // check but overflow the cast below. Doubles are integral past 2^52, so
            // Math.Round cannot push an accepted value back out of range.
            if (!double.IsFinite(value) || value >= -(double)long.MinValue || value < long.MinValue)
            {
                throw new InvalidDataException($"Group transform {what} is out of range.");
            }

            return (long)Math.Round(value);
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
