using System.Globalization;

namespace Lokad.OoxPdf.Pdf;

internal sealed record PdfShadingPatternResource(string ResourceName, PdfShadingPattern Pattern);

// A pattern matrix maps shading space to the parent stream's initial user space.
// It is independent of later cm operators used to construct a stroked path.
internal sealed class PdfShadingPattern
{
    public PdfShadingPattern(PdfShading shading, PdfPatternMatrix matrix)
    {
        ArgumentNullException.ThrowIfNull(shading);
        double Printed(double value)
        {
            if (!double.IsFinite(value)) { throw new InvalidDataException("Shading patterns require finite geometry."); }
            return double.Parse(PdfDocumentWriter.FormatNumber(value), CultureInfo.InvariantCulture);
        }
        Matrix = new PdfPatternMatrix(Printed(matrix.A), Printed(matrix.B), Printed(matrix.C), Printed(matrix.D), Printed(matrix.E), Printed(matrix.F));
        double determinant = Matrix.A * Matrix.D - Matrix.B * Matrix.C;
        if (!double.IsFinite(determinant) || determinant == 0d)
        {
            throw new InvalidDataException("Shading patterns require a nonsingular printed matrix.");
        }
        if (shading.Stops.Any(stop => !double.IsFinite(stop.Offset)))
        {
            throw new InvalidDataException("Shading patterns require finite stop offsets.");
        }
        if (shading is PdfAxialShading axial)
        {
            double x0 = Printed(axial.X0), y0 = Printed(axial.Y0), x1 = Printed(axial.X1), y1 = Printed(axial.Y1);
            double lengthSquared = (x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0);
            if (!double.IsFinite(lengthSquared) || lengthSquared <= 0d)
            {
                throw new InvalidDataException("Shading patterns require a finite nonzero printed axis.");
            }
        }
        Shading = shading;
        ResourceKey = shading.ResourceKey + ":pattern:" + Matrix;
    }

    public PdfShading Shading { get; }
    public PdfPatternMatrix Matrix { get; }
    public string ResourceKey { get; }
}
