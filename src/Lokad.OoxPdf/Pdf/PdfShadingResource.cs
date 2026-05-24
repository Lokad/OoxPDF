using System.Globalization;

namespace Lokad.OoxPdf.Pdf;

internal sealed record PdfShadingResource(string ResourceName, PdfAxialShading Shading);

internal sealed class PdfAxialShading
{
    private string? resourceKey;

    public PdfAxialShading(double X0, double Y0, double X1, double Y1, byte StartRed, byte StartGreen, byte StartBlue, byte EndRed, byte EndGreen, byte EndBlue)
    {
        this.X0 = X0;
        this.Y0 = Y0;
        this.X1 = X1;
        this.Y1 = Y1;
        this.StartRed = StartRed;
        this.StartGreen = StartGreen;
        this.StartBlue = StartBlue;
        this.EndRed = EndRed;
        this.EndGreen = EndGreen;
        this.EndBlue = EndBlue;
    }

    public double X0 { get; }

    public double Y0 { get; }

    public double X1 { get; }

    public double Y1 { get; }

    public byte StartRed { get; }

    public byte StartGreen { get; }

    public byte StartBlue { get; }

    public byte EndRed { get; }

    public byte EndGreen { get; }

    public byte EndBlue { get; }

    public string ResourceKey => resourceKey ??= string.Create(CultureInfo.InvariantCulture, $"axial:{X0:0.###}:{Y0:0.###}:{X1:0.###}:{Y1:0.###}:{StartRed:X2}{StartGreen:X2}{StartBlue:X2}:{EndRed:X2}{EndGreen:X2}{EndBlue:X2}");
}
