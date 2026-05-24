using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed record PdfShadingResource(string ResourceName, PdfAxialShading Shading);

internal readonly record struct PdfShadingStop(double Offset, byte Red, byte Green, byte Blue);

internal sealed class PdfAxialShading
{
    private string? resourceKey;

    public PdfAxialShading(double X0, double Y0, double X1, double Y1, byte StartRed, byte StartGreen, byte StartBlue, byte EndRed, byte EndGreen, byte EndBlue)
        : this(X0, Y0, X1, Y1, [new PdfShadingStop(0d, StartRed, StartGreen, StartBlue), new PdfShadingStop(1d, EndRed, EndGreen, EndBlue)])
    {
    }

    public PdfAxialShading(double X0, double Y0, double X1, double Y1, IReadOnlyList<PdfShadingStop> Stops)
    {
        this.X0 = X0;
        this.Y0 = Y0;
        this.X1 = X1;
        this.Y1 = Y1;
        this.Stops = NormalizeStops(Stops);
    }

    public double X0 { get; }

    public double Y0 { get; }

    public double X1 { get; }

    public double Y1 { get; }

    public IReadOnlyList<PdfShadingStop> Stops { get; }

    public string ResourceKey => resourceKey ??= BuildResourceKey();

    private string BuildResourceKey()
    {
        var builder = new StringBuilder();
        builder.Append(CultureInfo.InvariantCulture, $"axial:{X0:0.###}:{Y0:0.###}:{X1:0.###}:{Y1:0.###}");
        foreach (PdfShadingStop stop in Stops)
        {
            builder.Append(CultureInfo.InvariantCulture, $":{stop.Offset:0.#####}:{stop.Red:X2}{stop.Green:X2}{stop.Blue:X2}");
        }

        return builder.ToString();
    }

    private static IReadOnlyList<PdfShadingStop> NormalizeStops(IReadOnlyList<PdfShadingStop> stops)
    {
        if (stops.Count < 2)
        {
            throw new ArgumentException("An axial shading requires at least two color stops.", nameof(stops));
        }

        PdfShadingStop[] ordered = stops
            .Select(stop => stop with { Offset = Math.Clamp(stop.Offset, 0d, 1d) })
            .OrderBy(stop => stop.Offset)
            .ToArray();

        var normalized = new List<PdfShadingStop>(ordered.Length + 2);
        PdfShadingStop first = ordered[0];
        if (first.Offset > 0d)
        {
            normalized.Add(first with { Offset = 0d });
        }

        foreach (PdfShadingStop stop in ordered)
        {
            if (normalized.Count != 0 && Math.Abs(normalized[^1].Offset - stop.Offset) < 0.000001d)
            {
                normalized[^1] = stop;
            }
            else
            {
                normalized.Add(stop);
            }
        }

        PdfShadingStop last = normalized[^1];
        if (last.Offset < 1d)
        {
            normalized.Add(last with { Offset = 1d });
        }

        if (normalized.Count < 2)
        {
            PdfShadingStop only = normalized[0];
            normalized.Add(only with { Offset = 1d });
        }

        return normalized;
    }
}
