using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed record PdfShadingResource(string ResourceName, PdfShading Shading);

internal readonly record struct PdfShadingStop(double Offset, byte Red, byte Green, byte Blue);

internal abstract class PdfShading
{
    private string? resourceKey;

    protected PdfShading(IReadOnlyList<PdfShadingStop> stops, bool preserveHardStops)
    {
        Stops = NormalizeStops(stops, preserveHardStops);
    }

    protected abstract string CoordinateKey { get; }

    public IReadOnlyList<PdfShadingStop> Stops { get; }

    public string ResourceKey => resourceKey ??= BuildResourceKey();

    private string BuildResourceKey()
    {
        var builder = new StringBuilder();
        builder.Append(CoordinateKey);
        foreach (PdfShadingStop stop in Stops)
        {
            builder.Append(':').Append(FormatStopOffset(stop.Offset));
            builder.Append(CultureInfo.InvariantCulture, $":{stop.Red:X2}{stop.Green:X2}{stop.Blue:X2}");
        }

        return builder.ToString();
    }

    // Normalized offsets are in [0, 1]. Preserve their round-trip precision for
    // identity and stitching bounds; PDF numbers cannot use exponent notation.
    internal static string FormatStopOffset(double offset)
    {
        string number = PdfDocumentWriter.FormatNumber(offset, "R");
        int exponentIndex = number.IndexOf('E');
        if (exponentIndex < 0) return number;
        int exponent = int.Parse(number.AsSpan(exponentIndex + 1), CultureInfo.InvariantCulture);
        return "0." + new string('0', -exponent - 1) + number[..exponentIndex].Replace(".", "", StringComparison.Ordinal);
    }

    private static IReadOnlyList<PdfShadingStop> NormalizeStops(IReadOnlyList<PdfShadingStop> stops, bool preserveHardStops)
    {
        if (stops.Count < 2)
        {
            throw new ArgumentException("A shading requires at least two color stops.", nameof(stops));
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
            if (!preserveHardStops && normalized.Count != 0 && Math.Abs(normalized[^1].Offset - stop.Offset) < 0.000001d)
            {
                normalized[^1] = stop;
            }
            else
            {
                normalized.Add(stop);
            }
        }

        // Coalescing a stop very near zero may replace the synthetic endpoint.
        // Its color still owns the start of the function's [0, 1] domain.
        if (normalized[0].Offset > 0d) normalized[0] = normalized[0] with { Offset = 0d };
        PdfShadingStop last = normalized[^1];
        if (last.Offset < 1d)
        {
            normalized.Add(last with { Offset = 1d });
        }

        return normalized;
    }
}

internal sealed class PdfAxialShading : PdfShading
{
    public PdfAxialShading(double X0, double Y0, double X1, double Y1, byte StartRed, byte StartGreen, byte StartBlue, byte EndRed, byte EndGreen, byte EndBlue)
        : this(X0, Y0, X1, Y1, [new PdfShadingStop(0d, StartRed, StartGreen, StartBlue), new PdfShadingStop(1d, EndRed, EndGreen, EndBlue)])
    {
    }

    public PdfAxialShading(double X0, double Y0, double X1, double Y1, IReadOnlyList<PdfShadingStop> Stops, bool preserveHardStops = false)
        : base(Stops, preserveHardStops)
    {
        this.X0 = X0;
        this.Y0 = Y0;
        this.X1 = X1;
        this.Y1 = Y1;
    }

    public double X0 { get; }
    public double Y0 { get; }
    public double X1 { get; }
    public double Y1 { get; }

    protected override string CoordinateKey => FormattableString.Invariant($"axial:{X0:0.###}:{Y0:0.###}:{X1:0.###}:{Y1:0.###}");
}

// A centered unit-circle gradient; the caller maps it into the printed ellipse.
internal sealed class PdfRadialShading(IReadOnlyList<PdfShadingStop> stops, int cycleCount = 1, bool reflect = false,
    bool preserveHardStops = false) : PdfShading(stops, preserveHardStops)
{
    public int CycleCount { get; } = cycleCount is >= 1 and <= 128 ? cycleCount :
        throw new ArgumentOutOfRangeException(nameof(cycleCount));

    public bool Reflect { get; } = reflect;

    protected override string CoordinateKey => FormattableString.Invariant($"radial:0:0:0:0:0:1:{CycleCount}:{Reflect}");
}
