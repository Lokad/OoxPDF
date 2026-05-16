using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using Lokad.OoxPdf.VisualDiff;

if (args.Length != 3)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.VisualDiff <reference-png-directory> <candidate-png-directory> <output-directory>");
    return 2;
}

string referenceDirectory = Path.GetFullPath(args[0]);
string candidateDirectory = Path.GetFullPath(args[1]);
string outputDirectory = Path.GetFullPath(args[2]);

if (!Directory.Exists(referenceDirectory))
{
    Console.Error.WriteLine($"Reference PNG directory was not found: {referenceDirectory}");
    return 1;
}

if (!Directory.Exists(candidateDirectory))
{
    Console.Error.WriteLine($"Candidate PNG directory was not found: {candidateDirectory}");
    return 1;
}

Directory.CreateDirectory(outputDirectory);

string[] referenceFiles = Directory.GetFiles(referenceDirectory, "*.png").Order(StringComparer.OrdinalIgnoreCase).ToArray();
string[] candidateFiles = Directory.GetFiles(candidateDirectory, "*.png").Order(StringComparer.OrdinalIgnoreCase).ToArray();
int pageCount = Math.Max(referenceFiles.Length, candidateFiles.Length);

var metrics = new List<PageMetric>();
for (int i = 0; i < pageCount; i++)
{
    string? referenceFile = i < referenceFiles.Length ? referenceFiles[i] : null;
    string? candidateFile = i < candidateFiles.Length ? candidateFiles[i] : null;
    metrics.Add(MeasurePage(i + 1, referenceFile, candidateFile));
}

var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
File.WriteAllText(Path.Combine(outputDirectory, "metrics.json"), JsonSerializer.Serialize(metrics, jsonOptions));
File.WriteAllText(Path.Combine(outputDirectory, "index.html"), BuildIndexHtml(metrics, referenceDirectory, candidateDirectory, outputDirectory));

Console.WriteLine($"Wrote {metrics.Count} visual comparison entries to {outputDirectory}");
return 0;

static string BuildIndexHtml(IReadOnlyList<PageMetric> metrics, string referenceDirectory, string candidateDirectory, string outputDirectory)
{
    var html = new StringBuilder();
    html.AppendLine("<!doctype html>");
    html.AppendLine("<html lang=\"en\">");
    html.AppendLine("<head><meta charset=\"utf-8\"><title>Lokad.OoxPdf VisualDiff</title>");
    html.AppendLine("<style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px}section{margin-bottom:32px}.pair{display:grid;grid-template-columns:1fr 1fr;gap:16px}img{max-width:100%;border:1px solid #bbb}h2{font-size:18px}</style>");
    html.AppendLine("</head><body><h1>Lokad.OoxPdf VisualDiff</h1>");

    foreach (PageMetric metric in metrics)
    {
        html.AppendLine($"<section><h2>Page {metric.Page}</h2><div class=\"pair\">");
        AppendImage(html, "Reference", metric.ReferenceFile, referenceDirectory, outputDirectory);
        AppendImage(html, "Candidate", metric.CandidateFile, candidateDirectory, outputDirectory);
        html.AppendLine("</div></section>");
    }

    html.AppendLine("</body></html>");
    return html.ToString();
}

static void AppendImage(StringBuilder html, string title, string? fileName, string sourceDirectory, string outputDirectory)
{
    html.Append("<div><h3>");
    html.Append(HtmlEncoder.Default.Encode(title));
    html.AppendLine("</h3>");

    if (fileName is null)
    {
        html.AppendLine("<p>Missing</p></div>");
        return;
    }

    string absolutePath = Path.Combine(sourceDirectory, fileName);
    string relativePath = Path.GetRelativePath(outputDirectory, absolutePath).Replace('\\', '/');
    html.Append("<img src=\"");
    html.Append(HtmlEncoder.Default.Encode(relativePath));
    html.Append("\" alt=\"");
    html.Append(HtmlEncoder.Default.Encode($"{title} {fileName}"));
    html.AppendLine("\"></div>");
}

static PageMetric MeasurePage(int page, string? referenceFile, string? candidateFile)
{
    PngImage? reference = referenceFile is null ? null : PngImage.Load(referenceFile);
    PngImage? candidate = candidateFile is null ? null : PngImage.Load(candidateFile);
    bool dimensionsMatch = reference is not null
        && candidate is not null
        && reference.Width == candidate.Width
        && reference.Height == candidate.Height;
    bool dimensionsNearMatch = reference is not null
        && candidate is not null
        && Math.Abs(reference.Width - candidate.Width) <= 2
        && Math.Abs(reference.Height - candidate.Height) <= 2;

    PixelMetric? pixelMetric = null;
    if ((dimensionsMatch || dimensionsNearMatch) && reference is not null && candidate is not null)
    {
        pixelMetric = MeasurePixels(reference, candidate, Math.Min(reference.Width, candidate.Width), Math.Min(reference.Height, candidate.Height));
    }

    return new PageMetric(
        Page: page,
        ReferenceFile: referenceFile is null ? null : Path.GetFileName(referenceFile),
        CandidateFile: candidateFile is null ? null : Path.GetFileName(candidateFile),
        ReferenceWidth: reference?.Width,
        ReferenceHeight: reference?.Height,
        CandidateWidth: candidate?.Width,
        CandidateHeight: candidate?.Height,
        MeanAbsoluteError: pixelMetric?.MeanAbsoluteError,
        RootMeanSquaredError: pixelMetric?.RootMeanSquaredError,
        ChangedPixelRatioAtThreshold16: pixelMetric?.ChangedPixelRatioAtThreshold16,
        ChangedPixelRatioAtThreshold32: pixelMetric?.ChangedPixelRatioAtThreshold32,
        StructuralSimilarity: pixelMetric?.StructuralSimilarity,
        ForegroundColorHistogramCorrelation: pixelMetric?.ForegroundColorHistogramCorrelation,
        DimensionsMatch: dimensionsMatch);
}

static PixelMetric MeasurePixels(PngImage reference, PngImage candidate, int width, int height)
{
    long absoluteError = 0;
    double squaredError = 0;
    int changed16 = 0;
    int changed32 = 0;
    int pixelCount = width * height;
    double referenceLumaSum = 0d;
    double candidateLumaSum = 0d;
    double referenceLumaSquaredSum = 0d;
    double candidateLumaSquaredSum = 0d;
    double lumaProductSum = 0d;
    int[] referenceHistogram = new int[96];
    int[] candidateHistogram = new int[96];
    int referenceForeground = 0;
    int candidateForeground = 0;

    for (int y = 0; y < height; y++)
    {
        for (int x = 0; x < width; x++)
        {
            int referenceOffset = (y * reference.Width + x) * 4;
            int candidateOffset = (y * candidate.Width + x) * 4;
            int maxChannelDelta = 0;
            int referenceRed = reference.Rgba[referenceOffset];
            int referenceGreen = reference.Rgba[referenceOffset + 1];
            int referenceBlue = reference.Rgba[referenceOffset + 2];
            int candidateRed = candidate.Rgba[candidateOffset];
            int candidateGreen = candidate.Rgba[candidateOffset + 1];
            int candidateBlue = candidate.Rgba[candidateOffset + 2];
            double referenceLuma = Luma(referenceRed, referenceGreen, referenceBlue);
            double candidateLuma = Luma(candidateRed, candidateGreen, candidateBlue);
            referenceLumaSum += referenceLuma;
            candidateLumaSum += candidateLuma;
            referenceLumaSquaredSum += referenceLuma * referenceLuma;
            candidateLumaSquaredSum += candidateLuma * candidateLuma;
            lumaProductSum += referenceLuma * candidateLuma;
            if (referenceLuma < 245d)
            {
                AddHistogram(referenceHistogram, referenceRed, referenceGreen, referenceBlue);
                referenceForeground++;
            }

            if (candidateLuma < 245d)
            {
                AddHistogram(candidateHistogram, candidateRed, candidateGreen, candidateBlue);
                candidateForeground++;
            }

            for (int channel = 0; channel < 4; channel++)
            {
                int delta = Math.Abs(reference.Rgba[referenceOffset + channel] - candidate.Rgba[candidateOffset + channel]);
                absoluteError += delta;
                squaredError += delta * delta;
                maxChannelDelta = Math.Max(maxChannelDelta, delta);
            }

            if (maxChannelDelta >= 16)
            {
                changed16++;
            }

            if (maxChannelDelta >= 32)
            {
                changed32++;
            }
        }
    }

    double sampleCount = pixelCount * 4d;
    return new PixelMetric(
        absoluteError / sampleCount,
        Math.Sqrt(squaredError / sampleCount),
        changed16 / (double)pixelCount,
        changed32 / (double)pixelCount,
        ComputeStructuralSimilarity(pixelCount, referenceLumaSum, candidateLumaSum, referenceLumaSquaredSum, candidateLumaSquaredSum, lumaProductSum),
        ComputeHistogramCorrelation(referenceHistogram, candidateHistogram, referenceForeground, candidateForeground, pixelCount));
}

static double Luma(int red, int green, int blue)
{
    return 0.2126d * red + 0.7152d * green + 0.0722d * blue;
}

static void AddHistogram(int[] histogram, int red, int green, int blue)
{
    histogram[red * 32 / 256]++;
    histogram[32 + green * 32 / 256]++;
    histogram[64 + blue * 32 / 256]++;
}

static double ComputeStructuralSimilarity(
    int sampleCount,
    double referenceSum,
    double candidateSum,
    double referenceSquaredSum,
    double candidateSquaredSum,
    double productSum)
{
    if (sampleCount <= 1)
    {
        return 1d;
    }

    double referenceMean = referenceSum / sampleCount;
    double candidateMean = candidateSum / sampleCount;
    double referenceVariance = referenceSquaredSum / sampleCount - referenceMean * referenceMean;
    double candidateVariance = candidateSquaredSum / sampleCount - candidateMean * candidateMean;
    double covariance = productSum / sampleCount - referenceMean * candidateMean;
    const double c1 = 6.5025d;
    const double c2 = 58.5225d;
    double numerator = (2d * referenceMean * candidateMean + c1) * (2d * covariance + c2);
    double denominator = (referenceMean * referenceMean + candidateMean * candidateMean + c1) * (referenceVariance + candidateVariance + c2);
    return denominator <= 0d ? 1d : Math.Clamp(numerator / denominator, -1d, 1d);
}

static double ComputeHistogramCorrelation(int[] reference, int[] candidate, int referenceForeground, int candidateForeground, int pixelCount)
{
    if (referenceForeground == 0 && candidateForeground == 0)
    {
        return 1d;
    }

    if (referenceForeground == 0 || candidateForeground == 0)
    {
        return 0d;
    }

    if (Math.Min(referenceForeground, candidateForeground) < pixelCount * 0.015d)
    {
        return 1d;
    }

    double referenceMean = reference.Average();
    double candidateMean = candidate.Average();
    double numerator = 0d;
    double referenceDenominator = 0d;
    double candidateDenominator = 0d;
    for (int i = 0; i < reference.Length; i++)
    {
        double rd = reference[i] - referenceMean;
        double cd = candidate[i] - candidateMean;
        numerator += rd * cd;
        referenceDenominator += rd * rd;
        candidateDenominator += cd * cd;
    }

    double denominator = Math.Sqrt(referenceDenominator * candidateDenominator);
    return denominator <= 0d ? 1d : Math.Clamp(numerator / denominator, -1d, 1d);
}
