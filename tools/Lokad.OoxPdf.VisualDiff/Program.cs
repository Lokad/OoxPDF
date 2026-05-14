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

    PixelMetric? pixelMetric = dimensionsMatch ? MeasurePixels(reference!, candidate!) : null;

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
        DimensionsMatch: dimensionsMatch);
}

static PixelMetric MeasurePixels(PngImage reference, PngImage candidate)
{
    long absoluteError = 0;
    double squaredError = 0;
    int changed16 = 0;
    int changed32 = 0;
    int pixelCount = reference.Width * reference.Height;

    for (int pixel = 0; pixel < pixelCount; pixel++)
    {
        int offset = pixel * 4;
        int maxChannelDelta = 0;
        for (int channel = 0; channel < 4; channel++)
        {
            int delta = Math.Abs(reference.Rgba[offset + channel] - candidate.Rgba[offset + channel]);
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

    double sampleCount = pixelCount * 4d;
    return new PixelMetric(
        absoluteError / sampleCount,
        Math.Sqrt(squaredError / sampleCount),
        changed16 / (double)pixelCount,
        changed32 / (double)pixelCount);
}
