using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.PptxInspect <input.pptx> <output-directory> [--slide <number>]... [--include-text]");
    return 2;
}

string inputPath = Path.GetFullPath(args[0]);
string outputDirectory = Path.GetFullPath(args[1]);
bool includeText = args.Any(arg => string.Equals(arg, "--include-text", StringComparison.Ordinal));
HashSet<int>? slideFilter = ReadSlideFilter(args);

Directory.CreateDirectory(outputDirectory);

using FileStream stream = File.OpenRead(inputPath);
OoxPackage package = OoxPackage.Open(stream);
PptxDocument document = new PptxReader().Read(package);
IEnumerable<PptxSlide> slides = document.Slides;
if (slideFilter is not null)
{
    slides = slides.Where(slide => slideFilter.Contains(slide.Index + 1));
}

var records = new List<PptxGlyphRunRecord>();
foreach (PptxSlide slide in slides)
{
    int slideNumber = slide.Index + 1;
    foreach (PptxTextGlyphRunSnapshot run in PptxRenderer.InspectTextGlyphRuns(document, package, slide.Index))
    {
        records.Add(new PptxGlyphRunRecord(
            slideNumber,
            includeText ? run.Text : null,
            run.Text.Length,
            Round(run.X),
            Round(run.BaselineY),
            Round(run.Width),
            run.FrameIndex,
            run.ParagraphIndex,
            run.LineIndex,
            Round(run.FrameFontScale),
            Round(run.FrameTextX),
            Round(run.FrameTextWidth),
            Round(run.FrameTextWrapWidth),
            Round(run.FrameTextHeight),
            Round(run.FrameClipX),
            Round(run.FrameClipWidth),
            Round(run.FrameClipY),
            Round(run.FrameClipHeight),
            run.FrameColumnCount,
            Round(run.FrameColumnSpacing),
            Round(run.LineTopY),
            Round(run.LineAdvance),
            Round(run.LineMaxFontSize),
            Round(run.LayoutFontSize),
            Round(run.PdfFontSize),
            run.GlyphCount,
            Round(run.FirstAdjustmentAfterOrigin)));
    }
}

var options = new JsonSerializerOptions
{
    WriteIndented = true,
    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
};
string outputPath = Path.Combine(outputDirectory, "glyph-runs.json");
File.WriteAllText(outputPath, JsonSerializer.Serialize(records, options), Encoding.UTF8);

Console.WriteLine(FormattableString.Invariant($"PPTX: {inputPath}"));
Console.WriteLine(FormattableString.Invariant($"Slides: {document.Slides.Count}"));
Console.WriteLine(FormattableString.Invariant($"Glyph runs: {records.Count}"));
Console.WriteLine(FormattableString.Invariant($"Output: {outputPath}"));

return 0;

static HashSet<int>? ReadSlideFilter(string[] args)
{
    var slides = new HashSet<int>();
    for (int i = 2; i < args.Length; i++)
    {
        if (!string.Equals(args[i], "--slide", StringComparison.Ordinal))
        {
            continue;
        }

        if (i + 1 >= args.Length ||
            !int.TryParse(args[i + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int slide) ||
            slide <= 0)
        {
            throw new ArgumentException("--slide expects a positive one-based slide number.");
        }

        slides.Add(slide);
        i++;
    }

    return slides.Count == 0 ? null : slides;
}

static double Round(double value)
{
    return Math.Round(value, 6, MidpointRounding.AwayFromZero);
}

internal sealed record PptxGlyphRunRecord(
    int Slide,
    string? Text,
    int TextLength,
    double X,
    double BaselineY,
    double Width,
    int FrameIndex,
    int ParagraphIndex,
    int LineIndex,
    double FrameFontScale,
    double FrameTextX,
    double FrameTextWidth,
    double FrameTextWrapWidth,
    double FrameTextHeight,
    double FrameClipX,
    double FrameClipWidth,
    double FrameClipY,
    double FrameClipHeight,
    int FrameColumnCount,
    double FrameColumnSpacing,
    double LineTopY,
    double LineAdvance,
    double LineMaxFontSize,
    double LayoutFontSize,
    double PdfFontSize,
    int GlyphCount,
    double FirstAdjustmentAfterOrigin);
