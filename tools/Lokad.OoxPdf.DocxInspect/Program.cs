using System.Text.Encodings.Web;
using System.Text.Json;

using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.DocxInspect <input.docx> <output-directory>");
    Environment.Exit(2);
}

string inputPath = Path.GetFullPath(args[0]);
string outputDirectory = Path.GetFullPath(args[1]);
Directory.CreateDirectory(outputDirectory);

using FileStream stream = File.OpenRead(inputPath);
OoxPackage package = OoxPackage.Open(stream);
DocxDocument document = new DocxReader().Read(package);
var renderer = new DocxRenderer();

var options = new JsonSerializerOptions
{
    WriteIndented = true,
    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
};

DocxLayoutSnapshot layout = renderer.InspectLayout(document);
DocxFontPlanSnapshot fontPlan = renderer.InspectFontPlan(document);
DocxStructureSnapshot structure = renderer.InspectStructure(document);
DocxTextEmissionSnapshot textEmission = renderer.InspectTextEmission(document);
File.WriteAllText(
    Path.Combine(outputDirectory, "layout-snapshot.json"),
    JsonSerializer.Serialize(layout, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "font-plan-snapshot.json"),
    JsonSerializer.Serialize(fontPlan, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "structure-snapshot.json"),
    JsonSerializer.Serialize(structure, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "text-emission-snapshot.json"),
    JsonSerializer.Serialize(textEmission, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "document-settings.json"),
    JsonSerializer.Serialize(document.Settings, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "page-summary.json"),
    JsonSerializer.Serialize(layout.Pages.Select((page, index) => new
    {
        Page = index + 1,
        page.Width,
        page.Height,
        page.ItemCount,
        page.TextLineCount,
        page.InlineImageCount,
        page.TableRowCount,
        page.SourceBlockCount,
        page.FirstSourceBlockIndex,
        page.LastSourceBlockIndex
    }), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "source-block-summary.json"),
    JsonSerializer.Serialize(layout.SourceBlocks, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "text-emission-summary.json"),
    JsonSerializer.Serialize(new
    {
        textEmission.LineCount,
        textEmission.SegmentCount,
        textEmission.TerminalSpaceSegmentCount,
        textEmission.NonzeroPdfCharacterSpacingSegmentCount,
        textEmission.CompensatedCharacterSpacingSegmentCount,
        CharacterProfile = SumCharacterProfiles(textEmission.Lines.SelectMany(line => line.Segments)),
        AdvanceProfile = SumAdvanceProfiles(textEmission.Lines.SelectMany(line => line.Segments)),
        LinesByPage = textEmission.Lines
            .GroupBy(line => line.PageIndex)
            .OrderBy(group => group.Key)
            .Select(group => new
            {
                PageIndex = group.Key,
                LineCount = group.Count(),
                StaticLineCount = group.Count(line => line.IsStaticStory),
                BodyLineCount = group.Count(line => !line.IsStaticStory),
                SegmentCount = group.Sum(line => line.SegmentCount),
                TerminalSpaceSegmentCount = group.Sum(line => line.TerminalSpaceSegmentCount),
                NonzeroPdfCharacterSpacingSegmentCount = group.Sum(line => line.NonzeroPdfCharacterSpacingSegmentCount),
                CharacterProfile = SumCharacterProfiles(group.SelectMany(line => line.Segments)),
                AdvanceProfile = SumAdvanceProfiles(group.SelectMany(line => line.Segments)),
                SourceBlockCount = group
                    .Select(line => line.SourceBlockIndex)
                    .Where(index => index is not null)
                    .Distinct()
                    .Count(),
                FirstSourceBlockIndex = group
                    .Select(line => line.SourceBlockIndex)
                    .Where(index => index is not null)
                    .FirstOrDefault(),
                LastSourceBlockIndex = group
                    .Select(line => line.SourceBlockIndex)
                    .Where(index => index is not null)
                    .LastOrDefault()
            })
            .ToArray()
    }, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "block-sequence.json"),
    JsonSerializer.Serialize(structure.Blocks, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "table-adjacency-summary.json"),
    JsonSerializer.Serialize(structure.TableAdjacency, options));

static DocxTextEmissionCharacterProfile SumCharacterProfiles(IEnumerable<DocxTextEmissionSegmentSnapshot> segments)
{
    int digitCount = 0;
    int letterCount = 0;
    int whitespaceCount = 0;
    int punctuationCount = 0;
    int symbolCount = 0;
    int otherCount = 0;
    foreach (DocxTextEmissionSegmentSnapshot segment in segments)
    {
        digitCount += segment.CharacterProfile.DigitCount;
        letterCount += segment.CharacterProfile.LetterCount;
        whitespaceCount += segment.CharacterProfile.WhitespaceCount;
        punctuationCount += segment.CharacterProfile.PunctuationCount;
        symbolCount += segment.CharacterProfile.SymbolCount;
        otherCount += segment.CharacterProfile.OtherCount;
    }

    return new(digitCount, letterCount, whitespaceCount, punctuationCount, symbolCount, otherCount);
}

static object SumAdvanceProfiles(IEnumerable<DocxTextEmissionSegmentSnapshot> segments)
{
    int glyphCount = 0;
    int glyphGapCount = 0;
    double naturalPdfWidth = 0d;
    double layoutWidth = 0d;
    double residual = 0d;
    foreach (DocxTextEmissionSegmentSnapshot segment in segments)
    {
        glyphCount += segment.AdvanceProfile.GlyphCount;
        glyphGapCount += segment.AdvanceProfile.GlyphGapCount;
        naturalPdfWidth += segment.AdvanceProfile.NaturalPdfWidth;
        layoutWidth += segment.AdvanceProfile.LayoutWidth;
        residual += segment.AdvanceProfile.LayoutToNaturalResidual;
    }

    return new
    {
        GlyphCount = glyphCount,
        GlyphGapCount = glyphGapCount,
        NaturalPdfWidth = Math.Round(naturalPdfWidth, 6),
        LayoutWidth = Math.Round(layoutWidth, 6),
        LayoutToNaturalResidual = Math.Round(residual, 6),
        UniformResidualPerGap = glyphGapCount == 0 ? (double?)null : Math.Round(residual / glyphGapCount, 6)
    };
}
