using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;

if (args.Length < 1)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.PdfInspect <input.pdf> [output-directory] [--text-only] [--page <number>]...");
    return 2;
}

string inputPath = Path.GetFullPath(args[0]);
bool textOnly = args.Any(arg => string.Equals(arg, "--text-only", StringComparison.Ordinal));
HashSet<int>? pageFilter = ReadPageFilter(args);
string? outputDirectory = ReadOutputDirectory(args);
outputDirectory = outputDirectory is null ? null : Path.GetFullPath(outputDirectory);
if (outputDirectory is not null)
{
    Directory.CreateDirectory(outputDirectory);
}

byte[] bytes = File.ReadAllBytes(inputPath);
string pdf = Encoding.Latin1.GetString(bytes);
var objects = PdfObject.ParseAll(pdf, bytes, skipImageDecode: textOnly);
Dictionary<int, int> contentPageNumbers = BuildContentPageMap(objects);
Dictionary<string, IReadOnlyDictionary<int, string>> fontUnicodeMaps = BuildFontUnicodeMaps(objects);
Dictionary<string, PdfFontWidthMap> fontWidthMaps = BuildFontWidthMaps(objects);
HashSet<int>? filteredContentObjects = pageFilter is null
    ? null
    : contentPageNumbers
        .Where(pair => pageFilter.Contains(pair.Value))
        .Select(pair => pair.Key)
        .ToHashSet();
var textOperations = new List<PdfTextOperation>();
var graphicsOperations = new List<PdfGraphicsOperation>();

Console.WriteLine(FormattableString.Invariant($"PDF: {inputPath}"));
Console.WriteLine(FormattableString.Invariant($"Objects: {objects.Count}"));
if (pageFilter is not null)
{
    Console.WriteLine(FormattableString.Invariant($"Page filter: {string.Join(",", pageFilter.Order())}"));
}

foreach (PdfObject item in objects)
{
    string streamNote = item.Stream is null
        ? string.Empty
        : FormattableString.Invariant($", stream {item.Stream.RawLength} bytes, decoded {item.Stream.DecodedLength} bytes, filters: {item.Stream.Filters}, {item.Stream.DecodeStatus}");
    Console.WriteLine(FormattableString.Invariant($"{item.Number} {item.Generation} obj: {Classify(item.Body)}{streamNote}"));

    if (outputDirectory is null || item.Stream is null)
    {
        continue;
    }

    if (filteredContentObjects is not null && !filteredContentObjects.Contains(item.Number))
    {
        continue;
    }

    string prefix = Path.Combine(outputDirectory, FormattableString.Invariant($"obj-{item.Number:0000}-{item.Generation}"));
    if (!textOnly)
    {
        File.WriteAllText(prefix + ".dict.txt", item.Dictionary, Encoding.UTF8);
    }

    if (LooksTextual(item.Stream.Decoded))
    {
        string text = Encoding.Latin1.GetString(item.Stream.Decoded);
        if (!textOnly)
        {
            File.WriteAllText(prefix + ".stream.txt", text, Encoding.UTF8);
        }

        int? pageNumber = contentPageNumbers.TryGetValue(item.Number, out int page) ? page : null;
        textOperations.AddRange(PdfTextOperation.Extract(pageNumber, item.Number, item.Generation, text, fontUnicodeMaps, fontWidthMaps));
        graphicsOperations.AddRange(PdfGraphicsOperation.Extract(pageNumber, item.Number, item.Generation, text));
    }
    else if (!textOnly)
    {
        File.WriteAllBytes(prefix + ".stream.bin", item.Stream.Decoded);
    }
}

if (outputDirectory is not null && textOperations.Count != 0)
{
    string textOperationsPath = Path.Combine(outputDirectory, "text-operations.json");
    var jsonOptions = new JsonSerializerOptions
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };
    File.WriteAllText(textOperationsPath, JsonSerializer.Serialize(textOperations, jsonOptions), Encoding.UTF8);
}

if (outputDirectory is not null && graphicsOperations.Count != 0)
{
    string graphicsOperationsPath = Path.Combine(outputDirectory, "graphics-operations.json");
    var jsonOptions = new JsonSerializerOptions
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };
    File.WriteAllText(graphicsOperationsPath, JsonSerializer.Serialize(graphicsOperations, jsonOptions), Encoding.UTF8);
}

return 0;

static HashSet<int>? ReadPageFilter(string[] args)
{
    var pages = new HashSet<int>();
    for (int i = 1; i < args.Length; i++)
    {
        if (!string.Equals(args[i], "--page", StringComparison.Ordinal))
        {
            continue;
        }

        if (i + 1 >= args.Length ||
            !int.TryParse(args[i + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int page) ||
            page <= 0)
        {
            throw new ArgumentException("--page expects a positive one-based page number.");
        }

        pages.Add(page);
        i++;
    }

    return pages.Count == 0 ? null : pages;
}

static string? ReadOutputDirectory(string[] args)
{
    string? outputDirectory = null;
    for (int i = 1; i < args.Length; i++)
    {
        string arg = args[i];
        if (string.Equals(arg, "--text-only", StringComparison.Ordinal))
        {
            continue;
        }

        if (string.Equals(arg, "--page", StringComparison.Ordinal))
        {
            i++;
            continue;
        }

        if (arg.StartsWith("--", StringComparison.Ordinal))
        {
            throw new ArgumentException($"Unknown option '{arg}'.");
        }

        if (outputDirectory is not null)
        {
            throw new ArgumentException("Only one output directory can be specified.");
        }

        outputDirectory = arg;
    }

    return outputDirectory;
}

static string Classify(string body)
{
    if (IsPageObject(body))
    {
        return "page";
    }

    if (body.Contains("/Type /Pages", StringComparison.Ordinal))
    {
        return "pages";
    }

    if (body.Contains("/Type /Font", StringComparison.Ordinal) || body.Contains("/Subtype /Type0", StringComparison.Ordinal))
    {
        return "font";
    }

    if (body.Contains("/Subtype /Image", StringComparison.Ordinal))
    {
        return "image";
    }

    if (body.Contains(" stream", StringComparison.Ordinal))
    {
        return "stream";
    }

    return "object";
}

static bool IsPageObject(string body)
{
    return Regex.IsMatch(body, @"/Type\s*/Page\b", RegexOptions.CultureInvariant) &&
        !Regex.IsMatch(body, @"/Type\s*/Pages\b", RegexOptions.CultureInvariant);
}

static Dictionary<int, int> BuildContentPageMap(IReadOnlyList<PdfObject> objects)
{
    var pageByContentObject = new Dictionary<int, int>();
    int pageNumber = 0;
    foreach (PdfObject item in objects)
    {
        if (!IsPageObject(item.Body))
        {
            continue;
        }

        pageNumber++;
        foreach (int contentObjectNumber in ReadContentObjectNumbers(item.Body))
        {
            pageByContentObject[contentObjectNumber] = pageNumber;
        }
    }

    return pageByContentObject;
}

static IReadOnlyList<int> ReadContentObjectNumbers(string body)
{
    Match contents = Regex.Match(body, @"(?s)/Contents\s*(?<value>\[(?<array>.*?)\]|(?<single>\d+\s+\d+\s+R))", RegexOptions.CultureInvariant);
    if (!contents.Success)
    {
        return Array.Empty<int>();
    }

    string value = contents.Groups["array"].Success
        ? contents.Groups["array"].Value
        : contents.Groups["single"].Value;
    return Regex.Matches(value, @"(?<number>\d+)\s+\d+\s+R", RegexOptions.CultureInvariant)
        .Select(match => int.Parse(match.Groups["number"].Value, CultureInfo.InvariantCulture))
        .ToArray();
}

static Dictionary<string, IReadOnlyDictionary<int, string>> BuildFontUnicodeMaps(IReadOnlyList<PdfObject> objects)
{
    var objectByNumber = objects.ToDictionary(item => item.Number);
    var maps = new Dictionary<string, IReadOnlyDictionary<int, string>>(StringComparer.Ordinal);
    foreach (PdfObject page in objects.Where(item => IsPageObject(item.Body)))
    {
        foreach ((string fontName, int fontObjectNumber) in ReadPageFontResources(page.Body))
        {
            if (!objectByNumber.TryGetValue(fontObjectNumber, out PdfObject? fontObject))
            {
                continue;
            }

            Match toUnicode = Regex.Match(fontObject.Body, @"/ToUnicode\s+(?<number>\d+)\s+\d+\s+R", RegexOptions.CultureInvariant);
            if (!toUnicode.Success ||
                !objectByNumber.TryGetValue(int.Parse(toUnicode.Groups["number"].Value, CultureInfo.InvariantCulture), out PdfObject? cmapObject) ||
                cmapObject.Stream is null)
            {
                continue;
            }

            string cmap = Encoding.Latin1.GetString(cmapObject.Stream.Decoded);
            maps[fontName] = PdfToUnicodeMap.Parse(cmap);
        }
    }

    return maps;
}

static Dictionary<string, PdfFontWidthMap> BuildFontWidthMaps(IReadOnlyList<PdfObject> objects)
{
    var objectByNumber = objects.ToDictionary(item => item.Number);
    var maps = new Dictionary<string, PdfFontWidthMap>(StringComparer.Ordinal);
    foreach (PdfObject page in objects.Where(item => IsPageObject(item.Body)))
    {
        foreach ((string fontName, int fontObjectNumber) in ReadPageFontResources(page.Body))
        {
            if (!objectByNumber.TryGetValue(fontObjectNumber, out PdfObject? fontObject))
            {
                continue;
            }

            PdfFontWidthMap? widthMap = ReadType0WidthMap(fontObject, objectByNumber) ??
                ReadSimpleWidthMap(fontObject, objectByNumber);
            if (widthMap is not null)
            {
                maps[fontName] = widthMap;
            }
        }
    }

    return maps;
}

static PdfFontWidthMap? ReadType0WidthMap(PdfObject fontObject, IReadOnlyDictionary<int, PdfObject> objectByNumber)
{
    Match descendantFonts = Regex.Match(fontObject.Body, @"/DescendantFonts\s+(?<number>\d+)\s+\d+\s+R", RegexOptions.CultureInvariant);
    if (!descendantFonts.Success ||
        !objectByNumber.TryGetValue(int.Parse(descendantFonts.Groups["number"].Value, CultureInfo.InvariantCulture), out PdfObject? descendantArray))
    {
        return null;
    }

    Match descendantRef = Regex.Match(descendantArray.Body, @"\[\s*(?<number>\d+)\s+\d+\s+R\s*\]", RegexOptions.CultureInvariant);
    if (!descendantRef.Success ||
        !objectByNumber.TryGetValue(int.Parse(descendantRef.Groups["number"].Value, CultureInfo.InvariantCulture), out PdfObject? descendant))
    {
        return null;
    }

    int defaultWidth = ReadDictionaryInt(descendant.Body, "DW");
    if (defaultWidth <= 0)
    {
        defaultWidth = 1000;
    }

    string widths = ReadReferencedArray(descendant.Body, "W", objectByNumber);
    return new PdfFontWidthMap(defaultWidth, ParseCidWidths(widths));
}

static PdfFontWidthMap? ReadSimpleWidthMap(PdfObject fontObject, IReadOnlyDictionary<int, PdfObject> objectByNumber)
{
    int firstChar = ReadDictionaryInt(fontObject.Body, "FirstChar");
    string widths = ReadReferencedArray(fontObject.Body, "Widths", objectByNumber);
    if (string.IsNullOrWhiteSpace(widths))
    {
        return null;
    }

    int code = firstChar;
    var map = new Dictionary<int, int>();
    foreach (Match match in Regex.Matches(widths, @"-?\d+", RegexOptions.CultureInvariant))
    {
        map[code++] = int.Parse(match.Value, CultureInfo.InvariantCulture);
    }

    return new PdfFontWidthMap(0, map);
}

static string ReadReferencedArray(string body, string key, IReadOnlyDictionary<int, PdfObject> objectByNumber)
{
    Match reference = Regex.Match(body, @"/" + Regex.Escape(key) + @"\s+(?<number>\d+)\s+\d+\s+R", RegexOptions.CultureInvariant);
    if (reference.Success &&
        objectByNumber.TryGetValue(int.Parse(reference.Groups["number"].Value, CultureInfo.InvariantCulture), out PdfObject? referenced))
    {
        return referenced.Body;
    }

    Match inline = Regex.Match(body, @"/" + Regex.Escape(key) + @"\s*(?<array>\[(?:[^\[\]]|\[(?:[^\[\]]|\[[^\[\]]*\])*\])*\])", RegexOptions.CultureInvariant);
    return inline.Success ? inline.Groups["array"].Value : string.Empty;
}

static IReadOnlyDictionary<int, int> ParseCidWidths(string widths)
{
    var map = new Dictionary<int, int>();
    foreach (Match match in Regex.Matches(widths, @"(?<start>\d+)\s*\[(?<values>[^\]]*)\]|(?<startRange>\d+)\s+(?<endRange>\d+)\s+(?<width>\d+)", RegexOptions.CultureInvariant))
    {
        if (match.Groups["values"].Success)
        {
            int code = int.Parse(match.Groups["start"].Value, CultureInfo.InvariantCulture);
            foreach (Match width in Regex.Matches(match.Groups["values"].Value, @"-?\d+", RegexOptions.CultureInvariant))
            {
                map[code++] = int.Parse(width.Value, CultureInfo.InvariantCulture);
            }
        }
        else
        {
            int start = int.Parse(match.Groups["startRange"].Value, CultureInfo.InvariantCulture);
            int end = int.Parse(match.Groups["endRange"].Value, CultureInfo.InvariantCulture);
            int width = int.Parse(match.Groups["width"].Value, CultureInfo.InvariantCulture);
            for (int code = start; code <= end; code++)
            {
                map[code] = width;
            }
        }
    }

    return map;
}

static int ReadDictionaryInt(string dictionary, string key)
{
    Match match = Regex.Match(dictionary, @"/" + Regex.Escape(key) + @"\s+(?<value>\d+)", RegexOptions.CultureInvariant);
    return match.Success ? int.Parse(match.Groups["value"].Value, CultureInfo.InvariantCulture) : 0;
}

static IReadOnlyList<(string FontName, int ObjectNumber)> ReadPageFontResources(string pageBody)
{
    Match fonts = Regex.Match(pageBody, @"(?s)/Font\s*<<(?<fonts>.*?)>>", RegexOptions.CultureInvariant);
    if (!fonts.Success)
    {
        return Array.Empty<(string FontName, int ObjectNumber)>();
    }

    return Regex.Matches(fonts.Groups["fonts"].Value, @"/(?<name>[A-Za-z0-9._+-]+)\s+(?<number>\d+)\s+\d+\s+R", RegexOptions.CultureInvariant)
        .Select(match => (
            match.Groups["name"].Value,
            int.Parse(match.Groups["number"].Value, CultureInfo.InvariantCulture)))
        .ToArray();
}

static bool LooksTextual(byte[] bytes)
{
    if (bytes.Length == 0)
    {
        return true;
    }

    int printable = 0;
    foreach (byte value in bytes)
    {
        if (value is 9 or 10 or 12 or 13 || value is >= 32 and <= 126)
        {
            printable++;
        }
    }

    return printable >= bytes.Length * 0.85;
}

internal sealed record PdfObject(int Number, int Generation, string Body, string Dictionary, PdfStream? Stream)
{
    private static readonly Regex ObjectRegex = new(
        @"(?s)(?<number>\d+)\s+(?<generation>\d+)\s+obj(?<body>.*?)endobj",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static IReadOnlyList<PdfObject> ParseAll(string pdf, byte[] bytes, bool skipImageDecode = false)
    {
        var objects = new List<PdfObject>();
        foreach (Match match in ObjectRegex.Matches(pdf))
        {
            int number = int.Parse(match.Groups["number"].Value, CultureInfo.InvariantCulture);
            int generation = int.Parse(match.Groups["generation"].Value, CultureInfo.InvariantCulture);
            string rawBody = match.Groups["body"].Value;
            string body = rawBody.Trim();
            string dictionary = ReadDictionary(body);
            PdfStream? stream = TryReadStream(match, rawBody, bytes, dictionary, skipImageDecode);
            objects.Add(new PdfObject(number, generation, body, dictionary, stream));
        }

        objects.AddRange(ParseObjectStreams(objects));
        return objects;
    }

    private static IReadOnlyList<PdfObject> ParseObjectStreams(IReadOnlyList<PdfObject> objects)
    {
        var existing = objects.Select(item => item.Number).ToHashSet();
        var embeddedObjects = new List<PdfObject>();
        foreach (PdfObject objectStream in objects)
        {
            if (objectStream.Stream is null ||
                !objectStream.Dictionary.Contains("/Type/ObjStm", StringComparison.Ordinal) &&
                !objectStream.Dictionary.Contains("/Type /ObjStm", StringComparison.Ordinal))
            {
                continue;
            }

            int count = ReadDictionaryInt(objectStream.Dictionary, "N");
            int first = ReadDictionaryInt(objectStream.Dictionary, "First");
            if (count <= 0 || first <= 0 || first >= objectStream.Stream.Decoded.Length)
            {
                continue;
            }

            string decoded = Encoding.Latin1.GetString(objectStream.Stream.Decoded);
            string header = decoded[..first];
            MatchCollection pairs = Regex.Matches(header, @"(?<number>\d+)\s+(?<offset>\d+)", RegexOptions.CultureInvariant);
            var entries = pairs
                .Take(count)
                .Select(match => (
                    Number: int.Parse(match.Groups["number"].Value, CultureInfo.InvariantCulture),
                    Offset: int.Parse(match.Groups["offset"].Value, CultureInfo.InvariantCulture)))
                .ToArray();

            for (int i = 0; i < entries.Length; i++)
            {
                int number = entries[i].Number;
                if (existing.Contains(number))
                {
                    continue;
                }

                int start = first + entries[i].Offset;
                int end = i + 1 < entries.Length ? first + entries[i + 1].Offset : decoded.Length;
                if (start < first || end <= start || end > decoded.Length)
                {
                    continue;
                }

                string body = decoded[start..end].Trim();
                string dictionary = ReadDictionary(body);
                embeddedObjects.Add(new PdfObject(number, 0, body, dictionary, null));
                existing.Add(number);
            }
        }

        return embeddedObjects;
    }

    private static int ReadDictionaryInt(string dictionary, string key)
    {
        Match match = Regex.Match(dictionary, @"/" + Regex.Escape(key) + @"\s+(?<value>\d+)", RegexOptions.CultureInvariant);
        return match.Success ? int.Parse(match.Groups["value"].Value, CultureInfo.InvariantCulture) : 0;
    }

    private static string ReadDictionary(string body)
    {
        int start = body.IndexOf("<<", StringComparison.Ordinal);
        int stream = body.IndexOf("stream", StringComparison.Ordinal);
        int end = stream >= 0
            ? body.LastIndexOf(">>", stream, StringComparison.Ordinal)
            : body.LastIndexOf(">>", StringComparison.Ordinal);
        return start >= 0 && end > start
            ? body.Substring(start, end - start + 2).Trim()
            : string.Empty;
    }

    private static PdfStream? TryReadStream(Match match, string body, byte[] bytes, string dictionary, bool skipImageDecode)
    {
        int bodyStream = body.IndexOf("stream", StringComparison.Ordinal);
        int bodyEndStream = body.LastIndexOf("endstream", StringComparison.Ordinal);
        if (bodyStream < 0 || bodyEndStream < bodyStream)
        {
            return null;
        }

        int streamStart = match.Groups["body"].Index + bodyStream + "stream".Length;
        if (streamStart < bytes.Length && bytes[streamStart] == (byte)'\r')
        {
            streamStart++;
        }

        if (streamStart < bytes.Length && bytes[streamStart] == (byte)'\n')
        {
            streamStart++;
        }

        int streamEnd = match.Groups["body"].Index + bodyEndStream;
        while (streamEnd > streamStart && bytes[streamEnd - 1] is (byte)'\r' or (byte)'\n')
        {
            streamEnd--;
        }

        byte[] raw = bytes[streamStart..streamEnd];
        string filters = ReadFilters(dictionary);
        if (skipImageDecode && dictionary.Contains("/Subtype /Image", StringComparison.Ordinal))
        {
            return new PdfStream(raw.Length, 0, filters, "skipped image", Array.Empty<byte>());
        }

        DecodeResult decoded = filters.Contains("FlateDecode", StringComparison.Ordinal)
            ? TryInflate(raw)
            : new DecodeResult(raw, "not decoded");
        return new PdfStream(raw.Length, decoded.Bytes.Length, filters, decoded.Status, decoded.Bytes);
    }

    private static string ReadFilters(string dictionary)
    {
        Match filter = Regex.Match(dictionary, @"/Filter\s*(?<filter>/[A-Za-z0-9]+|\[[^\]]+\])", RegexOptions.CultureInvariant);
        return filter.Success ? filter.Groups["filter"].Value : "none";
    }

    private static DecodeResult TryInflate(byte[] raw)
    {
        try
        {
            using var input = new MemoryStream(raw);
            using var zlib = new ZLibStream(input, CompressionMode.Decompress);
            using var output = new MemoryStream();
            zlib.CopyTo(output);
            return new DecodeResult(output.ToArray(), "decoded");
        }
        catch (InvalidDataException)
        {
            try
            {
                using var input = new MemoryStream(raw);
                using var deflate = new DeflateStream(input, CompressionMode.Decompress);
                using var output = new MemoryStream();
                deflate.CopyTo(output);
                return new DecodeResult(output.ToArray(), "decoded raw deflate");
            }
            catch (InvalidDataException ex)
            {
                return new DecodeResult(raw, "decode failed: " + ex.Message);
            }
        }
    }
}

internal sealed record DecodeResult(byte[] Bytes, string Status);

internal sealed record PdfStream(int RawLength, int DecodedLength, string Filters, string DecodeStatus, byte[] Decoded);

internal sealed record PdfTextOperation(
    int? PageNumber,
    int ObjectNumber,
    int Generation,
    string Font,
    double FontSize,
    double CharacterSpacing,
    double A,
    double B,
    double C,
    double D,
    double X,
    double Y,
    double EffectiveA,
    double EffectiveB,
    double EffectiveC,
    double EffectiveD,
    double EffectiveX,
    double EffectiveY,
    string Operator,
    string Payload,
    int TextChunkCount,
    int AdjustmentCount,
    double AdjustmentSum,
    double AdjustmentMin,
    double AdjustmentMax,
    double AverageAdjustmentPoints,
    int DecodedRuneCount,
    int CharacterSpacingGapCount,
    double CharacterSpacingGapTotalPoints,
    double AdjustmentTotalPoints,
    double NetSpacingGapTotalPoints,
    double? NaturalWidthPoints,
    double? EmittedAdvancePoints,
    double NetAverageCharacterSpacing,
    string DecodedText)
{
    private static readonly Regex SaveGraphicsStateRegex = new(
        @"(?:^|\s)q(?:\s|$)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex RestoreGraphicsStateRegex = new(
        @"(?:^|\s)Q(?:\s|$)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex CurrentMatrixRegex = new(
        @"(?<a>-?(?:\d+\.?\d*|\.\d+))\s+(?<b>-?(?:\d+\.?\d*|\.\d+))\s+(?<c>-?(?:\d+\.?\d*|\.\d+))\s+(?<d>-?(?:\d+\.?\d*|\.\d+))\s+(?<x>-?(?:\d+\.?\d*|\.\d+))\s+(?<y>-?(?:\d+\.?\d*|\.\d+))\s+cm",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex FontRegex = new(
        @"/(?<font>[A-Za-z0-9._+-]+)\s+(?<size>-?(?:\d+\.?\d*|\.\d+))\s+Tf",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex CharacterSpacingRegex = new(
        @"(?<spacing>-?(?:\d+\.?\d*|\.\d+))\s+Tc",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex MatrixRegex = new(
        @"(?<a>-?(?:\d+\.?\d*|\.\d+))\s+(?<b>-?(?:\d+\.?\d*|\.\d+))\s+(?<c>-?(?:\d+\.?\d*|\.\d+))\s+(?<d>-?(?:\d+\.?\d*|\.\d+))\s+(?<x>-?(?:\d+\.?\d*|\.\d+))\s+(?<y>-?(?:\d+\.?\d*|\.\d+))\s+Tm",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex ShowRegex = new(
        @"(?<payload>\[(?:[^\[\]]|\([^)]*\)|<[^>]*>)*\]|\([^)]*\)|<[^>]*>)\s*(?<operator>TJ|Tj)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex PayloadTokenRegex = new(
        @"\((?:\\.|[^\\)])*\)|<[^>]*>|-?(?:\d+\.?\d*|\.\d+)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static IReadOnlyList<PdfTextOperation> Extract(
        int? pageNumber,
        int objectNumber,
        int generation,
        string stream,
        IReadOnlyDictionary<string, IReadOnlyDictionary<int, string>> fontUnicodeMaps,
        IReadOnlyDictionary<string, PdfFontWidthMap> fontWidthMaps)
    {
        var operations = new List<PdfTextOperation>();
        string font = string.Empty;
        double fontSize = 0d;
        double characterSpacing = 0d;
        double a = 1d;
        double b = 0d;
        double c = 0d;
        double d = 1d;
        double x = 0d;
        double y = 0d;
        PdfMatrix currentMatrix = PdfMatrix.Identity;
        var graphicsStack = new Stack<TextExtractionState>();

        foreach (string line in stream.Split('\n'))
        {
            foreach (Match _ in SaveGraphicsStateRegex.Matches(line))
            {
                graphicsStack.Push(new TextExtractionState(font, fontSize, characterSpacing, currentMatrix));
            }

            foreach (Match match in CurrentMatrixRegex.Matches(line))
            {
                PdfMatrix matrix = new(
                    ReadDouble(match.Groups["a"].Value),
                    ReadDouble(match.Groups["b"].Value),
                    ReadDouble(match.Groups["c"].Value),
                    ReadDouble(match.Groups["d"].Value),
                    ReadDouble(match.Groups["x"].Value),
                    ReadDouble(match.Groups["y"].Value));
                currentMatrix = currentMatrix.Multiply(matrix);
            }

            foreach (Match match in FontRegex.Matches(line))
            {
                font = match.Groups["font"].Value;
                fontSize = ReadDouble(match.Groups["size"].Value);
            }

            foreach (Match match in CharacterSpacingRegex.Matches(line))
            {
                characterSpacing = ReadDouble(match.Groups["spacing"].Value);
            }

            foreach (Match match in MatrixRegex.Matches(line))
            {
                a = ReadDouble(match.Groups["a"].Value);
                b = ReadDouble(match.Groups["b"].Value);
                c = ReadDouble(match.Groups["c"].Value);
                d = ReadDouble(match.Groups["d"].Value);
                x = ReadDouble(match.Groups["x"].Value);
                y = ReadDouble(match.Groups["y"].Value);
            }

            foreach (Match match in ShowRegex.Matches(line))
            {
                PdfMatrix effective = currentMatrix.Multiply(new PdfMatrix(a, b, c, d, x, y));
                string payload = match.Groups["payload"].Value;
                string decodedText = DecodePayload(payload, fontUnicodeMaps.TryGetValue(font, out IReadOnlyDictionary<int, string>? unicodeMap)
                    ? unicodeMap
                    : null);
                PdfTextPayloadProfile payloadProfile = ReadPayloadProfile(payload);
                double averageAdjustmentPoints = payloadProfile.AdjustmentCount == 0
                    ? 0d
                    : -payloadProfile.AdjustmentSum / payloadProfile.AdjustmentCount * fontSize / 1000d;
                int decodedRuneCount = decodedText.EnumerateRunes().Count();
                int characterSpacingGapCount = Math.Max(0, decodedRuneCount - 1);
                double characterSpacingGapTotalPoints = characterSpacing * characterSpacingGapCount;
                double adjustmentTotalPoints = -payloadProfile.AdjustmentSum * fontSize / 1000d;
                double? naturalWidthPoints = fontWidthMaps.TryGetValue(font, out PdfFontWidthMap? widthMap)
                    ? MeasurePayloadWidth(payload, fontUnicodeMaps.TryGetValue(font, out IReadOnlyDictionary<int, string>? unicodeMapForCodes) ? unicodeMapForCodes : null, widthMap, fontSize)
                    : null;
                operations.Add(new PdfTextOperation(
                    pageNumber,
                    objectNumber,
                    generation,
                    font,
                    fontSize,
                    characterSpacing,
                    a,
                    b,
                    c,
                    d,
                    x,
                    y,
                    effective.A,
                    effective.B,
                    effective.C,
                    effective.D,
                    effective.X,
                    effective.Y,
                    match.Groups["operator"].Value,
                    payload,
                    payloadProfile.TextChunkCount,
                    payloadProfile.AdjustmentCount,
                    payloadProfile.AdjustmentSum,
                    payloadProfile.AdjustmentMin,
                    payloadProfile.AdjustmentMax,
                    averageAdjustmentPoints,
                    decodedRuneCount,
                    characterSpacingGapCount,
                    characterSpacingGapTotalPoints,
                    adjustmentTotalPoints,
                    characterSpacingGapTotalPoints + adjustmentTotalPoints,
                    naturalWidthPoints,
                    naturalWidthPoints + characterSpacingGapTotalPoints + adjustmentTotalPoints,
                    characterSpacing + averageAdjustmentPoints,
                    decodedText));
            }

            foreach (Match _ in RestoreGraphicsStateRegex.Matches(line))
            {
                if (graphicsStack.Count == 0)
                {
                    font = string.Empty;
                    fontSize = 0d;
                    characterSpacing = 0d;
                    currentMatrix = PdfMatrix.Identity;
                }
                else
                {
                    TextExtractionState restored = graphicsStack.Pop();
                    font = restored.Font;
                    fontSize = restored.FontSize;
                    characterSpacing = restored.CharacterSpacing;
                    currentMatrix = restored.CurrentMatrix;
                }
            }
        }

        return operations;
    }

    private sealed record TextExtractionState(
        string Font,
        double FontSize,
        double CharacterSpacing,
        PdfMatrix CurrentMatrix);

    private static double ReadDouble(string value) => double.Parse(value, CultureInfo.InvariantCulture);

    private static PdfTextPayloadProfile ReadPayloadProfile(string payload)
    {
        int textChunkCount = 0;
        int adjustmentCount = 0;
        double adjustmentSum = 0d;
        double adjustmentMin = 0d;
        double adjustmentMax = 0d;
        foreach (Match match in PayloadTokenRegex.Matches(payload))
        {
            string token = match.Value;
            if (token.StartsWith('(') || token.StartsWith('<'))
            {
                textChunkCount++;
                continue;
            }

            double adjustment = ReadDouble(token);
            if (adjustmentCount == 0)
            {
                adjustmentMin = adjustment;
                adjustmentMax = adjustment;
            }
            else
            {
                adjustmentMin = Math.Min(adjustmentMin, adjustment);
                adjustmentMax = Math.Max(adjustmentMax, adjustment);
            }

            adjustmentCount++;
            adjustmentSum += adjustment;
        }

        return new PdfTextPayloadProfile(
            textChunkCount,
            adjustmentCount,
            adjustmentSum,
            adjustmentCount == 0 ? 0d : adjustmentMin,
            adjustmentCount == 0 ? 0d : adjustmentMax);
    }

    private static double MeasurePayloadWidth(
        string payload,
        IReadOnlyDictionary<int, string>? unicodeMap,
        PdfFontWidthMap widthMap,
        double fontSize)
    {
        int widthUnits = 0;
        foreach (int code in ReadPayloadTextCodes(payload, unicodeMap))
        {
            widthUnits += widthMap.Widths.TryGetValue(code, out int width)
                ? width
                : widthMap.DefaultWidth;
        }

        return widthUnits * fontSize / 1000d;
    }

    private static IReadOnlyList<int> ReadPayloadTextCodes(string payload, IReadOnlyDictionary<int, string>? unicodeMap)
    {
        var codes = new List<int>();
        foreach (Match match in Regex.Matches(payload, @"\((?<literal>(?:\\.|[^\\)])*)\)|<(?<hex>[0-9A-Fa-f\s]+)>", RegexOptions.CultureInvariant))
        {
            if (match.Groups["literal"].Success)
            {
                foreach (char value in DecodeLiteralString(match.Groups["literal"].Value))
                {
                    codes.Add(value);
                }
            }
            else if (match.Groups["hex"].Success)
            {
                ReadHexCodes(match.Groups["hex"].Value, unicodeMap, codes);
            }
        }

        return codes;
    }

    private static void ReadHexCodes(string value, IReadOnlyDictionary<int, string>? unicodeMap, List<int> codes)
    {
        string hex = Regex.Replace(value, @"\s+", string.Empty);
        int codeHexLength = GetHexCodeLength(hex, unicodeMap);
        if (unicodeMap is null && hex.Length % codeHexLength != 0)
        {
            foreach (byte valueByte in Convert.FromHexString(hex))
            {
                codes.Add(valueByte);
            }

            return;
        }

        for (int index = 0; index + codeHexLength - 1 < hex.Length; index += codeHexLength)
        {
            codes.Add(int.Parse(hex.Substring(index, codeHexLength), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        }
    }

    private static string DecodePayload(string payload, IReadOnlyDictionary<int, string>? unicodeMap)
    {
        var builder = new StringBuilder();
        foreach (Match match in Regex.Matches(payload, @"\((?<literal>(?:\\.|[^\\)])*)\)|<(?<hex>[0-9A-Fa-f\s]+)>", RegexOptions.CultureInvariant))
        {
            if (match.Groups["literal"].Success)
            {
                builder.Append(DecodeLiteralString(match.Groups["literal"].Value));
            }
            else if (match.Groups["hex"].Success)
            {
                builder.Append(DecodeHexString(match.Groups["hex"].Value, unicodeMap));
            }
        }

        return builder.ToString();
    }

    private static string DecodeLiteralString(string value)
    {
        var builder = new StringBuilder();
        for (int i = 0; i < value.Length; i++)
        {
            char current = value[i];
            if (current != '\\' || i == value.Length - 1)
            {
                builder.Append(current);
                continue;
            }

            char escaped = value[++i];
            builder.Append(escaped switch
            {
                'n' => '\n',
                'r' => '\r',
                't' => '\t',
                'b' => '\b',
                'f' => '\f',
                '(' => '(',
                ')' => ')',
                '\\' => '\\',
                _ => escaped
            });
        }

        return builder.ToString();
    }

    private static string DecodeHexString(string value, IReadOnlyDictionary<int, string>? unicodeMap)
    {
        string hex = Regex.Replace(value, @"\s+", string.Empty);
        var builder = new StringBuilder();
        int codeHexLength = GetHexCodeLength(hex, unicodeMap);

        if (unicodeMap is null && hex.Length % codeHexLength != 0)
        {
            return Encoding.Latin1.GetString(Convert.FromHexString(hex));
        }

        for (int index = 0; index + codeHexLength - 1 < hex.Length; index += codeHexLength)
        {
            int code = int.Parse(hex.Substring(index, codeHexLength), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            if (unicodeMap is not null && unicodeMap.TryGetValue(code, out string? mapped))
            {
                builder.Append(mapped);
            }
            else
            {
                builder.Append(char.ConvertFromUtf32(code));
            }
        }

        return builder.ToString();
    }

    private static int GetHexCodeLength(string hex, IReadOnlyDictionary<int, string>? unicodeMap)
    {
        int codeHexLength = 4;
        if (unicodeMap is not null && hex.Length >= 4)
        {
            int firstTwoByteCode = int.Parse(hex.Substring(0, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            if (!unicodeMap.ContainsKey(firstTwoByteCode) && hex.Length >= 2)
            {
                int firstOneByteCode = int.Parse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                if (unicodeMap.ContainsKey(firstOneByteCode))
                {
                    codeHexLength = 2;
                }
            }
        }

        return codeHexLength;
    }
}

internal sealed record PdfTextPayloadProfile(
    int TextChunkCount,
    int AdjustmentCount,
    double AdjustmentSum,
    double AdjustmentMin,
    double AdjustmentMax);

internal sealed record PdfFontWidthMap(int DefaultWidth, IReadOnlyDictionary<int, int> Widths);

internal sealed record PdfPathCommand(string Operator, IReadOnlyList<double> Values);

internal sealed record PdfGraphicsOperation(
    int? PageNumber,
    int ObjectNumber,
    int Generation,
    string Kind,
    string Operator,
    int SegmentCount,
    int MoveCount,
    int LineCount,
    int CurveCount,
    int CloseCount,
    double MinX,
    double MinY,
    double MaxX,
    double MaxY,
    double LineWidth,
    string StrokeColor,
    string FillColor,
    string Dash,
    int LineCap,
    int LineJoin,
    IReadOnlyList<PdfPathCommand> PathCommands)
{
    private static readonly Regex TokenRegex = new(
        @"(?s)\[[^\]]*\]|\((?:\\.|[^\\)])*\)|<[^>]*>|/[^\s\[\]()<>{}%]+|-?(?:\d+\.?\d*|\.\d+)|[A-Za-z\*]+",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly HashSet<string> Operators = new(StringComparer.Ordinal)
    {
        "q", "Q", "cm", "w", "J", "j", "M", "d",
        "G", "g", "RG", "rg", "K", "k",
        "m", "l", "c", "v", "y", "h", "re",
        "S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "n", "W", "W*"
    };

    public static IReadOnlyList<PdfGraphicsOperation> Extract(int? pageNumber, int objectNumber, int generation, string stream)
    {
        var operations = new List<PdfGraphicsOperation>();
        var operands = new List<string>();
        var graphicsStack = new Stack<GraphicsState>();
        GraphicsState state = GraphicsState.Default;
        PdfPathBuilder path = new();

        foreach (Match match in TokenRegex.Matches(stream))
        {
            string token = match.Value;
            if (!Operators.Contains(token))
            {
                operands.Add(token);
                continue;
            }

            switch (token)
            {
                case "q":
                    graphicsStack.Push(state);
                    break;
                case "Q":
                    state = graphicsStack.Count == 0 ? GraphicsState.Default : graphicsStack.Pop();
                    break;
                case "cm" when TryReadNumbers(operands, 6, out double[] cm):
                    state = state with
                    {
                        CurrentMatrix = state.CurrentMatrix.Multiply(new PdfMatrix(cm[0], cm[1], cm[2], cm[3], cm[4], cm[5]))
                    };
                    break;
                case "w" when TryReadNumbers(operands, 1, out double[] width):
                    state = state with { LineWidth = width[0] };
                    break;
                case "J" when TryReadNumbers(operands, 1, out double[] cap):
                    state = state with { LineCap = (int)Math.Round(cap[0]) };
                    break;
                case "j" when TryReadNumbers(operands, 1, out double[] join):
                    state = state with { LineJoin = (int)Math.Round(join[0]) };
                    break;
                case "d":
                    state = state with { Dash = string.Join(" ", operands) };
                    break;
                case "G" when TryReadNumbers(operands, 1, out double[] grayStroke):
                    state = state with { StrokeColor = FormatColor("G", grayStroke) };
                    break;
                case "g" when TryReadNumbers(operands, 1, out double[] grayFill):
                    state = state with { FillColor = FormatColor("g", grayFill) };
                    break;
                case "RG" when TryReadNumbers(operands, 3, out double[] rgbStroke):
                    state = state with { StrokeColor = FormatColor("RG", rgbStroke) };
                    break;
                case "rg" when TryReadNumbers(operands, 3, out double[] rgbFill):
                    state = state with { FillColor = FormatColor("rg", rgbFill) };
                    break;
                case "K" when TryReadNumbers(operands, 4, out double[] cmykStroke):
                    state = state with { StrokeColor = FormatColor("K", cmykStroke) };
                    break;
                case "k" when TryReadNumbers(operands, 4, out double[] cmykFill):
                    state = state with { FillColor = FormatColor("k", cmykFill) };
                    break;
                case "m" when TryReadNumbers(operands, 2, out double[] move):
                    path.MoveTo(state.CurrentMatrix.Transform(move[0], move[1]));
                    break;
                case "l" when TryReadNumbers(operands, 2, out double[] line):
                    path.LineTo(state.CurrentMatrix.Transform(line[0], line[1]));
                    break;
                case "c" when TryReadNumbers(operands, 6, out double[] cubic):
                    path.CurveTo(
                        state.CurrentMatrix.Transform(cubic[0], cubic[1]),
                        state.CurrentMatrix.Transform(cubic[2], cubic[3]),
                        state.CurrentMatrix.Transform(cubic[4], cubic[5]));
                    break;
                case "v" when TryReadNumbers(operands, 4, out double[] cubicV):
                    path.ShorthandCurveTo(
                        "v",
                        state.CurrentMatrix.Transform(cubicV[0], cubicV[1]),
                        state.CurrentMatrix.Transform(cubicV[2], cubicV[3]));
                    break;
                case "y" when TryReadNumbers(operands, 4, out double[] cubicY):
                    path.ShorthandCurveTo(
                        "y",
                        state.CurrentMatrix.Transform(cubicY[0], cubicY[1]),
                        state.CurrentMatrix.Transform(cubicY[2], cubicY[3]));
                    break;
                case "h":
                    path.Close();
                    break;
                case "re" when TryReadNumbers(operands, 4, out double[] rect):
                    path.Rectangle(state.CurrentMatrix, rect[0], rect[1], rect[2], rect[3]);
                    break;
                case "W":
                case "W*":
                    AddOperation(operations, pageNumber, objectNumber, generation, path, state, "Clip", token);
                    break;
                case "S":
                case "s":
                    AddOperation(operations, pageNumber, objectNumber, generation, path, state, "Stroke", token);
                    path.Clear();
                    break;
                case "f":
                case "F":
                case "f*":
                    AddOperation(operations, pageNumber, objectNumber, generation, path, state, "Fill", token);
                    path.Clear();
                    break;
                case "B":
                case "B*":
                case "b":
                case "b*":
                    AddOperation(operations, pageNumber, objectNumber, generation, path, state, "FillStroke", token);
                    path.Clear();
                    break;
                case "n":
                    path.Clear();
                    break;
            }

            operands.Clear();
        }

        return operations;
    }

    private static void AddOperation(
        List<PdfGraphicsOperation> operations,
        int? pageNumber,
        int objectNumber,
        int generation,
        PdfPathBuilder path,
        GraphicsState state,
        string kind,
        string op)
    {
        if (!path.TryGetBounds(
            out double minX,
            out double minY,
            out double maxX,
            out double maxY,
            out int segmentCount,
            out int moveCount,
            out int lineCount,
            out int curveCount,
            out int closeCount))
        {
            return;
        }

        operations.Add(new PdfGraphicsOperation(
            pageNumber,
            objectNumber,
            generation,
            kind,
            op,
            segmentCount,
            moveCount,
            lineCount,
            curveCount,
            closeCount,
            Round(minX),
            Round(minY),
            Round(maxX),
            Round(maxY),
            Round(state.LineWidth),
            state.StrokeColor,
            state.FillColor,
            state.Dash,
            state.LineCap,
            state.LineJoin,
            path.GetCommands()));
    }

    private static bool TryReadNumbers(IReadOnlyList<string> operands, int count, out double[] values)
    {
        values = Array.Empty<double>();
        if (operands.Count < count)
        {
            return false;
        }

        var parsed = new double[count];
        int start = operands.Count - count;
        for (int i = 0; i < count; i++)
        {
            if (!double.TryParse(operands[start + i], NumberStyles.Float, CultureInfo.InvariantCulture, out parsed[i]))
            {
                return false;
            }
        }

        values = parsed;
        return true;
    }

    private static string FormatColor(string colorSpace, IReadOnlyList<double> values) =>
        colorSpace + ":" + string.Join(",", values.Select(value => Round(value).ToString("0.######", CultureInfo.InvariantCulture)));

    private static double Round(double value) => Math.Round(value, 6);

    private sealed record GraphicsState(
        PdfMatrix CurrentMatrix,
        double LineWidth,
        string StrokeColor,
        string FillColor,
        string Dash,
        int LineCap,
        int LineJoin)
    {
        public static GraphicsState Default { get; } = new(PdfMatrix.Identity, 1d, string.Empty, string.Empty, string.Empty, 0, 0);
    }

    private sealed class PdfPathBuilder
    {
        private readonly List<PdfPoint> points = new();
        private readonly List<PdfPathCommand> commands = new();
        private int segments;
        private int moves;
        private int lines;
        private int curves;
        private int closes;
        private PdfPoint? startPoint;
        private PdfPoint? currentPoint;

        public void MoveTo(PdfPoint point)
        {
            points.Add(point);
            commands.Add(new PdfPathCommand("m", new[] { Round(point.X), Round(point.Y) }));
            moves++;
            startPoint = point;
            currentPoint = point;
        }

        public void LineTo(PdfPoint point)
        {
            points.Add(point);
            commands.Add(new PdfPathCommand("l", new[] { Round(point.X), Round(point.Y) }));
            currentPoint = point;
            segments++;
            lines++;
        }

        public void CurveTo(PdfPoint firstControl, PdfPoint secondControl, PdfPoint endPoint)
        {
            points.Add(firstControl);
            points.Add(secondControl);
            points.Add(endPoint);
            commands.Add(new PdfPathCommand(
                "c",
                new[]
                {
                    Round(firstControl.X),
                    Round(firstControl.Y),
                    Round(secondControl.X),
                    Round(secondControl.Y),
                    Round(endPoint.X),
                    Round(endPoint.Y)
                }));
            currentPoint = endPoint;
            segments++;
            curves++;
        }

        public void ShorthandCurveTo(string op, PdfPoint control, PdfPoint endPoint)
        {
            points.Add(control);
            points.Add(endPoint);
            commands.Add(new PdfPathCommand(
                op,
                new[]
                {
                    Round(control.X),
                    Round(control.Y),
                    Round(endPoint.X),
                    Round(endPoint.Y)
                }));
            currentPoint = endPoint;
            segments++;
            curves++;
        }

        public void Close()
        {
            if (startPoint is not null)
            {
                commands.Add(new PdfPathCommand("h", Array.Empty<double>()));
                currentPoint = startPoint;
                segments++;
                closes++;
            }
        }

        public void Rectangle(PdfMatrix matrix, double x, double y, double width, double height)
        {
            MoveTo(matrix.Transform(x, y));
            LineTo(matrix.Transform(x + width, y));
            LineTo(matrix.Transform(x + width, y + height));
            LineTo(matrix.Transform(x, y + height));
            Close();
        }

        public bool TryGetBounds(
            out double minX,
            out double minY,
            out double maxX,
            out double maxY,
            out int segmentCount,
            out int moveCount,
            out int lineCount,
            out int curveCount,
            out int closeCount)
        {
            segmentCount = segments;
            moveCount = moves;
            lineCount = lines;
            curveCount = curves;
            closeCount = closes;
            if (points.Count == 0)
            {
                minX = minY = maxX = maxY = 0d;
                return false;
            }

            minX = points.Min(point => point.X);
            minY = points.Min(point => point.Y);
            maxX = points.Max(point => point.X);
            maxY = points.Max(point => point.Y);
            return true;
        }

        public IReadOnlyList<PdfPathCommand> GetCommands() => commands.ToArray();

        public void Clear()
        {
            points.Clear();
            commands.Clear();
            segments = 0;
            moves = 0;
            lines = 0;
            curves = 0;
            closes = 0;
            startPoint = null;
            currentPoint = null;
        }
    }
}

internal static class PdfToUnicodeMap
{
    private static readonly Regex BfCharRegex = new(
        @"<(?<source>[0-9A-Fa-f]+)>\s*<(?<target>[0-9A-Fa-f]+)>",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex BfRangeRegex = new(
        @"<(?<start>[0-9A-Fa-f]+)>\s*<(?<end>[0-9A-Fa-f]+)>\s*(?:<(?<target>[0-9A-Fa-f]+)>|\[(?<array>[^\]]*)\])",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static IReadOnlyDictionary<int, string> Parse(string cmap)
    {
        var map = new Dictionary<int, string>();
        foreach (string block in ReadBlocks(cmap, "beginbfchar", "endbfchar"))
        {
            foreach (Match match in BfCharRegex.Matches(block))
            {
                map[ReadHexInt(match.Groups["source"].Value)] = DecodeUtf16Hex(match.Groups["target"].Value);
            }
        }

        foreach (string block in ReadBlocks(cmap, "beginbfrange", "endbfrange"))
        {
            foreach (Match match in BfRangeRegex.Matches(block))
            {
                int start = ReadHexInt(match.Groups["start"].Value);
                int end = ReadHexInt(match.Groups["end"].Value);
                if (match.Groups["array"].Success)
                {
                    MatchCollection targets = Regex.Matches(match.Groups["array"].Value, @"<(?<target>[0-9A-Fa-f]+)>", RegexOptions.CultureInvariant);
                    for (int i = 0; i < targets.Count && start + i <= end; i++)
                    {
                        map[start + i] = DecodeUtf16Hex(targets[i].Groups["target"].Value);
                    }
                }
                else
                {
                    int targetStart = ReadHexInt(match.Groups["target"].Value);
                    for (int code = start; code <= end; code++)
                    {
                        map[code] = char.ConvertFromUtf32(targetStart + code - start);
                    }
                }
            }
        }

        return map;
    }

    private static IEnumerable<string> ReadBlocks(string cmap, string begin, string end)
    {
        var regex = new Regex(@"(?s)" + Regex.Escape(begin) + @"(?<block>.*?)" + Regex.Escape(end), RegexOptions.CultureInvariant);
        foreach (Match match in regex.Matches(cmap))
        {
            yield return match.Groups["block"].Value;
        }
    }

    private static int ReadHexInt(string value) => int.Parse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture);

    private static string DecodeUtf16Hex(string value)
    {
        byte[] bytes = Convert.FromHexString(value);
        return Encoding.BigEndianUnicode.GetString(bytes);
    }
}

internal readonly record struct PdfMatrix(double A, double B, double C, double D, double X, double Y)
{
    public static PdfMatrix Identity { get; } = new(1d, 0d, 0d, 1d, 0d, 0d);

    public PdfPoint Transform(double x, double y) => new(A * x + C * y + X, B * x + D * y + Y);

    public PdfMatrix Multiply(PdfMatrix right)
    {
        return new PdfMatrix(
            A * right.A + C * right.B,
            B * right.A + D * right.B,
            A * right.C + C * right.D,
            B * right.C + D * right.D,
            A * right.X + C * right.Y + X,
            B * right.X + D * right.Y + Y);
    }
}

internal readonly record struct PdfPoint(double X, double Y);
