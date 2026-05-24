using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;

if (args.Length is < 1 or > 3)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.PdfInspect <input.pdf> [output-directory] [--text-only]");
    return 2;
}

string inputPath = Path.GetFullPath(args[0]);
bool textOnly = args.Any(arg => string.Equals(arg, "--text-only", StringComparison.Ordinal));
string? outputDirectory = args
    .Skip(1)
    .FirstOrDefault(arg => !arg.StartsWith("--", StringComparison.Ordinal));
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
var textOperations = new List<PdfTextOperation>();

Console.WriteLine(FormattableString.Invariant($"PDF: {inputPath}"));
Console.WriteLine(FormattableString.Invariant($"Objects: {objects.Count}"));

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
        textOperations.AddRange(PdfTextOperation.Extract(pageNumber, item.Number, item.Generation, text, fontUnicodeMaps));
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

return 0;

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

        return objects;
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

    public static IReadOnlyList<PdfTextOperation> Extract(
        int? pageNumber,
        int objectNumber,
        int generation,
        string stream,
        IReadOnlyDictionary<string, IReadOnlyDictionary<int, string>> fontUnicodeMaps)
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
        var graphicsStack = new Stack<PdfMatrix>();

        foreach (string line in stream.Split('\n'))
        {
            foreach (Match _ in SaveGraphicsStateRegex.Matches(line))
            {
                graphicsStack.Push(currentMatrix);
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
                    decodedText));
            }

            foreach (Match _ in RestoreGraphicsStateRegex.Matches(line))
            {
                currentMatrix = graphicsStack.Count == 0 ? PdfMatrix.Identity : graphicsStack.Pop();
            }
        }

        return operations;
    }

    private static double ReadDouble(string value) => double.Parse(value, CultureInfo.InvariantCulture);

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
