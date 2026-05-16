using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;

if (args.Length is < 1 or > 2)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.PdfInspect <input.pdf> [output-directory]");
    return 2;
}

string inputPath = Path.GetFullPath(args[0]);
string? outputDirectory = args.Length == 2 ? Path.GetFullPath(args[1]) : null;
if (outputDirectory is not null)
{
    Directory.CreateDirectory(outputDirectory);
}

byte[] bytes = File.ReadAllBytes(inputPath);
string pdf = Encoding.Latin1.GetString(bytes);
var objects = PdfObject.ParseAll(pdf, bytes);
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
    File.WriteAllText(prefix + ".dict.txt", item.Dictionary, Encoding.UTF8);
    if (LooksTextual(item.Stream.Decoded))
    {
        string text = Encoding.Latin1.GetString(item.Stream.Decoded);
        File.WriteAllText(prefix + ".stream.txt", text, Encoding.UTF8);
        textOperations.AddRange(PdfTextOperation.Extract(item.Number, item.Generation, text));
    }
    else
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
    if (body.Contains("/Type /Page", StringComparison.Ordinal))
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

    public static IReadOnlyList<PdfObject> ParseAll(string pdf, byte[] bytes)
    {
        var objects = new List<PdfObject>();
        foreach (Match match in ObjectRegex.Matches(pdf))
        {
            int number = int.Parse(match.Groups["number"].Value, CultureInfo.InvariantCulture);
            int generation = int.Parse(match.Groups["generation"].Value, CultureInfo.InvariantCulture);
            string rawBody = match.Groups["body"].Value;
            string body = rawBody.Trim();
            string dictionary = ReadDictionary(body);
            PdfStream? stream = TryReadStream(match, rawBody, bytes, dictionary);
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

    private static PdfStream? TryReadStream(Match match, string body, byte[] bytes, string dictionary)
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
    string Operator,
    string Payload)
{
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

    public static IReadOnlyList<PdfTextOperation> Extract(int objectNumber, int generation, string stream)
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

        foreach (string line in stream.Split('\n'))
        {
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
                operations.Add(new PdfTextOperation(
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
                    match.Groups["operator"].Value,
                    match.Groups["payload"].Value));
            }
        }

        return operations;
    }

    private static double ReadDouble(string value) => double.Parse(value, CultureInfo.InvariantCulture);
}
