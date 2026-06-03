namespace Lokad.OoxPdf.Fonts;

public sealed class WindowsFontResolver : IFontResolver
{
    private static readonly object CacheLock = new();
    private static readonly Dictionary<string, Lazy<IReadOnlyList<FontFaceResolution>>> DiscoveryCaches = new(StringComparer.OrdinalIgnoreCase);
    private readonly Lazy<IReadOnlyList<FontFaceResolution>> cache;

    public WindowsFontResolver()
        : this(GetDefaultFontDirectories())
    {
    }

    internal WindowsFontResolver(string fontsDirectory)
        : this([fontsDirectory])
    {
    }

    private WindowsFontResolver(IReadOnlyList<string> fontDirectories)
    {
        cache = GetOrCreateCache(fontDirectories);
    }

    public FontFaceResolution Resolve(FontRequest request)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(request.FamilyName);

        FontFaceResolution[] exact = cache.Value
            .Where(f => f.FamilyName.Equals(request.FamilyName, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (exact.Length != 0)
        {
            return SelectBest(exact, request);
        }

        FontFaceResolution[] textFonts = cache.Value
            .Where(f => !f.HasMathTable)
            .ToArray();
        if (textFonts.Length != 0)
        {
            return SelectBest(textFonts, request) with { RequestedFamily = request.FamilyName, IsFallback = true };
        }

        return cache.Value.FirstOrDefault() is { } first
            ? first with { RequestedFamily = request.FamilyName, IsFallback = true }
            : new FontFaceResolution(
                request.FamilyName,
                request.FamilyName,
                new FontStyleKey(request.Bold, request.Italic),
                new MemoryFontProgramSource("missing:" + request.FamilyName, ReadOnlyMemory<byte>.Empty),
                IsFallback: true);
    }

    internal FontFaceResolution ResolvePresentationTextFace(FontRequest request)
    {
        return Resolve(request);
    }

    internal IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
    {
        return cache.Value;
    }

    private static Lazy<IReadOnlyList<FontFaceResolution>> GetOrCreateCache(IReadOnlyList<string> fontDirectories)
    {
        string cacheKey = string.Join(
            "|",
            fontDirectories
                .Where(d => !string.IsNullOrWhiteSpace(d))
                .Select(Path.GetFullPath)
                .Order(StringComparer.OrdinalIgnoreCase));
        lock (CacheLock)
        {
            if (!DiscoveryCaches.TryGetValue(cacheKey, out Lazy<IReadOnlyList<FontFaceResolution>>? cached))
            {
                cached = new Lazy<IReadOnlyList<FontFaceResolution>>(() => Discover(fontDirectories));
                DiscoveryCaches[cacheKey] = cached;
            }

            return cached;
        }
    }

    private static IReadOnlyList<FontFaceResolution> Discover(IReadOnlyList<string> fontDirectories)
    {
        var fonts = new List<FontFaceResolution>();
        foreach (string fontsDirectory in fontDirectories.Where(Directory.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            SearchOption searchOption = fontsDirectory.Contains("CloudFonts", StringComparison.OrdinalIgnoreCase)
                ? SearchOption.AllDirectories
                : SearchOption.TopDirectoryOnly;
            foreach (string path in Directory.EnumerateFiles(fontsDirectory, "*.*", searchOption)
                         .Where(p => p.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) ||
                             p.EndsWith(".otf", StringComparison.OrdinalIgnoreCase) ||
                             p.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase))
                         .Order(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    byte[] bytes = File.ReadAllBytes(path);
                    var source = new FileFontProgramSource(path);
                    int faceCount = OpenTypeFont.GetCollectionFontCount(bytes);
                    for (int faceIndex = 0; faceIndex < faceCount; faceIndex++)
                    {
                        OpenTypeFont font = OpenTypeFont.Load(bytes, faceIndex);
                        if (!string.IsNullOrWhiteSpace(font.FamilyName))
                        {
                            fonts.Add(new FontFaceResolution(
                                font.FamilyName,
                                font.FamilyName,
                                new FontStyleKey(
                                    Bold: font.Os2.WeightClass >= 600,
                                    Italic: Math.Abs(font.Post.ItalicAngle) > 0.01d,
                                    WeightClass: font.Os2.WeightClass,
                                    FaceIndex: faceIndex,
                                    HasMathTable: font.TableTags.Contains("MATH")),
                                source,
                                IsFallback: false));
                        }
                    }
                }
                catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
                {
                    // Ignore fonts outside the minimal parser's current scope.
                }
            }
        }

        return fonts
            .OrderBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static IReadOnlyList<string> GetDefaultFontDirectories()
    {
        string windowsFonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string cloudFonts = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft",
            "FontCache",
            "4",
            "CloudFonts");
        return [windowsFonts, cloudFonts];
    }

    private static FontFaceResolution SelectBest(IReadOnlyList<FontFaceResolution> candidates, FontRequest request)
    {
        int targetWeight = request.Bold ? 700 : 400;
        return candidates
            .OrderBy(f => f.Italic == request.Italic ? 0 : 1000)
            .ThenBy(f => Math.Abs(f.WeightClass - targetWeight))
            .ThenBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(f => f.Source.StableId, StringComparer.OrdinalIgnoreCase)
            .ThenBy(f => f.FontFaceIndex)
            .First();
    }
}
