namespace Lokad.OoxPdf.Fonts;

public sealed class WindowsFontResolver : IFontResolver
{
    private static readonly object CacheLock = new();
    private static readonly Dictionary<string, Lazy<IReadOnlyList<FontResolution>>> DiscoveryCaches = new(StringComparer.OrdinalIgnoreCase);
    private readonly Lazy<IReadOnlyList<FontResolution>> cache;

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

    public FontResolution Resolve(FontRequest request)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(request.FamilyName);

        FontResolution[] exact = cache.Value
            .Where(f => f.FamilyName.Equals(request.FamilyName, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (exact.Length != 0)
        {
            return SelectBest(exact, request);
        }

        FontResolution[] textFonts = cache.Value
            .Where(f => !f.HasMathTable)
            .ToArray();
        if (textFonts.Length != 0)
        {
            return SelectBest(textFonts, request) with { IsFallback = true };
        }

        return cache.Value.FirstOrDefault() is { } first
            ? first with { IsFallback = true }
            : new FontResolution(request.FamilyName, null, IsFallback: true);
    }

    internal FontResolution ResolvePresentationTextFace(FontRequest request)
    {
        FontResolution resolved = Resolve(request);
        if (!resolved.HasMathTable || string.IsNullOrWhiteSpace(resolved.FontFilePath))
        {
            return resolved;
        }

        FontResolution[] collectionTextFaces = cache.Value
            .Where(f => !f.HasMathTable &&
                f.FontFilePath is not null &&
                f.FontFilePath.Equals(resolved.FontFilePath, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        return collectionTextFaces.Length == 0
            ? resolved
            : SelectBest(collectionTextFaces, request) with { IsFallback = resolved.IsFallback };
    }

    internal IReadOnlyList<FontResolution> GetDiscoveredFonts()
    {
        return cache.Value;
    }

    private static Lazy<IReadOnlyList<FontResolution>> GetOrCreateCache(IReadOnlyList<string> fontDirectories)
    {
        string cacheKey = string.Join(
            "|",
            fontDirectories
                .Where(d => !string.IsNullOrWhiteSpace(d))
                .Select(Path.GetFullPath)
                .Order(StringComparer.OrdinalIgnoreCase));
        lock (CacheLock)
        {
            if (!DiscoveryCaches.TryGetValue(cacheKey, out Lazy<IReadOnlyList<FontResolution>>? cached))
            {
                cached = new Lazy<IReadOnlyList<FontResolution>>(() => Discover(fontDirectories));
                DiscoveryCaches[cacheKey] = cached;
            }

            return cached;
        }
    }

    private static IReadOnlyList<FontResolution> Discover(IReadOnlyList<string> fontDirectories)
    {
        var fonts = new List<FontResolution>();
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
                    int faceCount = OpenTypeFont.GetCollectionFontCount(bytes);
                    for (int faceIndex = 0; faceIndex < faceCount; faceIndex++)
                    {
                        OpenTypeFont font = OpenTypeFont.Load(bytes, faceIndex);
                        if (!string.IsNullOrWhiteSpace(font.FamilyName))
                        {
                            fonts.Add(new FontResolution(
                                font.FamilyName,
                                path,
                                IsFallback: false,
                                Bold: font.Os2.WeightClass >= 600,
                                Italic: Math.Abs(font.Post.ItalicAngle) > 0.01d,
                                WeightClass: font.Os2.WeightClass,
                                FontFaceIndex: faceIndex,
                                HasMathTable: font.TableTags.Contains("MATH")));
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

    private static FontResolution SelectBest(IReadOnlyList<FontResolution> candidates, FontRequest request)
    {
        int targetWeight = request.Bold ? 700 : 400;
        return candidates
            .OrderBy(f => f.Italic == request.Italic ? 0 : 1000)
            .ThenBy(f => Math.Abs(f.WeightClass - targetWeight))
            .ThenBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(f => f.FontFilePath, StringComparer.OrdinalIgnoreCase)
            .First();
    }
}
