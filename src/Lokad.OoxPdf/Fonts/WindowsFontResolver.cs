namespace Lokad.OoxPdf.Fonts;

public sealed class WindowsFontResolver : IFontResolver
{
    private readonly Lazy<IReadOnlyList<FontResolution>> cache;

    public WindowsFontResolver()
        : this(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"))
    {
    }

    internal WindowsFontResolver(string fontsDirectory)
    {
        cache = new Lazy<IReadOnlyList<FontResolution>>(() => Discover(fontsDirectory));
    }

    public FontResolution Resolve(FontRequest request)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(request.FamilyName);

        if (request.FamilyName.Equals("Cambria Math", StringComparison.OrdinalIgnoreCase))
        {
            FontResolution[] cambria = cache.Value
                .Where(f => f.FamilyName.Equals("Cambria", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (cambria.Length != 0)
            {
                return SelectBest(cambria, request) with { IsFallback = true };
            }
        }

        FontResolution[] exact = cache.Value
            .Where(f => f.FamilyName.Equals(request.FamilyName, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (exact.Length != 0)
        {
            return SelectBest(exact, request);
        }

        foreach (string alias in ResolveAliases(request.FamilyName))
        {
            FontResolution[] aliasMatches = cache.Value
                .Where(f => f.FamilyName.Equals(alias, StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (aliasMatches.Length != 0)
            {
                return SelectBest(aliasMatches, request) with { IsFallback = true };
            }
        }

        FontResolution[] arial = cache.Value
            .Where(f => f.FamilyName.Equals("Arial", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (arial.Length != 0)
        {
            return SelectBest(arial, request) with { IsFallback = true };
        }

        return cache.Value.FirstOrDefault() is { } first
            ? first with { IsFallback = true }
            : new FontResolution(request.FamilyName, null, IsFallback: true);
    }

    internal IReadOnlyList<FontResolution> GetDiscoveredFonts()
    {
        return cache.Value;
    }

    private static IReadOnlyList<FontResolution> Discover(string fontsDirectory)
    {
        if (!Directory.Exists(fontsDirectory))
        {
            return [];
        }

        var fonts = new List<FontResolution>();
        foreach (string path in Directory.EnumerateFiles(fontsDirectory, "*.*", SearchOption.TopDirectoryOnly)
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
                            FontFaceIndex: faceIndex));
                    }
                }
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException)
            {
                // Ignore fonts outside the minimal parser's current scope.
            }
        }

        return fonts
            .OrderBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static FontResolution SelectBest(IReadOnlyList<FontResolution> candidates, FontRequest request)
    {
        int targetWeight = request.Bold ? 700 : 400;
        return candidates
            .OrderBy(f => f.Italic == request.Italic ? 0 : 1000)
            .ThenBy(f => Math.Abs(f.WeightClass - targetWeight))
            .ThenBy(f => f.FontFilePath, StringComparer.OrdinalIgnoreCase)
            .First();
    }

    private static IReadOnlyList<string> ResolveAliases(string familyName)
    {
        if (familyName.Equals("Aptos Display", StringComparison.OrdinalIgnoreCase))
        {
            return ["Calibri Light", "Calibri"];
        }

        return familyName.Equals("Aptos", StringComparison.OrdinalIgnoreCase)
            ? ["Calibri"]
            : [];
    }
}
