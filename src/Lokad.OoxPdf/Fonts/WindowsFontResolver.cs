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

        FontResolution? exact = cache.Value.FirstOrDefault(f =>
            f.FamilyName.Equals(request.FamilyName, StringComparison.OrdinalIgnoreCase));
        if (exact is not null)
        {
            return exact;
        }

        FontResolution? arial = cache.Value.FirstOrDefault(f =>
            f.FamilyName.Equals("Arial", StringComparison.OrdinalIgnoreCase));
        if (arial is not null)
        {
            return arial with { IsFallback = true };
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
                     .Where(p => p.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) || p.EndsWith(".otf", StringComparison.OrdinalIgnoreCase))
                     .Order(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                OpenTypeFont font = OpenTypeFont.Load(path);
                if (!string.IsNullOrWhiteSpace(font.FamilyName))
                {
                    fonts.Add(new FontResolution(font.FamilyName, path, IsFallback: false));
                }
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException)
            {
                // Ignore fonts outside the minimal parser's current scope.
            }
        }

        return fonts
            .GroupBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .OrderBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }
}
