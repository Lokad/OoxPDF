using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Fonts;

// Diagnostic resolver decoration only: keep native resolution, opaque stable
// identities, face indices, catalog access, bytes and cancellation. No font cache
// or parsed-font reuse is introduced.
internal static class FontLoadCounts
{
    internal static int Run(string[] inputs, string reportPath, IFontResolver? explicitResolver, bool review)
    {
        var rows = new List<object>();
        foreach (string input in inputs)
        {
            byte[] bytes = File.ReadAllBytes(input);
            IFontResolver native = explicitResolver ?? new WindowsFontResolver();
            CountingResolver resolver = native is IFontCatalog catalog
                ? new CountingCatalogResolver(native, catalog) : new CountingResolver(native);
            using var source = new MemoryStream(bytes, writable: false);
            using var output = new MemoryStream();
            OoxPdfConverter.Convert(source, output, new OoxPdfOptions
            {
                InputKind = input.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) ? OoxPdfInputKind.Pptx : OoxPdfInputKind.Docx,
                FontResolver = resolver,
                DocxMarkupMode = review ? OoxPdfDocxMarkupMode.AllMarkup : OoxPdfDocxMarkupMode.Final,
                DocxMarkupGeometryMode = review ? OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout
            });
            rows.Add(new
            {
                name = Path.GetFileName(input),
                inputSha256 = Convert.ToHexString(SHA256.HashData(bytes)),
                outputSha256 = Convert.ToHexString(SHA256.HashData(output.GetBuffer().AsSpan(0, checked((int)output.Length)))),
                outputBytes = output.Length,
                faces = resolver.Snapshot()
            });
        }
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(reportPath))!);
        File.WriteAllText(reportPath, JsonSerializer.Serialize(new
        {
            scope = "One diagnostic conversion/input. Delegated program-source reads indicate loader requests, not physical file reads, parse success or layout-pass counts. No timing/allocation attribution.",
            librarySha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(typeof(OoxPdfConverter).Assembly.Location))),
            mode = review ? "all-word-compatible" : "final-preserve",
            inputs = rows
        }, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"Traced font-program loads for {rows.Count} inputs; no allocation benchmark.");
        return 0;
    }

    private class CountingResolver(IFontResolver inner) : IFontResolver
    {
        private readonly object gate = new();
        private readonly Dictionary<(string Id, int Face), CountingSource> sources = new();

        public FontFaceResolution Resolve(FontRequest request) => Wrap(inner.Resolve(request));

        protected FontFaceResolution Wrap(FontFaceResolution resolution)
        {
            lock (gate)
            {
                var key = (resolution.Source.StableId, resolution.Style.FaceIndex);
                if (!sources.TryGetValue(key, out var source))
                {
                    source = new CountingSource(resolution.Source, resolution.ResolvedFamily, resolution.Style.FaceIndex);
                    sources.Add(key, source);
                }
                return resolution with { Source = source };
            }
        }

        internal object[] Snapshot()
        {
            lock (gate)
            {
                return sources.Values.Where(source => source.Requests != 0).OrderBy(source => source.Family, StringComparer.Ordinal)
                    .ThenBy(source => source.StableId, StringComparer.Ordinal).ThenBy(source => source.FaceIndex)
                    .Select(source => (object)new
                    {
                        family = source.Family, faceIndex = source.FaceIndex,
                        sourceIdentitySha256 = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(source.StableId))),
                        programRequests = source.Requests
                    }).ToArray();
            }
        }
    }

    private sealed class CountingCatalogResolver(IFontResolver inner, IFontCatalog catalog) : CountingResolver(inner), IFontCatalog
    {
        IReadOnlyList<FontFaceResolution> IFontCatalog.GetDiscoveredFonts() => catalog.GetDiscoveredFonts().Select(Wrap).ToArray();
    }

    private sealed class CountingSource(IFontProgramSource inner, string family, int faceIndex) : IFontProgramSource
    {
        private int requests;
        internal string Family => family;
        internal int FaceIndex => faceIndex;
        internal int Requests => Volatile.Read(ref requests);
        public string StableId => inner.StableId;
        public ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct = default)
        {
            Interlocked.Increment(ref requests);
            return inner.GetBytesAsync(ct);
        }
    }
}
