using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxFooterRevisionFrameTests
{
    public static void DocxFooterRevisionFramesMatchIndependentWordBoundsAndUiFaces()
    {
        if (!OperatingSystem.IsWindows() || !File.Exists(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc")))
            TestAssert.Skip("Independent footer frames require installed Word-compatible fonts.");
        using var gold = ReadGold();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var document = Read(row);
            var renderer = new DocxRenderer(null, OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup);
            var placement = renderer.InspectMarkupBalloons(document).Single();
            var page = renderer.RenderPagesAsync(document, null, default).ToBlockingEnumerable().Single();
            double[] bounds = [placement.X, page.Height - placement.Y - placement.Height,
                placement.X + placement.Width, page.Height - placement.Y];
            for (int edge = 0; edge < bounds.Length; edge++)
                TestAssert.True(Math.Abs(bounds[edge] - row.GetProperty("Frame")[edge].GetDouble()) < 1d,
                    "The source-owned footer frame retains Word's absolute bounds: " + row.GetProperty("Id").GetString());
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).ToArray();
            var title = shows.Single(show => show.Text.Trim() == "Formatted:");
            var body = shows.Single(show => show.Text.Trim() == row.GetProperty("Caption").GetString());
            TestAssert.True(Math.Abs(title.X - row.GetProperty("TitleOrigin")[0].GetDouble()) < 1d &&
                Math.Abs(page.Height - title.Y - row.GetProperty("TitleOrigin")[1].GetDouble()) < 1d,
                "Footer title origins retain the independent one-point gate.");
            TestAssert.Equal(700, (int)title.Font.Font.Os2.WeightClass);
            TestAssert.Equal(400, (int)body.Font.Font.Os2.WeightClass);
            TestAssert.Equal("Segoe UI", body.Font.Font.FamilyName);
        }
    }

    public static void DocxFooterRevisionFramesKeepPortableContentAndOwnedCaptionFaces()
    {
        using var gold = ReadGold();
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var resolver = new ReviewResolver();
            var page = Render(Read(row), resolver);
            var body = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                .Single(show => show.Text.Trim() == row.GetProperty("Caption").GetString());
            TestAssert.Equal("DejaVu Sans", body.Font.Font.FamilyName);
            TestAssert.Equal(1, resolver.RegularUiRequests);
            TestAssert.True(page.Content.Length != 0 && page.Fonts.Count != 0, "Caption preparation retains source rendering.");
            var math = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true).Where(show => page.Fonts.Any(font =>
                font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) && ReferenceEquals(font.Font, show.Font))).ToArray();
            TestAssert.Equal(row.GetProperty("ExpectedEquationCount").GetInt32() == 0 ? 0 : 18, math.Length);
        }
        var groupedParagraphs = new[] { "Public alpha α", "Public beta β", "Public gamma γ" }
            .Select((text, index) => new DocxParagraphElement(DocxTests.CreateDocxLayoutParagraph(text, 11d, 8d) with
            {
                LineSpacingPoints = 8d,
                Spacing = DocxParagraphSpacing.Empty with { LineRuleValue = "exact", LineValue = "160" },
                Revisions = [new DocxRevisionInfo(DocxRevisionKind.RunPropertiesChange, index.ToString(), "Public Reviewer", null,
                    "rPrChange", null, propertyElementNames: ["b"])]
            })).ToArray();
        var grouped = new DocxDocument(612d, 792d, 72d, 300d, 72d, 72d, DocxPageSettings.Empty, [], [], [], groupedParagraphs, [], [])
            { MarkupMode = OoxPdfDocxMarkupMode.AllMarkup };
        var groupedPage = Render(grouped, new ReviewResolver());
        string captions = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(groupedPage, true)
            .Where(show => show.Font.Font.FamilyName == "DejaVu Sans").Select(show => show.Text));
        foreach (string expected in new[] { "Public alpha α", "Public beta β", "Public gamma γ" })
            TestAssert.Contains(expected, captions);
        var commentResolver = new ReviewResolver();
        _ = Render(DocxTests.ReadDocx(Path.Combine(Cases, "UnitPrintOrigins", "arial-unscaled.docx"),
            OoxPdfDocxMarkupMode.AllMarkup), commentResolver);
        TestAssert.Equal(0, commentResolver.RegularUiRequests);
    }

    public static void DocxFooterRevisionFramesKeepFontFallbackAndTerminalAcquisitionFailures()
    {
        using var gold = ReadGold();
        var document = Read(gold.RootElement.GetProperty("Rows")[0]);
        foreach (string mode in new[] { "missing", "unloadable", "incomplete", "cff" })
        {
            var resolver = new ReviewResolver(mode);
            var page = Render(document, resolver);
            var body = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true)
                .Single(show => show.Text.Trim() == "Font: Bold");
            TestAssert.Equal("DejaVu Math TeX Gyre", body.Font.Font.FamilyName);
        }
        TestAssert.Throws<OoxPdfLimitExceededException>(() => Render(document, new ReviewResolver("refused")));
        using var cancellation = new CancellationTokenSource();
        TestAssert.Throws<OperationCanceledException>(() =>
            Render(document, new ReviewResolver("cancel", cancellation), cancellation.Token));
    }

    private static PdfPage Render(DocxDocument document, IFontResolver resolver, CancellationToken token = default) =>
        new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup, OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup)
            .RenderPagesAsync(document, null, token).ToBlockingEnumerable(token).Single();

    private static DocxDocument Read(JsonElement row)
    {
        string path = Path.Combine(Cases, row.GetProperty("Input").GetString()!);
        TestAssert.Equal(row.GetProperty("Sha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        return DocxTests.ReadDocx(path, OoxPdfDocxMarkupMode.AllMarkup);
    }

    private static string Cases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases"));
    private static JsonDocument ReadGold() => JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "FooterRevisionFrames", "office.json")));

    private sealed class ReviewResolver(string mode = "available", CancellationTokenSource? cancellation = null) : IFontResolver
    {
        private readonly PortableRenderingTests.PinnedResolver source = new(mathFont: true);
        private readonly PortableRenderingTests.PinnedResolver ui = new();
        public int RegularUiRequests { get; private set; }

        public FontFaceResolution Resolve(FontRequest request)
        {
            bool uiFamily = request.FamilyName.Equals("Segoe UI", StringComparison.OrdinalIgnoreCase);
            if (!uiFamily || request.Bold)
                return source.Resolve(request) with { IsFallback = uiFamily };
            RegularUiRequests++;
            var face = ui.Resolve(request);
            return mode switch
            {
                "missing" => face with { IsFallback = true },
                "incomplete" => face with { Source = new MemoryFontProgramSource("footer-ui-incomplete", IncompleteFont()) },
                "cff" => face with { Source = new MemoryFontProgramSource("footer-ui-cff", TestFontBuilder.CreateCffKindFont()) },
                "unloadable" or "refused" or "cancel" => face with { Source = new FailedUiSource(mode, cancellation) },
                _ => face
            };
        }

        private static byte[] IncompleteFont()
        {
            byte[] bytes = TestFontBuilder.CreateTestFont();
            var cmap = TestFontBuilder.GetTableRange(bytes, "cmap");
            bytes[cmap.Offset + 2] = bytes[cmap.Offset + 3] = 0;
            TestAssert.Equal((ushort)0, OpenTypeFont.Load(bytes).MapCodePoint('F'));
            return bytes;
        }
    }

    private sealed class FailedUiSource(string mode, CancellationTokenSource? cancellation) : IFontProgramSource
    {
        public string StableId => "footer-ui-" + mode;
        public ValueTask<IOoxPdfFontProgramLease> AcquireAsync(IMemoryBudget? operationBudget, CancellationToken ct = default)
        {
            if (mode == "refused") throw new OoxPdfLimitExceededException("Synthetic font admission refusal.",
                OoxPdfLimitKind.MemoryAdmission, OoxPdfLimitScope.Conversion);
            if (mode == "cancel") { cancellation!.Cancel(); ct.ThrowIfCancellationRequested(); }
            throw new IOException("Synthetic unavailable UI font.");
        }
    }
}
