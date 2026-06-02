using System.Globalization;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed class DocxRenderer
{
    internal const string DefaultDocumentTypefaceRequest = "OOXPDF_DOCUMENT_DEFAULT";
    private readonly IFontResolver fontResolver;

    public DocxRenderer(IFontResolver? fontResolver = null)
    {
        this.fontResolver = fontResolver ?? new WindowsFontResolver();
    }

    public IReadOnlyList<PdfPage> RenderBlankPages(DocxDocument document, Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        if (document.BodyElements.Count == 0 && document.HeaderParagraphs.Count == 0 && document.FooterParagraphs.Count == 0)
        {
            return [new PdfPage(document.PageWidthPoints, document.PageHeightPoints)];
        }

        return RenderParagraphs(document, fontResolver, diagnosticSink);
    }

    internal DocxLayoutSnapshot InspectLayout(DocxDocument document)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver);
        DocxLayout layout = new DocxLayoutEngine().Create(document, fontResources.TextMeasurer);
        return DocxLayoutSnapshot.FromLayout(layout);
    }

    internal DocxFontPlanSnapshot InspectFontPlan(DocxDocument document)
    {
        return DocxFontPlanSnapshot.FromPlan(DocxFontPlan.Create(document, fontResolver));
    }

    internal DocxTextEmissionSnapshot InspectTextEmission(DocxDocument document)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver);
        DocxLayout layout = new DocxLayoutEngine().Create(document, fontResources.TextMeasurer);
        var lines = new List<DocxTextEmissionLineSnapshot>();
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            DocxLayoutPage page = layout.Pages[pageIndex];
            int pageNumber = pageIndex + 1;
            foreach (DocxTextLineLayout line in page.StaticTextLines)
            {
                lines.Add(ToTextEmissionLineSnapshot(pageIndex, isStaticStory: true, line, fontResources, pageNumber, layout.Pages.Count));
            }

            foreach (DocxTextLineLayout line in EnumerateBodyTextLines(page))
            {
                lines.Add(ToTextEmissionLineSnapshot(pageIndex, isStaticStory: false, line, fontResources, pageNumber, layout.Pages.Count));
            }
        }

        return new DocxTextEmissionSnapshot(
            lines.Count,
            lines.Sum(line => line.SegmentCount),
            lines.Sum(line => line.TerminalSpaceSegmentCount),
            lines.Sum(line => line.NonzeroPdfCharacterSpacingSegmentCount),
            lines.Sum(line => line.Segments.Count(segment => segment.CompensatePdfCharacterSpacing)),
            lines);
    }

    internal DocxStructureSnapshot InspectStructure(DocxDocument document)
    {
        return DocxStructureSnapshot.FromDocument(document);
    }

    private static IReadOnlyList<PdfPage> RenderParagraphs(DocxDocument document, IFontResolver fontResolver, Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver);

        DocxLayout layout = new DocxLayoutEngine().Create(document, fontResources.TextMeasurer);
        var pages = new List<PdfPage>(layout.Pages.Count);
        int imageIndex = 1;

        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            DocxLayoutPage layoutPage = layout.Pages[pageIndex];
            var graphics = new PdfGraphicsBuilder();
            var pageImages = new List<PdfImageResource>();
            int pageNumber = pageIndex + 1;
            foreach (DocxTextLineLayout staticLine in layoutPage.StaticTextLines)
            {
                RenderTextLine(staticLine, graphics, fontResources, pageNumber, layout.Pages.Count);
            }

            for (int itemIndex = 0; itemIndex < layoutPage.Items.Count; itemIndex++)
            {
                DocxLayoutItem item = layoutPage.Items[itemIndex];
                DocxTableRowLayout? previousRow = itemIndex > 0 ? layoutPage.Items[itemIndex - 1] as DocxTableRowLayout : null;
                DocxTableRowLayout? nextRow = itemIndex + 1 < layoutPage.Items.Count ? layoutPage.Items[itemIndex + 1] as DocxTableRowLayout : null;
                RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, diagnosticSink, pageIndex + 1, layout.Pages.Count, ref imageIndex);
            }

            pages.Add(new PdfPage(layoutPage.Width, layoutPage.Height, graphics.ToString(), fontResources.Resources, pageImages.ToArray()));
        }

        return pages;
    }

    private static DocxFontResources PrepareFontResources(DocxDocument document, IFontResolver fontResolver)
    {
        DocxFontPlan plan = DocxFontPlan.Create(document, fontResolver);
        var resources = new List<PdfFontResource>();
        var runResources = new Dictionary<DocxTextRun, DocxRunFontResource>();
        PrepareResolvedRunFontResources(plan, resources, runResources);
        DocxRunFontResource? fallback = PrepareFallbackFontResource(plan, fontResolver, resources, runResources);
        IDocxTextMeasurer? measurer = plan.Runs.Any(run => run.Resolution?.FontFilePath is not null && File.Exists(run.Resolution.FontFilePath)) || fallback is not null
            ? new DocxFontPlanTextMeasurer(plan, fallback?.Resolution)
            : null;
        return new DocxFontResources(plan, measurer, resources, runResources, fallback);
    }

    private static void PrepareResolvedRunFontResources(
        DocxFontPlan plan,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources)
    {
        foreach (IGrouping<(string Path, int FaceIndex), DocxResolvedRunTypeface> group in plan.Runs
            .Where(run => run.Resolution?.FontFilePath is not null && File.Exists(run.Resolution.FontFilePath))
            .GroupBy(run => (run.Resolution!.FontFilePath!, run.Resolution.FontFaceIndex)))
        {
            FontResolution resolution = group.First().Resolution!;
            IReadOnlyList<int> glyphs = CollectRunGlyphs(group);
            if (glyphs.Count == 0)
            {
                continue;
            }

            OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath!, resolution.FontFaceIndex);
            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, glyphs);
            string name = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            var runResource = new DocxRunFontResource(name, embedded, resolution);
            resources.Add(new PdfFontResource(name, embedded));
            foreach (DocxResolvedRunTypeface run in group)
            {
                runResources[run.Run] = runResource;
            }
        }
    }

    private static DocxRunFontResource? PrepareFallbackFontResource(
        DocxFontPlan plan,
        IFontResolver fontResolver,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources)
    {
        FontResolution resolution = ResolveDocumentBaseFont(plan, fontResolver);
        if (resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath))
        {
            return null;
        }

        DocxResolvedRunTypeface[] fallbackRuns = plan.Runs
            .Where(run => !runResources.ContainsKey(run.Run))
            .ToArray();
        if (fallbackRuns.Length == 0)
        {
            return null;
        }

        IReadOnlyList<int> glyphs = CollectRunGlyphs(fallbackRuns);
        if (glyphs.Count == 0)
        {
            return null;
        }

        OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath, resolution.FontFaceIndex);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, glyphs);
        string name = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
        var runResource = new DocxRunFontResource(name, embedded, resolution);
        resources.Add(new PdfFontResource(name, embedded));
        foreach (DocxResolvedRunTypeface run in fallbackRuns)
        {
            runResources[run.Run] = runResource;
        }

        return runResource;
    }

    private static IReadOnlyList<int> CollectRunGlyphs(IEnumerable<DocxResolvedRunTypeface> runs)
    {
        return runs
            .SelectMany(run => run.Run.Text.EnumerateRunes().Select(rune => rune.Value))
            .Concat(" ".EnumerateRunes().Select(rune => rune.Value))
            .Concat("0123456789".EnumerateRunes().Select(rune => rune.Value))
            .Distinct()
            .ToArray();
    }

    private static FontResolution ResolveDocumentBaseFont(DocxFontPlan plan, IFontResolver fontResolver)
    {
        foreach (DocxResolvedRunTypeface run in plan.Runs)
        {
            if (run.Resolution is { FontFilePath: not null } resolution)
            {
                return resolution;
            }
        }

        return fontResolver.Resolve(new FontRequest(DefaultDocumentTypefaceRequest));
    }

    private static void RenderLayoutItem(
        DocxLayoutItem item,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageNumber,
        int pageCount,
        ref int imageIndex)
    {
        switch (item)
        {
            case DocxTextLineLayout textLine:
                RenderTextLine(textLine, graphics, fontResources, pageNumber, pageCount);
                break;
            case DocxInlineImageLayout image:
                RenderInlineImage(image, graphics, pageImages, diagnosticSink, ref imageIndex);
                break;
            case DocxTableRowLayout row:
                RenderTableRow(row, IsAdjacentTableRow(previousRow, row) ? previousRow : null, IsAdjacentTableRow(row, nextRow) ? nextRow : null, graphics, pageImages, fontResources, diagnosticSink, pageNumber, pageCount, ref imageIndex);
                break;
        }
    }

    private static bool IsAdjacentTableRow(DocxTableRowLayout? first, DocxTableRowLayout? second)
    {
        return first is not null &&
            second is not null &&
            first.Table.TableIndex == second.Table.TableIndex &&
            (first.RowIndex + 1 == second.RowIndex ||
                (first.RowIndex == second.RowIndex && first.FragmentIndex + 1 == second.FragmentIndex));
    }

    private static void RenderTextLine(DocxTextLineLayout line, PdfGraphicsBuilder graphics, DocxFontResources fontResources, int pageNumber, int pageCount)
    {
        foreach (DocxTextEmissionSegment segment in CreateTextEmissionSegments(line, fontResources, pageNumber, pageCount))
        {
            RenderTextEmissionSegment(segment, graphics);
        }
    }

    private static double GetSegmentFontSize(DocxTextSegmentLayout segment, double lineFontSize)
    {
        return segment.FontSize ?? lineFontSize;
    }

    private static double GetSegmentBaselineY(DocxTextSegmentLayout segment, double lineBaselineY)
    {
        return lineBaselineY + segment.BaselineOffsetY;
    }

    private static void RenderTextEmissionSegment(DocxTextEmissionSegment segment, PdfGraphicsBuilder graphics)
    {
        DocxTextRun style = segment.StyleRun;
        RgbColor color = segment.Color;
        if (!segment.IsTerminalLineSpace)
        {
            RenderRunBackground(style, segment.Resource.Embedded.Font, segment.X, segment.Width, segment.FontSize, segment.BaselineY, graphics);
        }

        DrawRunGlyphText(graphics, segment.Resource, style, segment.Text, segment.FontSize, segment.X, segment.BaselineY, color, segment.PdfCharacterSpacing, segment.CompensatePdfCharacterSpacing);
        if (!segment.IsTerminalLineSpace && segment.SyntheticBold)
        {
            DrawRunGlyphText(graphics, segment.Resource, style, segment.Text, segment.FontSize, segment.X + 0.35d, segment.BaselineY, color, segment.PdfCharacterSpacing, segment.CompensatePdfCharacterSpacing);
        }

        if (!segment.IsTerminalLineSpace)
        {
            RenderTextDecorations(style, segment.Resource.Embedded.Font, segment.X, segment.Width, segment.FontSize, segment.BaselineY, color, graphics);
        }
    }

    private static DocxRunFontResource? ResolveFontResource(DocxTextRun run, DocxFontResources fontResources)
    {
        return fontResources.RunResources.TryGetValue(run, out DocxRunFontResource? resource)
            ? resource
            : fontResources.Fallback;
    }

    private static void RenderInlineImage(
        DocxInlineImageLayout image,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        PdfImageXObject? xObject = CreateImage(image.Image, diagnosticSink, image.PageIndex);
        if (xObject is null)
        {
            return;
        }

        string imageName = "Im" + imageIndex++;
        graphics.DrawImage(imageName, image.X, image.Y, image.Width, image.Height);
        pageImages.Add(new PdfImageResource(imageName, xObject));
    }

    private static void RenderTableRow(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageNumber,
        int pageCount,
        ref int imageIndex)
    {
        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow))
            {
                continue;
            }

            DocxTableCell cell = ResolveVisualCell(cellLayout);
            if (TryResolveShadingColor(cell.FillHex, cell.ShadingValue, cell.ShadingColor, out RgbColor fill))
            {
                graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
                graphics.FillRectangle(cellLayout.X, cellLayout.Y, cellLayout.Width, cellLayout.Height);
            }
        }

        RenderTableRowBorders(row, previousRow, nextRow, graphics);

        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            if (!ShouldRenderTableCellContentFragment(cellLayout))
            {
                continue;
            }

            if (cellLayout.TextLines.Count != 0 || cellLayout.InlineImages.Count != 0)
            {
                graphics.SaveState();
                graphics.ClipRectangle(cellLayout.X, cellLayout.Y, cellLayout.Width, cellLayout.Height);
                foreach (DocxTextLineLayout line in cellLayout.TextLines)
                {
                    RenderTextLine(line, graphics, fontResources, pageNumber, pageCount);
                }

                foreach (DocxInlineImageLayout image in cellLayout.InlineImages)
                {
                    RenderInlineImage(image, graphics, pageImages, diagnosticSink, ref imageIndex);
                }

                graphics.RestoreState();
            }
        }
    }

    private static IReadOnlyList<DocxTextEmissionSegment> CreateTextEmissionSegments(
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        int pageNumber,
        int pageCount)
    {
        IReadOnlyList<DocxTextSegmentLayout> segments = line.Segments.Count == 0
            ? [new DocxTextSegmentLayout(line.Text, line.StyleRun, line.X, line.Width)]
            : line.Segments;
        var emissionSegments = new List<DocxTextEmissionSegment>(segments.Count + 1);
        foreach (DocxTextSegmentLayout segment in segments)
        {
            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            if (resource is null)
            {
                continue;
            }

            double fontSize = GetSegmentFontSize(segment, line.FontSize);
            double baselineY = GetSegmentBaselineY(segment, line.BaselineY);
            foreach (TextEmissionPart part in SplitOfficeTextOperationParts(segment, fontSize, fontResources.TextMeasurer))
            {
                emissionSegments.Add(new DocxTextEmissionSegment(
                    ResolveStaticFieldPlaceholders(part.Text, pageNumber, pageCount),
                    segment.StyleRun,
                    resource,
                    ReadColor(segment.StyleRun.ColorHex),
                    part.X,
                    baselineY,
                    part.Width,
                    fontSize,
                    segment.PdfCharacterSpacing,
                    segment.CompensatePdfCharacterSpacing,
                    ShouldApplySyntheticBold(segment.StyleRun, resource),
                    segment.StyleRun.Italic && !resource.Resolution.Italic,
                    IsTerminalLineSpace: false));
            }
        }

        if (line.EndsWithIntraTokenBreak)
        {
            return emissionSegments;
        }

        for (int i = segments.Count - 1; i >= 0; i--)
        {
            DocxTextSegmentLayout segment = segments[i];
            if (string.IsNullOrEmpty(segment.Text) || string.IsNullOrWhiteSpace(segment.Text))
            {
                continue;
            }

            if (char.IsWhiteSpace(segment.Text[^1]))
            {
                return emissionSegments;
            }

            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            if (resource is null)
            {
                return emissionSegments;
            }

            double fontSize = GetSegmentFontSize(segment, line.FontSize);
            double baselineY = GetSegmentBaselineY(segment, line.BaselineY);
            RgbColor color = ReadColor(segment.StyleRun.ColorHex);
            emissionSegments.Add(new DocxTextEmissionSegment(
                " ",
                segment.StyleRun,
                resource,
                color,
                segment.X + segment.Width,
                baselineY,
                0d,
                fontSize,
                segment.PdfCharacterSpacing,
                segment.CompensatePdfCharacterSpacing,
                SyntheticBold: false,
                SyntheticItalic: segment.StyleRun.Italic && !resource.Resolution.Italic,
                IsTerminalLineSpace: true));
            break;
        }

        return emissionSegments;
    }

    private static IReadOnlyList<TextEmissionPart> SplitOfficeTextOperationParts(
        DocxTextSegmentLayout segment,
        double fontSize,
        IDocxTextMeasurer? textMeasurer)
    {
        if (textMeasurer is null ||
            segment.Text.Length == 0 ||
            !segment.Text.Any(IsOfficeTextOperationBoundaryPunctuation))
        {
            return [new TextEmissionPart(segment.Text, segment.X, segment.Width)];
        }

        var parts = new List<TextEmissionPart>();
        int partStart = 0;
        for (int i = 0; i < segment.Text.Length; i++)
        {
            if (!IsOfficeTextOperationBoundaryPunctuation(segment.Text[i]))
            {
                continue;
            }

            AddTextEmissionPart(segment, fontSize, textMeasurer, partStart, i - partStart, parts);
            AddTextEmissionPart(segment, fontSize, textMeasurer, i, 1, parts);
            partStart = i + 1;
        }

        AddTextEmissionPart(segment, fontSize, textMeasurer, partStart, segment.Text.Length - partStart, parts);
        return parts.Count == 0 ? [new TextEmissionPart(segment.Text, segment.X, segment.Width)] : parts;
    }

    private static void AddTextEmissionPart(
        DocxTextSegmentLayout segment,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        int start,
        int length,
        List<TextEmissionPart> parts)
    {
        if (length <= 0)
        {
            return;
        }

        string prefix = start == 0 ? string.Empty : segment.Text[..start];
        string text = segment.Text.Substring(start, length);
        double x = segment.X + textMeasurer.MeasureText(segment.StyleRun, prefix, fontSize);
        double width = start + length == segment.Text.Length
            ? Math.Max(0d, segment.X + segment.Width - x)
            : textMeasurer.MeasureText(segment.StyleRun, text, fontSize);
        parts.Add(new TextEmissionPart(text, x, width));
    }

    private static bool IsOfficeTextOperationBoundaryPunctuation(char value)
    {
        return CharUnicodeInfo.GetUnicodeCategory(value) == UnicodeCategory.DashPunctuation;
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateBodyTextLines(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            switch (item)
            {
                case DocxTextLineLayout line:
                    yield return line;
                    break;
                case DocxTableRowLayout row:
                    foreach (DocxTextLineLayout cellLine in row.Cells.SelectMany(cell => cell.TextLines))
                    {
                        yield return cellLine;
                    }

                    break;
            }
        }
    }

    private static DocxTextEmissionLineSnapshot ToTextEmissionLineSnapshot(
        int pageIndex,
        bool isStaticStory,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        int pageNumber,
        int pageCount)
    {
        DocxTextEmissionSegmentSnapshot[] segments = CreateTextEmissionSegments(line, fontResources, pageNumber, pageCount)
            .Select(ToTextEmissionSegmentSnapshot)
            .ToArray();
        return new DocxTextEmissionLineSnapshot(
            pageIndex,
            isStaticStory,
            line.SourceBlockIndex,
            line.SourceLineIndex,
            line.EndsWithIntraTokenBreak,
            segments.Length,
            segments.Sum(segment => segment.TextLength),
            segments.Count(segment => segment.IsTerminalLineSpace),
            segments.Count(segment => Math.Abs(segment.PdfCharacterSpacing) > 0.0001d),
            segments);
    }

    private static DocxTextEmissionSegmentSnapshot ToTextEmissionSegmentSnapshot(DocxTextEmissionSegment segment)
    {
        return new DocxTextEmissionSegmentSnapshot(
            segment.Text.Length,
            segment.X,
            segment.BaselineY,
            segment.Width,
            segment.FontSize,
            OfficePdfTextEmissionProfile.FontSize(segment.FontSize),
            segment.StyleRun.CharacterSpacingPoints,
            segment.PdfCharacterSpacing,
            ResolvePositioningCharacterSpacing(segment.StyleRun, segment.PdfCharacterSpacing, segment.CompensatePdfCharacterSpacing),
            segment.CompensatePdfCharacterSpacing,
            segment.IsTerminalLineSpace,
            segment.Resource.Name,
            segment.SyntheticBold,
            segment.SyntheticItalic);
    }

    private static void DrawRunGlyphText(
        PdfGraphicsBuilder graphics,
        DocxRunFontResource resource,
        DocxTextRun style,
        string text,
        double fontSize,
        double x,
        double baselineY,
        RgbColor color,
        double pdfCharacterSpacing = 0d,
        bool compensatePdfCharacterSpacing = true)
    {
        bool syntheticItalic = style.Italic && !resource.Resolution.Italic;
        double pdfFontSize = OfficePdfTextEmissionProfile.FontSize(fontSize);
        double positioningCharacterSpacing = ResolvePositioningCharacterSpacing(style, pdfCharacterSpacing, compensatePdfCharacterSpacing);
        string? positioningArray = resource.Embedded.EncodeGlyphPositioningArray(text, positioningCharacterSpacing, pdfFontSize, forcePositioningArray: true);
        if (positioningArray is not null)
        {
            graphics.DrawGlyphPositionedText(resource.Name, pdfFontSize, x, baselineY, color.Red, color.Green, color.Blue, positioningArray, syntheticItalic, pdfCharacterSpacing);
            return;
        }

        graphics.DrawGlyphText(resource.Name, pdfFontSize, x, baselineY, color.Red, color.Green, color.Blue, resource.Embedded.EncodeGlyphHex(text), syntheticItalic, pdfCharacterSpacing);
    }

    private static double ResolvePositioningCharacterSpacing(DocxTextRun style, double pdfCharacterSpacing, bool compensatePdfCharacterSpacing)
    {
        return compensatePdfCharacterSpacing
            ? style.CharacterSpacingPoints - pdfCharacterSpacing
            : style.CharacterSpacingPoints;
    }

    private static bool ShouldApplySyntheticBold(DocxTextRun style, DocxRunFontResource resource)
    {
        return style.Bold && !resource.Resolution.Bold;
    }

    private static void RenderRunBackground(
        DocxTextRun style,
        OpenTypeFont font,
        double x,
        double width,
        double fontSize,
        double baselineY,
        PdfGraphicsBuilder graphics)
    {
        if (width <= 0d || !TryResolveRunBackgroundColor(style, out RgbColor color))
        {
            return;
        }

        double ascender = DocxLineMetrics.MeasureWindowsAscender(font, fontSize);
        double descender = DocxLineMetrics.MeasureWindowsDescender(font, fontSize);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        graphics.FillRectangle(x, baselineY - descender, width, ascender + descender);
    }

    private static bool TryResolveRunBackgroundColor(DocxTextRun style, out RgbColor color)
    {
        if (TryResolveHighlightColor(style.HighlightValue, out color))
        {
            return true;
        }

        return TryResolveShadingColor(style.ShadingFillHex, style.ShadingValue, style.ShadingColor, out color);
    }

    private static bool TryResolveShadingColor(string? fillHex, string? value, string? foregroundHex, out RgbColor color)
    {
        if (value is null || value.Equals("clear", StringComparison.OrdinalIgnoreCase))
        {
            return RgbColor.TryParse(fillHex, out color);
        }

        if (TryResolvePercentageShadingColor(fillHex, value, foregroundHex, out color))
        {
            return true;
        }

        color = default;
        return false;
    }

    private static bool TryResolvePercentageShadingColor(string? fillHex, string? value, string? foregroundHex, out RgbColor color)
    {
        color = default;
        if (value is null ||
            !value.StartsWith("pct", StringComparison.OrdinalIgnoreCase) ||
            !int.TryParse(value.AsSpan(3), NumberStyles.Integer, CultureInfo.InvariantCulture, out int percent) ||
            !RgbColor.TryParse(fillHex, out RgbColor background) ||
            !RgbColor.TryParse(foregroundHex, out RgbColor foreground))
        {
            return false;
        }

        double weight = Math.Clamp(percent, 0, 100) / 100d;
        color = new RgbColor(
            BlendByte(background.Red, foreground.Red, weight),
            BlendByte(background.Green, foreground.Green, weight),
            BlendByte(background.Blue, foreground.Blue, weight));
        return true;
    }

    private static byte BlendByte(byte background, byte foreground, double foregroundWeight)
    {
        return (byte)Math.Round(background * (1d - foregroundWeight) + foreground * foregroundWeight);
    }

    private static bool TryResolveHighlightColor(string? value, out RgbColor color)
    {
        switch (value)
        {
            case "black":
                color = new RgbColor(0x00, 0x00, 0x00);
                return true;
            case "blue":
                color = new RgbColor(0x00, 0x00, 0xFF);
                return true;
            case "cyan":
                color = new RgbColor(0x00, 0xFF, 0xFF);
                return true;
            case "green":
                color = new RgbColor(0x00, 0xFF, 0x00);
                return true;
            case "magenta":
                color = new RgbColor(0xFF, 0x00, 0xFF);
                return true;
            case "red":
                color = new RgbColor(0xFF, 0x00, 0x00);
                return true;
            case "yellow":
                color = new RgbColor(0xFF, 0xFF, 0x00);
                return true;
            case "white":
                color = new RgbColor(0xFF, 0xFF, 0xFF);
                return true;
            case "darkBlue":
                color = new RgbColor(0x00, 0x00, 0x80);
                return true;
            case "darkCyan":
                color = new RgbColor(0x00, 0x80, 0x80);
                return true;
            case "darkGreen":
                color = new RgbColor(0x00, 0x80, 0x00);
                return true;
            case "darkMagenta":
                color = new RgbColor(0x80, 0x00, 0x80);
                return true;
            case "darkRed":
                color = new RgbColor(0x80, 0x00, 0x00);
                return true;
            case "darkYellow":
                color = new RgbColor(0x80, 0x80, 0x00);
                return true;
            case "darkGray":
                color = new RgbColor(0x80, 0x80, 0x80);
                return true;
            case "lightGray":
                color = new RgbColor(0xC0, 0xC0, 0xC0);
                return true;
            default:
                color = default;
                return false;
        }
    }

    private static void RenderTextDecorations(
        DocxTextRun style,
        OpenTypeFont font,
        double x,
        double width,
        double fontSize,
        double baselineY,
        RgbColor color,
        PdfGraphicsBuilder graphics)
    {
        if (width <= 0d)
        {
            return;
        }

        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        if (style.Underline)
        {
            double thickness = ResolveDecorationThickness(font.Post.UnderlineThickness, font, fontSize);
            double y = baselineY + font.Post.UnderlinePosition * fontSize / font.UnitsPerEm;
            graphics.FillRectangle(x, y - thickness / 2d, width, thickness);
        }

        if (style.Strike || style.DoubleStrike)
        {
            double thickness = ResolveDecorationThickness(font.Os2.StrikeoutSize, font, fontSize);
            double y = baselineY + font.Os2.StrikeoutPosition * fontSize / font.UnitsPerEm;
            if (style.DoubleStrike)
            {
                double offset = Math.Max(thickness, fontSize / 18d);
                graphics.FillRectangle(x, y - offset - thickness / 2d, width, thickness);
                graphics.FillRectangle(x, y + offset - thickness / 2d, width, thickness);
            }
            else
            {
                graphics.FillRectangle(x, y - thickness / 2d, width, thickness);
            }
        }
    }

    private static double ResolveDecorationThickness(short metricValue, OpenTypeFont font, double fontSize)
    {
        return Math.Max(0.25d, Math.Abs(metricValue) * fontSize / font.UnitsPerEm);
    }

    private static void RenderTableRowBorders(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics)
    {
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCellLayout cellLayout = row.Cells[cellIndex];
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow))
            {
                continue;
            }

            if (previousRow is null)
            {
                RenderHorizontalTableCellBorder(cellLayout, "top", graphics);
            }

            if (nextRow is null)
            {
                RenderHorizontalTableCellBorder(cellLayout, "bottom", graphics);
            }

            DocxTableCell visualCell = ResolveVisualCell(cellLayout);
            DocxTableCellBorder? left = DocxTableBorderGeometry.Find(visualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "start");
            if (cellIndex == 0)
            {
                RenderVerticalTableCellBorder(cellLayout.X, cellLayout.Y, cellLayout.Height, left, graphics);
            }

            DocxTableCellBorder? right = DocxTableBorderGeometry.Find(visualCell.Borders, "right") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "end");
            if (cellIndex == row.Cells.Count - 1)
            {
                RenderVerticalTableCellBorder(cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.Height, right, graphics);
                continue;
            }

            DocxTableCellLayout nextCell = row.Cells[cellIndex + 1];
            DocxTableCell nextVisualCell = ResolveVisualCell(nextCell);
            DocxTableCellBorder? nextLeft = DocxTableBorderGeometry.Find(nextVisualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(nextVisualCell.Borders, "start");
            RenderSharedVerticalTableBorder(cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.Height, right, nextLeft, graphics);
        }

        if (nextRow is not null && nextRow.RowIndex != row.RowIndex)
        {
            RenderSharedHorizontalTableBorders(row, nextRow, graphics);
        }
    }

    private static void RenderSharedHorizontalTableBorders(
        DocxTableRowLayout row,
        DocxTableRowLayout nextRow,
        PdfGraphicsBuilder graphics)
    {
        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow: null))
            {
                continue;
            }

            DocxTableCellLayout[] overlappingNextCells = nextRow.Cells
                .Where(nextCell => ShouldRenderTableCellVisualFragment(nextCell, row) && HorizontalOverlap(cellLayout, nextCell) > 0d)
                .ToArray();
            if (overlappingNextCells.Length == 0)
            {
                RenderHorizontalTableCellBorder(cellLayout, "bottom", graphics);
                continue;
            }

            foreach (DocxTableCellLayout nextRowCell in overlappingNextCells)
            {
                RenderSharedHorizontalTableBorderSegment(cellLayout, nextRowCell, graphics);
            }
        }
    }

    private static void RenderSharedHorizontalTableBorderSegment(
        DocxTableCellLayout cellLayout,
        DocxTableCellLayout nextRowCell,
        PdfGraphicsBuilder graphics)
    {
        DocxTableCell cell = ResolveVisualCell(cellLayout);
        DocxTableCell nextCell = ResolveVisualCell(nextRowCell);
        DocxTableCellBorder? bottom = DocxTableBorderGeometry.Find(cell.Borders, "bottom");
        DocxTableCellBorder? nextTop = DocxTableBorderGeometry.Find(nextCell.Borders, "top");
        if (DocxTableBorderGeometry.IsSuppressed(bottom) || DocxTableBorderGeometry.IsSuppressed(nextTop))
        {
            return;
        }

        DocxTableCellBorder? border = DocxTableBorderGeometry.SelectStronger(bottom, nextTop);
        if (border is null)
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        double x = Math.Max(cellLayout.X, nextRowCell.X);
        double right = Math.Min(cellLayout.X + cellLayout.Width, nextRowCell.X + nextRowCell.Width);
        if (right <= x)
        {
            return;
        }

        double leftBorderWidth = ResolveLeftVerticalBorderWidth(cellLayout);
        double segmentX = Math.Min(right, x + leftBorderWidth);
        if (right <= segmentX)
        {
            return;
        }

        graphics.FillRectangle(segmentX, cellLayout.Y - width / 2d, right - segmentX, width);
    }

    private static double HorizontalOverlap(DocxTableCellLayout first, DocxTableCellLayout second)
    {
        return Math.Min(first.X + first.Width, second.X + second.Width) - Math.Max(first.X, second.X);
    }

    private static void RenderHorizontalTableCellBorder(DocxTableCellLayout cellLayout, string edge, PdfGraphicsBuilder graphics)
    {
        DocxTableCell visualCell = ResolveVisualCell(cellLayout);
        DocxTableCellBorder? border = DocxTableBorderGeometry.Find(visualCell.Borders, edge);
        if (border is null || DocxTableBorderGeometry.IsSuppressed(border))
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        switch (edge)
        {
            case "top":
                double topX = cellLayout.X + ResolveLeftVerticalBorderWidth(cellLayout);
                double topWidth = cellLayout.Width - (topX - cellLayout.X);
                if (topWidth > 0d)
                {
                    graphics.FillRectangle(topX, cellLayout.Y + cellLayout.Height - width, topWidth, width);
                }
                break;
            case "bottom":
                double bottomX = cellLayout.X + ResolveLeftVerticalBorderWidth(cellLayout);
                double bottomWidth = cellLayout.Width - (bottomX - cellLayout.X);
                if (bottomWidth > 0d)
                {
                    graphics.FillRectangle(bottomX, cellLayout.Y, bottomWidth, width);
                }
                break;
        }
    }

    private static double ResolveLeftVerticalBorderWidth(DocxTableCellLayout cellLayout)
    {
        DocxTableCell visualCell = ResolveVisualCell(cellLayout);
        DocxTableCellBorder? left = DocxTableBorderGeometry.Find(visualCell.Borders, "left") ??
            DocxTableBorderGeometry.Find(visualCell.Borders, "start");
        return DocxTableBorderGeometry.IsSuppressed(left)
            ? 0d
            : DocxTableBorderGeometry.ResolveVisibleWidth(left);
    }

    private static void RenderSharedVerticalTableBorder(
        double boundaryX,
        double y,
        double height,
        DocxTableCellBorder? leftCellRight,
        DocxTableCellBorder? rightCellLeft,
        PdfGraphicsBuilder graphics)
    {
        if (DocxTableBorderGeometry.IsSuppressed(leftCellRight) || DocxTableBorderGeometry.IsSuppressed(rightCellLeft))
        {
            return;
        }

        DocxTableCellBorder? border = DocxTableBorderGeometry.SelectStronger(leftCellRight, rightCellLeft);
        if (border is null)
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        graphics.FillRectangle(boundaryX, y, width, height);
    }

    private static void RenderVerticalTableCellBorder(
        double boundaryX,
        double y,
        double height,
        DocxTableCellBorder? border,
        PdfGraphicsBuilder graphics)
    {
        if (border is null || DocxTableBorderGeometry.IsSuppressed(border))
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        graphics.FillRectangle(boundaryX, y, width, height);
    }

    private static bool ShouldRenderTableCellContentFragment(DocxTableCellLayout cellLayout)
    {
        return !cellLayout.IsVerticalMergeContinuation;
    }

    private static bool ShouldRenderTableCellVisualFragment(
        DocxTableCellLayout cellLayout,
        DocxTableRowLayout? previousRow)
    {
        if (!cellLayout.IsVerticalMergeContinuation)
        {
            return true;
        }

        if (cellLayout.VerticalMergeOwnerCell is null)
        {
            return false;
        }

        return previousRow is null || !previousRow.Cells.Any(previousCell =>
            !previousCell.IsVerticalMergeContinuation &&
            HorizontalOverlap(previousCell, cellLayout) > 0d &&
            previousCell.Y <= cellLayout.Y + 0.001d &&
            previousCell.Y + previousCell.Height >= cellLayout.Y + cellLayout.Height - 0.001d);
    }

    private static DocxTableCell ResolveVisualCell(DocxTableCellLayout cellLayout)
    {
        return cellLayout.VerticalMergeOwnerCell ?? cellLayout.Cell;
    }

    private static string ResolveStaticFieldPlaceholders(string text, int pageNumber, int pageCount)
    {
        return text
            .Replace("{NUMPAGES}", pageCount.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
            .Replace("{PAGE}", pageNumber.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
    }

    private static RgbColor ReadColor(string? hex)
    {
        return RgbColor.TryParse(hex, out RgbColor color) ? color : new RgbColor(0, 0, 0);
    }

    private static PdfImageXObject? CreateImage(DocxInlineImage image, Action<OoxPdfDiagnostic>? diagnosticSink, int pageIndex)
    {
        try
        {
            if (image.ContentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                image.ContentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase))
            {
                JpegInfo info = JpegInfo.Read(image.Bytes);
                return PdfImageXObject.Jpeg(info.Width, info.Height, image.Bytes, info.ComponentCount, info.BitsPerComponent);
            }

            if (image.ContentType.Equals("image/png", StringComparison.OrdinalIgnoreCase))
            {
                PngImage png = PngImage.Read(image.Bytes);
                return PdfImageXObject.RgbPng(png.Width, png.Height, png.Rgb, png.Alpha);
            }

            if (image.ContentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) ||
                image.ContentType.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase))
            {
                BmpImage bmp = BmpImage.Read(image.Bytes);
                return PdfImageXObject.RgbPng(bmp.Width, bmp.Height, bmp.Rgb, bmp.Alpha);
            }
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
        {
            EmitImageDiagnostic(diagnosticSink, image, pageIndex, ex.Message);
            return null;
        }

        EmitImageDiagnostic(diagnosticSink, image, pageIndex, "Unsupported image content type.");
        return null;
    }

    private static void EmitImageDiagnostic(Action<OoxPdfDiagnostic>? diagnosticSink, DocxInlineImage image, int pageIndex, string reason)
    {
        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            "IMAGE_UNSUPPORTED_FORMAT",
            OoxPdfSeverity.Error,
            $"Image '{image.ContentType}' could not be rendered and was ignored: {reason}",
            image.PartName,
            PageIndex: pageIndex,
            Feature: image.ContentType,
            Fallback: "Ignored"));
    }

    private readonly record struct TextEmissionPart(string Text, double X, double Width);
}
