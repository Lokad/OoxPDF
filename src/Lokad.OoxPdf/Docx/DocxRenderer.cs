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
        IReadOnlyDictionary<string, PdfLinkDestination> bookmarkDestinations = CreateBookmarkDestinations(layout, fontResources);
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

            IReadOnlyList<PdfLinkAnnotation> annotations = CreateHyperlinkAnnotations(layoutPage, fontResources, pageNumber, layout.Pages.Count, bookmarkDestinations);
            pages.Add(new PdfPage(layoutPage.Width, layoutPage.Height, graphics.ToString(), fontResources.Resources, pageImages.ToArray(), [], [], [], annotations));
        }

        return pages;
    }

    private static IReadOnlyList<PdfLinkAnnotation> CreateHyperlinkAnnotations(
        DocxLayoutPage page,
        DocxFontResources fontResources,
        int pageNumber,
        int pageCount,
        IReadOnlyDictionary<string, PdfLinkDestination> bookmarkDestinations)
    {
        var annotations = new List<PdfLinkAnnotation>();
        foreach (DocxTextLineLayout line in page.StaticTextLines.Concat(EnumerateBodyTextLines(page)))
        {
            if (line.SourceParagraph is not { } paragraph ||
                paragraph.Hyperlinks.Count == 0)
            {
                continue;
            }

            IReadOnlyList<DocxHyperlinkSpan> links = paragraph.Hyperlinks;
            foreach (DocxTextEmissionSegment segment in CreateTextEmissionSegments(line, fontResources, pageNumber, pageCount))
            {
                if (segment.IsTerminalLineSpace || segment.SourceTextRunIndex < 0 || segment.Width <= 0d)
                {
                    continue;
                }

                DocxHyperlinkSpan? link = links.FirstOrDefault(item => IsHyperlinkSegment(item, segment.SourceTextRunIndex));
                if (link is null)
                {
                    continue;
                }

                double ascender = segment.Resource.Embedded.Font.Os2.WindowsAscender * segment.FontSize / segment.Resource.Embedded.Font.UnitsPerEm;
                double descender = segment.Resource.Embedded.Font.Os2.WindowsDescender * segment.FontSize / segment.Resource.Embedded.Font.UnitsPerEm;
                if (IsExternalHyperlink(link))
                {
                    annotations.Add(PdfLinkAnnotation.ToUri(
                        segment.X,
                        segment.BaselineY - descender,
                        segment.Width,
                        ascender + descender,
                        link.Target!));
                }
                else if (!string.IsNullOrEmpty(link.Anchor) &&
                    bookmarkDestinations.TryGetValue(link.Anchor, out PdfLinkDestination destination))
                {
                    annotations.Add(PdfLinkAnnotation.ToDestination(
                        segment.X,
                        segment.BaselineY - descender,
                        segment.Width,
                        ascender + descender,
                        destination));
                }
            }
        }

        return annotations;
    }

    private static IReadOnlyDictionary<string, PdfLinkDestination> CreateBookmarkDestinations(
        DocxLayout layout,
        DocxFontResources fontResources)
    {
        var destinations = new Dictionary<string, PdfLinkDestination>(StringComparer.Ordinal);
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            DocxLayoutPage page = layout.Pages[pageIndex];
            int pageNumber = pageIndex + 1;
            foreach (DocxTextLineLayout line in page.StaticTextLines.Concat(EnumerateBodyTextLines(page)))
            {
                if (line.SourceParagraph is not { } paragraph ||
                    paragraph.BookmarkAnchors.Count == 0)
                {
                    continue;
                }

                IReadOnlyList<DocxTextEmissionSegment> segments = CreateTextEmissionSegments(line, fontResources, pageNumber, layout.Pages.Count)
                    .Where(segment => !segment.IsTerminalLineSpace && segment.SourceTextRunIndex >= 0 && segment.Width > 0d)
                    .ToArray();
                if (segments.Count == 0)
                {
                    continue;
                }

                foreach (DocxBookmarkAnchor bookmark in paragraph.BookmarkAnchors)
                {
                    if (string.IsNullOrEmpty(bookmark.Name) || destinations.ContainsKey(bookmark.Name))
                    {
                        continue;
                    }

                    DocxTextEmissionSegment? target = segments.FirstOrDefault(segment => segment.SourceTextRunIndex >= bookmark.TextRunIndex);
                    if (target is null)
                    {
                        continue;
                    }

                    double ascender = target.Resource.Embedded.Font.Os2.WindowsAscender * target.FontSize / target.Resource.Embedded.Font.UnitsPerEm;
                    destinations[bookmark.Name] = new PdfLinkDestination(
                        pageIndex,
                        target.X,
                        target.BaselineY + ascender,
                        Zoom: null);
                }
            }
        }

        return destinations;
    }

    private static bool IsHyperlinkSegment(DocxHyperlinkSpan link, int sourceTextRunIndex)
    {
        return sourceTextRunIndex >= link.TextRunStartIndex &&
            sourceTextRunIndex < link.TextRunStartIndex + link.TextRunCount;
    }

    private static bool IsExternalHyperlink(DocxHyperlinkSpan link)
    {
        return !string.IsNullOrEmpty(link.Target) &&
            string.Equals(link.TargetMode, "External", StringComparison.OrdinalIgnoreCase);
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
            foreach (DocxTextEmissionPart part in DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, fontSize, fontResources.TextMeasurer))
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
                    segment.PdfCharacterSpacingSource,
                    segment.CompensatePdfCharacterSpacing,
                    ShouldApplySyntheticBold(segment.StyleRun, resource),
                    segment.StyleRun.Italic && !resource.Resolution.Italic,
                    IsTerminalLineSpace: false,
                    segment.SourceTextRunIndex,
                    segment.Role));
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
                PdfCharacterSpacing: 0d,
                PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.TerminalLineSpace,
                CompensatePdfCharacterSpacing: true,
                SyntheticBold: false,
                SyntheticItalic: segment.StyleRun.Italic && !resource.Resolution.Italic,
                IsTerminalLineSpace: true,
                segment.SourceTextRunIndex,
                segment.Role));
            break;
        }

        return emissionSegments;
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
            .Select(segment => ToTextEmissionSegmentSnapshot(segment, line))
            .ToArray();
        return new DocxTextEmissionLineSnapshot(
            pageIndex,
            isStaticStory,
            line.SourceBlockIndex,
            line.SourceParagraphIndex,
            line.SourceLineIndex,
            line.EndsWithIntraTokenBreak,
            segments.Length,
            segments.Sum(segment => segment.TextLength),
            segments.Count(segment => segment.IsTerminalLineSpace),
            segments.Count(segment => Math.Abs(segment.PdfCharacterSpacing) > 0.0001d),
            segments);
    }

    private static DocxTextEmissionSegmentSnapshot ToTextEmissionSegmentSnapshot(
        DocxTextEmissionSegment segment,
        DocxTextLineLayout line)
    {
        DocxTextEmissionPlan plan = segment.IsTerminalLineSpace
            ? DocxTextEmissionPlanner.CreateTerminalLineSpace(segment.StyleRun, segment.FontSize)
            : DocxTextEmissionPlanner.Create(
                segment.StyleRun,
                segment.FontSize,
                segment.PdfCharacterSpacing,
                segment.CompensatePdfCharacterSpacing,
                segment.PdfCharacterSpacingSource);
        return new DocxTextEmissionSegmentSnapshot(
            segment.Text.Length,
            line.SourceBlockIndex,
            line.SourceParagraphIndex,
            line.SourceLineIndex,
            segment.Role.ToString(),
            segment.X,
            segment.BaselineY,
            segment.Width,
            segment.FontSize,
            plan.PdfFontSize,
            segment.StyleRun.CharacterSpacingPoints,
            plan.PdfCharacterSpacing,
            plan.PdfCharacterSpacingSource.ToString(),
            plan.PositioningCharacterSpacing,
            plan.CompensatePdfCharacterSpacing,
            DocxTextEmissionPlanner.ClassifyText(segment.Text),
            DocxTextEmissionPlanner.MeasureAdvanceProfile(segment.Text, segment.Resource.Embedded, segment.Width, plan),
            DocxTextEmissionPlanner.CreateGlyphAdvanceSignature(segment.Text, segment.Resource.Embedded),
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
        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.Create(style, fontSize, pdfCharacterSpacing, compensatePdfCharacterSpacing);
        string? positioningArray = resource.Embedded.EncodeGlyphPositioningArray(text, plan.PositioningCharacterSpacing, plan.PdfFontSize, forcePositioningArray: true);
        if (positioningArray is not null)
        {
            graphics.DrawGlyphPositionedText(resource.Name, plan.PdfFontSize, x, baselineY, color.Red, color.Green, color.Blue, positioningArray, syntheticItalic, plan.PdfCharacterSpacing);
            return;
        }

        graphics.DrawGlyphText(resource.Name, plan.PdfFontSize, x, baselineY, color.Red, color.Green, color.Blue, resource.Embedded.EncodeGlyphHex(text), syntheticItalic, plan.PdfCharacterSpacing);
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
        return cellLayout.VisualOwnership == DocxTableCellVisualOwnership.OwnCell;
    }

    private static bool ShouldRenderTableCellVisualFragment(
        DocxTableCellLayout cellLayout,
        DocxTableRowLayout? previousRow)
    {
        if (cellLayout.VisualOwnership == DocxTableCellVisualOwnership.OwnCell)
        {
            return true;
        }

        if (cellLayout.VisualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner)
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

}
