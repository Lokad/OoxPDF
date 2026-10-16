namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static DocxStaticStoryLayoutResult CreatePageStaticStoryLayout(
        DocxLayoutPage page, DocxSelectedStaticStory story, double x, double width, double startY,
        bool isHeader, int pageNumber, int pageCount, IDocxTextMeasurer textMeasurer,
        IDocxStaticTextMetricsProvider staticMetrics, double defaultTabStopPoints, double scale,
        IDocxLineMetricsProvider? unscaledLineMetrics, CancellationToken token,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations,
        IDocxTextMeasurer? unscaledTextMeasurer)
    {
        bool canMap = scale is > 0d and < 1d && double.IsFinite(width / scale) &&
            page.ColumnFrames.Count <= 1 &&
            page.Items.FirstOrDefault() is DocxTextLineLayout &&
            unscaledTextMeasurer is IDocxStaticTextMetricsProvider &&
            page.PageSettings.HeaderFloatingDrawingsByType.Values.All(drawings => drawings.Count == 0) &&
            page.PageSettings.FooterFloatingDrawingsByType.Values.All(drawings => drawings.Count == 0) &&
            IsNaturalStaticStory(story, width / scale, mathEquations, token) &&
            IsNaturalStaticStory(SelectStaticHeaderFooter(
                isHeader ? page.PageSettings.FooterBodyElementsByType : page.PageSettings.HeaderBodyElementsByType,
                isHeader ? page.PageSettings.FooterParagraphsByType : page.PageSettings.HeaderParagraphsByType,
                page.PageSettings, pageNumber), width / scale, mathEquations, token);
        if (!canMap)
            return CreateStaticStoryLayout(story, x, width, startY, isHeader, pageNumber, pageCount,
                textMeasurer, staticMetrics, defaultTabStopPoints, scale, unscaledLineMetrics, token, mathEquations);

        // Lay out one source block using the existing natural metrics. Mapping
        // its complete frame keeps the first inset, pitches and footer anchor
        // in the same coordinate system, including prepared mathematical boxes.
        var source = CreateStaticStoryLayout(story, x, width / scale, startY, isHeader, pageNumber, pageCount,
            unscaledTextMeasurer!, (IDocxStaticTextMetricsProvider)unscaledTextMeasurer!,
            defaultTabStopPoints, 1d, unscaledLineMetrics, token, mathEquations);
        if (!HasVisibleStaticContent(source)) { return source; }
        double center = page.Height / 2d;
        double frameAnchor = isHeader ? page.Height - page.MarginTop : page.MarginBottom;
        DocxTextLineLayout MapLine(DocxTextLineLayout line)
        {
            token.ThrowIfCancellationRequested();
            return line with
            {
                X = x + (line.X - x) * scale,
                BaselineY = center + (line.BaselineY - center) * scale,
                Width = line.Width * scale,
                FrameOriginOffsetX = line.FrameOriginOffsetX * scale,
                Segments = line.Segments.Select(segment =>
                {
                    token.ThrowIfCancellationRequested();
                    return segment with { X = x + (segment.X - x) * scale, Width = segment.Width * scale };
                }).ToArray(),
                LineHeight = line.LineHeight * scale,
                AppliedBeforeSpacing = line.AppliedBeforeSpacing * scale,
                PendingAfterSpacing = line.PendingAfterSpacing * scale,
                ParagraphBeforeSpacing = line.ParagraphBeforeSpacing * scale,
                ParagraphAfterSpacing = line.ParagraphAfterSpacing * scale,
                MathScale = line.MathScale * scale,
                InlineMath = line.InlineMath.Select(placement =>
                {
                    token.ThrowIfCancellationRequested();
                    return placement with { OffsetX = placement.OffsetX * scale };
                }).ToArray(),
                InlineFrame = line.InlineFrame is { } frame ? frame.Scale(scale) : null,
                InlineObjectEndOffsetX = line.InlineObjectEndOffsetX * scale,
                InlineSourceExtents = line.InlineSourceExtents.Select(extent =>
                {
                    token.ThrowIfCancellationRequested();
                    return extent.Scale(scale);
                }).ToArray(),
                StaticStoryUsesPhysicalY = true
            };
        }
        DocxInlineImageLayout MapImage(DocxInlineImageLayout image)
        {
            token.ThrowIfCancellationRequested();
            return image with
            {
                X = x + (image.X - x) * scale, Y = center + (image.Y - center) * scale,
                Width = image.Width * scale, Height = image.Height * scale, StaticStoryUsesPhysicalY = true,
                ClipBounds = image.ClipBounds is { } clip ? clip with
                {
                    X = x + (clip.X - x) * scale, Y = center + (clip.Y - center) * scale,
                    Width = clip.Width * scale, Height = clip.Height * scale
                } : null,
                LineFrame = image.LineFrame is { } frame ? new(center + (frame.Bottom - center) * scale, frame.Height * scale) : null
            };
        }
        var mappedTables = new Dictionary<DocxTableLayoutContext, DocxTableLayoutContext>(ReferenceEqualityComparer.Instance);
        DocxTableRowLayout MapRow(DocxTableRowLayout row)
        {
            token.ThrowIfCancellationRequested();
            if (!mappedTables.TryGetValue(row.Table, out var table))
            {
                table = row.Table with
                {
                    TableX = x + (row.Table.TableX - x) * scale,
                    GridColumnsWidthSum = row.Table.GridColumnsWidthSum * scale,
                    ResolvedTableWidth = row.Table.ResolvedTableWidth * scale,
                    ResolvedColumnWidths = row.Table.ResolvedColumnWidths.Select(value => value * scale).ToArray(),
                    AlignedMinimumRowBottomMargins = row.Table.AlignedMinimumRowBottomMargins.Select(value => value * scale).ToArray()
                };
                mappedTables.Add(row.Table, table);
            }
            return row with
            {
                Y = center + (row.Y - center) * scale, Height = row.Height * scale,
                FullRowHeight = row.FullRowHeight * scale, FragmentOffsetFromRowTop = row.FragmentOffsetFromRowTop * scale,
                StaticStoryUsesPhysicalY = true, ReviewBorderPrintScale = scale,
                Table = table,
                Cells = row.Cells.Select(cell =>
                {
                    token.ThrowIfCancellationRequested();
                    return cell with
                    {
                        X = x + (cell.X - x) * scale, Y = center + (cell.Y - center) * scale,
                        Width = cell.Width * scale, Height = cell.Height * scale,
                        ContentPaddingLeft = cell.ContentPaddingLeft * scale, ContentPaddingRight = cell.ContentPaddingRight * scale,
                        ContentPaddingTop = cell.ContentPaddingTop * scale, ContentPaddingBottom = cell.ContentPaddingBottom * scale,
                        TextLines = cell.TextLines.Select(MapLine).ToArray(), InlineImages = cell.InlineImages.Select(MapImage).ToArray()
                    };
                }).ToArray()
            };
        }
        double footerTop = isHeader ? double.NegativeInfinity :
            ResolveStaticFooterContentTop(source, unscaledLineMetrics);
        return source with
        {
            TextLines = source.TextLines.Select(MapLine).ToArray(),
            InlineImages = source.InlineImages.Select(MapImage).ToArray(),
            TableRows = source.TableRows.Select(MapRow).ToArray(),
            EndCursorY = frameAnchor + (source.EndCursorY - frameAnchor) * scale,
            EndPendingAfterSpacing = source.EndPendingAfterSpacing * scale,
            MappedFooterContentTop = isHeader ? null : frameAnchor + (footerTop - frameAnchor) * scale,
            PhysicalFooterFrameTop = double.IsFinite(footerTop) ? center + (footerTop - center) * scale : null
        };
    }

    private static bool IsNaturalStaticStory(DocxSelectedStaticStory story, double width,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations, CancellationToken token)
    {
        foreach (var element in story.BodyElements)
        {
            token.ThrowIfCancellationRequested();
            IEnumerable<DocxParagraph> paragraphs;
            if (element is DocxParagraphElement direct &&
                DocxInlineContent.HasSupportedRunRevisionOwners(direct.Revisions, direct.Paragraph.Runs, token))
                { paragraphs = [direct.Paragraph]; }
            else if (element is DocxTableElement table && table.Table.Revisions.Count == 0 &&
                table.Table.Rows.All(row =>
                {
                    token.ThrowIfCancellationRequested();
                    return DocxMathParagraphPolicy.IsSupportedSimpleRow(row) && row.Cells.All(cell =>
                    {
                        token.ThrowIfCancellationRequested();
                        return DocxMathParagraphPolicy.IsSupportedSimpleCell(cell);
                    });
                }))
                paragraphs = table.Table.Rows.SelectMany(row => row.Cells).SelectMany(cell => cell.Paragraphs);
            else { return false; }
            foreach (DocxParagraph paragraph in paragraphs)
            {
                token.ThrowIfCancellationRequested();
                if (paragraph.Equations.Count != 0)
                {
                    if (!TryGetStandaloneMath(paragraph, width, 1d, mathEquations, out _) &&
                        (mathEquations is null || paragraph.InlineContent is null ||
                        DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasPageInputs: true).Count != paragraph.Equations.Count ||
                        paragraph.Equations.Any(equation => !mathEquations.TryGetValue(equation, out var math) || math.Box.AdvanceWidth > width) ||
                        paragraph.Images.Any(image => image.WidthPoints > width))) { return false; }
                }
                else if (paragraph.Alignment is not (DocxTextAlignment.Left or DocxTextAlignment.Center or DocxTextAlignment.Right) ||
                    paragraph.Images.Count != 0 || !DocxPlainBodyFrame.HasAutomaticLineFlow(paragraph) ||
                    !DocxInlineContent.HasSupportedSourceOwner(paragraph, token: token)) { return false; }
            }
        }
        return true;
    }
}
