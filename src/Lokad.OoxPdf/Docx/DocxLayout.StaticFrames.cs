namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static bool HasCompleteStaticContent(DocxStaticStoryLayoutResult? story,
        IReadOnlyList<DocxBodyElement> source, DocxLayoutPage page, double bottom, double top, CancellationToken token)
    {
        HashSet<DocxEquation>? remaining = null;
        foreach (var paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(source))
        foreach (var equation in paragraph.Equations)
        {
            token.ThrowIfCancellationRequested();
            if (!(remaining ??= new(ReferenceEqualityComparer.Instance)).Add(equation)) return false;
        }
        if (story is null) return remaining is null;
        // Containers supply their existing measured block edges. Margin
        // independence is a body-metrics policy, not a note-placement boundary.
        if (!double.IsFinite(bottom) || !double.IsFinite(top) || bottom < 0d || top > page.Height || bottom > top ||
            story.InlineTextBoxes.Count != 0 ||
            story.InlineImages.Any(image => !DocxPlainBodyFrame.HasCompleteImageFrame(image, page))) return false;
        foreach (var line in story.TextLines)
            if (!Line(line, null)) return false;
        if (story.TableRows.Count != 0 &&
            (!Rows(story.TableRows, null) || !DocxTableCellClipGeometry.HasCompleteMeasuredRows(story.TableRows, token, Line))) return false;
        return remaining is null || remaining.Count == 0;

        bool Line(DocxTextLineLayout line, DocxTableCellClipSnapshot? clip)
        {
            token.ThrowIfCancellationRequested();
            foreach (var segment in line.Segments)
                if (segment.Width != 0d && !DocxTableCellClipGeometry.ContainsSegment(clip, segment)) return false;
            if (!double.IsFinite(line.X) || !double.IsFinite(line.Width) || line.Width < 0d ||
                line.X < 0d || line.X + line.Width > page.Width) return false;
            if (line.Math is null && line.InlineMath.Count == 0) return true;
            if (!DocxInlineLineGeometry.TryGetPreparedMathBounds(line, out double lineBottom, out double lineTop) ||
                lineBottom < 0d || lineTop > page.Height) return false;
            if (line.Math is { } standalone && !Consume(standalone,
                line.X + (line.IsInlineMath ? line.InlineMathOffsetX : 0d), line.MathScale, clip)) return false;
            foreach (var placement in line.InlineMath)
                if (!Consume(placement.Math, line.X + placement.OffsetX, line.MathScale, clip)) return false;
            return true;
        }

        bool Consume(DocxPreparedMathParagraph math, double x, double scale, DocxTableCellClipSnapshot? clip)
        {
            double width = Math.Max(math.Box.Width, math.Box.AdvanceWidth) * scale;
            return double.IsFinite(x) && x >= 0d && double.IsFinite(scale) && scale > 0d &&
                double.IsFinite(width) && width >= 0d && x + width <= page.Width &&
                DocxTableCellClipGeometry.ContainsHorizontalRange(clip, x, width) &&
                remaining?.Remove(math.Equation) == true;
        }

        bool Rows(IReadOnlyList<DocxTableRowLayout> rows, DocxTableCellClipSnapshot? parent)
        {
            foreach (var row in rows)
            {
                token.ThrowIfCancellationRequested();
                if (row.FragmentCount != 1 || row.HeightRuleValue == "exact" ||
                    !double.IsFinite(row.Y) || !double.IsFinite(row.Height) || row.Height <= 0d ||
                    !double.IsFinite(row.Table.TableX) || !double.IsFinite(row.Table.ResolvedTableWidth) ||
                    row.Table.ResolvedTableWidth <= 0d || row.Table.TableX < 0d ||
                    row.Table.TableX + row.Table.ResolvedTableWidth > page.Width ||
                    row.Y < 0d || row.Y + row.Height > page.Height ||
                    !DocxTableCellClipGeometry.ContainsHorizontalRange(parent, row.Table.TableX, row.Table.ResolvedTableWidth) ||
                    !DocxTableCellClipGeometry.ContainsVerticalRange(parent, row.Y, row.Y + row.Height)) return false;
                foreach (var cell in row.Cells)
                {
                    token.ThrowIfCancellationRequested();
                    if (cell.VisualOwnership != DocxTableCellVisualOwnership.OwnCell || cell.InlineTextBoxes.Count != 0) return false;
                    var clip = DocxTableCellClipGeometry.Resolve(row, cell, parent);
                    foreach (var image in cell.InlineImages)
                        if (!DocxPlainBodyFrame.HasCompleteImageFrame(image, page) ||
                            !DocxTableCellClipGeometry.ContainsHorizontalRange(clip, image.X, image.Width) ||
                            !DocxTableCellClipGeometry.ContainsVerticalRange(clip, image.Y, image.Y + image.Height)) return false;
                    if (!Rows(cell.NestedRows, clip)) return false;
                }
            }
            return true;
        }
    }

    private static DocxStaticStoryLayoutResult CreatePageStaticStoryLayout(
        DocxLayoutPage page, DocxSelectedStaticStory story, double x, double width, double startY,
        bool isHeader, int pageNumber, int pageCount, IDocxTextMeasurer textMeasurer,
        IDocxStaticTextMetricsProvider staticMetrics, double defaultTabStopPoints, double scale,
        IDocxLineMetricsProvider? unscaledLineMetrics, CancellationToken token,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations,
        IDocxTextMeasurer? unscaledTextMeasurer, bool hasDocumentBodyOrigin = false)
    {
        bool canMap = scale is > 0d and < 1d && double.IsFinite(width / scale) &&
            page.ColumnFrames.Count <= 1 &&
            (page.Items.FirstOrDefault() is DocxTextLineLayout || page.Items.Count == 0 && hasDocumentBodyOrigin) &&
            unscaledTextMeasurer is IDocxStaticTextMetricsProvider &&
            // Each story owns its measured frame. A counterpart's fallback
            // cannot revoke physical placement of otherwise supported content;
            // floating drawings already own independent page/margin frames.
            CanUseStaticStoryFrame(story, width / scale, mathEquations, token);
        if (!canMap)
        {
            var layout = CreateStaticStoryLayout(story, x, width, startY, isHeader, pageNumber, pageCount,
                textMeasurer, staticMetrics, defaultTabStopPoints, scale, unscaledLineMetrics, token, mathEquations);
            if (!isHeader && scale == 1d && CanUseStaticStoryFrame(story, width, mathEquations, token))
            {
                double top = ResolveStaticFooterContentTop(layout, unscaledLineMetrics);
                layout = layout with { PhysicalFooterFrameTop = double.IsFinite(top) ? top : null };
                return PlaceFooterTableDrawingFrames(layout, page.MarginBottom, token);
            }
            return layout;
        }

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
                ParagraphShadingFrame = line.ParagraphShadingFrame?.Scale(scale),
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
                OccupiedLineBox = line.OccupiedLineBox is { } occupied
                    ? new(occupied.Height * scale, occupied.BaselineInset * scale) : null,
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
                ParagraphShadingFrame = image.ParagraphShadingFrame?.Scale(scale),
                ClipBounds = image.ClipBounds is { } clip ? clip with
                {
                    X = x + (clip.X - x) * scale, Y = center + (clip.Y - center) * scale,
                    Width = clip.Width * scale, Height = clip.Height * scale
                } : null,
                LineFrame = image.LineFrame is { } frame ? new(center + (frame.Bottom - center) * scale, frame.Height * scale) { ImageGrowth = frame.ImageGrowth * scale } : null
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
                        TextLines = cell.TextLines.Select(MapLine).ToArray(), InlineImages = cell.InlineImages.Select(MapImage).ToArray(),
                        NestedTableRows = cell.NestedRows.Select(MapRow).ToArray(),
                        InlineTextBoxes = cell.InlineTextBoxes.Select(MapBox).ToArray()
                    };
                }).ToArray()
            };
        }
        DocxInlineTextBoxLayout MapBox(DocxInlineTextBoxLayout box)
        {
            token.ThrowIfCancellationRequested();
            return box with
            {
                BoxX = x + (box.BoxX - x) * scale,
                BoxTop = center + (box.BoxTop - center) * scale,
                BoxWidth = box.BoxWidth * scale, BoxHeight = box.BoxHeight * scale,
                PrintScale = box.PrintScale * scale, StaticStoryUsesPhysicalY = true,
                TextLines = box.TextLines.Select(MapLine).ToArray(),
                InlineImages = box.InlineImages.Select(MapImage).ToArray(),
                TableRows = box.TableRows.Select(MapRow).ToArray()
            };
        }
        double footerTop = isHeader ? double.NegativeInfinity :
            ResolveStaticFooterContentTop(source, unscaledLineMetrics);
        var mapped = source with
        {
            TextLines = source.TextLines.Select(MapLine).ToArray(),
            InlineImages = source.InlineImages.Select(MapImage).ToArray(),
            TableRows = source.TableRows.Select(MapRow).ToArray(),
            InlineTextBoxes = source.InlineTextBoxes.Select(MapBox).ToArray(),
            EndCursorY = frameAnchor + (source.EndCursorY - frameAnchor) * scale,
            EndPendingAfterSpacing = source.EndPendingAfterSpacing * scale,
            MappedFooterContentTop = isHeader ? null : frameAnchor + (footerTop - frameAnchor) * scale,
            PhysicalFooterFrameTop = double.IsFinite(footerTop) ? center + (footerTop - center) * scale : null,
            PhysicalHeaderFrameBottom = isHeader
                ? center + (source.EndCursorY - source.EndPendingAfterSpacing - center) * scale : null
        };
        return isHeader ? mapped : PlaceFooterTableDrawingFrames(mapped,
            center + (page.MarginBottom - center) * scale, token);
    }

    // A complete source frame can enter its drawing viewport without preceding
    // its measured paragraph start. Any tail that cannot fit keeps its clip.
    // Independent wrapped/primary-line Word controls qualify these two bounds;
    // explicit breaks and unknown owners retain their existing placement.
    private static DocxStaticStoryLayoutResult PlaceFooterTableDrawingFrames(
        DocxStaticStoryLayoutResult layout, double bottom, CancellationToken token)
    {
        if (!double.IsFinite(bottom) || !layout.TableRows.Any(row => row.Cells.Any(cell =>
            cell.InlineTextBoxes.Any(box => box.BoxTop - box.BoxHeight < bottom)))) return layout;
        return layout with
        {
            TableRows = layout.TableRows.Select(row =>
            {
                token.ThrowIfCancellationRequested();
                return row with { Cells = row.Cells.Select(cell =>
                {
                    token.ThrowIfCancellationRequested();
                    if (!cell.InlineTextBoxes.Any(box => box.BoxTop - box.BoxHeight < bottom)) return cell;
                    var starts = new Dictionary<int, (DocxParagraph Paragraph, double Top)>();
                    foreach (var line in cell.TextLines)
                    {
                        token.ThrowIfCancellationRequested();
                        if (line.SourceParagraphIndex is not { } index || line.SourceLineIndex != 0 ||
                            line.SourceParagraph is not { InlineContent: { } items } paragraph ||
                            line.InlineFrame is not { } frame || items.Any(item => item is DocxInlineBreak)) continue;
                        double top = line.BaselineY + frame.BaselineInset;
                        if (!starts.TryAdd(index, (paragraph, top))) starts[index] = (paragraph, double.NaN);
                    }
                    return cell with { InlineTextBoxes = cell.InlineTextBoxes.Select(box =>
                    {
                        token.ThrowIfCancellationRequested();
                        if (box.BoxTop - box.BoxHeight >= bottom) return box;
                        if (box.SourceParagraphIndex is { } index && starts.TryGetValue(index, out var start) &&
                            double.IsFinite(start.Top) && start.Top >= box.BoxTop &&
                            start.Paragraph.InlineTextBoxes.Any(source => ReferenceEquals(source, box.TextBox)))
                        {
                            double admittedTop = Math.Min(start.Top, bottom + box.BoxHeight);
                            box = ShiftInlineTextBox(box, admittedTop - box.BoxTop, 0d);
                        }
                        return box with { OutlineViewportBottom = bottom };
                    }).ToArray() };
                }).ToArray() };
            }).ToArray()
        };
    }

    private static bool CanUseStaticStoryFrame(DocxSelectedStaticStory story, double width,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations, CancellationToken token)
        => CanUseStaticBodyFrame(story.BodyElements, width, mathEquations, token);

    private static bool CanUseStaticBodyFrame(IReadOnlyList<DocxBodyElement> elements, double width,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations, CancellationToken token)
    {
        foreach (var element in elements)
        {
            token.ThrowIfCancellationRequested();
            if (element is DocxParagraphElement direct &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(direct.Paragraph, token, direct.Revisions))
            {
                if (!CanUseStaticParagraphFrame(direct.Paragraph, width, mathEquations, token)) return false;
            }
            else if (element is DocxTableElement table)
            {
                if (!CanUseStaticTableFrame(table.Table, width, mathEquations, token)) return false;
            }
            else { return false; }
        }
        return true;
    }

    private static bool CanUseStaticTableFrame(DocxTable table, double width,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations, CancellationToken token)
    {
        if (table.Revisions.Count != 0) return false;
        foreach (var row in table.Rows)
        {
            token.ThrowIfCancellationRequested();
            if (!DocxMathParagraphPolicy.IsSupportedSimpleRow(row)) return false;
            foreach (var cell in row.Cells)
            {
                token.ThrowIfCancellationRequested();
                if (!DocxMathParagraphPolicy.IsSupportedCellProperties(cell)) return false;
                foreach (var child in DocxTableCellContent.GetBodyElements(cell))
                {
                    token.ThrowIfCancellationRequested();
                    if (child is DocxParagraphElement paragraph)
                    {
                        if (!CanUseStaticParagraphFrame(paragraph.Paragraph, width, mathEquations, token)) return false;
                    }
                    else if (child is DocxTableElement nested)
                    {
                        if (!CanUseStaticTableFrame(nested.Table, width, mathEquations, token)) return false;
                    }
                    else return false;
                }
            }
        }
        return true;
    }

    private static bool CanUseStaticParagraphFrame(DocxParagraph paragraph, double width,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (!paragraph.InlineTextBoxes.All(box => CanUseInlineBoxFrame(box, width, mathEquations, token))) return false;
        // Legacy box-only models may lack a parsed ordered view. Mixed parsed
        // paragraphs use the same source ownership as body and cell containers.
        if (paragraph.InlineTextBoxes.Count != 0 && paragraph.Equations.Count == 0 &&
            paragraph.Images.Count == 0 && paragraph.ListLabel is null &&
            paragraph.Runs.All(run => run.EffectiveProperties.Hidden || run.Text.Length == 0))
            return DocxInlineContent.HasSupportedSourceMetadata(paragraph, token: token);
        if (paragraph.Equations.Count != 0)
            return (TryGetStandaloneMath(paragraph, width, 1d, mathEquations, out _) ||
                mathEquations is not null && paragraph.InlineContent is not null &&
                DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasPageInputs: true).Count == paragraph.Equations.Count &&
                // Unavailable math keeps its diagnosed fallback; the remaining
                // ordered source items still own a static frame.
                paragraph.Equations.All(equation => !mathEquations.TryGetValue(equation, out var math) || math.Box.AdvanceWidth <= width) &&
                paragraph.Images.All(image => image.WidthPoints <= width));
        return DocxMathParagraphPolicy.HasSupportedOrderedAlignment(paragraph) &&
            paragraph.Images.All(image => image.SourceRunIndex >= 0 && image.RotationDegrees == 0d &&
                double.IsFinite(image.WidthPoints) && double.IsFinite(image.HeightPoints) &&
                image.WidthPoints > 0d && image.HeightPoints > 0d && image.WidthPoints <= width) &&
            // A finite fixed line box retains its declared advance in the same
            // static frame as automatic/minimum flow, including boxed children.
            (paragraph.InlineTextBoxes.Count == 0 ?
                DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph) || DocxPlainBodyFrame.HasFixedLineFlow(paragraph) :
                paragraph.ListLabel is null && DocxPlainBodyFrame.HasAutomaticLineFlow(paragraph)) &&
            // Static placement resolves page inputs; boxed ordered content
            // separately verifies its actual counter dependencies at measurement.
            DocxInlineContent.HasSupportedSourceOwner(paragraph, hasPageInputs: true, token: token);
    }

    private static bool CanUseInlineBoxFrame(DocxInlineTextBox box, double width,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (box.Revisions.Count != 0 || ReadEmuPoints(box.ExtentCxValue) is not { } boxWidth || boxWidth <= 0d || boxWidth > width ||
            ReadEmuPoints(box.ExtentCyValue) is not > 0d || box.BodyElements.Count == 0) return false;
        // Nested boxes retain their existing placement until their own source
        // frames are supported. Tables inside one independent box may still map.
        foreach (var paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(box.BodyElements))
        {
            token.ThrowIfCancellationRequested();
            if (paragraph.InlineTextBoxes.Count != 0) return false;
        }
        ResolveInlineTextBoxContentInsets(box, out double left, out _, out double right, out _);
        double contentWidth = boxWidth - left - right;
        return contentWidth > 0d && CanUseStaticBodyFrame(box.BodyElements, contentWidth, mathEquations, token);
    }
}
