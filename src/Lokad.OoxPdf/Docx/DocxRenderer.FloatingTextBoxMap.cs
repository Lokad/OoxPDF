namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static DocxLayout ReframeFloatingParagraphTops(DocxLayout layout, DocxMarkupContext context,
        CancellationToken cancellationToken)
    {
        if (layout.FloatingDrawings.Count == 0 || layout.Pages.Count == 0 ||
            !FloatingTextBoxEmissionMap.TryCreate(context, layout.Pages[0].Height, out _)) return layout;
        DocxFloatingDrawingLayout[]? changed = null;
        bool separateHeaders = CanSeparatePlainHeaderOrigins(layout);
        int contextPage = -1;
        var pageContext = context;
        for (int index = 0; index < layout.FloatingDrawings.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var drawing = layout.FloatingDrawings[index];
            // Top-offset anchors follow the measured printed paragraph, not a
            // second scaling of its already spaced body coordinates. Wrapped,
            // fragmented, non-body and unmeasured owners retain their fallback.
            if (drawing.Story is not null || drawing.TextBoxLayout is null || drawing.WrapExclusionTop is not null ||
                drawing.PageStartIndex != drawing.PageEndIndex ||
                !string.Equals(drawing.Drawing.VerticalRelativeFromValue, "paragraph", StringComparison.OrdinalIgnoreCase) ||
                drawing.VerticalPlacementSource != DocxAnchorPlacementSource.Offset ||
                drawing.AnchorTextLine is not { BodyLineBoxBaselineInsetPoints: double printedInset } line ||
                drawing.AnchorPageIndex is not int pageIndex || drawing.VerticalReferenceTop is not double oldTop ||
                drawing.AnchorBlockVerticalBottom is not double blockBottom || drawing.PlacedTop is not double placedTop)
                continue;
            if (contextPage != pageIndex)
            {
                pageContext = WithPageTextEmissionOffsets(context, layout, pageIndex, separateHeaders);
                contextPage = pageIndex;
            }
            var frame = new DocxPrintedPageFrame(layout.Pages[pageIndex].Height, pageContext.WordCompatiblePrintScale);
            double top = frame.UnmapY(line.BaselineY - pageContext.WordCompatibleTextYOffset + printedInset);
            double bottom = frame.UnmapY(blockBottom - pageContext.WordCompatibleTextYOffset);
            if (top == oldTop && bottom == drawing.VerticalReferenceBottom) continue;
            changed ??= layout.FloatingDrawings.ToArray();
            changed[index] = drawing with
            {
                VerticalReferenceTop = top, VerticalReferenceBottom = bottom,
                PlacedTop = placedTop + top - oldTop
            };
        }
        return changed is null ? layout : layout with { FloatingDrawings = changed };
    }

    // Floating placement stays in design coordinates. On scaled Word-compatible
    // pages, the frame and its content map uniformly to emission space (X about
    // the page origin, Y about the page center), as in retained Word PDF controls.
    // Body text instead keeps the first-pin shift, which strands textbox content ~18pt low
    // and below its own content-box clip (invisible). Pre-compensating design coordinates
    // into body-layout space lets every existing offset consumer (render, balloon anchors,
    // revision bar, snapshots, bookmarks, links) land mapped without touching emission math.
    // Unit-scale and non-WC pages keep the legacy path (TryCreate returns false).
    internal readonly record struct FloatingTextBoxEmissionMap(double PrintScale, double XOffset, double YOffset, double PageHeight)
    {
        public static bool TryCreate(DocxMarkupContext markupContext, double pageHeight, out FloatingTextBoxEmissionMap map)
        {
            if (markupContext.Mode != OoxPdfDocxMarkupMode.AllMarkup ||
                markupContext.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ||
                !markupContext.ExpandsMarkupMargin ||
                Math.Abs(markupContext.WordCompatiblePrintScale - 1d) < 0.000000001d)
            {
                map = default;
                return false;
            }

            map = new FloatingTextBoxEmissionMap(markupContext.WordCompatiblePrintScale, markupContext.WordCompatibleTextXOffset, markupContext.WordCompatibleTextYOffset, pageHeight);
            return true;
        }

        public double MapEmissionX(double designX)
        {
            return designX * PrintScale;
        }

        public double MapEmissionY(double designY)
        {
            return new DocxPrintedPageFrame(PageHeight, PrintScale).MapY(designY);
        }

        public double PrecompensateX(double designX)
        {
            return MapEmissionX(designX) - XOffset;
        }

        public double PrecompensateY(double designY)
        {
            return MapEmissionY(designY) + YOffset;
        }

        public double ScaleExtent(double designExtent)
        {
            return designExtent * PrintScale;
        }

        public DocxTextLineLayout PrecompensateLine(DocxTextLineLayout line, double originX, double originY)
        {
            // Struct lambdas cannot capture this (CS1673); work through a copy.
            FloatingTextBoxEmissionMap self = this;
            return line with
            {
                X = self.PrecompensateX(originX + line.X),
                BaselineY = self.PrecompensateY(originY + line.BaselineY),
                Width = self.ScaleExtent(line.Width),
                FrameOriginOffsetX = line.FrameOriginOffsetX * self.PrintScale,
                ParagraphShadingFrame = line.ParagraphShadingFrame?.Scale(self.PrintScale),
                MathScale = line.MathScale * self.PrintScale,
                InlineMathOffsetX = line.InlineMathOffsetX * self.PrintScale,
                InlineObjectEndOffsetX = line.InlineObjectEndOffsetX * self.PrintScale,
                InlineMath = line.InlineMath.Select(math => math with { OffsetX = math.OffsetX * self.PrintScale }).ToArray(),
                InlineFrame = line.InlineFrame is { } frame
                    ? frame.Scale(self.PrintScale) : null,
                OccupiedLineBox = line.OccupiedLineBox is { } occupied
                    ? new(occupied.Height * self.PrintScale, occupied.BaselineInset * self.PrintScale) : null,
                InlineSourceExtents = line.InlineSourceExtents.Select(extent => extent.Scale(self.PrintScale)).ToArray(),
                Segments = line.Segments
                    .Select(segment => segment with
                    {
                        X = self.PrecompensateX(originX + segment.X),
                        Width = self.ScaleExtent(segment.Width),
                        BaselineOffsetY = segment.BaselineOffsetY * self.PrintScale
                    })
                    .ToArray()
            };
        }

        public DocxInlineImageLayout PrecompensateImage(DocxInlineImageLayout image, int pageIndex)
        {
            // Design-space pictures, their clipping and their occupied line
            // frame join the same printed map as text and container bounds.
            return image with
            {
                X = PrecompensateX(image.X), Y = PrecompensateY(image.Y),
                Width = ScaleExtent(image.Width), Height = ScaleExtent(image.Height), PageIndex = pageIndex,
                ParagraphShadingFrame = image.ParagraphShadingFrame?.Scale(PrintScale),
                ClipBounds = image.ClipBounds is { } clip
                    ? new(PrecompensateX(clip.X), PrecompensateY(clip.Y), ScaleExtent(clip.Width), ScaleExtent(clip.Height)) : null,
                LineFrame = image.LineFrame is { } frame
                    ? new(PrecompensateY(frame.Bottom), ScaleExtent(frame.Height)) { ImageGrowth = ScaleExtent(frame.ImageGrowth) } : null
            };
        }

        public DocxInlineImageLayout PrecompensateImage(DocxInlineImageLayout image, double originX, double originY, int pageIndex)
        {
            return PrecompensateImage(image.Translate(originX, originY), pageIndex);
        }

        public DocxInlineTextBoxLayout PrecompensateBox(DocxInlineTextBoxLayout box)
        {
            FloatingTextBoxEmissionMap self = this;
            return box with
            {
                BoxX = self.PrecompensateX(box.BoxX),
                BoxTop = self.PrecompensateY(box.BoxTop),
                BoxWidth = self.ScaleExtent(box.BoxWidth),
                BoxHeight = self.ScaleExtent(box.BoxHeight),
                PrintScale = box.PrintScale * self.PrintScale,
                TextLines = box.TextLines.Select(line => self.PrecompensateLine(line, 0d, 0d)).ToArray(),
                InlineImages = box.InlineImages.Select(image => self.PrecompensateImage(image, image.PageIndex)).ToArray(),
                TableRows = box.TableRows.Select(self.PrecompensateRow).ToArray()
            };
        }

        public DocxTableRowLayout PrecompensateRow(DocxTableRowLayout row)
        {
            // Rows arrive translated to absolute design space. Paint bounds and
            // descendants join the same map as glyphs; unscaled cell extents can
            // otherwise cover the preceding row or clip a nested source.
            FloatingTextBoxEmissionMap self = this;
            return row with
            {
                Table = row.Table with
                {
                    TableX = self.PrecompensateX(row.Table.TableX),
                    ResolvedTableWidth = self.ScaleExtent(row.Table.ResolvedTableWidth),
                    AlignedMinimumRowBottomMargins = row.Table.AlignedMinimumRowBottomMargins.Select(self.ScaleExtent).ToArray()
                },
                Y = self.PrecompensateY(row.Y),
                Height = self.ScaleExtent(row.Height),
                Cells = row.Cells
                    .Select(cell => cell with
                    {
                        X = self.PrecompensateX(cell.X),
                        Y = self.PrecompensateY(cell.Y),
                        Width = self.ScaleExtent(cell.Width),
                        Height = self.ScaleExtent(cell.Height),
                        ContentPaddingLeft = self.ScaleExtent(cell.ContentPaddingLeft),
                        ContentPaddingRight = self.ScaleExtent(cell.ContentPaddingRight),
                        ContentPaddingTop = self.ScaleExtent(cell.ContentPaddingTop),
                        ContentPaddingBottom = self.ScaleExtent(cell.ContentPaddingBottom),
                        TextLines = cell.TextLines
                            .Select(cellLine => self.PrecompensateLine(cellLine, 0d, 0d))
                            .ToArray(),
                        InlineImages = cell.InlineImages
                            .Select(cellImage => self.PrecompensateImage(cellImage, cellImage.PageIndex))
                            .ToArray(),
                        InlineTextBoxes = cell.InlineTextBoxes.Select(self.PrecompensateBox).ToArray(),
                        NestedTableRows = cell.NestedRows.Select(self.PrecompensateRow).ToArray()
                    })
                    .ToArray()
            };
        }

        public (double X, double Y, double Width, double Height) MapRectangle(
            double x,
            double top,
            double width,
            double height)
        {
            // Graphics consumes physical coordinates without body text offsets.
            double physicalHeight = ScaleExtent(height);
            return (MapEmissionX(x), MapEmissionY(top) - physicalHeight, ScaleExtent(width), physicalHeight);
        }
    }
}
