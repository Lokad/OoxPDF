namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    // Office A/B (tbxrev probe, Word-COM rendered): on scaled Word-compatible pages the
    // floating-textbox frame stays unscaled page geometry while its design-space content
    // maps uniformly to emission space (X about the page origin, Y about the page center).
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
            return designY * PrintScale + (PageHeight / 2d) * (1d - PrintScale);
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

        public DocxInlineImageLayout PrecompensateImage(DocxInlineImageLayout image, double originX, double originY, int pageIndex)
        {
            // Image dimensions stay file values, like body images on scaled pages; only the
            // position joins the uniform map.
            return image with
            {
                X = PrecompensateX(originX + image.X),
                Y = PrecompensateY(originY + image.Y),
                PageIndex = pageIndex
            };
        }

        public DocxTableRowLayout PrecompensateRow(DocxTableRowLayout row)
        {
            // Rows arrive translated to absolute design space (TranslateTableRow first); only
            // positions and text extents join the map, row/cell heights keep file behavior.
            FloatingTextBoxEmissionMap self = this;
            return row with
            {
                Table = row.Table with { TableX = self.PrecompensateX(row.Table.TableX) },
                Y = self.PrecompensateY(row.Y),
                Cells = row.Cells
                    .Select(cell => cell with
                    {
                        X = self.PrecompensateX(cell.X),
                        Y = self.PrecompensateY(cell.Y),
                        TextLines = cell.TextLines
                            .Select(cellLine => self.PrecompensateLine(cellLine, 0d, 0d))
                            .ToArray(),
                        InlineImages = cell.InlineImages
                            .Select(cellImage => cellImage with
                            {
                                X = self.PrecompensateX(cellImage.X),
                                Y = self.PrecompensateY(cellImage.Y)
                            })
                            .ToArray()
                    })
                    .ToArray()
            };
        }

        public (double X, double Y, double Width, double Height) MapClipRectangle(
            double contentX,
            double contentTop,
            double contentWidth,
            double contentHeight)
        {
            // The clip feeds graphics directly (no body offsets apply), so it maps to
            // emission space, not pre-compensated space.
            double top = MapEmissionY(contentTop);
            double height = ScaleExtent(contentHeight);
            return (MapEmissionX(contentX), top - height, ScaleExtent(contentWidth), height);
        }
    }
}