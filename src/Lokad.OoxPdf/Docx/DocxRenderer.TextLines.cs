using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
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
                    foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                    {
                        yield return cellLine;
                    }

                    break;
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateRenderedPageTextLines(
        FloatingDrawingPageIndex.PageIndexPair drawingPages,
        DocxLayoutPage page,
        int pageIndex,
        DocxMarkupContext markupContext,
        double pageHeight)
    {
        // R12: page drawings arrive from the once-per-pass page index instead of
        // re-filtering both drawing lists for every page.
        FloatingTextBoxEmissionMap? map = TryCreateFloatingTextBoxEmissionMap(markupContext, pageHeight);
        return EnumerateStaticTextLines(page)
            .Concat(EnumerateBodyTextLines(page))
            .Concat(EnumeratePlacedRelatedStoryTextLines(page))
            .Concat(EnumerateInlineTextBoxTextLines(page))
            .Concat(EnumerateMappedFloatingDrawingTextBoxTextLines(drawingPages.PageAll(pageIndex), map))
            .Concat(page.PlacedRelatedStories.SelectMany(story => EnumerateMappedFloatingDrawingTextBoxTextLines(story.FloatingDrawings, map)));
    }

    private static FloatingTextBoxEmissionMap? TryCreateFloatingTextBoxEmissionMap(DocxMarkupContext markupContext, double pageHeight)
    {
        return FloatingTextBoxEmissionMap.TryCreate(markupContext, pageHeight, out FloatingTextBoxEmissionMap map)
            ? map
            : null;
    }

    // Mapped legs translate by the inset content origin (like the renderer) and
    // pre-compensate design coordinates into body-layout space on scaled WC pages; every
    // other mode keeps the legacy frame-origin translation bit-identically.
    // Box-only walk for comment suppression: unlike EnumerateTableRowTextLines it yields
    // no regular cell lines, so body-cell comments keep ballooning while box-anchored
    // ones suppress by paragraph identity.
    private static IEnumerable<DocxTextLineLayout> EnumerateTableRowTextBoxLines(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            if (item is not DocxTableRowLayout row)
            {
                continue;
            }

            foreach (DocxTextLineLayout line in EnumerateTableRowTextBoxLines(row))
            {
                yield return line;
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateTableRowTextBoxLines(DocxTableRowLayout row)
    {
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            foreach (DocxInlineTextBoxLayout box in cell.InlineTextBoxes)
            {
                foreach (DocxTextLineLayout boxLine in box.TextLines)
                {
                    yield return boxLine;
                }

                foreach (DocxTableRowLayout boxRow in box.TableRows)
                {
                    foreach (DocxTextLineLayout boxCellLine in EnumerateTableRowTextLines(boxRow))
                    {
                        yield return boxCellLine;
                    }

                    foreach (DocxTextLineLayout nestedBoxLine in EnumerateTableRowTextBoxLines(boxRow))
                    {
                        yield return nestedBoxLine;
                    }
                }
            }

            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                foreach (DocxTextLineLayout nestedBoxLine in EnumerateTableRowTextBoxLines(nestedRow))
                {
                    yield return nestedBoxLine;
                }
            }
        }
    }
    private static IEnumerable<DocxTextLineLayout> EnumerateMappedFloatingDrawingTextBoxTextLines(
        IEnumerable<DocxFloatingDrawingLayout> drawings,
        FloatingTextBoxEmissionMap? map)
    {
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            if (drawing.PlacedX is not { } placedX ||
                drawing.PlacedTop is not { } placedTop ||
                drawing.TextBoxLayout is not { } textBoxLayout)
            {
                continue;
            }

            if (map is not { } emissionMap)
            {
                foreach (DocxTextLineLayout line in textBoxLayout.TextLines)
                {
                    yield return TranslateTextLine(line, placedX, placedTop);
                }

                foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
                {
                    foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                    {
                        yield return TranslateTextLine(cellLine, placedX, placedTop);
                    }
                }

                continue;
            }

            DocxLayoutEngine.ResolveTextBoxContentInsets(drawing.Drawing, out double insetLeft, out double insetTop, out _, out _);
            double contentX = placedX + insetLeft;
            double contentTop = placedTop - insetTop;
            foreach (DocxTextLineLayout line in textBoxLayout.TextLines)
            {
                yield return emissionMap.PrecompensateLine(line, contentX, contentTop);
            }

            foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
            {
                foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                {
                    yield return emissionMap.PrecompensateLine(TranslateTextLine(cellLine, contentX, contentTop), 0d, 0d);
                }
            }
        }
    }

    // Office A/B (w6-tbxctl probe): Word never balloons body-flow floating-textbox
    // comments; static header/footer floatings keep the legacy path (unprobed).
    internal static bool IsStaticStoryFloatingDrawing(DocxFloatingDrawingLayout drawing)
    {
        return drawing.Story?.Kind is DocxStoryKind.Header or DocxStoryKind.Footer;
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateMarkupBalloonAnchorTextLines(
        DocxLayoutPage page,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        DocxMarkupContext markupContext,
        double pageHeight)
    {
        FloatingTextBoxEmissionMap? map = TryCreateFloatingTextBoxEmissionMap(markupContext, pageHeight);
        return EnumerateStaticTextLines(page)
            .Concat(EnumerateBodyTextLines(page))
            .Concat(EnumeratePlacedRelatedStoryTextLines(page))
            .Concat(EnumerateInlineTextBoxTextLines(page))
            .Concat(EnumerateMappedFloatingDrawingTextBoxTextLines(floatingDrawings, map))
            .Concat(page.PlacedRelatedStories.SelectMany(story => EnumerateMappedFloatingDrawingTextBoxTextLines(story.FloatingDrawings, map)));
    }

    // Inline-textbox content lives in absolute flow space (placed at layout time), so
    // anchor/snapshot/link consumers take box lines directly with the body offsets.
    private static IEnumerable<DocxTextLineLayout> EnumerateInlineTextBoxTextLines(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            if (item is not DocxInlineTextBoxLayout box)
            {
                continue;
            }

            foreach (DocxTextLineLayout line in box.TextLines)
            {
                yield return line;
            }

            foreach (DocxTableRowLayout row in box.TableRows)
            {
                foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                {
                    yield return cellLine;
                }
            }
        }
    }
    private sealed record DocxTextEmissionLineSource(
        DocxTextLineLayout Line,
        bool IsStaticStory,
        string StoryKind,
        string? StoryVariantType,
        string? ContainerStoryKind,
        string? ContainerStoryVariantType);

    private static IEnumerable<DocxTextEmissionLineSource> EnumerateRenderedFloatingDrawingTextBoxTextLines(
        FloatingDrawingPageIndex.PageIndexPair drawingPages,
        DocxLayoutPage page,
        int pageIndex,
        DocxMarkupContext markupContext,
        double pageHeight)
    {
        FloatingTextBoxEmissionMap? map = TryCreateFloatingTextBoxEmissionMap(markupContext, pageHeight);
        foreach (DocxFloatingDrawingLayout drawing in drawingPages.PageAll(pageIndex))
        {
            bool isStaticStory = IsStaticStoryFloatingDrawing(drawing);
            foreach (DocxTextLineLayout line in EnumerateMappedFloatingDrawingTextBoxTextLines([drawing], map))
            {
                yield return new DocxTextEmissionLineSource(
                    line,
                    isStaticStory,
                    "TextBox",
                    line.Story?.VariantType,
                    drawing.Story?.ToKindString() ?? "Body",
                    drawing.Story?.VariantType);
            }
        }

        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTextLineLayout line in EnumerateMappedFloatingDrawingTextBoxTextLines(story.FloatingDrawings, map))
            {
                yield return new DocxTextEmissionLineSource(
                    line,
                    IsStaticStory: false,
                    "TextBox",
                    line.Story?.VariantType,
                    story.StoryLayout.Story.Kind.ToValueString(),
                    story.StoryLayout.Story.Id);
            }
        }

        foreach (DocxTextLineLayout line in EnumerateInlineTextBoxTextLines(page))
        {
            yield return new DocxTextEmissionLineSource(
                line,
                IsStaticStory: false,
                "TextBox",
                line.Story?.VariantType,
                "Body",
                null);
        }
    }

    private static IEnumerable<DocxLayoutItem> EnumerateStaticLayoutItems(DocxLayoutPage page)
    {
        return page.StaticTextLines
            .Cast<DocxLayoutItem>()
            .Concat(page.StaticInlineImages)
            .Concat(page.StaticInlineTextBoxes)
            .Concat(page.StaticTableRows)
            .OrderByDescending(item => item switch
            {
                DocxTextLineLayout textLine => textLine.BaselineY,
                DocxInlineImageLayout image => image.Y + image.Height,
                DocxTableRowLayout row => row.Y + row.Height,
                _ => 0d
            });
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateStaticTextLines(DocxLayoutPage page)
    {
        return EnumerateStaticTextLines(page, includeTextBoxes: true);
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateStaticTextLines(DocxLayoutPage page, bool includeTextBoxes)
    {
        foreach (DocxTextLineLayout line in page.StaticTextLines)
        {
            yield return line;
        }

        if (includeTextBoxes)
        {
            foreach (DocxTextLineLayout boxLine in EnumerateStaticTextBoxLines(page))
            {
                yield return boxLine;
            }
        }

        foreach (DocxTableRowLayout row in page.StaticTableRows)
        {
            foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
            {
                yield return cellLine;
            }
        }
    }

    // Box-only static walk for comment suppression: static text comments keep
    // ballooning while static-box ones suppress by paragraph identity.
    private static IEnumerable<DocxTextLineLayout> EnumerateStaticTextBoxLines(DocxLayoutPage page)
    {
        foreach (DocxInlineTextBoxLayout box in page.StaticInlineTextBoxes)
        {
            foreach (DocxTextLineLayout boxLine in box.TextLines)
            {
                yield return boxLine;
            }

            foreach (DocxTableRowLayout boxRow in box.TableRows)
            {
                foreach (DocxTextLineLayout boxCellLine in EnumerateTableRowTextLines(boxRow))
                {
                    yield return boxCellLine;
                }
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumeratePlacedRelatedStoryTextLines(DocxLayoutPage page)
    {
        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTextLineLayout line in story.TextLines)
            {
                yield return line;
            }

            foreach (DocxTableRowLayout row in story.TableRows)
            {
                foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                {
                    yield return cellLine;
                }
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateFloatingDrawingTextBoxTextLines(DocxFloatingDrawingLayout drawing)
    {
        if (drawing.PlacedX is not { } placedX ||
            drawing.PlacedTop is not { } placedTop ||
            drawing.TextBoxLayout is not { } textBoxLayout)
        {
            yield break;
        }

        foreach (DocxTextLineLayout line in textBoxLayout.TextLines)
        {
            yield return TranslateTextLine(line, placedX, placedTop);
        }

        foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
        {
            foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
            {
                yield return TranslateTextLine(cellLine, placedX, placedTop);
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateFloatingDrawingTextBoxTableRows(IEnumerable<DocxFloatingDrawingLayout> drawings)
    {
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            if (drawing.PlacedX is not { } placedX ||
                drawing.PlacedTop is not { } placedTop ||
                drawing.TextBoxLayout is not { } textBoxLayout)
            {
                continue;
            }

            foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(TranslateTableRow(row, placedX, placedTop)))
                {
                    yield return nested;
                }
            }
        }
    }

    private static DocxTableRowLayout TranslateTableRow(DocxTableRowLayout row, double deltaX, double deltaY)
    {
        return row with
        {
            Table = row.Table with
            {
                TableX = row.Table.TableX + deltaX
            },
            Y = row.Y + deltaY,
            Cells = row.Cells
                .Select(cell => TranslateTableCell(cell, deltaX, deltaY))
                .ToArray()
        };
    }

    private static DocxTableCellLayout TranslateTableCell(DocxTableCellLayout cell, double deltaX, double deltaY)
    {
        return cell with
        {
            X = cell.X + deltaX,
            Y = cell.Y + deltaY,
            TextLines = cell.TextLines
                .Select(line => TranslateTextLine(line, deltaX, deltaY))
                .ToArray(),
            InlineImages = cell.InlineImages
                .Select(image => image with { X = image.X + deltaX, Y = image.Y + deltaY })
                .ToArray(),
            NestedTableRows = cell.NestedRows
                .Select(row => TranslateTableRow(row, deltaX, deltaY))
                .ToArray()
        };
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateTableRowTextLines(DocxTableRowLayout row)
    {
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            foreach (DocxTextLineLayout line in cell.TextLines)
            {
                yield return line;
            }

            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                foreach (DocxTextLineLayout nestedLine in EnumerateTableRowTextLines(nestedRow))
                {
                    yield return nestedLine;
                }
            }

            foreach (DocxInlineTextBoxLayout box in cell.InlineTextBoxes)
            {
                foreach (DocxTextLineLayout boxLine in box.TextLines)
                {
                    yield return boxLine;
                }

                foreach (DocxTableRowLayout boxRow in box.TableRows)
                {
                    foreach (DocxTextLineLayout boxCellLine in EnumerateTableRowTextLines(boxRow))
                    {
                        yield return boxCellLine;
                    }
                }
            }
        }
    }

    private static DocxTextEmissionLineSnapshot ToTextEmissionLineSnapshot(
        int pageIndex,
        bool isStaticStory,
        string storyKind,
        string? storyVariantType,
        string? containerStoryKind,
        string? containerStoryVariantType,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        int pageNumber,
        int pageCount,
        double fontScale,
        double baselineOffsetY,
        double xOffset,
        bool suppressCommentReferenceSpacer,
        bool useWordCompatibleTextProfile)
    {
        DocxTextEmissionSegmentSnapshot[] segments = CreateTextEmissionSegments(line, fontResources, pageNumber, pageCount, fontScale, baselineOffsetY, xOffset, suppressCommentReferenceSpacer, useWordCompatibleTextProfile)
            .Select(segment => ToTextEmissionSegmentSnapshot(segment, line))
            .ToArray();
        return new DocxTextEmissionLineSnapshot(
            pageIndex,
            isStaticStory,
            storyKind,
            storyVariantType,
            containerStoryKind,
            containerStoryVariantType,
            line.SourceBlockIndex,
            line.SourceParagraphIndex,
            line.SourceLineIndex,
            line.EndsWithIntraTokenBreak,
            segments.Length,
            segments.Sum(segment => segment.TextLength),
            segments.Count(segment => segment.RevisionKind is not null),
            segments.Count(segment => segment.RevisionKind == "Insertion"),
            segments.Count(segment => segment.RevisionKind == "Deletion"),
            segments.Count(segment => segment.RevisionKind == "MoveFrom"),
            segments.Count(segment => segment.RevisionKind == "MoveTo"),
            segments.Count(segment => segment.RevisionKind is not null &&
                segment.RevisionKind != "Insertion" &&
                segment.RevisionKind != "Deletion" &&
                segment.RevisionKind != "MoveFrom" &&
                segment.RevisionKind != "MoveTo"),
            line.SourceParagraph?.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment) ?? 0,
            segments.Count(segment => segment.IsTerminalLineSpace),
            segments.Count(segment => Math.Abs(segment.PdfCharacterSpacing) > 0.0001d),
            segments);
    }

    private static DocxTextEmissionSegmentSnapshot ToTextEmissionSegmentSnapshot(
        DocxTextEmissionSegment segment,
        DocxTextLineLayout line)
    {
        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            segment.StyleRun,
            segment.FontSize,
            segment.PdfCharacterSpacing,
            segment.PdfCharacterSpacingSource,
            segment.CompensatePdfCharacterSpacing,
            segment.IsTerminalLineSpace);
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
            segment.StyleRun.EffectiveProperties.CharacterSpacingPoints,
            plan.PdfCharacterSpacing,
            plan.PdfCharacterSpacingSource.ToString(),
            plan.PositioningCharacterSpacing,
            plan.CompensatePdfCharacterSpacing,
            DocxTextEmissionPlanner.ClassifyText(segment.Text),
            DocxTextEmissionPlanner.MeasureAdvanceProfile(segment.Text, segment.Resource?.Embedded, segment.FallbackFace?.Font, segment.Width, plan),
            DocxTextEmissionPlanner.CreateGlyphAdvanceSignature(segment.Text, segment.Resource?.Embedded, segment.FallbackFace?.Font),
            segment.IsTerminalLineSpace,
            segment.Resource?.Name ?? segment.FallbackFace?.ResourceName ?? string.Empty,
            segment.SyntheticBold,
            segment.SyntheticItalic,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleId,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleFound,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleDepth,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasDocumentDefaultRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasParagraphStyleRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasCharacterStyleRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasDirectRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasTableStyleRunProperties,
            segment.StyleRun.Revision?.Kind.ToValueString(),
            segment.StyleRun.Revision?.SourceElement);
    }
}
