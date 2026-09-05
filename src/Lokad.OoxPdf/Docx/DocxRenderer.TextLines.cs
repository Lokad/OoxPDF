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
        DocxLayout layout,
        DocxLayoutPage page,
        int pageIndex)
    {
        return EnumerateStaticTextLines(page)
            .Concat(EnumerateBodyTextLines(page))
            .Concat(EnumeratePlacedRelatedStoryTextLines(page))
            .Concat(EnumerateFloatingDrawingTextBoxTextLines(EnumeratePageFloatingDrawings(layout, pageIndex)))
            .Concat(page.PlacedRelatedStories.SelectMany(story => EnumerateFloatingDrawingTextBoxTextLines(story.FloatingDrawings)));
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateMarkupBalloonAnchorTextLines(
        DocxLayoutPage page,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings)
    {
        return EnumerateStaticTextLines(page)
            .Concat(EnumerateBodyTextLines(page))
            .Concat(EnumeratePlacedRelatedStoryTextLines(page))
            .Concat(EnumerateFloatingDrawingTextBoxTextLines(floatingDrawings))
            .Concat(page.PlacedRelatedStories.SelectMany(story => EnumerateFloatingDrawingTextBoxTextLines(story.FloatingDrawings)));
    }

    private sealed record DocxTextEmissionLineSource(
        DocxTextLineLayout Line,
        bool IsStaticStory,
        string StoryKind,
        string? StoryVariantType,
        string? ContainerStoryKind,
        string? ContainerStoryVariantType);

    private static IEnumerable<DocxTextEmissionLineSource> EnumerateRenderedFloatingDrawingTextBoxTextLines(
        DocxLayout layout,
        DocxLayoutPage page,
        int pageIndex)
    {
        foreach (DocxFloatingDrawingLayout drawing in EnumeratePageFloatingDrawings(layout, pageIndex))
        {
            bool isStaticStory = string.Equals(drawing.StoryKind, "Header", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(drawing.StoryKind, "Footer", StringComparison.OrdinalIgnoreCase);
            foreach (DocxTextLineLayout line in EnumerateFloatingDrawingTextBoxTextLines(drawing))
            {
                yield return new DocxTextEmissionLineSource(
                    line,
                    isStaticStory,
                    "TextBox",
                    line.StoryVariantType,
                    drawing.StoryKind ?? "Body",
                    drawing.StoryVariantType);
            }
        }

        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTextLineLayout line in EnumerateFloatingDrawingTextBoxTextLines(story.FloatingDrawings))
            {
                yield return new DocxTextEmissionLineSource(
                    line,
                    IsStaticStory: false,
                    "TextBox",
                    line.StoryVariantType,
                    story.StoryLayout.Story.Kind.ToValueString(),
                    story.StoryLayout.Story.Id);
            }
        }
    }

    private static IEnumerable<DocxLayoutItem> EnumerateStaticLayoutItems(DocxLayoutPage page)
    {
        return page.StaticTextLines
            .Cast<DocxLayoutItem>()
            .Concat(page.StaticInlineImages)
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
        foreach (DocxTextLineLayout line in page.StaticTextLines)
        {
            yield return line;
        }

        foreach (DocxTableRowLayout row in page.StaticTableRows)
        {
            foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
            {
                yield return cellLine;
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

    private static IEnumerable<DocxTextLineLayout> EnumerateFloatingDrawingTextBoxTextLines(IEnumerable<DocxFloatingDrawingLayout> drawings)
    {
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            foreach (DocxTextLineLayout line in EnumerateFloatingDrawingTextBoxTextLines(drawing))
            {
                yield return line;
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
            DocxTextEmissionPlanner.MeasureAdvanceProfile(segment.Text, segment.Resource.Embedded, segment.Width, plan),
            DocxTextEmissionPlanner.CreateGlyphAdvanceSignature(segment.Text, segment.Resource.Embedded),
            segment.IsTerminalLineSpace,
            segment.Resource.Name,
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
