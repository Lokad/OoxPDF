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
    private static void RenderTextEmissionSegment(
        DocxTextEmissionSegment segment,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        DocxTextRun style = segment.StyleRun;
        RgbColor color = segment.Color;
        if (!segment.IsTerminalLineSpace)
        {
            RenderRunBackground(style, segment.Resource.Embedded.Font, segment.X, segment.Width, segment.FontSize, segment.BaselineY, graphics);
        }

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            style,
            segment.FontSize,
            segment.PdfCharacterSpacing,
            segment.PdfCharacterSpacingSource,
            segment.CompensatePdfCharacterSpacing,
            segment.IsTerminalLineSpace);
        DrawRunGlyphText(graphics, segment.Resource, segment.Text, segment.X, segment.BaselineY, color, plan, segment.SyntheticItalic);
        if (!segment.IsTerminalLineSpace && segment.SyntheticBold)
        {
            DrawRunGlyphText(graphics, segment.Resource, segment.Text, segment.X + 0.35d, segment.BaselineY, color, plan, segment.SyntheticItalic);
        }

        if (!segment.IsTerminalLineSpace)
        {
            bool useWordCompatibleRevisionDecorationProfile = ShouldUseWordCompatibleRevisionDecorationProfile(segment, markupContext);
            RgbColor decorationColor = useWordCompatibleRevisionDecorationProfile
                ? new RgbColor(WordCompatibleAllMarkupReviewStrokeRgb.Red, WordCompatibleAllMarkupReviewStrokeRgb.Green, WordCompatibleAllMarkupReviewStrokeRgb.Blue)
                : color;
            double decorationWidth = ResolveRevisionDecorationWidth(segment, useWordCompatibleRevisionDecorationProfile);
            RenderTextDecorations(
                style,
                segment.Resource.Embedded,
                segment.Text,
                segment.X,
                decorationWidth,
                segment.FontSize,
                segment.BaselineY,
                decorationColor,
                plan,
                useWordCompatibleRevisionDecorationProfile,
                graphics);
        }
    }

    private static bool ShouldUseWordCompatibleRevisionDecorationProfile(
        DocxTextEmissionSegment segment,
        DocxMarkupContext markupContext)
    {
        return UsesWordCompatibleAllMarkupTextProfile(markupContext) &&
            (segment.StyleRun.Revision is not null || segment.StyleRun.Revisions.Count != 0);
    }

    private static double ResolveRevisionDecorationWidth(
        DocxTextEmissionSegment segment,
        bool useWordCompatibleRevisionDecorationProfile)
    {
        if (!useWordCompatibleRevisionDecorationProfile ||
            !HasInsertionLikeRevisionKind(segment.StyleRun) ||
            segment.Text.EnumerateRunes().Count() > 8)
        {
            return segment.Width;
        }

        return Math.Max(0d, segment.Width - WordCompatibleAllMarkupShortInsertionDecorationWidthInsetPoints);
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

    private static IReadOnlyList<DocxTextEmissionSegment> CreateTextEmissionSegments(
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        int pageNumber,
        int pageCount,
        double fontScale = 1d,
        double baselineOffsetY = 0d,
        double xOffset = 0d,
        bool suppressCommentReferenceSpacer = false,
        bool useWordCompatibleTextProfile = false)
    {
        IReadOnlyList<DocxTextSegmentLayout> segments = line.Segments.Count == 0
            ? [new DocxTextSegmentLayout(line.Text, line.StyleRun, line.X, line.Width, null, 0d, 0d, DocxTextStateCharacterSpacingSource.None, true, -1, 0, DocxTextSegmentRole.Text)]
            : line.Segments;
        var emissionSegments = new List<DocxTextEmissionSegment>(segments.Count + 1);
        double substitutedFieldXAdjustment = 0d;
        foreach (DocxTextSegmentLayout segment in segments)
        {
            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            if (resource is null)
            {
                continue;
            }

            double fontSize = ResolveTextEmissionSegmentFontSize(segment, line, fontScale, useWordCompatibleTextProfile);
            double baselineY = GetSegmentBaselineY(segment, line.BaselineY) - baselineOffsetY;
            int partSourceTextOffset = segment.SourceTextOffsetInRun;
            foreach (DocxTextEmissionPart part in DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, fontSize, fontResources.TextMeasurer))
            {
                int currentPartSourceTextOffset = partSourceTextOffset;
                partSourceTextOffset += part.Text.Length;
                if (ShouldSuppressCommentReferenceSpacerPart(line, part, suppressCommentReferenceSpacer))
                {
                    continue;
                }

                DocxTextRun emissionStyleRun = ResolveWordCompatibleAllMarkupEmissionStyleRun(
                    segment.StyleRun,
                    segment.Role,
                    part.Text,
                    fontSize,
                    useWordCompatibleTextProfile);
                DocxEffectiveRunProperties emissionEffective = emissionStyleRun.EffectiveProperties;
                RgbColor textColor = ResolveTextEmissionColor(
                    emissionStyleRun,
                    useWordCompatibleTextProfile,
                    ReadColor(emissionEffective.ColorHex));
                string emittedText = ResolveStaticFieldPlaceholders(part.Text, pageNumber, pageCount);
                double emittedWidth = ResolveSubstitutedFieldEmissionWidth(part.Text, emittedText, emissionStyleRun, fontSize, fontResources.TextMeasurer, part.Width);
                emissionSegments.Add(new DocxTextEmissionSegment(
                    emittedText,
                    emissionStyleRun,
                    resource,
                    textColor,
                    part.X + xOffset + substitutedFieldXAdjustment + ResolveWordCompatibleAllMarkupBodyXOffset(segment, part, emissionStyleRun, fontSize, line.X, useWordCompatibleTextProfile),
                    baselineY,
                    emittedWidth,
                    fontSize,
                    segment.PdfCharacterSpacing,
                    segment.PdfCharacterSpacingSource,
                    segment.CompensatePdfCharacterSpacing,
                    ShouldApplySyntheticBold(emissionStyleRun, resource),
                    emissionEffective.Italic && !resource.Resolution.Italic,
                    IsTerminalLineSpace: false,
                    segment.SourceTextRunIndex,
                    currentPartSourceTextOffset,
                    segment.Role));
                substitutedFieldXAdjustment += emittedWidth - part.Width;
            }
        }

        if (line.EndsWithIntraTokenBreak)
        {
            return emissionSegments;
        }

        if (line.EmitsTerminalParagraphMark)
        {
            AddTerminalLineSpace(emissionSegments, segments, line, fontResources, fontScale, baselineOffsetY, xOffset, useWordCompatibleTextProfile);
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

            double fontSize = ResolveTerminalLineSpaceFontSize(segment, line, fontScale, useWordCompatibleTextProfile);
            double baselineY = GetSegmentBaselineY(segment, line.BaselineY) - baselineOffsetY;
            DocxEffectiveRunProperties effective = segment.StyleRun.EffectiveProperties;
            RgbColor color = ResolveTextEmissionColor(
                segment.StyleRun,
                useWordCompatibleTextProfile,
                ReadColor(effective.ColorHex));
            emissionSegments.Add(new DocxTextEmissionSegment(
                " ",
                segment.StyleRun,
                resource,
                color,
                ResolveTerminalLineSpaceX(segments, line, fontResources, fontScale, xOffset, useWordCompatibleTextProfile),
                baselineY,
                0d,
                fontSize,
                PdfCharacterSpacing: 0d,
                PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.TerminalLineSpace,
                CompensatePdfCharacterSpacing: true,
                SyntheticBold: false,
                SyntheticItalic: effective.Italic && !resource.Resolution.Italic,
                IsTerminalLineSpace: true,
                segment.SourceTextRunIndex,
                segment.SourceTextOffsetInRun + segment.Text.Length,
                segment.Role));
            break;
        }

        return emissionSegments;
    }

    private static RgbColor ResolveTextEmissionColor(
        DocxTextRun styleRun,
        bool useWordCompatibleTextProfile,
        RgbColor fallbackColor)
    {
        if (!useWordCompatibleTextProfile ||
            (styleRun.Revision is null && styleRun.Revisions.Count == 0))
        {
            return fallbackColor;
        }

        return new RgbColor(
            WordCompatibleAllMarkupReviewStrokeRgb.Red,
            WordCompatibleAllMarkupReviewStrokeRgb.Green,
            WordCompatibleAllMarkupReviewStrokeRgb.Blue);
    }

    private static DocxTextRun ResolveWordCompatibleAllMarkupEmissionStyleRun(
        DocxTextRun styleRun,
        DocxTextSegmentRole role,
        string text,
        double fontSize,
        bool useWordCompatibleTextProfile)
    {
        if (!ShouldApplyWordCompatibleAllMarkupBodyPositioningSpacing(role, text, useWordCompatibleTextProfile))
        {
            return styleRun;
        }

        return styleRun with
        {
            CharacterSpacingPoints = styleRun.CharacterSpacingPoints +
                ResolveWordCompatibleAllMarkupBodyPositioningCharacterSpacing(styleRun, text, fontSize)
        };
    }

    private static double ResolveWordCompatibleAllMarkupBodyPositioningCharacterSpacing(
        DocxTextRun styleRun,
        string text,
        double fontSize)
    {
        int runeCount = text.EnumerateRunes().Count();
        if (runeCount <= 8)
        {
            return WordCompatibleAllMarkupShortWordPositioningCharacterSpacingPoints;
        }

        if (HasInsertionLikeRevisionKind(styleRun))
        {
            return WordCompatibleAllMarkupInsertionPositioningCharacterSpacingPoints;
        }

        if (HasDeletionLikeRevisionKind(styleRun))
        {
            return WordCompatibleAllMarkupDeletionPositioningCharacterSpacingPoints;
        }

        if (fontSize >= WordCompatibleAllMarkupMaxBodyTextFontSizePoints - 0.001d)
        {
            return WordCompatibleAllMarkupHeadingPositioningCharacterSpacingPoints;
        }

        if (text.Any(char.IsPunctuation))
        {
            return WordCompatibleAllMarkupPunctuationPositioningCharacterSpacingPoints;
        }

        return WordCompatibleAllMarkupBodyPositioningCharacterSpacingPoints;
    }

    private static bool HasRevisionKind(DocxTextRun styleRun, DocxRevisionKind kind)
    {
        return IsRevisionKind(styleRun.Revision, kind) ||
            styleRun.Revisions.Any(revision => IsRevisionKind(revision, kind));
    }

    private static bool HasInsertionLikeRevisionKind(DocxTextRun styleRun)
    {
        return HasRevisionKind(styleRun, DocxRevisionKind.Insertion) ||
            HasRevisionKind(styleRun, DocxRevisionKind.MoveTo);
    }

    private static bool HasDeletionLikeRevisionKind(DocxTextRun styleRun)
    {
        return HasRevisionKind(styleRun, DocxRevisionKind.Deletion) ||
            HasRevisionKind(styleRun, DocxRevisionKind.MoveFrom);
    }

    private static bool IsRevisionKind(DocxRevisionInfo? revision, DocxRevisionKind kind)
    {
        return revision is not null &&
            revision.Kind == kind;
    }

    private static bool ShouldApplyWordCompatibleAllMarkupBodyPositioningSpacing(
        DocxTextSegmentRole role,
        string text,
        bool useWordCompatibleTextProfile)
    {
        return useWordCompatibleTextProfile &&
            role == DocxTextSegmentRole.Text &&
            text.EnumerateRunes().Take(2).Count() > 1 &&
            text.Any(character => !char.IsWhiteSpace(character));
    }

    private static double ResolveWordCompatibleAllMarkupBodyXOffset(
        DocxTextSegmentLayout sourceSegment,
        DocxTextEmissionPart part,
        DocxTextRun emissionStyleRun,
        double fontSize,
        double lineX,
        bool useWordCompatibleTextProfile)
    {
        if (!ShouldApplyWordCompatibleAllMarkupBodyPositioningSpacing(sourceSegment.Role, part.Text, useWordCompatibleTextProfile) ||
            fontSize >= WordCompatibleAllMarkupMaxBodyTextFontSizePoints - 0.001d)
        {
            return 0d;
        }

        double lineRelativeX = Math.Max(0d, part.X - lineX);
        if (lineRelativeX < 0.001d)
        {
            return 0d;
        }

        if (HasDeletionLikeRevisionKind(emissionStyleRun))
        {
            return WordCompatibleAllMarkupDeletionXOffsetPoints;
        }

        if (HasInsertionLikeRevisionKind(emissionStyleRun))
        {
            return WordCompatibleAllMarkupInsertionXOffsetPoints;
        }

        return WordCompatibleAllMarkupBodyXOffsetAsymptotePoints *
            (1d - Math.Exp(-lineRelativeX / WordCompatibleAllMarkupBodyXOffsetDecayPoints));
    }

    private static double ResolveTextEmissionSegmentFontSize(
        DocxTextSegmentLayout segment,
        DocxTextLineLayout line,
        double fontScale,
        bool useWordCompatibleTextProfile)
    {
        double fontSize = GetSegmentFontSize(segment, line.FontSize) * fontScale;
        return ShouldCapWordCompatibleAllMarkupTextFontSize(fontSize, useWordCompatibleTextProfile)
            ? WordCompatibleAllMarkupMaxBodyTextFontSizePoints
            : fontSize;
    }

    private static double ResolveTerminalLineSpaceFontSize(
        DocxTextSegmentLayout segment,
        DocxTextLineLayout line,
        double fontScale,
        bool useWordCompatibleTextProfile)
    {
        double fontSize = GetSegmentFontSize(segment, line.FontSize) * fontScale;
        return ShouldCapWordCompatibleAllMarkupTextFontSize(fontSize, useWordCompatibleTextProfile)
            ? WordCompatibleAllMarkupTerminalLineSpaceFontSizePoints
            : fontSize;
    }

    private static double ResolveTerminalLineSpaceX(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        double fontScale,
        double xOffset,
        bool useWordCompatibleTextProfile)
    {
        if (segments.Count == 0)
        {
            return line.X + xOffset;
        }

        DocxTextSegmentLayout terminalSegment = segments[^1];
        if (!ShouldUseWordCompatibleAllMarkupEmittedTerminalAdvance(segments, line, fontScale, useWordCompatibleTextProfile))
        {
            return terminalSegment.X + terminalSegment.Width + xOffset;
        }

        double? emittedEndX = TryResolveWordCompatibleAllMarkupTextEmissionEndX(
            segments,
            line,
            fontResources,
            fontScale,
            xOffset,
            useWordCompatibleTextProfile);
        return emittedEndX ?? terminalSegment.X + terminalSegment.Width + xOffset;
    }

    private static double? TryResolveWordCompatibleAllMarkupTextEmissionEndX(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        double fontScale,
        double xOffset,
        bool useWordCompatibleTextProfile)
    {
        double emittedEndX = double.NegativeInfinity;
        foreach (DocxTextSegmentLayout segment in segments)
        {
            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            if (resource is null)
            {
                continue;
            }

            double fontSize = ResolveTextEmissionSegmentFontSize(segment, line, fontScale, useWordCompatibleTextProfile);
            foreach (DocxTextEmissionPart part in DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, fontSize, fontResources.TextMeasurer))
            {
                DocxTextRun emissionStyleRun = ResolveWordCompatibleAllMarkupEmissionStyleRun(
                    segment.StyleRun,
                    segment.Role,
                    part.Text,
                    fontSize,
                    useWordCompatibleTextProfile);
                DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(
                    emissionStyleRun,
                    fontSize,
                    segment.PdfCharacterSpacing,
                    segment.PdfCharacterSpacingSource,
                    segment.CompensatePdfCharacterSpacing,
                    isTerminalLineSpace: false);
                double partX = part.X + xOffset + ResolveWordCompatibleAllMarkupBodyXOffset(
                    segment,
                    part,
                    emissionStyleRun,
                    fontSize,
                    line.X,
                    useWordCompatibleTextProfile);
                double partAdvance = DocxTextEmissionPlanner.MeasureAdvanceProfile(part.Text, resource.Embedded, part.Width, plan)
                    .PlannedEmittedAdvance;
                emittedEndX = Math.Max(emittedEndX, partX + partAdvance);
            }
        }

        return double.IsNegativeInfinity(emittedEndX) ? null : emittedEndX;
    }

    private static bool ShouldUseWordCompatibleAllMarkupEmittedTerminalAdvance(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        double fontScale,
        bool useWordCompatibleTextProfile)
    {
        return useWordCompatibleTextProfile &&
            segments.Any(segment => ShouldCapWordCompatibleAllMarkupTextFontSize(GetSegmentFontSize(segment, line.FontSize) * fontScale, useWordCompatibleTextProfile));
    }

    private static bool ShouldCapWordCompatibleAllMarkupTextFontSize(double fontSize, bool useWordCompatibleTextProfile)
    {
        return useWordCompatibleTextProfile && fontSize > WordCompatibleAllMarkupMaxBodyTextFontSizePoints;
    }

    private static bool ShouldSuppressCommentReferenceSpacerPart(
        DocxTextLineLayout line,
        DocxTextEmissionPart part,
        bool suppressCommentReferenceSpacer)
    {
        return suppressCommentReferenceSpacer &&
            part.Width > 0d &&
            !string.IsNullOrEmpty(part.Text) &&
            string.IsNullOrWhiteSpace(part.Text) &&
            line.SourceParagraph?.InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment) == true;
    }

    private static void AddTerminalLineSpace(
        List<DocxTextEmissionSegment> emissionSegments,
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        double fontScale = 1d,
        double baselineOffsetY = 0d,
        double xOffset = 0d,
        bool useWordCompatibleTextProfile = false)
    {
        if (segments.Count == 0)
        {
            return;
        }

        DocxTextSegmentLayout segment = segments[^1];
        DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
        if (resource is null)
        {
            return;
        }

        double fontSize = ResolveTerminalLineSpaceFontSize(segment, line, fontScale, useWordCompatibleTextProfile);
        double baselineY = GetSegmentBaselineY(segment, line.BaselineY) - baselineOffsetY;
        DocxEffectiveRunProperties effective = segment.StyleRun.EffectiveProperties;
        emissionSegments.Add(new DocxTextEmissionSegment(
            " ",
            segment.StyleRun,
            resource,
            ReadColor(effective.ColorHex),
            ResolveTerminalLineSpaceX(segments, line, fontResources, fontScale, xOffset, useWordCompatibleTextProfile),
            baselineY,
            0d,
            fontSize,
            PdfCharacterSpacing: 0d,
            PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.TerminalLineSpace,
            CompensatePdfCharacterSpacing: true,
            SyntheticBold: false,
            SyntheticItalic: effective.Italic && !resource.Resolution.Italic,
            IsTerminalLineSpace: true,
            segment.SourceTextRunIndex,
            segment.SourceTextOffsetInRun + segment.Text.Length,
            segment.Role));
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

    private static string ResolveTextEmissionStoryKind(DocxTextLineLayout line, string fallback)
    {
        return string.IsNullOrWhiteSpace(line.StoryKind) ? fallback : line.StoryKind;
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
