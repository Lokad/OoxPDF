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
        bool ShouldUseWordCompatibleRevisionDecorationProfile()
        {
            return UsesWordCompatibleAllMarkupTextProfile(markupContext) &&
                (segment.StyleRun.Revision is not null || segment.StyleRun.Revisions.Count != 0);
        }

        RgbColor color = segment.Color;
        if (segment.FallbackFace is { } fallback)
        {
            DrawFallbackTextEmissionSegment(segment, fallback, graphics, markupContext);
            return;
        }

        DocxRunFontResource? resource = segment.Resource;
        if (resource is null)
        {
            return;
        }

        if (!segment.IsTerminalLineSpace)
        {
            OpenTypeFont font = resource.Embedded.Font;
            RenderRunBackground(style, DocxLineMetrics.MeasureWindowsAscender(font, segment.FontSize), DocxLineMetrics.MeasureWindowsDescender(font, segment.FontSize), segment.X, segment.Width, segment.BaselineY, graphics);
        }

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            style,
            segment.FontSize,
            segment.PdfCharacterSpacing,
            segment.PdfCharacterSpacingSource,
            segment.CompensatePdfCharacterSpacing,
            segment.IsTerminalLineSpace);
        DrawRunGlyphText(graphics, resource, segment.Text, segment.X, segment.BaselineY, color, plan, segment.SyntheticItalic);
        if (!segment.IsTerminalLineSpace && segment.SyntheticBold)
        {
            DrawRunGlyphText(graphics, resource, segment.Text, segment.X + 0.35d, segment.BaselineY, color, plan, segment.SyntheticItalic);
        }

        if (!segment.IsTerminalLineSpace)
        {
            bool useWordCompatibleRevisionDecorationProfile = ShouldUseWordCompatibleRevisionDecorationProfile();
            RgbColor decorationColor = useWordCompatibleRevisionDecorationProfile
                ? new RgbColor(WordCompatibleAllMarkupReviewStrokeRgb.Red, WordCompatibleAllMarkupReviewStrokeRgb.Green, WordCompatibleAllMarkupReviewStrokeRgb.Blue)
                : color;
            double ResolveRevisionDecorationWidth()
            {
                if (!useWordCompatibleRevisionDecorationProfile ||
                    !HasInsertionLikeRevisionKind(segment.StyleRun) ||
                    segment.Text.EnumerateRunes().Count() > 8)
                {
                    return segment.Width;
                }
        
                return Math.Max(0d, segment.Width - WordCompatibleAllMarkupShortInsertionDecorationWidthInsetPoints);
            }

            double decorationWidth = ResolveRevisionDecorationWidth();
            RenderTextDecorations(
                style,
                resource.Embedded,
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
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex)
    {
        cancellationToken.ThrowIfCancellationRequested();
        PdfImageXObject? xObject = CreateImage(image.Image, imageCache, diagnosticSink, image.PageIndex, cancellationToken);
        if (xObject is null)
        {
            return;
        }

        string imageName = "Im" + imageIndex++;
        graphics.DrawImage(imageName, image.X, image.Y, image.Width, image.Height);
        pageImages.Add(new PdfImageResource(imageName, xObject));
    }

    private static void AddEmissionSegmentWithFontFallback(
        List<DocxTextEmissionSegment> emissionSegments,
        DocxTextEmissionSegment segment,
        DocxFontResources fontResources,
        CancellationToken cancellationToken)
    {
        if (segment.IsTerminalLineSpace ||
            string.IsNullOrEmpty(segment.Text) ||
            !fontResources.FallbackChains.TryGetValue(segment.StyleRun, out IReadOnlyList<DocxFallbackFontEntry>? chain) ||
            chain.Count <= 1)
        {
            emissionSegments.Add(segment);
            return;
        }

        var fonts = new OpenTypeFont?[chain.Count];
        for (int i = 0; i < chain.Count; i++)
        {
            fonts[i] = chain[i].Font;
        }

        IReadOnlyList<FontCoverageSpan> spans = FontCoverageFallback.SplitByCoverage(segment.Text, fonts, cancellationToken);
        bool needsFallback = false;
        foreach (FontCoverageSpan span in spans)
        {
            if (span.FontIndex > 0)
            {
                needsFallback = true;
                break;
            }
        }

        if (!needsFallback)
        {
            emissionSegments.Add(segment);
            return;
        }

        IDocxTextMeasurer? measurer = fontResources.TextMeasurer;
        if (measurer is null)
        {
            emissionSegments.Add(segment);
            return;
        }

        var widths = new double[spans.Count];
        double total = 0d;
        for (int i = 0; i < spans.Count; i++)
        {
            widths[i] = measurer.MeasureText(segment.StyleRun, segment.Text.Substring(spans[i].Start, spans[i].Length), segment.FontSize);
            total += widths[i];
        }

        if (total <= 0d)
        {
            emissionSegments.Add(segment);
            return;
        }

        bool italic = segment.StyleRun.EffectiveProperties.Italic;
        double scale = segment.Width / total;
        double x = segment.X;
        for (int i = 0; i < spans.Count; i++)
        {
            FontCoverageSpan span = spans[i];
            DocxFallbackFontEntry entry = span.FontIndex < 0 ? chain[0] : chain[span.FontIndex];
            double spanWidth = i == spans.Count - 1 ? segment.X + segment.Width - x : widths[i] * scale;
            emissionSegments.Add(segment with
            {
                Text = segment.Text.Substring(span.Start, span.Length),
                Resource = entry.Resource,
                X = x,
                Width = spanWidth,
                SyntheticBold = ShouldApplySyntheticBold(segment.StyleRun, entry.Resource),
                SyntheticItalic = italic && !entry.Resource.Resolution.Italic,
                SourceTextOffsetInRun = segment.SourceTextOffsetInRun + span.Start
            });
            x += spanWidth;
        }
    }

    private static IReadOnlyList<DocxTextEmissionSegment> CreateTextEmissionSegments(
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        int pageNumber,
        int pageCount,
        double fontScale,
        double baselineOffsetY,
        double xOffset,
        bool suppressCommentReferenceSpacer,
        bool useWordCompatibleTextProfile,
        CancellationToken cancellationToken = default)
    {
        IReadOnlyList<DocxTextSegmentLayout> segments = line.Segments.Count == 0
            ? [new DocxTextSegmentLayout(line.Text, line.StyleRun, line.X, line.Width, null, 0d, 0d, DocxTextStateCharacterSpacingSource.None, true, -1, 0, DocxTextSegmentRole.Text)]
            : line.Segments;
        var emissionSegments = new List<DocxTextEmissionSegment>(segments.Count + 1);
        double substitutedFieldXAdjustment = 0d;
        foreach (DocxTextSegmentLayout segment in segments)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            PdfFallbackFontResource? fallbackFace = null;
            if (resource is null)
            {
                fontResources.FallbackFaces.TryGetValue(segment.StyleRun, out fallbackFace);
            }

            if (resource is null && fallbackFace is null)
            {
                continue;
            }

            double fontSize = ResolveTextEmissionSegmentFontSize(segment, line, fontScale);
            double baselineY = GetSegmentBaselineY(segment, line.BaselineY) - baselineOffsetY;
            int partSourceTextOffset = segment.SourceTextOffsetInRun;
            bool ShouldSuppressCommentReferenceSpacerPart(DocxTextEmissionPart part)
            {
            return suppressCommentReferenceSpacer &&
                part.Width > 0d &&
                !string.IsNullOrEmpty(part.Text) &&
                string.IsNullOrWhiteSpace(part.Text) &&
                line.SourceParagraph?.InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment) == true;
            }

            foreach (DocxTextEmissionPart part in DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, fontSize, fontResources.TextMeasurer))
            {
                int currentPartSourceTextOffset = partSourceTextOffset;
                partSourceTextOffset += part.Text.Length;
                if (ShouldSuppressCommentReferenceSpacerPart(part))
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
                string emittedText = ResolveStaticFieldPlaceholders(emissionStyleRun, part.Text, pageNumber, pageCount);
                double emittedWidth = ResolveSubstitutedFieldEmissionWidth(part.Text, emittedText, emissionStyleRun, fontSize, fontResources.TextMeasurer, part.Width);
                AddEmissionSegmentWithFontFallback(emissionSegments, new DocxTextEmissionSegment(
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
                    resource is not null && ShouldApplySyntheticBold(emissionStyleRun, resource),
                    resource is not null && emissionEffective.Italic && !resource.Resolution.Italic,
                    IsTerminalLineSpace: false,
                    segment.SourceTextRunIndex,
                    currentPartSourceTextOffset,
                    segment.Role,
                    FallbackFace: fallbackFace), fontResources, cancellationToken);
                substitutedFieldXAdjustment += emittedWidth - part.Width;
            }
        }

        if (line.EndsWithIntraTokenBreak)
        {
            return emissionSegments;
        }

        if (line.EmitsTerminalParagraphMark)
        {
        void AddTerminalLineSpace()
        {
            if (segments.Count == 0)
            {
                return;
            }
    
            DocxTextSegmentLayout segment = segments[^1];
            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            fontResources.FallbackFaces.TryGetValue(segment.StyleRun, out PdfFallbackFontResource? terminalFallbackFace);
            if (resource is null && terminalFallbackFace is null)
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
                ResolveTerminalLineSpaceX(segments, line, fontResources, fontScale, xOffset, useWordCompatibleTextProfile, cancellationToken),
                baselineY,
                0d,
                fontSize,
                PdfCharacterSpacing: 0d,
                PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.TerminalLineSpace,
                CompensatePdfCharacterSpacing: true,
                SyntheticBold: false,
                SyntheticItalic: resource is not null && effective.Italic && !resource.Resolution.Italic,
                IsTerminalLineSpace: true,
                segment.SourceTextRunIndex,
                segment.SourceTextOffsetInRun + segment.Text.Length,
                segment.Role,
                FallbackFace: terminalFallbackFace));
        }

            AddTerminalLineSpace();
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
                ResolveTerminalLineSpaceX(segments, line, fontResources, fontScale, xOffset, useWordCompatibleTextProfile, cancellationToken),
                baselineY,
                0d,
                fontSize,
                PdfCharacterSpacing: 0d,
                PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.TerminalLineSpace,
                CompensatePdfCharacterSpacing: true,
                SyntheticBold: false,
                SyntheticItalic: resource is not null && effective.Italic && !resource.Resolution.Italic,
                IsTerminalLineSpace: true,
                segment.SourceTextRunIndex,
                segment.SourceTextOffsetInRun + segment.Text.Length,
                segment.Role,
                FallbackFace: null));
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
        double fontScale)
    {
        // Office A/B (W5-C1 size probe w5-cap1422): a 22pt run prints at full lane-fit scale
        // (16.656), not the 15pt-design cap. Large body text scales uniformly; the cap only
        // ever bound above 15pt design and is removed (terminal spacing keeps its own rule).
        return GetSegmentFontSize(segment, line.FontSize) * fontScale;
    }

    private static double ResolveTerminalLineSpaceFontSize(
        DocxTextSegmentLayout segment,
        DocxTextLineLayout line,
        double fontScale,
        bool useWordCompatibleTextProfile)
    {
        double fontSize = GetSegmentFontSize(segment, line.FontSize) * fontScale;
        return ShouldCapWordCompatibleAllMarkupTextFontSize(fontSize, useWordCompatibleTextProfile)
            ? 11d * fontScale
            : fontSize;
    }

    private static double ResolveTerminalLineSpaceX(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        double fontScale,
        double xOffset,
        bool useWordCompatibleTextProfile,
        CancellationToken cancellationToken)
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
            useWordCompatibleTextProfile,
            cancellationToken);
        return emittedEndX ?? terminalSegment.X + terminalSegment.Width + xOffset;
    }

    private static double? TryResolveWordCompatibleAllMarkupTextEmissionEndX(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        double fontScale,
        double xOffset,
        bool useWordCompatibleTextProfile,
        CancellationToken cancellationToken)
    {
        double emittedEndX = double.NegativeInfinity;
        foreach (DocxTextSegmentLayout segment in segments)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            PdfFallbackFontResource? fallbackFace = null;
            if (resource is null)
            {
                fontResources.FallbackFaces.TryGetValue(segment.StyleRun, out fallbackFace);
            }

            if (resource is null && fallbackFace is null)
            {
                continue;
            }

            double fontSize = ResolveTextEmissionSegmentFontSize(segment, line, fontScale);
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
                // RV01: fallback runs have no advance profile; measured part width stands in.
                double partAdvance = resource is not null
                    ? DocxTextEmissionPlanner.MeasureAdvanceProfile(part.Text, resource.Embedded, part.Width, plan).PlannedEmittedAdvance
                    : part.Width;
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

    // RV01: diagnosed fallback emission: background from fallback constants, every
    // rune positioned absolutely with fallback advances, decorations as solid
    // constant-metric rectangles. Terminal spaces draw their space glyph like the
    // embedded path.
    private static void DrawFallbackTextEmissionSegment(
        DocxTextEmissionSegment segment,
        PdfFallbackFontResource fallback,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        if (!segment.IsTerminalLineSpace)
        {
            RenderRunBackground(
                segment.StyleRun,
                segment.FontSize * PdfFallbackFont.AscentEm,
                segment.FontSize * PdfFallbackFont.DescentEm,
                segment.X,
                segment.Width,
                segment.BaselineY,
                graphics);
        }

        DrawFallbackRunGlyphText(
            graphics,
            fallback,
            segment.Text,
            segment.X,
            segment.BaselineY,
            segment.Color,
            segment.FontSize,
            segment.StyleRun.EffectiveProperties.CharacterSpacingPoints);

        if (!segment.IsTerminalLineSpace)
        {
            bool revisionProfile = UsesWordCompatibleAllMarkupTextProfile(markupContext) &&
                (segment.StyleRun.Revision is not null || segment.StyleRun.Revisions.Count != 0);
            RgbColor decorationColor = revisionProfile
                ? new RgbColor(WordCompatibleAllMarkupReviewStrokeRgb.Red, WordCompatibleAllMarkupReviewStrokeRgb.Green, WordCompatibleAllMarkupReviewStrokeRgb.Blue)
                : segment.Color;
            RenderFallbackTextDecorations(
                segment.StyleRun,
                segment.X,
                segment.Width,
                segment.FontSize,
                segment.BaselineY,
                decorationColor,
                graphics);
        }
    }

    private static void RenderTextLine(
        DocxTextLineLayout line,
        PdfGraphicsBuilder graphics,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        CancellationToken cancellationToken)
    {
        RenderMarkupIndicators(line, graphics, fontResources, markupContext);
        foreach (DocxTextEmissionSegment segment in CreateTextEmissionSegments(
            line,
            fontResources,
            pageNumber,
            pageCount,
            ResolveTextEmissionFontScale(markupContext),
            ResolveTextEmissionBaselineOffset(markupContext),
            ResolveTextEmissionXOffset(markupContext),
            ShouldSuppressWordCompatibleCommentReferenceSpacer(markupContext),
            UsesWordCompatibleAllMarkupTextProfile(markupContext), cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            RenderTextEmissionSegment(segment, graphics, markupContext);
        }
    }

    private static void RenderMarkupIndicators(
        DocxTextLineLayout line,
        PdfGraphicsBuilder graphics,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext)
    {
        if (!markupContext.DrawsChangeBars && !markupContext.DrawsCommentMarkers)
        {
            return;
        }

        IReadOnlyList<DocxRevisionInfo> lineRevisions = markupContext.DrawsChangeBars && !UsesWordCompatibleAllMarkupTextProfile(markupContext)
            ? CollectTextLineRevisions(line)
            : [];
        if (markupContext.DrawsChangeBars &&
            lineRevisions.Count != 0 &&
            !UsesWordCompatibleAllMarkupTextProfile(markupContext))
        {
            double scaledFontSize = line.FontSize * ResolveTextEmissionFontScale(markupContext);
            double height = Math.Max(6d, line.LineHeight ?? scaledFontSize * 1.2d);
            double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
            double y = baselineY - height * 0.25d;
            DocxMarkupBalloonRgb color = ResolveRevisionAuthorColor(lineRevisions);
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.FillRectangle(Math.Max(0d, line.X - 7d), y, 1.5d, height);
        }

        if (line.SourceParagraph is not { } paragraph)
        {
            return;
        }

        if (markupContext.DrawsCommentMarkers && paragraph.InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment))
        {
            if (UsesWordCompatibleAllMarkupTextProfile(markupContext))
            {
                RenderWordCompatibleCommentRangeMarkers(line, paragraph, graphics, markupContext);
                return;
            }

            double scaledFontSize = line.FontSize * ResolveTextEmissionFontScale(markupContext);
            string label = ResolveCommentMarkerLabel(paragraph);
            double labelFontSize = Math.Max(4.5d, Math.Min(7d, scaledFontSize * 0.55d));
            double markerHeight = Math.Max(6d, labelFontSize + 2d);
            double markerWidth = Math.Max(markerHeight, label.Length * labelFontSize * 0.55d + 3d);
            double markerX = line.X + Math.Max(0d, line.Width) + 2d;
            double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
            double markerY = baselineY + markerHeight * 0.15d;
            graphics.SetFillRgb(255, 192, 0);
            graphics.FillRectangle(markerX, markerY, markerWidth, markerHeight);
            graphics.SetStrokeRgb(217, 151, 0);
            graphics.SetLineWidth(0.5d);
            graphics.StrokeRectangle(markerX, markerY, markerWidth, markerHeight);
            if (!ShouldDrawCommentMarkerLabel(markupContext))
            {
                return;
            }

            DocxRunFontResource? labelResource = fontResources.Fallback ??
                line.Segments
                    .Select(segment => ResolveFontResource(segment.StyleRun, fontResources))
                    .FirstOrDefault(resource => resource is not null);
            if (labelResource is not null)
            {
                string glyphHex = labelResource.Embedded.EncodeGlyphHex(SubstituteUncoveredGlyphs(labelResource.Embedded, label));
                if (glyphHex.Length != 0)
                {
                    graphics.DrawGlyphText(labelResource.Name, labelFontSize, markerX + 1.5d, markerY + 1.4d, 0, 0, 0, glyphHex, italic: false, characterSpacing: 0d, textRenderingMode: 0, strokeRed: 0, strokeGreen: 0, strokeBlue: 0, strokeWidth: 0d);
                }
            }
            else
            {
                // RV01: marker labels without any embedded resource use the first
                // fallback face of the line instead of vanishing.
                PdfFallbackFontResource? fallbackLabel = line.Segments
                    .Select(segment => fontResources.FallbackFaces.TryGetValue(segment.StyleRun, out PdfFallbackFontResource? face) ? face : null)
                    .FirstOrDefault(face => face is not null);
                if (fallbackLabel is not null)
                {
                    DrawFallbackRunGlyphText(graphics, fallbackLabel, label, markerX + 1.5d, markerY + 1.4d, new RgbColor(0, 0, 0), labelFontSize, 0d);
                }
            }
        }
    }
}
