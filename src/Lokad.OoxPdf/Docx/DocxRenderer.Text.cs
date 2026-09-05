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
        double fontScale,
        double baselineOffsetY,
        double xOffset,
        bool suppressCommentReferenceSpacer,
        bool useWordCompatibleTextProfile)
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
        void AddTerminalLineSpace()
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
}
