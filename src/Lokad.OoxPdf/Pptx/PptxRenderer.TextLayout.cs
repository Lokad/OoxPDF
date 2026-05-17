using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal static IReadOnlyList<PptxTextRunSnapshot> InspectTextRuns(PptxDocument document, OoxPackage package, int slideIndex)
    {
        return ReadSlideTextRunsForInspection(document, package, slideIndex)
            .Select(ToSnapshot)
            .ToArray();
    }

    private static IReadOnlyList<TextRun> ReadSlideTextRunsForInspection(PptxDocument document, OoxPackage package, int slideIndex)
    {
        return ReadSlideTextSpansForInspection(document, package, slideIndex).Select(span => span.Run).ToArray();
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadSlideTextSpansForInspection(PptxDocument document, OoxPackage package, int slideIndex)
    {
        PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);
        PptxRenderContext? context = TryLoadRenderContext(document, package, theme, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null);
        if (context is null)
        {
            return [];
        }

        return context.InheritedXml
            .SelectMany(xml => ReadTextSpans(context, xml, includePlaceholders: false, placeholderSources: []))
            .Concat(ReadTextSpans(context, context.SlideXml, includePlaceholders: true, context.InheritedXml))
            .ToArray();
    }

    internal static PptxTextLayoutSnapshot InspectTextLayout(PptxDocument document, OoxPackage package, int slideIndex)
    {
        PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);
        PptxRenderContext? context = TryLoadRenderContext(document, package, theme, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null);
        if (context is null)
        {
            return new PptxTextLayoutSnapshot([]);
        }

        PptxTextLayoutModel inheritedLayout = BuildTextLayoutModelForSources(context.InheritedXml, context);
        PptxTextLayoutModel slideLayout = BuildTextLayoutModel(context, context.SlideXml, includePlaceholders: true, context.InheritedXml);
        return ToSnapshot(new PptxTextLayoutModel(inheritedLayout.Frames.Concat(slideLayout.Frames).ToArray()));
    }

    internal static PptxTextFlowSnapshot InspectTextFlow(PptxDocument document, OoxPackage package, int slideIndex)
    {
        PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);
        PptxRenderContext? context = TryLoadRenderContext(document, package, theme, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null);
        if (context is null)
        {
            return new PptxTextFlowSnapshot([]);
        }

        PptxTextFlowModel inheritedFlow = BuildTextFlowModelForSources(context.InheritedXml, context);
        PptxTextFlowModel slideFlow = BuildTextFlowModel(context, context.SlideXml, includePlaceholders: true, context.InheritedXml);
        return ToSnapshot(new PptxTextFlowModel(inheritedFlow.Frames.Concat(slideFlow.Frames).ToArray()));
    }

    private static PptxTextLayoutModel BuildTextLayoutModelForSources(
        IReadOnlyList<XDocument> sources,
        PptxRenderContext context)
    {
        var frames = new List<PptxTextFrameLayout>();
        foreach (XDocument source in sources)
        {
            frames.AddRange(BuildTextLayoutModel(context, source, includePlaceholders: false, placeholderSources: []).Frames);
        }

        return new PptxTextLayoutModel(frames);
    }

    private static PptxTextFlowModel BuildTextFlowModelForSources(
        IReadOnlyList<XDocument> sources,
        PptxRenderContext context)
    {
        var frames = new List<PptxTextFlowFrame>();
        foreach (XDocument source in sources)
        {
            frames.AddRange(BuildTextFlowModel(context, source, includePlaceholders: false, placeholderSources: []).Frames);
        }

        return new PptxTextFlowModel(frames);
    }

    private static PptxTextRunSnapshot ToSnapshot(TextRun run)
    {
        return new PptxTextRunSnapshot(
            run.Text,
            run.X,
            run.Y,
            run.Width,
            run.FontSize,
            run.CharacterSpacing,
            run.Color,
            run.Alpha,
            run.HighlightColor,
            run.Bold,
            run.Italic,
            run.Underline,
            run.Strike,
            run.Alignment.ToString(),
            run.FontFamily);
    }

    private static PptxTextLayoutSnapshot ToSnapshot(PptxTextLayoutModel layout)
    {
        return new PptxTextLayoutSnapshot(layout.Frames.Select(ToSnapshot).ToArray());
    }

    private static PptxTextFlowSnapshot ToSnapshot(PptxTextFlowModel flow)
    {
        return new PptxTextFlowSnapshot(flow.Frames.Select(ToSnapshot).ToArray());
    }

    private static PptxTextFlowFrameSnapshot ToSnapshot(PptxTextFlowFrame frame)
    {
        return new PptxTextFlowFrameSnapshot(
            frame.Box.TextX,
            frame.Box.TextWidth,
            frame.Box.TextHeight,
            frame.Box.ClipY,
            frame.Box.ClipHeight,
            frame.Box.CursorTop,
            frame.Paragraphs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextFlowParagraphSnapshot ToSnapshot(PptxTextFlowParagraph paragraph)
    {
        return new PptxTextFlowParagraphSnapshot(
            paragraph.Model.Level,
            paragraph.Style.Alignment.ToString(),
            paragraph.Style.FontSize,
            paragraph.Runs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextFlowRunSnapshot ToSnapshot(PptxTextFlowRun run)
    {
        return new PptxTextFlowRunSnapshot(
            run.Source.Kind.ToString(),
            run.Source.Text,
            run.Style.FontSize,
            run.Style.Typeface,
            run.Segments.Select(ToSnapshot).ToArray());
    }

    private static PptxTextFlowSegmentSnapshot ToSnapshot(PptxTextFlowSegment segment)
    {
        return new PptxTextFlowSegmentSnapshot(
            segment.Kind.ToString(),
            segment.Text,
            segment.AdvanceText,
            segment.Draw,
            segment.PreventCoalesce,
            segment.FontScale,
            segment.AdvanceFontSizeFactor);
    }

    private static PptxTextFrameLayoutSnapshot ToSnapshot(PptxTextFrameLayout frame)
    {
        return new PptxTextFrameLayoutSnapshot(frame.Paragraphs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextParagraphLayoutSnapshot ToSnapshot(PptxTextParagraphLayout paragraph)
    {
        return new PptxTextParagraphLayoutSnapshot(paragraph.Model.Level, paragraph.Lines.Select(ToSnapshot).ToArray());
    }

    private static PptxTextLineLayoutSnapshot ToSnapshot(PptxTextLineLayout line)
    {
        return new PptxTextLineLayoutSnapshot(
            line.Box.TopY,
            line.Box.BaselineY,
            line.Box.Advance,
            line.Box.BaselineOffset,
            line.Box.MaxFontSize,
            line.Box.LineSpacing.IsAbsolute ? "Absolute" : line.Box.LineSpacing.IsExplicit ? "Multiple" : "Default",
            ToSnapshot(line.Box.BaselineMetric),
            line.StartX,
            line.EndX,
            line.Alignment.ToString(),
            line.Spans.Select(ToSnapshot).ToArray());
    }

    private static PptxTextBaselineMetricSnapshot ToSnapshot(PptxTextBaselineMetricLayout metric)
    {
        return new PptxTextBaselineMetricSnapshot(
            metric.Source,
            metric.Typeface,
            metric.Bold,
            metric.Italic,
            metric.FontSize,
            metric.Ratio,
            metric.UnitsPerEm,
            metric.WindowsAscender,
            metric.WindowsDescender,
            metric.TypographicAscender,
            metric.TypographicDescender,
            metric.TypographicLineGap);
    }

    private static PptxTextSpanLayoutSnapshot ToSnapshot(PptxTextSpanLayout span)
    {
        return new PptxTextSpanLayoutSnapshot(
            span.SourceRun?.Text,
            span.Run.Text,
            span.Run.X,
            span.Run.Y,
            span.Run.Width,
            span.Run.FontSize,
            span.Atoms.Select(ToSnapshot).ToArray(),
            ToSnapshot(span.GlyphSpan));
    }

    private static PptxTextAtomLayoutSnapshot ToSnapshot(PptxTextAtomLayout atom)
    {
        return new PptxTextAtomLayoutSnapshot(atom.Kind.ToString(), atom.Text, atom.X, atom.Width, atom.Draw);
    }

    private static PptxTextGlyphSpanLayoutSnapshot ToSnapshot(PptxTextGlyphSpanLayout span)
    {
        return new PptxTextGlyphSpanLayoutSnapshot(
            span.Text,
            span.Typeface,
            span.FontSize,
            span.NaturalWidth,
            span.LayoutWidth,
            span.Glyphs.Count,
            span.Glyphs.Skip(1).FirstOrDefault()?.AdjustmentBefore ?? 0d,
            span.Glyphs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextGlyphLayoutSnapshot ToSnapshot(PptxTextGlyphLayout glyph)
    {
        return new PptxTextGlyphLayoutSnapshot(glyph.CodePoint, glyph.GlyphId, glyph.Advance, glyph.AdjustmentBefore);
    }

    private static IReadOnlyList<TextRun> ReadInheritedTextRuns(PptxRenderContext context)
    {
        return ReadInheritedTextSpans(context).Select(span => span.Run).ToArray();
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadInheritedTextSpans(PptxRenderContext context)
    {
        return context.InheritedXml
            .SelectMany(xml => ReadTextSpans(context, xml, includePlaceholders: false, placeholderSources: []))
            .ToArray();
    }

    private static IReadOnlyList<TextRun> ReadSlideTextRuns(PptxRenderContext context)
    {
        return ReadSlideTextSpans(context).Select(span => span.Run).ToArray();
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadSlideTextSpans(PptxRenderContext context)
    {
        return ReadTextSpans(context, context.SlideXml, includePlaceholders: true, context.InheritedXml);
    }

    private static IReadOnlyList<TextRun> ReadTextRuns(
        PptxRenderContext context,
        XDocument slideXml,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return ReadTextSpans(context, slideXml, includePlaceholders, placeholderSources).Select(span => span.Run).ToArray();
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpans(
        PptxRenderContext context,
        XDocument slideXml,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return FlattenTextLayoutToSpans(BuildTextLayoutModel(context, slideXml, includePlaceholders, placeholderSources));
    }

    private static PptxTextLayoutModel BuildTextLayoutModel(
        PptxRenderContext context,
        XDocument slideXml,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return BuildTextLayoutModel(slideXml, context.Document, context.Theme, context.SlideNumber, includePlaceholders, placeholderSources);
    }

    private static PptxTextLayoutModel BuildTextLayoutModel(
        XDocument slideXml,
        PptxDocument document,
        PptxTheme theme,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        var advanceEstimator = new TextAdvanceEstimator();
        var frames = new List<PptxTextFrameLayout>();
        PptxTextFlowModel flow = BuildTextFlowModel(slideXml, document, theme, slideNumber, includePlaceholders, placeholderSources);
        foreach (PptxTextFlowFrame frame in flow.Frames)
        {
            frames.Add(BuildTextFrameLayout(frame, document, advanceEstimator));
        }

        return new PptxTextLayoutModel(frames);
    }

    private static PptxTextFlowModel BuildTextFlowModel(
        PptxRenderContext context,
        XDocument slideXml,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return BuildTextFlowModel(slideXml, context.Document, context.Theme, context.SlideNumber, includePlaceholders, placeholderSources);
    }

    private static PptxTextFlowModel BuildTextFlowModel(
        XDocument slideXml,
        PptxDocument document,
        PptxTheme theme,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return BuildTextFlowModel(BuildTextFrameModels(slideXml, document, theme, slideNumber, includePlaceholders, placeholderSources), document);
    }

    private static PptxTextFlowModel BuildTextFlowModel(IReadOnlyList<PptxTextFrameModel> frames, PptxDocument document)
    {
        return new PptxTextFlowModel(frames.Select(frame => BuildTextFlowFrame(frame, document)).ToArray());
    }

    private static PptxTextFlowFrame BuildTextFlowFrame(PptxTextFrameModel frame, PptxDocument document)
    {
        double yTop = OoxUnits.EmuToPoints(frame.Bounds.Y);
        var box = new PptxTextFlowBox(
            yTop,
            document.SlideHeightPoints - yTop - frame.Insets.Top - frame.VerticalOffset,
            frame.TextX,
            frame.TextWidth,
            frame.TextHeight,
            frame.TextClipY,
            frame.TextClipHeight,
            frame.RotationCenterX,
            frame.RotationCenterY);
        return new PptxTextFlowFrame(frame, box, frame.Paragraphs.Select(BuildTextFlowParagraph).ToArray());
    }

    private static PptxTextFlowParagraph BuildTextFlowParagraph(PptxTextParagraphModel paragraph)
    {
        return new PptxTextFlowParagraph(paragraph, paragraph.Style, paragraph.Runs.Select(run => BuildTextFlowRun(run, paragraph.Style.DefaultRunProperties)).ToArray());
    }

    private static PptxTextFlowRun BuildTextFlowRun(PptxTextRunModel run, XElement? defaultRunProperties)
    {
        if (run.Kind == PptxTextRunKind.Break)
        {
            return new PptxTextFlowRun(run, run.Style, [new PptxTextFlowSegment("\n", "\n", PptxTextFlowSegmentKind.Break, Draw: false, PreventCoalesce: true)]);
        }

        var segments = new List<PptxTextFlowSegment>();
        string[] tabParts = run.Text.Split('\t');
        for (int tabPartIndex = 0; tabPartIndex < tabParts.Length; tabPartIndex++)
        {
            if (tabPartIndex > 0)
            {
                segments.Add(new PptxTextFlowSegment(" ", " ", PptxTextFlowSegmentKind.Tab, Draw: true, PreventCoalesce: true));
            }

            foreach (TextCapsFragment fragment in ApplyTextCaps(tabParts[tabPartIndex], run.Properties, defaultRunProperties))
            {
                if (fragment.Text.Length == 0)
                {
                    continue;
                }

                foreach (PptxTextFlowSegment segment in SplitFlowSegments(fragment.Text))
                {
                    segments.Add(segment with
                    {
                        FontScale = fragment.FontScale
                    });
                }
            }
        }

        return new PptxTextFlowRun(run, run.Style, segments);
    }

    private static IReadOnlyList<TextRun> FlattenTextLayout(PptxTextLayoutModel layout)
    {
        return FlattenTextLayoutToSpans(layout).Select(span => span.Run).ToArray();
    }

    private static IReadOnlyList<PptxPositionedTextSpan> FlattenTextLayoutToSpans(PptxTextLayoutModel layout)
    {
        return layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans.Select(span => new PptxPositionedTextSpan(span.SourceRun, line.Box, span.Run, span.EndX, span.Atoms, span.GlyphSpan)))
            .ToArray();
    }

    private static PptxTextFrameLayout BuildTextFrameLayout(PptxTextFlowFrame flowFrame, PptxDocument document, TextAdvanceEstimator advanceEstimator)
    {
        PptxTextFrameModel frame = flowFrame.Model;
        double cursorLineTop = flowFrame.Box.CursorTop;
        int autoNumberValue = 1;
        bool hasPlacedParagraph = false;
        var paragraphLayouts = new List<PptxTextParagraphLayout>();

        foreach (PptxTextFlowParagraph flowParagraph in flowFrame.Paragraphs)
        {
            PptxTextParagraphModel paragraph = flowParagraph.Model;
            var lineLayouts = new List<PptxTextLineLayout>();
            ResolvedParagraphTextStyle paragraphStyle = flowParagraph.Style;
            if (!ParagraphHasVisibleContent(paragraph.Source))
            {
                if (ParagraphHasLayoutContent(paragraph.Source))
                {
                    XElement? endRunProperties = paragraph.Source.Element(DrawingNamespace + "endParaRPr");
                    double emptyFontSize = ReadFontSize(endRunProperties, paragraphStyle.DefaultRunProperties) * frame.FontScale;
                    double emptySpacingBefore = ReadParagraphSpacing(paragraph.Properties, paragraph.DefaultProperties, "spcBef", emptyFontSize);
                    double emptySpacingAfter = ReadParagraphSpacing(paragraph.Properties, paragraph.DefaultProperties, "spcAft", emptyFontSize);
                    cursorLineTop -= (hasPlacedParagraph ? emptySpacingBefore : 0d) + ReadParagraphAdvance(paragraphStyle.LineSpacing, emptyFontSize) + emptySpacingAfter;
                    hasPlacedParagraph = true;
                }

                paragraphLayouts.Add(new PptxTextParagraphLayout(paragraph, lineLayouts));
                continue;
            }

            string? bulletText = ReadBulletText(paragraph.Properties, ref autoNumberValue);
            bool bulletPending = bulletText is not null;
            double bulletX = frame.TextX + Math.Max(0d, paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging);
            double paragraphTextX = bulletText is null
                ? frame.TextX + Math.Max(0d, paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                : frame.TextX + Math.Max(0d, paragraphStyle.Indent.MarginLeft);
            if (hasPlacedParagraph)
            {
                cursorLineTop -= paragraphStyle.SpacingBefore;
            }

            bool afterManualLineBreak = false;
            double cursorY = cursorLineTop - ReadFirstLineBaselineOffset(paragraph, paragraphStyle.LineSpacing, advanceEstimator);
            double cursorX = paragraphTextX;
            double maxFontSize = paragraphStyle.FontSize;
            var line = new TextLayoutLine(paragraphTextX);
            foreach (PptxTextFlowRun flowRun in flowParagraph.Runs)
            {
                PptxTextRunModel modelRun = flowRun.Source;
                if (modelRun.Kind == PptxTextRunKind.Break)
                {
                    AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, maxFontSize, line, advanceEstimator), paragraphStyle.Alignment, frame.TextX, frame.TextWidth, justify: false, advanceEstimator);
                    cursorLineTop -= ReadManualBreakLineAdvance(paragraphStyle.LineSpacing, maxFontSize);
                    cursorY = double.NaN;
                    afterManualLineBreak = true;
                    cursorX = paragraphTextX;
                    line.Reset(paragraphTextX);
                    maxFontSize = paragraphStyle.FontSize;
                    continue;
                }

                ResolvedRunTextStyle runStyle = flowRun.Style;
                maxFontSize = Math.Max(maxFontSize, runStyle.NominalFontSize);
                if (double.IsNaN(cursorY))
                {
                    cursorY = cursorLineTop - (afterManualLineBreak
                        ? ManualBreakBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing)
                        : LineBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator));
                    afterManualLineBreak = false;
                }

                if (bulletPending)
                {
                    BulletStyle bulletStyle = ReadBulletStyle(paragraph.Properties, frame.Theme, runStyle.FontSize, runStyle.Color, runStyle.Typeface);
                    double bulletWidth = Math.Max(1d, frame.TextWidth - (bulletX - frame.TextX));
                    double bulletEndX = bulletX + advanceEstimator.Measure(bulletText!, bulletStyle.FontSize, bulletStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing);
                    TextRun bulletRun = new(bulletText!, bulletX, cursorY, bulletWidth, frame.TextHeight, frame.TextX, frame.TextClipY, frame.TextWidth, frame.TextClipHeight, bulletStyle.FontSize, runStyle.CharacterSpacing, 0d, bulletStyle.Color, 1d, null, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, bulletStyle.Typeface, frame.Bounds.RotationDegrees, frame.RotationCenterX, frame.RotationCenterY);
                    line.Add(modelRun, bulletRun, bulletEndX, BuildTextAtoms(bulletRun, advanceEstimator, PptxTextAtomKind.Word), BuildGlyphSpan(bulletRun, advanceEstimator));
                    bulletPending = false;
                }

                foreach (PptxTextFlowSegment flowSegment in flowRun.Segments)
                {
                    if (flowSegment.Kind == PptxTextFlowSegmentKind.Tab)
                    {
                        double tabSpaceWidth = advanceEstimator.Measure(" ", runStyle.FontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                        TextRun tabRun = new(" ", cursorX, cursorY, Math.Max(1d, tabSpaceWidth), frame.TextHeight, frame.TextX, frame.TextClipY, frame.TextWidth, frame.TextClipHeight, runStyle.FontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.Bounds.RotationDegrees, frame.RotationCenterX, frame.RotationCenterY, PreventCoalesce: true);
                        line.Add(modelRun, tabRun, cursorX + tabSpaceWidth, BuildTextAtoms(tabRun, advanceEstimator, PptxTextAtomKind.Tab), BuildGlyphSpan(tabRun, advanceEstimator));
                        cursorX = ResolveNextTabX(cursorX, paragraphTextX, paragraphStyle.TabStops);
                        line.AdvanceTo(cursorX);
                        continue;
                    }

                    double fragmentFontSize = runStyle.FontSize * flowSegment.FontScale;
                    string currentSegment = flowSegment.Text;
                    string currentAdvanceText = flowSegment.AdvanceText;
                    double segmentWidth = MeasureFlowSegmentAdvance(advanceEstimator, flowSegment, currentAdvanceText, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                    bool overflowsLine = cursorX > paragraphTextX &&
                        (cursorX + segmentWidth > frame.TextX + frame.TextWidth ||
                            (runStyle.CharacterSpacing > 0d && cursorX + segmentWidth > frame.TextX + frame.TextWidth - fragmentFontSize));
                    if (overflowsLine)
                    {
                        AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, maxFontSize, line, advanceEstimator), paragraphStyle.Alignment, frame.TextX, frame.TextWidth, justify: paragraphStyle.Alignment == TextAlignment.Justify, advanceEstimator);
                        cursorLineTop -= ReadLineAdvance(paragraphStyle.LineSpacing, maxFontSize);
                        cursorY = cursorLineTop - LineBaselineOffset(fragmentFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator);
                        cursorX = paragraphTextX;
                        line.Reset(paragraphTextX);
                        maxFontSize = Math.Max(runStyle.NominalFontSize, fragmentFontSize);
                        currentSegment = currentSegment.TrimStart();
                        currentAdvanceText = currentAdvanceText.TrimStart();
                        segmentWidth = MeasureFlowSegmentAdvance(advanceEstimator, flowSegment, currentAdvanceText, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                    }

                    if (currentAdvanceText.Length == 0 && flowSegment.AdvanceFontSizeFactor is null)
                    {
                        continue;
                    }

                    if (flowSegment.Draw && currentSegment.Length != 0)
                    {
                        TextRun textRun = new(currentSegment, cursorX, cursorY, Math.Max(1d, segmentWidth), frame.TextHeight, frame.TextX, frame.TextClipY, frame.TextWidth, frame.TextClipHeight, fragmentFontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.Bounds.RotationDegrees, frame.RotationCenterX, frame.RotationCenterY, flowSegment.PreventCoalesce);
                        line.Add(modelRun, textRun, cursorX + segmentWidth, BuildTextAtoms(textRun, advanceEstimator), BuildGlyphSpan(textRun, advanceEstimator));
                    }

                    cursorX += segmentWidth;
                    line.AdvanceTo(cursorX);
                }
            }

            AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, maxFontSize, line, advanceEstimator), paragraphStyle.Alignment, frame.TextX, frame.TextWidth, justify: false, advanceEstimator);
            cursorLineTop -= ReadParagraphAdvance(paragraphStyle.LineSpacing, maxFontSize) + paragraphStyle.SpacingAfter;
            hasPlacedParagraph = true;
            paragraphLayouts.Add(new PptxTextParagraphLayout(paragraph, lineLayouts));
        }

        return new PptxTextFrameLayout(frame, paragraphLayouts);
    }
    private static GroupTransform ReadAncestorGroupTransform(XElement shape)
    {
        GroupTransform transform = GroupTransform.Identity;
        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp").Reverse())
        {
            transform = transform.Combine(ReadGroupTransform(group));
        }

        return transform;
    }

    private static IReadOnlyList<TextRun> ReadTextRunsForShape(
        XElement shape,
        PptxRenderContext context,
        bool includePlaceholders)
    {
        return ReadTextSpansForShape(shape, context, includePlaceholders).Select(span => span.Run).ToArray();
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpansForShape(
        XElement shape,
        PptxRenderContext context,
        bool includePlaceholders)
    {
        return ReadTextSpansForShape(shape, context.Document, context.Theme, context.SlideNumber, includePlaceholders, context.InheritedXml);
    }

    private static IReadOnlyList<TextRun> ReadTextRunsForShape(
        XElement shape,
        PptxDocument document,
        PptxTheme theme,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return ReadTextSpansForShape(shape, document, theme, slideNumber, includePlaceholders, placeholderSources)
            .Select(span => span.Run)
            .ToArray();
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpansForShape(
        XElement shape,
        PptxDocument document,
        PptxTheme theme,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        XElement current = new(shape);
        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp"))
        {
            var groupCopy = new XElement(PresentationNamespace + "grpSp");
            if (group.Element(PresentationNamespace + "grpSpPr") is { } properties)
            {
                groupCopy.Add(new XElement(properties));
            }

            groupCopy.Add(current);
            current = groupCopy;
        }

        var slide = new XDocument(
            new XElement(PresentationNamespace + "sld",
                new XElement(PresentationNamespace + "cSld",
                    new XElement(PresentationNamespace + "spTree", current))));
        return FlattenTextLayoutToSpans(BuildTextLayoutModel(slide, document, theme, slideNumber, includePlaceholders, placeholderSources));
    }

    private static IReadOnlyList<XElement> FindInheritedPlaceholderShapes(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        XElement? placeholder = shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph");
        if (placeholder is null)
        {
            return [];
        }

        var matches = new List<XElement>();
        string? type = (string?)placeholder.Attribute("type");
        string? index = (string?)placeholder.Attribute("idx");
        foreach (XDocument source in placeholderSources)
        {
            foreach (XElement candidate in source.Descendants(PresentationNamespace + "sp"))
            {
                XElement? candidatePlaceholder = candidate
                    .Element(PresentationNamespace + "nvSpPr")
                    ?.Element(PresentationNamespace + "nvPr")
                    ?.Element(PresentationNamespace + "ph");
                if (candidatePlaceholder is null)
                {
                    continue;
                }

                string? candidateType = (string?)candidatePlaceholder.Attribute("type");
                string? candidateIndex = (string?)candidatePlaceholder.Attribute("idx");
                bool indexMatches = index is not null && candidateIndex == index;
                bool typeMatches = index is null && type is not null && candidateType == type;
                if (!indexMatches && !typeMatches)
                {
                    continue;
                }

                matches.Add(candidate);
                break;
            }
        }

        return matches;
    }

    private static XElement? FindInheritedTextStyle(XElement shape, IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        string? placeholderType = (string?)shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph")
            ?.Attribute("type");
        string styleName = placeholderType switch
        {
            "title" or "ctrTitle" => "titleStyle",
            "body" or "subTitle" => "bodyStyle",
            _ => "otherStyle"
        };

        foreach (XDocument source in placeholderSources)
        {
            XElement? style = source.Root?
                .Element(PresentationNamespace + "txStyles")
                ?.Element(PresentationNamespace + styleName);
            XElement? level = style?.Element(DrawingNamespace + levelName) ??
                style?.Element(DrawingNamespace + "defPPr");
            if (level is not null)
            {
                return level;
            }
        }

        return null;
    }

    private static XElement? FindDefaultTextStyle(IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        foreach (XDocument source in placeholderSources)
        {
            XElement? defaultTextStyle = source.Root?.Element(PresentationNamespace + "defaultTextStyle");
            XElement? level = defaultTextStyle?.Element(DrawingNamespace + levelName) ??
                defaultTextStyle?.Element(DrawingNamespace + "defPPr");
            if (level is not null)
            {
                return level;
            }
        }

        return null;
    }

    private static IEnumerable<PptxTextFlowSegment> SplitFlowSegments(string text)
    {
        int index = 0;
        while (index < text.Length)
        {
            int start = index;
            while (index < text.Length && text[index] == ' ')
            {
                index++;
            }

            while (index < text.Length && text[index] != ' ')
            {
                index++;
                if (text[index - 1] == '-' && index < text.Length && text[index] != ' ')
                {
                    break;
                }
            }

            while (index < text.Length && text[index] == ' ')
            {
                index++;
            }

            if (index > start)
            {
                foreach (PptxTextFlowSegment segment in SplitControlSegments(text[start..index]))
                {
                    yield return segment;
                }
            }
        }
    }

    private static IEnumerable<PptxTextFlowSegment> SplitControlSegments(string text)
    {
        var builder = new StringBuilder();
        bool nextPreventsCoalesce = false;
        foreach (char c in text)
        {
            if (c == '\u00A0' || c == '\u202F')
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce);
                    builder.Clear();
                }

                yield return new PptxTextFlowSegment(string.Empty, c.ToString(), PptxTextFlowSegmentKind.HiddenAdvance, Draw: false, PreventCoalesce: true);
                nextPreventsCoalesce = true;
                continue;
            }

            if (c == '\u00AD')
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce);
                    builder.Clear();
                }

                nextPreventsCoalesce = true;
                continue;
            }

            if (IsOfficeTextOperationBoundaryPunctuation(c))
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce);
                    builder.Clear();
                }

                yield return new PptxTextFlowSegment(c.ToString(), c.ToString(), PptxTextFlowSegmentKind.BoundaryPunctuation, Draw: true, PreventCoalesce: true);
                nextPreventsCoalesce = true;
                continue;
            }

            builder.Append(c);
        }

        if (builder.Length > 0)
        {
            yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce);
        }
    }

    private static bool IsOfficeTextOperationBoundaryPunctuation(char value)
    {
        return value is '-' or '\u2010' or '\u2011' or '\u2012' or '\u2013';
    }

    private static double MeasureFlowSegmentAdvance(TextAdvanceEstimator advanceEstimator, PptxTextFlowSegment segment, string advanceText, double fontSize, string? typeface, bool bold, bool italic, double characterSpacing, bool kerningEnabled)
    {
        return segment.AdvanceFontSizeFactor is { } factor
            ? Math.Max(0d, fontSize * factor)
            : advanceEstimator.Measure(advanceText, fontSize, typeface, bold, italic, characterSpacing, kerningEnabled);
    }

    private static IReadOnlyList<PptxTextAtomLayout> BuildTextAtoms(TextRun run, TextAdvanceEstimator advanceEstimator, PptxTextAtomKind? forcedKind = null)
    {
        if (run.Text.Length == 0)
        {
            return [];
        }

        if (forcedKind is { } kind)
        {
            return [new PptxTextAtomLayout(kind, run.Text, run.X, run.Width, Draw: kind != PptxTextAtomKind.HiddenAdvance)];
        }

        var atoms = new List<PptxTextAtomLayout>();
        double cursorX = run.X;
        int index = 0;
        while (index < run.Text.Length)
        {
            int start = index;
            bool isSpace = run.Text[index] == ' ';
            while (index < run.Text.Length && (run.Text[index] == ' ') == isSpace)
            {
                index++;
            }

            string text = run.Text[start..index];
            double width = advanceEstimator.Measure(text, run.FontSize, run.FontFamily, run.Bold, run.Italic, run.CharacterSpacing, run.KerningEnabled);
            atoms.Add(new PptxTextAtomLayout(isSpace ? PptxTextAtomKind.Space : PptxTextAtomKind.Word, text, cursorX, width, Draw: true));
            cursorX += width;
        }

        if (atoms.Count > 0)
        {
            PptxTextAtomLayout last = atoms[^1];
            double delta = run.X + run.Width - (last.X + last.Width);
            if (Math.Abs(delta) > 0.001d)
            {
                atoms[^1] = last with { Width = Math.Max(0d, last.Width + delta) };
            }
        }

        return atoms;
    }

    private static PptxTextGlyphSpanLayout BuildGlyphSpan(TextRun run, TextAdvanceEstimator advanceEstimator)
    {
        OpenTypeFont? font = advanceEstimator.ResolveOpenTypeFont(run.FontFamily, run.Bold, run.Italic);
        if (font is null || font.UnitsPerEm == 0)
        {
            return PptxTextGlyphSpanLayout.Empty(run);
        }

        var glyphs = new List<PptxTextGlyphLayout>();
        ushort previousGlyph = 0;
        foreach (Rune rune in run.Text.EnumerateRunes())
        {
            ushort glyph = font.MapCodePoint(rune.Value);
            if (glyph == 0)
            {
                continue;
            }

            double adjustmentBefore = 0d;
            if (glyphs.Count > 0)
            {
                adjustmentBefore += run.CharacterSpacing;
                if (run.KerningEnabled && previousGlyph != 0)
                {
                    adjustmentBefore += font.GetKerning(previousGlyph, glyph) * run.FontSize / font.UnitsPerEm;
                }
            }

            double advance = font.GetAdvanceWidth(glyph) * run.FontSize / font.UnitsPerEm;
            glyphs.Add(new PptxTextGlyphLayout(rune.Value, glyph, advance, adjustmentBefore));
            previousGlyph = glyph;
        }

        double naturalWidth = glyphs.Sum(glyph => glyph.Advance) + glyphs.Sum(glyph => glyph.AdjustmentBefore);
        return new PptxTextGlyphSpanLayout(
            run.Text,
            run.FontFamily,
            run.Bold,
            run.Italic,
            run.FontSize,
            run.CharacterSpacing,
            run.KerningEnabled,
            Math.Max(0d, naturalWidth),
            run.Width,
            glyphs);
    }

    private static PptxTextLineBoxLayout CreateLineBox(
        double lineTopY,
        double baselineY,
        LineSpacing lineSpacing,
        double maxFontSize,
        TextLayoutLine line,
        TextAdvanceEstimator advanceEstimator)
    {
        double advance = ReadLineAdvance(lineSpacing, maxFontSize);
        PptxTextSpanLayout? baselineSpan = line.Spans.FirstOrDefault();
        ResolvedRunTextStyle? baselineStyle = baselineSpan?.SourceRun?.Style;
        double baselineFontSize = baselineSpan?.Run.FontSize ?? maxFontSize;
        PptxTextBaselineMetricLayout baselineMetric = ReadBaselineMetric(baselineFontSize, baselineStyle, advanceEstimator);
        return new PptxTextLineBoxLayout(
            lineTopY,
            baselineY,
            advance,
            lineTopY - baselineY,
            maxFontSize,
            lineSpacing,
            baselineMetric);
    }

    private static void AddAlignedParagraphLine(
        List<PptxTextLineLayout> lines,
        TextLayoutLine line,
        PptxTextLineBoxLayout box,
        TextAlignment alignment,
        double textX,
        double textWidth,
        bool justify,
        TextAdvanceEstimator advanceEstimator)
    {
        if (line.Spans.Count == 0)
        {
            return;
        }

        double paragraphWidth = Math.Max(0d, line.EndX - textX);
        bool justifyLine = justify && paragraphWidth > 0d && paragraphWidth < textWidth;
        double offset = alignment switch
        {
            TextAlignment.Center => Math.Max(0d, textWidth - paragraphWidth) / 2d,
            TextAlignment.Right => Math.Max(0d, textWidth - paragraphWidth),
            _ => 0d
        };

        if (justifyLine)
        {
            PptxTextLineLayout? justified = TryJustifyLine(line, box, textX, textWidth, advanceEstimator);
            if (justified is not null)
            {
                lines.Add(justified);
                return;
            }
        }

        PptxTextSpanLayout[] spans = line.Spans
            .Select(span => span with
            {
                Run = span.Run with
                {
                    X = span.Run.X + offset,
                    Width = span.Run.Width,
                    Alignment = TextAlignment.Left
                },
                EndX = span.EndX + offset,
                Atoms = OffsetAtoms(span.Atoms, offset)
            })
            .ToArray();
        lines.Add(new PptxTextLineLayout(box, textX + offset, line.EndX + offset, alignment, spans));
    }

    private static IReadOnlyList<PptxTextAtomLayout> OffsetAtoms(IReadOnlyList<PptxTextAtomLayout> atoms, double offset)
    {
        if (Math.Abs(offset) <= 0.001d)
        {
            return atoms;
        }

        return atoms.Select(atom => atom with { X = atom.X + offset }).ToArray();
    }

    private static PptxTextLineLayout? TryJustifyLine(TextLayoutLine line, PptxTextLineBoxLayout box, double textX, double textWidth, TextAdvanceEstimator advanceEstimator)
    {
        int spaceCount = line.Spans.Sum(span => span.Run.Text.Count(static c => c == ' '));
        if (spaceCount == 0)
        {
            return null;
        }

        double extraWidth = textWidth - Math.Max(0d, line.EndX - textX);
        if (extraWidth <= 0.001d)
        {
            return null;
        }

        double extraPerSpace = extraWidth / spaceCount;
        double shift = 0d;
        var spans = new List<PptxTextSpanLayout>(line.Spans.Count);
        foreach (PptxTextSpanLayout span in line.Spans)
        {
            int spanSpaces = span.Run.Text.Count(static c => c == ' ');
            double spanExtra = spanSpaces * extraPerSpace;
            PptxTextSpanLayout justifiedSpan = span with
            {
                Run = span.Run with
                {
                    X = span.Run.X + shift,
                    Width = span.Run.Width + spanExtra,
                    Alignment = TextAlignment.Left,
                    PreventCoalesce = true
                },
                EndX = span.EndX + shift + spanExtra,
                Atoms = JustifyAtoms(span.Atoms, shift, extraPerSpace),
                GlyphSpan = span.GlyphSpan with { LayoutWidth = span.GlyphSpan.LayoutWidth + spanExtra }
            };
            foreach (PptxTextSpanLayout wordSpan in SplitJustifiedWordSpans(justifiedSpan, advanceEstimator))
            {
                spans.Add(wordSpan);
            }

            shift += spanExtra;
        }

        return new PptxTextLineLayout(box, textX, textX + textWidth, TextAlignment.Justify, spans);
    }

    private static IEnumerable<PptxTextSpanLayout> SplitJustifiedWordSpans(PptxTextSpanLayout span, TextAdvanceEstimator advanceEstimator)
    {
        PptxTextAtomLayout[] words = span.Atoms
            .Where(static atom => atom.Kind == PptxTextAtomKind.Word && atom.Draw && atom.Text.Length != 0)
            .ToArray();
        if (words.Length == 0)
        {
            yield return span;
            yield break;
        }

        foreach (PptxTextAtomLayout word in words)
        {
            TextRun wordRun = span.Run with
            {
                Text = word.Text,
                X = word.X,
                Width = word.Width,
                PreventCoalesce = true
            };
            yield return span with
            {
                Run = wordRun,
                EndX = word.X + word.Width,
                Atoms = [word],
                GlyphSpan = BuildGlyphSpan(wordRun, advanceEstimator)
            };
        }
    }

    private static IReadOnlyList<PptxTextAtomLayout> JustifyAtoms(IReadOnlyList<PptxTextAtomLayout> atoms, double initialShift, double extraPerSpace)
    {
        double shift = initialShift;
        var justified = new PptxTextAtomLayout[atoms.Count];
        for (int i = 0; i < atoms.Count; i++)
        {
            PptxTextAtomLayout atom = atoms[i];
            double extra = atom.Kind == PptxTextAtomKind.Space
                ? atom.Text.Count(static c => c == ' ') * extraPerSpace
                : 0d;
            justified[i] = atom with
            {
                X = atom.X + shift,
                Width = atom.Width + extra
            };
            shift += extra;
        }

        return justified;
    }

    private static XElement? MergeParagraphProperties(params XElement?[] sources)
    {
        XElement? merged = null;
        foreach (XElement? source in sources.Reverse())
        {
            if (source is null)
            {
                continue;
            }

            merged = merged is null
                ? new XElement(source)
                : OverlayParagraphProperties(source, merged);
        }

        return merged;
    }

    private static XElement OverlayParagraphProperties(XElement primary, XElement fallback)
    {
        XElement merged = new(primary);
        foreach (XAttribute attribute in fallback.Attributes())
        {
            if (merged.Attribute(attribute.Name) is null)
            {
                merged.Add(new XAttribute(attribute));
            }
        }

        MergeChildElement(merged, fallback, DrawingNamespace + "defRPr");
        return merged;
    }

    private static void MergeChildElement(XElement primaryParent, XElement fallbackParent, XName childName)
    {
        XElement? primary = primaryParent.Element(childName);
        XElement? fallback = fallbackParent.Element(childName);
        if (fallback is null)
        {
            return;
        }

        if (primary is null)
        {
            primaryParent.Add(new XElement(fallback));
            return;
        }

        foreach (XAttribute attribute in fallback.Attributes())
        {
            if (primary.Attribute(attribute.Name) is null)
            {
                primary.Add(new XAttribute(attribute));
            }
        }

        foreach (XElement child in fallback.Elements())
        {
            if (primary.Element(child.Name) is null)
            {
                primary.Add(new XElement(child));
            }
        }
    }

    private static bool ParagraphHasVisibleContent(XElement paragraph)
    {
        return paragraph.Elements().Any(child =>
            child.Name == DrawingNamespace + "r" ||
            child.Name == DrawingNamespace + "fld" ||
            child.Name == DrawingNamespace + "br");
    }

    private static bool ParagraphHasLayoutContent(XElement paragraph)
    {
        return paragraph.Element(DrawingNamespace + "pPr") is not null ||
            paragraph.Element(DrawingNamespace + "endParaRPr") is not null;
    }

    private static ResolvedParagraphTextStyle ResolveParagraphTextStyle(
        XElement paragraph,
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        double fontScale)
    {
        XElement? defaultRunProperties = paragraphProperties?.Element(DrawingNamespace + "defRPr") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
        double fontSize = ReadFirstParagraphFontSize(paragraph, defaultRunProperties) * fontScale;
        return new ResolvedParagraphTextStyle(
            ReadAlignment(paragraph, defaultParagraphProperties),
            paragraphProperties,
            defaultRunProperties,
            fontSize,
            ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef", fontSize),
            ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", fontSize),
            ReadLineSpacing(paragraphProperties, defaultParagraphProperties),
            ReadParagraphIndent(paragraphProperties),
            ReadTabStops(paragraphProperties));
    }

    private static ResolvedRunTextStyle ResolveRunTextStyle(
        XElement? runProperties,
        XElement? defaultRunProperties,
        RgbColor? shapeFontColor,
        PptxTheme theme,
        double fontScale)
    {
        double nominalFontSize = ReadFontSize(runProperties, defaultRunProperties) * fontScale;
        double baselineOffset = ReadBaselineOffset(runProperties, defaultRunProperties, nominalFontSize);
        double fontSize = Math.Abs(baselineOffset) > 0.001d
            ? nominalFontSize * 2d / 3d
            : nominalFontSize;
        double alpha = 1d;
        RgbColor color;
        if (TryReadSolidColorWithAlpha(runProperties, theme, out RgbColor runColor, out double runAlpha))
        {
            color = runColor;
            alpha = runAlpha;
        }
        else if (TryReadSolidColorWithAlpha(defaultRunProperties, theme, out RgbColor defaultColor, out double defaultAlpha))
        {
            color = defaultColor;
            alpha = defaultAlpha;
        }
        else
        {
            color = shapeFontColor ?? new RgbColor(0, 0, 0);
        }

        string? typeface = theme.ResolveTypeface((string?)(runProperties?
            .Element(DrawingNamespace + "latin") ??
            defaultRunProperties?.Element(DrawingNamespace + "latin"))
            ?.Attribute("typeface"));
        bool bold = ParseOptionalBoolAttribute(runProperties, "b") ||
            (runProperties?.Attribute("b") is null && ParseOptionalBoolAttribute(defaultRunProperties, "b"));
        bool italic = ParseOptionalBoolAttribute(runProperties, "i") ||
            (runProperties?.Attribute("i") is null && ParseOptionalBoolAttribute(defaultRunProperties, "i"));
        bool underline = ((string?)(runProperties?.Attribute("u") ?? defaultRunProperties?.Attribute("u"))) is { } underlineValue
            && !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);

        return new ResolvedRunTextStyle(
            nominalFontSize,
            fontSize,
            ReadCharacterSpacing(runProperties, defaultRunProperties),
            baselineOffset,
            color,
            alpha,
            TryReadHighlightColor(runProperties, out RgbColor highlightColor) ? highlightColor : null,
            bold,
            italic,
            underline,
            IsStrikeEnabled(runProperties, defaultRunProperties),
            IsKerningEnabled(runProperties, defaultRunProperties, fontSize),
            typeface);
    }

    private static bool IsKerningEnabled(XElement? runProperties, XElement? defaultRunProperties, double fontSize)
    {
        XAttribute? threshold = runProperties?.Attribute("kern") ?? defaultRunProperties?.Attribute("kern");
        if (threshold is null)
        {
            return true;
        }

        double minimumFontSize = int.Parse(threshold.Value, CultureInfo.InvariantCulture) / 100d;
        return minimumFontSize <= 0d || fontSize >= minimumFontSize;
    }

    private static bool IsTextRunElement(XElement element)
    {
        return element.Name == DrawingNamespace + "r" ||
            element.Name == DrawingNamespace + "fld";
    }

    private static string ReadTextElementText(XElement element, int slideNumber)
    {
        if (slideNumber > 0 &&
            element.Name == DrawingNamespace + "fld" &&
            string.Equals((string?)element.Attribute("type"), "slidenum", StringComparison.OrdinalIgnoreCase))
        {
            return slideNumber.ToString(CultureInfo.InvariantCulture);
        }

        return NormalizeText((string?)element.Element(DrawingNamespace + "t") ?? string.Empty);
    }

    private static string NormalizeText(string text)
    {
        return text;
    }

    private static double ReadFontSize(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("sz") ?? defaultRunProperties?.Attribute("sz")) is { } size
            ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
            : 18d;
    }

    private static double ReadCharacterSpacing(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("spc") ?? defaultRunProperties?.Attribute("spc")) is { } spacing
            ? int.Parse(spacing.Value, CultureInfo.InvariantCulture) / 100d
            : 0d;
    }

    private static double ReadBaselineOffset(XElement? runProperties, XElement? defaultRunProperties, double fontSize)
    {
        return (runProperties?.Attribute("baseline") ?? defaultRunProperties?.Attribute("baseline")) is { } baseline
            ? fontSize * int.Parse(baseline.Value, CultureInfo.InvariantCulture) / 100000d
            : 0d;
    }

    private static bool IsStrikeEnabled(XElement? runProperties, XElement? defaultRunProperties)
    {
        string? value = (string?)(runProperties?.Attribute("strike") ?? defaultRunProperties?.Attribute("strike"));
        return value is not null && !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyList<TextCapsFragment> ApplyTextCaps(string text, XElement? runProperties, XElement? defaultRunProperties)
    {
        string? value = (string?)(runProperties?.Attribute("cap") ?? defaultRunProperties?.Attribute("cap"));
        if (text.Length == 0)
        {
            return [];
        }

        if (value is "all")
        {
            return [new TextCapsFragment(text.ToUpperInvariant(), 1d)];
        }

        if (value is not "small")
        {
            return [new TextCapsFragment(text, 1d)];
        }

        var fragments = new List<TextCapsFragment>();
        var builder = new StringBuilder();
        bool? currentSmall = null;
        foreach (char character in text)
        {
            bool isSmall = char.IsLetter(character) && char.IsLower(character);
            if (currentSmall is not null && currentSmall != isSmall)
            {
                fragments.Add(new TextCapsFragment(builder.ToString(), currentSmall.Value ? 0.8d : 1d));
                builder.Clear();
            }

            currentSmall = isSmall;
            builder.Append(char.ToUpperInvariant(character));
        }

        if (builder.Length > 0 && currentSmall is not null)
        {
            fragments.Add(new TextCapsFragment(builder.ToString(), currentSmall.Value ? 0.8d : 1d));
        }

        return fragments;
    }

    private static TextInsets ReadTextInsets(XElement textBody)
    {
        XElement? bodyProperties = textBody.Element(DrawingNamespace + "bodyPr");
        return new TextInsets(
            ReadInset(bodyProperties, "lIns", 91440),
            ReadInset(bodyProperties, "rIns", 91440),
            ReadInset(bodyProperties, "tIns", 45720),
            ReadInset(bodyProperties, "bIns", 45720));
    }

    private static double ReadNormAutofitFontScale(XElement textBody)
    {
        XElement? normAutofit = textBody
            .Element(DrawingNamespace + "bodyPr")
            ?.Element(DrawingNamespace + "normAutofit");
        if (normAutofit?.Attribute("fontScale") is not { } fontScale)
        {
            return 1d;
        }

        return Math.Clamp(int.Parse(fontScale.Value, CultureInfo.InvariantCulture) / 100000d, 0.01d, 10d);
    }

    private static bool ClipsVerticalOverflow(XElement textBody)
    {
        string? overflow = (string?)textBody
            .Element(DrawingNamespace + "bodyPr")
            ?.Attribute("vertOverflow");
        return overflow?.Equals("clip", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static double ReadInset(XElement? element, string attributeName, long defaultEmu)
    {
        long emu = element?.Attribute(attributeName) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : defaultEmu;
        return OoxUnits.EmuToPoints(emu);
    }

    private static double ReadParagraphSpacing(
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        string elementName,
        double referenceFontSize)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + elementName) ??
            defaultParagraphProperties?.Element(DrawingNamespace + elementName);
        if (spacing?.Element(DrawingNamespace + "spcPts")?.Attribute("val") is { } points)
        {
            return int.Parse(points.Value, CultureInfo.InvariantCulture) / 100d;
        }

        if (spacing?.Element(DrawingNamespace + "spcPct")?.Attribute("val") is { } percent)
        {
            return referenceFontSize * Math.Max(0d, int.Parse(percent.Value, CultureInfo.InvariantCulture) / 100000d);
        }

        return 0d;
    }

    private static ParagraphIndent ReadParagraphIndent(XElement? paragraphProperties)
    {
        return new ParagraphIndent(
            ReadParagraphEmuAttribute(paragraphProperties, "marL"),
            ReadParagraphEmuAttribute(paragraphProperties, "indent"));
    }

    private static double ReadParagraphEmuAttribute(XElement? paragraphProperties, string attributeName)
    {
        return paragraphProperties?.Attribute(attributeName) is { } attribute
            ? OoxUnits.EmuToPoints(long.Parse(attribute.Value, CultureInfo.InvariantCulture))
            : 0d;
    }

    private static IReadOnlyList<double> ReadTabStops(XElement? paragraphProperties)
    {
        if (paragraphProperties?.Element(DrawingNamespace + "tabLst") is not { } tabList)
        {
            return Array.Empty<double>();
        }

        return tabList
            .Elements(DrawingNamespace + "tab")
            .Select(tab => tab.Attribute("pos") is { } position
                ? OoxUnits.EmuToPoints(long.Parse(position.Value, CultureInfo.InvariantCulture))
                : double.NaN)
            .Where(position => !double.IsNaN(position))
            .Order()
            .ToArray();
    }

    private static double ResolveNextTabX(double cursorX, double paragraphTextX, IReadOnlyList<double> tabStops)
    {
        double current = cursorX - paragraphTextX;
        foreach (double tabStop in tabStops)
        {
            if (tabStop > current + 0.01d)
            {
                return paragraphTextX + tabStop;
            }
        }

        const long defaultTabStopEmus = 914400;
        double defaultTabStop = OoxUnits.EmuToPoints(defaultTabStopEmus);
        return paragraphTextX + Math.Ceiling((current + 0.01d) / defaultTabStop) * defaultTabStop;
    }

    private static LineSpacing ReadLineSpacing(XElement? paragraphProperties, XElement? defaultParagraphProperties)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + "lnSpc") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "lnSpc");
        if (spacing?.Element(DrawingNamespace + "spcPts")?.Attribute("val") is { } points)
        {
            return LineSpacing.Absolute(Math.Max(0.1d, int.Parse(points.Value, CultureInfo.InvariantCulture) / 100d));
        }

        if (spacing?.Element(DrawingNamespace + "spcPct")?.Attribute("val") is { } percent)
        {
            return LineSpacing.Multiple(Math.Max(0.1d, int.Parse(percent.Value, CultureInfo.InvariantCulture) / 100000d), true);
        }

        return LineSpacing.Multiple(1d, false);
    }

    private static double ReadParagraphAdvance(LineSpacing lineSpacing, double fontSize)
    {
        return lineSpacing.IsExplicit
            ? lineSpacing.Resolve(fontSize)
            : fontSize * 1.2d;
    }

    private static double ReadLineAdvance(LineSpacing lineSpacing, double fontSize)
    {
        return lineSpacing.IsExplicit
            ? lineSpacing.Resolve(fontSize)
            : fontSize <= 12d ? fontSize : fontSize * 1.2d;
    }

    private static double ReadFirstLineBaselineOffset(PptxTextParagraphModel paragraph, LineSpacing lineSpacing, TextAdvanceEstimator advanceEstimator)
    {
        PptxTextRunModel? firstRun = paragraph.Runs.FirstOrDefault(run => run.Kind != PptxTextRunKind.Break);
        double fontSize = firstRun?.Style.NominalFontSize ?? ReadFirstParagraphFontSize(paragraph.Source, paragraph.Style.DefaultRunProperties);
        return ParagraphHasManualLineBreak(paragraph.Source)
            ? ManualBreakBaselineOffset(fontSize, lineSpacing)
            : LineBaselineOffset(fontSize, lineSpacing, firstRun?.Style, advanceEstimator);
    }

    private static bool ParagraphHasManualLineBreak(XElement paragraph)
    {
        return paragraph.Elements(DrawingNamespace + "br").Any();
    }

    private static double ReadManualBreakLineAdvance(LineSpacing lineSpacing, double fontSize)
    {
        return lineSpacing.IsExplicit ? ReadLineAdvance(lineSpacing, fontSize) : fontSize * 1.24d;
    }

    private static double ReadFirstParagraphFontSize(XElement paragraph, XElement? defaultRunProperties)
    {
        const double defaultFontSize = 18d;
        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == DrawingNamespace + "br")
            {
                return defaultFontSize;
            }

            if (!IsTextRunElement(child))
            {
                continue;
            }

            XElement? runProperties = child.Element(DrawingNamespace + "rPr");
            if (runProperties?.Attribute("sz") is { } size)
            {
                return int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d;
            }

            if (defaultRunProperties?.Attribute("sz") is { } defaultSize)
            {
                return int.Parse(defaultSize.Value, CultureInfo.InvariantCulture) / 100d;
            }

            return defaultFontSize;
        }

        return defaultFontSize;
    }

    private static double LineBaselineOffset(double fontSize, LineSpacing lineSpacing)
    {
        if (lineSpacing.IsAbsolute)
        {
            return Math.Max(BaselineOffset(fontSize), lineSpacing.Value - fontSize * 0.374d);
        }

        return lineSpacing.IsExplicit
            ? lineSpacing.Resolve(fontSize) - fontSize * 0.234d
            : BaselineOffset(fontSize);
    }

    private static double LineBaselineOffset(double fontSize, LineSpacing lineSpacing, ResolvedRunTextStyle? style, TextAdvanceEstimator advanceEstimator)
    {
        if (lineSpacing.IsAbsolute)
        {
            return Math.Max(BaselineOffset(fontSize, style, advanceEstimator), lineSpacing.Value - fontSize * 0.374d);
        }

        return lineSpacing.IsExplicit
            ? lineSpacing.Resolve(fontSize) - fontSize * 0.234d
            : BaselineOffset(fontSize, style, advanceEstimator);
    }

    private static double ManualBreakBaselineOffset(double fontSize, LineSpacing lineSpacing)
    {
        return lineSpacing.IsExplicit ? LineBaselineOffset(fontSize, lineSpacing) : fontSize * 0.9344d;
    }

    private static double BaselineOffset(double fontSize)
    {
        const double baselineOffsetFactor = 0.974d;
        return fontSize * baselineOffsetFactor;
    }

    private static double BaselineOffset(double fontSize, ResolvedRunTextStyle? style, TextAdvanceEstimator advanceEstimator)
    {
        if (style is null)
        {
            return BaselineOffset(fontSize);
        }

        ResolvedRunTextStyle runStyle = style.Value;
        OpenTypeFont? font = advanceEstimator.ResolveOpenTypeFont(runStyle.Typeface, runStyle.Bold, runStyle.Italic);
        if (font is null || font.UnitsPerEm == 0)
        {
            return BaselineOffset(fontSize);
        }

        double ascenderRatio = font.Os2.WindowsAscender / (double)font.UnitsPerEm;
        if (ascenderRatio <= 0d)
        {
            return BaselineOffset(fontSize);
        }

        double metricRatio = Math.Clamp(ascenderRatio, 0.75d, 1.05d);
        if (fontSize >= 24d)
        {
            metricRatio = Math.Max(0.974d, metricRatio);
        }

        return fontSize * metricRatio;
    }

    private static PptxTextBaselineMetricLayout ReadBaselineMetric(double fontSize, ResolvedRunTextStyle? style, TextAdvanceEstimator advanceEstimator)
    {
        const double fallbackRatio = 0.974d;
        if (style is null)
        {
            return new PptxTextBaselineMetricLayout("Fallback", null, false, false, fontSize, fallbackRatio, 0, 0, 0, 0, 0, 0);
        }

        ResolvedRunTextStyle runStyle = style.Value;
        OpenTypeFont? font = advanceEstimator.ResolveOpenTypeFont(runStyle.Typeface, runStyle.Bold, runStyle.Italic);
        if (font is null || font.UnitsPerEm == 0)
        {
            return new PptxTextBaselineMetricLayout("Fallback", runStyle.Typeface, runStyle.Bold, runStyle.Italic, fontSize, fallbackRatio, 0, 0, 0, 0, 0, 0);
        }

        double ascenderRatio = font.Os2.WindowsAscender / (double)font.UnitsPerEm;
        double ratio = ascenderRatio > 0d
            ? Math.Clamp(ascenderRatio, 0.75d, 1.05d)
            : fallbackRatio;
        if (fontSize >= 24d)
        {
            ratio = Math.Max(fallbackRatio, ratio);
        }
        string source = ascenderRatio > 0d ? "OS/2 usWinAscent" : "Fallback";
        return new PptxTextBaselineMetricLayout(
            source,
            runStyle.Typeface,
            runStyle.Bold,
            runStyle.Italic,
            fontSize,
            ratio,
            font.UnitsPerEm,
            font.Os2.WindowsAscender,
            font.Os2.WindowsDescender,
            font.Os2.TypographicAscender,
            font.Os2.TypographicDescender,
            font.Os2.TypographicLineGap);
    }

    private static string? ReadBulletText(XElement? paragraphProperties, ref int autoNumberValue)
    {
        if (paragraphProperties is null || paragraphProperties.Element(DrawingNamespace + "buNone") is not null)
        {
            return null;
        }

        if ((string?)paragraphProperties.Element(DrawingNamespace + "buChar")?.Attribute("char") is { } bullet)
        {
            return bullet;
        }

        XElement? autoNumber = paragraphProperties.Element(DrawingNamespace + "buAutoNum");
        if (autoNumber is null)
        {
            return null;
        }

        if (autoNumber.Attribute("startAt") is { } startAt &&
            int.TryParse(startAt.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int start) &&
            start > 0)
        {
            autoNumberValue = start;
        }

        string result = FormatAutoNumber(autoNumberValue, (string?)autoNumber.Attribute("type"));
        autoNumberValue++;
        return result;
    }

    private static string FormatAutoNumber(int value, string? type)
    {
        return type switch
        {
            "arabicParenBoth" => $"({value})",
            "arabicParenR" => $"{value})",
            "alphaLcPeriod" => $"{FormatAlphaNumber(value, upper: false)}.",
            "alphaUcPeriod" => $"{FormatAlphaNumber(value, upper: true)}.",
            "alphaLcParenR" => $"{FormatAlphaNumber(value, upper: false)})",
            "alphaUcParenR" => $"{FormatAlphaNumber(value, upper: true)})",
            "romanLcPeriod" => $"{FormatRomanNumber(value, upper: false)}.",
            "romanUcPeriod" => $"{FormatRomanNumber(value, upper: true)}.",
            "romanLcParenR" => $"{FormatRomanNumber(value, upper: false)})",
            "romanUcParenR" => $"{FormatRomanNumber(value, upper: true)})",
            _ => $"{value}."
        };
    }

    private static string FormatAlphaNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        int current = value;
        while (current > 0)
        {
            current--;
            char letter = (char)((upper ? 'A' : 'a') + current % 26);
            builder.Insert(0, letter);
            current /= 26;
        }

        return builder.ToString();
    }

    private static string FormatRomanNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return string.Empty;
        }

        (int Value, string Numeral)[] numerals =
        [
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I")
        ];

        var builder = new StringBuilder();
        int current = value;
        foreach ((int numeralValue, string numeral) in numerals)
        {
            while (current >= numeralValue)
            {
                builder.Append(numeral);
                current -= numeralValue;
            }
        }

        string result = builder.ToString();
        return upper ? result : result.ToLowerInvariant();
    }

    private static BulletStyle ReadBulletStyle(XElement? paragraphProperties, PptxTheme theme, double textFontSize, RgbColor textColor, string? textTypeface)
    {
        XElement? bulletFont = FindBulletProperty(paragraphProperties, "buFont");
        XElement? bulletColor = FindBulletProperty(paragraphProperties, "buClr");
        XElement? bulletSizePercent = FindBulletProperty(paragraphProperties, "buSzPct");
        XElement? bulletSizePoints = FindBulletProperty(paragraphProperties, "buSzPts");

        string? typeface = theme.ResolveTypeface((string?)bulletFont?.Attribute("typeface"));
        RgbColor color = bulletColor is not null &&
            TryReadSolidColor(bulletColor, theme, out RgbColor explicitColor)
                ? explicitColor
                : textColor;
        double fontSize = textFontSize;
        if (bulletSizePercent?.Attribute("val") is { } sizePercent)
        {
            fontSize = textFontSize * Math.Max(0.1d, int.Parse(sizePercent.Value, CultureInfo.InvariantCulture) / 100000d);
        }
        else if (bulletSizePoints?.Attribute("val") is { } sizePoints)
        {
            fontSize = Math.Max(0.1d, int.Parse(sizePoints.Value, CultureInfo.InvariantCulture) / 100d);
        }

        return new BulletStyle(fontSize, color, typeface ?? textTypeface);
    }

    private static XElement? FindBulletProperty(XElement? paragraphProperties, string localName)
    {
        if (paragraphProperties is null)
        {
            return null;
        }

        XName propertyName = DrawingNamespace + localName;
        XElement? marker = paragraphProperties
            .Elements()
            .FirstOrDefault(element => element.Name == DrawingNamespace + "buChar" ||
                element.Name == DrawingNamespace + "buAutoNum" ||
                element.Name == DrawingNamespace + "buBlip");
        IEnumerable<XElement> candidates = marker is null
            ? paragraphProperties.Elements()
            : paragraphProperties.Elements().TakeWhile(element => element != marker);
        return candidates.FirstOrDefault(element => element.Name == propertyName);
    }

    private static bool TryReadHighlightColor(XElement? runProperties, out RgbColor color)
    {
        XElement? highlight = runProperties?.Element(DrawingNamespace + "highlight");
        string? hex = (string?)highlight?.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        return RgbColor.TryParse(hex, out color);
    }

    private static TextVerticalAnchor ReadVerticalAnchor(XElement textBody)
    {
        string? anchor = (string?)textBody
            .Element(DrawingNamespace + "bodyPr")
            ?.Attribute("anchor");
        return anchor switch
        {
            "ctr" => TextVerticalAnchor.Middle,
            "b" => TextVerticalAnchor.Bottom,
            _ => TextVerticalAnchor.Top
        };
    }

    private static double EstimateTextHeight(XElement textBody, XElement? defaultParagraphProperties)
    {
        double height = 0d;
        foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
        {
            XElement? paragraphProperties = paragraph.Element(DrawingNamespace + "pPr");
            XElement? defaultRunProperties = paragraphProperties?.Element(DrawingNamespace + "defRPr") ??
                defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
            LineSpacing lineSpacing = ReadLineSpacing(paragraphProperties, defaultParagraphProperties);
            double paragraphFontSize = ReadFirstParagraphFontSize(paragraph, defaultRunProperties);
            height += ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef", paragraphFontSize);
            if (!ParagraphHasVisibleContent(paragraph))
            {
                if (ParagraphHasLayoutContent(paragraph))
                {
                    XElement? endRunProperties = paragraph.Element(DrawingNamespace + "endParaRPr");
                    height += ReadParagraphAdvance(lineSpacing, ReadFontSize(endRunProperties, defaultRunProperties));
                    double emptyFontSize = ReadFontSize(endRunProperties, defaultRunProperties);
                    height += ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", emptyFontSize);
                }

                continue;
            }

            double maxFontSize = 18d;
            bool hasLineContent = false;
            foreach (XElement child in paragraph.Elements())
            {
                if (child.Name == DrawingNamespace + "br")
                {
                    height += ReadLineAdvance(lineSpacing, maxFontSize);
                    maxFontSize = 18d;
                    hasLineContent = false;
                    continue;
                }

                if (!IsTextRunElement(child))
                {
                    continue;
                }

                XElement? runProperties = child.Element(DrawingNamespace + "rPr");
                double fontSize = ReadFontSize(runProperties, defaultRunProperties);
                maxFontSize = Math.Max(maxFontSize, fontSize);
                hasLineContent = true;
            }

            if (hasLineContent || maxFontSize > 0d)
            {
                height += ReadLineAdvance(lineSpacing, maxFontSize);
            }

            height += ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", paragraphFontSize);
        }

        return height;
    }

}
