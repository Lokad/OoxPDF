using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxScene(PptxDocument Document, PptxTheme Theme, IReadOnlyList<PptxSceneSlide> Slides);

internal sealed record PptxSceneSnapshot(IReadOnlyList<PptxSceneSlideSnapshot> Slides);

internal sealed record PptxSceneSlideSnapshot(
    int Index,
    string PartName,
    string? MasterPartName,
    string? LayoutPartName,
    bool HasMasterXml,
    bool HasLayoutXml,
    bool HasSlideXml,
    int MasterRelationshipCount,
    int LayoutRelationshipCount,
    int SlideRelationshipCount,
    bool HasMasterBackground,
    bool HasLayoutBackground,
    bool HasSlideBackground,
    IReadOnlyDictionary<string, string> MasterColorMap,
    IReadOnlyDictionary<string, string> LayoutColorMap,
    IReadOnlyDictionary<string, string> SlideColorMap,
    IReadOnlyList<PptxSceneNodeSnapshot> MasterNodes,
    IReadOnlyList<PptxSceneNodeSnapshot> LayoutNodes,
    IReadOnlyList<PptxSceneNodeSnapshot> SlideNodes);

internal sealed record PptxSceneNodeSnapshot(
    string Kind,
    bool IsPlaceholder,
    bool IsSmartArtGraphicFrame,
    bool HasHyperlinkClick,
    string? HyperlinkClickId,
    string? HyperlinkClickAction,
    bool HasBounds,
    double RotationDegrees,
    bool FlipHorizontal,
    bool FlipVertical,
    bool HasShape,
    string ShapePreset,
    bool HasTextBody,
    int TextParagraphCount,
    int TextRunCount,
    bool HasPicture,
    bool HasPictureResource,
    string PictureContentType,
    double PictureAlpha,
    string PictureAlphaValue,
    bool HasShapePictureFillResource,
    string ShapePictureFillContentType,
    double ShapePictureFillAlpha,
    string ShapePictureFillAlphaValue,
    string PictureRecolorKind,
    string PictureRecolorKindValue,
    double? PictureRecolorBrightness,
    double? PictureRecolorContrast,
    double? PictureRecolorThreshold,
    string PictureRecolorBrightnessValue,
    string PictureRecolorContrastValue,
    string PictureRecolorThresholdValue,
    bool HasTable,
    int TableRowCount,
    int TableCellCount,
    string TableStyleId,
    string TableStyleName,
    string TableStyleAccent,
    bool TableStyleIsSupported,
    bool TableStyleFirstRow,
    string TableStyleFirstRowValue,
    bool TableStyleLastRow,
    string TableStyleLastRowValue,
    bool TableStyleFirstColumn,
    string TableStyleFirstColumnValue,
    bool TableStyleLastColumn,
    string TableStyleLastColumnValue,
    bool TableStyleBandRow,
    string TableStyleBandRowValue,
    bool TableStyleBandColumn,
    string TableStyleBandColumnValue,
    int TableStyleFillCellCount,
    int TableStyleTextColorCellCount,
    int TableStyleTextBoldCellCount,
    bool HasChart,
    string ChartRelationshipId,
    string ChartTargetPartName,
    int ChartPlotCount,
    int ChartAxisCount,
    int ChartSeriesCount,
    int ChartSeriesMarkerCount,
    int ChartSeriesPointStyleCount,
    int ChartSeriesPointExplosionCount,
    int ChartDataLabelsDefinedCount,
    int ChartDataLabelOverrideCount,
    int ChartDataLabelManualLayoutCount,
    bool HasChartColorStyle,
    string ChartColorStylePartName,
    string ChartColorStyleMethod,
    string ChartColorStyleId,
    int ChartColorStyleColorCount,
    int ChartColorStyleVariationCount,
    int ChartColorStyleDeclarationCount,
    int ChartColorStyleResolvedDeclarationCount,
    IReadOnlyList<string> ChartColorStyleDeclarationKinds,
    bool HasChartStylePart,
    string ChartStylePartName,
    string ChartStylePartId,
    int ChartStyleEntryCount,
    IReadOnlyList<string> ChartStyleEntryRoles,
    int ChartStyleShapeStyleCount,
    int ChartStyleShapeFillCount,
    int ChartStyleFillReferenceCount,
    int ChartStyleEffectReferenceCount,
    int ChartStyleFontReferenceCount,
    int ChartRejectedPointStyleIndexCount,
    IReadOnlyList<string> ChartRejectedPointStyleIndexValues,
    int ChartRejectedDataLabelOverrideIndexCount,
    IReadOnlyList<string> ChartRejectedDataLabelOverrideIndexValues,
    bool HasChartExternalData,
    string ChartExternalDataRelationshipId,
    string ChartExternalDataTargetPartName,
    bool? ChartExternalDataAutoUpdate,
    string ChartExternalDataAutoUpdateValue,
    bool HasChartExternalDataResource,
    string ChartExternalDataContentType,
    bool? ChartDate1904,
    string ChartDate1904Value,
    bool? ChartRoundedCorners,
    string ChartRoundedCornersValue,
    bool? ChartPlotVisibleOnly,
    string ChartPlotVisibleOnlyValue,
    bool? ChartShowDataLabelsOverMaximum,
    string ChartShowDataLabelsOverMaximumValue,
    string ChartDisplayBlanksAs,
    bool HasChartLegend,
    string ChartLegendPosition,
    bool? ChartLegendOverlay,
    string ChartLegendOverlayValue,
    bool? ChartLegendDeleted,
    string ChartLegendDeletedValue,
    bool HasChartLegendManualLayout,
    double? ChartLegendLayoutX,
    double? ChartLegendLayoutY,
    double? ChartLegendLayoutWidth,
    double? ChartLegendLayoutHeight,
    bool HasGroupTransform,
    IReadOnlyList<PptxSceneNodeSnapshot> Children);

internal sealed record PptxTextRunSnapshot(
    string Text,
    double X,
    double Y,
    double Width,
    double FontSize,
    double CharacterSpacing,
    RgbColor Color,
    double Alpha,
    RgbColor? Highlight,
    bool Bold,
    bool Italic,
    bool Underline,
    bool Strike,
    string Alignment,
    string? FontFamily);

internal sealed record PptxTextFrameModelSnapshot(
    double TextX,
    double TextWidth,
    double TextHeight,
    double VerticalOffset,
    double InsetLeft,
    double InsetRight,
    double InsetTop,
    double InsetBottom,
    string? InsetLeftValue,
    string? InsetRightValue,
    string? InsetTopValue,
    string? InsetBottomValue,
    double FontScale,
    string? FontScaleValue,
    string FontScaleSource,
    double LineSpacingScale,
    string? LineSpacingReductionValue,
    string LineSpacingScaleSource,
    bool CompatibleLineSpacing,
    string CompatibleLineSpacingSource,
    double? RotationDegrees,
    string? RotationValue,
    string RotationDegreesSource,
    int InheritedPlaceholderCount,
    bool HasInheritedTextBody,
    bool UsesInheritedShapeBounds,
    string InsetLeftSource,
    string InsetRightSource,
    string InsetTopSource,
    string InsetBottomSource,
    string Orientation,
    string? OrientationValue,
    string OrientationSource,
    string VerticalAnchor,
    string? VerticalAnchorValue,
    string VerticalAnchorSource,
    bool? AnchorCenter,
    string? AnchorCenterValue,
    string AnchorCenterSource,
    string WrapMode,
    string? WrapValue,
    string WrapSource,
    string VerticalOverflow,
    string? VerticalOverflowValue,
    string VerticalOverflowSource,
    int ColumnCount,
    double ColumnSpacing,
    string ColumnSource,
    string ColumnCountSource,
    string ColumnSpacingSource,
    string? ColumnCountValue,
    string? ColumnSpacingValue,
    string AutofitModeValue,
    string AutofitModeSource,
    IReadOnlyList<PptxTextParagraphModelSnapshot> Paragraphs);

internal sealed record PptxTextParagraphModelSnapshot(
    int Level,
    string CascadeLevelName,
    int ResolvedCascadeSourceCount,
    IReadOnlyList<string> CascadeLayerNames,
    IReadOnlyList<string> CascadeLayerKinds,
    int ResolvedStyleSourceCount,
    IReadOnlyList<string> ResolvedStyleLayerNames,
    IReadOnlyList<string> ResolvedStyleLayerKinds,
    string Alignment,
    string? AlignmentValue,
    double FontSize,
    double SpacingBefore,
    double SpacingAfter,
    double LineSpacingValue,
    string LineSpacingKind,
    bool LineSpacingUseNormalLineAdvance,
    double MarginLeft,
    double HangingIndent,
    IReadOnlyList<PptxTextRunModelSnapshot> Runs);

internal sealed record PptxTextRunModelSnapshot(
    string Kind,
    string Text,
    int ResolvedCascadeSourceCount,
    IReadOnlyList<string> CascadeLayerNames,
    IReadOnlyList<string> CascadeLayerKinds,
    double FontSize,
    double CharacterSpacing,
    string? Typeface,
    string TypefaceSource,
    string ColorSource,
    bool HasHyperlinkClick,
    string? HyperlinkClickId,
    bool Underline,
    string? UnderlineValue,
    bool Strike,
    string? StrikeValue,
    string? CapsValue,
    RgbColor? Highlight);

internal sealed record PptxTextFlowSnapshot(IReadOnlyList<PptxTextFlowFrameSnapshot> Frames);

internal sealed record PptxTextFlowFrameSnapshot(
    double TextX,
    double TextWidth,
    double TextHeight,
    double ClipY,
    double ClipHeight,
    double CursorTop,
    IReadOnlyList<PptxTextFlowParagraphSnapshot> Paragraphs);

internal sealed record PptxTextFlowParagraphSnapshot(
    int Level,
    string Alignment,
    double FontSize,
    IReadOnlyList<PptxTextFlowRunSnapshot> Runs);

internal sealed record PptxTextFlowRunSnapshot(
    string SourceKind,
    string SourceText,
    double FontSize,
    string? Typeface,
    bool HasHyperlinkClick,
    string? HyperlinkClickId,
    IReadOnlyList<PptxTextFlowSegmentSnapshot> Segments);

internal sealed record PptxTextFlowSegmentSnapshot(
    string Kind,
    string Text,
    string AdvanceText,
    bool Draw,
    bool PreventCoalesce,
    double FontScale);

internal sealed record PptxTextLayoutSnapshot(IReadOnlyList<PptxTextFrameLayoutSnapshot> Frames);

internal sealed record PptxTextFrameLayoutSnapshot(IReadOnlyList<PptxTextParagraphLayoutSnapshot> Paragraphs);

internal sealed record PptxTextParagraphLayoutSnapshot(
    int Level,
    IReadOnlyList<PptxTextLineLayoutSnapshot> Lines);

internal sealed record PptxTextLineLayoutSnapshot(
    double TopY,
    double BaselineY,
    double Advance,
    double BaselineOffset,
    double MaxFontSize,
    string LineSpacingKind,
    PptxTextBaselineMetricSnapshot BaselineMetric,
    double StartX,
    double EndX,
    double NaturalEndX,
    string Alignment,
    IReadOnlyList<PptxTextSpanLayoutSnapshot> Spans);

internal sealed record PptxTextBaselineMetricSnapshot(
    string Source,
    string? Typeface,
    bool Bold,
    bool Italic,
    double FontSize,
    double Ratio,
    int UnitsPerEm,
    int WindowsAscender,
    int WindowsDescender,
    int TypographicAscender,
    int TypographicDescender,
    int TypographicLineGap);

internal sealed record PptxTextSpanLayoutSnapshot(
    string? SourceText,
    string Text,
    double X,
    double Y,
    double Width,
    double FontSize,
    IReadOnlyList<PptxTextAtomLayoutSnapshot> Atoms,
    PptxTextGlyphSpanLayoutSnapshot GlyphSpan);

internal sealed record PptxTextGlyphSpanLayoutSnapshot(
    string Text,
    string? Typeface,
    double FontSize,
    double LeadingAdjustment,
    double NaturalWidth,
    double LayoutWidth,
    int GlyphCount,
    double FirstAdjustmentAfterOrigin,
    IReadOnlyList<PptxTextGlyphLayoutSnapshot> Glyphs);

internal sealed record PptxTextGlyphLayoutSnapshot(
    int CodePoint,
    string? Typeface,
    ushort GlyphId,
    double Advance,
    double AdjustmentBefore);

internal sealed record PptxTextAtomLayoutSnapshot(
    string Kind,
    string Text,
    double X,
    double Width,
    bool Draw);

internal sealed record PptxTextGlyphRunSnapshot(
    string Text,
    double X,
    double BaselineY,
    double Width,
    RgbColor? HighlightColor,
    double? HighlightX,
    double? HighlightY,
    double? HighlightWidth,
    double? HighlightHeight,
    double? UnderlineX,
    double? UnderlineY,
    double? UnderlineWidth,
    double? UnderlineHeight,
    double? StrikeX,
    double? StrikeY,
    double? StrikeWidth,
    double? StrikeHeight,
    int FrameIndex,
    int ParagraphIndex,
    int LineIndex,
    int SpanIndex,
    int LineSpanCount,
    double FrameFontScale,
    double FrameShapeX,
    double FrameShapeTopY,
    double FrameShapeWidth,
    double FrameShapeHeight,
    double FrameInsetLeft,
    double FrameInsetRight,
    double FrameInsetTop,
    double FrameInsetBottom,
    string FrameWrapMode,
    string? FrameWrapValue,
    string FrameAutofitMode,
    double FrameTextX,
    double FrameTextWidth,
    double FrameTextWrapWidth,
    double FrameTextHeight,
    double FrameClipX,
    double FrameClipWidth,
    double FrameClipY,
    double FrameClipHeight,
    int FrameColumnCount,
    double FrameColumnSpacing,
    double LineTopY,
    double LineAdvance,
    double LineMaxFontSize,
    double LayoutFontSize,
    double PdfFontSize,
    int GlyphCount,
    double FirstAdjustmentAfterOrigin,
    IReadOnlyList<PptxTextGlyphRunAtomSnapshot> Glyphs);

internal sealed record PptxTextGlyphRunAtomSnapshot(
    int CodePoint,
    string? Typeface,
    string ResourceName,
    ushort GlyphId,
    double Advance,
    double AdjustmentBefore);

internal sealed record PptxSceneSlide(
    int Index,
    string PartName,
    string? MasterPartName,
    string? LayoutPartName,
    XDocument? MasterXml,
    XDocument? LayoutXml,
    XDocument SlideXml,
    IReadOnlyDictionary<string, OoxRelationship> MasterRelationships,
    IReadOnlyDictionary<string, OoxRelationship> LayoutRelationships,
    IReadOnlyDictionary<string, OoxRelationship> SlideRelationships,
    PptxColorMap MasterColorMap,
    PptxColorMap LayoutColorMap,
    PptxColorMap SlideColorMap,
    PptxSceneBackground MasterBackground,
    PptxSceneBackground LayoutBackground,
    PptxSceneBackground SlideBackground,
    IReadOnlyList<PptxSceneNode> MasterNodes,
    IReadOnlyList<PptxSceneNode> LayoutNodes,
    IReadOnlyList<PptxSceneNode> SlideNodes);

internal readonly record struct PptxSceneBackground(
    bool HasFill,
    RgbColor Color,
    double Alpha);

internal sealed record PptxSceneNode(
    PptxSceneNodeKind Kind,
    string Id,
    string Name,
    bool IsPlaceholder,
    bool IsSmartArtGraphicFrame,
    PptxSceneHyperlinkClick HyperlinkClick,
    PptxSceneBounds? Bounds,
    PptxSceneShape? Shape,
    PptxSceneTextBody? TextBody,
    PptxScenePicture? Picture,
    PptxSceneTable? Table,
    PptxSceneChart? Chart,
    PptxSceneGroupTransform GroupTransform,
    IReadOnlyList<PptxSceneNode> Children,
    XElement Source);

internal readonly record struct PptxSceneHyperlinkClick(
    bool IsDefined,
    string? RelationshipId,
    string? Action);

internal sealed record PptxSceneShape(
    string Preset,
    IReadOnlyDictionary<string, double> PresetAdjustments,
    bool HasCustomGeometry,
    PptxSceneCustomGeometry CustomGeometry,
    PptxSceneFillStyle Fill,
    PptxSceneGradientFill GradientFill,
    PptxScenePatternFill PatternFill,
    PptxSceneShapePictureFill PictureFill,
    PptxSceneGlow Glow,
    PptxSceneOuterShadow OuterShadow,
    PptxSceneLineStyle Line,
    PptxSceneLineEnd HeadEnd,
    PptxSceneLineEnd TailEnd);

internal sealed record PptxSceneCustomGeometry(
    bool HasGeometry,
    IReadOnlyList<PptxSceneCustomGuide> Guides,
    IReadOnlyList<PptxSceneCustomPath> Paths);

internal sealed record PptxSceneCustomGuide(
    string Name,
    string Formula);

internal sealed record PptxSceneCustomPath(
    double Width,
    double Height,
    bool AllowsFill,
    bool AllowsStroke,
    IReadOnlyList<PptxSceneCustomCommand> Commands);

internal sealed record PptxSceneCustomCommand(
    PptxSceneCustomCommandKind Kind,
    IReadOnlyList<PptxSceneCustomPoint> Points,
    string RadiusX,
    string RadiusY,
    string StartAngle,
    string SweepAngle);

internal enum PptxSceneCustomCommandKind
{
    MoveTo,
    LineTo,
    CubicBezierTo,
    QuadraticBezierTo,
    ArcTo,
    Close
}

internal readonly record struct PptxSceneCustomPoint(
    string X,
    string Y);

internal readonly record struct PptxSceneFillStyle(
    bool HasFill,
    RgbColor Color,
    double Alpha);

internal sealed record PptxSceneGradientFill(
    bool HasGradient,
    double AngleDegrees,
    IReadOnlyList<PptxSceneGradientStop> Stops);

internal readonly record struct PptxSceneGradientStop(
    double Offset,
    RgbColor Color,
    double Alpha);

internal readonly record struct PptxScenePatternFill(
    bool HasPattern,
    string Preset,
    RgbColor Foreground,
    RgbColor Background,
    double Alpha);

internal readonly record struct PptxSceneShapePictureFill(
    bool HasPicture,
    string RelationshipId,
    string? TargetPartName,
    PptxSceneImageResource? Resource,
    PptxSceneRect Crop,
    PptxSceneRect Fill,
    double Alpha,
    string? AlphaValue,
    PptxScenePictureTile Tile);

internal readonly record struct PptxSceneGlow(
    bool HasGlow,
    RgbColor Color,
    double Alpha,
    double Radius);

internal readonly record struct PptxSceneOuterShadow(
    bool HasShadow,
    RgbColor Color,
    double Alpha,
    double OffsetX,
    double OffsetY);

internal readonly record struct PptxSceneLineStyle(
    bool HasLine,
    RgbColor Color,
    double Width,
    double Alpha,
    IReadOnlyList<double> DashPattern,
    string? DashPreset,
    PptxSceneLineCompound? Compound,
    string? CompoundValue,
    int? Cap,
    string? CapValue,
    int? Join,
    string? JoinValue,
    bool WidthSpecified = true)
{
    public bool HasDash => DashPattern is { Count: > 0 };
}

internal enum PptxSceneLineCompound
{
    Single,
    Double,
    ThickThin,
    ThinThick,
    Triple
}

internal enum PptxSceneLineEndKind
{
    None,
    Triangle,
    Arrow,
    Stealth,
    Diamond,
    Oval
}

internal readonly record struct PptxSceneLineEnd(
    PptxSceneLineEndKind Kind,
    string? TypeValue,
    double WidthScale,
    string? WidthValue,
    double LengthScale,
    string? LengthValue)
{
    public bool IsNone => Kind == PptxSceneLineEndKind.None;
}

internal sealed record PptxScenePicture(
    string? RelationshipId,
    string? TargetPartName,
    PptxSceneImageResource? Resource,
    PptxSceneRect Crop,
    PptxSceneRect Fill,
    double Alpha,
    string? AlphaValue,
    PptxSceneImageRecolor Recolor,
    PptxScenePictureTile Tile);

internal sealed record PptxSceneImageResource(
    string PartName,
    string ContentType,
    byte[] Bytes);

internal sealed record PptxScenePackageResource(
    string PartName,
    string ContentType,
    byte[] Bytes);

internal sealed record PptxSceneChart(
    string? RelationshipId,
    string? TargetPartName,
    XDocument? ChartXml,
    PptxSceneChartExternalData ExternalData,
    PptxSceneChartOptions Options,
    IReadOnlyList<RgbColor>? PaletteColors,
    PptxSceneChartColorStyle ColorStyle,
    PptxSceneChartStyle StylePart,
    string StyleId,
    IReadOnlyList<PptxSceneChartPlot> Plots,
    IReadOnlyList<PptxSceneChartAxis> Axes,
    PptxSceneChartTitle Title,
    PptxSceneChartLegend Legend,
    PptxSceneChartTextStyleOverride TextStyle,
    PptxSceneChartManualLayout PlotAreaLayout,
    PptxSceneChartShapeStyle ChartAreaStyle,
    PptxSceneChartShapeStyle PlotAreaStyle);

internal readonly record struct PptxSceneChartExternalData(
    bool IsDefined,
    string? RelationshipId,
    string? TargetPartName,
    PptxScenePackageResource? Resource,
    bool? AutoUpdate,
    string AutoUpdateValue);

internal readonly record struct PptxSceneChartOptions(
    bool? Date1904,
    string Date1904Value,
    bool? RoundedCorners,
    string RoundedCornersValue,
    bool? PlotVisibleOnly,
    string PlotVisibleOnlyValue,
    bool? ShowDataLabelsOverMaximum,
    string ShowDataLabelsOverMaximumValue,
    PptxSceneChartDisplayBlanksAs DisplayBlanksAsKind,
    string DisplayBlanksAs);

internal enum PptxSceneChartDisplayBlanksAs
{
    Unknown,
    Gap,
    Span,
    Zero
}

internal readonly record struct PptxSceneChartColorStyle(
    bool IsDefined,
    string? PartName,
    string Method,
    string Id,
    IReadOnlyList<RgbColor> Colors,
    int VariationCount,
    IReadOnlyList<PptxSceneChartColorDeclaration> Declarations,
    XDocument? ColorStyleXml);

internal readonly record struct PptxSceneChartColorDeclaration(
    string Kind,
    string Value,
    int? VariationIndex,
    bool IsResolved,
    RgbColor? Color,
    double Alpha);

internal sealed record PptxSceneChartStyle(
    bool IsDefined,
    string? PartName,
    string Id,
    XDocument? StyleXml,
    IReadOnlyList<PptxSceneChartStyleEntry> Entries);

internal readonly record struct PptxSceneChartStyleEntry(
    string Role,
    int? LineReferenceIndex,
    int? FillReferenceIndex,
    int? EffectReferenceIndex,
    string FontReferenceIndex,
    PptxSceneLineStyle Line,
    PptxSceneChartShapeStyle ShapeStyle,
    PptxSceneLineStyle ShapeLine,
    PptxSceneChartTextStyleOverride TextStyle);

internal sealed record PptxSceneChartPlot(
    PptxSceneChartPlotKind PlotKind,
    string Kind,
    int PlotAreaIndex,
    int KindIndex,
    int SeriesCount,
    IReadOnlyList<string> AxisIds,
    IReadOnlyList<PptxSceneChartSeries> Series,
    PptxSceneChartGrouping GroupingKind,
    string Grouping,
    PptxSceneChartBarDirection BarDirectionKind,
    string BarDirection,
    PptxSceneChartScatterStyle ScatterStyleKind,
    string ScatterStyle,
    PptxSceneChartRadarStyle RadarStyleKind,
    string RadarStyle,
    bool? MarkersEnabled,
    string MarkersEnabledValue,
    bool? VaryColors,
    string VaryColorsValue,
    double? GapWidth,
    string GapWidthValue,
    double? Overlap,
    string OverlapValue,
    double? HoleSize,
    string HoleSizeValue,
    double? FirstSliceAngle,
    string FirstSliceAngleValue,
    PptxSceneChartDataLabels DataLabels,
    XElement Source);

internal enum PptxSceneChartGrouping
{
    Clustered,
    PercentStacked,
    Stacked,
    Standard,
    Unknown
}

internal enum PptxSceneChartBarDirection
{
    Bar,
    Column,
    Unknown
}

internal enum PptxSceneChartScatterStyle
{
    Line,
    LineMarker,
    Marker,
    None,
    Smooth,
    SmoothMarker,
    Unknown
}

internal enum PptxSceneChartRadarStyle
{
    Filled,
    Marker,
    Standard,
    Unknown
}

internal enum PptxSceneChartPlotKind
{
    Area,
    Bar,
    Bubble,
    Doughnut,
    Line,
    Pie,
    Radar,
    Scatter,
    Unknown
}

internal sealed record PptxSceneChartDataLabels(
    bool? ShowValue,
    string ShowValueValue,
    bool? ShowPercent,
    string ShowPercentValue,
    bool? ShowCategoryName,
    string ShowCategoryNameValue,
    bool? ShowSeriesName,
    string ShowSeriesNameValue,
    bool? ShowLeaderLines,
    string ShowLeaderLinesValue,
    bool? ShowLegendKey,
    string ShowLegendKeyValue,
    bool? ShowBubbleSize,
    string ShowBubbleSizeValue,
    PptxSceneChartLeaderLines LeaderLines,
    PptxSceneChartDataLabelPosition PositionKind,
    string Position,
    string Separator,
    string NumberFormat,
    PptxSceneChartNumberFormat NumberFormatInfo,
    PptxSceneChartManualLayout Layout,
    PptxSceneChartTextStyleOverride TextStyle,
    PptxSceneChartTextBodyProperties TextBodyProperties,
    PptxSceneChartShapeStyle ShapeStyle,
    IReadOnlyList<string> RejectedOverrideIndexValues,
    IReadOnlyList<PptxSceneChartDataLabelOverride> Overrides,
    bool IsDefined);

internal sealed record PptxSceneChartDataLabelOverride(
    int Index,
    string IndexValue,
    bool? ShowValue,
    string ShowValueValue,
    bool? ShowPercent,
    string ShowPercentValue,
    bool? ShowCategoryName,
    string ShowCategoryNameValue,
    bool? ShowSeriesName,
    string ShowSeriesNameValue,
    bool? ShowLeaderLines,
    string ShowLeaderLinesValue,
    bool? ShowLegendKey,
    string ShowLegendKeyValue,
    bool? ShowBubbleSize,
    string ShowBubbleSizeValue,
    PptxSceneChartLeaderLines LeaderLines,
    string CustomText,
    IReadOnlyList<PptxSceneChartTextRun> CustomTextRuns,
    PptxSceneChartDataLabelPosition PositionKind,
    string Position,
    string Separator,
    string NumberFormat,
    PptxSceneChartNumberFormat NumberFormatInfo,
    PptxSceneChartManualLayout Layout,
    PptxSceneChartTextStyleOverride TextStyle,
    PptxSceneChartTextBodyProperties TextBodyProperties,
    PptxSceneChartShapeStyle ShapeStyle);

internal readonly record struct PptxSceneChartTextBodyProperties(
    double? RotationDegrees,
    string RotationValue);

internal readonly record struct PptxSceneChartNumberFormat(
    bool IsDefined,
    string FormatCode,
    bool? SourceLinked,
    string SourceLinkedValue);

internal readonly record struct PptxSceneChartLeaderLines(
    bool IsDefined,
    PptxSceneLineStyle Line);

internal sealed record PptxSceneChartTextRun(
    string Text,
    PptxSceneChartTextStyleOverride TextStyle);

internal enum PptxSceneChartDataLabelPosition
{
    BestFit,
    Bottom,
    Center,
    InsideBase,
    InsideEnd,
    Left,
    OutsideEnd,
    Right,
    Top,
    Unknown
}

internal sealed record PptxSceneChartShapeStyle(
    bool NoFill,
    PptxSceneFillStyle Fill,
    PptxSceneGradientFill? GradientFill,
    PptxScenePatternFill PatternFill,
    PptxSceneShapePictureFill PictureFill,
    PptxSceneLineStyle Line,
    PptxSceneGlow Glow,
    PptxSceneOuterShadow OuterShadow,
    PptxSceneChartEffectFamily Effects);

internal readonly record struct PptxSceneChartEffectFamily(
    bool HasEffectList,
    bool HasEffectDag,
    IReadOnlyList<string> UnsupportedEffectNames);

internal readonly record struct PptxSceneChartTextStyleOverride(
    string? FontFamily,
    double? FontSize,
    RgbColor? Color,
    double? Alpha,
    bool? Bold,
    bool? Italic,
    bool? Underline,
    bool? Strike);

internal readonly record struct PptxSceneChartManualLayout(
    bool HasLayout,
    double? X,
    string XValue,
    double? Y,
    string YValue,
    double? Width,
    string WidthValue,
    double? Height,
    string HeightValue,
    PptxSceneChartManualLayoutTarget LayoutTargetKind,
    string LayoutTarget,
    PptxSceneChartManualLayoutMode XModeKind,
    string XMode,
    PptxSceneChartManualLayoutMode YModeKind,
    string YMode,
    PptxSceneChartManualLayoutMode WidthModeKind,
    string WidthMode,
    PptxSceneChartManualLayoutMode HeightModeKind,
    string HeightMode);

internal enum PptxSceneChartManualLayoutTarget
{
    Inner,
    Outer,
    Unknown
}

internal enum PptxSceneChartManualLayoutMode
{
    Edge,
    Factor,
    Unknown
}

internal sealed record PptxSceneChartSeries(
    int? Index,
    string IndexValue,
    int? Order,
    string OrderValue,
    string? Name,
    PptxSceneChartSeriesDataSources DataSources,
    IReadOnlyList<double> Values,
    IReadOnlyList<PptxSceneChartNumberPoint> ValuePoints,
    int? ValuePointCount,
    string ValuePointCountValue,
    string? ValueFormatCode,
    IReadOnlyList<string> Categories,
    IReadOnlyList<PptxSceneChartStringPoint> CategoryPoints,
    int? CategoryPointCount,
    string CategoryPointCountValue,
    IReadOnlyList<IReadOnlyList<PptxSceneChartStringPoint>> CategoryLevels,
    IReadOnlyList<double> XValues,
    IReadOnlyList<PptxSceneChartNumberPoint> XValuePoints,
    int? XValuePointCount,
    string XValuePointCountValue,
    string? XValueFormatCode,
    IReadOnlyList<double> YValues,
    IReadOnlyList<PptxSceneChartNumberPoint> YValuePoints,
    int? YValuePointCount,
    string YValuePointCountValue,
    string? YValueFormatCode,
    IReadOnlyList<double> BubbleSizes,
    IReadOnlyList<PptxSceneChartNumberPoint> BubbleSizePoints,
    int? BubbleSizePointCount,
    string BubbleSizePointCountValue,
    string? BubbleSizeFormatCode,
    PptxSceneFillStyle Fill,
    PptxScenePatternFill PatternFill,
    PptxSceneLineStyle Line,
    PptxSceneChartMarker Marker,
    IReadOnlyList<PptxSceneChartPointStyle> PointStyles,
    IReadOnlyList<string> RejectedPointStyleIndexValues,
    double? Explosion,
    string ExplosionValue,
    bool? Smooth,
    string SmoothValue,
    PptxSceneChartDataLabels DataLabels);

internal readonly record struct PptxSceneChartNumberPoint(
    int Index,
    string IndexValue,
    bool HasParsedIndex,
    double? Value,
    string Text,
    bool HasValueElement);

internal readonly record struct PptxSceneChartStringPoint(
    int Index,
    string IndexValue,
    bool HasParsedIndex,
    string Text,
    bool HasText);

internal sealed record PptxSceneChartSeriesDataSources(
    PptxSceneChartDataSource Name,
    PptxSceneChartDataSource Values,
    PptxSceneChartDataSource Categories,
    PptxSceneChartDataSource XValues,
    PptxSceneChartDataSource YValues,
    PptxSceneChartDataSource BubbleSizes);

internal readonly record struct PptxSceneChartDataSource(
    string? Formula,
    PptxSceneChartDataSourceReferenceKind ReferenceKindValue,
    string ReferenceKind,
    PptxSceneChartDataSourceCacheKind CacheKindValue,
    string CacheKind,
    bool HasCachedPoints);

internal enum PptxSceneChartDataSourceReferenceKind
{
    Unknown,
    StringReference,
    NumberReference,
    MultiLevelStringReference
}

internal enum PptxSceneChartDataSourceCacheKind
{
    Unknown,
    StringCache,
    NumberCache,
    MultiLevelStringCache
}

internal sealed record PptxSceneChartMarker(
    bool IsDefined,
    PptxSceneChartMarkerSymbol SymbolKind,
    string Symbol,
    string? SizeValue,
    double Size,
    PptxSceneFillStyle Fill,
    PptxSceneLineStyle Line);

internal enum PptxSceneChartMarkerSymbol
{
    Circle,
    Dash,
    Diamond,
    Dot,
    None,
    Plus,
    Square,
    Star,
    Triangle,
    X,
    Unknown
}

internal sealed record PptxSceneChartPointStyle(
    int Index,
    string IndexValue,
    PptxSceneFillStyle Fill,
    PptxScenePatternFill PatternFill,
    PptxSceneLineStyle Line,
    double? Explosion,
    string ExplosionValue);

internal sealed record PptxSceneChartAxis(
    string Id,
    PptxSceneChartAxisKind AxisKind,
    string Kind,
    PptxSceneChartAxisPosition PositionKind,
    string Position,
    string CrossAxisId,
    PptxSceneChartAxisCrosses CrossesKind,
    string Crosses,
    double? CrossesAt,
    string CrossesAtValue,
    PptxSceneChartAxisCrossBetween CrossBetweenKind,
    string CrossBetween,
    PptxSceneChartAxisOrientation OrientationKind,
    string Orientation,
    bool IsReversed,
    bool? IsDeleted,
    string IsDeletedValue,
    bool HasScaling,
    double? Minimum,
    string MinimumValue,
    double? Maximum,
    string MaximumValue,
    double? MajorUnit,
    string MajorUnitValue,
    double? MinorUnit,
    string MinorUnitValue,
    bool HasMajorGridlines,
    bool HasMinorGridlines,
    bool HasMajorGridlineElement,
    bool HasMinorGridlineElement,
    PptxSceneLineStyle Line,
    PptxSceneLineStyle MajorGridlineLine,
    PptxSceneLineStyle MinorGridlineLine,
    PptxSceneLineStyle MajorGridlineStyleLine,
    PptxSceneLineStyle MinorGridlineStyleLine,
    PptxSceneChartTextStyleOverride TextStyle,
    PptxSceneChartTickLabelPosition TickLabelPositionKind,
    string TickLabelPosition,
    PptxSceneChartAxisTickMark MajorTickMarkKind,
    string MajorTickMark,
    PptxSceneChartAxisTickMark MinorTickMarkKind,
    string MinorTickMark,
    int? LabelOffset,
    string LabelOffsetValue,
    int? TickLabelSkip,
    string TickLabelSkipValue,
    int? TickMarkSkip,
    string TickMarkSkipValue,
    bool? NoMultiLevelLabels,
    string NoMultiLevelLabelsValue,
    string? NumberFormat,
    PptxSceneChartNumberFormat NumberFormatInfo,
    PptxSceneChartTitle Title);

internal enum PptxSceneChartAxisKind
{
    Category,
    Date,
    Series,
    Value,
    Unknown
}

internal enum PptxSceneChartAxisCrosses
{
    AutoZero,
    Maximum,
    Minimum,
    Unknown
}

internal enum PptxSceneChartAxisCrossBetween
{
    Between,
    MidpointCategory,
    Unknown
}

internal enum PptxSceneChartAxisOrientation
{
    MinimumMaximum,
    MaximumMinimum,
    Unknown
}

internal enum PptxSceneChartAxisTickMark
{
    Cross,
    Inside,
    None,
    Outside,
    Unknown
}

internal enum PptxSceneChartTickLabelPosition
{
    High,
    Low,
    NextTo,
    None,
    Unknown
}

internal enum PptxSceneChartAxisPosition
{
    Bottom,
    Left,
    Right,
    Top,
    Unknown
}

internal sealed record PptxSceneChartTitle(
    string? Text,
    IReadOnlyList<PptxSceneChartTextRun> TextRuns,
    bool? IsAutoDeleted,
    string IsAutoDeletedValue,
    bool IsAutoGenerated,
    bool? Overlay,
    string OverlayValue,
    PptxSceneChartManualLayout Layout,
    PptxSceneChartShapeStyle ShapeStyle,
    PptxSceneChartTextBodyProperties TextBodyProperties,
    PptxSceneChartTextStyleOverride TextStyle);

internal sealed record PptxSceneChartLegend(
    PptxSceneChartLegendPosition PositionKind,
    string Position,
    bool? Overlay,
    string OverlayValue,
    bool IsDefined,
    bool? IsDeleted,
    string IsDeletedValue,
    PptxSceneChartManualLayout Layout,
    PptxSceneChartShapeStyle ShapeStyle,
    PptxSceneChartTextBodyProperties TextBodyProperties,
    PptxSceneChartTextStyleOverride TextStyle);

internal enum PptxSceneChartLegendPosition
{
    Bottom,
    Left,
    Right,
    Top,
    TopRight,
    Unknown
}

internal sealed record PptxSceneTable(
    IReadOnlyList<double> ColumnWidths,
    IReadOnlyList<double> RowHeights,
    IReadOnlyList<PptxSceneTableRow> Rows,
    PptxSceneTableStyle Style,
    XElement? Source);

internal readonly record struct PptxSceneTableStyle(
    string? StyleId,
    string Name,
    string Accent,
    bool IsSupported,
    bool FirstRow,
    string FirstRowValue,
    bool LastRow,
    string LastRowValue,
    bool FirstColumn,
    string FirstColumnValue,
    bool LastColumn,
    string LastColumnValue,
    bool BandRow,
    string BandRowValue,
    bool BandColumn,
    string BandColumnValue)
{
    public bool HasStyle => StyleId is not null;
}

internal readonly record struct PptxBuiltInTableStyle(string Name, string Accent);

internal static class PptxBuiltInTableStyles
{
    public static bool TryGet(string? styleId, out PptxBuiltInTableStyle style)
    {
        return Styles.TryGetValue(styleId ?? string.Empty, out style);
    }

    private static IReadOnlyDictionary<string, PptxBuiltInTableStyle> Styles { get; } =
        new Dictionary<string, PptxBuiltInTableStyle>(StringComparer.OrdinalIgnoreCase)
        {
            ["{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}"] = new("Light-Style-1", "tx1"),
            ["{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"] = new("Light-Style-1", "accent1"),
            ["{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}"] = new("Light-Style-1", "accent2"),
            ["{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}"] = new("Light-Style-1", "accent3"),
            ["{D27102A9-8310-4765-A935-A1911B00CA55}"] = new("Light-Style-1", "accent4"),
            ["{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}"] = new("Light-Style-1", "accent5"),
            ["{68D230F3-CF80-4859-8CE7-A43EE81993B5}"] = new("Light-Style-1", "accent6"),
            ["{E8034E78-7F5D-4C2E-B375-FC64B27BC917}"] = new("Dark-Style-1", "dk1"),
            ["{125E5076-3810-47DD-B79F-674D7AD40C01}"] = new("Dark-Style-1", "accent1"),
            ["{37CE84F3-28C3-443E-9E96-99CF82512B78}"] = new("Dark-Style-1", "accent2"),
            ["{D03447BB-5D67-496B-8E87-E561075AD55C}"] = new("Dark-Style-1", "accent3"),
            ["{E929F9F4-4A8F-4326-A1B4-22849713DDAB}"] = new("Dark-Style-1", "accent4"),
            ["{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}"] = new("Dark-Style-1", "accent5"),
            ["{AF606853-7671-496A-8E4F-DF71F8EC918B}"] = new("Dark-Style-1", "accent6"),
            ["{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"] = new("Medium-Style-2", "tx1"),
            ["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"] = new("Medium-Style-2", "accent1"),
            ["{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}"] = new("Medium-Style-2", "accent2"),
            ["{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}"] = new("Medium-Style-2", "accent3"),
            ["{00A15C55-8517-42AA-B614-E9B94910E393}"] = new("Medium-Style-2", "accent4"),
            ["{7DF18680-E054-41AD-8BC1-D1AEF772440D}"] = new("Medium-Style-2", "accent5"),
            ["{93296810-A885-4BE3-A3E7-6D5BEEA58F35}"] = new("Medium-Style-2", "accent6")
        };
}

internal sealed record PptxSceneTableRow(IReadOnlyList<PptxSceneTableCell> Cells);

internal readonly record struct PptxSceneTableCell(
    int ColumnSpan,
    int RowSpan,
    bool IsMergedContinuation,
    PptxSceneTextInsets TextInsets,
    PptxSceneTableCellVerticalAnchor VerticalAnchor,
    string? VerticalAnchorValue,
    PptxSceneFillStyle Fill,
    PptxSceneTableCellBorders Borders,
    PptxSceneFillStyle StyleFill,
    PptxSceneTableCellTextStyle StyleText,
    XElement? TextBody);

internal readonly record struct PptxSceneTableCellTextStyle(RgbColor? Color, bool Bold);

internal readonly record struct PptxSceneTableCellBorders(
    PptxSceneTableCellBorder Left,
    PptxSceneTableCellBorder Right,
    PptxSceneTableCellBorder Top,
    PptxSceneTableCellBorder Bottom)
{
    public bool HasExplicitBorder => Left.IsSpecified || Right.IsSpecified || Top.IsSpecified || Bottom.IsSpecified;
}

internal readonly record struct PptxSceneTableCellBorder(bool IsSpecified, PptxSceneLineStyle Line);

internal readonly record struct PptxSceneTextInsets(double Left, double Right, double Top, double Bottom);

internal enum PptxSceneTableCellVerticalAnchor
{
    Top,
    Middle,
    Bottom
}

internal readonly record struct PptxSceneGroupTransform(
    long OffsetX,
    long OffsetY,
    long Width,
    long Height,
    long ChildOffsetX,
    long ChildOffsetY,
    double ScaleX,
    double ScaleY,
    double RotationDegrees,
    bool FlipHorizontal,
    bool FlipVertical)
{
    public static PptxSceneGroupTransform Identity { get; } = new(0, 0, 0, 0, 0, 0, 1d, 1d, 0d, FlipHorizontal: false, FlipVertical: false);
}

internal readonly record struct PptxSceneRect(
    double Left,
    double Top,
    double Right,
    double Bottom,
    string? LeftValue,
    string? TopValue,
    string? RightValue,
    string? BottomValue)
{
    public bool IsEmpty => Left == 0d && Top == 0d && Right == 0d && Bottom == 0d;
}

internal readonly record struct PptxScenePictureTile(
    bool HasTile,
    string? TileValue,
    string? AlignmentValue,
    string? FlipValue,
    string? ScaleXValue,
    string? ScaleYValue,
    string? OffsetXValue,
    string? OffsetYValue);

internal enum PptxSceneImageRecolorKind
{
    None,
    Luminance,
    Duotone,
    Grayscale,
    BiLevel
}

internal readonly record struct PptxSceneImageRecolor(
    PptxSceneImageRecolorKind Kind,
    double Brightness,
    double Contrast,
    RgbColor Dark,
    RgbColor Light,
    double Threshold,
    string? KindValue,
    string? BrightnessValue,
    string? ContrastValue,
    string? ThresholdValue)
{
    public static PptxSceneImageRecolor None { get; } = new(PptxSceneImageRecolorKind.None, 0d, 0d, default, default, 0d, null, null, null, null);

    public static PptxSceneImageRecolor Luminance(
        double brightness,
        double contrast,
        string? kindValue = "lum",
        string? brightnessValue = null,
        string? contrastValue = null)
    {
        return new PptxSceneImageRecolor(
            PptxSceneImageRecolorKind.Luminance,
            Math.Clamp(brightness, -1d, 1d),
            Math.Clamp(contrast, -1d, 1d),
            default,
            default,
            0d,
            kindValue,
            brightnessValue,
            contrastValue,
            null);
    }

    public static PptxSceneImageRecolor Duotone(RgbColor dark, RgbColor light, string? kindValue = "duotone")
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.Duotone, 0d, 0d, dark, light, 0d, kindValue, null, null, null);
    }

    public static PptxSceneImageRecolor Grayscale(string? kindValue = "grayscl")
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.Grayscale, 0d, 0d, default, default, 0d, kindValue, null, null, null);
    }

    public static PptxSceneImageRecolor BiLevel(double threshold, string? kindValue = "biLevel", string? thresholdValue = null)
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.BiLevel, 0d, 0d, default, default, Math.Clamp(threshold, 0d, 1d), kindValue, null, null, thresholdValue);
    }
}

internal sealed record PptxSceneTextBody(
    XElement? BodyProperties,
    XElement? ListStyle,
    IReadOnlyList<PptxSceneTextParagraph> Paragraphs);

internal sealed record PptxSceneTextParagraph(
    XElement? Properties,
    XElement? EndParagraphProperties,
    int Level,
    PptxSceneParagraphStyle ResolvedStyle,
    IReadOnlyList<PptxSceneTextRun> Runs);

internal sealed record PptxSceneTextRun(
    PptxSceneTextRunKind Kind,
    string Text,
    XElement? Properties,
    PptxSceneRunStyle ResolvedStyle,
    XElement Source);

internal sealed record PptxSceneParagraphStyle(
    int Level,
    string Alignment,
    double FontSize,
    RgbColor Color,
    double Alpha,
    string? Typeface,
    bool Bold,
    bool Italic,
    double CharacterSpacing);

internal sealed record PptxSceneRunStyle(
    double FontSize,
    RgbColor Color,
    double Alpha,
    string? Typeface,
    bool Bold,
    bool Italic,
    bool Underline,
    string? UnderlineValue,
    bool Strike,
    string? StrikeValue,
    string? CapsValue,
    double CharacterSpacing,
    double BaselineOffset,
    RgbColor? Highlight);

internal enum PptxSceneTextRunKind
{
    Text,
    Break,
    Field
}

internal sealed record PptxSceneBounds(
    long XEmu,
    long YEmu,
    long WidthEmu,
    long HeightEmu,
    double RotationDegrees,
    bool FlipHorizontal,
    bool FlipVertical)
{
    public double X => OoxUnits.EmuToPoints(XEmu);
    public double Y => OoxUnits.EmuToPoints(YEmu);
    public double Width => OoxUnits.EmuToPoints(WidthEmu);
    public double Height => OoxUnits.EmuToPoints(HeightEmu);
}

internal enum PptxSceneNodeKind
{
    Shape,
    Picture,
    Table,
    Chart,
    Group,
    Connector,
    UnknownGraphicFrame,
    Unknown
}

internal sealed class PptxSceneBuilder
{
    private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    internal static PptxSceneChartMarkerSymbol ParseChartMarkerSymbol(string? symbol)
    {
        return symbol switch
        {
            "circle" => PptxSceneChartMarkerSymbol.Circle,
            "dash" => PptxSceneChartMarkerSymbol.Dash,
            "diamond" => PptxSceneChartMarkerSymbol.Diamond,
            "dot" => PptxSceneChartMarkerSymbol.Dot,
            "none" => PptxSceneChartMarkerSymbol.None,
            "plus" => PptxSceneChartMarkerSymbol.Plus,
            "square" => PptxSceneChartMarkerSymbol.Square,
            "star" => PptxSceneChartMarkerSymbol.Star,
            "triangle" => PptxSceneChartMarkerSymbol.Triangle,
            "x" => PptxSceneChartMarkerSymbol.X,
            _ when symbol?.Equals("circle", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Circle,
            _ when symbol?.Equals("dash", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Dash,
            _ when symbol?.Equals("diamond", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Diamond,
            _ when symbol?.Equals("dot", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Dot,
            _ when symbol?.Equals("none", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.None,
            _ when symbol?.Equals("plus", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Plus,
            _ when symbol?.Equals("square", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Square,
            _ when symbol?.Equals("star", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Star,
            _ when symbol?.Equals("triangle", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.Triangle,
            _ when symbol?.Equals("x", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartMarkerSymbol.X,
            _ => PptxSceneChartMarkerSymbol.Unknown
        };
    }

    internal static PptxSceneChartDataSourceReferenceKind ParseChartDataSourceReferenceKind(string? referenceKind)
    {
        return referenceKind switch
        {
            "strRef" => PptxSceneChartDataSourceReferenceKind.StringReference,
            "numRef" => PptxSceneChartDataSourceReferenceKind.NumberReference,
            "multiLvlStrRef" => PptxSceneChartDataSourceReferenceKind.MultiLevelStringReference,
            _ => PptxSceneChartDataSourceReferenceKind.Unknown
        };
    }

    internal static PptxSceneChartDataSourceCacheKind ParseChartDataSourceCacheKind(string? cacheKind)
    {
        return cacheKind switch
        {
            "strCache" => PptxSceneChartDataSourceCacheKind.StringCache,
            "numCache" => PptxSceneChartDataSourceCacheKind.NumberCache,
            "multiLvlStrCache" => PptxSceneChartDataSourceCacheKind.MultiLevelStringCache,
            _ => PptxSceneChartDataSourceCacheKind.Unknown
        };
    }

    internal static PptxSceneChartDisplayBlanksAs ParseChartDisplayBlanksAs(string? displayBlanksAs)
    {
        return displayBlanksAs switch
        {
            "gap" => PptxSceneChartDisplayBlanksAs.Gap,
            "span" => PptxSceneChartDisplayBlanksAs.Span,
            "zero" => PptxSceneChartDisplayBlanksAs.Zero,
            _ when displayBlanksAs?.Equals("gap", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDisplayBlanksAs.Gap,
            _ when displayBlanksAs?.Equals("span", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDisplayBlanksAs.Span,
            _ when displayBlanksAs?.Equals("zero", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDisplayBlanksAs.Zero,
            _ => PptxSceneChartDisplayBlanksAs.Unknown
        };
    }

    internal static PptxSceneChartDataLabelPosition ParseChartDataLabelPosition(string? position)
    {
        return position switch
        {
            "bestFit" => PptxSceneChartDataLabelPosition.BestFit,
            "b" => PptxSceneChartDataLabelPosition.Bottom,
            "ctr" => PptxSceneChartDataLabelPosition.Center,
            "inBase" => PptxSceneChartDataLabelPosition.InsideBase,
            "inEnd" => PptxSceneChartDataLabelPosition.InsideEnd,
            "l" => PptxSceneChartDataLabelPosition.Left,
            "outEnd" => PptxSceneChartDataLabelPosition.OutsideEnd,
            "r" => PptxSceneChartDataLabelPosition.Right,
            "t" => PptxSceneChartDataLabelPosition.Top,
            _ when position?.Equals("bestFit", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.BestFit,
            _ when position?.Equals("b", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.Bottom,
            _ when position?.Equals("ctr", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.Center,
            _ when position?.Equals("inBase", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.InsideBase,
            _ when position?.Equals("inEnd", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.InsideEnd,
            _ when position?.Equals("l", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.Left,
            _ when position?.Equals("outEnd", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.OutsideEnd,
            _ when position?.Equals("r", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.Right,
            _ when position?.Equals("t", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartDataLabelPosition.Top,
            _ => PptxSceneChartDataLabelPosition.Unknown
        };
    }

    internal static PptxSceneChartLegendPosition ParseChartLegendPosition(string? position)
    {
        return position switch
        {
            "b" => PptxSceneChartLegendPosition.Bottom,
            "l" => PptxSceneChartLegendPosition.Left,
            "r" => PptxSceneChartLegendPosition.Right,
            "t" => PptxSceneChartLegendPosition.Top,
            "tr" => PptxSceneChartLegendPosition.TopRight,
            _ when position?.Equals("b", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartLegendPosition.Bottom,
            _ when position?.Equals("l", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartLegendPosition.Left,
            _ when position?.Equals("r", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartLegendPosition.Right,
            _ when position?.Equals("t", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartLegendPosition.Top,
            _ when position?.Equals("tr", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartLegendPosition.TopRight,
            _ => PptxSceneChartLegendPosition.Unknown
        };
    }

    internal static PptxSceneChartPlotKind ParseChartPlotKind(string? kind)
    {
        return kind switch
        {
            "areaChart" => PptxSceneChartPlotKind.Area,
            "barChart" => PptxSceneChartPlotKind.Bar,
            "bubbleChart" => PptxSceneChartPlotKind.Bubble,
            "doughnutChart" => PptxSceneChartPlotKind.Doughnut,
            "lineChart" => PptxSceneChartPlotKind.Line,
            "pieChart" => PptxSceneChartPlotKind.Pie,
            "radarChart" => PptxSceneChartPlotKind.Radar,
            "scatterChart" => PptxSceneChartPlotKind.Scatter,
            _ when kind?.Equals("areaChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Area,
            _ when kind?.Equals("barChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Bar,
            _ when kind?.Equals("bubbleChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Bubble,
            _ when kind?.Equals("doughnutChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Doughnut,
            _ when kind?.Equals("lineChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Line,
            _ when kind?.Equals("pieChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Pie,
            _ when kind?.Equals("radarChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Radar,
            _ when kind?.Equals("scatterChart", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartPlotKind.Scatter,
            _ => PptxSceneChartPlotKind.Unknown
        };
    }

    internal static PptxSceneChartGrouping ParseChartGrouping(string? grouping)
    {
        return grouping switch
        {
            "clustered" => PptxSceneChartGrouping.Clustered,
            "percentStacked" => PptxSceneChartGrouping.PercentStacked,
            "stacked" => PptxSceneChartGrouping.Stacked,
            "standard" => PptxSceneChartGrouping.Standard,
            _ when grouping?.Equals("clustered", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartGrouping.Clustered,
            _ when grouping?.Equals("percentStacked", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartGrouping.PercentStacked,
            _ when grouping?.Equals("stacked", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartGrouping.Stacked,
            _ when grouping?.Equals("standard", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartGrouping.Standard,
            _ => PptxSceneChartGrouping.Unknown
        };
    }

    internal static PptxSceneChartBarDirection ParseChartBarDirection(string? direction)
    {
        return direction switch
        {
            "bar" => PptxSceneChartBarDirection.Bar,
            "col" => PptxSceneChartBarDirection.Column,
            _ when direction?.Equals("bar", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartBarDirection.Bar,
            _ when direction?.Equals("col", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartBarDirection.Column,
            _ => PptxSceneChartBarDirection.Unknown
        };
    }

    internal static PptxSceneChartScatterStyle ParseChartScatterStyle(string? style)
    {
        return style switch
        {
            "line" => PptxSceneChartScatterStyle.Line,
            "lineMarker" => PptxSceneChartScatterStyle.LineMarker,
            "marker" => PptxSceneChartScatterStyle.Marker,
            "none" => PptxSceneChartScatterStyle.None,
            "smooth" => PptxSceneChartScatterStyle.Smooth,
            "smoothMarker" => PptxSceneChartScatterStyle.SmoothMarker,
            _ when style?.Equals("line", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.Line,
            _ when style?.Equals("lineMarker", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.LineMarker,
            _ when style?.Equals("marker", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.Marker,
            _ when style?.Equals("none", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.None,
            _ when style?.Equals("smooth", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.Smooth,
            _ when style?.Equals("smoothMarker", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartScatterStyle.SmoothMarker,
            _ => PptxSceneChartScatterStyle.Unknown
        };
    }

    internal static PptxSceneChartRadarStyle ParseChartRadarStyle(string? style)
    {
        return style switch
        {
            "filled" => PptxSceneChartRadarStyle.Filled,
            "marker" => PptxSceneChartRadarStyle.Marker,
            "standard" => PptxSceneChartRadarStyle.Standard,
            _ when style?.Equals("filled", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartRadarStyle.Filled,
            _ when style?.Equals("marker", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartRadarStyle.Marker,
            _ when style?.Equals("standard", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartRadarStyle.Standard,
            _ => PptxSceneChartRadarStyle.Unknown
        };
    }

    internal static PptxSceneChartAxisPosition ParseChartAxisPosition(string? position)
    {
        return position switch
        {
            "b" => PptxSceneChartAxisPosition.Bottom,
            "l" => PptxSceneChartAxisPosition.Left,
            "r" => PptxSceneChartAxisPosition.Right,
            "t" => PptxSceneChartAxisPosition.Top,
            _ when position?.Equals("b", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisPosition.Bottom,
            _ when position?.Equals("l", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisPosition.Left,
            _ when position?.Equals("r", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisPosition.Right,
            _ when position?.Equals("t", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisPosition.Top,
            _ => PptxSceneChartAxisPosition.Unknown
        };
    }

    internal static PptxSceneChartAxisKind ParseChartAxisKind(string? kind)
    {
        return kind switch
        {
            "catAx" => PptxSceneChartAxisKind.Category,
            "dateAx" => PptxSceneChartAxisKind.Date,
            "serAx" => PptxSceneChartAxisKind.Series,
            "valAx" => PptxSceneChartAxisKind.Value,
            _ when kind?.Equals("catAx", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisKind.Category,
            _ when kind?.Equals("dateAx", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisKind.Date,
            _ when kind?.Equals("serAx", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisKind.Series,
            _ when kind?.Equals("valAx", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisKind.Value,
            _ => PptxSceneChartAxisKind.Unknown
        };
    }

    internal static PptxSceneChartManualLayoutTarget ParseChartManualLayoutTarget(string? target)
    {
        return target switch
        {
            "inner" => PptxSceneChartManualLayoutTarget.Inner,
            "outer" => PptxSceneChartManualLayoutTarget.Outer,
            _ when target?.Equals("inner", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartManualLayoutTarget.Inner,
            _ when target?.Equals("outer", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartManualLayoutTarget.Outer,
            _ => PptxSceneChartManualLayoutTarget.Unknown
        };
    }

    internal static PptxSceneChartManualLayoutMode ParseChartManualLayoutMode(string? mode)
    {
        return mode switch
        {
            "edge" => PptxSceneChartManualLayoutMode.Edge,
            "factor" => PptxSceneChartManualLayoutMode.Factor,
            _ when mode?.Equals("edge", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartManualLayoutMode.Edge,
            _ when mode?.Equals("factor", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartManualLayoutMode.Factor,
            _ => PptxSceneChartManualLayoutMode.Unknown
        };
    }

    internal static PptxSceneChartTickLabelPosition ParseChartTickLabelPosition(string? position)
    {
        return position switch
        {
            "high" => PptxSceneChartTickLabelPosition.High,
            "low" => PptxSceneChartTickLabelPosition.Low,
            "nextTo" => PptxSceneChartTickLabelPosition.NextTo,
            "none" => PptxSceneChartTickLabelPosition.None,
            _ when position?.Equals("high", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartTickLabelPosition.High,
            _ when position?.Equals("low", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartTickLabelPosition.Low,
            _ when position?.Equals("nextTo", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartTickLabelPosition.NextTo,
            _ when position?.Equals("none", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartTickLabelPosition.None,
            _ => PptxSceneChartTickLabelPosition.Unknown
        };
    }

    internal static PptxSceneChartAxisCrosses ParseChartAxisCrosses(string? crosses)
    {
        return crosses switch
        {
            "autoZero" => PptxSceneChartAxisCrosses.AutoZero,
            "max" => PptxSceneChartAxisCrosses.Maximum,
            "min" => PptxSceneChartAxisCrosses.Minimum,
            _ when crosses?.Equals("autoZero", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisCrosses.AutoZero,
            _ when crosses?.Equals("max", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisCrosses.Maximum,
            _ when crosses?.Equals("min", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisCrosses.Minimum,
            _ => PptxSceneChartAxisCrosses.Unknown
        };
    }

    internal static PptxSceneChartAxisCrossBetween ParseChartAxisCrossBetween(string? crossBetween)
    {
        return crossBetween switch
        {
            "between" => PptxSceneChartAxisCrossBetween.Between,
            "midCat" => PptxSceneChartAxisCrossBetween.MidpointCategory,
            _ when crossBetween?.Equals("between", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisCrossBetween.Between,
            _ when crossBetween?.Equals("midCat", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisCrossBetween.MidpointCategory,
            _ => PptxSceneChartAxisCrossBetween.Unknown
        };
    }

    internal static PptxSceneChartAxisOrientation ParseChartAxisOrientation(string? orientation)
    {
        return orientation switch
        {
            "minMax" => PptxSceneChartAxisOrientation.MinimumMaximum,
            "maxMin" => PptxSceneChartAxisOrientation.MaximumMinimum,
            _ when orientation?.Equals("minMax", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisOrientation.MinimumMaximum,
            _ when orientation?.Equals("maxMin", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisOrientation.MaximumMinimum,
            _ => PptxSceneChartAxisOrientation.Unknown
        };
    }

    internal static PptxSceneChartAxisTickMark ParseChartAxisTickMark(string? tickMark)
    {
        return tickMark switch
        {
            "cross" => PptxSceneChartAxisTickMark.Cross,
            "in" => PptxSceneChartAxisTickMark.Inside,
            "none" => PptxSceneChartAxisTickMark.None,
            "out" => PptxSceneChartAxisTickMark.Outside,
            _ when tickMark?.Equals("cross", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisTickMark.Cross,
            _ when tickMark?.Equals("in", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisTickMark.Inside,
            _ when tickMark?.Equals("none", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisTickMark.None,
            _ when tickMark?.Equals("out", StringComparison.OrdinalIgnoreCase) == true => PptxSceneChartAxisTickMark.Outside,
            _ => PptxSceneChartAxisTickMark.Unknown
        };
    }

    private const double MinimumStrokeWidth = 0.1d;
    private const double SceneEffectTolerance = 0.001d;
    private const string SlideLayoutRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private const string SlideMasterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    private const string ChartExternalDataPackageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
    private const string ChartColorStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
    private const string ChartStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";

    public PptxScene Build(PptxDocument document, OoxPackage package)
    {
        PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);
        var slides = new List<PptxSceneSlide>(document.Slides.Count);
        foreach (PptxSlide slide in document.Slides)
        {
            OoxPart? slidePart = package.GetPart(slide.PartName);
            if (slidePart is null)
            {
                slides.Add(new PptxSceneSlide(
                    slide.Index,
                    slide.PartName,
                    null,
                    null,
                    null,
                    null,
                    new XDocument(),
                    new Dictionary<string, OoxRelationship>(),
                    new Dictionary<string, OoxRelationship>(),
                    new Dictionary<string, OoxRelationship>(),
                    PptxColorMap.Default,
                    PptxColorMap.Default,
                    PptxColorMap.Default,
                    default,
                    default,
                    default,
                    [],
                    [],
                    []));
                continue;
            }

            XDocument slideXml = LoadXml(slidePart);
            OoxPart? layoutPart = GetRelatedPart(package, slide.PartName, SlideLayoutRelationshipType);
            OoxPart? masterPart = layoutPart is null ? null : GetRelatedPart(package, layoutPart.Name, SlideMasterRelationshipType);
            XDocument? masterXml = masterPart is null ? null : LoadXml(masterPart);
            XDocument? layoutXml = layoutPart is null ? null : LoadXml(layoutPart);
            IReadOnlyDictionary<string, OoxRelationship> masterRelationships = masterPart is null ? new Dictionary<string, OoxRelationship>() : ReadRelationships(package, masterPart.Name);
            IReadOnlyDictionary<string, OoxRelationship> layoutRelationships = layoutPart is null ? new Dictionary<string, OoxRelationship>() : ReadRelationships(package, layoutPart.Name);
            IReadOnlyDictionary<string, OoxRelationship> slideRelationships = ReadRelationships(package, slide.PartName);
            IReadOnlyList<XDocument> layoutSources = masterXml is null ? [] : [masterXml];
            IReadOnlyList<XDocument> slideSources = masterXml is null
                ? layoutXml is null ? [] : [layoutXml]
                : layoutXml is null ? [masterXml] : [masterXml, layoutXml];
            PptxColorMap masterColorMap = ReadMasterColorMap(masterXml);
            PptxColorMap layoutColorMap = ReadColorMapOverride(layoutXml, masterColorMap);
            PptxColorMap slideColorMap = ReadColorMapOverride(slideXml, layoutColorMap);
            slides.Add(new PptxSceneSlide(
                slide.Index,
                slide.PartName,
                masterPart?.Name,
                layoutPart?.Name,
                masterXml,
                layoutXml,
                slideXml,
                masterRelationships,
                layoutRelationships,
                slideRelationships,
                masterColorMap,
                layoutColorMap,
                slideColorMap,
                ReadBackground(masterXml, theme, masterColorMap),
                ReadBackground(layoutXml, theme, layoutColorMap),
                ReadBackground(slideXml, theme, slideColorMap),
                masterXml is null ? [] : ReadNodes(masterXml, [], theme, masterColorMap, package, masterRelationships),
                layoutXml is null ? [] : ReadNodes(layoutXml, layoutSources, theme, layoutColorMap, package, layoutRelationships),
                ReadNodes(slideXml, slideSources, theme, slideColorMap, package, slideRelationships)));
        }

        return new PptxScene(document, theme, slides);
    }

    private static XDocument LoadXml(OoxPart part)
    {
        using Stream stream = part.OpenRead();
        return SafeXml.Load(stream);
    }

    private static PptxColorMap ReadMasterColorMap(XDocument? xml)
    {
        return PptxColorMap.FromElement(xml?.Root?.Element(PresentationNamespace + "clrMap"));
    }

    private static PptxColorMap ReadColorMapOverride(XDocument? xml, PptxColorMap inheritedColorMap)
    {
        XElement? overrideColorMap = xml?.Root?
            .Element(PresentationNamespace + "clrMapOvr")?
            .Element(DrawingNamespace + "overrideClrMapping");
        return PptxColorMap.FromElement(overrideColorMap, inheritedColorMap);
    }

    private static OoxPart? GetRelatedPart(OoxPackage package, string sourcePartName, string relationshipType)
    {
        OoxRelationship? relationship = package.GetRelationships(sourcePartName)
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        return relationship?.ResolvedTarget is null ? null : package.GetPart(relationship.ResolvedTarget);
    }

    private static IReadOnlyDictionary<string, OoxRelationship> ReadRelationships(OoxPackage package, string sourcePartName)
    {
        return package.GetRelationships(sourcePartName)
            .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
    }

    private static IReadOnlyList<PptxSceneNode> ReadNodes(
        XDocument xml,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        PptxColorMap colorMap,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        var nodes = new List<PptxSceneNode>();
        foreach (XElement shapeTree in xml.Descendants(PresentationNamespace + "spTree"))
        {
            nodes.AddRange(ReadChildNodes(shapeTree, placeholderSources, theme, colorMap, package, relationships));
        }

        return nodes;
    }

    private static PptxSceneBackground ReadBackground(XDocument? xml, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? background = xml?.Root?
            .Element(PresentationNamespace + "cSld")?
            .Element(PresentationNamespace + "bg")?
            .Element(PresentationNamespace + "bgPr");
        return TryReadSolidColorWithAlpha(background, theme, colorMap, out RgbColor color, out double alpha)
            ? new PptxSceneBackground(true, color, alpha)
            : default;
    }

    private static IReadOnlyList<PptxSceneNode> ReadChildNodes(
        XElement container,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        PptxColorMap colorMap,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        var nodes = new List<PptxSceneNode>();
        foreach (XElement child in container.Elements())
        {
            PptxSceneNodeKind kind = ReadNodeKind(child);
            if (kind == PptxSceneNodeKind.Unknown)
            {
                continue;
            }

            XElement? nonVisualProperties = ReadNonVisualProperties(child);
            nodes.Add(new PptxSceneNode(
                kind,
                ReadNonVisualId(nonVisualProperties),
                ReadNonVisualName(nonVisualProperties),
                IsPlaceholder(child),
                kind == PptxSceneNodeKind.UnknownGraphicFrame && IsSmartArtGraphicFrame(child),
                ReadHyperlinkClick(nonVisualProperties),
                ReadBounds(child),
                kind is PptxSceneNodeKind.Shape or PptxSceneNodeKind.Connector ? ReadShape(child, theme, package, relationships) : null,
                ReadTextBody(child, placeholderSources, theme, colorMap),
                kind == PptxSceneNodeKind.Picture ? ReadPicture(child, theme, package, relationships) : null,
                kind == PptxSceneNodeKind.Table ? ReadTable(child, theme) : null,
                kind == PptxSceneNodeKind.Chart ? ReadChart(child, package, theme, relationships) : null,
                kind == PptxSceneNodeKind.Group ? ReadGroupTransform(child) : PptxSceneGroupTransform.Identity,
                kind == PptxSceneNodeKind.Group ? ReadChildNodes(child, placeholderSources, theme, colorMap, package, relationships) : [],
                child));
        }

        return nodes;
    }

    internal static PptxSceneNodeKind ReadNodeKind(XElement element)
    {
        if (element.Name == PresentationNamespace + "sp")
        {
            return PptxSceneNodeKind.Shape;
        }

        if (element.Name == PresentationNamespace + "cxnSp")
        {
            return PptxSceneNodeKind.Connector;
        }

        if (element.Name == PresentationNamespace + "pic")
        {
            return PptxSceneNodeKind.Picture;
        }

        if (element.Name == PresentationNamespace + "grpSp")
        {
            return PptxSceneNodeKind.Group;
        }

        if (element.Name != PresentationNamespace + "graphicFrame")
        {
            return PptxSceneNodeKind.Unknown;
        }

        XElement? graphicData = element
            .Descendants(DrawingNamespace + "graphicData")
            .FirstOrDefault();
        string uri = (string?)graphicData?.Attribute("uri") ?? string.Empty;
        if (graphicData?.Descendants(DrawingNamespace + "tbl").Any() == true)
        {
            return PptxSceneNodeKind.Table;
        }

        return uri.Contains("chart", StringComparison.OrdinalIgnoreCase)
            ? PptxSceneNodeKind.Chart
            : PptxSceneNodeKind.UnknownGraphicFrame;
    }

    internal static bool IsSmartArtGraphicFrame(XElement graphicFrame)
    {
        return graphicFrame
            .Descendants(DrawingNamespace + "graphicData")
            .Select(element => (string?)element.Attribute("uri"))
            .Any(uri => uri?.Contains("drawingml/2006/diagram", StringComparison.OrdinalIgnoreCase) == true);
    }

    private static XElement? ReadNonVisualProperties(XElement element)
    {
        return element
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvGraphicFramePr" or "nvCxnSpPr")
            ?.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "cNvPr");
    }

    private static string ReadNonVisualId(XElement? nonVisualProperties)
    {
        return (string?)nonVisualProperties?.Attribute("id") ?? string.Empty;
    }

    private static string ReadNonVisualName(XElement? nonVisualProperties)
    {
        return (string?)nonVisualProperties?.Attribute("name") ?? string.Empty;
    }

    private static PptxSceneHyperlinkClick ReadHyperlinkClick(XElement? nonVisualProperties)
    {
        XElement? hyperlink = nonVisualProperties?.Element(DrawingNamespace + "hlinkClick");
        if (hyperlink is null)
        {
            return default;
        }

        return new PptxSceneHyperlinkClick(
            true,
            (string?)hyperlink.Attribute(RelationshipsNamespace + "id"),
            (string?)hyperlink.Attribute("action"));
    }

    private static bool IsPlaceholder(XElement element)
    {
        return element
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvGraphicFramePr" or "nvCxnSpPr")
            ?.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "nvPr")
            ?.Element(PresentationNamespace + "ph") is not null;
    }

    private static PptxSceneBounds? ReadBounds(XElement element)
    {
        XElement? transform = element
            .Element(PresentationNamespace + "spPr")?
            .Element(DrawingNamespace + "xfrm") ??
            element.Element(PresentationNamespace + "grpSpPr")?
                .Element(DrawingNamespace + "xfrm") ??
            element.Element(PresentationNamespace + "xfrm");
        if (transform is null)
        {
            return null;
        }

        XElement? offset = transform.Element(DrawingNamespace + "off");
        XElement? extents = transform.Element(DrawingNamespace + "ext");
        if (offset is null || extents is null)
        {
            return null;
        }

        return new PptxSceneBounds(
            ReadLong(offset, "x"),
            ReadLong(offset, "y"),
            ReadLong(extents, "cx"),
            ReadLong(extents, "cy"),
            transform.Attribute("rot") is { } rotation ? long.Parse(rotation.Value, CultureInfo.InvariantCulture) / 60000d : 0d,
            ReadBool(transform, "flipH"),
            ReadBool(transform, "flipV"));
    }

    internal static string? ReadPictureRelationshipId(XElement picture)
    {
        XElement? blip = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip");
        return (string?)blip?.Attribute(RelationshipsNamespace + "embed") ??
            blip?.Descendants()
                .FirstOrDefault(element => element.Name.LocalName == "svgBlip")
                ?.Attribute(RelationshipsNamespace + "embed")
                ?.Value;
    }

    private static PptxScenePicture ReadPicture(
        XElement picture,
        PptxTheme theme,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        string? relationshipId = ReadPictureRelationshipId(picture);
        string? targetPartName = ResolveRelationshipTarget(relationshipId, relationships);
        return new PptxScenePicture(
            relationshipId,
            targetPartName,
            ReadImageResource(package, targetPartName),
            ReadPictureCrop(picture),
            ReadPictureFill(picture),
            ReadPictureAlpha(picture),
            ReadPictureAlphaValue(picture),
            ReadImageRecolor(picture, theme),
            ReadPictureTile(picture));
    }

    private static PptxSceneChart ReadChart(
        XElement frame,
        OoxPackage package,
        PptxTheme theme,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        XElement? graphicData = frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData");
        string? relationshipId = (string?)graphicData
            ?.Element(ChartNamespace + "chart")
            ?.Attribute(RelationshipsNamespace + "id");
        string? targetPartName = relationshipId is not null &&
            relationships.TryGetValue(relationshipId, out OoxRelationship? relationship)
            ? relationship.ResolvedTarget
            : null;
        OoxPart? chartPart = targetPartName is null ? null : package.GetPart(targetPartName);
        XDocument? chartXml = chartPart is null ? null : LoadXml(chartPart);
        PptxSceneChartExternalData externalData = chartPart is null
            ? default
            : ReadChartExternalData(package, chartPart.Name, chartXml);
        PptxSceneChartColorStyle colorStyle = chartPart is null
            ? new PptxSceneChartColorStyle(false, null, string.Empty, string.Empty, [], 0, [], null)
            : ReadChartColorStyle(package, chartPart.Name, theme);
        PptxSceneChartStyle stylePart = chartPart is null
            ? new PptxSceneChartStyle(false, null, string.Empty, null, [])
            : ReadChartStylePart(package, chartPart.Name, theme);
        IReadOnlyList<RgbColor>? paletteColors = colorStyle.Colors.Count == 0 ? null : colorStyle.Colors;
        IReadOnlyList<PptxSceneChartPlot> plots = ReadChartPlots(chartXml, theme);
        IReadOnlyList<PptxSceneChartAxis> axes = ReadChartAxes(chartXml, theme, stylePart);
        return new PptxSceneChart(
            relationshipId,
            targetPartName,
            chartXml,
            externalData,
            ReadChartOptions(chartXml),
            paletteColors,
            colorStyle,
            stylePart,
            (string?)chartXml?.Root?.Element(ChartNamespace + "style")?.Attribute("val") ?? string.Empty,
            plots,
            axes,
            ReadChartTitle(chartXml, theme, plots),
            ReadChartLegend(chartXml, theme),
            ReadChartTextStyleOverride(chartXml?.Root, theme),
            ReadChartPlotAreaManualLayout(chartXml),
            ReadChartShapeStyle(chartXml?.Root?.Element(ChartNamespace + "spPr"), theme),
            ReadChartShapeStyle(chartXml?
                .Descendants(ChartNamespace + "plotArea")
                .FirstOrDefault()
                ?.Element(ChartNamespace + "spPr"), theme));
    }

    private static PptxSceneChartOptions ReadChartOptions(XDocument? chartXml)
    {
        XElement? chartSpace = chartXml?.Root;
        XElement? chart = chartSpace?.Element(ChartNamespace + "chart");
        string displayBlanksAs = chart is null ? string.Empty : ReadChartElementValue(chart, "dispBlanksAs");
        (bool? date1904, string date1904Value) = ReadOptionalOoxmlBooleanElementWithValue(chartSpace, "date1904");
        (bool? roundedCorners, string roundedCornersValue) = ReadOptionalOoxmlBooleanElementWithValue(chartSpace, "roundedCorners");
        (bool? plotVisibleOnly, string plotVisibleOnlyValue) = ReadOptionalOoxmlBooleanElementWithValue(chart, "plotVisOnly");
        (bool? showDataLabelsOverMaximum, string showDataLabelsOverMaximumValue) = ReadOptionalOoxmlBooleanElementWithValue(chart, "showDLblsOverMax");
        return new PptxSceneChartOptions(
            date1904,
            date1904Value,
            roundedCorners,
            roundedCornersValue,
            plotVisibleOnly,
            plotVisibleOnlyValue,
            showDataLabelsOverMaximum,
            showDataLabelsOverMaximumValue,
            ParseChartDisplayBlanksAs(displayBlanksAs),
            displayBlanksAs);
    }

    private static IReadOnlyList<PptxSceneChartPlot> ReadChartPlots(XDocument? chartXml, PptxTheme theme)
    {
        XElement? plotArea = chartXml?
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault();
        if (plotArea is null)
        {
            return [];
        }

        var plots = new List<PptxSceneChartPlot>();
        var kindIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (XElement plot in plotArea.Elements().Where(element => element.Name.Namespace == ChartNamespace && element.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal)))
        {
            string kind = plot.Name.LocalName;
            int kindIndex = kindIndexes.TryGetValue(kind, out int nextKindIndex) ? nextKindIndex : 0;
            kindIndexes[kind] = kindIndex + 1;
            string grouping = ReadChartElementValue(plot, "grouping");
            string barDirection = ReadChartElementValue(plot, "barDir");
            string scatterStyle = ReadChartElementValue(plot, "scatterStyle");
            string radarStyle = ReadChartElementValue(plot, "radarStyle");
            (bool? markersEnabled, string markersEnabledValue) = ReadChartPlotMarkersEnabled(plot);
            (bool? varyColors, string varyColorsValue) = ReadChartPlotVaryColors(plot);
            PptxSceneChartPlotKind plotKind = ParseChartPlotKind(kind);
            (double? gapWidth, string gapWidthValue) = ReadChartElementDoubleWithValue(plot, "gapWidth");
            (double? overlap, string overlapValue) = ReadChartElementDoubleWithValue(plot, "overlap");
            (double? holeSize, string holeSizeValue) = ReadChartElementDoubleWithValue(plot, "holeSize");
            (double? firstSliceAngle, string firstSliceAngleValue) = ReadChartElementDoubleWithValue(plot, "firstSliceAng");
            string[] axisIds = plot
                .Elements(ChartNamespace + "axId")
                .Select(axis => (string?)axis.Attribute("val") ?? string.Empty)
                .Where(value => value.Length != 0)
                .ToArray();
            plots.Add(new PptxSceneChartPlot(
                plotKind,
                kind,
                plots.Count,
                kindIndex,
                plot.Elements(ChartNamespace + "ser").Count(),
                axisIds,
                ReadChartSeries(plot, theme, plotKind, markersEnabled == true),
                ParseChartGrouping(grouping),
                grouping,
                ParseChartBarDirection(barDirection),
                barDirection,
                ParseChartScatterStyle(scatterStyle),
                scatterStyle,
                ParseChartRadarStyle(radarStyle),
                radarStyle,
                markersEnabled,
                markersEnabledValue,
                varyColors,
                varyColorsValue,
                gapWidth,
                gapWidthValue,
                overlap,
                overlapValue,
                holeSize,
                holeSizeValue,
                firstSliceAngle,
                firstSliceAngleValue,
                ReadChartDataLabels(plot, theme),
                plot));
        }

        return plots;
    }

    private static PptxSceneChartDataLabels ReadChartDataLabels(XElement plot, PptxTheme theme)
    {
        XElement? labels = plot.Element(ChartNamespace + "dLbls") ??
            plot.Elements(ChartNamespace + "ser")
                .Select(series => series.Element(ChartNamespace + "dLbls"))
                .FirstOrDefault(element => element is not null);
        if (labels is null)
        {
            return new PptxSceneChartDataLabels(
                ShowValue: null,
                ShowValueValue: string.Empty,
                ShowPercent: null,
                ShowPercentValue: string.Empty,
                ShowCategoryName: null,
                ShowCategoryNameValue: string.Empty,
                ShowSeriesName: null,
                ShowSeriesNameValue: string.Empty,
                ShowLeaderLines: null,
                ShowLeaderLinesValue: string.Empty,
                ShowLegendKey: null,
                ShowLegendKeyValue: string.Empty,
                ShowBubbleSize: null,
                ShowBubbleSizeValue: string.Empty,
                LeaderLines: default,
                PositionKind: PptxSceneChartDataLabelPosition.Unknown,
                Position: string.Empty,
                Separator: string.Empty,
                NumberFormat: string.Empty,
                NumberFormatInfo: default,
                Layout: default,
                TextStyle: new PptxSceneChartTextStyleOverride(null, null, null, null, null, null, null, null),
                TextBodyProperties: default,
                ShapeStyle: new PptxSceneChartShapeStyle(false, default, default, default, default, default, default, default, default),
                RejectedOverrideIndexValues: [],
                Overrides: [],
                IsDefined: false);
        }

        PptxSceneChartNumberFormat numberFormat = ReadChartNumberFormat(labels);
        (bool? showValue, string showValueValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showVal");
        (bool? showPercent, string showPercentValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showPercent");
        (bool? showCategoryName, string showCategoryNameValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showCatName");
        (bool? showSeriesName, string showSeriesNameValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showSerName");
        (bool? showLeaderLines, string showLeaderLinesValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showLeaderLines");
        (bool? showLegendKey, string showLegendKeyValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showLegendKey");
        (bool? showBubbleSize, string showBubbleSizeValue) = ReadOptionalOoxmlBooleanElementWithValue(labels, "showBubbleSize");
        return new PptxSceneChartDataLabels(
                showValue,
                showValueValue,
                showPercent,
                showPercentValue,
                showCategoryName,
                showCategoryNameValue,
                showSeriesName,
                showSeriesNameValue,
                showLeaderLines,
                showLeaderLinesValue,
                showLegendKey,
                showLegendKeyValue,
                showBubbleSize,
                showBubbleSizeValue,
                ReadChartLeaderLines(labels, theme),
                ParseChartDataLabelPosition(ReadChartElementValue(labels, "dLblPos")),
                ReadChartElementValue(labels, "dLblPos"),
                labels.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                numberFormat.FormatCode,
                numberFormat,
                ReadChartManualLayout(labels),
                ReadChartTextStyleOverride(labels, theme),
                ReadChartTextBodyProperties(labels),
                ReadChartShapeStyle(labels.Element(ChartNamespace + "spPr"), theme),
                ReadRejectedChartNonNegativeIndexValues(labels, "dLbl"),
                ReadChartDataLabelOverrides(labels, theme),
                IsDefined: true);
    }

    private static PptxSceneChartLeaderLines ReadChartLeaderLines(XElement labels, PptxTheme theme)
    {
        XElement? leaderLines = labels.Element(ChartNamespace + "leaderLines");
        return leaderLines is null
            ? default
            : new PptxSceneChartLeaderLines(true, ReadChartLine(leaderLines.Element(ChartNamespace + "spPr"), theme));
    }

    private static IReadOnlyList<PptxSceneChartDataLabelOverride> ReadChartDataLabelOverrides(XElement labels, PptxTheme theme)
    {
        var overrides = new List<PptxSceneChartDataLabelOverride>();
        foreach (XElement label in labels.Elements(ChartNamespace + "dLbl"))
        {
            if (!TryReadChartNonNegativeIndex(label, out int index, out string indexValue))
            {
                continue;
            }

            PptxSceneChartNumberFormat numberFormat = ReadChartNumberFormat(label);
            (bool? showValue, string showValueValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showVal");
            (bool? showPercent, string showPercentValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showPercent");
            (bool? showCategoryName, string showCategoryNameValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showCatName");
            (bool? showSeriesName, string showSeriesNameValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showSerName");
            (bool? showLeaderLines, string showLeaderLinesValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showLeaderLines");
            (bool? showLegendKey, string showLegendKeyValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showLegendKey");
            (bool? showBubbleSize, string showBubbleSizeValue) = ReadOptionalOoxmlBooleanElementWithValue(label, "showBubbleSize");
            overrides.Add(new PptxSceneChartDataLabelOverride(
                index,
                indexValue,
                showValue,
                showValueValue,
                showPercent,
                showPercentValue,
                showCategoryName,
                showCategoryNameValue,
                showSeriesName,
                showSeriesNameValue,
                showLeaderLines,
                showLeaderLinesValue,
                showLegendKey,
                showLegendKeyValue,
                showBubbleSize,
                showBubbleSizeValue,
                ReadChartLeaderLines(label, theme),
                ReadChartText(label.Element(ChartNamespace + "tx")) ?? string.Empty,
                ReadChartTextRuns(label.Element(ChartNamespace + "tx"), theme),
                ParseChartDataLabelPosition(ReadChartElementValue(label, "dLblPos")),
                ReadChartElementValue(label, "dLblPos"),
                label.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                numberFormat.FormatCode,
                numberFormat,
                ReadChartManualLayout(label),
                ReadChartTextStyleOverride(label, theme),
                ReadChartTextBodyProperties(label),
                ReadChartShapeStyle(label.Element(ChartNamespace + "spPr"), theme)));
        }

        return overrides;
    }

    private static bool? ReadOptionalOoxmlBooleanElement(XElement parent, string elementName)
    {
        XElement? element = parent.Element(ChartNamespace + elementName);
        return element is null ? null : IsOoxmlBooleanElementEnabled(element);
    }

    private static (bool? Value, string RawValue) ReadOptionalOoxmlBooleanElementWithValue(XElement? parent, string elementName)
    {
        XElement? element = parent?.Element(ChartNamespace + elementName);
        return element is null
            ? (null, string.Empty)
            : (IsOoxmlBooleanElementEnabled(element), (string?)element.Attribute("val") ?? string.Empty);
    }

    private static PptxSceneChartNumberFormat ReadChartNumberFormat(XElement parent)
    {
        XElement? numberFormat = parent.Element(ChartNamespace + "numFmt");
        return numberFormat is null
            ? default
            : new PptxSceneChartNumberFormat(
                IsDefined: true,
                FormatCode: (string?)numberFormat.Attribute("formatCode") ?? string.Empty,
                SourceLinked: ReadOptionalOoxmlBooleanAttribute(numberFormat, "sourceLinked"),
                SourceLinkedValue: (string?)numberFormat.Attribute("sourceLinked") ?? string.Empty);
    }

    private static PptxSceneChartShapeStyle ReadChartShapeStyle(XElement? shapeProperties, PptxTheme theme)
    {
        bool noFill = shapeProperties?.Element(DrawingNamespace + "noFill") is not null;
        PptxSceneFillStyle fill = !noFill && TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
            ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
            : default;
        return new PptxSceneChartShapeStyle(
            noFill,
            fill,
            !noFill && TryReadShapeGradientFill(shapeProperties, theme, out PptxSceneGradientFill gradientFill) ? gradientFill : new PptxSceneGradientFill(false, 0d, []),
            noFill ? default : ReadChartPatternFill(shapeProperties, theme),
            noFill ? default : ReadChartPictureFill(shapeProperties),
            ReadChartLine(shapeProperties, theme),
            TryReadGlow(shapeProperties, theme, out PptxSceneGlow glow) ? glow : default,
            TryReadOuterShadow(shapeProperties, theme, out PptxSceneOuterShadow outerShadow) ? outerShadow : default,
            ReadChartEffects(shapeProperties));
    }

    private static PptxSceneShapePictureFill ReadChartPictureFill(XElement? shapeProperties)
    {
        XElement? blipFill = shapeProperties?.Element(DrawingNamespace + "blipFill");
        if (blipFill is null)
        {
            return default;
        }

        string relationshipId = (string?)blipFill
            .Element(DrawingNamespace + "blip")
            ?.Attribute(RelationshipsNamespace + "embed")
            ?? string.Empty;
        return new PptxSceneShapePictureFill(
            true,
            relationshipId,
            null,
            null,
            ReadPictureCrop(shapeProperties!),
            ReadPictureFill(shapeProperties!),
            ReadPictureAlpha(shapeProperties!),
            ReadPictureAlphaValue(shapeProperties!),
            ReadPictureTile(shapeProperties!));
    }

    private static PptxSceneChartEffectFamily ReadChartEffects(XElement? shapeProperties)
    {
        XElement? effectList = shapeProperties?.Element(DrawingNamespace + "effectLst");
        XElement? effectDag = shapeProperties?.Element(DrawingNamespace + "effectDag");
        IReadOnlyList<string> unsupportedEffects = effectList is null
            ? []
            : effectList
                .Elements()
                .Where(IsUnsupportedChartEffect)
                .Select(effect => effect.Name.LocalName)
                .ToArray();
        return new PptxSceneChartEffectFamily(effectList is not null, effectDag is not null, unsupportedEffects);
    }

    private static bool IsUnsupportedChartEffect(XElement effect)
    {
        return effect.Name != DrawingNamespace + "outerShdw" &&
            effect.Name != DrawingNamespace + "glow";
    }

    private static string ReadChartElementValue(XElement element, string childName)
    {
        return (string?)element.Element(ChartNamespace + childName)?.Attribute("val") ?? string.Empty;
    }

    private static double? ReadChartElementDouble(XElement element, string childName)
    {
        (double? parsed, _) = ReadChartElementDoubleWithValue(element, childName);
        return parsed;
    }

    private static (double? Value, string RawValue) ReadChartElementDoubleWithValue(XElement element, string childName)
    {
        string value = (string?)element.Element(ChartNamespace + childName)?.Attribute("val") ?? string.Empty;
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? (parsed, value)
            : (null, value);
    }

    private static int? ReadChartElementInt(XElement element, string childName)
    {
        (int? parsed, _) = ReadChartElementIntWithValue(element, childName);
        return parsed;
    }

    private static (int? Value, string RawValue) ReadChartElementIntWithValue(XElement element, string childName)
    {
        string value = (string?)element.Element(ChartNamespace + childName)?.Attribute("val") ?? string.Empty;
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? (parsed, value)
            : (null, value);
    }

    private static (bool? Value, string RawValue) ReadChartPlotVaryColors(XElement plot)
    {
        XElement? varyColors = plot.Element(ChartNamespace + "varyColors");
        return varyColors is null
            ? (null, string.Empty)
            : (IsOoxmlBooleanElementEnabled(varyColors, defaultValue: true), (string?)varyColors.Attribute("val") ?? string.Empty);
    }

    private static (bool? Value, string RawValue) ReadChartPlotMarkersEnabled(XElement plot)
    {
        XElement? marker = plot.Element(ChartNamespace + "marker");
        return marker is null
            ? (null, string.Empty)
            : (IsOoxmlBooleanElementEnabled(marker), (string?)marker.Attribute("val") ?? string.Empty);
    }

    private static IReadOnlyList<PptxSceneChartSeries> ReadChartSeries(XElement plot, PptxTheme theme, PptxSceneChartPlotKind plotKind, bool chartMarkersEnabled)
    {
        var series = new List<PptxSceneChartSeries>();
        foreach (XElement seriesElement in plot.Elements(ChartNamespace + "ser"))
        {
            int seriesIndex = series.Count;
            (int? index, string indexValue) = ReadChartElementIntWithValue(seriesElement, "idx");
            (int? order, string orderValue) = ReadChartElementIntWithValue(seriesElement, "order");
            (double? explosion, string explosionValue) = ReadChartElementDoubleWithValue(seriesElement, "explosion");
            (int? valuePointCount, string valuePointCountValue) = ReadChartSeriesPointCountWithValue(seriesElement, "val");
            (int? categoryPointCount, string categoryPointCountValue) = ReadChartSeriesPointCountWithValue(seriesElement, "cat");
            (int? xValuePointCount, string xValuePointCountValue) = ReadChartSeriesPointCountWithValue(seriesElement, "xVal");
            (int? yValuePointCount, string yValuePointCountValue) = ReadChartSeriesPointCountWithValue(seriesElement, "yVal");
            (int? bubbleSizePointCount, string bubbleSizePointCountValue) = ReadChartSeriesPointCountWithValue(seriesElement, "bubbleSize");
            (bool? smooth, string smoothValue) = ReadChartSeriesSmooth(seriesElement);
            series.Add(new PptxSceneChartSeries(
                index,
                indexValue,
                order,
                orderValue,
                ReadChartSeriesName(seriesElement),
                ReadChartSeriesDataSources(seriesElement),
                ReadChartSeriesValues(seriesElement),
                ReadChartSeriesNumberPoints(seriesElement, "val"),
                valuePointCount,
                valuePointCountValue,
                ReadChartSeriesNumberFormatCode(seriesElement, "val"),
                ReadChartSeriesCategories(seriesElement),
                ReadChartSeriesStringPoints(seriesElement, "cat"),
                categoryPointCount,
                categoryPointCountValue,
                ReadChartSeriesStringLevels(seriesElement, "cat"),
                ReadChartSeriesNumbers(seriesElement, "xVal"),
                ReadChartSeriesNumberPoints(seriesElement, "xVal"),
                xValuePointCount,
                xValuePointCountValue,
                ReadChartSeriesNumberFormatCode(seriesElement, "xVal"),
                ReadChartSeriesNumbers(seriesElement, "yVal"),
                ReadChartSeriesNumberPoints(seriesElement, "yVal"),
                yValuePointCount,
                yValuePointCountValue,
                ReadChartSeriesNumberFormatCode(seriesElement, "yVal"),
                ReadChartSeriesNumbers(seriesElement, "bubbleSize"),
                ReadChartSeriesNumberPoints(seriesElement, "bubbleSize"),
                bubbleSizePointCount,
                bubbleSizePointCountValue,
                ReadChartSeriesNumberFormatCode(seriesElement, "bubbleSize"),
                ReadChartSeriesFill(seriesElement, theme),
                ReadChartSeriesPatternFill(seriesElement, theme),
                ReadChartSeriesLine(seriesElement, theme),
                ReadChartMarker(seriesElement, theme, plotKind, chartMarkersEnabled, seriesIndex),
                ReadChartPointStyles(seriesElement, theme),
                ReadRejectedChartNonNegativeIndexValues(seriesElement, "dPt"),
                explosion,
                explosionValue,
                smooth,
                smoothValue,
                ReadChartDataLabels(seriesElement, theme)));
        }

        return series;
    }

    private static PptxSceneFillStyle ReadChartSeriesFill(XElement series, PptxTheme theme)
    {
        XElement? shapeProperties = series.Element(ChartNamespace + "spPr");
        return TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    private static PptxSceneLineStyle ReadChartSeriesLine(XElement series, PptxTheme theme)
    {
        return ReadChartLine(series.Element(ChartNamespace + "spPr"), theme);
    }

    private static PptxScenePatternFill ReadChartSeriesPatternFill(XElement series, PptxTheme theme)
    {
        return ReadChartPatternFill(series.Element(ChartNamespace + "spPr"), theme);
    }

    private static PptxScenePatternFill ReadChartPatternFill(XElement? shapeProperties, PptxTheme theme)
    {
        XElement? patternFill = shapeProperties?.Element(DrawingNamespace + "pattFill");
        if (patternFill is null)
        {
            return default;
        }

        RgbColor foreground = TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "fgClr"), theme, out RgbColor foregroundColor, out _)
            ? foregroundColor
            : new RgbColor(0, 0, 0);
        RgbColor background = TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "bgClr"), theme, out RgbColor backgroundColor, out _)
            ? backgroundColor
            : new RgbColor(255, 255, 255);
        return new PptxScenePatternFill(
            HasPattern: true,
            (string?)patternFill.Attribute("prst") ?? "pct50",
            foreground,
            background,
            Alpha: 1d);
    }

    private static PptxSceneChartMarker ReadChartMarker(XElement series, PptxTheme theme, PptxSceneChartPlotKind plotKind, bool chartMarkersEnabled, int seriesIndex)
    {
        XElement? marker = series.Element(ChartNamespace + "marker");
        string symbol = (string?)marker?.Element(ChartNamespace + "symbol")?.Attribute("val") ??
            ReadDefaultChartMarkerSymbol(plotKind, chartMarkersEnabled, seriesIndex);
        string? sizeValue = (string?)marker?.Element(ChartNamespace + "size")?.Attribute("val");
        double size = PptxChartMarkerMetricRules.ResolveSize(
            sizeValue,
            plotKind,
            chartMarkersEnabled,
            marker is not null,
            marker?.Element(ChartNamespace + "spPr") is not null);
        XElement? shapeProperties = marker?.Element(ChartNamespace + "spPr");
        PptxSceneFillStyle fill = TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
            ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
            : default;
        return new PptxSceneChartMarker(marker is not null, ParseChartMarkerSymbol(symbol), symbol, sizeValue, size, fill, ReadChartLine(shapeProperties, theme));
    }

    private static string ReadDefaultChartMarkerSymbol(PptxSceneChartPlotKind plotKind, bool chartMarkersEnabled, int seriesIndex)
    {
        if (plotKind == PptxSceneChartPlotKind.Line)
        {
            return chartMarkersEnabled ? AutoLineChartMarkerSymbol(seriesIndex) : "none";
        }

        return "circle";
    }

    private static string AutoLineChartMarkerSymbol(int seriesIndex)
    {
        ReadOnlySpan<string> symbols = ["diamond", "square", "triangle", "x", "star", "circle"];
        return symbols[seriesIndex % symbols.Length];
    }

    private static IReadOnlyList<PptxSceneChartPointStyle> ReadChartPointStyles(XElement series, PptxTheme theme)
    {
        var styles = new List<PptxSceneChartPointStyle>();
        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (!TryReadChartNonNegativeIndex(point, out int index, out string indexValue))
            {
                continue;
            }

            XElement? shapeProperties = point.Element(ChartNamespace + "spPr");
            (double? explosion, string explosionValue) = ReadChartElementDoubleWithValue(point, "explosion");
            styles.Add(new PptxSceneChartPointStyle(
                index,
                indexValue,
                ReadChartPointFill(shapeProperties, theme),
                ReadChartPointPatternFill(shapeProperties, theme),
                ReadChartLine(shapeProperties, theme),
                explosion,
                explosionValue));
        }

        return styles;
    }

    private static bool TryReadChartNonNegativeIndex(XElement element, out int index, out string indexValue)
    {
        indexValue = (string?)element.Element(ChartNamespace + "idx")?.Attribute("val") ?? string.Empty;
        return int.TryParse(indexValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out index) &&
            index >= 0;
    }

    private static IReadOnlyList<string> ReadRejectedChartNonNegativeIndexValues(XElement parent, string elementName)
    {
        var rejected = new List<string>();
        foreach (XElement element in parent.Elements(ChartNamespace + elementName))
        {
            if (!TryReadChartNonNegativeIndex(element, out _, out string indexValue))
            {
                rejected.Add(indexValue);
            }
        }

        return rejected;
    }

    private static PptxSceneFillStyle ReadChartPointFill(XElement? shapeProperties, PptxTheme theme)
    {
        return TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    private static PptxScenePatternFill ReadChartPointPatternFill(XElement? shapeProperties, PptxTheme theme)
    {
        return ReadChartPatternFill(shapeProperties, theme);
    }

    private static PptxSceneLineStyle ReadChartLine(XElement? shapeProperties, PptxTheme theme)
    {
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        bool widthSpecified = line?.Attribute("w") is not null;
        return shapeProperties is not null &&
            TryReadLineWithAlpha(shapeProperties, theme, out RgbColor color, out double lineWidth, out double alpha)
                ? new PptxSceneLineStyle(
                    true,
                    color,
                    lineWidth,
                    alpha,
                    TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> dashPattern) ? dashPattern : [],
                    ReadPresetDashValue(shapeProperties),
                    ReadLineCompound(shapeProperties),
                    ReadLineCompoundValue(shapeProperties),
                    ReadLineCap(shapeProperties) switch
                    {
                        "rnd" => 1,
                        "sq" => 2,
                        _ => null
                    },
                    ReadLineCap(shapeProperties),
                    ReadLineJoin(shapeProperties),
                    ReadLineJoinValue(shapeProperties),
                    widthSpecified)
                : default;
    }

    private static (bool? Value, string RawValue) ReadChartSeriesSmooth(XElement series)
    {
        XElement? smooth = series.Element(ChartNamespace + "smooth");
        return smooth is null
            ? (null, string.Empty)
            : (IsOoxmlBooleanElementEnabled(smooth), (string?)smooth.Attribute("val") ?? string.Empty);
    }

    private static string? ReadChartSeriesName(XElement series)
    {
        return ReadChartText(series.Element(ChartNamespace + "tx"));
    }

    private static PptxSceneChartSeriesDataSources ReadChartSeriesDataSources(XElement series)
    {
        return new PptxSceneChartSeriesDataSources(
            ReadChartDataSource(series.Element(ChartNamespace + "tx"), "strRef", "numRef"),
            ReadChartDataSource(series.Element(ChartNamespace + "val"), "numRef"),
            ReadChartDataSource(series.Element(ChartNamespace + "cat"), "strRef", "numRef", "multiLvlStrRef"),
            ReadChartDataSource(series.Element(ChartNamespace + "xVal"), "numRef", "strRef", "multiLvlStrRef"),
            ReadChartDataSource(series.Element(ChartNamespace + "yVal"), "numRef"),
            ReadChartDataSource(series.Element(ChartNamespace + "bubbleSize"), "numRef"));
    }

    private static PptxSceneChartDataSource ReadChartDataSource(XElement? container, params string[] referenceKinds)
    {
        if (container is null)
        {
            return default;
        }

        foreach (string referenceKind in referenceKinds)
        {
            XElement? reference = container.Element(ChartNamespace + referenceKind);
            if (reference is null)
            {
                continue;
            }

            XElement? cache = reference
                .Elements()
                .FirstOrDefault(element => element.Name.Namespace == ChartNamespace && element.Name.LocalName.EndsWith("Cache", StringComparison.Ordinal));
            bool hasCachedPoints = cache?
                .Descendants(ChartNamespace + "pt")
                .Any() == true;
            return new PptxSceneChartDataSource(
                (string?)reference.Element(ChartNamespace + "f"),
                ParseChartDataSourceReferenceKind(referenceKind),
                referenceKind,
                ParseChartDataSourceCacheKind(cache?.Name.LocalName),
                cache?.Name.LocalName ?? string.Empty,
                hasCachedPoints);
        }

        return default;
    }

    private static string? ReadChartText(XElement? text)
    {
        string? literal = text?
            .Descendants(ChartNamespace + "v")
            .Select(value => value.Value)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (!string.IsNullOrWhiteSpace(literal))
        {
            return literal;
        }

        string? richText = text?
            .Descendants(DrawingNamespace + "t")
            .Aggregate(string.Empty, (current, textElement) => current + textElement.Value);
        return string.IsNullOrWhiteSpace(richText) ? null : richText;
    }

    internal static IReadOnlyList<PptxSceneChartTextRun> ReadChartTextRuns(XElement? text, PptxTheme theme)
    {
        XElement? rich = text?.Element(ChartNamespace + "rich");
        if (rich is null)
        {
            string? literal = text?
                .Descendants(ChartNamespace + "v")
                .Select(value => value.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            return string.IsNullOrWhiteSpace(literal)
                ? []
                : [new PptxSceneChartTextRun(literal!, default)];
        }

        var runs = new List<PptxSceneChartTextRun>();
        foreach (XElement run in rich.Descendants(DrawingNamespace + "r"))
        {
            string runText = run.Element(DrawingNamespace + "t")?.Value ?? string.Empty;
            if (runText.Length == 0)
            {
                continue;
            }

            runs.Add(new PptxSceneChartTextRun(runText, ReadChartTextRunStyle(run.Element(DrawingNamespace + "rPr"), theme)));
        }

        return runs;
    }

    private static IReadOnlyList<double> ReadChartSeriesValues(XElement series)
    {
        return ReadChartSeriesNumbers(series, "val");
    }

    private static IReadOnlyList<PptxSceneChartNumberPoint> ReadChartSeriesNumberPoints(XElement series, string elementName)
    {
        var points = new List<PptxSceneChartNumberPoint>();
        int ordinal = 0;
        foreach (XElement point in series
            .Elements(ChartNamespace + elementName)
            .Descendants(ChartNamespace + "pt"))
        {
            string indexValue = (string?)point.Attribute("idx") ?? string.Empty;
            bool hasParsedIndex = int.TryParse(indexValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedIndex);
            int index = hasParsedIndex
                ? parsedIndex
                : ordinal;
            XElement? valueElement = point.Element(ChartNamespace + "v");
            string text = (string?)valueElement ?? string.Empty;
            double? value = double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                ? parsed
                : null;
            points.Add(new PptxSceneChartNumberPoint(index, indexValue, hasParsedIndex, value, text, valueElement is not null));
            ordinal++;
        }

        return points;
    }

    private static int? ReadChartSeriesPointCount(XElement series, string elementName)
    {
        (int? parsed, _) = ReadChartSeriesPointCountWithValue(series, elementName);
        return parsed;
    }

    private static (int? Value, string RawValue) ReadChartSeriesPointCountWithValue(XElement series, string elementName)
    {
        XElement? cache = series
            .Elements(ChartNamespace + elementName)
            .Descendants()
            .FirstOrDefault(child => child.Name.Namespace == ChartNamespace &&
                (child.Name.LocalName == "numLit" ||
                 child.Name.LocalName == "numCache" ||
                 child.Name.LocalName == "strLit" ||
                 child.Name.LocalName == "strCache" ||
                 child.Name.LocalName == "multiLvlStrCache"));
        string value = (string?)cache?
            .Element(ChartNamespace + "ptCount")
            ?.Attribute("val") ?? string.Empty;
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
            ? (parsed, value)
            : (null, value);
    }

    private static string? ReadChartSeriesNumberFormatCode(XElement series, string elementName)
    {
        XElement? cache = series
            .Elements(ChartNamespace + elementName)
            .Descendants()
            .FirstOrDefault(child => child.Name.Namespace == ChartNamespace &&
                (child.Name.LocalName == "numLit" ||
                 child.Name.LocalName == "numCache"));
        return (string?)cache?.Element(ChartNamespace + "formatCode");
    }

    private static IReadOnlyList<double> ReadChartSeriesNumbers(XElement series, string elementName)
    {
        return series
            .Elements(ChartNamespace + elementName)
            .Descendants(ChartNamespace + "pt")
            .Select(point => (string?)point.Element(ChartNamespace + "v"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : double.NaN)
            .Where(value => !double.IsNaN(value))
            .ToArray();
    }

    private static IReadOnlyList<string> ReadChartSeriesCategories(XElement series)
    {
        return series
            .Elements(ChartNamespace + "cat")
            .Descendants(ChartNamespace + "pt")
            .Select(point => (string?)point.Element(ChartNamespace + "v"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value!)
            .ToArray();
    }

    private static IReadOnlyList<PptxSceneChartStringPoint> ReadChartSeriesStringPoints(XElement series, string elementName)
    {
        return ReadChartStringPoints(series
            .Elements(ChartNamespace + elementName)
            .Descendants(ChartNamespace + "pt"));
    }

    private static IReadOnlyList<IReadOnlyList<PptxSceneChartStringPoint>> ReadChartSeriesStringLevels(XElement series, string elementName)
    {
        return series
            .Elements(ChartNamespace + elementName)
            .Descendants(ChartNamespace + "lvl")
            .Select(level => ReadChartStringPoints(level.Elements(ChartNamespace + "pt")))
            .Where(points => points.Count != 0)
            .ToArray();
    }

    private static IReadOnlyList<PptxSceneChartStringPoint> ReadChartStringPoints(IEnumerable<XElement> sourcePoints)
    {
        var points = new List<PptxSceneChartStringPoint>();
        int ordinal = 0;
        foreach (XElement point in sourcePoints)
        {
            string indexValue = (string?)point.Attribute("idx") ?? string.Empty;
            bool hasParsedIndex = int.TryParse(indexValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedIndex);
            int index = hasParsedIndex
                ? parsedIndex
                : ordinal;
            XElement? valueElement = point.Element(ChartNamespace + "v");
            points.Add(new PptxSceneChartStringPoint(index, indexValue, hasParsedIndex, valueElement?.Value ?? string.Empty, valueElement is not null));
            ordinal++;
        }

        return points;
    }

    private static IReadOnlyList<PptxSceneChartAxis> ReadChartAxes(XDocument? chartXml, PptxTheme theme, PptxSceneChartStyle stylePart)
    {
        XElement? plotArea = chartXml?
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault();
        if (plotArea is null)
        {
            return [];
        }

        var axes = new List<PptxSceneChartAxis>();
        foreach (XElement axis in plotArea.Elements().Where(element => element.Name.Namespace == ChartNamespace && element.Name.LocalName.EndsWith("Ax", StringComparison.Ordinal)))
        {
            string id = (string?)axis.Element(ChartNamespace + "axId")?.Attribute("val") ?? string.Empty;
            if (id.Length == 0)
            {
                continue;
            }

            string axisPosition = (string?)axis.Element(ChartNamespace + "axPos")?.Attribute("val") ?? string.Empty;
            string crosses = (string?)axis.Element(ChartNamespace + "crosses")?.Attribute("val") ?? string.Empty;
            string crossBetween = (string?)axis.Element(ChartNamespace + "crossBetween")?.Attribute("val") ?? string.Empty;
            string orientation = (string?)axis.Element(ChartNamespace + "scaling")?.Element(ChartNamespace + "orientation")?.Attribute("val") ?? string.Empty;
            PptxSceneChartAxisOrientation orientationKind = ParseChartAxisOrientation(orientation);
            string tickLabelPosition = (string?)axis.Element(ChartNamespace + "tickLblPos")?.Attribute("val") ?? string.Empty;
            string majorTickMark = ReadChartElementValue(axis, "majorTickMark");
            string minorTickMark = ReadChartElementValue(axis, "minorTickMark");
            (double? crossesAt, string crossesAtValue) = ReadChartElementDoubleWithValue(axis, "crossesAt");
            (double? minimum, string minimumValue) = ReadChartAxisScalingValueWithValue(axis, "min");
            (double? maximum, string maximumValue) = ReadChartAxisScalingValueWithValue(axis, "max");
            (double? majorUnit, string majorUnitValue) = ReadChartAxisUnitValueWithValue(axis, "majorUnit");
            (double? minorUnit, string minorUnitValue) = ReadChartAxisUnitValueWithValue(axis, "minorUnit");
            (int? labelOffset, string labelOffsetValue) = ReadChartElementIntWithValue(axis, "lblOffset");
            (int? tickLabelSkip, string tickLabelSkipValue) = ReadChartElementIntWithValue(axis, "tickLblSkip");
            (int? tickMarkSkip, string tickMarkSkipValue) = ReadChartElementIntWithValue(axis, "tickMarkSkip");
            (bool? isDeleted, string isDeletedValue) = ReadOptionalOoxmlBooleanElementWithValue(axis, "delete");
            (bool? noMultiLevelLabels, string noMultiLevelLabelsValue) = ReadOptionalOoxmlBooleanElementWithValue(axis, "noMultiLvlLbl");
            XElement? majorGridlines = axis.Element(ChartNamespace + "majorGridlines");
            XElement? minorGridlines = axis.Element(ChartNamespace + "minorGridlines");
            axes.Add(new PptxSceneChartAxis(
                id,
                ParseChartAxisKind(axis.Name.LocalName),
                axis.Name.LocalName,
                ParseChartAxisPosition(axisPosition),
                axisPosition,
                (string?)axis.Element(ChartNamespace + "crossAx")?.Attribute("val") ?? string.Empty,
                ParseChartAxisCrosses(crosses),
                crosses,
                crossesAt,
                crossesAtValue,
                ParseChartAxisCrossBetween(crossBetween),
                crossBetween,
                orientationKind,
                orientation,
                orientationKind == PptxSceneChartAxisOrientation.MaximumMinimum,
                isDeleted,
                isDeletedValue,
                axis.Element(ChartNamespace + "scaling") is not null,
                minimum,
                minimumValue,
                maximum,
                maximumValue,
                majorUnit,
                majorUnitValue,
                minorUnit,
                minorUnitValue,
                IsChartGridlineVisible(majorGridlines),
                IsChartGridlineVisible(minorGridlines),
                majorGridlines is not null,
                minorGridlines is not null,
                ReadChartAxisLine(axis, theme),
                ReadChartGridlineLine(majorGridlines, theme),
                ReadChartGridlineLine(minorGridlines, theme),
                ReadChartStyleRoleLine(stylePart, "gridlineMajor"),
                ReadChartStyleRoleLine(stylePart, "gridlineMinor"),
                ReadChartTextStyleOverride(axis, theme),
                ParseChartTickLabelPosition(tickLabelPosition),
                tickLabelPosition,
                ParseChartAxisTickMark(majorTickMark),
                majorTickMark,
                ParseChartAxisTickMark(minorTickMark),
                minorTickMark,
                labelOffset,
                labelOffsetValue,
                tickLabelSkip,
                tickLabelSkipValue,
                tickMarkSkip,
                tickMarkSkipValue,
                noMultiLevelLabels,
                noMultiLevelLabelsValue,
                ReadChartAxisNumberFormat(axis),
                ReadChartNumberFormat(axis),
                ReadChartTitleElement(axis.Element(ChartNamespace + "title"), theme)));
        }

        return axes;
    }

    private static PptxSceneLineStyle ReadChartStyleRoleLine(PptxSceneChartStyle stylePart, string role)
    {
        PptxSceneChartStyleEntry entry = stylePart.Entries.FirstOrDefault(item => item.Role == role);
        if (entry.ShapeLine.HasLine)
        {
            return entry.ShapeLine;
        }

        return entry.Line;
    }

    private static PptxSceneLineStyle ReadChartAxisLine(XElement axis, PptxTheme theme)
    {
        XElement? shapeProperties = axis.Element(ChartNamespace + "spPr");
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "noFill") is not null)
        {
            return new PptxSceneLineStyle(true, new RgbColor(0, 0, 0), 0d, 0d, [], null, null, null, null, null, null, null);
        }

        return ReadChartLine(shapeProperties, theme);
    }

    private static PptxSceneChartTextStyleOverride ReadChartTextStyleOverride(XElement? parent, PptxTheme theme)
    {
        XElement? defaultRunProperties = parent?
            .Element(ChartNamespace + "txPr")?
            .Elements(DrawingNamespace + "p")
            .Select(paragraph => paragraph.Element(DrawingNamespace + "pPr")?.Element(DrawingNamespace + "defRPr"))
            .FirstOrDefault(element => element is not null);
        if (defaultRunProperties is null)
        {
            return default;
        }

        string? typeface = (string?)defaultRunProperties.Element(DrawingNamespace + "latin")?.Attribute("typeface") ??
            (string?)defaultRunProperties.Element(DrawingNamespace + "ea")?.Attribute("typeface") ??
            (string?)defaultRunProperties.Element(DrawingNamespace + "cs")?.Attribute("typeface");
        string? fontFamily = string.IsNullOrWhiteSpace(typeface)
            ? null
            : theme.ResolveTypeface(typeface);
        double? fontSize = defaultRunProperties.Attribute("sz") is { } sizeAttribute &&
            int.TryParse(sizeAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizeHundredths) &&
            sizeHundredths > 0
                ? sizeHundredths / 100d
                : null;
        RgbColor? color = TryReadSolidColorWithAlpha(defaultRunProperties.Element(DrawingNamespace + "solidFill"), theme, out RgbColor parsedColor, out double alpha)
            ? parsedColor
            : null;
        bool? bold = ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "b");
        bool? italic = ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "i");
        bool? underline = ReadChartUnderline(defaultRunProperties);
        bool? strike = ReadChartStrike(defaultRunProperties);
        return new PptxSceneChartTextStyleOverride(fontFamily, fontSize, color, color is null ? null : alpha, bold, italic, underline, strike);
    }

    private static PptxSceneChartTextBodyProperties ReadChartTextBodyProperties(XElement? parent)
    {
        string rotation = (string?)parent?
            .Element(ChartNamespace + "txPr")?
            .Element(DrawingNamespace + "bodyPr")?
            .Attribute("rot") ?? string.Empty;
        return new PptxSceneChartTextBodyProperties(ParseOptionalOoxmlAngle(rotation), rotation);
    }

    private static double? ParseOptionalOoxmlAngle(string value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long rawAngle)
            ? rawAngle / 60000d
            : null;
    }

    private static PptxSceneChartTextStyleOverride ReadChartTextRunStyle(XElement? runProperties, PptxTheme theme)
    {
        if (runProperties is null)
        {
            return default;
        }

        string? typeface = (string?)runProperties.Element(DrawingNamespace + "latin")?.Attribute("typeface") ??
            (string?)runProperties.Element(DrawingNamespace + "ea")?.Attribute("typeface") ??
            (string?)runProperties.Element(DrawingNamespace + "cs")?.Attribute("typeface");
        string? fontFamily = string.IsNullOrWhiteSpace(typeface)
            ? null
            : theme.ResolveTypeface(typeface);
        double? fontSize = runProperties.Attribute("sz") is { } sizeAttribute &&
            int.TryParse(sizeAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizeHundredths) &&
            sizeHundredths > 0
                ? sizeHundredths / 100d
                : null;
        RgbColor? color = TryReadSolidColorWithAlpha(runProperties.Element(DrawingNamespace + "solidFill"), theme, out RgbColor parsedColor, out double alpha)
            ? parsedColor
            : null;
        bool? bold = ReadOptionalOoxmlBooleanAttribute(runProperties, "b");
        bool? italic = ReadOptionalOoxmlBooleanAttribute(runProperties, "i");
        bool? underline = ReadChartUnderline(runProperties);
        bool? strike = ReadChartStrike(runProperties);
        return new PptxSceneChartTextStyleOverride(fontFamily, fontSize, color, color is null ? null : alpha, bold, italic, underline, strike);
    }

    private static bool? ReadChartUnderline(XElement runProperties)
    {
        string? value = (string?)runProperties.Attribute("u");
        return value is null ? null : !value.Equals("none", StringComparison.OrdinalIgnoreCase);
    }

    private static bool? ReadChartStrike(XElement runProperties)
    {
        string? value = (string?)runProperties.Attribute("strike");
        return value is null ? null : !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }

    private static PptxSceneChartManualLayout ReadChartPlotAreaManualLayout(XDocument? chartXml)
    {
        XElement? plotArea = chartXml?
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault();
        return ReadChartManualLayout(plotArea);
    }

    private static PptxSceneChartManualLayout ReadChartManualLayout(XElement? container)
    {
        XElement? manualLayout = container
            ?.Element(ChartNamespace + "layout")
            ?.Element(ChartNamespace + "manualLayout");
        if (manualLayout is null)
        {
            return default;
        }

        (double? x, string xValue) = ReadChartManualLayoutFactorWithValue(manualLayout, "x");
        (double? y, string yValue) = ReadChartManualLayoutFactorWithValue(manualLayout, "y");
        (double? width, string widthValue) = ReadChartManualLayoutFactorWithValue(manualLayout, "w");
        (double? height, string heightValue) = ReadChartManualLayoutFactorWithValue(manualLayout, "h");
        string layoutTarget = ReadChartElementValue(manualLayout, "layoutTarget");
        string xMode = ReadChartElementValue(manualLayout, "xMode");
        string yMode = ReadChartElementValue(manualLayout, "yMode");
        string widthMode = ReadChartElementValue(manualLayout, "wMode");
        string heightMode = ReadChartElementValue(manualLayout, "hMode");
        return new PptxSceneChartManualLayout(
            true,
            x,
            xValue,
            y,
            yValue,
            width,
            widthValue,
            height,
            heightValue,
            ParseChartManualLayoutTarget(layoutTarget),
            layoutTarget,
            ParseChartManualLayoutMode(xMode),
            xMode,
            ParseChartManualLayoutMode(yMode),
            yMode,
            ParseChartManualLayoutMode(widthMode),
            widthMode,
            ParseChartManualLayoutMode(heightMode),
            heightMode);
    }

    private static double? ReadChartManualLayoutFactor(XElement manualLayout, string elementName)
    {
        (double? parsed, _) = ReadChartManualLayoutFactorWithValue(manualLayout, elementName);
        return parsed;
    }

    private static (double? Value, string RawValue) ReadChartManualLayoutFactorWithValue(XElement manualLayout, string elementName)
    {
        string value = (string?)manualLayout.Element(ChartNamespace + elementName)?.Attribute("val") ?? string.Empty;
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? (parsed, value)
            : (null, value);
    }

    private static PptxSceneLineStyle ReadChartGridlineLine(XElement? gridlines, PptxTheme theme)
    {
        XElement? shapeProperties = gridlines?.Element(ChartNamespace + "spPr");
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "noFill") is not null)
        {
            return new PptxSceneLineStyle(true, new RgbColor(0, 0, 0), 0d, 0d, [], null, null, null, null, null, null, null);
        }

        return ReadChartLine(shapeProperties, theme);
    }

    private static double? ReadChartAxisScalingValue(XElement axis, string elementName)
    {
        (double? parsed, _) = ReadChartAxisScalingValueWithValue(axis, elementName);
        return parsed;
    }

    private static (double? Value, string RawValue) ReadChartAxisScalingValueWithValue(XElement axis, string elementName)
    {
        string value = (string?)axis
            .Element(ChartNamespace + "scaling")
            ?.Element(ChartNamespace + elementName)
            ?.Attribute("val") ?? string.Empty;
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? (parsed, value)
            : (null, value);
    }

    private static double? ReadChartAxisUnitValue(XElement axis, string elementName)
    {
        (double? parsed, _) = ReadChartAxisUnitValueWithValue(axis, elementName);
        return parsed;
    }

    private static (double? Value, string RawValue) ReadChartAxisUnitValueWithValue(XElement axis, string elementName)
    {
        string value = (string?)axis.Element(ChartNamespace + elementName)?.Attribute("val") ?? string.Empty;
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) && parsed > 0d
            ? (parsed, value)
            : (null, value);
    }

    private static bool IsChartGridlineVisible(XElement? gridlines)
    {
        XElement? line = gridlines?
            .Element(ChartNamespace + "spPr")
            ?.Element(DrawingNamespace + "ln");
        return gridlines is not null && line?.Element(DrawingNamespace + "noFill") is null;
    }

    private static string? ReadChartAxisNumberFormat(XElement axis)
    {
        string? format = (string?)axis
            .Element(ChartNamespace + "numFmt")
            ?.Attribute("formatCode");
        return string.IsNullOrWhiteSpace(format) ? null : format;
    }

    private static PptxSceneChartTitle ReadChartTitle(XDocument? chartXml, PptxTheme theme, IReadOnlyList<PptxSceneChartPlot> plots)
    {
        XElement? chart = chartXml?
            .Descendants(ChartNamespace + "chart")
            .FirstOrDefault();
        if (chart is null)
        {
            return EmptyChartTitle(IsAutoDeleted: null);
        }

        (bool? isAutoDeleted, string isAutoDeletedValue) = ReadOptionalOoxmlBooleanElementWithValue(chart, "autoTitleDeleted");
        return ReadChartTitleElement(chart.Element(ChartNamespace + "title"), theme, isAutoDeleted, isAutoDeletedValue, plots);
    }

    private static PptxSceneChartTitle ReadChartTitleElement(XElement? title, PptxTheme theme, bool? isAutoDeleted = null, string isAutoDeletedValue = "", IReadOnlyList<PptxSceneChartPlot>? plots = null)
    {
        if (title is null)
        {
            if (isAutoDeleted != false || plots is null)
            {
                return EmptyChartTitle(isAutoDeleted, isAutoDeletedValue);
            }

            string? inferredText = InferAutoChartTitleText(plots);
            return string.IsNullOrWhiteSpace(inferredText)
                ? EmptyChartTitle(isAutoDeleted, isAutoDeletedValue)
                : new PptxSceneChartTitle(
                    inferredText,
                    [new PptxSceneChartTextRun(inferredText!, default)],
                    isAutoDeleted,
                    isAutoDeletedValue,
                    IsAutoGenerated: true,
                    Overlay: null,
                    OverlayValue: string.Empty,
                    default,
                    new PptxSceneChartShapeStyle(false, default, default, default, default, default, default, default, default),
                    default,
                    default);
        }

        (bool? overlay, string overlayValue) = ReadOptionalOoxmlBooleanElementWithValue(title, "overlay");
        XElement? textElement = title?.Element(ChartNamespace + "tx");
        string? text = ReadChartText(textElement);

        return new PptxSceneChartTitle(
            string.IsNullOrWhiteSpace(text) ? null : text,
            ReadChartTextRuns(textElement, theme),
            isAutoDeleted,
            isAutoDeletedValue,
            IsAutoGenerated: false,
            overlay,
            overlayValue,
            ReadChartManualLayout(title),
            ReadChartShapeStyle(title?.Element(ChartNamespace + "spPr"), theme),
            ReadChartTextBodyProperties(title),
            ReadChartTextStyleOverride(title, theme));
    }

    private static string? InferAutoChartTitleText(IReadOnlyList<PptxSceneChartPlot> plots)
    {
        IReadOnlyList<PptxSceneChartSeries> series = plots
            .SelectMany(plot => plot.Series)
            .ToArray();
        if (series.Count != 1)
        {
            return null;
        }

        string? name = series[0].Name?.Trim();
        return string.IsNullOrWhiteSpace(name) ? null : name;
    }

    private static PptxSceneChartTitle EmptyChartTitle(bool? IsAutoDeleted, string IsAutoDeletedValue = "")
    {
        return new PptxSceneChartTitle(
            null,
            [],
            IsAutoDeleted,
            IsAutoDeletedValue,
            IsAutoGenerated: false,
            Overlay: null,
            OverlayValue: string.Empty,
            default,
            new PptxSceneChartShapeStyle(false, default, default, default, default, default, default, default, default),
            default,
            default);
    }

    private static PptxSceneChartLegend ReadChartLegend(XDocument? chartXml, PptxTheme theme)
    {
        XElement? legend = chartXml?
            .Descendants(ChartNamespace + "legend")
            .FirstOrDefault();
        if (legend is null)
        {
            return new PptxSceneChartLegend(
                PptxSceneChartLegendPosition.Right,
                "r",
                Overlay: null,
                OverlayValue: string.Empty,
                IsDefined: false,
                IsDeleted: null,
                IsDeletedValue: string.Empty,
                default,
                new PptxSceneChartShapeStyle(false, default, default, default, default, default, default, default, default),
                default,
                default);
        }

        string position = (string?)legend.Element(ChartNamespace + "legendPos")?.Attribute("val") ?? "r";
        (bool? overlay, string overlayValue) = ReadOptionalOoxmlBooleanElementWithValue(legend, "overlay");
        (bool? isDeleted, string isDeletedValue) = ReadOptionalOoxmlBooleanElementWithValue(legend, "delete");
        return new PptxSceneChartLegend(
            ParseChartLegendPosition(position),
            position,
            overlay,
            overlayValue,
            IsDefined: true,
            isDeleted,
            isDeletedValue,
            ReadChartManualLayout(legend),
            ReadChartShapeStyle(legend.Element(ChartNamespace + "spPr"), theme),
            ReadChartTextBodyProperties(legend),
            ReadChartTextStyleOverride(legend, theme));
    }

    private static PptxSceneChartExternalData ReadChartExternalData(OoxPackage package, string chartPartName, XDocument? chartXml)
    {
        XElement? externalData = chartXml?.Root?.Element(ChartNamespace + "externalData");
        if (externalData is null)
        {
            return default;
        }

        string? relationshipId = (string?)externalData.Attribute(RelationshipsNamespace + "id");
        string? targetPartName = null;
        if (!string.IsNullOrWhiteSpace(relationshipId))
        {
            targetPartName = package.GetRelationships(chartPartName)
                .FirstOrDefault(relationship => !relationship.IsExternal &&
                    relationship.Id == relationshipId &&
                    relationship.Type == ChartExternalDataPackageRelationshipType)
                ?.ResolvedTarget;
        }

        (bool? autoUpdate, string autoUpdateValue) = ReadOptionalOoxmlBooleanElementWithValue(externalData, "autoUpdate");
        return new PptxSceneChartExternalData(
            true,
            relationshipId,
            targetPartName,
            ReadPackageResource(package, targetPartName),
            autoUpdate,
            autoUpdateValue);
    }

    private static PptxSceneChartColorStyle ReadChartColorStyle(OoxPackage package, string chartPartName, PptxTheme theme)
    {
        OoxRelationship? colorRelationship = package.GetRelationships(chartPartName)
            .FirstOrDefault(relationship => !relationship.IsExternal &&
                relationship.Type == ChartColorStyleRelationshipType &&
                relationship.ResolvedTarget is not null);
        if (colorRelationship?.ResolvedTarget is null)
        {
            return new PptxSceneChartColorStyle(false, null, string.Empty, string.Empty, [], 0, [], null);
        }

        OoxPart? colorPart = package.GetPart(colorRelationship.ResolvedTarget);
        if (colorPart is null)
        {
            return new PptxSceneChartColorStyle(false, colorRelationship.ResolvedTarget, string.Empty, string.Empty, [], 0, [], null);
        }

        XDocument document = LoadXml(colorPart);
        IReadOnlyList<PptxSceneChartColorDeclaration> declarations = ReadChartColorStyleDeclarations(document, theme);
        var colors = new List<RgbColor>();
        foreach (XElement colorElement in document.Root?.Elements().Where(element => element.Name.Namespace == DrawingNamespace) ?? [])
        {
            var wrapper = new XElement(DrawingNamespace + "solidFill", new XElement(colorElement));
            if (TryReadSolidColorWithAlpha(wrapper, theme, out RgbColor color, out _))
            {
                colors.Add(color);
            }
        }

        return new PptxSceneChartColorStyle(
            true,
            colorPart.Name,
            (string?)document.Root?.Attribute("meth") ?? string.Empty,
            (string?)document.Root?.Attribute("id") ?? string.Empty,
            colors,
            CountChartColorStyleVariations(document),
            declarations,
            document);
    }

    private static IReadOnlyList<PptxSceneChartColorDeclaration> ReadChartColorStyleDeclarations(XDocument document, PptxTheme theme)
    {
        if (document.Root is null)
        {
            return [];
        }

        var declarations = new List<PptxSceneChartColorDeclaration>();
        foreach (XElement colorElement in document.Root.Elements().Where(IsDrawingColorElement))
        {
            declarations.Add(ReadChartColorStyleDeclaration(colorElement, theme, variationIndex: null));
        }

        int variationIndex = 0;
        foreach (XElement variation in document.Root.Elements().Where(IsChartColorStyleVariationElement))
        {
            foreach (XElement colorElement in variation.Descendants().Where(IsDrawingColorElement))
            {
                declarations.Add(ReadChartColorStyleDeclaration(colorElement, theme, variationIndex));
            }

            variationIndex++;
        }

        return declarations;
    }

    private static PptxSceneChartColorDeclaration ReadChartColorStyleDeclaration(XElement colorElement, PptxTheme theme, int? variationIndex)
    {
        var wrapper = new XElement(DrawingNamespace + "solidFill", new XElement(colorElement));
        bool isResolved = TryReadSolidColorWithAlpha(wrapper, theme, out RgbColor color, out double alpha);
        return new PptxSceneChartColorDeclaration(
            colorElement.Name.LocalName,
            (string?)colorElement.Attribute("val") ?? string.Empty,
            variationIndex,
            isResolved,
            isResolved ? color : null,
            isResolved ? alpha : 1d);
    }

    private static int CountChartColorStyleVariations(XDocument document)
    {
        return document.Root?.Elements().Count(IsChartColorStyleVariationElement) ?? 0;
    }

    private static bool IsChartColorStyleVariationElement(XElement element)
    {
        return element.Name.LocalName == "variation";
    }

    private static bool IsDrawingColorElement(XElement element)
    {
        if (element.Name.Namespace != DrawingNamespace)
        {
            return false;
        }

        return element.Name.LocalName is "srgbClr" or "schemeClr" or "scrgbClr" or "prstClr" or "sysClr" or "hslClr";
    }

    private static PptxSceneChartStyle ReadChartStylePart(OoxPackage package, string chartPartName, PptxTheme theme)
    {
        OoxRelationship? styleRelationship = package.GetRelationships(chartPartName)
            .FirstOrDefault(relationship => !relationship.IsExternal &&
                relationship.Type == ChartStyleRelationshipType &&
                relationship.ResolvedTarget is not null);
        if (styleRelationship?.ResolvedTarget is null)
        {
            return new PptxSceneChartStyle(false, null, string.Empty, null, []);
        }

        OoxPart? stylePart = package.GetPart(styleRelationship.ResolvedTarget);
        if (stylePart is null)
        {
            return new PptxSceneChartStyle(false, styleRelationship.ResolvedTarget, string.Empty, null, []);
        }

        XDocument document = LoadXml(stylePart);
        return new PptxSceneChartStyle(
            true,
            stylePart.Name,
            (string?)document.Root?.Attribute("id") ?? string.Empty,
            document,
            ReadChartStyleEntries(document, theme));
    }

    private static IReadOnlyList<PptxSceneChartStyleEntry> ReadChartStyleEntries(XDocument document, PptxTheme theme)
    {
        if (document.Root is null)
        {
            return [];
        }

        var entries = new List<PptxSceneChartStyleEntry>();
        foreach (XElement roleElement in document.Root.Elements())
        {
            XElement? lineReference = roleElement
                .Descendants()
                .FirstOrDefault(element => element.Name.LocalName == "lnRef");
            int lineReferenceIndexValue = lineReference is null ? 0 : ParseOptionalIntAttribute(lineReference, "idx", 0);
            int? lineReferenceIndex = lineReferenceIndexValue > 0 ? lineReferenceIndexValue : null;
            XElement? fillReference = roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "fillRef");
            int fillReferenceIndexValue = fillReference is null ? 0 : ParseOptionalIntAttribute(fillReference, "idx", 0);
            int? fillReferenceIndex = fillReferenceIndexValue > 0 ? fillReferenceIndexValue : null;
            XElement? effectReference = roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "effectRef");
            int effectReferenceIndexValue = effectReference is null ? 0 : ParseOptionalIntAttribute(effectReference, "idx", 0);
            int? effectReferenceIndex = effectReferenceIndexValue > 0 ? effectReferenceIndexValue : null;
            string fontReferenceIndex = (string?)roleElement
                .Elements()
                .FirstOrDefault(element => element.Name.LocalName == "fontRef")
                ?.Attribute("idx") ?? string.Empty;
            PptxSceneLineStyle line = lineReferenceIndex is not null &&
                lineReference is not null &&
                TryReadThemeLineReference(lineReference, theme, out PptxSceneLineStyle resolvedLine)
                    ? resolvedLine
                    : default;
            PptxSceneLineStyle shapeLine = ReadChartLine(
                roleElement
                    .Elements()
                    .FirstOrDefault(element => element.Name.LocalName == "spPr"),
                theme);
            PptxSceneChartShapeStyle shapeStyle = ReadChartShapeStyle(
                roleElement
                    .Elements()
                    .FirstOrDefault(element => element.Name.LocalName == "spPr"),
                theme);
            PptxSceneChartTextStyleOverride textStyle = ReadChartStyleRoleTextStyle(roleElement, theme);
            if (lineReference is null &&
                fillReference is null &&
                effectReference is null &&
                string.IsNullOrWhiteSpace(fontReferenceIndex) &&
                !HasChartShapeStyle(shapeStyle) &&
                !shapeLine.HasLine &&
                !HasChartTextStyleOverride(textStyle))
            {
                continue;
            }

            entries.Add(new PptxSceneChartStyleEntry(
                roleElement.Name.LocalName,
                lineReferenceIndex,
                fillReferenceIndex,
                effectReferenceIndex,
                fontReferenceIndex,
                line,
                shapeStyle,
                shapeLine,
                textStyle));
        }

        return entries;
    }

    private static bool HasChartShapeStyle(PptxSceneChartShapeStyle style)
    {
        return style.NoFill ||
            style.Fill.HasFill ||
            style.GradientFill?.HasGradient == true ||
            style.PatternFill.HasPattern ||
            style.PictureFill.HasPicture ||
            style.Line.HasLine ||
            style.Glow.HasGlow ||
            style.OuterShadow.HasShadow ||
            style.Effects.HasEffectDag ||
            style.Effects.UnsupportedEffectNames?.Count > 0;
    }

    private static bool HasChartTextStyleOverride(PptxSceneChartTextStyleOverride textStyle)
    {
        return textStyle.FontFamily is not null ||
            textStyle.FontSize is not null ||
            textStyle.Color is not null ||
            textStyle.Alpha is not null ||
            textStyle.Bold is not null ||
            textStyle.Italic is not null ||
            textStyle.Underline is not null ||
            textStyle.Strike is not null;
    }

    private static PptxSceneChartTextStyleOverride ReadChartStyleRoleTextStyle(XElement roleElement, PptxTheme theme)
    {
        XElement? defaultRunProperties = roleElement
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName == "defRPr");
        XElement? fontReference = roleElement
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName == "fontRef");
        string? typeface = (string?)defaultRunProperties?.Element(DrawingNamespace + "latin")?.Attribute("typeface") ??
            (string?)defaultRunProperties?.Element(DrawingNamespace + "ea")?.Attribute("typeface") ??
            (string?)defaultRunProperties?.Element(DrawingNamespace + "cs")?.Attribute("typeface");
        if (string.IsNullOrWhiteSpace(typeface))
        {
            typeface = (string?)fontReference?.Attribute("idx") switch
            {
                "major" => "+mj-lt",
                "minor" => "+mn-lt",
                _ => null
            };
        }

        string? fontFamily = string.IsNullOrWhiteSpace(typeface)
            ? null
            : theme.ResolveTypeface(typeface);
        double? fontSize = defaultRunProperties?.Attribute("sz") is { } sizeAttribute &&
            int.TryParse(sizeAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizeHundredths) &&
            sizeHundredths > 0
                ? sizeHundredths / 100d
                : null;
        RgbColor? color = TryReadSolidColorWithAlpha(defaultRunProperties?.Element(DrawingNamespace + "solidFill"), theme, out RgbColor parsedColor, out double alpha) ||
            TryReadSolidColorWithAlpha(fontReference, theme, out parsedColor, out alpha)
                ? parsedColor
                : null;
        bool? bold = defaultRunProperties is null ? null : ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "b");
        bool? italic = defaultRunProperties is null ? null : ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "i");
        bool? underline = defaultRunProperties is null ? null : ReadChartUnderline(defaultRunProperties);
        bool? strike = defaultRunProperties is null ? null : ReadChartStrike(defaultRunProperties);
        return new PptxSceneChartTextStyleOverride(fontFamily, fontSize, color, color is null ? null : alpha, bold, italic, underline, strike);
    }

    private static bool IsOoxmlBooleanElementEnabled(XElement? element)
    {
        return OoxBoolean.ParseElement(element);
    }

    private static bool IsOoxmlBooleanElementEnabled(XElement? element, bool defaultValue)
    {
        return OoxBoolean.ParseElement(element, defaultValue);
    }

    private static bool? ReadOptionalOoxmlBooleanAttribute(XElement element, string attributeName)
    {
        return OoxBoolean.ParseOptionalAttribute(element, attributeName);
    }

    private static PptxSceneTable ReadTable(XElement frame, PptxTheme theme)
    {
        XElement? table = ReadTableElement(frame);
        IReadOnlyList<double> columnWidths = ReadTableColumnWidths(table);
        IReadOnlyList<double> rowHeights = ReadTableRowHeights(table);
        PptxSceneTableStyle style = ReadTableStyle(table);
        return new PptxSceneTable(
            columnWidths,
            rowHeights,
            ReadTableRows(table, theme, style, rowHeights.Count, columnWidths.Count),
            style,
            table);
    }

    internal static XElement? ReadTableElement(XElement frame)
    {
        return frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData")
            ?.Element(DrawingNamespace + "tbl");
    }

    internal static IReadOnlyList<double> ReadTableColumnWidths(XElement? table)
    {
        return table
            ?.Element(DrawingNamespace + "tblGrid")
            ?.Elements(DrawingNamespace + "gridCol")
            .Select(column => Math.Max(1d, ReadLong(column, "w", 1)))
            .ToArray() ?? [];
    }

    internal static IReadOnlyList<double> ReadTableRowHeights(XElement? table)
    {
        return table
            ?.Elements(DrawingNamespace + "tr")
            .Select(row => Math.Max(1d, ReadLong(row, "h", 1)))
            .ToArray() ?? [];
    }

    internal static IReadOnlyList<PptxSceneTableRow> ReadTableRows(XElement? table, PptxTheme theme, PptxSceneTableStyle tableStyle, int rowCount, int columnCount)
    {
        if (table is null)
        {
            return [];
        }

        var rows = new List<PptxSceneTableRow>();
        int rowIndex = 0;
        foreach (XElement row in table.Elements(DrawingNamespace + "tr"))
        {
            var cells = new List<PptxSceneTableCell>();
            int columnIndex = 0;
            foreach (XElement cell in row.Elements(DrawingNamespace + "tc"))
            {
                PptxSceneTableCell sceneCell = ReadTableCell(cell, theme, tableStyle, rowIndex, columnIndex, rowCount, columnCount);
                cells.Add(sceneCell);
                columnIndex += sceneCell.IsMergedContinuation ? 1 : sceneCell.ColumnSpan;
            }

            rows.Add(new PptxSceneTableRow(cells));
            rowIndex++;
        }

        return rows;
    }

    internal static PptxSceneTableStyle ReadTableStyle(XElement? table)
    {
        XElement? tableProperties = table?.Element(DrawingNamespace + "tblPr");
        string? styleId = (string?)tableProperties?.Element(DrawingNamespace + "tableStyleId");
        bool supported = PptxBuiltInTableStyles.TryGet(styleId, out PptxBuiltInTableStyle style);
        return new PptxSceneTableStyle(
            styleId,
            supported ? style.Name : string.Empty,
            supported ? style.Accent : string.Empty,
            supported,
            ReadTablePropertyFlag(tableProperties, "firstRow"),
            ReadTablePropertyFlagValue(tableProperties, "firstRow"),
            ReadTablePropertyFlag(tableProperties, "lastRow"),
            ReadTablePropertyFlagValue(tableProperties, "lastRow"),
            ReadTablePropertyFlag(tableProperties, "firstCol"),
            ReadTablePropertyFlagValue(tableProperties, "firstCol"),
            ReadTablePropertyFlag(tableProperties, "lastCol"),
            ReadTablePropertyFlagValue(tableProperties, "lastCol"),
            ReadTablePropertyFlag(tableProperties, "bandRow"),
            ReadTablePropertyFlagValue(tableProperties, "bandRow"),
            ReadTablePropertyFlag(tableProperties, "bandCol"),
            ReadTablePropertyFlagValue(tableProperties, "bandCol"));
    }

    private static bool ReadTablePropertyFlag(XElement? tableProperties, string name)
    {
        if (tableProperties is null)
        {
            return false;
        }

        if (tableProperties.Attribute(name) is { } attribute)
        {
            return OoxBoolean.IsTrue(attribute.Value);
        }

        return tableProperties.Element(DrawingNamespace + name) is not null;
    }

    private static string ReadTablePropertyFlagValue(XElement? tableProperties, string name)
    {
        return (string?)tableProperties?.Attribute(name) ?? string.Empty;
    }

    internal static PptxSceneTableCell ReadTableCell(XElement cell, PptxTheme theme, PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount)
    {
        return new PptxSceneTableCell(
            ReadTableCellColumnSpan(cell),
            ReadTableCellRowSpan(cell),
            IsMergedTableCellContinuation(cell),
            ReadTableCellTextInsets(cell),
            ReadTableCellVerticalAnchor(cell),
            ReadTableCellVerticalAnchorValue(cell),
            ReadTableCellFill(cell, theme),
            ReadTableCellBorders(cell, theme),
            PptxTableStyleResolver.ReadCellFill(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme),
            PptxTableStyleResolver.ReadCellTextStyle(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme),
            cell.Element(DrawingNamespace + "txBody"));
    }

    internal static bool IsMergedTableCellContinuation(XElement cell)
    {
        return ReadBool(cell, "hMerge") || ReadBool(cell, "vMerge");
    }

    internal static int ReadTableCellColumnSpan(XElement cell)
    {
        return ReadTableCellSpan(cell, "gridSpan");
    }

    internal static int ReadTableCellRowSpan(XElement cell)
    {
        return ReadTableCellSpan(cell, "rowSpan");
    }

    private static int ReadTableCellSpan(XElement cell, string attributeName)
    {
        return cell.Attribute(attributeName) is { } spanAttribute &&
            int.TryParse(spanAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            ? Math.Max(1, span)
            : 1;
    }

    internal static PptxSceneTextInsets ReadTableCellTextInsets(XElement cell)
    {
        XElement? textBody = cell.Element(DrawingNamespace + "txBody");
        XElement? bodyProperties = textBody?.Element(DrawingNamespace + "bodyPr");
        PptxSceneTextInsets bodyInsets = new(
            ReadInset(bodyProperties, "lIns", 91440),
            ReadInset(bodyProperties, "rIns", 91440),
            ReadInset(bodyProperties, "tIns", 45720),
            ReadInset(bodyProperties, "bIns", 45720));
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        return new PptxSceneTextInsets(
            ReadTableCellMargin(cellProperties, "marL", bodyInsets.Left),
            ReadTableCellMargin(cellProperties, "marR", bodyInsets.Right),
            ReadTableCellMargin(cellProperties, "marT", bodyInsets.Top),
            ReadTableCellMargin(cellProperties, "marB", bodyInsets.Bottom));
    }

    private static double ReadTableCellMargin(XElement? cellProperties, string attributeName, double fallback)
    {
        return cellProperties?.Attribute(attributeName) is { } margin &&
            long.TryParse(margin.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emus)
                ? OoxUnits.EmuToPoints(emus)
                : fallback;
    }

    private static double ReadInset(XElement? element, string attributeName, long defaultEmu)
    {
        long emu = element?.Attribute(attributeName) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : defaultEmu;
        return OoxUnits.EmuToPoints(emu);
    }

    internal static PptxSceneTableCellVerticalAnchor ReadTableCellVerticalAnchor(XElement cell)
    {
        string? anchor = ReadTableCellVerticalAnchorValue(cell);
        return anchor switch
        {
            "ctr" => PptxSceneTableCellVerticalAnchor.Middle,
            "b" => PptxSceneTableCellVerticalAnchor.Bottom,
            _ => PptxSceneTableCellVerticalAnchor.Top
        };
    }

    internal static string? ReadTableCellVerticalAnchorValue(XElement cell)
    {
        return (string?)cell
            .Element(DrawingNamespace + "tcPr")
            ?.Attribute("anchor");
    }

    internal static PptxSceneFillStyle ReadTableCellFill(XElement cell, PptxTheme theme)
    {
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        return TryReadSolidColorWithAlpha(cellProperties, theme, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    internal static PptxSceneTableCellBorders ReadTableCellBorders(XElement cell, PptxTheme theme)
    {
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        return new PptxSceneTableCellBorders(
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnL"), theme),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnR"), theme),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnT"), theme),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnB"), theme));
    }

    private static PptxSceneTableCellBorder ReadTableCellBorder(XElement? line, PptxTheme theme)
    {
        if (line is null)
        {
            return default;
        }

        if (line.Element(DrawingNamespace + "noFill") is not null ||
            !TryReadSolidColorWithAlpha(line, theme, out RgbColor color, out double alpha))
        {
            return new PptxSceneTableCellBorder(IsSpecified: true, default);
        }

        double lineWidth = line.Attribute("w") is { } widthAttribute
            ? Math.Max(1d, OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture)) / 2d)
            : 0.75d;
        XElement shapeProperties = WrapTableCellBorderLine(line);
        IReadOnlyList<double> dashPattern = TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> parsedDashPattern)
            ? parsedDashPattern
            : [];
        string? lineCap = ReadLineCap(shapeProperties);
        return new PptxSceneTableCellBorder(
            IsSpecified: true,
            new PptxSceneLineStyle(
                true,
                color,
                lineWidth,
                alpha,
                dashPattern,
                ReadPresetDashValue(shapeProperties),
                ReadLineCompound(shapeProperties),
                ReadLineCompoundValue(shapeProperties),
                lineCap switch
                {
                    "rnd" => 1,
                    "sq" => 2,
                    _ => null
                },
                lineCap,
                ReadLineJoin(shapeProperties),
                ReadLineJoinValue(shapeProperties)));
    }

    private static XElement WrapTableCellBorderLine(XElement line)
    {
        return new XElement(
            DrawingNamespace + "spPr",
            new XElement(DrawingNamespace + "ln", line.Attributes(), line.Nodes()));
    }

    internal static PptxSceneGroupTransform ReadGroupTransform(XElement group)
    {
        XElement? transform = group
            .Element(PresentationNamespace + "grpSpPr")
            ?.Element(DrawingNamespace + "xfrm");
        XElement? offset = transform?.Element(DrawingNamespace + "off");
        XElement? extents = transform?.Element(DrawingNamespace + "ext");
        XElement? childOffset = transform?.Element(DrawingNamespace + "chOff");
        XElement? childExtents = transform?.Element(DrawingNamespace + "chExt");
        if (offset is null || extents is null || childOffset is null || childExtents is null)
        {
            return PptxSceneGroupTransform.Identity;
        }

        long width = ReadLong(extents, "cx");
        long height = ReadLong(extents, "cy");
        long childWidth = Math.Max(1, ReadLong(childExtents, "cx"));
        long childHeight = Math.Max(1, ReadLong(childExtents, "cy"));
        return new PptxSceneGroupTransform(
            ReadLong(offset, "x"),
            ReadLong(offset, "y"),
            width,
            height,
            ReadLong(childOffset, "x"),
            ReadLong(childOffset, "y"),
            width / (double)childWidth,
            height / (double)childHeight,
            transform!.Attribute("rot") is { } rotation ? long.Parse(rotation.Value, CultureInfo.InvariantCulture) / 60000d : 0d,
            ReadBool(transform, "flipH"),
            ReadBool(transform, "flipV"));
    }

    private static PptxSceneShape ReadShape(
        XElement shape,
        PptxTheme theme,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        PptxSceneLineStyle line = TryReadShapeLine(shape, shapeProperties, theme, out RgbColor lineColor, out double lineWidth, out double lineAlpha)
            ? new PptxSceneLineStyle(
                true,
                lineColor,
                lineWidth,
                lineAlpha,
                TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> dashPattern) ? dashPattern : [],
                ReadPresetDashValue(shapeProperties),
                ReadLineCompound(shapeProperties),
                ReadLineCompoundValue(shapeProperties),
                ReadLineCap(shapeProperties) switch
                {
                    "rnd" => 1,
                    "sq" => 2,
                    _ => null
                },
                ReadLineCap(shapeProperties),
                ReadLineJoin(shapeProperties),
                ReadLineJoinValue(shapeProperties))
            : default;
        return new PptxSceneShape(
            ReadShapePreset(shapeProperties),
            ReadPresetAdjustments(shapeProperties),
            shapeProperties?.Element(DrawingNamespace + "custGeom") is not null,
            ReadCustomGeometry(shapeProperties),
            TryReadShapeFill(shape, shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
                ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
                : default,
            TryReadShapeGradientFill(shapeProperties, theme, out PptxSceneGradientFill gradientFill) ? gradientFill : new PptxSceneGradientFill(false, 0d, []),
            TryReadShapePatternFill(shapeProperties, theme, out PptxScenePatternFill patternFill) ? patternFill : default,
            ReadShapePictureFill(shapeProperties, package, relationships),
            TryReadGlow(shapeProperties, theme, out PptxSceneGlow glow) ? glow : default,
            TryReadOuterShadow(shapeProperties, theme, out PptxSceneOuterShadow outerShadow) ? outerShadow : default,
            line,
            ReadLineEnd(shapeProperties, "headEnd"),
            ReadLineEnd(shapeProperties, "tailEnd"));
    }

    internal static bool TryReadShapeGradientFill(XElement? shapeProperties, PptxTheme theme, out PptxSceneGradientFill fill)
    {
        XElement? gradientFill = shapeProperties?.Element(DrawingNamespace + "gradFill");
        XElement? gradientStopList = gradientFill?.Element(DrawingNamespace + "gsLst");
        if (gradientFill is null || gradientStopList is null)
        {
            fill = new PptxSceneGradientFill(false, 0d, []);
            return false;
        }

        PptxSceneGradientStop[] stops = gradientStopList
            .Elements(DrawingNamespace + "gs")
            .Select(stop => TryReadGradientStop(stop, theme, out PptxSceneGradientStop parsed) ? parsed : (PptxSceneGradientStop?)null)
            .Where(stop => stop is not null)
            .Select(stop => stop!.Value)
            .OrderBy(stop => stop.Offset)
            .ToArray();
        if (stops.Length < 2 ||
            gradientFill.Element(DrawingNamespace + "lin") is not { } linear)
        {
            fill = new PptxSceneGradientFill(false, 0d, []);
            return false;
        }

        double angleDegrees = ReadLong(linear, "ang", 0) / 60000d;
        fill = new PptxSceneGradientFill(true, angleDegrees, stops);
        return true;
    }

    private static bool TryReadGradientStop(XElement gradientStop, PptxTheme theme, out PptxSceneGradientStop stop)
    {
        if (!TryReadSolidColorWithAlpha(gradientStop, theme, out RgbColor color, out double alpha))
        {
            stop = default;
            return false;
        }

        stop = new PptxSceneGradientStop(ParseGradientPercentage(gradientStop, "pos"), color, alpha);
        return true;
    }

    private static double ParseGradientPercentage(XElement element, string attribute)
    {
        return element.Attribute(attribute) is { } value
            ? Math.Clamp(int.Parse(value.Value, CultureInfo.InvariantCulture) / 100000d, 0d, 1d)
            : 0d;
    }

    private static IReadOnlyDictionary<string, double> ReadPresetAdjustments(XElement? shapeProperties)
    {
        Dictionary<string, double> adjustments = new(StringComparer.Ordinal);
        foreach (XElement guide in shapeProperties
                     ?.Element(DrawingNamespace + "prstGeom")
                     ?.Element(DrawingNamespace + "avLst")
                     ?.Elements(DrawingNamespace + "gd") ?? [])
        {
            string? name = (string?)guide.Attribute("name");
            string? formula = (string?)guide.Attribute("fmla");
            if (!string.IsNullOrWhiteSpace(name) &&
                formula is not null &&
                formula.StartsWith("val ", StringComparison.Ordinal) &&
                double.TryParse(formula[4..], NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
            {
                adjustments[name] = value;
            }
        }

        return adjustments;
    }

    private static PptxSceneCustomGeometry ReadCustomGeometry(XElement? shapeProperties)
    {
        XElement? customGeometry = shapeProperties?.Element(DrawingNamespace + "custGeom");
        if (customGeometry is null)
        {
            return new PptxSceneCustomGeometry(false, [], []);
        }

        IReadOnlyList<PptxSceneCustomGuide> guides = customGeometry
            .Element(DrawingNamespace + "gdLst")
            ?.Elements(DrawingNamespace + "gd")
            .Select(guide => new PptxSceneCustomGuide(
                (string?)guide.Attribute("name") ?? string.Empty,
                (string?)guide.Attribute("fmla") ?? string.Empty))
            .Where(guide => !string.IsNullOrWhiteSpace(guide.Name) && !string.IsNullOrWhiteSpace(guide.Formula))
            .ToArray() ?? [];

        IReadOnlyList<PptxSceneCustomPath> paths = customGeometry
            .Element(DrawingNamespace + "pathLst")
            ?.Elements(DrawingNamespace + "path")
            .Select(ReadCustomPath)
            .Where(path => path.Commands.Count > 0)
            .ToArray() ?? [];

        return new PptxSceneCustomGeometry(paths.Count > 0, guides, paths);
    }

    private static PptxSceneCustomPath ReadCustomPath(XElement path)
    {
        PptxSceneCustomCommand?[] commands = path.Elements().Select(ReadCustomCommand).ToArray();
        if (commands.Any(command => command is null))
        {
            return new PptxSceneCustomPath(0d, 0d, false, false, []);
        }

        return new PptxSceneCustomPath(
            ParseOptionalDoubleAttribute(path, "w", 21600d),
            ParseOptionalDoubleAttribute(path, "h", 21600d),
            !string.Equals((string?)path.Attribute("fill"), "none", StringComparison.Ordinal),
            ParseBoolAttribute(path, "stroke", defaultValue: true),
            commands
                .Cast<PptxSceneCustomCommand>()
                .ToArray());
    }

    private static PptxSceneCustomCommand? ReadCustomCommand(XElement command)
    {
        return command.Name.LocalName switch
        {
            "moveTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.MoveTo,
                ReadCustomPoints(command).Take(1).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "lnTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.LineTo,
                ReadCustomPoints(command).Take(1).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "cubicBezTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.CubicBezierTo,
                ReadCustomPoints(command).Take(3).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "quadBezTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.QuadraticBezierTo,
                ReadCustomPoints(command).Take(2).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "arcTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.ArcTo,
                [],
                (string?)command.Attribute("wR") ?? string.Empty,
                (string?)command.Attribute("hR") ?? string.Empty,
                (string?)command.Attribute("stAng") ?? string.Empty,
                (string?)command.Attribute("swAng") ?? string.Empty),
            "close" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.Close,
                [],
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            _ => null
        };
    }

    private static IReadOnlyList<PptxSceneCustomPoint> ReadCustomPoints(XElement command)
    {
        XElement[] points = command.Elements(DrawingNamespace + "pt").ToArray();
        if (points.Length == 0 && command.Name.LocalName is "moveTo" or "lnTo")
        {
            points = [command];
        }

        return points
            .Select(point => new PptxSceneCustomPoint(
                (string?)point.Attribute("x") ?? string.Empty,
                (string?)point.Attribute("y") ?? string.Empty))
            .ToArray();
    }

    private static PptxSceneShapePictureFill ReadShapePictureFill(
        XElement? shapeProperties,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        XElement? blipFill = shapeProperties?.Element(DrawingNamespace + "blipFill");
        XElement? blip = blipFill?.Element(DrawingNamespace + "blip");
        string? relationshipId = (string?)blip?.Attribute(RelationshipsNamespace + "embed");
        string? targetPartName = ResolveRelationshipTarget(relationshipId, relationships);
        return blipFill is null
            ? default
            : new PptxSceneShapePictureFill(
                true,
                relationshipId ?? string.Empty,
                targetPartName,
                ReadImageResource(package, targetPartName),
                ReadPictureCrop(shapeProperties!),
                ReadPictureFill(shapeProperties!),
                ReadPictureAlpha(shapeProperties!),
                ReadPictureAlphaValue(shapeProperties!),
                ReadPictureTile(shapeProperties!));
    }

    private static PptxSceneImageResource? ReadImageResource(OoxPackage package, string? targetPartName)
    {
        if (targetPartName is null)
        {
            return null;
        }

        OoxPart? imagePart = package.GetPart(targetPartName);
        return imagePart is null
            ? null
            : new PptxSceneImageResource(imagePart.Name, imagePart.ContentType, imagePart.Bytes);
    }

    private static PptxScenePackageResource? ReadPackageResource(OoxPackage package, string? targetPartName)
    {
        if (targetPartName is null)
        {
            return null;
        }

        OoxPart? packagePart = package.GetPart(targetPartName);
        return packagePart is null
            ? null
            : new PptxScenePackageResource(packagePart.Name, packagePart.ContentType, packagePart.Bytes);
    }

    private static string? ResolveRelationshipTarget(
        string? relationshipId,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        return relationshipId is not null &&
            relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) &&
            !relationship.IsExternal
            ? relationship.ResolvedTarget
            : null;
    }

    internal static bool TryReadGlow(XElement? shapeProperties, PptxTheme theme, out PptxSceneGlow glow)
    {
        XElement? glowElement = shapeProperties
            ?.Element(DrawingNamespace + "effectLst")
            ?.Element(DrawingNamespace + "glow");
        if (glowElement is null)
        {
            glow = default;
            return false;
        }

        XElement? colorElement = glowElement.Elements().FirstOrDefault(element =>
            element.Name.LocalName is "srgbClr" or "schemeClr" or "prstClr");
        if (colorElement is not null &&
            TryReadImageRecolorColor(colorElement, theme, out RgbColor color))
        {
            double radius = OoxUnits.EmuToPoints(ReadLong(glowElement, "rad", 0));
            glow = new PptxSceneGlow(
                true,
                color,
                ReadAlpha(new XElement(DrawingNamespace + "solidFill", new XElement(colorElement))),
                radius);
            return radius > SceneEffectTolerance;
        }

        glow = default;
        return false;
    }

    internal static bool TryReadOuterShadow(XElement? shapeProperties, PptxTheme theme, out PptxSceneOuterShadow shadow)
    {
        XElement? outerShadow = shapeProperties
            ?.Element(DrawingNamespace + "effectLst")
            ?.Element(DrawingNamespace + "outerShdw");
        if (outerShadow is null)
        {
            shadow = default;
            return false;
        }

        XElement? colorElement = outerShadow.Elements().FirstOrDefault(element =>
            element.Name.LocalName is "srgbClr" or "schemeClr" or "prstClr");
        if (colorElement is not null &&
            TryReadImageRecolorColor(colorElement, theme, out RgbColor color))
        {
            double alpha = ReadAlpha(new XElement(DrawingNamespace + "solidFill", new XElement(colorElement)));
            double distance = OoxUnits.EmuToPoints(ReadLong(outerShadow, "dist", 0));
            double direction = ReadLong(outerShadow, "dir", 0) / 60000d * Math.PI / 180d;
            shadow = new PptxSceneOuterShadow(
                true,
                color,
                alpha,
                distance * Math.Cos(direction),
                -distance * Math.Sin(direction));
            return true;
        }

        shadow = default;
        return false;
    }

    private static bool TryReadShapePatternFill(XElement? shapeProperties, PptxTheme theme, out PptxScenePatternFill fill)
    {
        XElement? patternFill = shapeProperties?.Element(DrawingNamespace + "pattFill");
        string? preset = (string?)patternFill?.Attribute("prst");
        if (patternFill is null || !IsSupportedDiagonalPatternFill(preset))
        {
            fill = default;
            return false;
        }

        RgbColor foreground = TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "fgClr"), theme, out RgbColor foregroundColor, out _)
            ? foregroundColor
            : new RgbColor(0, 0, 0);
        RgbColor background = TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "bgClr"), theme, out RgbColor backgroundColor, out _)
            ? backgroundColor
            : new RgbColor(255, 255, 255);
        fill = new PptxScenePatternFill(true, preset!, foreground, background, 1d);
        return true;
    }

    private static bool IsSupportedDiagonalPatternFill(string? preset)
    {
        return preset is not null &&
            (preset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase) ||
             preset.Contains("DnDiag", StringComparison.OrdinalIgnoreCase));
    }

    private static bool TryReadShapeFill(XElement shape, XElement? shapeProperties, PptxTheme theme, out RgbColor color, out double alpha)
    {
        if (shapeProperties?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            alpha = 1d;
            return false;
        }

        if (TryReadSolidColorWithAlpha(shapeProperties, theme, out color, out alpha))
        {
            return true;
        }

        PptxFormatSchemeReference fillReference = PptxFormatSchemeResolver.ResolveFillReference(shape, theme);
        if (fillReference.Style is not null &&
            TryReadSolidColorWithAlpha(fillReference.Style, theme, fillReference.Reference, out color, out alpha))
        {
            return true;
        }

        return fillReference.Index > 0 && TryReadSolidColorWithAlpha(fillReference.Reference, theme, out color, out alpha);
    }

    private static bool TryReadShapeLine(XElement shape, XElement? shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth, out double alpha)
    {
        XElement? explicitLine = shapeProperties?.Element(DrawingNamespace + "ln");
        if (explicitLine?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        double? styleLineWidth = TryReadStyleLineWidth(shape, theme, out double inheritedLineWidth)
            ? inheritedLineWidth
            : null;
        if (shapeProperties is not null && explicitLine is not null && TryReadLineWithAlpha(shapeProperties, theme, out color, out lineWidth, out alpha, styleLineWidth))
        {
            return true;
        }

        PptxFormatSchemeReference lineReference = PptxFormatSchemeResolver.ResolveLineReference(shape, theme);
        if (lineReference.Style is null)
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        if (TryReadThemeLineReference(lineReference, theme, out PptxSceneLineStyle line))
        {
            color = line.Color;
            lineWidth = line.Width;
            alpha = line.Alpha;
            return true;
        }

        color = default;
        lineWidth = 0d;
        alpha = 1d;
        return false;
    }

    private static bool TryReadStyleLineWidth(XElement shape, PptxTheme theme, out double lineWidth)
    {
        PptxFormatSchemeReference lineReference = PptxFormatSchemeResolver.ResolveLineReference(shape, theme);
        if (lineReference.Style?.Attribute("w") is { } widthAttribute)
        {
            lineWidth = OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture));
            return true;
        }

        lineWidth = 0d;
        return false;
    }

    private static bool TryReadThemeLineReference(XElement? lineReference, PptxTheme theme, out PptxSceneLineStyle line)
    {
        return TryReadThemeLineReference(
            new PptxFormatSchemeReference(
                lineReference,
                PptxFormatSchemeResolver.ReadIndex(lineReference),
                null),
            theme,
            out line);
    }

    private static bool TryReadThemeLineReference(PptxFormatSchemeReference lineReference, PptxTheme theme, out PptxSceneLineStyle line)
    {
        XElement? lineStyle = lineReference.Style;
        if (lineStyle is null &&
            lineReference.Index > 0 &&
            theme.TryGetLineStyle(lineReference.Index, out XElement resolvedLineStyle))
        {
            lineStyle = resolvedLineStyle;
        }

        if (lineStyle is null ||
            lineStyle.Element(DrawingNamespace + "noFill") is not null ||
            !TryReadSolidColorWithAlpha(lineStyle, theme, lineReference.Reference, out RgbColor color, out double alpha))
        {
            line = default;
            return false;
        }

        double lineWidth = lineStyle.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        XElement shapeProperties = new(DrawingNamespace + "spPr", new XElement(lineStyle));
        string? lineCap = ReadLineCap(shapeProperties);
        line = new PptxSceneLineStyle(
            true,
            color,
            lineWidth,
            alpha,
            TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> dashPattern) ? dashPattern : [],
            ReadPresetDashValue(shapeProperties),
            ReadLineCompound(shapeProperties),
            ReadLineCompoundValue(shapeProperties),
            lineCap switch
            {
                "rnd" => 1,
                "sq" => 2,
                _ => null
            },
            lineCap,
            ReadLineJoin(shapeProperties),
            ReadLineJoinValue(shapeProperties));
        return true;
    }

    private static bool TryReadLineWithAlpha(
        XElement shapeProperties,
        PptxTheme theme,
        out RgbColor color,
        out double lineWidth,
        out double alpha,
        double? fallbackLineWidth = null)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        lineWidth = line?.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : fallbackLineWidth ?? 1d;
        return TryReadSolidColorWithAlpha(line, theme, out color, out alpha);
    }

    private static bool TryReadPresetDash(XElement? shapeProperties, double lineWidth, out IReadOnlyList<double> dashPattern)
    {
        string? presetDash = ReadPresetDashValue(shapeProperties);
        double w = Math.Max(lineWidth, MinimumStrokeWidth);
        dashPattern = presetDash switch
        {
            "dot" or "sysDot" => [w, w * 2d],
            "dash" or "sysDash" => [w * 4d, w * 3d],
            "lgDash" => [w * 8d, w * 3d],
            "dashDot" or "sysDashDot" => [w * 4d, w * 3d, w, w * 3d],
            "lgDashDot" => [w * 8d, w * 3d, w, w * 3d],
            "lgDashDotDot" or "sysDashDotDot" => [w * 8d, w * 3d, w, w * 3d, w, w * 3d],
            _ => []
        };
        return dashPattern.Count > 0;
    }

    private static string? ReadPresetDashValue(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + "prstDash")
            ?.Attribute("val");
    }

    private static PptxSceneLineCompound? ReadLineCompound(XElement? shapeProperties)
    {
        string? compound = ReadLineCompoundValue(shapeProperties);
        return compound switch
        {
            "sng" => PptxSceneLineCompound.Single,
            "dbl" => PptxSceneLineCompound.Double,
            "thickThin" => PptxSceneLineCompound.ThickThin,
            "thinThick" => PptxSceneLineCompound.ThinThick,
            "tri" => PptxSceneLineCompound.Triple,
            _ => null
        };
    }

    private static string? ReadLineCompoundValue(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Attribute("cmpd");
    }

    private static string? ReadLineCap(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Attribute("cap");
    }

    private static int? ReadLineJoin(XElement? shapeProperties)
    {
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "round") is not null)
        {
            return 1;
        }

        if (line?.Element(DrawingNamespace + "bevel") is not null)
        {
            return 2;
        }

        if (line?.Element(DrawingNamespace + "miter") is not null)
        {
            return 0;
        }

        return null;
    }

    private static string? ReadLineJoinValue(XElement? shapeProperties)
    {
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "round") is not null)
        {
            return "round";
        }

        if (line?.Element(DrawingNamespace + "bevel") is not null)
        {
            return "bevel";
        }

        if (line?.Element(DrawingNamespace + "miter") is not null)
        {
            return "miter";
        }

        return null;
    }

    private static string ReadShapePreset(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "prstGeom")
            ?.Attribute("prst") ?? "rect";
    }

    private static PptxSceneLineEnd ReadLineEnd(XElement? shapeProperties, string elementName)
    {
        XElement? end = shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + elementName);
        string? type = (string?)end?.Attribute("type");
        string? width = (string?)end?.Attribute("w");
        string? length = (string?)end?.Attribute("len");
        return new PptxSceneLineEnd(
            ReadLineEndKind(type),
            type,
            ReadLineEndScale(width),
            width,
            ReadLineEndScale(length),
            length);
    }

    private static PptxSceneLineEndKind ReadLineEndKind(string? type)
    {
        return type switch
        {
            "triangle" => PptxSceneLineEndKind.Triangle,
            "arrow" => PptxSceneLineEndKind.Arrow,
            "stealth" => PptxSceneLineEndKind.Stealth,
            "diamond" => PptxSceneLineEndKind.Diamond,
            "oval" => PptxSceneLineEndKind.Oval,
            _ => PptxSceneLineEndKind.None
        };
    }

    private static double ReadLineEndScale(string? value)
    {
        return value switch
        {
            "sm" => 0.5d,
            "lg" => 1.5d,
            _ => 1d
        };
    }

    internal static PptxSceneRect ReadPictureCrop(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        XElement? sourceRectangle = blipFill?.Element(DrawingNamespace + "srcRect");
        return sourceRectangle is null
            ? default
            : ReadPercentageRectangle(sourceRectangle);
    }

    internal static PptxSceneRect ReadPictureFill(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        XElement? fillRectangle = blipFill
            ?.Element(DrawingNamespace + "stretch")
            ?.Element(DrawingNamespace + "fillRect");
        return fillRectangle is null
            ? default
            : ReadPercentageRectangle(fillRectangle);
    }

    internal static PptxScenePictureTile ReadPictureTile(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        XElement? tile = blipFill?.Element(DrawingNamespace + "tile");
        return tile is null
            ? default
            : new PptxScenePictureTile(
                true,
                tile.Name.LocalName,
                (string?)tile.Attribute("algn"),
                (string?)tile.Attribute("flip"),
                (string?)tile.Attribute("sx"),
                (string?)tile.Attribute("sy"),
                (string?)tile.Attribute("tx"),
                (string?)tile.Attribute("ty"));
    }

    internal static double ReadPictureAlpha(XElement picture)
    {
        XElement? blip = ReadPictureBlip(picture);
        XElement? alphaModFix = blip?.Element(DrawingNamespace + "alphaModFix");
        if (alphaModFix?.Attribute("amt") is { } amount &&
            int.TryParse(amount.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedAmount))
        {
            return Math.Clamp(parsedAmount / 100000d, 0d, 1d);
        }

        return 1d;
    }

    internal static string? ReadPictureAlphaValue(XElement picture)
    {
        XElement? blip = ReadPictureBlip(picture);
        return (string?)blip
            ?.Element(DrawingNamespace + "alphaModFix")
            ?.Attribute("amt");
    }

    private static XElement? ReadPictureBlip(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        return blipFill?.Element(DrawingNamespace + "blip");
    }

    internal static PptxSceneImageRecolor ReadImageRecolor(XElement picture, PptxTheme theme)
    {
        XElement? blip = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip");
        if (blip is null)
        {
            return PptxSceneImageRecolor.None;
        }

        XElement? grayscale = blip.Element(DrawingNamespace + "grayscl");
        if (grayscale is not null)
        {
            return PptxSceneImageRecolor.Grayscale(grayscale.Name.LocalName);
        }

        XElement? biLevel = blip.Element(DrawingNamespace + "biLevel");
        if (biLevel is not null)
        {
            string? thresholdValue = (string?)biLevel.Attribute("thresh");
            double threshold = thresholdValue is not null
                ? Math.Clamp(int.Parse(thresholdValue, CultureInfo.InvariantCulture) / 100000d, 0d, 1d)
                : 0.5d;
            return PptxSceneImageRecolor.BiLevel(threshold, biLevel.Name.LocalName, thresholdValue);
        }

        XElement? luminance = blip.Element(DrawingNamespace + "lum");
        if (luminance is not null)
        {
            string? brightnessValue = (string?)luminance.Attribute("bright");
            string? contrastValue = (string?)luminance.Attribute("contrast");
            double brightness = brightnessValue is not null
                ? Math.Clamp(int.Parse(brightnessValue, CultureInfo.InvariantCulture) / 100000d, -1d, 1d)
                : 0d;
            double contrast = contrastValue is not null
                ? Math.Clamp(int.Parse(contrastValue, CultureInfo.InvariantCulture) / 100000d, -1d, 1d)
                : 0d;
            return PptxSceneImageRecolor.Luminance(brightness, contrast, luminance.Name.LocalName, brightnessValue, contrastValue);
        }

        XElement? duotone = blip.Element(DrawingNamespace + "duotone");
        if (duotone is not null)
        {
            XElement[] colors = duotone.Elements().Take(2).ToArray();
            if (colors.Length == 2 &&
                TryReadImageRecolorColor(colors[0], theme, out RgbColor dark) &&
                TryReadImageRecolorColor(colors[1], theme, out RgbColor light))
            {
                return PptxSceneImageRecolor.Duotone(dark, light, duotone.Name.LocalName);
            }
        }

        return PptxSceneImageRecolor.None;
    }

    private static bool TryReadImageRecolorColor(XElement colorElement, PptxTheme theme, out RgbColor color)
    {
        if (colorElement.Name == DrawingNamespace + "prstClr")
        {
            string? preset = (string?)colorElement.Attribute("val");
            color = preset switch
            {
                "black" => new RgbColor(0, 0, 0),
                "white" => new RgbColor(255, 255, 255),
                _ => default
            };
            return preset is "black" or "white";
        }

        XElement wrapper = new(DrawingNamespace + "solidFill", new XElement(colorElement));
        return TryReadSolidColorWithAlpha(wrapper, theme, out color, out _);
    }

    private static PptxSceneRect ReadPercentageRectangle(XElement element)
    {
        string? left = (string?)element.Attribute("l");
        string? top = (string?)element.Attribute("t");
        string? right = (string?)element.Attribute("r");
        string? bottom = (string?)element.Attribute("b");
        return new PptxSceneRect(
            ParsePercentage(left),
            ParsePercentage(top),
            ParsePercentage(right),
            ParsePercentage(bottom),
            left,
            top,
            right,
            bottom);
    }

    private static double ParsePercentage(string? value)
    {
        return value is not null
            ? Math.Clamp(int.Parse(value, CultureInfo.InvariantCulture) / 100000d, 0d, 0.999d)
            : 0d;
    }

    private static PptxSceneTextBody? ReadTextBody(XElement element, IReadOnlyList<XDocument> placeholderSources, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? textBody = element.Element(PresentationNamespace + "txBody");
        if (textBody is null)
        {
            return null;
        }

        IReadOnlyList<XElement> inheritedTextBodies = FindInheritedPlaceholderShapes(element, placeholderSources)
            .Select(shape => shape.Element(PresentationNamespace + "txBody"))
            .Where(textBody => textBody is not null)
            .Cast<XElement>()
            .ToArray();
        return new PptxSceneTextBody(
            textBody.Element(DrawingNamespace + "bodyPr"),
            textBody.Element(DrawingNamespace + "lstStyle"),
            textBody.Elements(DrawingNamespace + "p").Select(paragraph => ReadParagraph(paragraph, element, textBody, inheritedTextBodies, placeholderSources, theme, colorMap)).ToArray());
    }

    private static PptxSceneTextParagraph ReadParagraph(
        XElement paragraph,
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedTextBodies,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        PptxColorMap colorMap)
    {
        XElement? properties = paragraph.Element(DrawingNamespace + "pPr");
        int level = properties?.Attribute("lvl") is { } levelAttribute
            ? int.Parse(levelAttribute.Value, CultureInfo.InvariantCulture)
            : 0;
        XElement? defaultParagraphProperties = ResolveDefaultParagraphProperties(
            level,
            shape,
            textBody,
            inheritedTextBodies,
            placeholderSources);
        XElement? defaultRunProperties = properties?.Element(DrawingNamespace + "defRPr") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
        PptxSceneParagraphStyle resolvedStyle = ResolveParagraphStyle(level, properties, defaultParagraphProperties, defaultRunProperties, shape, theme, colorMap);
        return new PptxSceneTextParagraph(
            properties,
            paragraph.Element(DrawingNamespace + "endParaRPr"),
            level,
            resolvedStyle,
            paragraph.Elements().Select(run => ReadRun(run, defaultRunProperties, resolvedStyle, theme, colorMap)).Where(run => run is not null).Cast<PptxSceneTextRun>().ToArray());
    }

    private static PptxSceneTextRun? ReadRun(XElement element, XElement? defaultRunProperties, PptxSceneParagraphStyle paragraphStyle, PptxTheme theme, PptxColorMap colorMap)
    {
        if (element.Name == DrawingNamespace + "r")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(
                PptxSceneTextRunKind.Text,
                (string?)element.Element(DrawingNamespace + "t") ?? string.Empty,
                runProperties,
                ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme, colorMap),
                element);
        }

        if (element.Name == DrawingNamespace + "br")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(PptxSceneTextRunKind.Break, "\n", runProperties, ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme, colorMap), element);
        }

        if (element.Name == DrawingNamespace + "fld")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(
                PptxSceneTextRunKind.Field,
                (string?)element.Element(DrawingNamespace + "t") ?? string.Empty,
                runProperties,
                ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme, colorMap),
                element);
        }

        return null;
    }

    private static XElement? ResolveDefaultParagraphProperties(
        int level,
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedTextBodies,
        IReadOnlyList<XDocument> placeholderSources)
    {
        string levelName = $"lvl{Math.Clamp(level + 1, 1, 9).ToString(CultureInfo.InvariantCulture)}pPr";
        var sources = new List<XElement?>();
        sources.Add(textBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName));
        sources.AddRange(inheritedTextBodies
            .Reverse()
            .Select(inheritedTextBody => inheritedTextBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName)));
        sources.Add(FindInheritedTextStyle(shape, placeholderSources, levelName));
        sources.Add(FindDefaultTextStyle(placeholderSources, levelName));
        return MergeParagraphProperties(sources.ToArray());
    }

    private static XElement? MergeParagraphProperties(params XElement?[] sources)
    {
        XElement? merged = null;
        foreach (XElement source in sources.Reverse().Where(source => source is not null).Cast<XElement>())
        {
            merged ??= new XElement(source.Name);
            foreach (XAttribute attribute in source.Attributes())
            {
                merged.SetAttributeValue(attribute.Name, attribute.Value);
            }

            foreach (XElement child in source.Elements())
            {
                XElement? existing = merged.Element(child.Name);
                if (existing is null)
                {
                    merged.Add(new XElement(child));
                }
                else
                {
                    MergeElementInto(existing, child);
                }
            }
        }

        return merged;
    }

    private static void MergeElementInto(XElement target, XElement source)
    {
        foreach (XAttribute attribute in source.Attributes())
        {
            target.SetAttributeValue(attribute.Name, attribute.Value);
        }

        foreach (XElement child in source.Elements())
        {
            target.Elements(child.Name).Remove();
            target.Add(new XElement(child));
        }
    }

    private static XElement? FindInheritedTextStyle(XElement shape, IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        return PptxTextStyleInheritance.FindInheritedTextStyle(shape, placeholderSources, levelName);
    }

    private static XElement? FindDefaultTextStyle(IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        return PptxTextStyleInheritance.FindDefaultTextStyle(placeholderSources, levelName);
    }

    private static IReadOnlyList<XElement> FindInheritedPlaceholderShapes(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        return PptxPlaceholderMatcher.FindInheritedPlaceholderShapes(shape, placeholderSources);
    }

    private static PptxSceneParagraphStyle ResolveParagraphStyle(
        int level,
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        XElement? defaultRunProperties,
        XElement shape,
        PptxTheme theme,
        PptxColorMap colorMap)
    {
        double fontSize = ReadFontSize(defaultRunProperties, null);
        RgbColor color = TryReadSolidColorWithAlpha(defaultRunProperties, theme, colorMap, out RgbColor defaultColor, out double alpha)
            ? defaultColor
            : TryReadShapeFontColor(shape, theme, colorMap, out RgbColor shapeColor)
                ? shapeColor
                : new RgbColor(0, 0, 0);
        if (!TryReadSolidColorWithAlpha(defaultRunProperties, theme, colorMap, out _, out alpha))
        {
            alpha = 1d;
        }

        string? typeface = theme.ResolveTypeface((string?)defaultRunProperties?.Element(DrawingNamespace + "latin")?.Attribute("typeface"));
        return new PptxSceneParagraphStyle(
            level,
            (string?)(paragraphProperties?.Attribute("algn") ?? defaultParagraphProperties?.Attribute("algn")) ?? "l",
            fontSize,
            color,
            alpha,
            typeface,
            ParseOptionalBoolAttribute(defaultRunProperties, "b"),
            ParseOptionalBoolAttribute(defaultRunProperties, "i"),
            ReadCharacterSpacing(defaultRunProperties, null));
    }

    private static PptxSceneRunStyle ResolveRunStyle(XElement? runProperties, XElement? defaultRunProperties, PptxSceneParagraphStyle paragraphStyle, PptxTheme theme, PptxColorMap colorMap)
    {
        double fontSize = ReadFontSize(runProperties, defaultRunProperties);
        double alpha = paragraphStyle.Alpha;
        RgbColor color = paragraphStyle.Color;
        if (TryReadSolidColorWithAlpha(runProperties, theme, colorMap, out RgbColor runColor, out double runAlpha))
        {
            color = runColor;
            alpha = runAlpha;
        }
        else if (TryReadSolidColorWithAlpha(defaultRunProperties, theme, colorMap, out RgbColor defaultColor, out double defaultAlpha))
        {
            color = defaultColor;
            alpha = defaultAlpha;
        }

        string? typeface = theme.ResolveTypeface((string?)(runProperties?.Element(DrawingNamespace + "latin") ??
            defaultRunProperties?.Element(DrawingNamespace + "latin"))?.Attribute("typeface")) ?? paragraphStyle.Typeface;
        bool bold = ParseOptionalBoolAttribute(runProperties, "b") ||
            (runProperties?.Attribute("b") is null && paragraphStyle.Bold);
        bool italic = ParseOptionalBoolAttribute(runProperties, "i") ||
            (runProperties?.Attribute("i") is null && paragraphStyle.Italic);
        string? underlineValue = ReadUnderlineValue(runProperties, defaultRunProperties);
        string? strikeValue = ReadStrikeValue(runProperties, defaultRunProperties);
        string? capsValue = ReadTextCapsValue(runProperties, defaultRunProperties);
        bool underline = underlineValue is not null &&
            !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
        bool strike = IsStrikeEnabled(strikeValue);
        return new PptxSceneRunStyle(
            fontSize,
            color,
            alpha,
            typeface,
            bold,
            italic,
            underline,
            underlineValue,
            strike,
            strikeValue,
            capsValue,
            ReadCharacterSpacing(runProperties, defaultRunProperties),
            ReadBaselineOffset(runProperties, defaultRunProperties, fontSize),
            TryReadHighlightColor(runProperties, out RgbColor highlight) ? highlight : null);
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
        return IsStrikeEnabled(ReadStrikeValue(runProperties, defaultRunProperties));
    }

    private static bool IsStrikeEnabled(string? value)
    {
        return value is not null && !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }

    private static string? ReadUnderlineValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("u") ?? defaultRunProperties?.Attribute("u"));
    }

    private static string? ReadStrikeValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("strike") ?? defaultRunProperties?.Attribute("strike"));
    }

    private static string? ReadTextCapsValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("cap") ?? defaultRunProperties?.Attribute("cap"));
    }

    private static bool ParseOptionalBoolAttribute(XElement? element, string attributeName)
    {
        return ParseBoolAttribute(element, attributeName, defaultValue: false);
    }

    private static bool ParseBoolAttribute(XElement? element, string attributeName, bool defaultValue)
    {
        return OoxBoolean.ParseAttribute(element, attributeName, defaultValue);
    }

    private static int ParseOptionalIntAttribute(XElement? element, string attributeName, int defaultValue)
    {
        return element?.Attribute(attributeName) is { } attribute &&
            int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)
            ? value
            : defaultValue;
    }

    private static double ParseOptionalDoubleAttribute(XElement element, string attributeName, double defaultValue)
    {
        return element.Attribute(attributeName) is { } attribute &&
            double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
            ? value
            : defaultValue;
    }

    private static bool TryReadHighlightColor(XElement? runProperties, out RgbColor color)
    {
        XElement? highlight = runProperties?.Element(DrawingNamespace + "highlight");
        string? hex = (string?)highlight?.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        return RgbColor.TryParse(hex, out color);
    }

    private static bool TryReadShapeFontColor(XElement shape, PptxTheme theme, PptxColorMap colorMap, out RgbColor color)
    {
        XElement? fontRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fontRef");
        return TryReadSolidColorWithAlpha(fontRef, theme, colorMap, out color, out _);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, out RgbColor color, out double alpha)
    {
        return PptxColorResolver.TryReadSolidColorWithAlpha(element, theme, out color, out alpha);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, PptxColorMap colorMap, out RgbColor color, out double alpha)
    {
        return PptxColorResolver.TryReadSolidColorWithAlpha(element, theme, colorMap, out color, out alpha);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, XElement? placeholderColorContainer, out RgbColor color, out double alpha)
    {
        return PptxColorResolver.TryReadSolidColorWithAlpha(element, theme, placeholderColorContainer, out color, out alpha);
    }

    private static double ReadAlpha(XElement? colorContainer)
    {
        return PptxColorResolver.ReadAlpha(colorContainer);
    }

    private static long ReadLong(XElement element, string name)
    {
        return element.Attribute(name) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : 0L;
    }

    private static long ReadLong(XElement element, string name, long defaultValue)
    {
        return element.Attribute(name) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : defaultValue;
    }

    private static bool ReadBool(XElement element, string name)
    {
        string? value = (string?)element.Attribute(name);
        return value is "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }
}
