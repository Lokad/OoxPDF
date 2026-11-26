using System.Reflection;

namespace Lokad.OoxPdf.Tests;

internal sealed record TestCase(string Group, Action Action, bool IsSlow = false);

internal static class TestCatalog
{
    public static IReadOnlyList<TestCase> All { get; } =
    [
        .. FromType("api", typeof(PublicApiTests)),
        .. FromType("api", typeof(TestRunnerTests)),
        .. FromType("api", typeof(TestTemporaryDirectoryTests)),
        .. FromType(PortableCompositionGroup, typeof(PortableRenderingTests)),
        .. FromType("rendering-portable", typeof(PptxScalarCapsTests)),
        .. FromType("rendering-portable", typeof(PptxBreakCompositionTests)),
        .. FromType("rendering-portable", typeof(PptxPictureRecolorTests)),
        .. FromType("rendering-portable", typeof(PptxAxisIdentityTests)),
        .. FromType("rendering-portable", typeof(PptxEffectColorTests)),
        .. FromType("rendering-portable", typeof(SvgTextTests)),
        .. FromType("rendering-portable", typeof(DocxMathCenterGroupTests)),
        .. FromType("rendering-portable", typeof(DocxTabCascadeTests)),
        .. FromType("rendering-portable", typeof(DocxNoteContinuationPaintTests)),
        .. FromType("rendering-portable", typeof(DocxNoteRevisionLaneTests)),
        .. FromType("rendering-portable", typeof(DocxUnitPrintOriginTests)),
        .. FromType("rendering-portable", typeof(DocxFooterRevisionFrameTests)),
        .. FromType("rendering-portable", typeof(DocxWrappedMinimumLeadingTests)),
        .. FromType("rendering-portable", typeof(DocxColumnAdmissionTests)),
        .. FromType("rendering-portable", typeof(DocxBalloonFrameTests)),
        .. FromType("rendering-portable", typeof(DocxCommentGlyphFrameTests)),
        .. FromType("rendering-portable", typeof(DocxEndnoteBlockTests)),
        .. FromType("rendering-portable", typeof(DocxEndnoteCounterPrefixTests)),
        .. FromType("rendering-portable", typeof(DocxNoteClippingTests)),
        .. FromType("rendering-portable", typeof(DocxDeletedParagraphMarkTests)),
        .. FromType("rendering-portable", typeof(DocxCellColumnBreakTests)),
        .. FromType("rendering-portable", typeof(DocxHardBreakVisibilityTests)),
        .. FromType("rendering-portable", typeof(DocxHiddenInlineBoxTests)),
        .. FromType("rendering-portable", typeof(DocxEmptyRevisionScopeTests)),
        .. FromType("rendering-portable", typeof(DocxFieldInstructionTests)),
        .. FromType(name => name == nameof(DocxMinimumCommentTests.DocxMinimumCommentsMatchIndependentWordFinalOrigins)
            ? "docx-text" : "rendering-portable", typeof(DocxMinimumCommentTests)),
        .. FromType(name => name == nameof(DocxPrintedMinimumTests.DocxPrintedMinimumMatchesIndependentWordParagraphOrigins)
            ? "docx-text" : "rendering-portable", typeof(DocxPrintedMinimumTests)),
        .. FromType("rendering-portable", typeof(DocxEndnoteParagraphBoundaryTests)),
        .. FromType("rendering-portable", typeof(DocxEndnoteFinalSectionTests)),
        .. FromType("rendering-portable", typeof(DocxRegisteredSeparatorTests)),
        .. FromType("rendering-portable", typeof(DocxRepeatedTableHeaderTests)),
        .. FromType("rendering-portable", typeof(DocxTableKeepContentTests)),
        .. FromType("rendering-portable", typeof(DocxRepeatedMixedRowTests)),
        .. FromType("rendering-portable", typeof(DocxNestedCellEndTests)),
        .. FromType("rendering-portable", typeof(DocxReviewTableFrameTests)),
        .. FromType("rendering-portable", typeof(DocxAutomaticReviewRowTests)),
        .. FromType("rendering-portable", typeof(DocxCommentAnchorTests)),
        new("rendering-portable", PptxColumnIntervalTests.PptxCategoryOffsetsShareMeasuredUnitsAcrossPreparedContainers),
        new("rendering-portable", PptxColumnIntervalTests.PptxShortRotatedCaptionsKeepFiniteSpacingAndAuthoredSkips),
        new("rendering-portable", PptxColumnIntervalTests.PptxSkippedCategoryPaintKeepsLogicalReferenceEnds),
        new("rendering-portable", PptxColumnIntervalTests.PptxOuterCategoryFittingSharesDefaultEllipsisAndOccupiedFrames),
        new("rendering-portable", PptxColumnIntervalTests.PptxAutomaticCategoryReferencesKeepSourceEndpoints),
        .. FromType("rendering-portable", typeof(DocxAutomaticColorTests)),
        .. FromType(name => name == nameof(DocxParagraphShadingTests.DocxParagraphShadingPrintedFramesMatchIndependentWord)
            ? "docx-text" : "rendering-portable", typeof(DocxParagraphShadingTests)),
        .. FromType("docx-page", typeof(DocxRevisionBarTests)),
        .. FromType("docx-page", typeof(DocxExplicitPageStartSpacingTests)),
        .. FromType("docx-page", typeof(DocxSpacedWrappedFootnoteTests)),
        .. FromType("docx-page", typeof(DocxMixedWrappedFootnoteTests)),
        .. FromType("fonts", typeof(FontProgramCacheTests)),
        .. FromType("fonts", typeof(FontFamilyAliasTests)),
        .. FromType("fonts", typeof(FallbackFontMetricsTests)),
        .. FromType("rendering-portable", typeof(FontRecoveryTests)),
        .. FromType("rendering-portable", typeof(DocxSectionNavigationTests)),
        .. FromType("rendering-portable", typeof(DocxPartOwnershipTests)),
        .. FromType("rendering-portable", typeof(DocxImageRecoveryTests)),
        .. FromType("rendering-portable", typeof(DocxGlyphlessAnchorTests)),
        .. FromType("rendering-portable", typeof(DocxStoryOwnershipTests)),
        .. FromType("rendering-portable", typeof(PptxSlideMembershipTests)),
        .. FromType("rendering-portable", typeof(PptxTableDefaultTests)),
        .. FromType("rendering-portable", typeof(PptxPictureRecoveryTests)),
        .. FromType("rendering-portable", typeof(PptxAffineGeometryTests)),
        .. FromType("rendering-portable", typeof(PptxWorkbookOwnershipTests)),
        .. FromType("rendering-portable", typeof(PptxDiagnosticAdmissionTests)),
        .. FromType("rendering-portable", typeof(PptxSlideNumberTests)),
        .. FromType("rendering-portable", typeof(PptxBulletChoiceTests)),
        .. FromType("rendering-portable", typeof(PptxVisibleDiagnosticTests)),
        .. FromType("rendering-portable", typeof(PptxMasterThemeTests)),
        .. FromType("rendering-portable", typeof(PptxInheritanceTests)),
        .. FromType("rendering-portable", typeof(DocxHiddenPictureTests)),
        .. FromType("rendering-portable", typeof(DocxStyleToggleTests)),
        .. FromType("rendering-portable", typeof(DocxParagraphIdentityTests)),
        .. FromType("rendering-portable", typeof(DocxLockedFieldTests)),
        .. FromType("rendering-portable", typeof(DocxVmlTextBoxTests)),
        .. FromType("rendering-portable", typeof(DocxFontSlotTests)),
        .. FromType("fonts", typeof(FontResolverConcurrencyTests)),
        .. FromType("fonts", typeof(FontDiscoveryAsyncTests)),
        .. FromType("fonts", typeof(MathFontTests)),
        .. FromType("api", typeof(DiagnosticOutcomeTests)),
        .. FromType("cli", typeof(CliTests)),
        .. FromType("tools", typeof(ToolBudgetTests)),
        .. FromType("ooxml", typeof(OoxmlTests)),
        .. FromType("ooxml", typeof(OoxmlStrictDocxTests)),
        .. FromType("ooxml", typeof(OoxMarkupCompatibilityTests)),
        .. FromType("ooxml", typeof(OoxComplexScriptTests)),
        .. FromType("ooxml", typeof(OoxmlBoundsTests)),
        .. FromType("ooxml", typeof(OoxLimitsTests)),
        .. FromType("ooxml", typeof(OoxResourceGuaranteeTests)),
        .. FromType("ooxml", typeof(OoxMemoryAdmissionTests)),
        .. FromType("ooxml", typeof(OoxMissingFontTests)),
        .. FromType("pdf", typeof(PdfWriterTests)),
        .. FromType("pdf", typeof(PdfTransparencyGroupTests)),
        .. FromType("pdf", typeof(PdfResourceIndexTests)),
        .. FromType("pdf", typeof(PdfIdentityTests)),
        .. FromType("rendering-portable", typeof(PdfUnicodeIdentityTests)),
        new("docx-core", DocxTests.DocxScaledMeasurerPreservesDesignInsetAndScalesAdvances),
        .. FromType("docx-images", typeof(DocxImagesTests)),
        .. FromType("docx-images", typeof(DocxPictureRevisionTests)),
        .. FromType("docx-text", typeof(DocxRevisionNeutralityTests)),
        .. FromType("docx-images", typeof(DocxImageParagraphAdvanceTests)),
        .. FromType("docx-math", typeof(DocxMathTests)),
        .. FromType("docx-math", typeof(DocxNativeMathTests)),
        .. FromType(name => name is nameof(DocxInlineCompositionTests.DocxEquationZonesKeepDrawingsWithoutMathOperands)
            or nameof(DocxInlineCompositionTests.DocxEquationDrawingCompanionsKeepInnerRevisionOwners)
            or nameof(DocxInlineCompositionTests.DocxEquationDrawingCompanionsKeepInnerHyperlinkCoverage)
            or nameof(DocxInlineCompositionTests.DocxEquationDrawingCompanionsKeepCombinedLinkAndRevisionScopes)
            or nameof(DocxInlineCompositionTests.DocxEquationWarningsIgnoreRemovedLiteralsAndRetainStructure)
            or nameof(DocxInlineCompositionTests.DocxLinearFractionsKeepSlashRowsAcrossDocumentStories)
            ? "docx-math-portable" : "docx-math", typeof(DocxInlineCompositionTests)),
        .. FromType("docx-math", typeof(DocxTableMathTests)),
        .. FromType(name => name.StartsWith("MathFont", StringComparison.Ordinal) ? "fonts" : "docx-math", typeof(DocxMathBoxTests)),
        .. FromType("docx-math-portable", typeof(DocxMathFractionTests)),
        .. FromType("docx-math-portable", typeof(DocxSkewedFractionTests)),
        .. FromType("docx-math-portable", typeof(DocxMathNaryTests)),
        .. FromType("docx-math-portable", typeof(DocxMathFunctionLayoutTests)),
        .. FromType("docx-math-portable", typeof(DocxMathMatrixLayoutTests)),
        .. FromType("docx-math-portable", typeof(MathScriptStyleTests)),
        .. FromType("docx-numbering", typeof(DocxNumberingTests)),
        .. FromType("docx-numbering", typeof(DocxListKeepLeadingTests)),
        .. FromType("docx-text", typeof(DocxTextTests)),
        .. FromType("docx-text", typeof(DocxTextWrapTests)),
        .. FromType("docx-text", typeof(DocxTextSpacingTests)),
        .. FromType("docx-page", typeof(DocxPageTests)),
        .. FromType("docx-page", typeof(DocxHeaderFooterTests)),
        .. FromType("docx-page", typeof(DocxPageMarkupTests)),
        .. FromType("docx-page", typeof(DocxFootnotesTests)),
        .. FromType("docx-page", typeof(DocxPageFootnoteNumberingTests)),
        .. FromType("docx-page", typeof(DocxWrappedFootnoteNumberingTests)),
        .. FromType("docx-text", typeof(DocxWrappedCalibriKerningTests)),
        .. FromType("docx-page", typeof(DocxMetWrappedFootnoteNumberingTests)),
        .. FromType("docx-page", typeof(DocxCompleteFootnoteReserveTests)),
        .. FromType("docx-text", typeof(DocxPlainFootnoteKerningTests)),
        .. FromType("docx-text", typeof(DocxOrdinaryKerningTests)),
        .. FromType("docx-text", typeof(DocxMinimumBodyLineLeadingTests)),
        .. FromType("docx-text", typeof(DocxPassiveStaticBodyTests)),
        .. FromType("docx-text", typeof(DocxPagePrintOriginTests)),
        .. FromType("docx-text", typeof(DocxMinimumBodyFlowTests)),
        .. FromType("docx-text", typeof(DocxMarkedMinimumLeadingTests)),
        .. FromType("docx-text", typeof(DocxMinimumBeforeLeadingTests)),
        .. FromType("docx-text", typeof(DocxOrdinaryMinimumSpacingTests)),
        .. FromType("docx-text", typeof(DocxContextualMinimumLeadingTests)),
        .. FromType("docx-text", typeof(DocxElevenPointBodyTests)),
        .. FromType("docx-text", typeof(DocxNominalGlyphAdvanceTests)),
        .. FromType("docx-text", typeof(DocxMetGlyphAdvanceTests)),
        .. FromType("docx-text", typeof(DocxLineSpacingInheritanceTests)),
        .. FromType("docx-text", typeof(DocxNaturalBodyLineTests)),
        .. FromType("docx-tables", typeof(DocxTablesTests)),
        .. FromType("docx-tables", typeof(DocxConditionalTableStyleTests)),
        .. FromType("docx-tables", typeof(DocxTableStyleMarginTests)),
        .. FromType("docx-tables", typeof(DocxTableStyleInheritanceTests)),
        .. FromType("docx-tables", typeof(DocxTableBandSizeTests)),
        .. FromType("docx-tables", typeof(DocxTableLookTests)),
        .. FromType("docx-tables", typeof(DocxAdjacentTableTests)),
        .. FromType("docx-tables", typeof(DocxAutofitSpanTests)),
        .. FromType("docx-tables", typeof(DocxNormalTableStyleTests)),
        .. FromType("docx-text", typeof(DocxNoteListOrderTests)),
        .. FromType("docx-tables", typeof(DocxTableCellsTests)),
        .. FromType("docx-tables", typeof(DocxCellMemoTests)),
        .. FromType("docx-tables", typeof(DocxTableFlowTests)),
        .. FromType("docx-tables", typeof(DocxBorderPlanTests)),
        .. FromType("docx-tables", typeof(DocxTablesLayoutTests)),
        .. FromType("docx-tables", typeof(DocxTablesMarkupTests)),
        .. FromType("docx-core", typeof(DocxCoreTests)),
        .. FromType("docx-core", typeof(DocxHyperlinksTests)),
        .. FromType(name => name is nameof(DocxFontRenderingTests.DocxRendererUsesThemeTypefaceForThemeOnlyDefaultRun)
            or nameof(DocxFontRenderingTests.DocxRendererUsesFontCatalogAlternateBeforeResolverFallback)
            ? "docx-composition-portable" : "docx-core", typeof(DocxFontRenderingTests)),
        .. FromType("docx-core", typeof(DocxFontCacheTests)),
        .. FromType("docx-core", typeof(DocxFieldsTests)),
        .. FromType("docx-core", typeof(DocxDynamicFieldFormatTests)),
        .. FromType("docx-core", typeof(DocxRevisionsTests)),
        .. FromType("docx-core", typeof(DocxMarkupTests)),
        .. FromType("docx-core", typeof(DocxCommentsTests)),
        .. FromType(name => name.StartsWith("DocxCombinedInspection", StringComparison.Ordinal)
            ? "rendering-portable" : "docx-core", typeof(DocxInspectionTests)),
        .. FromType("pptx-composition", typeof(PptxCompositionTests)),
        .. FromType("pptx-core", typeof(PptxCoreTests)),
        .. FromType("pptx-tables", typeof(PptxTablesTests)),
        .. FromType("pptx-tables", typeof(PptxTableMergeTests)),
        .. FromType("pptx-images", typeof(PptxImagesTests)),
        .. FromType("pptx-model", typeof(PptxModelTests)),
        .. FromType(name => name is nameof(PptxDiagnosticsTests.PptxSyntheticInheritedPlaceholderTextIsSkipped)
            or nameof(PptxDiagnosticsTests.PptxSyntheticSlidePlaceholderTextUsesInheritedBounds)
            ? "rendering-portable" : "pptx-model", typeof(PptxDiagnosticsTests)),
        .. FromType("pptx-shapes", typeof(PptxShapesTests)),
        .. FromType("pptx-shapes", typeof(PptxSceneTextAgreementTests)),
        .. FromType(name => name is nameof(PptxVisibilityTests.PptxHiddenObjectsLeaveNoPaintResourcesOrLinks)
            or nameof(PptxVisibilityTests.PptxMasterShapeVisibilityPreservesOwnedFills)
            or nameof(PptxVisibilityTests.PptxGlyphEmissionInspectionUsesSceneVisibilityAndMasterAdmission)
            or nameof(PptxVisibilityTests.PptxTableInspectionUsesMasterAdmission)
            ? "rendering-portable" : "pptx-composition", typeof(PptxVisibilityTests)),
        .. FromType("pptx-shapes", typeof(PptxHyperlinksTests)),
        .. FromType("pptx-shapes", typeof(PptxTextHyperlinksTests)),
        .. FromType("pptx-shapes", typeof(PptxFrameHyperlinksTests)),
        .. FromType("pptx-charts", typeof(PptxChartTextHyperlinksTests)),
        .. FromType("pptx-charts", typeof(PptxChartsTests)),
        .. FromType("rendering-portable", typeof(PptxChartPaletteTests)),
        .. FromType("rendering-portable", typeof(PptxLinePointStyleTests)),
        .. FromType("rendering-portable", typeof(PptxAxisCrossingTests)),
        .. FromType("rendering-portable", typeof(PptxRightLegendAdmissionTests)),
        .. FromType("rendering-portable", typeof(PptxColumnCrossingTests)),
        .. FromType("rendering-portable", typeof(PptxColumnIntervalTests)),
        .. FromType("rendering-portable", typeof(PptxHorizontalScaleTests)),
        .. FromType("rendering-portable", typeof(PptxSourceUnitBoundsTests)),
        .. FromType("rendering-portable", typeof(PptxScatterIntervalTests)),
        .. FromType("rendering-portable", typeof(PptxScatterSmoothingTests)),
        .. FromType("rendering-portable", typeof(PptxCartesianIntervalTests)),
        .. FromType("rendering-portable", typeof(PptxCategorySlotTests)),
        .. FromType("rendering-portable", typeof(PptxAreaFrameTests)),
        .. FromType("rendering-portable", typeof(PptxAuthoredColorTests)),
        .. FromType("rendering-portable", typeof(PptxChartSpacingTests)),
        .. FromType("rendering-portable", typeof(PptxAutomaticInsetTests)),
        .. FromType("rendering-portable", typeof(PptxHorizontalFrameTests)),
        .. FromType("rendering-portable", typeof(PptxTitleShapeTests)),
        .. FromType("rendering-portable", typeof(PptxCartesianHeaderTests)),
        .. FromType("rendering-portable", typeof(PptxCartesianSideTests)),
        .. FromType("rendering-portable", typeof(PptxChartLabelDefaultsTests)),
        .. FromType("rendering-portable", typeof(PptxColumnLabelWrappingTests)),
        .. FromType("rendering-portable", typeof(PptxLegendDefaultTests)),
        .. FromType("rendering-portable", typeof(PptxLegendRowsTests)),
        .. FromType("rendering-portable", typeof(PptxPointLegendTests)),
        .. FromType("rendering-portable", typeof(PptxSeriesOrderTests)),
        .. FromType("rendering-portable", typeof(PptxBuiltinPaletteTests)),
        .. FromType("pptx-charts", typeof(PptxChartLegendsTests)),
        .. FromType("pptx-charts", typeof(PptxChartTitleDefaultsTests)),
        .. FromType("pptx-charts", typeof(PptxChartAxesTests)),
        .. FromType("rendering-portable", typeof(PptxAutomaticScaleTests)),
        .. FromType("pptx-charts", typeof(PptxChartRenderingTests)),
        .. FromType("rendering-portable", typeof(PptxChartNegativeInversionTests)),
        .. FromType("pptx-charts", typeof(PptxChartRangeTests)),
        .. FromType("pptx-charts", typeof(PptxChartDensifyTests)),
        .. FromType(name => name is nameof(PptxChartNumberFormatsTests.PptxChartNumberMetadataKeepsConditionalAxisAndDataLabelContent)
            or nameof(PptxChartNumberFormatsTests.PptxChartNumberFractionLabelsKeepLargeMagnitudes)
            or nameof(PptxChartNumberFormatsTests.PptxChartNumberBubbleLabelsKeepSourceDateContext)
            or nameof(PptxChartNumberFormatsTests.PptxChartNumberLabelFormatsKeepCodeAndWorkbookLinkOwners)
            ? "rendering-portable" : "pptx-charts", typeof(PptxChartNumberFormatsTests)),
        .. FromType("pptx-charts", typeof(PptxChartPlotOptionsAgreementTests)),
        .. FromType("pptx-typography", typeof(PptxTypographyTests)),
        .. FromType("pptx-typography", typeof(PptxTextSpacingTests)),
        .. FromType("pptx-typography", typeof(PptxTextLayoutTests)),
        .. FromType("pptx-typography", typeof(PptxNumberingTests)),
        .. FromType("pptx-typography", typeof(PptxNoWrapDistributionTests)),
        .. FromType("pptx-typography", typeof(PptxFontCacheTests)),
        .. FromType("imaging", typeof(ImagingTests)),
        .. FromType("imaging", typeof(JpegPassthroughTests)),
        .. FromType("imaging", typeof(ImageRecoveryTests)),
        .. FromType("fonts", typeof(FontFormatTests)),
        .. FromType("fonts", typeof(FontTests))
    ];

    // The shared fixture class also contains standalone presentation tests.
    // These stay in the portable suite without entering the DOCX milestone.
    private static string PortableCompositionGroup(string name) => name switch
    {
        nameof(PortableRenderingTests.PortableRenderingPptxJustificationPreservesWordAndGlyphGeometry) or
        nameof(PortableRenderingTests.PortableRenderingPptxPreservesCompositeChartAndText) or
        nameof(PortableRenderingTests.PortableRenderingOperationsDecksKeepCompleteLegendsAndPictures) or
        nameof(PortableRenderingTests.PortableRenderingChartCachedFieldsKeepWordsAndHyperlinksAcrossTextOwners) or
        nameof(PortableRenderingTests.PortableRenderingRichChartTitleKeepsAuthoredLinesAndTheirLinks) or
        nameof(PortableRenderingTests.PortableRenderingChartSeriesOrderSurvivesXmlReordering) or
        nameof(PortableRenderingTests.PortableRenderingAutomaticColumnLabelsWrapInsideChartFrames) or
        nameof(PortableRenderingTests.PortableRenderingComposedColumnsKeepOwnedLabelsAndOfficeColors) or
        nameof(PortableRenderingTests.PortableRenderingColumnLabelWrappingKeepsUnicodeAndBreaks) or
        nameof(PortableRenderingTests.PortableRenderingBarPointKeysKeepSingleSeriesCategoryIdentity) or
        nameof(PortableRenderingTests.PortableRenderingBarPalettesFollowIndependentOfficeSourceIndices) or
        nameof(PortableRenderingTests.PortableRenderingBarPaletteIndicesKeepBoundsAndFallbacks) or
        nameof(PortableRenderingTests.PortableRenderingSharedAxisFramesMatchIndependentOfficePairs) or
        nameof(PortableRenderingTests.PortableRenderingSharedAxisPlacesEveryPlotOnItsResolvedScale) or
        nameof(PortableRenderingTests.PortableRenderingComposedPlotsShareOnePrimaryAxisLabelStrip) or
        nameof(PortableRenderingTests.PortableRenderingComposedAxisLabelsRetainSourceVisibility) or
        nameof(PortableRenderingTests.PortableRenderingComposedSameSideAxesFitAutomaticContainer) or
        nameof(PortableRenderingTests.PortableRenderingComposedAutoAxesUseTheirOwningScale) or
        nameof(PortableRenderingTests.PortableRenderingComposedOppositeAxesKeepIndependentBands) => "rendering-portable",
        _ => "docx-composition-portable"
    };

    private static IReadOnlyList<TestCase> FromType(string group, Type type)
    {
        return FromType(_ => group, type);
    }

    private static IReadOnlyList<TestCase> FromType(Func<string, string> classify, Type type)
    {
        return type
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(method => method.ReturnType == typeof(void) && method.GetParameters().Length == 0)
            .OrderBy(method => method.MetadataToken)
            .Select(method => new TestCase(classify(method.Name), (Action)Delegate.CreateDelegate(typeof(Action), method),
                IsSlowTest(method.Name)))
            .ToArray();
    }

    private static bool IsSlowTest(string name) => name is
        nameof(DocxCommentsTests.DocxWordCompatibleMixedBalloonWrapsParagraphBodiesIndependently) or
        nameof(PptxTypographyTests.PptxSyntheticRotatedTextBoxProducesTransform) or
        nameof(PptxTypographyTests.PptxSyntheticTextBoxEmbedsFontAndDrawsGlyphs) or
        nameof(FontTests.DiscoveryHeadersMatchFullLoadAcrossWindowsFonts);
}
