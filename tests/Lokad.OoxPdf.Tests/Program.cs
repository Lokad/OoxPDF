using Lokad.OoxPdf.Tests;

return TestRunner.Run(
    PublicApiTests.PublicApiRejectsMissingInput,
    PublicApiTests.AutoDetectsPptxExtension,
    PublicApiTests.AutoDetectsDocxExtension,
    OoxmlTests.ParsesContentTypesAndRelationships,
    OoxmlTests.RejectsPackagePartPathTraversal,
    OoxmlTests.ResolvesRelationshipTargets,
    OoxmlTests.ConvertsCommonOfficeUnits,
    PdfWriterTests.WritesSingleBlankPagePdfStructure,
    PdfWriterTests.WritesMultipleBlankPagesWithPageSizes,
    PdfWriterTests.WritesDrawingOperators,
    PdfWriterTests.WritesEmbeddedTrueTypeFontObjects,
    PptxTests.PptxSyntheticTwoSlidesProducesTwoPdfPages,
    PptxTests.PptxSyntheticShapesProduceDrawingOperators,
    PptxTests.PptxSyntheticRotatedShapeProducesTransform,
    PptxTests.PptxSyntheticTextBoxEmbedsFontAndDrawsGlyphs,
    DocxTests.DocxSyntheticDocumentProducesOnePdfPage,
    FontTests.WindowsFontResolverFindsInstalledFonts,
    FontTests.OpenTypeParserMapsBasicLatinGlyphs);
