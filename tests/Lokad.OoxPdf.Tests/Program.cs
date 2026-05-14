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
    PptxTests.PptxSyntheticTwoSlidesProducesTwoPdfPages,
    PptxTests.PptxSyntheticShapesProduceDrawingOperators,
    PptxTests.PptxSyntheticRotatedShapeProducesTransform,
    DocxTests.DocxSyntheticDocumentProducesOnePdfPage,
    FontTests.WindowsFontResolverFindsInstalledFonts,
    FontTests.OpenTypeParserMapsBasicLatinGlyphs);
