using Lokad.OoxPdf;

namespace Lokad.OoxPdf.Tests;

internal static class PublicApiTests
{
    public static void PublicApiRejectsMissingInput()
    {
        string missingInput = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");

        FileNotFoundException ex = TestAssert.Throws<FileNotFoundException>(
            () => OoxPdfConverter.Convert(missingInput, output));

        TestAssert.Equal(missingInput, ex.FileName);
    }

    public static void AutoDetectsPptxExtension()
    {
        OoxPdfInputKind kind = OoxPdfConverter.DetectInputKind("deck.PPTX");

        TestAssert.Equal(OoxPdfInputKind.Pptx, kind);
    }

    public static void AutoDetectsDocxExtension()
    {
        OoxPdfInputKind kind = OoxPdfConverter.DetectInputKind("document.docx");

        TestAssert.Equal(OoxPdfInputKind.Docx, kind);
    }

    public static void DeterministicConversionProducesStableBytes()
    {
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
                </w:document>
                """
        });
        string output1 = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        string output2 = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output1, new OoxPdfOptions { Deterministic = true });
        OoxPdfConverter.Convert(input, output2, new OoxPdfOptions { Deterministic = true });

        byte[] first = File.ReadAllBytes(output1);
        byte[] second = File.ReadAllBytes(output2);
        TestAssert.True(first.SequenceEqual(second), "Deterministic conversion should produce stable PDF bytes.");
    }
}
