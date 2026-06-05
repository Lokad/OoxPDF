using System.IO.Compression;
using System.Text;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static class OoxmlTests
{
    public static void ParsesContentTypesAndRelationships()
    {
        using MemoryStream packageStream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = "<p:presentation xmlns:p=\"p\"/>",
            ["ppt/slides/slide1.xml"] = "<p:sld xmlns:p=\"p\"/>"
        });

        OoxPackage package = OoxPackage.Open(packageStream);

        OoxPart presentation = TestAssert.NotNull(package.GetPart("/ppt/presentation.xml"));
        TestAssert.Equal("application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml", presentation.ContentType);

        IReadOnlyList<OoxRelationship> relationships = package.GetRelationships("/");
        TestAssert.Equal(1, relationships.Count);
        TestAssert.Equal("/ppt/presentation.xml", relationships[0].ResolvedTarget);
    }

    public static void RejectsPackagePartPathTraversal()
    {
        TestAssert.Throws<InvalidDataException>(() => OoxPath.NormalizePartName("../evil.xml"));
    }

    public static void ResolvesRelationshipTargets()
    {
        string resolved = OoxPath.ResolveRelationshipTarget("/ppt/slides/slide1.xml", "../media/image1.png");

        TestAssert.Equal("/ppt/media/image1.png", resolved);
    }

    public static void ConvertsCommonOfficeUnits()
    {
        TestAssert.Equal(72d, OoxUnits.EmuToPoints(914400));
        TestAssert.Equal(12d, OoxUnits.TwipsToPoints(240));
        TestAssert.Equal(9d, OoxUnits.HalfPointsToPoints(18));
    }
}
