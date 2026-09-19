using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxFontCacheTests
{
    public static void PptxEstimatorsSharingResolverReuseLoadedPrograms()
    {
        var program = new MemoryFontProgramSource("cache-test", TestFontBuilder.CreateTestFont());
        var resolver = new PresentationFontResolver(new SingleProgramResolver(program));
        Type estimatorType = typeof(PptxRenderer).GetNestedType(
            "TextAdvanceEstimator",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected text advance estimator.");
        object first = System.Activator.CreateInstance(estimatorType, resolver, CancellationToken.None) ?? throw new InvalidOperationException("Expected estimator instance.");
        object second = System.Activator.CreateInstance(estimatorType, resolver, CancellationToken.None) ?? throw new InvalidOperationException("Expected estimator instance.");
        System.Reflection.MethodInfo resolve = estimatorType.GetMethod("ResolveOpenTypeFont") ?? throw new InvalidOperationException("Expected program resolution bridge.");

        object? firstFont = resolve.Invoke(first, ["TestFont", false, false]);
        object? secondFont = resolve.Invoke(second, ["TestFont", false, false]);

        TestAssert.True(firstFont is not null, "The test program should resolve.");
        TestAssert.True(ReferenceEquals(firstFont, secondFont), "Estimators sharing a resolver must reuse the loaded program instead of reloading per estimator.");
    }

    public static void PptxSlideRenderLoadsEachFontProgramOnce()
    {
        byte[] fontBytes = TestFontBuilder.CreateTestFont();
        var counters = new System.Collections.Concurrent.ConcurrentDictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
        var resolver = new PerFamilyCountingResolver(fontBytes, counters);
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="2" name="First"/><p:nvPr/></p:nvSpPr>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="TestFont"/></a:rPr><a:t>Alpha</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="3" name="Second"/><p:nvPr/></p:nvSpPr>
                      <p:spPr><a:xfrm><a:off x="914400" y="2743200"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="TestFont"/></a:rPr><a:t>Beta</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                    <p:graphicFrame>
                      <p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:nvPr/></p:nvGraphicFramePr>
                      <p:xfrm><a:off x="914400" y="4572000"/><a:ext cx="3657600" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>
                        <a:tr h="914400">
                          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="TestFont"/></a:rPr><a:t>Gamma</a:t></a:r></a:p></a:txBody><a:tcPr anchor="ctr"/></a:tc>
                          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="TestFont"/></a:rPr><a:t>Delta</a:t></a:r></a:p></a:txBody><a:tcPr anchor="ctr"/></a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        int pageCount;
        using (System.IO.FileStream stream = System.IO.File.OpenRead(input))
        {
            OoxPackage package = OoxPackage.Open(stream, System.Threading.CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, System.Threading.CancellationToken.None);
            pageCount = new PptxRenderer(resolver).RenderPages(document, package, diagnosticSink: null, System.Threading.CancellationToken.None).Count;
        }

        TestAssert.Equal(1, pageCount);
        TestAssert.Equal(1, counters.Count);
        TestAssert.Equal("counting:TestFont", counters.Keys.Single());
        TestAssert.Equal(1, counters.Values.Single());
    }

    private sealed class PerFamilyCountingResolver(byte[] fontBytes, System.Collections.Concurrent.ConcurrentDictionary<string, int> counters) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            string family = string.IsNullOrWhiteSpace(request.FamilyName) ? "(default)" : request.FamilyName;
            var source = new CountingProgramSource("counting:" + family, fontBytes, counters);
            return new FontFaceResolution(
                request.FamilyName,
                family,
                new FontStyleKey(request.Bold, request.Italic, request.Bold ? 700 : 400, 0, false),
                source,
                IsFallback: false);
        }

        private sealed class CountingProgramSource(string stableId, byte[] bytes, System.Collections.Concurrent.ConcurrentDictionary<string, int> counters) : IFontProgramSource
        {
            public string StableId => stableId;

            public ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct)
            {
                counters.AddOrUpdate(stableId, 1, (_, count) => count + 1);
                ct.ThrowIfCancellationRequested();
                return ValueTask.FromResult((ReadOnlyMemory<byte>)bytes);
            }
        }
    }
    public static void PptxSharedResolverDeduplicatesIdenticalSubsets()
    {
        var program = new MemoryFontProgramSource("cache-test", TestFontBuilder.CreateTestFont());
        var resolver = new PresentationFontResolver(new SingleProgramResolver(program));
        var resolved = resolver.ResolvePresentationOpenTypeFont(new FontRequest("TestFont", false, false), CancellationToken.None) ?? throw new InvalidOperationException("Expected test program resolution.");

        PdfEmbeddedFont first = resolver.GetOrCreateSubset(resolved.Resolution, resolved.Font, [65, 66, 65], CancellationToken.None);
        PdfEmbeddedFont same = resolver.GetOrCreateSubset(resolved.Resolution, resolved.Font, [66, 65], CancellationToken.None);
        PdfEmbeddedFont other = resolver.GetOrCreateSubset(resolved.Resolution, resolved.Font, [67], CancellationToken.None);

        TestAssert.True(ReferenceEquals(first, same), "Identical codepoint sets must share one subset regardless of order or duplicates.");
        TestAssert.True(!ReferenceEquals(first, other), "Different codepoint sets must keep distinct subsets.");
    }
    private sealed class SingleProgramResolver(MemoryFontProgramSource program) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request)
        {
            return new FontFaceResolution(
                request.FamilyName,
                "TestFont",
                new FontStyleKey(request.Bold, request.Italic, request.Bold ? 700 : 400, 0, false),
                program,
                IsFallback: false);
        }
    }
}
