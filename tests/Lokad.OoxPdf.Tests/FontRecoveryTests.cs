using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class FontRecoveryTests
{
    public static void FontResolutionLimitsPreserveFilePublication()
    {
        foreach (OoxPdfInputKind kind in new[] { OoxPdfInputKind.Pptx, OoxPdfInputKind.Docx })
            WithInput(kind, (input, directory) => {
                var expected = new OoxPdfLimitExceededException("font admission refused",
                    OoxPdfLimitKind.MemoryAdmission, OoxPdfLimitScope.SharedResource);
                var resolver = new OneFailureResolver(expected);
                string output = Path.Combine(directory, "result.pdf");
                File.WriteAllText(output, "previous");
                var diagnostics = new DiagnosticCollector();
                Exception actual = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output,
                    new OoxPdfOptions { FontResolver=resolver, DiagnosticSink=diagnostics.Add }));
                TestAssert.True(ReferenceEquals(expected, actual), "Font admission refusal must preserve the original exception.");
                TestAssert.Equal(1, resolver.Calls);
                TestAssert.Equal("previous", File.ReadAllText(output));
                TestAssert.Equal(0, Directory.GetFiles(directory, "result.pdf.tmp-*").Length);
                TestAssert.True(diagnostics.Diagnostics.All(d => d.Id != "PPTX_NODE_RENDER_FAILED"), "Admission refusal is not malformed content.");
            });
    }

    public static void FontResolutionLimitsLeaveStreamsEmptyAndOpen()
    {
        foreach (OoxPdfInputKind kind in new[] { OoxPdfInputKind.Pptx, OoxPdfInputKind.Docx })
            WithInput(kind, (path, _) => {
                var expected = new OoxPdfLimitExceededException("font admission refused",
                    OoxPdfLimitKind.MemoryAdmission, OoxPdfLimitScope.SharedResource);
                var resolver = new OneFailureResolver(expected);
                using FileStream input = File.OpenRead(path);
                using var output = new MemoryStream();
                Exception actual = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output,
                    new OoxPdfOptions { InputKind=kind, FontResolver=resolver }));
                TestAssert.True(ReferenceEquals(expected, actual), "Stream conversion must retain fatal font failures.");
                TestAssert.Equal(1, resolver.Calls);
                TestAssert.Equal(0L, output.Length);
                TestAssert.True(input.CanRead && output.CanWrite, "Caller-owned streams remain open.");
            });
    }

    public static void FontResolutionKeepsOrdinaryPptxIoRecovery()
    {
        WithInput(OoxPdfInputKind.Pptx, (path, _) => {
            foreach (Exception? failure in new Exception?[] { null, new IOException("unavailable font") }) {
                var resolver = new OneFailureResolver(failure);
                using FileStream input = File.OpenRead(path);
                using var output = new MemoryStream();
                OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind=OoxPdfInputKind.Pptx, FontResolver=resolver });
                TestAssert.True(Encoding.ASCII.GetString(output.ToArray(), 0, 5) == "%PDF-", "Ordinary font I/O still permits the existing fallback.");
                TestAssert.True(resolver.Calls > 1, "Recovery must continue resolving usable paint fonts.");
            }
        });
    }

    public static void FontResolutionCancellationStillPropagates()
    {
        WithInput(OoxPdfInputKind.Pptx, (path, directory) => {
            var expected = new OperationCanceledException("font cancellation");
            string outputPath = Path.Combine(directory, "result.pdf");
            File.WriteAllText(outputPath, "previous");
            Exception actual = TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(path, outputPath,
                new OoxPdfOptions { FontResolver=new OneFailureResolver(expected) }));
            TestAssert.True(ReferenceEquals(expected, actual), "Font cancellation is not a missing-font result.");
            TestAssert.Equal("previous", File.ReadAllText(outputPath));
            using FileStream input = File.OpenRead(path);
            using var output = new MemoryStream();
            actual = TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(input, output,
                new OoxPdfOptions { InputKind=OoxPdfInputKind.Pptx, FontResolver=new OneFailureResolver(expected) }));
            TestAssert.True(ReferenceEquals(expected, actual), "Stream cancellation preserves the original failure.");
            TestAssert.Equal(0L, output.Length);
        });
    }

    public static void HiddenDocxGlyphDiagnosticsFollowEffectiveVisibility()
    {
        foreach (OoxPdfDocxMarkupMode mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        foreach (var (hidden, text, warns) in new (string?, string, bool)[] {
            (null,"",false), ("<w:vanish/>","\u0378",false), (null,"\u0378",true),
            ("<w:vanish/>","Visible text",false), ("<w:vanish w:val=\"0\"/>","\u0378",true)
        }) {
            string path = CreateDocx(hidden, text);
            try {
                DocxDocument document = DocxTests.ReadDocx(path, mode);
                var diagnostics = new DiagnosticCollector();
                var renderer = new DocxRenderer(new TestFaceFontResolver(), mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
                TestAssert.Equal(1, renderer.RenderBlankPages(document, diagnostics.Add, default).Count());
                TestAssert.Equal(warns ? 1 : 0, diagnostics.Diagnostics.Count(d => d.Id == "FONT_MISSING_GLYPHS"));
                TestAssert.Equal(warns, diagnostics.HasWarningsOrErrors);
                if (hidden == "<w:vanish/>") {
                    var source = document.Paragraphs.Single().Runs.Last();
                    TestAssert.Equal(text, source.Text);
                    TestAssert.True(source.EffectiveProperties.Hidden, "The source model must retain hidden text and its flag.");
                }
            }
            finally { File.Delete(path); }
        }
    }

    private static void WithInput(OoxPdfInputKind kind, Action<string,string> check)
    {
        string directory = Path.Combine(Path.GetTempPath(), "font-recovery-"+Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = kind == OoxPdfInputKind.Pptx ? CreatePptx() : CreateDocx(null, "");
        try { check(path, directory); }
        finally { File.Delete(path); Directory.Delete(directory, recursive:true); }
    }

    private static string CreatePptx() => TestFixtures.WriteTempPackage(".pptx", new Dictionary<string,string> {
        ["[Content_Types].xml"]=PptxTests.BasicContentTypes(),
        ["_rels/.rels"]=PptxTests.PackageRelationship(),
        ["ppt/_rels/presentation.xml.rels"]=PptxTests.PresentationRelationship(),
        ["ppt/presentation.xml"]=PptxTests.BasicPresentation(),
        ["ppt/slides/slide1.xml"]="""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="1" name="text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
                <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"><a:latin typeface="TestFace"/></a:rPr><a:t>Visible text</a:t></a:r></a:p></p:txBody>
              </p:sp></p:spTree></p:cSld>
            </p:sld>
            """
    });

    private static string CreateDocx(string? hidden, string text)
    {
        const string fonts="<w:rFonts w:ascii=\"TestFace\" w:hAnsi=\"TestFace\" w:eastAsia=\"TestFace\"/><w:sz w:val=\"24\"/>";
        string extra = text.Length == 0 ? "" : "<w:r><w:rPr>"+fonts+hidden+"</w:rPr><w:t>"+text+"</w:t></w:r>";
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string,string> {
            ["[Content_Types].xml"]="<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["_rels/.rels"]="<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/document.xml"]="<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:rPr>"+fonts+"</w:rPr><w:t>Visible text</w:t></w:r>"+extra+"</w:p><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr></w:body></w:document>"
        });
    }

    private sealed class OneFailureResolver(Exception? failure) : IFontResolver
    {
        private readonly TestFaceFontResolver usable = new();
        public int Calls { get; private set; }
        public FontFaceResolution Resolve(FontRequest request)
        {
            if (++Calls == 1 && failure is not null) throw failure;
            return usable.Resolve(request);
        }
    }
}
