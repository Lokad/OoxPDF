using System.Buffers.Binary;
using System.Security;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class DocxFontSlotTests
{
    private const string Properties = "<w:rPr><w:rFonts w:ascii=\"Slot A\" w:hAnsi=\"Slot H\" w:eastAsia=\"Slot E\" w:cs=\"Slot C\"/></w:rPr>";

    public static void DocxFontSlotsKeepSourceOffsetsAndLinkRevisionOwnership()
    {
        string paragraph = "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"target\"/><w:hyperlink w:anchor=\"target\"><w:ins w:id=\"2\" w:author=\"Review\"><w:r>" + Properties +
            "<w:t>AΩ中😀Z</w:t></w:r></w:ins></w:hyperlink><w:bookmarkEnd w:id=\"1\"/></w:p>";
        DocxDocument document = Read(paragraph, "body");
        DocxParagraph p = document.Paragraphs.Single();
        TestAssert.Equal("A|Ω|中|😀|Z", string.Join("|", p.Runs.Select(run => run.Text)));
        TestAssert.Equal("0|1|2|3|5", string.Join("|", p.Runs.Select(run => run.SourceTextOffsetInRun)));
        TestAssert.Equal(1, p.Runs.Select(run => run.SourceRunIndex).Distinct().Count());
        TestAssert.True(p.Runs.All(run => run.Revision is not null), "Fragmenting must retain source revision ownership.");
        TestAssert.Equal(1, p.Hyperlinks.Count);
        TestAssert.Equal(1, p.Hyperlinks[0].SourceRunCount);
        TestAssert.Equal(5, p.Hyperlinks[0].TextRunCount);
        TestAssert.Equal(1, p.BookmarkAnchors.Count);
        TestAssert.Equal("Slot A|Slot H|Slot E|Slot H|Slot A",
            string.Join("|", p.Runs.Select(run => DocxFontResolver.ResolveTypeface(run, document.FontCatalog).Primary)));
    }

    public static void DocxFontSlotsKeepEquivalentDeclarationsTogether()
    {
        string properties = Properties.Replace("Slot H", "Slot A").Replace("Slot E", "Slot A");
        DocxDocument document = Read("<w:p><w:r>"+properties+"<w:t>AΩ中😀Z</w:t></w:r></w:p>", "body");
        TestAssert.Equal(1, document.Paragraphs.Single().Runs.Count);
    }

    public static void DocxFontSlotsResolveDeclaredThemesAndMissingSlotFallbacks()
    {
        var fonts = new DocxRunFonts("Slot A", "Slot H", "Slot E", "Slot C", "majorHAnsi", "minorHAnsi", "majorEastAsia", "minorBidi");
        var catalog = new DocxFontCatalog([], new DocxThemeFonts("Theme A", "Theme H", null, "Theme C", "Theme E", null));
        foreach (var (text, family, theme) in new[] { ("A", "Slot A", "Theme A"), ("ΩЖ", "Slot H", "Theme H"), ("中", "Slot E", "Theme E"), ("א", "Slot C", "Theme C") }) {
            var run = new DocxTextRun(text, 12, null, false, false, false, null, "Slot A") { Fonts=fonts };
            DocxTypefaceCandidates candidates = DocxFontResolver.ResolveTypeface(run, catalog);
            TestAssert.Equal(family, candidates.Primary!);
            TestAssert.Equal(theme, candidates.Theme!);
        }
        var fallback = new DocxTextRun("Ω", 12, null, false, false, false, null, null) { Fonts=fonts with { HighAnsi=null, HighAnsiTheme=null } };
        DocxTypefaceCandidates selected = DocxFontResolver.ResolveTypeface(fallback, catalog);
        TestAssert.Equal("Slot A", selected.Primary!);
        TestAssert.Equal("Theme A", selected.Theme!);
    }

    public static void DocxFontSlotsMeasureAndPaintJoinedSplitTextAcrossContainers()
    {
        var resolver = new SlotResolver();
        foreach (string container in new[] { "body", "cell", "header", "footer" }) {
            string[] pieces = ["Latin ABC 123 ", "ΩЖ", "中", " tail"];
            string Paragraph(IEnumerable<string> texts) => "<w:p>" + string.Concat(texts.Select(text => "<w:r>" + Properties + "<w:t xml:space=\"preserve\">" + SecurityElement.Escape(text) + "</w:t></w:r>")) + "</w:p>";
            DocxDocument joined = Read(Paragraph([string.Concat(pieces)]), container);
            DocxDocument split = Read(Paragraph(pieces), container);
            PdfPage a = new DocxRenderer(resolver).RenderBlankPages(joined, null, default).Single();
            PdfPage b = new DocxRenderer(resolver).RenderBlankPages(split, null, default).Single();
            TestAssert.Equal(a.Content, b.Content);
            TestAssert.True(a.Fonts.Any(font => font.Font.Font.FamilyName == "Slot A") &&
                a.Fonts.Any(font => font.Font.Font.FamilyName == "Slot H") &&
                a.Fonts.Any(font => font.Font.Font.FamilyName == "Slot E"), "All declared faces must reach actual paint resources in " + container);
        }
    }

    private static DocxDocument Read(string paragraph, string container)
    {
        var entries = new Dictionary<string, string> {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["_rels/.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"
        };
        string body = container == "cell" ? "<w:tbl><w:tblPr><w:tblLayout w:type=\"fixed\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"9360\"/></w:tblGrid><w:tr><w:tc>" + paragraph + "</w:tc></w:tr></w:tbl>" : paragraph;
        string reference = "";
        if (container is "header" or "footer") {
            string root = container == "header" ? "hdr" : "ftr";
            entries["word/"+container+"1.xml"] = "<w:"+root+" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"+paragraph+"</w:"+root+">";
            entries["[Content_Types].xml"] = entries["[Content_Types].xml"].Replace("</Types>", "<Override PartName=\"/word/"+container+"1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml."+container+"+xml\"/></Types>");
            entries["word/_rels/document.xml.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"story\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"+container+"\" Target=\""+container+"1.xml\"/></Relationships>";
            reference = "<w:"+container+"Reference w:type=\"default\" r:id=\"story\"/>";
            body = "<w:p/>";
        }
        entries["word/document.xml"] = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" + body + "<w:sectPr>"+reference+"<w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>";
        return DocxTests.ReadDocx(TestFixtures.WriteTempPackage(".docx", entries), OoxPdfDocxMarkupMode.Final);
    }

    private sealed class SlotResolver : IFontResolver
    {
        private readonly Dictionary<string, IFontProgramSource> sources = new(StringComparer.OrdinalIgnoreCase);
        public SlotResolver()
        {
            foreach (var (family, width) in new[] { ("Slot A", 450), ("Slot H", 650), ("Slot E", 1000) }) {
                byte[] bytes = TestFontBuilder.CreateFontWithFamilyNames((0x0409, family));
                var groups = new[] { (0x20, 0x9F, 1), (0x0301, 0x0301, 129), (0x03A9, 0x03A9, 126), (0x0416, 0x0416, 127), (0x4E2D, 0x4E2D, 128), (0x1F600, 0x1F600, 130) };
                byte[] cmap = new byte[28 + groups.Length*12];
                BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(2,2), 1);
                BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(4,2), 3);
                BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(6,2), 10);
                BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(8,4), 12);
                BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(12,2), 12);
                BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(16,4), (uint)(cmap.Length-12));
                BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(24,4), (uint)groups.Length);
                for (int index=0; index < groups.Length; index++) {
                    BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(28+index*12,4), (uint)groups[index].Item1);
                    BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(32+index*12,4), (uint)groups[index].Item2);
                    BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(36+index*12,4), (uint)groups[index].Item3);
                }
                bytes = TestFontBuilder.ReplaceTable(bytes, "cmap", cmap);
                int hmtx = TestFontBuilder.GetTableRange(bytes, "hmtx").Offset;
                for (int glyph=0; glyph < TestFontBuilder.GlyphCount; glyph++)
                    if (glyph != 129) BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(hmtx+glyph*4,2), (ushort)width);
                sources.Add(family, new MemoryFontProgramSource("slots:"+family, bytes));
            }
        }
        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, request.FamilyName,
            new FontStyleKey(request.Bold, request.Italic), sources.GetValueOrDefault(request.FamilyName) ?? sources["Slot A"], IsFallback:false);
    }
}
