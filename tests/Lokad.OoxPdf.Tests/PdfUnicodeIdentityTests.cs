using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PdfUnicodeIdentityTests
{
    private const string AliasedText = "A\u0391\U0001F600";

    public static void PdfSourceUnicodeAliasesHaveDistinctCidsAndSharedOutlines()
    {
        foreach (bool whole in new[] { false, true })
        {
            var font = AliasFont(whole);
            var points = AliasedText.EnumerateRunes().Select(rune => rune.Value).ToArray();
            var embedded = PdfEmbeddedFont.Create(font, points, default);
            var reversed = PdfEmbeddedFont.Create(font, points.Reverse(), default);
            TestAssert.Equal(!whole, embedded.UsesSubsetFontProgram);
            TestAssert.Equal(embedded.ResourceKey, reversed.ResourceKey);
            TestAssert.True(embedded.FontProgramBytes.Span.SequenceEqual(reversed.FontProgramBytes.Span), "Source order cannot change the canonical font mapping.");
            TestAssert.True(ReferenceEquals(embedded, PdfEmbeddedFont.Merge([embedded, reversed], default)), "Equivalent aliases reuse one font resource.");
            ushort glyph = font.MapCodePoint('A');
            TestAssert.True(embedded.TryGetEncodedCid(glyph, out ushort outline), "The ordinary outline retains its original encoding.");
            var cids = points.Select(point => {
                TestAssert.Equal(glyph, font.MapCodePoint(point));
                TestAssert.True(embedded.TryGetEncodedCid(glyph, point, out ushort cid), "Each source identity owns a CID.");
                TestAssert.Equal(point, embedded.UnicodeByCid[cid]);
                return cid;
            }).ToArray();
            TestAssert.Equal(3, cids.Distinct().Count());
            string expected = string.Concat(cids.Select(cid => cid.ToString("X4", CultureInfo.InvariantCulture)));
            TestAssert.Equal(expected, embedded.EncodeGlyphHex(AliasedText));
            TestAssert.Equal(expected, reversed.EncodeGlyphHex(AliasedText));
            TestAssert.Equal("[<" + expected + ">]", embedded.EncodeGlyphPositioningArray(AliasedText, 0, 24, true, false));
            byte[] map = ReadGlyphMap(embedded);
            TestAssert.Equal(embedded.CidToGidMapLength, map.Length);
            string widths = PdfWriterTests.ReadWidthArray(embedded);
            foreach (ushort cid in cids)
            {
                TestAssert.Equal(outline, BinaryPrimitives.ReadUInt16BigEndian(map.AsSpan(2 * cid)));
                TestAssert.Contains(cid + " [" + Math.Round(font.GetAdvanceWidth(glyph) * 1000d / font.UnitsPerEm) + "]", widths);
            }
            PdfEmbeddedFont merged = PdfEmbeddedFont.Merge([PdfEmbeddedFont.Create(font, ['A'], default),
                PdfEmbeddedFont.Create(font, [0x391, 0x1F600], default)], default);
            TestAssert.Equal(expected, merged.EncodeGlyphHex(AliasedText));
        }
    }

    public static void PdfSourceUnicodeAliasesRetainSharedSpacesAndExplicitMathGlyphs()
    {
        var font = AliasFont();
        ushort a = font.MapCodePoint('A'), b = font.MapCodePoint('B');
        var explicitGlyphs = PdfEmbeddedFont.CreateForGlyphMappings(font, [(a, 65), (a, 0x391), (b, 65)], default);
        TestAssert.True(explicitGlyphs.TryGetEncodedCid(a, 65, out ushort ordinary), "The ordinary source glyph encodes.");
        TestAssert.True(explicitGlyphs.TryGetEncodedCid(a, 0x391, out ushort alias), "The distinct source letter encodes.");
        TestAssert.True(explicitGlyphs.TryGetEncodedCid(b, out ushort variant), "The explicit math variant encodes.");
        TestAssert.True(ordinary != alias && alias != variant && ordinary != variant, "A source alias does not replace an explicit variant outline.");
        TestAssert.Equal(65, explicitGlyphs.UnicodeByCid[variant]);
        byte[] map = ReadGlyphMap(explicitGlyphs);
        TestAssert.Equal(ordinary, BinaryPrimitives.ReadUInt16BigEndian(map.AsSpan(2 * alias)));
        TestAssert.Equal(variant, BinaryPrimitives.ReadUInt16BigEndian(map.AsSpan(2 * variant)));
        foreach (int[] order in new[] { new[] { 0x20, 0xA0 }, new[] { 0xA0, 0x20 } })
        {
            var spaces = PdfEmbeddedFont.Create(font, order, default);
            TestAssert.Equal(0, spaces.CidToGidMapLength);
            ushort glyph = font.MapCodePoint(' ');
            TestAssert.Equal(glyph, font.MapCodePoint(0xA0));
            TestAssert.True(spaces.TryGetEncodedCid(glyph, out ushort cid), "The shared space still encodes.");
            string encoded = cid.ToString("X4", CultureInfo.InvariantCulture);
            TestAssert.Equal(encoded + encoded, spaces.EncodeGlyphHex(" \u00A0"));
            TestAssert.Equal(0x20, spaces.UnicodeByCid[cid]);
            string toUnicode = PdfWriterTests.ReadToUnicodeCMap(spaces);
            TestAssert.Contains("<" + encoded + "> <0020>", toUnicode);
            TestAssert.DoesNotContain("<00A0>", toUnicode);
        }
    }

    public static void PdfSourceUnicodeAliasesPreserveCancellationAdmissionAndLifetime()
    {
        var font = AliasFont();
        foreach (bool cancel in new[] { false, true })
        {
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                if (after == 0) return false;
                if (cancel) cancellation.Cancel();
                return !cancel;
            }};
            using (OoxConversionBudget.BeginScope(null, pool))
            {
                void Create() => PdfEmbeddedFont.Create(font, [65, 0x391, 0x1F600], cancellation.Token);
                if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Create).CancellationToken);
                else TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(Create).Kind);
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
        var admitted = new TestMemoryBudget();
        using (OoxConversionBudget.BeginScope(null, admitted))
        {
            var embedded = PdfEmbeddedFont.Create(font, [65, 0x391, 0x1F600], default);
            TestAssert.True(admitted.Reserved > 0, "Font aliases remain funded through their resource lifetime.");
            TestAssert.True(ReadGlyphMap(embedded).Length > 0, "The admitted lookup emits the real CID map.");
            using var untouched = new MemoryStream();
            TestAssert.Throws<OperationCanceledException>(() => embedded.WriteCidToGidMapAsync(
                new PdfObjectWriter(untouched, default, new byte[4096]), new CancellationToken(true)).GetAwaiter().GetResult());
            TestAssert.Equal(0L, untouched.Length);
        }
        TestAssert.Equal(0L, admitted.Reserved);
        TestAssert.Equal(admitted.ScopesCreated, admitted.ScopesDisposed);
        TestAssert.Throws<OoxPdfLimitExceededException>(() => PdfEmbeddedFont.CreateForGlyphMappings(font,
            Enumerable.Range(0x10000, ushort.MaxValue + 1).Select(point => (font.MapCodePoint('A'), point)), default));
    }

    public static void PdfSourceUnicodeAliasesSurviveDocxAndPptxPaintPaths()
    {
        var resolver = new AliasResolver();
        foreach (bool docx in new[] { false, true })
        foreach (bool split in new[] { false, true })
        {
            using var source = Input(docx, split);
            using var package = OoxPackage.Open(source, default);
            var diagnostics = new List<OoxPdfDiagnostic>();
            PdfPage page = docx
                ? new DocxRenderer(resolver, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                    .RenderPagesAsync(new DocxReader().Read(package, diagnostics.Add, default, OoxPdfDocxMarkupMode.Final), diagnostics.Add, default).ToBlockingEnumerable().Single()
                : new PptxRenderer(resolver).RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default).ToBlockingEnumerable().Single();
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false).Select(show => show.Text));
            TestAssert.Equal(AliasedText, text.Trim());
            TestAssert.Equal(0, diagnostics.Count);
            PdfContentValidator.ValidatePage(page, 0, default);
            using var output = new MemoryStream();
            PdfDocumentWriter.WriteStagedAsync(output, AsyncEnumerable.ToAsyncEnumerable([page]), new OoxConversionLimits(), null, default).GetAwaiter().GetResult();
            TestAssert.Contains("/CIDToGIDMap ", Encoding.ASCII.GetString(output.ToArray()));
            TestAssert.DoesNotContain("/CIDToGIDMap /Identity", Encoding.ASCII.GetString(output.ToArray()));
        }
    }

    private static byte[] ReadGlyphMap(PdfEmbeddedFont font)
    {
        using var output = new MemoryStream();
        font.WriteCidToGidMapAsync(new PdfObjectWriter(output, default, new byte[4096]), default).GetAwaiter().GetResult();
        return output.ToArray();
    }

    public static void PdfSourceUnicodeAliasesSurviveLiteralMathPreparation()
    {
        foreach (bool split in new[] { false, true })
        {
            using var input = Input(docx: true, split, math: true);
            using var package = OoxPackage.Open(input, default);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var document = new DocxReader().Read(package, diagnostics.Add, default, OoxPdfDocxMarkupMode.Final);
            var page = new DocxRenderer(new AliasResolver(), OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
                .RenderPagesAsync(document, diagnostics.Add, default).ToBlockingEnumerable().Single();
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, true, honorActualText: false,
                fontResourceFilter: name => name.StartsWith("Eq", StringComparison.Ordinal)).Select(show => show.Text));
            TestAssert.Equal(AliasedText, text);
            TestAssert.True(diagnostics.All(d => d.Id == "DOCX_APPROXIMATED_EQUATION"), "Supported literal equations retain only their usual approximation diagnostic.");
            PdfContentValidator.ValidatePage(page, 0, default);
        }
    }

    private static OpenTypeFont AliasFont(bool whole = false) => OpenTypeFont.Load(AliasBytes(whole));
    private static byte[] AliasBytes(bool whole = false)
    {
        var source = TestFontBuilder.LoadTestFont();
        var points = new[] { (0x20, source.MapCodePoint(0x20)), (0x3F, source.MapCodePoint(0x3F)),
            (0x41, source.MapCodePoint(0x41)), (0x42, source.MapCodePoint(0x42)), (0xA0, source.MapCodePoint(0x20)),
            (0x391, source.MapCodePoint(0x41)), (0x1F600, source.MapCodePoint(0x41)) };
        byte[] cmap = new byte[12 + 16 + 12 * points.Length];
        BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(2), 1);
        BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(4), 3);
        BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(6), 10);
        BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(8), 12);
        BinaryPrimitives.WriteUInt16BigEndian(cmap.AsSpan(12), 12);
        BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(16), (uint)cmap.Length - 12);
        BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(24), (uint)points.Length);
        for (int i = 0; i < points.Length; i++)
        {
            int offset = 28 + 12 * i;
            BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(offset), (uint)points[i].Item1);
            BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(offset + 4), (uint)points[i].Item1);
            BinaryPrimitives.WriteUInt32BigEndian(cmap.AsSpan(offset + 8), points[i].Item2);
        }
        byte[] bytes = TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "cmap", cmap);
        if (whole) BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(TestFontBuilder.GetTableRange(bytes, "head").Offset + 50), 2);
        return bytes;
    }

    private sealed class AliasResolver : IFontResolver
    {
        private readonly MemoryFontProgramSource source = new("unicode-alias-test", AliasBytes());
        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, "TestFont", new FontStyleKey(), source, false);
    }

    private static MemoryStream Input(bool docx, bool split, bool math = false)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        string[] text = split ? AliasedText.EnumerateRunes().Select(rune => rune.ToString()).ToArray() : [AliasedText];
        string paragraph = math
            ? "<w:p xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath>" +
                string.Concat(text.Select(value => "<m:r><m:rPr><m:nor/></m:rPr><w:rPr><w:rFonts w:ascii=\"TestFont\" w:hAnsi=\"TestFont\"/><w:sz w:val=\"48\"/></w:rPr><m:t>" + value + "</m:t></m:r>")) + "</m:oMath></w:p>"
            : "<w:p>" + string.Concat(text.Select(value => "<w:r><w:rPr><w:rFonts w:ascii=\"TestFont\" w:hAnsi=\"TestFont\"/><w:sz w:val=\"48\"/></w:rPr><w:t>" + value + "</w:t></w:r>")) + "</w:p>";
        if (docx) return TestFixtures.CreateZipPackage(new Dictionary<string, string> {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["_rels/.rels"] = $"<Relationships xmlns=\"{rel}\"><Relationship Id=\"main\" Type=\"{office}/officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/document.xml"] = $"<w:document xmlns:w=\"{w}\"><w:body>" + paragraph + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr></w:body></w:document>"
        });
        return TestFixtures.CreateZipPackage(new Dictionary<string, string> {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(), ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(), ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/slides/slide1.xml"] = $"<p:sld xmlns:p=\"{OoxNamespaces.PresentationNamespace}\" xmlns:a=\"{OoxNamespaces.DrawingNamespace}\"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Aliases\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"3657600\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\"rect\"/><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>" +
                string.Concat(text.Select(value => "<a:r><a:rPr sz=\"2400\"><a:latin typeface=\"TestFont\"/></a:rPr><a:t>" + value + "</a:t></a:r>")) + "</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"
        });
    }
}
