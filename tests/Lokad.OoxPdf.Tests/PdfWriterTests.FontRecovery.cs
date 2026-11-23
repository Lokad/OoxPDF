using System.Buffers.Binary;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PdfWriterTests
{
    public static void FontResourcesKeepDistinctProgramsDuringWholeFontRecovery()
    {
        foreach (bool wholeFont in new[] { false, true })
        {
            byte[] firstBytes = CreateIdentityFont(wholeFont, alternateAdvance: false);
            byte[] secondBytes = CreateIdentityFont(wholeFont, alternateAdvance: true);
            var first = PdfEmbeddedFont.Create(OpenTypeFont.Load(firstBytes), [65], default);
            var second = PdfEmbeddedFont.Create(OpenTypeFont.Load(secondBytes), [65], default);
            var repeat = PdfEmbeddedFont.Create(OpenTypeFont.Load(firstBytes.ToArray()), [65], default);
            TestAssert.Equal(!wholeFont, first.UsesSubsetFontProgram);
            TestAssert.Equal(!wholeFont, second.UsesSubsetFontProgram);
            TestAssert.Equal(first.BaseFontName, wholeFont ? second.BaseFontName : repeat.BaseFontName);
            TestAssert.True(first.ResourceKey != second.ResourceKey, "Different programs with the same family/map need distinct resources in either embedding mode.");
            TestAssert.Equal(first.ResourceKey, repeat.ResourceKey);
            TestAssert.True(!first.FontProgramBytes.Span.SequenceEqual(second.FontProgramBytes.Span), "Both distinct font payloads must survive preparation.");
            if (wholeFont)
            {
                TestAssert.True(first.FontProgramBytes.Span.SequenceEqual(firstBytes), "Recovery retains the first whole program.");
                TestAssert.True(second.FontProgramBytes.Span.SequenceEqual(secondBytes), "Recovery retains the second whole program.");
            }
            var page = new PdfPage(200, 100, string.Empty,
                [new PdfFontResource("F1", first), new PdfFontResource("F2", second), new PdfFontResource("F3", repeat)]);
            string pdf = WritePdfText([page]);
            TestAssert.Equal(2, pdf.AsSpan().Count("/Subtype /Type0"));
            TestAssert.Equal(2, pdf.AsSpan().Count("/FontFile2 "));
        }
    }

    public static void FontResourceMergeRejectsDifferentProgramsAndKeepsEquivalentCopies()
    {
        foreach (bool wholeFont in new[] { false, true })
        {
            byte[] bytes = CreateIdentityFont(wholeFont, alternateAdvance: false);
            var first = PdfEmbeddedFont.Create(OpenTypeFont.Load(bytes), [65], default);
            var different = PdfEmbeddedFont.Create(OpenTypeFont.Load(CreateIdentityFont(wholeFont, alternateAdvance: true)), [66], default);
            TestAssert.Throws<InvalidDataException>(() => PdfEmbeddedFont.Merge([first, different], default));

            var equivalent = PdfEmbeddedFont.Create(OpenTypeFont.Load(bytes.ToArray()), [66], default);
            var merged = PdfEmbeddedFont.Merge([first, equivalent], default);
            var expected = PdfEmbeddedFont.Create(OpenTypeFont.Load(bytes.ToArray()), [65, 66], default);
            TestAssert.Equal(expected.ResourceKey, merged.ResourceKey);
            TestAssert.Equal(expected.EncodeGlyphHex("AB"), merged.EncodeGlyphHex("AB"));
            TestAssert.True(expected.FontProgramBytes.Span.SequenceEqual(merged.FontProgramBytes.Span), "Equal program copies can still share a merged mapping.");
        }
    }

    private static byte[] CreateIdentityFont(bool wholeFont, bool alternateAdvance)
    {
        byte[] bytes = TestFontBuilder.CreateTestFont();
        if (alternateAdvance)
        {
            int hmtx = TestFontBuilder.GetTableRange(bytes, "hmtx").Offset;
            int glyph = TestFontBuilder.LoadTestFont().MapCodePoint('A');
            BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(hmtx + glyph * 4), 731);
        }
        if (wholeFont)
        {
            int head = TestFontBuilder.GetTableRange(bytes, "head").Offset;
            BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(head + 50), 2);
        }
        return bytes;
    }
}
