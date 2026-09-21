using Lokad.OoxPdf.Imaging;

namespace Lokad.OoxPdf.Tests;

internal static class ImagingTests
{
    public static void SharedDecoderObservesCancellationMidDecode()
    {
        // Q01: long pixel decodes stay cancellable through the shared path.
        byte[] png = TestFixtures.CreateGrayscalePng(256, 256, new byte[256 * 256]);
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();

        TestAssert.Throws<OperationCanceledException>(
            () => OoxImageDecoder.Decode("image/png", png, static rgb => rgb, cancelled.Token));
        TestAssert.Throws<OperationCanceledException>(
            () => PngImage.Read(png, cancelled.Token));
        TestAssert.Throws<OperationCanceledException>(
            () => BmpImage.Read(TestFixtures.CreateRgbBmp(64, 64, new byte[64 * 64 * 3]), cancelled.Token));

        Pdf.PdfImageXObject image = OoxImageDecoder.Decode("image/png", png, static rgb => rgb, CancellationToken.None);
        TestAssert.Equal(256, image.Width);
        TestAssert.Equal(256, image.Height);
    }

    public static void SharedDecoderClassifiesContentTypes()
    {
        // D02: one classification shared by the DOCX/PPTX image paths.
        foreach (string jpeg in new[] { "image/jpeg", "image/jpg", "IMAGE/JPEG" })
        {
            TestAssert.True(OoxImageDecoder.IsJpegContentType(jpeg), "Expected JPEG classification for " + jpeg);
            TestAssert.True(OoxImageDecoder.IsSupportedContentType(jpeg), "Expected supported classification for " + jpeg);
        }

        TestAssert.True(OoxImageDecoder.IsPngContentType("image/png"), "Expected PNG classification.");
        TestAssert.True(OoxImageDecoder.IsBmpContentType("image/bmp"), "Expected BMP classification.");
        TestAssert.True(OoxImageDecoder.IsBmpContentType("image/x-ms-bmp"), "Expected BMP classification.");
        foreach (string other in new[] { "image/gif", "image/svg+xml", "image/webp", "application/octet-stream" })
        {
            TestAssert.True(!OoxImageDecoder.IsSupportedContentType(other), "Expected unsupported classification for " + other);
        }
    }

    public static void SharedDecoderRejectsUnknownContentType()
    {
        // D02: unknown types throw with the exact legacy message so both renderers
        // keep byte-identical IMAGE_UNSUPPORTED_FORMAT diagnostics.
        NotSupportedException pixels = TestAssert.Throws<NotSupportedException>(
            () => OoxImageDecoder.DecodePixels("image/gif", [71, 73, 70]));
        TestAssert.Equal("Unsupported image content type.", pixels.Message);
        NotSupportedException image = TestAssert.Throws<NotSupportedException>(
            () => OoxImageDecoder.Decode("image/gif", [71, 73, 70], static rgb => rgb));
        TestAssert.Equal("Unsupported image content type.", image.Message);
    }

    public static void SharedDecoderPassesThroughJpegDimensions()
    {
        // D02: JPEG stays a DCT passthrough (no pixel decode) through the shared path.
        byte[] jpegHeader =
        [
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            0x00, 0x03,
            0x00, 0x05,
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        ];

        Pdf.PdfImageXObject image = OoxImageDecoder.Decode("image/jpg", jpegHeader, static rgb => rgb);

        TestAssert.Equal(5, image.Width);
        TestAssert.Equal(3, image.Height);
        TestAssert.Equal("/DCTDecode", image.Filter);
        TestAssert.True(ReferenceEquals(jpegHeader, image.Bytes), "JPEG passthrough must keep the original bytes.");
    }

    public static void SharedDecoderDecodesPngAndBmpPixels()
    {
        // D02: PNG/BMP share one pixel path, with the caller recolor map applied.
        byte[] png = TestFixtures.CreateGrayscalePng(3, 1, [0, 128, 255]);
        Pdf.PdfImageXObject plain = OoxImageDecoder.Decode("image/png", png, static rgb => rgb);

        TestAssert.Equal(3, plain.Width);
        TestAssert.Equal(1, plain.Height);
        TestAssert.Equal("/FlateDecode", plain.Filter);

        Pdf.PdfImageXObject inverted = OoxImageDecoder.Decode(
            "image/png", png, rgb => rgb.Select(channel => (byte)(255 - channel)).ToArray());

        TestAssert.True(!plain.Bytes.AsSpan().SequenceEqual(inverted.Bytes), "The recolor map must transform decoded pixels.");

        (int width, int height, byte[] rgb, byte[]? alpha) = OoxImageDecoder.DecodePixels(
            "image/x-ms-bmp", TestFixtures.CreateRgbBmp(2, 1, [255, 0, 0, 0, 0, 255]));

        TestAssert.Equal(2, width);
        TestAssert.Equal(1, height);
        TestAssert.True(rgb.SequenceEqual(new byte[] { 255, 0, 0, 0, 0, 255 }), "BMP pixels must decode through the shared path.");
        TestAssert.True(alpha is null, "24-bit BMP must not produce an alpha channel.");
    }

    public static void JpegInfoReadsDimensions()
    {
        byte[] jpegHeader =
        [
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            0x00, 0x03,
            0x00, 0x05,
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        ];

        JpegInfo info = JpegInfo.Read(jpegHeader);

        TestAssert.Equal(5, info.Width);
        TestAssert.Equal(3, info.Height);
        TestAssert.Equal(8, info.BitsPerComponent);
        TestAssert.Equal(3, info.ComponentCount);
        TestAssert.Equal(0xC0, info.FrameMarker);
        TestAssert.True(info.IsBaselineDct, "SOF0 should be classified as baseline DCT.");
        TestAssert.True(!info.IsProgressiveDct, "SOF0 should not be classified as progressive DCT.");
        TestAssert.Equal("baseline DCT", info.FrameProfileName);
    }

    public static void JpegInfoReadsGrayscaleFrameMetadata()
    {
        byte[] jpegHeader =
        [
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x0B,
            0x08,
            0x00, 0x01,
            0x00, 0x01,
            0x01,
            0x01, 0x11, 0x00,
            0xFF, 0xD9
        ];

        JpegInfo info = JpegInfo.Read(jpegHeader);

        TestAssert.Equal(1, info.Width);
        TestAssert.Equal(1, info.Height);
        TestAssert.Equal(8, info.BitsPerComponent);
        TestAssert.Equal(1, info.ComponentCount);
        TestAssert.Equal(0xC0, info.FrameMarker);
        TestAssert.True(info.IsBaselineDct, "SOF0 grayscale should be classified as baseline DCT.");
    }

    public static void JpegInfoClassifiesProgressiveFrameMetadata()
    {
        byte[] jpegHeader =
        [
            0xFF, 0xD8,
            0xFF, 0xC2,
            0x00, 0x11,
            0x08,
            0x00, 0x02,
            0x00, 0x04,
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        ];

        JpegInfo info = JpegInfo.Read(jpegHeader);

        TestAssert.Equal(4, info.Width);
        TestAssert.Equal(2, info.Height);
        TestAssert.Equal(0xC2, info.FrameMarker);
        TestAssert.True(!info.IsBaselineDct, "SOF2 should not be classified as baseline DCT.");
        TestAssert.True(info.IsProgressiveDct, "SOF2 should be classified as progressive DCT.");
        TestAssert.Equal("progressive DCT", info.FrameProfileName);
    }

    public static void JpegImageDecodesBaselineDctPixels()
    {
        byte[] jpeg = Convert.FromBase64String(
            "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAIDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD+Rb4g/wDI++N/+xv8S/8Ap5vaKKK/7o/An/kyHg3/ANmq8PP/AFkcoPyHxn/5PD4r/wDZyuOv/WozQ//Z");

        JpegImage image = JpegImage.Read(jpeg);

        TestAssert.Equal(2, image.Width);
        TestAssert.Equal(1, image.Height);
        TestAssert.True(
            image.Rgb.SequenceEqual(new byte[] { 150, 24, 150, 103, 0, 103 }),
            "Baseline JPEG decoder should match the public Windows-decoded 2x1 oracle pixels.");
    }

    public static void PngImageReadsIndexedPalettePixels()
    {
        byte[] png = TestFixtures.CreateIndexedPng(
            2,
            1,
            [255, 0, 0, 0, 0, 255],
            [0, 1]);

        PngImage image = PngImage.Read(png);

        TestAssert.Equal(2, image.Width);
        TestAssert.Equal(1, image.Height);
        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 255, 0, 0, 0, 0, 255 }), "Indexed PNG palette should expand to RGB pixels.");
    }

    public static void PngImageReadsPackedIndexedPalettePixels()
    {
        byte[] png = TestFixtures.CreatePackedIndexedPng(
            4,
            1,
            4,
            [0, 0, 0, 255, 0, 0, 0, 255, 0],
            [0, 1, 2, 1]);

        PngImage image = PngImage.Read(png);

        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 0, 0, 0, 255, 0, 0, 0, 255, 0, 255, 0, 0 }), "Packed indexed PNG samples should expand to RGB pixels.");
    }

    public static void PngImageReadsGrayscalePixels()
    {
        byte[] png = TestFixtures.CreateGrayscalePng(3, 1, [0, 128, 255]);

        PngImage image = PngImage.Read(png);

        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 0, 0, 0, 128, 128, 128, 255, 255, 255 }), "Grayscale PNG samples should expand to RGB pixels.");
    }

    public static void PngImageReadsAdam7TruecolorAlphaPixels()
    {
        byte[] rgba =
        [
            255, 0, 0, 255,
            0, 255, 0, 128,
            0, 0, 255, 64,
            255, 255, 255, 0,
            10, 20, 30, 255,
            40, 50, 60, 255,
            70, 80, 90, 255,
            100, 110, 120, 255,
            130, 140, 150, 255,
            160, 170, 180, 255,
            190, 200, 210, 255,
            220, 230, 240, 255
        ];
        byte[] png = TestFixtures.CreateInterlacedRgbaPng(4, 3, rgba);

        PngImage image = PngImage.Read(png);

        TestAssert.Equal(4, image.Width);
        TestAssert.Equal(3, image.Height);
        TestAssert.True(image.Rgb.SequenceEqual(new byte[]
        {
            255, 0, 0,
            0, 255, 0,
            0, 0, 255,
            255, 255, 255,
            10, 20, 30,
            40, 50, 60,
            70, 80, 90,
            100, 110, 120,
            130, 140, 150,
            160, 170, 180,
            190, 200, 210,
            220, 230, 240
        }), "Adam7 RGBA PNG should expand to RGB pixels in final image order.");
        TestAssert.True(image.Alpha is not null && image.Alpha.SequenceEqual(new byte[] { 255, 128, 64, 0, 255, 255, 255, 255, 255, 255, 255, 255 }), "Adam7 RGBA PNG should preserve alpha in final image order.");
    }

    public static void BmpImageReadsBottomUpRgbPixels()
    {
        byte[] bmp = TestFixtures.CreateRgbBmp(2, 1, [255, 0, 0, 0, 0, 255]);

        BmpImage image = BmpImage.Read(bmp);

        TestAssert.Equal(2, image.Width);
        TestAssert.Equal(1, image.Height);
        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 255, 0, 0, 0, 0, 255 }), "BMP BGR pixels should expand to RGB pixels.");
        TestAssert.True(image.Alpha is null, "24-bit BMP should not produce an alpha channel.");
    }

    public static void BmpImageIgnoresRgb32AlphaByte()
    {
        byte[] bmp = TestFixtures.CreateRgbaBmp(1, 2, [255, 0, 0, 128, 0, 0, 255, 64]);

        BmpImage image = BmpImage.Read(bmp);

        TestAssert.Equal(1, image.Width);
        TestAssert.Equal(2, image.Height);
        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 255, 0, 0, 0, 0, 255 }), "32-bit BMP BGRA pixels should expand to RGB pixels.");
        TestAssert.True(image.Alpha is null, "Office treats 32-bit BI_RGB BMP alpha bytes as unused.");
    }
    public static void RejectsPngLyingDimensions()
    {
        byte[] png = BuildPng(100000, 100000, 8, 2, 0);
        TestAssert.Throws<InvalidDataException>(() => PngImage.Read(png));
    }

    public static void RejectsPngChunkBeyondEndOfData()
    {
        var bytes = new List<byte> { 137, 80, 78, 71, 13, 10, 26, 10 };
        WritePngChunk(bytes, "IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 });
        WritePngChunkHeaderOnly(bytes, "IDAT", 1000000);
        TestAssert.Throws<InvalidDataException>(() => PngImage.Read(bytes.ToArray()));
    }

    public static void RejectsJpegLyingDimensions()
    {
        var bytes = new List<byte> { 0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0x11, 0x08, 0x9C, 0x40, 0x9C, 0x40, 0x03 };
        for (int component = 1; component <= 3; component++)
        {
            bytes.Add((byte)component);
            bytes.Add(0x11);
            bytes.Add(0x00);
        }

        bytes.Add(0xFF);
        bytes.Add(0xD9);
        TestAssert.Throws<InvalidDataException>(() => JpegImage.Read(bytes.ToArray()));
    }

    public static void RejectsBmpLyingDimensions()
    {
        byte[] bmp = BuildBmpHeader(100000, 1, 24);
        TestAssert.Throws<InvalidDataException>(() => BmpImage.Read(bmp));
    }

    public static void RejectsBmpMinimumHeight()
    {
        byte[] bmp = BuildBmpHeader(1, int.MinValue, 24);
        TestAssert.Throws<InvalidDataException>(() => BmpImage.Read(bmp));
    }

    public static void PngStoredBlocksDecodeToIdenticalPixels()
    {
        // Positive control for the in-place inflate-buffer decode (M08): stored
        // (uncompressed) deflate rows must expand to the same pixels.
        byte[] payload = [0, 10, 20, 30, 40, 50, 60, 0, 70, 80, 90, 100, 110, 120];
        byte[] png = BuildStoredPng(2, 2, 8, 2, 0, payload);
        PngImage image = PngImage.Read(png);
        TestAssert.True(image.Rgb.SequenceEqual(new byte[] { 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120 }), "Stored-deflate rows must decode to identical pixels.");
    }

    public static void PngTruncatedFilterByteThrowsIndexOutOfRange()
    {
        // Only the first of two rows is present: the missing filter byte must fail
        // (not render garbage), preserving the crop-fallback contract.
        byte[] payload = [0, 10, 20, 30, 40, 50, 60];
        byte[] png = BuildStoredPng(2, 2, 8, 2, 0, payload);
        TestAssert.Throws<IndexOutOfRangeException>(() => PngImage.Read(png));
    }

    public static void PngTruncatedRowSpanThrowsArgumentOutOfRange()
    {
        // The second filter byte is present but its row is short: the span copy must
        // fail (not render garbage), preserving the crop-fallback contract.
        byte[] payload = [0, 10, 20, 30, 40, 50, 60, 0, 70, 80, 90];
        byte[] png = BuildStoredPng(2, 2, 8, 2, 0, payload);
        TestAssert.Throws<ArgumentOutOfRangeException>(() => PngImage.Read(png));
    }

    public static void PngInterlacedTruncatedSpanThrowsArgumentOutOfRange()
    {
        byte[] payload = [0, 10, 20, 30];
        byte[] png = BuildStoredPng(1, 1, 8, 6, 1, payload);
        TestAssert.Throws<ArgumentOutOfRangeException>(() => PngImage.Read(png));
    }

    private static byte[] BuildStoredPng(int width, int height, int bitDepth, int colorType, int interlace, byte[] rawRows)
    {
        var bytes = new List<byte> { 137, 80, 78, 71, 13, 10, 26, 10 };
        WritePngChunk(bytes, "IHDR", new byte[]
        {
            (byte)(width >> 24), (byte)(width >> 16), (byte)(width >> 8), (byte)width,
            (byte)(height >> 24), (byte)(height >> 16), (byte)(height >> 8), (byte)height,
            (byte)bitDepth, (byte)colorType, 0, 0, (byte)interlace,
        });
        WritePngChunk(bytes, "IDAT", BuildStoredDeflate(rawRows));
        WritePngChunk(bytes, "IEND", Array.Empty<byte>());
        return bytes.ToArray();
    }

    private static byte[] BuildStoredDeflate(byte[] raw)
    {
        // Minimal zlib stream with one final stored block; the reader never checks
        // Adler-32, but ZLibStream validates it, so compute it for real.
        var output = new List<byte> { 0x78, 0x01, 0x01 };
        output.Add((byte)raw.Length);
        output.Add((byte)(raw.Length >> 8));
        output.Add((byte)(~raw.Length));
        output.Add((byte)(~raw.Length >> 8));
        output.AddRange(raw);
        uint checksum = Adler32(raw);
        output.Add((byte)(checksum >> 24));
        output.Add((byte)(checksum >> 16));
        output.Add((byte)(checksum >> 8));
        output.Add((byte)checksum);
        return output.ToArray();
    }

    private static uint Adler32(byte[] data)
    {
        uint a = 1;
        uint b = 0;
        foreach (byte value in data)
        {
            a = (a + value) % 65521;
            b = (b + a) % 65521;
        }

        return (b << 16) | a;
    }

    private static byte[] BuildPng(int width, int height, int bitDepth, int colorType, int interlace)
    {
        var bytes = new List<byte> { 137, 80, 78, 71, 13, 10, 26, 10 };
        WritePngChunk(bytes, "IHDR", new byte[]
        {
            (byte)(width >> 24), (byte)(width >> 16), (byte)(width >> 8), (byte)width,
            (byte)(height >> 24), (byte)(height >> 16), (byte)(height >> 8), (byte)height,
            (byte)bitDepth, (byte)colorType, 0, 0, (byte)interlace,
        });
        WritePngChunk(bytes, "IDAT", Array.Empty<byte>());
        WritePngChunk(bytes, "IEND", Array.Empty<byte>());
        return bytes.ToArray();
    }

    private static void WritePngChunk(List<byte> bytes, string type, byte[] data)
    {
        bytes.Add((byte)(data.Length >> 24));
        bytes.Add((byte)(data.Length >> 16));
        bytes.Add((byte)(data.Length >> 8));
        bytes.Add((byte)data.Length);
        foreach (char c in type)
        {
            bytes.Add((byte)c);
        }

        bytes.AddRange(data);
        bytes.AddRange(new byte[] { 0, 0, 0, 0 });
    }

    private static void WritePngChunkHeaderOnly(List<byte> bytes, string type, int length)
    {
        bytes.Add((byte)(length >> 24));
        bytes.Add((byte)(length >> 16));
        bytes.Add((byte)(length >> 8));
        bytes.Add((byte)length);
        foreach (char c in type)
        {
            bytes.Add((byte)c);
        }
    }

    private static byte[] BuildBmpHeader(int width, int height, int bitsPerPixel)
    {
        byte[] bmp = new byte[54];
        bmp[0] = (byte)66;
        bmp[1] = (byte)77;
        BitConverter.GetBytes(54).CopyTo(bmp, 2);
        BitConverter.GetBytes(54).CopyTo(bmp, 10);
        BitConverter.GetBytes(40).CopyTo(bmp, 14);
        BitConverter.GetBytes(width).CopyTo(bmp, 18);
        BitConverter.GetBytes(height).CopyTo(bmp, 22);
        BitConverter.GetBytes((ushort)1).CopyTo(bmp, 26);
        BitConverter.GetBytes((ushort)bitsPerPixel).CopyTo(bmp, 28);
        return bmp;
    }
}

