using Lokad.OoxPdf.Imaging;

namespace Lokad.OoxPdf.Tests;

internal static class ImagingTests
{
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
}
