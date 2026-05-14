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
}
