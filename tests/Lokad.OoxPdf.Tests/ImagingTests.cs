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
}
