using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Imaging;

/// <summary>
/// Shared Office content-image decoding (D02). JPEG passes through as DCTDecode;
/// PNG and BMP decode to RGB(A) pixels. A caller-supplied recolor map covers the
/// PPTX recolor case; DOCX passes the identity. Per-format diagnostics (with
/// page/slide location and severity) and PPTX crop/recolor policy stay in the
/// renderers. Throws <see cref="InvalidDataException"/> for corrupt data and
/// <see cref="NotSupportedException"/> for unknown content types.
/// </summary>
internal static class OoxImageDecoder
{
    public static bool IsJpegContentType(string contentType) =>
        contentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase);

    public static bool IsPngContentType(string contentType) =>
        contentType.Equals("image/png", StringComparison.OrdinalIgnoreCase);

    public static bool IsBmpContentType(string contentType) =>
        contentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase);

    public static bool IsSupportedContentType(string contentType) =>
        IsJpegContentType(contentType) || IsPngContentType(contentType) || IsBmpContentType(contentType);

    public static PdfImageXObject Decode(string contentType, byte[] bytes, Func<byte[], byte[]> recolorRgb, CancellationToken cancellationToken = default)
    {
        // PLAN Q01: conversion-wide cumulative charge before decoding allocates.
        OoxConversionBudget.Current?.ChargeImagesDecoded(1);
        if (IsJpegContentType(contentType))
        {
            JpegInfo info = JpegInfo.Read(bytes);
            return PdfImageXObject.Jpeg(info.Width, info.Height, bytes, info.ComponentCount, info.BitsPerComponent);
        }

        (int width, int height, byte[] rgb, byte[]? alpha) = DecodePixels(contentType, bytes, cancellationToken);
        return PdfImageXObject.RgbPng(width, height, recolorRgb(rgb), alpha);
    }

    public static (int Width, int Height, byte[] Rgb, byte[]? Alpha) DecodePixels(string contentType, byte[] bytes, CancellationToken cancellationToken = default)
    {
        if (IsPngContentType(contentType))
        {
            PngImage png = PngImage.Read(bytes, cancellationToken);
            return (png.Width, png.Height, png.Rgb, png.Alpha);
        }

        if (IsBmpContentType(contentType))
        {
            BmpImage bmp = BmpImage.Read(bytes, cancellationToken);
            return (bmp.Width, bmp.Height, bmp.Rgb, bmp.Alpha);
        }

        if (IsJpegContentType(contentType))
        {
            JpegImage jpeg = JpegImage.Read(bytes, cancellationToken);
            return (jpeg.Width, jpeg.Height, jpeg.Rgb, null);
        }

        throw new NotSupportedException("Unsupported image content type.");
    }
}
