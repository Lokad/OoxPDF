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
        if (IsJpegContentType(contentType))
        {
            OoxConversionBudget.Current?.ChargeImagesDecoded(1);
            JpegInfo info = JpegInfo.Read(bytes);
            return PdfImageXObject.Jpeg(info.Width, info.Height, bytes, info.ComponentCount, info.BitsPerComponent);
        }

        // R02: the pixel reservation spans decoding, the recolor transform, and PDF
        // compression as one scoped operation instead of releasing at decoder return.
        using DecodedPixels decoded = DecodePixelsOwned(contentType, bytes, cancellationToken);
        return PdfImageXObject.RgbPng(decoded.Width, decoded.Height, recolorRgb(decoded.Rgb), decoded.Alpha);
    }

    public static (int Width, int Height, byte[] Rgb, byte[]? Alpha) DecodePixels(string contentType, byte[] bytes, CancellationToken cancellationToken = default)
    {
        using DecodedPixels owned = DecodePixelsOwned(contentType, bytes, cancellationToken);
        return (owned.Width, owned.Height, owned.Rgb, owned.Alpha);
    }

    // R02/R03: single shared charge for every pixel decode (content, crop, recolor
    // variants). Cache hits return before this boundary and do not charge. The
    // returned ownership keeps the working-set reservation alive across the
    // caller transform/compress steps that consume the pixels.
    public static DecodedPixels DecodePixelsOwned(string contentType, byte[] bytes, CancellationToken cancellationToken = default)
    {
        if (IsPngContentType(contentType))
        {
            OoxConversionBudget.Current?.ChargeImagesDecoded(1);
            return PngImage.ReadOwned(bytes, cancellationToken);
        }

        if (IsBmpContentType(contentType))
        {
            OoxConversionBudget.Current?.ChargeImagesDecoded(1);
            return BmpImage.ReadOwned(bytes, cancellationToken);
        }

        if (IsJpegContentType(contentType))
        {
            OoxConversionBudget.Current?.ChargeImagesDecoded(1);
            return JpegImage.ReadOwned(bytes, cancellationToken);
        }

        throw new NotSupportedException("Unsupported image content type.");
    }
}
