using System.Security.Cryptography;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class PdfIdentityTests
{
    public static void ImageIdentityUsesFullDigests()
    {
        PdfImageXObject first = PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8);
        PdfImageXObject second = PdfImageXObject.Jpeg(2, 1, [1, 2, 4], 3, 8);
        TestAssert.True(first.ResourceKey != second.ResourceKey, "One payload byte must change the image identity.");
        string expectedDigest = Convert.ToHexString(SHA256.HashData([1, 2, 3]));
        TestAssert.Equal(64, expectedDigest.Length);
        TestAssert.True(first.ResourceKey.Contains(expectedDigest, StringComparison.Ordinal), "Identity must carry the full content digest, not a truncated prefix.");
        TestAssert.True(!second.ResourceKey.Contains(expectedDigest, StringComparison.Ordinal), "Different payloads must not share a digest.");
    }

    public static void ImageContentEqualityIsExact()
    {
        PdfImageXObject image = PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8);
        TestAssert.True(image.HasIdenticalContent(PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8)), "Separate equal instances must compare identical.");
        TestAssert.True(!image.HasIdenticalContent(null), "Null content must not compare identical.");
        TestAssert.True(!image.HasIdenticalContent(PdfImageXObject.Jpeg(3, 1, [1, 2, 3], 3, 8)), "Width must participate in identity.");
        TestAssert.True(!image.HasIdenticalContent(PdfImageXObject.Jpeg(2, 2, [1, 2, 3], 3, 8)), "Height must participate in identity.");
        TestAssert.True(!image.HasIdenticalContent(PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 1, 8)), "Color space must participate in identity.");
        TestAssert.True(!image.HasIdenticalContent(PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 4)), "Bits per component must participate in identity.");
        TestAssert.True(!image.HasIdenticalContent(PdfImageXObject.Jpeg(2, 1, [1, 2, 4], 3, 8)), "One payload byte must break identity.");
        TestAssert.True(!image.HasIdenticalContent(PdfImageXObject.RgbPng(2, 1, [1, 2, 3, 4, 5, 6], null)), "Filter and payload must participate in identity.");

        PdfImageXObject withAlpha = PdfImageXObject.RgbPng(2, 1, [1, 2, 3, 4, 5, 6], [7, 8]);
        TestAssert.True(!withAlpha.HasIdenticalContent(PdfImageXObject.RgbPng(2, 1, [1, 2, 3, 4, 5, 6], null)), "Alpha presence must participate in identity.");
        TestAssert.True(!withAlpha.HasIdenticalContent(PdfImageXObject.RgbPng(2, 1, [1, 2, 3, 4, 5, 6], [7])), "Alpha length must participate in identity.");
        TestAssert.True(!withAlpha.HasIdenticalContent(PdfImageXObject.RgbPng(2, 1, [1, 2, 3, 4, 5, 6], [7, 9])), "Alpha content must participate in identity.");
        TestAssert.True(withAlpha.HasIdenticalContent(PdfImageXObject.RgbPng(2, 1, [1, 2, 3, 4, 5, 6], [7, 8])), "Equal alpha channels must compare identical.");
    }

    public static void DeduplicateImagesKeepsDistinctContent()
    {
        PdfImageXObject first = PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8);
        PdfImageXObject second = PdfImageXObject.Jpeg(2, 1, [1, 2, 4], 3, 8);
        PdfImageXObject duplicate = PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8);
        List<PdfImageXObject> deduped = PdfDocumentWriter.DeduplicateImages([first, second, duplicate], CancellationToken.None);
        TestAssert.Equal(2, deduped.Count);
        TestAssert.True(ReferenceEquals(first, deduped[0]), "Deduplication must keep the first identical instance.");
        TestAssert.True(ReferenceEquals(second, deduped[1]), "Distinct content must survive deduplication in order.");
    }

    public static void DeduplicateImagesRejectsStaleContentIdentity()
    {
        PdfImageXObject first = PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8);
        PdfImageXObject second = PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8);
        string identity = first.ResourceKey;
        first.Bytes[2] = 4; // Model-only invariant check; no live conversion mutation is claimed.
        TestAssert.Equal(identity, second.ResourceKey);
        TestAssert.Throws<InvalidDataException>(() => PdfDocumentWriter.DeduplicateImages([first, second], CancellationToken.None));
        List<PdfImageXObject> deduped = PdfDocumentWriter.DeduplicateImages([second, PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8)], CancellationToken.None);
        TestAssert.Equal(1, deduped.Count);
    }

    public static void DeduplicateImagesObservesCancellation()
    {
        PdfImageXObject image = PdfImageXObject.Jpeg(2, 1, [1, 2, 3], 3, 8);
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => PdfDocumentWriter.DeduplicateImages([image], cancelled.Token));
    }
}
