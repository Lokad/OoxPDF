using System.IO.Compression;
using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class JpegPassthroughTests
{
    // Pillow 11.1, 16x16 uniform (220,40,20), quality=100, subsampling=0.
    // keep_rgb=true stores RGB samples with Adobe transform 0 and R/G/B IDs.
    // Default storage uses YCbCr; independent decoded pixels are (221,40,21).
    private const string RgbSamples = "/9j/7gAOQWRvYmUAZAAAAAAA/9sAQwABAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEB/8AAEQgAEAAQA1IRAEcRAEIRAP/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/aAAwDUgBHAEIAAD8A/rgr+T+v4n6KKKKKKKKK/9k=";
    private const string YCbCrSamples = "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAAQABADAREAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD836/y/P8ArgCgAoAKAP/Z";

    public static void JpegPixelDecodeKeepsRgbAndYCbCrAcrossMarkerPlacement()
    {
        byte[] rgb = Convert.FromBase64String(RgbSamples);
        var (withoutAdobe, adobe) = SplitAdobeMarker(rgb);
        int frame = withoutAdobe.AsSpan().IndexOf(new byte[] { 0xFF, 0xC0 });
        int frameEnd = frame + 2 + (withoutAdobe[frame + 2] << 8 | withoutAdobe[frame + 3]);
        byte[] afterFrame = [.. withoutAdobe[..frameEnd], .. adobe, .. withoutAdobe[frameEnd..]];
        byte[] numericIds = (byte[])rgb.Clone();
        int originalFrame = numericIds.AsSpan().IndexOf(new byte[] { 0xFF, 0xC0 });
        int scan = numericIds.AsSpan().IndexOf(new byte[] { 0xFF, 0xDA });
        for (int index = 0; index < 3; index++)
        {
            numericIds[originalFrame + 10 + index * 3] = (byte)(index + 1);
            numericIds[scan + 5 + index * 2] = (byte)(index + 1);
        }
        // Adobe transform 0 still owns RGB interpretation with numeric IDs.
        foreach (byte[] bytes in new[] { rgb, withoutAdobe, afterFrame, numericIds })
        {
            using var decoded = JpegImage.ReadOwned(bytes);
            AssertSamplePlane(decoded.Rgb, [220, 40, 20]);
        }
        using var ycbcr = JpegImage.ReadOwned(Convert.FromBase64String(YCbCrSamples));
        AssertSamplePlane(ycbcr.Rgb, [221, 40, 21]);
        using var gray = JpegImage.ReadOwned(TestFixtures.CreateGrayscaleBaselineJpeg());
        AssertSamplePlane(gray.Rgb, [128, 128, 128]);
    }

    public static void PptxJpegCropAndRecolorKeepSampleInterpretation()
    {
        foreach (var source in new[] {
            (Bytes: Convert.FromBase64String(RgbSamples), Pixel: new byte[] { 220, 40, 20 }, Gray: (byte)77),
            (Bytes: Convert.FromBase64String(YCbCrSamples), Pixel: new byte[] { 221, 40, 21 }, Gray: (byte)77),
            (Bytes: TestFixtures.CreateGrayscaleBaselineJpeg(), Pixel: new byte[] { 128, 128, 128 }, Gray: (byte)128) })
        foreach (bool crop in new[] { false, true })
        foreach (bool recolor in new[] { false, true })
        {
            var entries = PptxEntries();
            entries["ppt/media/probe.jpg"] = source.Bytes;
            string xml = Encoding.UTF8.GetString(entries["ppt/slides/slide1.xml"]);
            if (crop) xml = xml.Replace("<a:stretch>", "<a:srcRect l=\"25000\" r=\"25000\"/><a:stretch>", StringComparison.Ordinal);
            if (recolor) xml = xml.Replace("<a:blip r:embed=\"image\"/>",
                "<a:blip r:embed=\"image\"><a:grayscl/></a:blip>", StringComparison.Ordinal);
            entries["ppt/slides/slide1.xml"] = TestFixtures.Utf8(xml);
            using var input = TestFixtures.CreateZipPackage(entries);
            using var package = OoxPackage.Open(input, default);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default)
                .ToBlockingEnumerable().Single();
            TestAssert.Equal(0, diagnostics.Count);
            var image = page.Images.Single().Image;
            if (!crop && !recolor)
            {
                TestAssert.Equal("/DCTDecode", image.Filter);
                TestAssert.True(source.Bytes.AsSpan().SequenceEqual(image.Bytes), "Passthrough preserves original JPEG bytes.");
                continue;
            }
            TestAssert.Equal("/FlateDecode", image.Filter);
            using var compressed = new MemoryStream(image.Bytes, writable: false);
            using var inflate = new ZLibStream(compressed, CompressionMode.Decompress);
            using var plane = new MemoryStream();
            inflate.CopyTo(plane);
            TestAssert.Equal(image.Width * image.Height * 3L, plane.Length);
            AssertSamplePlane(plane.ToArray(), recolor ? [source.Gray, source.Gray, source.Gray] : source.Pixel);
        }
    }

    public static void JpegPixelDecodeRejectsUnsupportedAndConflictingAdobeInterpretation()
    {
        byte[] rgb = Convert.FromBase64String(RgbSamples);
        var (withoutAdobe, adobe) = SplitAdobeMarker(rgb);
        byte[] unsupported = (byte[])rgb.Clone();
        int signature = unsupported.AsSpan().IndexOf("Adobe"u8);
        unsupported[signature + 11] = 2; // A YCCK transform cannot describe three components.
        var pool = new TestMemoryBudget();
        using (var scope = OoxConversionBudget.BeginScope(null, pool))
            TestAssert.Throws<NotSupportedException>(() => JpegImage.ReadOwned(unsupported));
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);

        byte[] conflicting = (byte[])adobe.Clone();
        conflicting[^1] = 1;
        byte[] conflict = [.. rgb[..2], .. conflicting, .. rgb[2..]];
        byte[] truncated = [.. withoutAdobe[..2], 0xFF, 0xEE, 0, 7, .. "Adobe"u8, .. withoutAdobe[2..]];
        foreach (byte[] damaged in new[] { conflict, truncated })
        {
            TestAssert.Throws<InvalidDataException>(() => JpegInfo.Read(damaged));
            TestAssert.Throws<InvalidDataException>(() => JpegImage.ReadOwned(damaged));
        }
    }

    private static (byte[] Body, byte[] Marker) SplitAdobeMarker(byte[] bytes)
    {
        int start = bytes.AsSpan().IndexOf("Adobe"u8) - 4;
        TestAssert.True(start >= 2, "The RGB fixture has an Adobe segment.");
        int end = start + 2 + (bytes[start + 2] << 8 | bytes[start + 3]);
        return ([.. bytes[..start], .. bytes[end..]], bytes[start..end]);
    }

    private static void AssertSamplePlane(byte[] actual, byte[] pixel)
    {
        TestAssert.True(actual.Length > 0 && actual.Length % 3 == 0, "A decoded RGB plane has complete pixels.");
        for (int offset = 0; offset < actual.Length; offset += 3)
            TestAssert.True(actual.AsSpan(offset, 3).SequenceEqual(pixel), "Every uniform source pixel keeps its independently decoded color.");
    }
}
