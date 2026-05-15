using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfDocumentWriter
{
    public static void WriteBlank(Stream stream, IReadOnlyList<PdfPage> pages)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (pages.Count == 0)
        {
            throw new ArgumentException("A PDF document must contain at least one page.", nameof(pages));
        }

        var writer = new PdfObjectWriter(stream);
        writer.WriteHeader();

        List<PdfEmbeddedFont> fonts = pages
            .SelectMany(p => p.Fonts.Select(f => f.Font))
            .DistinctBy(f => f.ResourceKey)
            .ToList();
        List<PdfImageXObject> images = pages
            .SelectMany(p => p.Images.Select(i => i.Image))
            .DistinctBy(i => i.ResourceKey)
            .ToList();

        int fontObjectBase = 3 + pages.Count * 2;
        var fontObjects = new Dictionary<string, FontObjectNumbers>(StringComparer.Ordinal);
        for (int i = 0; i < fonts.Count; i++)
        {
            int baseObject = fontObjectBase + i * 5;
            fontObjects[fonts[i].ResourceKey] = new FontObjectNumbers(
                Type0: baseObject,
                CidFont: baseObject + 1,
                Descriptor: baseObject + 2,
                FontFile: baseObject + 3,
                ToUnicode: baseObject + 4);
        }

        int imageObjectBase = fontObjectBase + fonts.Count * 5;
        var imageObjects = new Dictionary<string, ImageObjectNumbers>(StringComparer.Ordinal);
        int nextImageObject = imageObjectBase;
        foreach (PdfImageXObject image in images)
        {
            int imageObject = nextImageObject++;
            int? softMaskObject = image.Alpha is null ? null : nextImageObject++;
            imageObjects[image.ResourceKey] = new ImageObjectNumbers(imageObject, softMaskObject);
        }

        int objectCount = nextImageObject - 1;
        writer.WriteObject(1, "<< /Type /Catalog /Pages 2 0 R >>\n");
        writer.WriteObject(2, BuildPagesObject(pages));

        for (int i = 0; i < pages.Count; i++)
        {
            int pageObjectNumber = 3 + i * 2;
            int contentObjectNumber = pageObjectNumber + 1;
            PdfPage page = pages[i];

            writer.WriteObject(pageObjectNumber, FormattableString.Invariant(
                $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {FormatNumber(page.Width)} {FormatNumber(page.Height)}] /Contents {contentObjectNumber} 0 R /Resources {BuildResources(page, fontObjects, imageObjects)} >>\n"));
            byte[] contentBytes = Encoding.ASCII.GetBytes(page.Content);
            writer.WriteObject(contentObjectNumber, FormattableString.Invariant(
                $"<< /Length {contentBytes.Length} >>\nstream\n{page.Content}endstream\n"));
        }

        foreach (PdfEmbeddedFont font in fonts)
        {
            WriteFontObjects(writer, font, fontObjects[font.ResourceKey]);
        }

        foreach (PdfImageXObject image in images)
        {
            WriteImageObjects(writer, image, imageObjects[image.ResourceKey]);
        }

        long xrefOffset = writer.Position;
        writer.WriteAscii(FormattableString.Invariant($"xref\n0 {objectCount + 1}\n"));
        writer.WriteAscii("0000000000 65535 f \n");
        foreach (long offset in writer.Offsets)
        {
            writer.WriteAscii(FormattableString.Invariant($"{offset:0000000000} 00000 n \n"));
        }

        writer.WriteAscii(FormattableString.Invariant(
            $"trailer\n<< /Size {objectCount + 1} /Root 1 0 R >>\nstartxref\n{xrefOffset}\n%%EOF\n"));
    }

    private static string BuildPagesObject(IReadOnlyList<PdfPage> pages)
    {
        var builder = new StringBuilder();
        builder.Append("<< /Type /Pages /Count ");
        builder.Append(CultureInfo.InvariantCulture, $"{pages.Count}");
        builder.Append(" /Kids [");
        for (int i = 0; i < pages.Count; i++)
        {
            if (i > 0)
            {
                builder.Append(' ');
            }

            builder.Append(CultureInfo.InvariantCulture, $"{3 + i * 2} 0 R");
        }

        builder.Append("] >>\n");
        return builder.ToString();
    }

    private static string BuildResources(PdfPage page, IReadOnlyDictionary<string, FontObjectNumbers> fontObjects, IReadOnlyDictionary<string, ImageObjectNumbers> imageObjects)
    {
        if (page.Fonts.Count == 0 && page.Images.Count == 0 && page.ExtGStates.Count == 0)
        {
            return "<< >>";
        }

        var builder = new StringBuilder("<<");
        if (page.Fonts.Count != 0)
        {
            builder.Append(" /Font <<");
            foreach (PdfFontResource font in page.Fonts)
            {
                FontObjectNumbers objects = fontObjects[font.Font.ResourceKey];
                builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(font.ResourceName)).Append(' ');
                builder.Append(CultureInfo.InvariantCulture, $"{objects.Type0} 0 R");
            }

            builder.Append(" >>");
        }

        if (page.Images.Count != 0)
        {
            builder.Append(" /XObject <<");
            foreach (PdfImageResource image in page.Images)
            {
                ImageObjectNumbers objects = imageObjects[image.Image.ResourceKey];
                builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(image.ResourceName)).Append(' ');
                builder.Append(CultureInfo.InvariantCulture, $"{objects.Image} 0 R");
            }

            builder.Append(" >>");
        }

        if (page.ExtGStates.Count != 0)
        {
            builder.Append(" /ExtGState <<");
            foreach (PdfExtGStateResource state in page.ExtGStates)
            {
                builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(state.ResourceName));
                builder.Append(CultureInfo.InvariantCulture, $" << /ca {FormatNumber(state.FillAlpha)} /CA {FormatNumber(state.StrokeAlpha)} >>");
            }

            builder.Append(" >>");
        }

        builder.Append(" >>");
        return builder.ToString();
    }

    private static void WriteFontObjects(PdfObjectWriter writer, PdfEmbeddedFont font, FontObjectNumbers objects)
    {
        string baseFont = PdfEmbeddedFont.SanitizeName(font.BaseFontName);
        writer.WriteObject(objects.Type0, FormattableString.Invariant(
            $"<< /Type /Font /Subtype /Type0 /BaseFont /{baseFont} /Encoding /Identity-H /DescendantFonts [{objects.CidFont} 0 R] /ToUnicode {objects.ToUnicode} 0 R >>\n"));

        writer.WriteObject(objects.CidFont, FormattableString.Invariant(
            $"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{baseFont} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor {objects.Descriptor} 0 R /CIDToGIDMap /Identity /W {font.BuildWidthArray()} >>\n"));

        OpenTypeFontMetrics metrics = OpenTypeFontMetrics.From(font.Font);
        writer.WriteObject(objects.Descriptor, FormattableString.Invariant(
            $"<< /Type /FontDescriptor /FontName /{baseFont} /Flags {metrics.Flags} /FontBBox [{metrics.XMin} {metrics.YMin} {metrics.XMax} {metrics.YMax}] /ItalicAngle {FormatNumber(metrics.ItalicAngle)} /Ascent {metrics.Ascent} /Descent {metrics.Descent} /CapHeight {metrics.CapHeight} /StemV 80 /FontFile2 {objects.FontFile} 0 R >>\n"));

        writer.WriteStreamObject(objects.FontFile, "/Filter /FlateDecode", Compress(font.Font.Bytes.Span));
        writer.WriteStreamObject(objects.ToUnicode, string.Empty, Encoding.ASCII.GetBytes(font.BuildToUnicodeCMap()));
    }

    private static void WriteImageObjects(PdfObjectWriter writer, PdfImageXObject image, ImageObjectNumbers objects)
    {
        string smask = objects.SoftMask is null ? string.Empty : FormattableString.Invariant($" /SMask {objects.SoftMask.Value} 0 R");
        writer.WriteStreamObject(objects.Image, FormattableString.Invariant(
            $"/Type /XObject /Subtype /Image /Width {image.Width} /Height {image.Height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter {image.Filter}{smask}"), image.Bytes);
        if (image.Alpha is not null && objects.SoftMask is not null)
        {
            writer.WriteStreamObject(objects.SoftMask.Value, FormattableString.Invariant(
                $"/Type /XObject /Subtype /Image /Width {image.Width} /Height {image.Height} /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode"), image.Alpha);
        }
    }

    private static byte[] Compress(ReadOnlySpan<byte> bytes)
    {
        using var output = new MemoryStream();
        using (var zlib = new System.IO.Compression.ZLibStream(output, System.IO.Compression.CompressionLevel.SmallestSize, leaveOpen: true))
        {
            zlib.Write(bytes);
        }

        return output.ToArray();
    }

    private static string FormatNumber(double value)
    {
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }

    private readonly record struct FontObjectNumbers(int Type0, int CidFont, int Descriptor, int FontFile, int ToUnicode);

    private readonly record struct ImageObjectNumbers(int Image, int? SoftMask);

    private readonly record struct OpenTypeFontMetrics(
        int XMin,
        int YMin,
        int XMax,
        int YMax,
        int Ascent,
        int Descent,
        int CapHeight,
        double ItalicAngle,
        int Flags)
    {
        public static OpenTypeFontMetrics From(Fonts.OpenTypeFont font)
        {
            double scale = 1000d / font.UnitsPerEm;
            int flags = font.Post.IsFixedPitch ? 1 : 32;
            if (Math.Abs(font.Post.ItalicAngle) > 0.001)
            {
                flags |= 64;
            }

            return new OpenTypeFontMetrics(
                Scale(font.Bounds.XMin, scale),
                Scale(font.Bounds.YMin, scale),
                Scale(font.Bounds.XMax, scale),
                Scale(font.Bounds.YMax, scale),
                Scale(font.Os2.WindowsAscender, scale),
                -Scale(font.Os2.WindowsDescender, scale),
                Scale(font.Os2.TypographicAscender, scale),
                font.Post.ItalicAngle,
                flags);
        }

        private static int Scale(double value, double scale)
        {
            return (int)Math.Round(value * scale);
        }
    }
}
