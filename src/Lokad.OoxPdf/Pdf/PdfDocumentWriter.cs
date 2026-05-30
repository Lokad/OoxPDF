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
            .GroupBy(f => f.ResourceKey, StringComparer.Ordinal)
            .Select(PdfEmbeddedFont.Merge)
            .ToList();
        List<PdfImageXObject> images = pages
            .SelectMany(p => p.Images.Select(i => i.Image).Concat(p.ExtGStates.Where(s => s.SoftMask is not null).Select(s => s.SoftMask!.Image)))
            .DistinctBy(i => i.ResourceKey)
            .ToList();
        List<PdfAxialShading> shadings = pages
            .SelectMany(p => p.Shadings.Select(s => s.Shading))
            .DistinctBy(s => s.ResourceKey)
            .ToList();
        List<PdfTilingPattern> patterns = pages
            .SelectMany(p => p.Patterns.Select(s => s.Pattern))
            .DistinctBy(s => s.ResourceKey)
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

        int shadingObjectBase = nextImageObject;
        var shadingObjects = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < shadings.Count; i++)
        {
            shadingObjects[shadings[i].ResourceKey] = shadingObjectBase + i;
        }

        int softMaskObjectBase = shadingObjectBase + shadings.Count;
        List<PdfLuminositySoftMask> softMasks = pages
            .SelectMany(p => p.ExtGStates)
            .Where(s => s.SoftMask is not null)
            .Select(s => s.SoftMask!)
            .DistinctBy(s => s.ResourceKey)
            .ToList();
        var softMaskObjects = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < softMasks.Count; i++)
        {
            softMaskObjects[softMasks[i].ResourceKey] = softMaskObjectBase + i;
        }

        int patternObjectBase = softMaskObjectBase + softMasks.Count;
        var patternObjects = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < patterns.Count; i++)
        {
            patternObjects[patterns[i].ResourceKey] = patternObjectBase + i;
        }

        int objectCount = patternObjectBase + patterns.Count - 1;
        writer.WriteObject(1, "<< /Type /Catalog /Pages 2 0 R >>\n");
        writer.WriteObject(2, BuildPagesObject(pages));

        for (int i = 0; i < pages.Count; i++)
        {
            int pageObjectNumber = 3 + i * 2;
            int contentObjectNumber = pageObjectNumber + 1;
            PdfPage page = pages[i];

            writer.WriteObject(pageObjectNumber, FormattableString.Invariant(
                $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {FormatNumber(page.Width)} {FormatNumber(page.Height)}] /Contents {contentObjectNumber} 0 R /Resources {BuildResources(page, fontObjects, imageObjects, shadingObjects, softMaskObjects, patternObjects)} >>\n"));
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

        foreach (PdfAxialShading shading in shadings)
        {
            WriteAxialShadingObject(writer, shading, shadingObjects[shading.ResourceKey]);
        }

        foreach (PdfLuminositySoftMask softMask in softMasks)
        {
            WriteLuminositySoftMaskObject(writer, softMask, softMaskObjects[softMask.ResourceKey], imageObjects);
        }

        foreach (PdfTilingPattern pattern in patterns)
        {
            WriteTilingPatternObject(writer, pattern, patternObjects[pattern.ResourceKey]);
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

    private static string BuildResources(
        PdfPage page,
        IReadOnlyDictionary<string, FontObjectNumbers> fontObjects,
        IReadOnlyDictionary<string, ImageObjectNumbers> imageObjects,
        IReadOnlyDictionary<string, int> shadingObjects,
        IReadOnlyDictionary<string, int> softMaskObjects,
        IReadOnlyDictionary<string, int> patternObjects)
    {
        if (page.Fonts.Count == 0 && page.Images.Count == 0 && page.ExtGStates.Count == 0 && page.Shadings.Count == 0 && page.Patterns.Count == 0)
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
                builder.Append(CultureInfo.InvariantCulture, $" << /ca {FormatNumber(state.FillAlpha)} /CA {FormatNumber(state.StrokeAlpha)}");
                if (state.SoftMask is not null)
                {
                    builder.Append(CultureInfo.InvariantCulture, $" /SMask << /S /Luminosity /G {softMaskObjects[state.SoftMask.ResourceKey]} 0 R >>");
                }

                builder.Append(" >>");
            }

            builder.Append(" >>");
        }

        if (page.Shadings.Count != 0)
        {
            builder.Append(" /Shading <<");
            foreach (PdfShadingResource shading in page.Shadings)
            {
                builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(shading.ResourceName));
                builder.Append(CultureInfo.InvariantCulture, $" {shadingObjects[shading.Shading.ResourceKey]} 0 R");
            }

            builder.Append(" >>");
        }

        if (page.Patterns.Count != 0)
        {
            builder.Append(" /Pattern <<");
            foreach (PdfTilingPatternResource pattern in page.Patterns)
            {
                builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(pattern.ResourceName));
                builder.Append(CultureInfo.InvariantCulture, $" {patternObjects[pattern.Pattern.ResourceKey]} 0 R");
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
            $"/Type /XObject /Subtype /Image /Width {image.Width} /Height {image.Height} /ColorSpace {image.ColorSpace} /BitsPerComponent {image.BitsPerComponent} /Filter {image.Filter}{smask}"), image.Bytes);
        if (image.Alpha is not null && objects.SoftMask is not null)
        {
            writer.WriteStreamObject(objects.SoftMask.Value, FormattableString.Invariant(
                $"/Type /XObject /Subtype /Image /Width {image.Width} /Height {image.Height} /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode"), image.Alpha);
        }
    }

    private static void WriteAxialShadingObject(PdfObjectWriter writer, PdfAxialShading shading, int objectNumber)
    {
        writer.WriteObject(objectNumber, FormattableString.Invariant(
            $"<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [{FormatNumber(shading.X0)} {FormatNumber(shading.Y0)} {FormatNumber(shading.X1)} {FormatNumber(shading.Y1)}] /Function {BuildAxialShadingFunction(shading.Stops)} /Extend [true true] >>\n"));
    }

    private static void WriteLuminositySoftMaskObject(
        PdfObjectWriter writer,
        PdfLuminositySoftMask softMask,
        int objectNumber,
        IReadOnlyDictionary<string, ImageObjectNumbers> imageObjects)
    {
        ImageObjectNumbers image = imageObjects[softMask.Image.ResourceKey];
        string content = BuildLuminositySoftMaskContent(softMask);
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        writer.WriteObject(objectNumber, FormattableString.Invariant(
            $"<< /Type /XObject /Subtype /Form /BBox [{FormatNumber(softMask.X)} {FormatNumber(softMask.Y)} {FormatNumber(softMask.X + softMask.Width)} {FormatNumber(softMask.Y + softMask.Height)}] /Group << /S /Transparency /CS /DeviceRGB >> /Resources << /XObject << /ImMask {image.Image} 0 R >> >> /Length {contentBytes.Length} >>\nstream\n{content}endstream\n"));
    }

    private static void WriteTilingPatternObject(PdfObjectWriter writer, PdfTilingPattern pattern, int objectNumber)
    {
        byte[] contentBytes = Encoding.ASCII.GetBytes(pattern.Content);
        string matrix = pattern.Matrix is { } m
            ? FormattableString.Invariant($" /Matrix [{FormatNumber(m.A)} {FormatNumber(m.B)} {FormatNumber(m.C)} {FormatNumber(m.D)} {FormatNumber(m.E)} {FormatNumber(m.F)}]")
            : string.Empty;
        writer.WriteObject(objectNumber, FormattableString.Invariant(
            $"<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType {pattern.TilingType} /BBox [0 0 {FormatNumber(pattern.Width)} {FormatNumber(pattern.Height)}]{matrix} /XStep {FormatNumber(pattern.XStep)} /YStep {FormatNumber(pattern.YStep)} /Resources << >> /Length {contentBytes.Length} >>\nstream\n{pattern.Content}endstream\n"));
    }

    private static string BuildLuminositySoftMaskContent(PdfLuminositySoftMask softMask)
    {
        double visibleWidth = Math.Max(0.001d, 1d - softMask.CropLeft - softMask.CropRight);
        double visibleHeight = Math.Max(0.001d, 1d - softMask.CropTop - softMask.CropBottom);
        double scaledWidth = softMask.Width / visibleWidth;
        double scaledHeight = softMask.Height / visibleHeight;
        double imageX = softMask.X - softMask.CropLeft * scaledWidth;
        double imageY = softMask.Y - softMask.CropBottom * scaledHeight;
        var builder = new StringBuilder();
        builder.AppendLine("q");
        builder.Append(FormatNumber(softMask.X)).Append(' ').Append(FormatNumber(softMask.Y)).Append(' ');
        builder.Append(FormatNumber(softMask.Width)).Append(' ').Append(FormatNumber(softMask.Height)).AppendLine(" re W n");
        builder.Append(FormatNumber(scaledWidth)).Append(" 0 0 ").Append(FormatNumber(scaledHeight)).Append(' ');
        builder.Append(FormatNumber(imageX)).Append(' ').Append(FormatNumber(imageY)).AppendLine(" cm");
        builder.AppendLine("/ImMask Do");
        builder.AppendLine("Q");
        return builder.ToString();
    }

    private static string BuildAxialShadingFunction(IReadOnlyList<PdfShadingStop> stops)
    {
        if (stops.Count == 2)
        {
            return BuildExponentialInterpolationFunction(stops[0], stops[1]);
        }

        var builder = new StringBuilder();
        builder.Append("<< /FunctionType 3 /Domain [0 1] /Functions [");
        for (int i = 0; i < stops.Count - 1; i++)
        {
            builder.Append(' ').Append(BuildExponentialInterpolationFunction(stops[i], stops[i + 1]));
        }

        builder.Append(" ] /Bounds [");
        for (int i = 1; i < stops.Count - 1; i++)
        {
            if (i > 1)
            {
                builder.Append(' ');
            }

            builder.Append(FormatNumber(stops[i].Offset));
        }

        builder.Append("] /Encode [");
        for (int i = 0; i < stops.Count - 1; i++)
        {
            if (i > 0)
            {
                builder.Append(' ');
            }

            builder.Append("0 1");
        }

        builder.Append("] >>");
        return builder.ToString();
    }

    private static string BuildExponentialInterpolationFunction(PdfShadingStop start, PdfShadingStop end)
    {
        return FormattableString.Invariant(
            $"<< /FunctionType 2 /Domain [0 1] /C0 [{FormatColor(start.Red)} {FormatColor(start.Green)} {FormatColor(start.Blue)}] /C1 [{FormatColor(end.Red)} {FormatColor(end.Green)} {FormatColor(end.Blue)}] /N 1 >>");
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

    private static string FormatColor(byte value)
    {
        return (value / 255d).ToString("0.###", CultureInfo.InvariantCulture);
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
