using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfDocumentWriter
{
    // R18: image deduplication groups by the full content identity and verifies byte
    // equality on every key match. A same-key content mismatch (only reachable through
    // a full-digest collision) fails loudly instead of silently emitting one image
    // under two names. The keySelector seam lets tests force collisions that SHA-256
    // makes unconstructible in production.
    internal static List<PdfImageXObject> DeduplicateImages(IEnumerable<PdfImageXObject> images, Func<PdfImageXObject, string> keySelector, CancellationToken cancellationToken)
    {
        var byKey = new Dictionary<string, PdfImageXObject>(StringComparer.Ordinal);
        var deduped = new List<PdfImageXObject>();
        foreach (PdfImageXObject image in images)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string key = keySelector(image);
            if (!byKey.TryGetValue(key, out PdfImageXObject? existing))
            {
                byKey[key] = image;
                deduped.Add(image);
                continue;
            }
            if (!existing.HasIdenticalContent(image))
            {
                throw new InvalidDataException($"PDF image identity collision for '{key}': distinct pixel content shares one resource identity.");
            }
        }
        return deduped;
    }
    // R06: returns the exact serialized byte count so the conversion scope (still open
    // at the call site) can charge output bytes after writing. Pages and encoded
    // content bytes charge up front, bounding serialization before it allocates.
    public static long WriteBlank(Stream stream, IReadOnlyList<PdfPage> pages, CancellationToken cancellationToken, DateTimeOffset? creationDate = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (pages.Count == 0)
        {
            throw new ArgumentException("A PDF document must contain at least one page.", nameof(pages));
        }

        OoxConversionBudget.Current?.ChargePdfPages(pages.Count);
        long contentByteTotal = 0;
        foreach (PdfPage page in pages)
        {
            contentByteTotal = checked(contentByteTotal + page.Content.Length);
        }

        OoxConversionBudget.Current?.ChargePdfContentBytes(contentByteTotal);

        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            PdfPage page = pages[pageIndex];
            if (!double.IsFinite(page.Width) || !double.IsFinite(page.Height) || page.Width <= 0 || page.Height <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(pages), "PDF pages must have finite positive dimensions.");
            }

            PdfContentValidator.ValidatePage(page, pageIndex, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();
        var writer = new PdfObjectWriter(stream, cancellationToken);
        writer.WriteHeader();

        PdfDocumentPlan plan = BuildDocumentPlan(pages, cancellationToken);
        PdfDocumentNumbers numbers = AssignDocumentNumbers(plan, pages, creationDate, cancellationToken);

        writer.WriteObject(1, "<< /Type /Catalog /Pages 2 0 R >>\n");
        writer.WriteObject(2, BuildPagesObject());

        for (int i = 0; i < pages.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int pageObjectNumber = 3 + i * 2;
            int contentObjectNumber = pageObjectNumber + 1;
            PdfPage page = pages[i];

            writer.WriteObject(pageObjectNumber, FormattableString.Invariant(
                $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {FormatNumber(page.Width)} {FormatNumber(page.Height)}] /Contents {contentObjectNumber} 0 R /Resources {BuildResources(page)}{BuildPageAnnotations(numbers.AnnotationObjectsByPage[i])} >>\n"));
            byte[] contentBytes = Encoding.ASCII.GetBytes(page.Content);
            writer.WriteContentStreamObject(contentObjectNumber, contentBytes);
        }

        foreach (PdfEmbeddedFont font in plan.Fonts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            WriteFontObjects(writer, font, numbers.FontObjects[font.ResourceKey], cancellationToken);
        }

        foreach (PdfImageXObject image in plan.Images)
        {
            cancellationToken.ThrowIfCancellationRequested();
            WriteImageObjects(writer, image, numbers.ImageObjects[image.ResourceKey]);
        }

        foreach (PdfAxialShading shading in plan.Shadings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            WriteAxialShadingObject(writer, shading, numbers.ShadingObjects[shading.ResourceKey]);
        }

        foreach (PdfLuminositySoftMask softMask in plan.SoftMasks)
        {
            cancellationToken.ThrowIfCancellationRequested();
            WriteLuminositySoftMaskObject(writer, softMask, numbers.SoftMaskObjects[softMask.ResourceKey], numbers.ImageObjects);
        }

        foreach (PdfTilingPattern pattern in plan.Patterns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            WriteTilingPatternObject(writer, pattern, numbers.PatternObjects[pattern.ResourceKey], numbers.ImageObjects);
        }

        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int[] annotationObjects = numbers.AnnotationObjectsByPage[pageIndex];
            IReadOnlyList<PdfLinkAnnotation> annotations = pages[pageIndex].Annotations;
            for (int annotationIndex = 0; annotationIndex < annotations.Count; annotationIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                WriteLinkAnnotationObject(writer, annotations[annotationIndex], annotationObjects[annotationIndex], pages.Count);
            }
        }

        if (numbers.InfoObjectNumber is not null && creationDate is not null)
        {
            writer.WriteObject(numbers.InfoObjectNumber.Value, BuildInfoObject(creationDate.Value));
        }

        if (writer.Offsets.Count != numbers.ObjectCount)
        {
            throw new InvalidDataException($"PDF object numbering is inconsistent: {writer.Offsets.Count} written objects but {numbers.ObjectCount} counted objects.");
        }

        long xrefOffset = writer.Position;
        writer.WriteAscii(FormattableString.Invariant($"xref\n0 {numbers.ObjectCount + 1}\n"));
        writer.WriteAscii("0000000000 65535 f \n");
        foreach (long offset in writer.Offsets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            writer.WriteAscii(FormattableString.Invariant($"{offset:0000000000} 00000 n \n"));
        }

        writer.WriteAscii(FormattableString.Invariant(
            $"trailer\n<< /Size {numbers.ObjectCount + 1} /Root 1 0 R{(numbers.InfoObjectNumber is null ? string.Empty : FormattableString.Invariant($" /Info {numbers.InfoObjectNumber.Value} 0 R"))} >>\nstartxref\n{xrefOffset}\n%%EOF\n"));

        return writer.Position;

        string BuildPagesObject()
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

        string BuildPageAnnotations(IReadOnlyList<int> annotationObjects)
        {
            if (annotationObjects.Count == 0)
            {
                return string.Empty;
            }

            var builder = new StringBuilder(" /Annots [");
            for (int i = 0; i < annotationObjects.Count; i++)
            {
                if (i > 0)
                {
                    builder.Append(' ');
                }

                builder.Append(CultureInfo.InvariantCulture, $"{annotationObjects[i]} 0 R");
            }

            builder.Append(']');
            return builder.ToString();
        }

        string BuildResources(PdfPage page)
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
                    FontObjectNumbers objects = numbers.FontObjects[font.Font.ResourceKey];
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
                    ImageObjectNumbers objects = numbers.ImageObjects[image.Image.ResourceKey];
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
                        builder.Append(CultureInfo.InvariantCulture, $" /SMask << /S /Luminosity /G {numbers.SoftMaskObjects[state.SoftMask.ResourceKey]} 0 R >>");
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
                    builder.Append(CultureInfo.InvariantCulture, $" {numbers.ShadingObjects[shading.Shading.ResourceKey]} 0 R");
                }

                builder.Append(" >>");
            }

            if (page.Patterns.Count != 0)
            {
                builder.Append(" /Pattern <<");
                foreach (PdfTilingPatternResource pattern in page.Patterns)
                {
                    builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(pattern.ResourceName));
                    builder.Append(CultureInfo.InvariantCulture, $" {numbers.PatternObjects[pattern.Pattern.ResourceKey]} 0 R");
                }

                builder.Append(" >>");
            }

            builder.Append(" >>");
            return builder.ToString();
        }
    }

    private static void WriteFontObjects(PdfObjectWriter writer, PdfEmbeddedFont font, FontObjectNumbers objects, CancellationToken cancellationToken)
    {
        string baseFont = PdfEmbeddedFont.SanitizeName(font.BaseFontName);
        writer.WriteObject(objects.Type0, FormattableString.Invariant(
            $"<< /Type /Font /Subtype /Type0 /BaseFont /{baseFont} /Encoding /Identity-H /DescendantFonts [{objects.CidFont} 0 R] /ToUnicode {objects.ToUnicode} 0 R >>\n"));

        writer.WriteObject(objects.CidFont, FormattableString.Invariant(
            $"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{baseFont} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor {objects.Descriptor} 0 R /CIDToGIDMap /Identity /W {font.BuildWidthArray(cancellationToken)} >>\n"));

        OpenTypeFontMetrics metrics = OpenTypeFontMetrics.From(font.Font);
        writer.WriteObject(objects.Descriptor, FormattableString.Invariant(
            $"<< /Type /FontDescriptor /FontName /{baseFont} /Flags {metrics.Flags} /FontBBox [{metrics.XMin} {metrics.YMin} {metrics.XMax} {metrics.YMax}] /ItalicAngle {FormatNumber(metrics.ItalicAngle)} /Ascent {metrics.Ascent} /Descent {metrics.Descent} /CapHeight {metrics.CapHeight} /StemV 80 /FontFile2 {objects.FontFile} 0 R >>\n"));

        ReadOnlyMemory<byte> fontProgram = font.FontProgramBytes;
        byte[] compressedFontProgram = Compress(fontProgram.Span, cancellationToken);
        byte[] toUnicode = Encoding.ASCII.GetBytes(font.BuildToUnicodeCMap(cancellationToken));
        // R06: retained font resources accumulate across embedded fonts; charge the
        // program plus ToUnicode bytes as they serialize (operation counts alone do
        // not bound retained resource bytes).
        OoxConversionBudget.Current?.ChargePdfFontBytes(checked((long)fontProgram.Length + toUnicode.Length));
        writer.WriteStreamObject(objects.FontFile, FormattableString.Invariant($"/Filter /FlateDecode /Length1 {fontProgram.Length}"), compressedFontProgram);
        writer.WriteStreamObject(objects.ToUnicode, string.Empty, toUnicode);
    }

    private static string BuildInfoObject(DateTimeOffset creationDate)
    {
        string producer = "Lokad.OoxPdf " + (typeof(PdfDocumentWriter).Assembly.GetName().Version?.ToString() ?? "0");
        string date = FormatPdfDate(creationDate);
        return FormattableString.Invariant($"<< /Producer ({producer}) /CreationDate ({date}) /ModDate ({date}) >>\n");
    }

    private static void WriteImageObjects(PdfObjectWriter writer, PdfImageXObject image, ImageObjectNumbers objects)
    {
        // R06: retained image resources accumulate across embedded images (soft masks
        // ride the same path); charge encoded bytes as they serialize.
        OoxConversionBudget.Current?.ChargePdfImageBytes(checked((long)image.Bytes.Length + (image.Alpha?.Length ?? 0)));
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

        string BuildAxialShadingFunction(IReadOnlyList<PdfShadingStop> stops)
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
    }

    private static void WriteLuminositySoftMaskObject(
        PdfObjectWriter writer,
        PdfLuminositySoftMask softMask,
        int objectNumber,
        IReadOnlyDictionary<string, ImageObjectNumbers> imageObjects)
    {
        ImageObjectNumbers image = imageObjects[softMask.Image.ResourceKey];
        string content = BuildLuminositySoftMaskContent();
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        writer.WriteObject(objectNumber, FormattableString.Invariant(
            $"<< /Type /XObject /Subtype /Form /BBox [{FormatNumber(softMask.X)} {FormatNumber(softMask.Y)} {FormatNumber(softMask.X + softMask.Width)} {FormatNumber(softMask.Y + softMask.Height)}] /Group << /S /Transparency /CS /DeviceRGB >> /Resources << /XObject << /ImMask {image.Image} 0 R >> >> /Length {contentBytes.Length} >>\nstream\n{content}endstream\n"));

        string BuildLuminositySoftMaskContent()
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
    }

    private static void WriteTilingPatternObject(
        PdfObjectWriter writer,
        PdfTilingPattern pattern,
        int objectNumber,
        IReadOnlyDictionary<string, ImageObjectNumbers> imageObjects)
    {
        byte[] contentBytes = Encoding.ASCII.GetBytes(pattern.Content);
        string matrix = pattern.Matrix is { } m
            ? FormattableString.Invariant($" /Matrix [{FormatNumber(m.A)} {FormatNumber(m.B)} {FormatNumber(m.C)} {FormatNumber(m.D)} {FormatNumber(m.E)} {FormatNumber(m.F)}]")
            : string.Empty;
        string resources = pattern.Images.Count == 0
            ? "<< >>"
            : BuildPatternResources();
        writer.WriteObject(objectNumber, FormattableString.Invariant(
            $"<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType {pattern.TilingType} /BBox [0 0 {FormatNumber(pattern.Width)} {FormatNumber(pattern.Height)}]{matrix} /XStep {FormatNumber(pattern.XStep)} /YStep {FormatNumber(pattern.YStep)} /Resources {resources} /Length {contentBytes.Length} >>\nstream\n{pattern.Content}endstream\n"));

        string BuildPatternResources()
        {
            var builder = new StringBuilder("<< /XObject <<");
            foreach (PdfImageResource image in pattern.Images)
            {
                ImageObjectNumbers objects = imageObjects[image.Image.ResourceKey];
                builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(image.ResourceName)).Append(' ');
                builder.Append(CultureInfo.InvariantCulture, $"{objects.Image} 0 R");
            }

            builder.Append(" >> >>");
            return builder.ToString();
        }
    }

    private static void WriteLinkAnnotationObject(
        PdfObjectWriter writer,
        PdfLinkAnnotation annotation,
        int objectNumber,
        int pageCount)
    {
        double x1 = annotation.X;
        double y1 = annotation.Y;
        double x2 = annotation.X + annotation.Width;
        double y2 = annotation.Y + annotation.Height;
        writer.WriteObject(objectNumber, FormattableString.Invariant(
            $"<< /Type /Annot /Subtype /Link /Rect [{FormatNumber(x1)} {FormatNumber(y1)} {FormatNumber(x2)} {FormatNumber(y2)}] /Border [0 0 0]{BuildLinkTarget()} >>\n"));

        string BuildLinkTarget()
        {
            if (!string.IsNullOrEmpty(annotation.Uri))
            {
                return $" /A << /S /URI /URI ({EscapePdfUriString(annotation.Uri)}) >>";
            }

            if (annotation.Destination is { } destination)
            {
                if (destination.PageIndex < 0 || destination.PageIndex >= pageCount)
                {
                    throw new InvalidOperationException(
                        FormattableString.Invariant($"PDF link destination page index {destination.PageIndex} is outside the document page range 0..{pageCount - 1}."));
                }

                int pageObjectNumber = 3 + destination.PageIndex * 2;
                string left = destination.Left is { } x ? FormatNumber(x) : "null";
                string top = destination.Top is { } y ? FormatNumber(y) : "null";
                string zoom = destination.Zoom is { } z ? FormatNumber(z) : "null";
                return FormattableString.Invariant($" /Dest [{pageObjectNumber} 0 R /XYZ {left} {top} {zoom}]");
            }

            throw new InvalidOperationException("PDF link annotations must have either a URI target or an internal destination target.");

        }
    }

    private static string BuildExponentialInterpolationFunction(PdfShadingStop start, PdfShadingStop end)
    {
        return FormattableString.Invariant(
            $"<< /FunctionType 2 /Domain [0 1] /C0 [{FormatColor(start.Red)} {FormatColor(start.Green)} {FormatColor(start.Blue)}] /C1 [{FormatColor(end.Red)} {FormatColor(end.Green)} {FormatColor(end.Blue)}] /N 1 >>");
    }

    private static byte[] Compress(ReadOnlySpan<byte> bytes, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using var output = new MemoryStream();
        using (var zlib = new System.IO.Compression.ZLibStream(output, System.IO.Compression.CompressionLevel.SmallestSize, leaveOpen: true))
        {
            cancellationToken.ThrowIfCancellationRequested();
            zlib.Write(bytes);
        }

        cancellationToken.ThrowIfCancellationRequested();
        return output.ToArray();
    }

    internal static string FormatNumber(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            throw new ArgumentOutOfRangeException(nameof(value), "PDF numbers must be finite.");
        }

        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }

    internal static string EscapePdfUriString(string value)
    {
        var builder = new StringBuilder(value.Length);
        Span<byte> utf8 = stackalloc byte[4];
        foreach (System.Text.Rune rune in value.EnumerateRunes())
        {
            if (rune.Value == 92 || rune.Value == 40 || rune.Value == 41)
            {
                builder.Append((char)92).Append((char)rune.Value);
            }
            else if (rune.Value <= 32 || rune.Value > 126)
            {
                int encodedLength = rune.EncodeToUtf8(utf8);
                for (int i = 0; i < encodedLength; i++)
                {
                    builder.Append((char)37);
                    builder.Append(utf8[i].ToString("X2", CultureInfo.InvariantCulture));
                }
            }
            else
            {
                builder.Append((char)rune.Value);
            }
        }

        return builder.ToString();
    }

    internal static string FormatPdfDate(DateTimeOffset value)
    {
        string sign = value.Offset < TimeSpan.Zero ? "-" : "+";
        TimeSpan absolute = value.Offset.Duration();
        return string.Format(CultureInfo.InvariantCulture, "D:{0:yyyyMMddHHmmss}{1}{2:00}\u0027{3:00}\u0027", value, sign, (int)absolute.TotalHours, absolute.Minutes);
    }

    internal static string FormatColor(byte value)
    {
        return (value / 255d).ToString("0.###", CultureInfo.InvariantCulture);
    }


    private sealed record PdfDocumentPlan(
        IReadOnlyList<PdfEmbeddedFont> Fonts,
        IReadOnlyList<PdfImageXObject> Images,
        IReadOnlyList<PdfAxialShading> Shadings,
        IReadOnlyList<PdfLuminositySoftMask> SoftMasks,
        IReadOnlyList<PdfTilingPattern> Patterns);

    private sealed record PdfDocumentNumbers(
        IReadOnlyDictionary<string, FontObjectNumbers> FontObjects,
        IReadOnlyDictionary<string, ImageObjectNumbers> ImageObjects,
        IReadOnlyDictionary<string, int> ShadingObjects,
        IReadOnlyDictionary<string, int> SoftMaskObjects,
        IReadOnlyDictionary<string, int> PatternObjects,
        IReadOnlyList<int[]> AnnotationObjectsByPage,
        int ObjectCount,
        int? InfoObjectNumber);

    private static PdfDocumentPlan BuildDocumentPlan(IReadOnlyList<PdfPage> pages, CancellationToken cancellationToken)
    {        List<PdfEmbeddedFont> fonts = pages
            .SelectMany(p => p.Fonts.Select(f => f.Font))
            .GroupBy(f => f.ResourceKey, StringComparer.Ordinal)
            .Select(group => PdfEmbeddedFont.Merge(group, cancellationToken))
            .ToList();
        cancellationToken.ThrowIfCancellationRequested();
        List<PdfImageXObject> images = DeduplicateImages(pages
            .SelectMany(p => p.Images
                .Select(i => i.Image)
                .Concat(p.ExtGStates.Select(s => s.SoftMask?.Image).OfType<PdfImageXObject>())
                .Concat(p.Patterns.SelectMany(pattern => pattern.Pattern.Images.Select(image => image.Image)))),
            static image => image.ResourceKey,
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        List<PdfAxialShading> shadings = pages
            .SelectMany(p => p.Shadings.Select(s => s.Shading))
            .DistinctBy(s => s.ResourceKey)
            .ToList();
        cancellationToken.ThrowIfCancellationRequested();
        List<PdfTilingPattern> patterns = pages
            .SelectMany(p => p.Patterns.Select(s => s.Pattern))
            .DistinctBy(s => s.ResourceKey)
            .ToList();
        List<PdfLuminositySoftMask> softMasks = pages
            .SelectMany(p => p.ExtGStates)
            .Select(s => s.SoftMask).OfType<PdfLuminositySoftMask>()
            .DistinctBy(s => s.ResourceKey)
            .ToList();

        return new PdfDocumentPlan(fonts, images, shadings, softMasks, patterns);
    }


    private static PdfDocumentNumbers AssignDocumentNumbers(PdfDocumentPlan plan, IReadOnlyList<PdfPage> pages, DateTimeOffset? creationDate, CancellationToken cancellationToken)
    {        int fontObjectBase = 3 + pages.Count * 2;
        var fontObjects = new Dictionary<string, FontObjectNumbers>(StringComparer.Ordinal);
        for (int i = 0; i < plan.Fonts.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int baseObject = fontObjectBase + i * 5;
            fontObjects[plan.Fonts[i].ResourceKey] = new FontObjectNumbers(
                Type0: baseObject,
                CidFont: baseObject + 1,
                Descriptor: baseObject + 2,
                FontFile: baseObject + 3,
                ToUnicode: baseObject + 4);
        }

        int imageObjectBase = fontObjectBase + plan.Fonts.Count * 5;
        var imageObjects = new Dictionary<string, ImageObjectNumbers>(StringComparer.Ordinal);
        int nextImageObject = imageObjectBase;
        foreach (PdfImageXObject image in plan.Images)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int imageObject = nextImageObject++;
            int? softMaskObject = image.Alpha is null ? null : nextImageObject++;
            imageObjects[image.ResourceKey] = new ImageObjectNumbers(imageObject, softMaskObject);
        }

        int shadingObjectBase = nextImageObject;
        var shadingObjects = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < plan.Shadings.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            shadingObjects[plan.Shadings[i].ResourceKey] = shadingObjectBase + i;
        }

        int softMaskObjectBase = shadingObjectBase + plan.Shadings.Count;
        var softMaskObjects = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < plan.SoftMasks.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            softMaskObjects[plan.SoftMasks[i].ResourceKey] = softMaskObjectBase + i;
        }

        int patternObjectBase = softMaskObjectBase + plan.SoftMasks.Count;
        var patternObjects = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < plan.Patterns.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            patternObjects[plan.Patterns[i].ResourceKey] = patternObjectBase + i;
        }

        int annotationObjectBase = patternObjectBase + plan.Patterns.Count;
        var annotationObjectsByPage = new List<int[]>(pages.Count);
        int nextAnnotationObject = annotationObjectBase;
        foreach (PdfPage page in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var annotationObjects = new int[page.Annotations.Count];
            for (int i = 0; i < annotationObjects.Length; i++)
            {
                annotationObjects[i] = nextAnnotationObject++;
            }

            annotationObjectsByPage.Add(annotationObjects);
        }

        int objectCount = nextAnnotationObject - 1;
        int? infoObjectNumber = null;
        if (creationDate is not null)
        {
            infoObjectNumber = nextAnnotationObject++;
            objectCount = nextAnnotationObject - 1;
        }

        return new PdfDocumentNumbers(fontObjects, imageObjects, shadingObjects, softMaskObjects, patternObjects, annotationObjectsByPage, objectCount, infoObjectNumber);
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
