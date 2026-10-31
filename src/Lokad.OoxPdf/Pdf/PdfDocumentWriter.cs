using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfDocumentWriter
{
    // R18: image deduplication groups by the full content identity and verifies byte
    // equality on every key match. A same-key content mismatch, including a stale
    // cached identity, fails loudly instead of emitting distinct content as one image.
    internal static List<PdfImageXObject> DeduplicateImages(IEnumerable<PdfImageXObject> images, CancellationToken cancellationToken)
    {
        var byKey = new Dictionary<string, PdfImageXObject>(StringComparer.Ordinal);
        var deduped = new List<PdfImageXObject>();
        foreach (PdfImageXObject image in images)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string key = image.ResourceKey;
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
    // R06.1: returns the exact serialized byte count. Every byte below crosses the
    // PdfObjectWriter boundary, which admits each chunk against the output budget
    // before writing, so a zero budget writes nothing and no call-site whole-output
    // charge remains. On success the admitted total equals the returned position.
    // Pages and content bytes admit per produced page during rendering (R06.2);
    // the writer keeps page validation, planning, and numbering.
    // R06.3: staged emission. The producer enumerable is drained once: each page is
    // validated, its page/group content bytes are encoded a single time and handed
    // to one spill store, and only blanked descriptors are retained. Planning, numbering, and
    // emission then run over descriptors with content re-read in order, so output
    // stays byte-identical while stageable payload retention follows the window.
    public static async ValueTask<long> WriteStagedAsync(
        Stream stream,
        IAsyncEnumerable<PdfPage> pages,
        OoxConversionLimits limits,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        DateTimeOffset? creationDate = null,
        IOoxPdfTemporaryStorage? temporaryStorage = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        PdfStagedDocument staged = await ProduceStagedPagesAsync(pages, limits, diagnosticSink,
            cancellationToken, temporaryStorage).ConfigureAwait(false);
        long result;
        await using (staged.ConfigureAwait(false))
        {
            result = await EmitStagedAsync(stream, staged, cancellationToken, creationDate).ConfigureAwait(false);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }

    // R06.3: drains the producer once (validate, encode, spill, blank) and hands
    // a staged document (descriptors plus store) to the caller, so the stream path can snapshot and report
    // between production and emission. The caller must dispose the document.
    public static async ValueTask<PdfStagedDocument> ProduceStagedPagesAsync(
        IAsyncEnumerable<PdfPage> pages,
        OoxConversionLimits limits,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        IOoxPdfTemporaryStorage? temporaryStorage = null)
    {
        ArgumentNullException.ThrowIfNull(pages);
        ArgumentNullException.ThrowIfNull(limits);
        // Ownership transfers to the caller; dispose on the drain-failure path below.
        var staging = new PdfPageContentStaging(limits.MaxResidentPageContentBytesPerConversion, diagnosticSink, temporaryStorage);
        try
        {
            var blanked = new List<PdfPage>();
            var pageEntries = new List<int>();
            // Weak source identities allow duplicate references to share their
            // staged descriptor without retaining previous pages' payload strings.
            var stagedGroups = new ConditionalWeakTable<PdfTransparencyGroup, PdfTransparencyGroup>();
            var groupEntries = new Dictionary<PdfTransparencyGroup, int>();
            void AdmitRetainedResources(IReadOnlyList<PdfExtGStateResource> states,
                IReadOnlyList<PdfShadingResource> shadings, IReadOnlyList<PdfTilingPatternResource> patterns,
                IReadOnlyList<PdfTransparencyGroupResource> groups, IReadOnlyList<PdfShadingPatternResource> shadingPatterns,
                long descriptorCapacity = 0)
            {
                if (OoxConversionBudget.Current?.Memory is null) return;
                // Coarse retained records, collection/map headroom and normalized
                // stops/keys. This funds staged descriptors and planning headroom,
                // not earlier painting.
                long bytes = checked(descriptorCapacity + 128L * ((long)states.Count + shadings.Count + patterns.Count + groups.Count + shadingPatterns.Count));
                foreach (var shading in shadings)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    bytes = checked(bytes + 128L * shading.Shading.Stops.Count);
                }
                foreach (var pattern in shadingPatterns)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    bytes = checked(bytes + 128L * pattern.Pattern.Shading.Stops.Count);
                }
                foreach (var pattern in patterns)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    bytes = checked(bytes + 4L * pattern.Pattern.Content.Length + 128L * pattern.Pattern.Images.Count);
                }
                if (bytes > 0) staging.RetainResourceCapacity(bytes);
                cancellationToken.ThrowIfCancellationRequested();
            }
            async ValueTask<PdfTransparencyGroupResource[]> StageGroupsAsync(IReadOnlyList<PdfTransparencyGroupResource> groups)
            {
                var result = new PdfTransparencyGroupResource[groups.Count];
                for (int index = 0; index < groups.Count; index++)
                    result[index] = new(groups[index].ResourceName, await StageGroupAsync(groups[index].Group).ConfigureAwait(false));
                return result;
            }
            async ValueTask<PdfTransparencyGroup> StageGroupAsync(PdfTransparencyGroup group)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (stagedGroups.TryGetValue(group, out PdfTransparencyGroup? existing)) return existing;
                AdmitRetainedResources(group.ExtGStates, group.Shadings, [], group.Groups, group.ShadingPatterns, 1024);
                var children = await StageGroupsAsync(group.Groups).ConfigureAwait(false);
                int entry = staging.Count;
                staging.BeginPage(cancellationToken);
                await AppendEncodedPageContentAsync(staging, group.Content, cancellationToken).ConfigureAwait(false);
                await staging.EndPageAsync(cancellationToken).ConfigureAwait(false);
                var descriptor = new PdfTransparencyGroup(group.Bounds, string.Empty, group.ExtGStates, group.Shadings, children, group.ShadingPatterns);
                stagedGroups.Add(group, descriptor);
                groupEntries.Add(descriptor, entry);
                return descriptor;
            }
            int pageIndex = 0;
            await foreach (PdfPage page in pages.WithCancellation(cancellationToken).ConfigureAwait(false))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!double.IsFinite(page.Width) || !double.IsFinite(page.Height) || page.Width <= 0 || page.Height <= 0)
                {
                    throw new ArgumentOutOfRangeException(nameof(pages), "PDF pages must have finite positive dimensions.");
                }

                // Page/store records survive content spilling; the same window
                // also covers ordinary resource-reference and numbering overlap.
                long pageMetadata = checked(1024 + 128L * ((long)page.Fonts.Count + page.FallbackFonts.Count + page.Images.Count + page.Annotations.Count));
                AdmitRetainedResources(page.ExtGStates, page.Shadings, page.Patterns, page.Groups, page.ShadingPatterns, pageMetadata);
                PdfContentValidator.ValidatePage(page, pageIndex, cancellationToken);
                pageEntries.Add(staging.Count);
                staging.BeginPage(cancellationToken);
                await AppendEncodedPageContentAsync(staging, page.Content, cancellationToken).ConfigureAwait(false);
                await staging.EndPageAsync(cancellationToken).ConfigureAwait(false);
                var groups = await StageGroupsAsync(page.Groups).ConfigureAwait(false);
                blanked.Add(page with { Content = string.Empty, Groups = groups });
                pageIndex++;
            }

            if (blanked.Count == 0)
            {
                throw new ArgumentException("A PDF document must contain at least one page.", nameof(pages));
            }

            if (staging.SpilledBytes > 0)
            {
                OoxConversionBudget.Current?.NotePageContentSpilled(staging.SpilledBytes);
            }

            return new PdfStagedDocument(blanked, staging, pageEntries, groupEntries);
        }
        catch
        {
            await staging.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    // RV11-P1: bounded-scratch ASCII encoding. One ASCII char always encodes to
    // exactly one byte (non-ASCII becomes ?), so any char-boundary chunking is
    // byte-identical to whole-string encoding. The store owns and admits the
    // single scratch chunk, reused for emission. Validation rejects non-ASCII.
    private static async ValueTask AppendEncodedPageContentAsync(PdfPageContentStaging staging, string content, CancellationToken cancellationToken)
    {
        byte[] scratch = staging.Scratch;
        int offset = 0;
        while (offset < content.Length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int chars = Math.Min(scratch.Length, content.Length - offset);
            int bytes = Encoding.ASCII.GetBytes(content, offset, chars, scratch, 0);
            await staging.AppendPageBytesAsync(scratch, 0, bytes, cancellationToken).ConfigureAwait(false);
            offset += chars;
        }
    }

    // R06.3: emits already-produced descriptors with content re-read in order.
    public static async ValueTask<long> EmitStagedAsync(
        Stream stream,
        PdfStagedDocument staged,
        CancellationToken cancellationToken,
        DateTimeOffset? creationDate = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(staged);
        staged.ThrowIfDisposed();
        IReadOnlyList<PdfPage> pages = staged.Pages;
        PdfPageContentStaging staging = staged.Staging;
        if (staging.Count != pages.Count + staged.GroupEntries.Count || staged.PageEntries.Count != pages.Count)
        {
            throw new InvalidDataException("Staged content entry count mismatch.");
        }

        return await WriteCoreAsync(stream, pages, staging, staged.PageEntries, staged.GroupEntries, cancellationToken, creationDate).ConfigureAwait(false);
    }

    private static async ValueTask<long> WriteCoreAsync(Stream stream, IReadOnlyList<PdfPage> pages, PdfPageContentStaging staging, IReadOnlyList<int> pageEntries, IReadOnlyDictionary<PdfTransparencyGroup, int> groupEntries, CancellationToken cancellationToken, DateTimeOffset? creationDate = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var writer = new PdfObjectWriter(stream, cancellationToken, staging.Scratch);
        await writer.WriteHeaderAsync().ConfigureAwait(false);

        PdfDocumentPlan plan = BuildDocumentPlan(pages, cancellationToken);
        PdfDocumentNumbers numbers = AssignDocumentNumbers(plan, pages, creationDate, cancellationToken);

        await writer.WriteObjectAsync(1, "<< /Type /Catalog /Pages 2 0 R >>\n").ConfigureAwait(false);
        await writer.WriteObjectAsync(2, BuildPagesObject()).ConfigureAwait(false);

        for (int i = 0; i < pages.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int pageObjectNumber = 3 + i * 2;
            int contentObjectNumber = pageObjectNumber + 1;
            PdfPage page = pages[i];

            await writer.WriteObjectAsync(pageObjectNumber, FormattableString.Invariant(
                $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {FormatNumber(page.Width)} {FormatNumber(page.Height)}] /Contents {contentObjectNumber} 0 R /Resources {BuildResources(page)}{BuildPageAnnotations(numbers.AnnotationObjectsByPage[i])} >>\n")).ConfigureAwait(false);
            int contentLength = await staging.GetPageLengthAsync(pageEntries[i], cancellationToken).ConfigureAwait(false);
            await writer.WriteContentStreamHeaderAsync(contentObjectNumber, contentLength).ConfigureAwait(false);
            await CopyStagedContentAsync(writer, staging, pageEntries[i], contentLength, cancellationToken).ConfigureAwait(false);

            await writer.WriteContentStreamTrailerAsync().ConfigureAwait(false);
        }

        foreach (PdfEmbeddedFont font in plan.Fonts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await WriteFontObjectsAsync(writer, font, numbers.FontObjects[font.ResourceKey], cancellationToken).ConfigureAwait(false);
        }

        foreach (PdfFallbackFont font in plan.FallbackFonts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await WriteFallbackFontObjectsAsync(writer, font, numbers.FallbackFontObjects[font.ResourceKey], cancellationToken).ConfigureAwait(false);
        }

        foreach (PdfImageXObject image in plan.Images)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await WriteImageObjectsAsync(writer, image, numbers.ImageObjects[image.ResourceKey]).ConfigureAwait(false);
        }

        foreach (PdfShading shading in plan.Shadings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await WriteShadingObjectAsync(writer, shading, numbers.ShadingObjects[shading.ResourceKey]).ConfigureAwait(false);
        }

        foreach (PdfTilingPattern pattern in plan.Patterns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await WriteTilingPatternObjectAsync(writer, pattern, numbers.PatternObjects[pattern.ResourceKey], numbers.ImageObjects).ConfigureAwait(false);
        }

        foreach (PdfShadingPattern pattern in plan.ShadingPatterns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PdfPatternMatrix m = pattern.Matrix;
            await writer.WriteObjectAsync(numbers.ShadingPatternObjects[pattern.ResourceKey], FormattableString.Invariant(
                $"<< /Type /Pattern /PatternType 2 /Shading {numbers.ShadingObjects[pattern.Shading.ResourceKey]} 0 R /Matrix [{FormatNumber(m.A)} {FormatNumber(m.B)} {FormatNumber(m.C)} {FormatNumber(m.D)} {FormatNumber(m.E)} {FormatNumber(m.F)}] >>\n")).ConfigureAwait(false);
        }

        foreach (PdfTransparencyGroup group in plan.Groups)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int entry = groupEntries[group];
            PdfRectangle bounds = group.Bounds;
            string resources = BuildResourceDictionary([], [], [], group.ExtGStates, group.Shadings, [], group.Groups, group.ShadingPatterns);
            int contentLength = await staging.GetPageLengthAsync(entry, cancellationToken).ConfigureAwait(false);
            await writer.WriteContentStreamHeaderAsync(numbers.GroupObjects[group], contentLength,
                $"/Type /XObject /Subtype /Form /BBox [{FormatNumber(bounds.X)} {FormatNumber(bounds.Y)} {FormatNumber(bounds.X + bounds.Width)} {FormatNumber(bounds.Y + bounds.Height)}] /Group << /S /Transparency /CS /DeviceRGB /I true /K false >> /Resources {resources}").ConfigureAwait(false);
            await CopyStagedContentAsync(writer, staging, entry, contentLength, cancellationToken).ConfigureAwait(false);
            await writer.WriteContentStreamTrailerAsync().ConfigureAwait(false);
        }

        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int[] annotationObjects = numbers.AnnotationObjectsByPage[pageIndex];
            IReadOnlyList<PdfLinkAnnotation> annotations = pages[pageIndex].Annotations;
            for (int annotationIndex = 0; annotationIndex < annotations.Count; annotationIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await WriteLinkAnnotationObjectAsync(writer, annotations[annotationIndex], annotationObjects[annotationIndex], pages.Count).ConfigureAwait(false);
            }
        }

        if (numbers.InfoObjectNumber is not null && creationDate is not null)
        {
            await writer.WriteObjectAsync(numbers.InfoObjectNumber.Value, BuildInfoObject(creationDate.Value)).ConfigureAwait(false);
        }

        if (writer.Offsets.Count != numbers.ObjectCount)
        {
            throw new InvalidDataException($"PDF object numbering is inconsistent: {writer.Offsets.Count} written objects but {numbers.ObjectCount} counted objects.");
        }

        long xrefOffset = writer.Position;
        await writer.WriteAsciiAsync(FormattableString.Invariant($"xref\n0 {numbers.ObjectCount + 1}\n")).ConfigureAwait(false);
        await writer.WriteAsciiAsync("0000000000 65535 f \n").ConfigureAwait(false);
        foreach (long offset in writer.Offsets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await writer.WriteAsciiAsync(FormattableString.Invariant($"{offset:0000000000} 00000 n \n")).ConfigureAwait(false);
        }

        await writer.WriteAsciiAsync(FormattableString.Invariant(
            $"trailer\n<< /Size {numbers.ObjectCount + 1} /Root 1 0 R{(numbers.InfoObjectNumber is null ? string.Empty : FormattableString.Invariant($" /Info {numbers.InfoObjectNumber.Value} 0 R"))} >>\nstartxref\n{xrefOffset}\n%%EOF\n")).ConfigureAwait(false);

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

        string BuildResources(PdfPage page) => BuildResourceDictionary(page.Fonts, page.FallbackFonts, page.Images, page.ExtGStates, page.Shadings, page.Patterns, page.Groups, page.ShadingPatterns);

        string BuildResourceDictionary(IReadOnlyList<PdfFontResource> fonts, IReadOnlyList<PdfFallbackFontResource> fallbackFonts,
            IReadOnlyList<PdfImageResource> images, IReadOnlyList<PdfExtGStateResource> states, IReadOnlyList<PdfShadingResource> shadings,
            IReadOnlyList<PdfTilingPatternResource> patterns, IReadOnlyList<PdfTransparencyGroupResource> groups, IReadOnlyList<PdfShadingPatternResource> shadingPatterns)
        {
            if (fonts.Count == 0 && fallbackFonts.Count == 0 && images.Count == 0 && states.Count == 0 && shadings.Count == 0 && patterns.Count == 0 && groups.Count == 0 && shadingPatterns.Count == 0)
            {
                return "<< >>";
            }

            var builder = new StringBuilder("<<");
            Dictionary<string, PdfTransparencyGroup>? maskGroups = null;
            if (fonts.Count != 0 || fallbackFonts.Count != 0)
            {
                builder.Append(" /Font <<");
                foreach (PdfFontResource font in fonts)
                {
                    FontObjectNumbers objects = numbers.FontObjects[font.Font.ResourceKey];
                    builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(font.ResourceName)).Append(' ');
                    builder.Append(CultureInfo.InvariantCulture, $"{objects.Type0} 0 R");
                }

                foreach (PdfFallbackFontResource font in fallbackFonts)
                {
                    FallbackFontObjectNumbers objects = numbers.FallbackFontObjects[font.Font.ResourceKey];
                    builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(font.ResourceName)).Append(' ');
                    builder.Append(CultureInfo.InvariantCulture, $"{objects.Font} 0 R");
                }

                builder.Append(" >>");
            }

            if (images.Count != 0 || groups.Count != 0)
            {
                builder.Append(" /XObject <<");
                foreach (PdfImageResource image in images)
                {
                    ImageObjectNumbers objects = numbers.ImageObjects[image.Image.ResourceKey];
                    builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(image.ResourceName)).Append(' ');
                    builder.Append(CultureInfo.InvariantCulture, $"{objects.Image} 0 R");
                }

                foreach (PdfTransparencyGroupResource group in groups)
                {
                    builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(group.ResourceName)).Append(' ');
                    builder.Append(CultureInfo.InvariantCulture, $"{numbers.GroupObjects[group.Group]} 0 R");
                }
                builder.Append(" >>");
            }

            if (states.Count != 0)
            {
                builder.Append(" /ExtGState <<");
                foreach (PdfExtGStateResource state in states)
                {
                    builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(state.ResourceName));
                    builder.Append(CultureInfo.InvariantCulture, $" << /ca {FormatNumber(state.FillAlpha)} /CA {FormatNumber(state.StrokeAlpha)}");
                    if (state.LuminosityGroupName is { } name)
                    {
                        maskGroups ??= groups.ToDictionary(resource => PdfEmbeddedFont.SanitizeName(resource.ResourceName), resource => resource.Group, StringComparer.Ordinal);
                        PdfTransparencyGroup group = maskGroups[PdfEmbeddedFont.SanitizeName(name)];
                        builder.Append(CultureInfo.InvariantCulture, $" /SMask << /S /Luminosity /G {numbers.GroupObjects[group]} 0 R /BC [0 0 0] >>");
                    }

                    builder.Append(" >>");
                }

                builder.Append(" >>");
            }

            if (shadings.Count != 0)
            {
                builder.Append(" /Shading <<");
                foreach (PdfShadingResource shading in shadings)
                {
                    builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(shading.ResourceName));
                    builder.Append(CultureInfo.InvariantCulture, $" {numbers.ShadingObjects[shading.Shading.ResourceKey]} 0 R");
                }

                builder.Append(" >>");
            }

            if (patterns.Count != 0 || shadingPatterns.Count != 0)
            {
                builder.Append(" /Pattern <<");
                foreach (PdfTilingPatternResource pattern in patterns)
                {
                    builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(pattern.ResourceName));
                    builder.Append(CultureInfo.InvariantCulture, $" {numbers.PatternObjects[pattern.Pattern.ResourceKey]} 0 R");
                }

                foreach (PdfShadingPatternResource pattern in shadingPatterns)
                {
                    builder.Append(" /").Append(PdfEmbeddedFont.SanitizeName(pattern.ResourceName));
                    builder.Append(CultureInfo.InvariantCulture, $" {numbers.ShadingPatternObjects[pattern.Pattern.ResourceKey]} 0 R");
                }

                builder.Append(" >>");
            }

            builder.Append(" >>");
            return builder.ToString();
        }
    }

    private static async ValueTask CopyStagedContentAsync(PdfObjectWriter writer, PdfPageContentStaging staging, int entry, int length, CancellationToken cancellationToken)
    {
        byte[] copy = staging.Scratch;
        int offset = 0;
        while (offset < length)
        {
            int read = await staging.ReadPageBytesAsync(entry, copy, 0, offset, Math.Min(copy.Length, length - offset), cancellationToken).ConfigureAwait(false);
            if (read == 0) { throw new InvalidDataException("Content spill store ended unexpectedly."); }
            await writer.WriteRawBytesAsync(copy.AsMemory(0, read)).ConfigureAwait(false);
            offset += read;
        }
    }

    private static async ValueTask WriteFontObjectsAsync(PdfObjectWriter writer, PdfEmbeddedFont font, FontObjectNumbers objects, CancellationToken cancellationToken)
    {
        string baseFont = PdfEmbeddedFont.SanitizeName(font.BaseFontName);
        await writer.WriteObjectAsync(objects.Type0, FormattableString.Invariant(
            $"<< /Type /Font /Subtype /Type0 /BaseFont /{baseFont} /Encoding /Identity-H /DescendantFonts [{objects.CidFont} 0 R] /ToUnicode {objects.ToUnicode} 0 R >>\n")).ConfigureAwait(false);

        await writer.WriteObjectHeaderAsync(objects.CidFont).ConfigureAwait(false);
        await writer.WriteAsciiAsync(FormattableString.Invariant(
            $"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{baseFont} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor {objects.Descriptor} 0 R /CIDToGIDMap /Identity /W ")).ConfigureAwait(false);
        await font.WriteWidthArrayAsync(writer, cancellationToken).ConfigureAwait(false);
        await writer.WriteAsciiAsync(" >>\n").ConfigureAwait(false);
        await writer.WriteObjectTrailerAsync().ConfigureAwait(false);

        OpenTypeFontMetrics metrics = OpenTypeFontMetrics.From(font.Font);
        await writer.WriteObjectAsync(objects.Descriptor, FormattableString.Invariant(
            $"<< /Type /FontDescriptor /FontName /{baseFont} /Flags {metrics.Flags} /FontBBox [{metrics.XMin} {metrics.YMin} {metrics.XMax} {metrics.YMax}] /ItalicAngle {FormatNumber(metrics.ItalicAngle)} /Ascent {metrics.Ascent} /Descent {metrics.Descent} /CapHeight {metrics.CapHeight} /StemV 80 /FontFile2 {objects.FontFile} 0 R >>\n")).ConfigureAwait(false);

        ReadOnlyMemory<byte> fontProgram = font.FontProgramBytes;
        int toUnicodeLength = font.MeasureToUnicodeCMap(writer.Scratch, cancellationToken);
        using (IMemoryReservation? storage = OoxMemory.Reserve(OoxConversionBudget.Current?.Memory, 0))
        {
            byte[] compressedFontProgram = PdfStreamCompression.Compress(fontProgram.Span, storage, 0, cancellationToken);
            // Count uncompressed font/ToUnicode payloads, separately from
            // dictionaries/output and from the compression buffer's admission.
            OoxConversionBudget.Current?.ChargePdfFontBytes(checked((long)fontProgram.Length + toUnicodeLength));
            await writer.WriteStreamObjectAsync(objects.FontFile, FormattableString.Invariant($"/Filter /FlateDecode /Length1 {fontProgram.Length}"), compressedFontProgram).ConfigureAwait(false);
        }
        await writer.WriteStreamObjectHeaderAsync(objects.ToUnicode, string.Empty, toUnicodeLength).ConfigureAwait(false);
        await font.WriteToUnicodeCMapAsync(writer, cancellationToken).ConfigureAwait(false);
        await writer.WriteStreamObjectTrailerAsync().ConfigureAwait(false);
    }

    // Standard-14 faces have no embedded font program. Their dictionaries and
    // static WinAnsi ToUnicode map count toward output, not the font-payload total.
    private static async ValueTask WriteFallbackFontObjectsAsync(PdfObjectWriter writer, PdfFallbackFont font, FallbackFontObjectNumbers objects, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await writer.WriteObjectAsync(objects.Font, FormattableString.Invariant($"<< /Type /Font /Subtype /Type1 /BaseFont /{font.FaceName} /Encoding /WinAnsiEncoding /ToUnicode {objects.ToUnicode} 0 R >>\n")).ConfigureAwait(false);

        string toUnicode = PdfFallbackFont.WinAnsiToUnicodeCMap;
        await writer.WriteStreamObjectHeaderAsync(objects.ToUnicode, string.Empty, Encoding.ASCII.GetByteCount(toUnicode)).ConfigureAwait(false);
        await writer.WriteAsciiAsync(toUnicode).ConfigureAwait(false);
        await writer.WriteStreamObjectTrailerAsync().ConfigureAwait(false);
    }

    private static string BuildInfoObject(DateTimeOffset creationDate)
    {
        string producer = "Lokad.OoxPdf " + (typeof(PdfDocumentWriter).Assembly.GetName().Version?.ToString() ?? "0");
        string date = FormatPdfDate(creationDate);
        return FormattableString.Invariant($"<< /Producer ({producer}) /CreationDate ({date}) /ModDate ({date}) >>\n");
    }

    private static async ValueTask WriteImageObjectsAsync(PdfObjectWriter writer, PdfImageXObject image, ImageObjectNumbers objects)
    {
        // R06: retained image resources accumulate across embedded images (soft masks
        // ride the same path); charge encoded bytes as they serialize.
        OoxConversionBudget.Current?.ChargePdfImageBytes(checked((long)image.Bytes.Length + (image.Alpha?.Length ?? 0)));
        string smask = objects.SoftMask is null ? string.Empty : FormattableString.Invariant($" /SMask {objects.SoftMask.Value} 0 R");
        string decode = image.InvertedCmyk ? " /Decode [1 0 1 0 1 0 1 0]" : string.Empty;
        await writer.WriteStreamObjectAsync(objects.Image, FormattableString.Invariant(
            $"/Type /XObject /Subtype /Image /Width {image.Width} /Height {image.Height} /ColorSpace {image.ColorSpace} /BitsPerComponent {image.BitsPerComponent} /Filter {image.Filter}{decode}{smask}"), image.Bytes).ConfigureAwait(false);
        if (image.Alpha is not null && objects.SoftMask is not null)
        {
            await writer.WriteStreamObjectAsync(objects.SoftMask.Value, FormattableString.Invariant(
                $"/Type /XObject /Subtype /Image /Width {image.Width} /Height {image.Height} /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode"), image.Alpha).ConfigureAwait(false);
        }
    }

    private static async ValueTask WriteShadingObjectAsync(PdfObjectWriter writer, PdfShading shading, int objectNumber)
    {
        (int type, string coordinates) = shading switch
        {
            PdfAxialShading axial => (2, $"{FormatNumber(axial.X0)} {FormatNumber(axial.Y0)} {FormatNumber(axial.X1)} {FormatNumber(axial.Y1)}"),
            PdfRadialShading => (3, "0 0 0 0 0 1"),
            _ => throw new InvalidOperationException("Unknown PDF shading type.")
        };
        string function = BuildShadingFunction(shading.Stops);
        if (shading is PdfRadialShading { CycleCount: > 1 } radial)
        {
            var repeated = new StringBuilder("<< /FunctionType 3 /Domain [0 1] /Functions [");
            for (int cycle = 0; cycle < radial.CycleCount; cycle++)
            {
                repeated.Append(' ').Append(function);
            }
            repeated.Append(" ] /Bounds [");
            for (int cycle = 1; cycle < radial.CycleCount; cycle++)
            {
                if (cycle > 1) { repeated.Append(' '); }
                repeated.Append(((double)cycle / radial.CycleCount).ToString("0.######", CultureInfo.InvariantCulture));
            }
            repeated.Append("] /Encode [");
            for (int cycle = 0; cycle < radial.CycleCount; cycle++)
            {
                if (cycle > 0) { repeated.Append(' '); }
                repeated.Append(radial.Reflect && cycle % 2 == 1 ? "1 0" : "0 1");
            }
            function = repeated.Append("] >>").ToString();
        }
        await writer.WriteObjectAsync(objectNumber, FormattableString.Invariant(
            $"<< /ShadingType {type} /ColorSpace /DeviceRGB /Coords [{coordinates}] /Function {function} /Extend [true true] >>\n")).ConfigureAwait(false);

        string BuildShadingFunction(IReadOnlyList<PdfShadingStop> stops)
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

    private static async ValueTask WriteTilingPatternObjectAsync(
        PdfObjectWriter writer,
        PdfTilingPattern pattern,
        int objectNumber,
        IReadOnlyDictionary<string, ImageObjectNumbers> imageObjects)
    {
        int contentLength = Encoding.ASCII.GetByteCount(pattern.Content);
        string matrix = pattern.Matrix is { } m
            ? FormattableString.Invariant($" /Matrix [{FormatNumber(m.A)} {FormatNumber(m.B)} {FormatNumber(m.C)} {FormatNumber(m.D)} {FormatNumber(m.E)} {FormatNumber(m.F)}]")
            : string.Empty;
        string resources = pattern.Images.Count == 0
            ? "<< >>"
            : BuildPatternResources();
        await writer.WriteObjectAsync(objectNumber, FormattableString.Invariant(
            $"<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType {pattern.TilingType} /BBox [0 0 {FormatNumber(pattern.Width)} {FormatNumber(pattern.Height)}]{matrix} /XStep {FormatNumber(pattern.XStep)} /YStep {FormatNumber(pattern.YStep)} /Resources {resources} /Length {contentLength} >>\nstream\n{pattern.Content}endstream\n")).ConfigureAwait(false);

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

    private static async ValueTask WriteLinkAnnotationObjectAsync(
        PdfObjectWriter writer,
        PdfLinkAnnotation annotation,
        int objectNumber,
        int pageCount)
    {
        double x1 = annotation.X;
        double y1 = annotation.Y;
        double x2 = annotation.X + annotation.Width;
        double y2 = annotation.Y + annotation.Height;
        await writer.WriteObjectAsync(objectNumber, FormattableString.Invariant(
            $"<< /Type /Annot /Subtype /Link /Rect [{FormatNumber(x1)} {FormatNumber(y1)} {FormatNumber(x2)} {FormatNumber(y2)}] /Border [0 0 0]{BuildLinkTarget()} >>\n")).ConfigureAwait(false);

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
        IReadOnlyList<PdfFallbackFont> FallbackFonts,
        IReadOnlyList<PdfImageXObject> Images,
        IReadOnlyList<PdfShading> Shadings,
        IReadOnlyList<PdfTilingPattern> Patterns,
        IReadOnlyList<PdfShadingPattern> ShadingPatterns,
        IReadOnlyList<PdfTransparencyGroup> Groups);

    private sealed record PdfDocumentNumbers(
        IReadOnlyDictionary<string, FontObjectNumbers> FontObjects,
        IReadOnlyDictionary<string, FallbackFontObjectNumbers> FallbackFontObjects,
        IReadOnlyDictionary<string, ImageObjectNumbers> ImageObjects,
        IReadOnlyDictionary<string, int> ShadingObjects,
        IReadOnlyDictionary<string, int> PatternObjects,
        IReadOnlyDictionary<string, int> ShadingPatternObjects,
        IReadOnlyDictionary<PdfTransparencyGroup, int> GroupObjects,
        IReadOnlyList<int[]> AnnotationObjectsByPage,
        int ObjectCount,
        int? InfoObjectNumber);

    private static PdfDocumentPlan BuildDocumentPlan(IReadOnlyList<PdfPage> pages, CancellationToken cancellationToken)
    {
        var groups = new List<PdfTransparencyGroup>();
        var seen = new HashSet<PdfTransparencyGroup>();
        void Gather(PdfTransparencyGroup group)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!seen.Add(group)) { return; }
            groups.Add(group);
            foreach (PdfTransparencyGroupResource child in group.Groups) { Gather(child.Group); }
        }
        foreach (PdfPage page in pages)
        {
            foreach (PdfTransparencyGroupResource group in page.Groups) { Gather(group.Group); }
        }
        List<PdfEmbeddedFont> fonts = pages
            .SelectMany(p => p.Fonts.Select(f => f.Font))
            .GroupBy(f => f.ResourceKey, StringComparer.Ordinal)
            .Select(group => PdfEmbeddedFont.Merge(group, cancellationToken))
            .ToList();
        cancellationToken.ThrowIfCancellationRequested();
        List<PdfFallbackFont> fallbackFonts = pages
            .SelectMany(p => p.FallbackFonts.Select(f => f.Font))
            .DistinctBy(f => f.ResourceKey, StringComparer.Ordinal)
            .ToList();
        cancellationToken.ThrowIfCancellationRequested();
        List<PdfImageXObject> images = DeduplicateImages(pages
            .SelectMany(p => p.Images
                .Select(i => i.Image)
                .Concat(p.Patterns.SelectMany(pattern => pattern.Pattern.Images.Select(image => image.Image)))),
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        List<PdfShading> shadings = pages
            .SelectMany(p => p.Shadings.Select(s => s.Shading))
            .Concat(groups.SelectMany(group => group.Shadings.Select(shading => shading.Shading)))
            .Concat(pages.SelectMany(page => page.ShadingPatterns.Select(pattern => pattern.Pattern.Shading)))
            .Concat(groups.SelectMany(group => group.ShadingPatterns.Select(pattern => pattern.Pattern.Shading)))
            .DistinctBy(s => s.ResourceKey)
            .ToList();
        cancellationToken.ThrowIfCancellationRequested();
        List<PdfTilingPattern> patterns = pages
            .SelectMany(p => p.Patterns.Select(s => s.Pattern))
            .DistinctBy(s => s.ResourceKey)
            .ToList();
        List<PdfShadingPattern> shadingPatterns = pages
            .SelectMany(page => page.ShadingPatterns.Select(pattern => pattern.Pattern))
            .Concat(groups.SelectMany(group => group.ShadingPatterns.Select(pattern => pattern.Pattern)))
            .DistinctBy(pattern => pattern.ResourceKey, StringComparer.Ordinal)
            .ToList();
        return new PdfDocumentPlan(fonts, fallbackFonts, images, shadings, patterns, shadingPatterns, groups);
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

        var fallbackFontObjects = new Dictionary<string, FallbackFontObjectNumbers>(StringComparer.Ordinal);
        for (int i = 0; i < plan.FallbackFonts.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int baseObject = fontObjectBase + plan.Fonts.Count * 5 + i * 2;
            fallbackFontObjects[plan.FallbackFonts[i].ResourceKey] = new FallbackFontObjectNumbers(
                Font: baseObject,
                ToUnicode: baseObject + 1);
        }

        int imageObjectBase = fontObjectBase + plan.Fonts.Count * 5 + plan.FallbackFonts.Count * 2;
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

        int patternObjectBase = shadingObjectBase + plan.Shadings.Count;
        var patternObjects = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < plan.Patterns.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            patternObjects[plan.Patterns[i].ResourceKey] = patternObjectBase + i;
        }

        int shadingPatternObjectBase = patternObjectBase + plan.Patterns.Count;
        var shadingPatternObjects = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < plan.ShadingPatterns.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            shadingPatternObjects[plan.ShadingPatterns[i].ResourceKey] = shadingPatternObjectBase + i;
        }
        int groupObjectBase = shadingPatternObjectBase + plan.ShadingPatterns.Count;
        var groupObjects = new Dictionary<PdfTransparencyGroup, int>();
        for (int i = 0; i < plan.Groups.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            groupObjects.Add(plan.Groups[i], groupObjectBase + i);
        }
        int annotationObjectBase = groupObjectBase + plan.Groups.Count;
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

        return new PdfDocumentNumbers(fontObjects, fallbackFontObjects, imageObjects, shadingObjects, patternObjects, shadingPatternObjects, groupObjects, annotationObjectsByPage, objectCount, infoObjectNumber);
    }
    private readonly record struct FontObjectNumbers(int Type0, int CidFont, int Descriptor, int FontFile, int ToUnicode);

    private readonly record struct FallbackFontObjectNumbers(int Font, int ToUnicode);

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
