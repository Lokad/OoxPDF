using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed class PdfGraphicsBuilder
{
    internal const double SyntheticItalicShear = 0.213d;

    private readonly StringBuilder builder = new();
    private readonly List<PdfExtGStateResource> extGStates = [];
    private readonly List<PdfShadingResource> shadings = [];
    private readonly List<PdfTilingPatternResource> patterns = [];
    private readonly List<PdfShadingPatternResource> shadingPatterns = [];
    private readonly List<PdfTransparencyGroupResource> groups = [];
    private readonly HashSet<string> usedFontResourceNames = new(StringComparer.Ordinal);
    private readonly List<string> fontUseOrder = [];
    // Each registration key belongs to one resource entry. Rolling back the
    // appended entries keeps the keys and slot numbers of surviving resources.
    private readonly HashSet<string> extGStateNames = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> shadingIndex = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> patternIndex = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> shadingPatternIndex = new(StringComparer.Ordinal);
    private int stateDepth;
    private bool actualTextOpen;

    // Borrow the producer's existing page grant; the builder never owns/refunds it.
    private readonly IMemoryReservation? construction;
    private readonly long constructionBaseBytes;
    private readonly long svgSourceBytes;
    private long admittedTextCapacity;

    public PdfGraphicsBuilder() { }

    public PdfGraphicsBuilder(IMemoryReservation? construction, long svgSourceBytes)
    {
        this.construction = construction;
        constructionBaseBytes = construction?.Bytes ?? 0L;
        this.svgSourceBytes = svgSourceBytes;
    }

    internal int ContentLength => builder.Length;

    internal void AdmitSvgConstruction(long sampledBytes, int sourceBytes, int precedingCharacters)
    {
        if (construction is null) return;
        // One observed expansion density covers all SVG placements on this page,
        // including completed opacity groups. Prefix operators coexist with the
        // expanded image. This is coarse construction/snapshot headroom, not a
        // heap ceiling, per-path reservation or cumulative work allowance.
        long density = checked((sampledBytes + sourceBytes - 1L) / sourceBytes);
        long capacity = checked(constructionBaseBytes + 4L * precedingCharacters + density * Math.Max(svgSourceBytes, sourceBytes));
        if (capacity > construction.Bytes) OoxMemory.Resize(construction, capacity);
    }

    private void AdmitTextConstruction(long characters)
    {
        if (construction is null) return;
        long required = checked(builder.Length + characters);
        // Append can grow the current chunk. Keep doubling headroom without
        // changing StringBuilder's allocation policy. The page owner holds
        // both its UTF16 storage and the final immutable content snapshot.
        long capacity = required > builder.Capacity ? Math.Max(required, 2L * builder.Capacity) : builder.Capacity;
        if (capacity <= admittedTextCapacity) return;
        long bytes = checked(constructionBaseBytes + 4L * capacity);
        if (bytes > construction.Bytes) OoxMemory.Resize(construction, bytes);
        admittedTextCapacity = capacity;
    }

    public IReadOnlyList<PdfExtGStateResource> ExtGStates => extGStates;

    public IReadOnlyList<PdfShadingResource> Shadings => shadings;

    public IReadOnlyList<PdfTilingPatternResource> Patterns => patterns;

    public IReadOnlyList<PdfShadingPatternResource> ShadingPatterns => shadingPatterns;

    public IReadOnlySet<string> UsedFontResourceNames => usedFontResourceNames;

    public int StateDepth => stateDepth;

    public IReadOnlyList<PdfTransparencyGroupResource> Groups => groups;

    public void DrawTransparencyGroup(PdfGraphicsBuilder child, PdfRectangle bounds, double opacity)
    {
        string name = RegisterTransparencyGroup(child, bounds);
        SaveState();
        SetAlpha(opacity, opacity);
        builder.Append('/').Append(name).AppendLine(" Do");
        RestoreState();
    }

    public void SetVectorLuminositySoftMask(PdfGraphicsBuilder child, PdfRectangle bounds, double fillAlpha, double strokeAlpha)
    {
        string groupName = RegisterTransparencyGroup(child, bounds);
        string stateName = "GSV" + (extGStates.Count + 1).ToString(CultureInfo.InvariantCulture);
        extGStateNames.Add(stateName);
        extGStates.Add(new PdfExtGStateResource(stateName, Math.Clamp(fillAlpha, 0d, 1d), Math.Clamp(strokeAlpha, 0d, 1d), groupName));
        builder.Append('/').Append(stateName).AppendLine(" gs");
    }

    private string RegisterTransparencyGroup(PdfGraphicsBuilder child, PdfRectangle bounds)
    {
        ArgumentNullException.ThrowIfNull(child);
        if (child.stateDepth != 0 || child.actualTextOpen || child.usedFontResourceNames.Count != 0 || child.patterns.Count != 0)
        {
            throw new InvalidDataException("SVG transparency groups require balanced vector content.");
        }
        // Count nested output once before snapshotting. Producer construction and
        // staged-resource reservations independently admit their memory lifetimes.
        OoxConversionBudget.Current?.ChargePdfContentBytes(child.builder.Length);
        var group = new PdfTransparencyGroup(bounds, child.builder.ToString(), child.extGStates, child.shadings, child.groups, child.shadingPatterns);
        string name = "Tr" + (groups.Count + 1).ToString(CultureInfo.InvariantCulture);
        groups.Add(new PdfTransparencyGroupResource(name, group));
        return name;
    }

    public void SetFillRgb(byte red, byte green, byte blue)
    {
        if (TryAppendFillGray(red, green, blue))
        {
            return;
        }

        builder.Append(PdfDocumentWriter.FormatColor(red)).Append(' ').Append(PdfDocumentWriter.FormatColor(green)).Append(' ').Append(PdfDocumentWriter.FormatColor(blue)).AppendLine(" rg");
    }

    public void SetStrokeRgb(byte red, byte green, byte blue)
    {
        if (TryAppendStrokeGray(red, green, blue))
        {
            return;
        }

        builder.Append(PdfDocumentWriter.FormatColor(red)).Append(' ').Append(PdfDocumentWriter.FormatColor(green)).Append(' ').Append(PdfDocumentWriter.FormatColor(blue)).AppendLine(" RG");
    }

    public void SetLineWidth(double width)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(width)).AppendLine(" w");
    }

    public void SetLineDash(double dashLength, double gapLength)
    {
        builder.Append('[').Append(PdfDocumentWriter.FormatNumber(dashLength)).Append(' ').Append(PdfDocumentWriter.FormatNumber(gapLength)).AppendLine("] 0 d");
    }

    public void SetLineDash(IReadOnlyList<double> lengths, double phase = 0)
    {
        builder.Append('[');
        foreach (double length in lengths)
        {
            builder.Append(PdfDocumentWriter.FormatNumber(length)).Append(' ');
        }

        builder.Append("] ");
        builder.Append(PdfDocumentWriter.FormatNumber(phase)).AppendLine(" d");
    }

    public void ClearLineDash()
    {
        builder.AppendLine("[] 0 d");
    }

    public void SetLineCap(int lineCap)
    {
        builder.Append(lineCap.ToString(CultureInfo.InvariantCulture)).AppendLine(" J");
    }

    public void SetLineJoin(int lineJoin)
    {
        builder.Append(lineJoin.ToString(CultureInfo.InvariantCulture)).AppendLine(" j");
    }

    public void SetMiterLimit(double miterLimit)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(miterLimit)).AppendLine(" M");
    }

    public void SaveState()
    {
        builder.AppendLine("q");
        stateDepth++;
    }

    public void RestoreState()
    {
        if (stateDepth == 0)
        {
            throw new InvalidDataException("PDF graphics state cannot be restored without a matching save.");
        }
        builder.AppendLine("Q");
        stateDepth--;
    }

    public void SetAlpha(double fillAlpha, double strokeAlpha)
    {
        fillAlpha = Math.Clamp(fillAlpha, 0d, 1d);
        strokeAlpha = Math.Clamp(strokeAlpha, 0d, 1d);
        string resourceName = "GS" + AlphaResourceComponent(fillAlpha) + "F" + AlphaResourceComponent(strokeAlpha) + "S";
        if (extGStateNames.Add(resourceName))
        {
            extGStates.Add(new PdfExtGStateResource(resourceName, fillAlpha, strokeAlpha));
        }

        builder.Append('/').Append(resourceName).AppendLine(" gs");
    }

    private static string AlphaResourceComponent(double alpha)
    {
        // Identity follows the writer's actual decimal rounding. Keep the
        // existing five-digit name units, and retain raw opacity in the model.
        decimal emitted = decimal.Parse(PdfDocumentWriter.FormatNumber(alpha), CultureInfo.InvariantCulture);
        return (emitted * 100000m).ToString("0", CultureInfo.InvariantCulture);
    }

    internal const string TransformCoefficientFormat = "0.######";

    public void Transform(double a, double b, double c, double d, double e, double f)
    {
        // Linear coefficients multiply page coordinates. Coordinate precision
        // here can move a distant rotation pivot by tenths of a point.
        builder.Append(PdfDocumentWriter.FormatNumber(a, TransformCoefficientFormat)).Append(' ').Append(PdfDocumentWriter.FormatNumber(b, TransformCoefficientFormat)).Append(' ');
        builder.Append(PdfDocumentWriter.FormatNumber(c, TransformCoefficientFormat)).Append(' ').Append(PdfDocumentWriter.FormatNumber(d, TransformCoefficientFormat)).Append(' ');
        builder.Append(PdfDocumentWriter.FormatNumber(e)).Append(' ').Append(PdfDocumentWriter.FormatNumber(f)).AppendLine(" cm");
    }

    public void FillRectangle(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re f");
    }

    public void FillRectangleEvenOdd(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re f*");
    }

    public void FillRectangleWithTilingPattern(double x, double y, double width, double height, PdfTilingPattern pattern)
    {
        string resourceName = "P" + (patterns.Count + 1).ToString(CultureInfo.InvariantCulture);
        if (!patternIndex.TryGetValue(pattern.ResourceKey, out int existingPatternIndex))
        {
            patternIndex[pattern.ResourceKey] = patterns.Count;
            patterns.Add(new PdfTilingPatternResource(resourceName, pattern));
        }
        else
        {
            resourceName = patterns[existingPatternIndex].ResourceName;
        }

        builder.Append("/Pattern cs /").Append(PdfEmbeddedFont.SanitizeName(resourceName)).AppendLine(" scn");
        FillRectangle(x, y, width, height);
    }

    public void PaintAxialShading(double x0, double y0, double x1, double y1, IReadOnlyList<PdfShadingStop> stops)
    {
        PaintShading(new PdfAxialShading(x0, y0, x1, y1, stops));
    }

    public void SetStrokeShadingPattern(PdfShadingPattern pattern) => SetShadingPattern(pattern, stroking: true);

    public void SetFillShadingPattern(PdfShadingPattern pattern) => SetShadingPattern(pattern, stroking: false);

    private void SetShadingPattern(PdfShadingPattern pattern, bool stroking)
    {
        ArgumentNullException.ThrowIfNull(pattern);
        if (!shadingPatternIndex.TryGetValue(pattern.ResourceKey, out int index))
        {
            index = shadingPatterns.Count;
            shadingPatterns.Add(new PdfShadingPatternResource("Ps" + (index + 1).ToString(CultureInfo.InvariantCulture), pattern));
            shadingPatternIndex.Add(pattern.ResourceKey, index);
        }
        builder.Append(stroking ? "/Pattern CS /" : "/Pattern cs /");
        builder.Append(PdfEmbeddedFont.SanitizeName(shadingPatterns[index].ResourceName)).AppendLine(stroking ? " SCN" : " scn");
    }

    internal void PaintShading(PdfShading shading)
    {
        string resourceName = "Sh" + (shadings.Count + 1).ToString(CultureInfo.InvariantCulture);
        if (!shadingIndex.TryGetValue(shading.ResourceKey, out int existingShadingIndex))
        {
            shadingIndex[shading.ResourceKey] = shadings.Count;
            shadings.Add(new PdfShadingResource(resourceName, shading));
        }
        else
        {
            resourceName = shadings[existingShadingIndex].ResourceName;
        }

        builder.Append('/').Append(resourceName).AppendLine(" sh");
    }

    public void StrokeRectangle(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re S");
    }

    public void FillStrokeRectangleEvenOdd(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re B*");
    }

    public void FillStrokeRoundedRectangleEvenOdd(double x, double y, double width, double height, double radius)
    {
        AppendRoundedRectanglePath(x, y, width, height, radius, explicitClosingLine: true);
        builder.AppendLine("B*");
    }

    public void FillRoundedRectangleEvenOdd(double x, double y, double width, double height, double radius)
    {
        AppendRoundedRectanglePath(x, y, width, height, radius);
        builder.AppendLine("f*");
    }

    public void StrokeRoundedRectangle(double x, double y, double width, double height, double radius)
    {
        AppendRoundedRectanglePath(x, y, width, height, radius);
        builder.AppendLine("S");
    }

    public void ClipRectangle(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re W n");
    }

    public void ClipRectangleEvenOdd(double x, double y, double width, double height)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).Append(' ').Append(PdfDocumentWriter.FormatNumber(width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(height)).AppendLine(" re W* n");
    }

    public void ClipEllipse(double x, double y, double width, double height)
    {
        AppendEllipsePath(x, y, width, height);
        builder.AppendLine("W n");
    }

    public void ClipRoundedRectangle(double x, double y, double width, double height, double radius)
    {
        AppendRoundedRectanglePath(x, y, width, height, radius);
        builder.AppendLine("W n");
    }

    public void ClipPolygon((double X, double Y)[] points)
    {
        AppendPolygonPath(points);
        builder.AppendLine("W n");
    }

    public void ClipCurrentPath(bool evenOdd = false)
    {
        builder.AppendLine(evenOdd ? "W* n" : "W n");
    }

    public void StrokeLine(double x1, double y1, double x2, double y2)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x1)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y1)).Append(" m ");
        builder.Append(PdfDocumentWriter.FormatNumber(x2)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y2)).AppendLine(" l S");
    }

    public void FillPolygon((double X, double Y)[] points)
    {
        AppendPolygonPath(points);
        builder.AppendLine("f");
    }

    public void FillPolygonEvenOdd((double X, double Y)[] points)
    {
        AppendPolygonPath(points);
        builder.AppendLine("f*");
    }

    public void MoveTo(double x, double y)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).AppendLine(" m");
    }

    public void LineTo(double x, double y)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).AppendLine(" l");
    }

    public void CurveTo(double x1, double y1, double x2, double y2, double x3, double y3)
    {
        Curve(x1, y1, x2, y2, x3, y3);
    }

    public void ClosePath()
    {
        builder.AppendLine("h");
    }

    public void FillCurrentPath()
    {
        builder.AppendLine("f");
    }

    public void FillCurrentPathEvenOdd()
    {
        builder.AppendLine("f*");
    }

    public void StrokeCurrentPath()
    {
        builder.AppendLine("S");
    }

    public void FillAndStrokeCurrentPathEvenOdd()
    {
        builder.AppendLine("B*");
    }

    public void FillAndStrokeCurrentPath()
    {
        builder.AppendLine("B");
    }

    public void StrokePolygon((double X, double Y)[] points)
    {
        AppendPolygonPath(points);
        builder.AppendLine("S");
    }

    public void FillEllipse(double x, double y, double width, double height)
    {
        AppendEllipsePath(x, y, width, height);
        builder.AppendLine("f");
    }

    public void FillEllipseEvenOdd(double x, double y, double width, double height)
    {
        AppendEllipsePath(x, y, width, height);
        builder.AppendLine("f*");
    }

    public void StrokeEllipse(double x, double y, double width, double height)
    {
        AppendEllipsePath(x, y, width, height);
        builder.AppendLine("S");
    }

    // RV01: diagnosed fallback text. Every rune is positioned absolutely with
    // the fallback advances, so emitted positions match fallback measurement
    // exactly without knowing viewer font advances. Character spacing is baked
    // into the rune positions, so Tc stays 0. Synthetic styles are covered by
    // face choice (bold/oblique faces), never double-struck here.
    public void DrawFallbackText(
        string fontResourceName,
        double fontSize,
        byte red,
        byte green,
        byte blue,
        IReadOnlyList<PdfFallbackGlyph> glyphs,
        CancellationToken cancellationToken = default)
    {
        if (glyphs.Count == 0)
        {
            return;
        }

        cancellationToken.ThrowIfCancellationRequested();
        string name = PdfEmbeddedFont.SanitizeName(fontResourceName);
        string size = PdfDocumentWriter.FormatNumber(fontSize);
        AdmitTextConstruction(64L + name.Length + size.Length);

        RegisterFontUse(fontResourceName);

        builder.AppendLine("BT");
        SetFillRgb(red, green, blue);

        builder.Append("/").Append(name).Append(" ").Append(size).AppendLine(" Tf");
        builder.AppendLine("0 Tc");
        foreach (PdfFallbackGlyph glyph in glyphs)
        {
            string x = PdfDocumentWriter.FormatNumber(glyph.X), y = PdfDocumentWriter.FormatNumber(glyph.Y);
            AdmitTextConstruction(32L + x.Length + y.Length);
            cancellationToken.ThrowIfCancellationRequested();
            builder.Append("1 0 0 1 ").Append(x).Append(" ").Append(y).AppendLine(" Tm");
            builder.Append("<").Append(glyph.Code.ToString("X2", CultureInfo.InvariantCulture)).AppendLine("> Tj");
        }

        builder.AppendLine("ET");
    }

    public void DrawGlyphText(
        string fontResourceName,
        double fontSize,
        double x,
        double y,
        byte red,
        byte green,
        byte blue,
        string glyphOperand,
        bool italic,
        double characterSpacing,
        int textRenderingMode,
        byte strokeRed,
        byte strokeGreen,
        byte strokeBlue,
        double strokeWidth,
        int textRotationQuarterTurns = 0,
        bool isPositioned = false)
    {
        string textOperator = isPositioned ? glyphOperand + " TJ" : '<' + glyphOperand + "> Tj";
        string name = PdfEmbeddedFont.SanitizeName(fontResourceName);
        string size = PdfDocumentWriter.FormatNumber(fontSize);
        string spacing = Math.Abs(characterSpacing) > 0.001d ? PdfDocumentWriter.FormatNumber(characterSpacing) : "0";
        string xPosition = PdfDocumentWriter.FormatNumber(x), yPosition = PdfDocumentWriter.FormatNumber(y);
        string stroke = textRenderingMode is 1 or 2 ? PdfDocumentWriter.FormatNumber(strokeWidth) : string.Empty;
        string shear = textRotationQuarterTurns == 0 ? PdfDocumentWriter.FormatNumber(italic ? SyntheticItalicShear : 0d) : string.Empty;
        // Fixed operators, byte colors, quarter turns and line endings fit in
        // 128 characters; all variable-length tokens are counted separately.
        AdmitTextConstruction(128L + textOperator.Length + name.Length + size.Length + spacing.Length +
            xPosition.Length + yPosition.Length + stroke.Length + shear.Length);
        RegisterFontUse(fontResourceName);
        builder.AppendLine("BT");
        SetFillRgb(red, green, blue);

        if (textRenderingMode is 1 or 2)
        {
            SetStrokeRgb(strokeRed, strokeGreen, strokeBlue);
            builder.Append(stroke).AppendLine(" w");
            builder.Append(textRenderingMode.ToString(CultureInfo.InvariantCulture)).AppendLine(" Tr");
        }

        builder.Append('/').Append(name).Append(' ').Append(size).AppendLine(" Tf");
        builder.Append(spacing).AppendLine(" Tc");

        if (textRotationQuarterTurns == 0)
        {
            builder.Append("1 0 ").Append(shear).Append(" 1 ").Append(xPosition).Append(' ').Append(yPosition).AppendLine(" Tm");
        }
        else
        {
            // Quarter-turn text matrices for vertical text: exact integer matrices,
            // no float dust (Office emits 0/-1/1/0 forms). Italic shear stays on the
            // identity path until a rotated-italic sample is observed.
            int turns = ((textRotationQuarterTurns % 4) + 4) % 4;
            (double ra, double rb, double rc, double rd) = turns switch
            {
                1 => (0d, -1d, 1d, 0d),
                2 => (-1d, 0d, 0d, -1d),
                3 => (0d, 1d, -1d, 0d),
                _ => (1d, 0d, 0d, 1d),
            };
            builder.Append(PdfDocumentWriter.FormatNumber(ra)).Append(' ').Append(PdfDocumentWriter.FormatNumber(rb)).Append(' ').Append(PdfDocumentWriter.FormatNumber(rc)).Append(' ').Append(PdfDocumentWriter.FormatNumber(rd)).Append(' ').Append(xPosition).Append(' ').Append(yPosition).AppendLine(" Tm");
        }
        builder.AppendLine(textOperator);
        if (textRenderingMode is 1 or 2)
        {
            builder.AppendLine("0 Tr");
        }

        builder.AppendLine("ET");
    }

    public void DrawImage(string imageResourceName, double x, double y, double width, double height)
    {
        builder.AppendLine("q");
        builder.Append(PdfDocumentWriter.FormatNumber(width)).Append(" 0 0 ").Append(PdfDocumentWriter.FormatNumber(height)).Append(' ').Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).AppendLine(" cm");
        builder.Append('/').Append(PdfEmbeddedFont.SanitizeName(imageResourceName)).AppendLine(" Do");
        builder.AppendLine("Q");
    }

    public void DrawImageCropped(string imageResourceName, double x, double y, double width, double height, double cropLeft, double cropTop, double cropRight, double cropBottom)
    {
        double visibleWidth = Math.Max(0.001d, 1d - cropLeft - cropRight);
        double visibleHeight = Math.Max(0.001d, 1d - cropTop - cropBottom);
        double scaledWidth = width / visibleWidth;
        double scaledHeight = height / visibleHeight;
        double imageX = x - cropLeft * scaledWidth;
        double imageY = y - cropBottom * scaledHeight;

        builder.AppendLine("q");
        builder.Append(PdfDocumentWriter.FormatNumber(scaledWidth)).Append(" 0 0 ").Append(PdfDocumentWriter.FormatNumber(scaledHeight)).Append(' ').Append(PdfDocumentWriter.FormatNumber(imageX)).Append(' ').Append(PdfDocumentWriter.FormatNumber(imageY)).AppendLine(" cm");
        builder.Append('/').Append(PdfEmbeddedFont.SanitizeName(imageResourceName)).AppendLine(" Do");
        builder.AppendLine("Q");
    }

    // Operation/resource boundary for node-level transactional recovery (S09).
    // Snapshots capture the append-only builder state and emitted font-use names
    // so a failed node rewinds its paint without disturbing earlier nodes. Font/image caches are intentionally
    // outside the boundary: orphan entries are inert, while index rollback could
    // dangle references held by surviving content.
    public readonly record struct ContentMark(int ContentLength, int ExtGStateCount, int ShadingCount, int PatternCount, int StateDepth, int FontUseCount, int GroupCount, int ShadingPatternCount)
    {
        public bool ActualTextOpen { get; init; }
    }

    public ContentMark MarkContent()
    {
        return new ContentMark(builder.Length, extGStates.Count, shadings.Count, patterns.Count, stateDepth, fontUseOrder.Count, groups.Count, shadingPatterns.Count)
            { ActualTextOpen = actualTextOpen };
    }

    public void TruncateContent(ContentMark mark)
    {
        if (mark.ContentLength < builder.Length)
        {
            builder.Length = Math.Max(0, mark.ContentLength);
        }

        while (fontUseOrder.Count > mark.FontUseCount)
        {
            usedFontResourceNames.Remove(fontUseOrder[^1]);
            fontUseOrder.RemoveAt(fontUseOrder.Count - 1);
        }

        while (extGStates.Count > mark.ExtGStateCount)
        {
            extGStateNames.Remove(extGStates[^1].ResourceName);
            extGStates.RemoveAt(extGStates.Count - 1);
        }

        while (shadings.Count > mark.ShadingCount)
        {
            shadingIndex.Remove(shadings[^1].Shading.ResourceKey);
            shadings.RemoveAt(shadings.Count - 1);
        }

        while (patterns.Count > mark.PatternCount)
        {
            patternIndex.Remove(patterns[^1].Pattern.ResourceKey);
            patterns.RemoveAt(patterns.Count - 1);
        }

        while (groups.Count > mark.GroupCount)
        {
            groups.RemoveAt(groups.Count - 1);
        }

        while (shadingPatterns.Count > mark.ShadingPatternCount)
        {
            shadingPatternIndex.Remove(shadingPatterns[^1].Pattern.ResourceKey);
            shadingPatterns.RemoveAt(shadingPatterns.Count - 1);
        }

        stateDepth = Math.Max(0, mark.StateDepth);
        actualTextOpen = mark.ActualTextOpen;
    }

    private void RegisterFontUse(string resourceName)
    {
        if (usedFontResourceNames.Add(resourceName)) { fontUseOrder.Add(resourceName); }
    }

    public override string ToString()
    {
        if (actualTextOpen) throw new InvalidOperationException("PDF content leaves an ActualText span open.");
        return builder.ToString();
    }

    // A structural symbol owns one Unicode scalar. An overprint has no text of its own.
    public void BeginActualText(int? sourceCodePoint)
    {
        if (sourceCodePoint is int scalar && !Rune.IsValid(scalar)) throw new ArgumentOutOfRangeException(nameof(sourceCodePoint));
        if (actualTextOpen) throw new InvalidOperationException("ActualText spans cannot be nested.");
        AdmitTextConstruction(64);
        builder.Append("/Span << /ActualText <FEFF");
        if (sourceCodePoint is int codePoint)
            builder.Append(Convert.ToHexString(Encoding.BigEndianUnicode.GetBytes(char.ConvertFromUtf32(codePoint))));
        builder.AppendLine("> >> BDC");
        actualTextOpen = true;
    }

    public void EndActualText()
    {
        if (!actualTextOpen) throw new InvalidOperationException("ActualText ends without an open span.");
        builder.AppendLine("EMC");
        actualTextOpen = false;
    }

    private void AppendEllipsePath(double x, double y, double width, double height)
    {
        const double kappa = 0.5522847498307936d;
        double rx = width / 2d;
        double ry = height / 2d;
        double cx = x + rx;
        double cy = y + ry;
        double ox = rx * kappa;
        double oy = ry * kappa;

        builder.Append(PdfDocumentWriter.FormatNumber(cx + rx)).Append(' ').Append(PdfDocumentWriter.FormatNumber(cy)).AppendLine(" m");
        Curve(cx + rx, cy + oy, cx + ox, cy + ry, cx, cy + ry);
        Curve(cx - ox, cy + ry, cx - rx, cy + oy, cx - rx, cy);
        Curve(cx - rx, cy - oy, cx - ox, cy - ry, cx, cy - ry);
        Curve(cx + ox, cy - ry, cx + rx, cy - oy, cx + rx, cy);
        builder.AppendLine("h");
    }

    private void AppendRoundedRectanglePath(double x, double y, double width, double height, double radius, bool explicitClosingLine = false)
    {
        const double kappa = 0.5522847498307936d;
        double r = Math.Clamp(radius, 0d, Math.Min(width, height) / 2d);
        double ox = r * kappa;

        builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y + height - r)).AppendLine(" m");
        Curve(x, y + height - r + ox, x + r - ox, y + height, x + r, y + height);
        builder.Append(PdfDocumentWriter.FormatNumber(x + width - r)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y + height)).AppendLine(" l");
        Curve(x + width - r + ox, y + height, x + width, y + height - r + ox, x + width, y + height - r);
        builder.Append(PdfDocumentWriter.FormatNumber(x + width)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y + r)).AppendLine(" l");
        Curve(x + width, y + r - ox, x + width - r + ox, y, x + width - r, y);
        builder.Append(PdfDocumentWriter.FormatNumber(x + r)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y)).AppendLine(" l");
        Curve(x + r - ox, y, x, y + r - ox, x, y + r);
        if (explicitClosingLine)
        {
            builder.Append(PdfDocumentWriter.FormatNumber(x)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y + height - r)).AppendLine(" l");
        }
        builder.AppendLine("h");
    }

    internal void AppendPolygonPath((double X, double Y)[] points)
    {
        if (points.Length == 0)
        {
            return;
        }

        builder.Append(PdfDocumentWriter.FormatNumber(points[0].X)).Append(' ').Append(PdfDocumentWriter.FormatNumber(points[0].Y)).AppendLine(" m");
        for (int i = 1; i < points.Length; i++)
        {
            builder.Append(PdfDocumentWriter.FormatNumber(points[i].X)).Append(' ').Append(PdfDocumentWriter.FormatNumber(points[i].Y)).AppendLine(" l");
        }

        builder.AppendLine("h");
    }

    private void Curve(double x1, double y1, double x2, double y2, double x3, double y3)
    {
        builder.Append(PdfDocumentWriter.FormatNumber(x1)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y1)).Append(' ');
        builder.Append(PdfDocumentWriter.FormatNumber(x2)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y2)).Append(' ');
        builder.Append(PdfDocumentWriter.FormatNumber(x3)).Append(' ').Append(PdfDocumentWriter.FormatNumber(y3)).AppendLine(" c");
    }

    private bool TryAppendFillGray(byte red, byte green, byte blue)
    {
        if (red != green || red != blue)
        {
            return false;
        }

        builder.Append(PdfDocumentWriter.FormatColor(red)).AppendLine(" g");
        return true;
    }

    private bool TryAppendStrokeGray(byte red, byte green, byte blue)
    {
        if (red != green || red != blue)
        {
            return false;
        }

        builder.Append(PdfDocumentWriter.FormatColor(red)).AppendLine(" G");
        return true;
    }
}
