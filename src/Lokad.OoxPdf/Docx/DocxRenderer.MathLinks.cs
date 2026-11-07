using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    // Glyphs and rules already contain the final argument/control translations.
    // Source link identity survives those translations and inline wrapping.
    private static IEnumerable<(DocxHyperlinkSpan Link, PdfRectangle Rectangle)> ReadMathHyperlinkBounds(
        DocxMathBox box, CancellationToken token)
    {
        Dictionary<DocxHyperlinkSpan, PdfRectangle>? bounds = null;
        foreach (var glyph in box.Glyphs)
        {
            token.ThrowIfCancellationRequested();
            var link = glyph.Hyperlink ?? (glyph.ControlOwner as DocxMathNode)?.Hyperlink;
            if (link is null) continue;
            var font = glyph.Font.Font.Font;
            double scale = glyph.FontSize / font.UnitsPerEm;
            if (glyph.AssemblyParts.Count == 0) AddGlyph(glyph.OriginalGlyph, glyph.X, glyph.BaselineY);
            else foreach (var part in glyph.AssemblyParts)
            {
                token.ThrowIfCancellationRequested();
                AddGlyph(part.OriginalGlyph, glyph.X + part.X, glyph.BaselineY + part.BaselineY);
            }

            void AddGlyph(ushort original, double x, double baseline)
            {
                if (!font.TryGetGlyphBounds(original, out var ink)) return;
                double shearBottom = glyph.TextShear * ink.YMin;
                double shearTop = glyph.TextShear * ink.YMax;
                double left = Math.Min(x, x + (ink.XMin + Math.Min(shearBottom, shearTop)) * scale);
                double right = Math.Max(x + glyph.Advance, x + (ink.XMax + Math.Max(shearBottom, shearTop)) * scale);
                Add(link, left, baseline + ink.YMin * scale, right, baseline + ink.YMax * scale);
            }
        }
        foreach (var rule in box.Rules)
        {
            token.ThrowIfCancellationRequested();
            if (rule.Hyperlink is { } link)
                Add(link, rule.X, rule.Y, rule.X + rule.Width, rule.Y + rule.Thickness);
        }
        if (bounds is not null)
            foreach (var pair in bounds) { token.ThrowIfCancellationRequested(); yield return (pair.Key, pair.Value); }

        void Add(DocxHyperlinkSpan link, double left, double bottom, double right, double top)
        {
            if (!(right > left) || !(top > bottom)) return;
            bounds ??= new(ReferenceEqualityComparer.Instance);
            if (bounds.TryGetValue(link, out var previous))
            {
                left = Math.Min(left, previous.X); bottom = Math.Min(bottom, previous.Y);
                right = Math.Max(right, previous.X + previous.Width); top = Math.Max(top, previous.Y + previous.Height);
            }
            bounds[link] = new(left, bottom, right - left, top - bottom);
        }
    }
}
