using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal static IReadOnlyList<PptxTextGlyphRunSnapshot> InspectTextGlyphRuns(PptxDocument document, OoxPackage package, int slideIndex)
    {
        IReadOnlyList<TextRun> textRuns = ReadSlideTextRunsForInspection(document, package, slideIndex);
        RenderedFonts renderedFonts = CreateRenderedFonts(textRuns);
        textRuns = CoalesceAdjacentTextRuns(textRuns, compareHighlight: false);
        textRuns = CoalesceUnderlineRuns(textRuns);
        var glyphRuns = new List<PptxTextGlyphRunSnapshot>();
        foreach (TextRun run in textRuns)
        {
            if (!renderedFonts.Fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
            {
                continue;
            }

            TextGlyphRun? glyphRun = BuildTextGlyphRun(rendered.ResourceName, rendered.Font, run, rendered.SyntheticBold, rendered.SyntheticItalic);
            if (glyphRun is null)
            {
                continue;
            }

            glyphRuns.Add(new PptxTextGlyphRunSnapshot(
                run.Text,
                glyphRun.X,
                glyphRun.BaselineY,
                glyphRun.Width,
                glyphRun.Glyphs.Count,
                glyphRun.Glyphs.Skip(1).FirstOrDefault()?.AdjustmentBefore ?? 0d));
        }

        return glyphRuns;
    }

    private static IReadOnlyList<PdfFontResource> RenderTextRuns(IReadOnlyList<TextRun> textRuns, PdfGraphicsBuilder graphics)
    {
        if (textRuns.Count == 0)
        {
            return [];
        }

        RenderedFonts renderedFonts = CreateRenderedFonts(textRuns);
        DrawTextRunsWithFonts(textRuns, graphics, renderedFonts.Fonts);
        return renderedFonts.Resources;
    }

    private static RenderedFonts CreateRenderedFonts(IReadOnlyList<TextRun> textRuns)
    {
        if (textRuns.Count == 0)
        {
            return new RenderedFonts(new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase), []);
        }

        textRuns = CoalesceAdjacentTextRuns(textRuns, compareHighlight: false);
        textRuns = CoalesceUnderlineRuns(textRuns);
        var resolver = new WindowsFontResolver();
        var fonts = new Dictionary<string, RenderedFont>(StringComparer.OrdinalIgnoreCase);
        var resources = new List<PdfFontResource>();
        foreach (IGrouping<string, TextRun> group in textRuns.GroupBy(r => FontKey(r), StringComparer.OrdinalIgnoreCase))
        {
            TextRun first = group.First();
            string familyName = string.IsNullOrWhiteSpace(first.FontFamily) ? "Arial" : first.FontFamily!;
            FontResolution resolution = resolver.Resolve(new FontRequest(familyName, first.Bold, first.Italic));
            if (resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath))
            {
                continue;
            }

            OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath, resolution.FontFaceIndex);
            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, group.SelectMany(r => r.Text.EnumerateRunes().Select(rune => rune.Value)));
            string resourceName = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            fonts[group.Key] = new RenderedFont(resourceName, embedded, first.Bold && !resolution.Bold, first.Italic && !resolution.Italic);
            resources.Add(new PdfFontResource(resourceName, embedded));
        }

        return new RenderedFonts(fonts, resources);
    }

    private static void DrawTextRunsWithFonts(IReadOnlyList<TextRun> textRuns, PdfGraphicsBuilder graphics, IReadOnlyDictionary<string, RenderedFont> fonts)
    {
        DrawHighlightRunsWithFonts(textRuns, graphics, fonts);
        textRuns = CoalesceAdjacentTextRuns(textRuns, compareHighlight: false);
        textRuns = CoalesceUnderlineRuns(textRuns);
        foreach (TextRun run in textRuns)
        {
            if (fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
            {
                DrawWrappedRun(graphics, rendered.ResourceName, rendered.Font, run, rendered.SyntheticBold, rendered.SyntheticItalic);
            }
        }
    }

    private static IReadOnlyList<TextRun> CoalesceAdjacentTextRuns(IReadOnlyList<TextRun> textRuns, bool compareHighlight = true)
    {
        var coalesced = new List<TextRun>(textRuns.Count);
        foreach (TextRun run in textRuns)
        {
            if (run.Text.Length == 0)
            {
                continue;
            }

            if (coalesced.Count != 0 && CanCoalesceTextRun(coalesced[^1], run, compareHighlight))
            {
                TextRun previous = coalesced[^1];
                coalesced[^1] = previous with
                {
                    Text = previous.Text + run.Text,
                    Width = run.X + run.Width - previous.X
                };
            }
            else
            {
                coalesced.Add(run);
            }
        }

        return coalesced;
    }

    private static bool CanCoalesceTextRun(TextRun left, TextRun right, bool compareHighlight = true)
    {
        return Math.Abs(left.Y - right.Y) < 0.01d &&
            Math.Abs(left.FontSize - right.FontSize) < 0.01d &&
            Math.Abs(left.CharacterSpacing - right.CharacterSpacing) < 0.01d &&
            Math.Abs(left.BaselineOffset - right.BaselineOffset) < 0.01d &&
            Math.Abs(left.RotationDegrees - right.RotationDegrees) < 0.01d &&
            Math.Abs(left.RotationCenterX - right.RotationCenterX) < 0.01d &&
            Math.Abs(left.RotationCenterY - right.RotationCenterY) < 0.01d &&
            Math.Abs(left.ClipX - right.ClipX) < 0.01d &&
            Math.Abs(left.ClipY - right.ClipY) < 0.01d &&
            Math.Abs(left.ClipWidth - right.ClipWidth) < 0.01d &&
            Math.Abs(left.ClipHeight - right.ClipHeight) < 0.01d &&
            left.Color.Equals(right.Color) &&
            Math.Abs(left.Alpha - right.Alpha) < 0.001d &&
            (!compareHighlight || left.HighlightColor.Equals(right.HighlightColor)) &&
            !left.PreventCoalesce &&
            !right.PreventCoalesce &&
            left.Bold == right.Bold &&
            left.Italic == right.Italic &&
            left.Underline == right.Underline &&
            left.Strike == right.Strike &&
            left.KerningEnabled == right.KerningEnabled &&
            left.Alignment == right.Alignment &&
            string.Equals(left.FontFamily, right.FontFamily, StringComparison.OrdinalIgnoreCase) &&
            right.X >= left.X &&
            Math.Abs(right.X - (left.X + left.Width)) < Math.Max(1d, left.FontSize * 0.2d);
    }

    private static void DrawHighlightRunsWithFonts(IReadOnlyList<TextRun> textRuns, PdfGraphicsBuilder graphics, IReadOnlyDictionary<string, RenderedFont> fonts)
    {
        foreach (TextRun run in CoalesceHighlightRuns(textRuns))
        {
            if (run.HighlightColor is null || !fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
            {
                continue;
            }

            DrawHighlightRun(graphics, rendered.Font, run);
        }
    }

    private static IReadOnlyList<TextRun> CoalesceHighlightRuns(IReadOnlyList<TextRun> textRuns)
    {
        var coalesced = new List<TextRun>(textRuns.Count);
        foreach (TextRun run in textRuns)
        {
            if (run.Text.Length == 0 || run.HighlightColor is null)
            {
                continue;
            }

            if (coalesced.Count != 0 && CanCoalesceTextRun(coalesced[^1], run))
            {
                TextRun previous = coalesced[^1];
                coalesced[^1] = previous with
                {
                    Text = previous.Text + run.Text,
                    Width = run.X + run.Width - previous.X
                };
            }
            else
            {
                coalesced.Add(run);
            }
        }

        return coalesced;
    }

    private static IReadOnlyList<TextRun> CoalesceUnderlineRuns(IReadOnlyList<TextRun> textRuns)
    {
        var coalesced = new List<TextRun>(textRuns.Count);
        foreach (TextRun run in textRuns)
        {
            if (coalesced.Count > 0 && CanCoalesceUnderlineRun(coalesced[^1], run))
            {
                TextRun previous = coalesced[^1];
                coalesced[^1] = previous with
                {
                    Text = previous.Text + run.Text,
                    Width = Math.Max(previous.Width, run.X + run.Width - previous.X)
                };
                continue;
            }

            coalesced.Add(run);
        }

        return coalesced;
    }

    private static bool CanCoalesceUnderlineRun(TextRun left, TextRun right)
    {
        return left.Underline &&
            right.Underline &&
            !left.Strike &&
            !right.Strike &&
            left.Bold == right.Bold &&
            left.Italic == right.Italic &&
            left.KerningEnabled == right.KerningEnabled &&
            left.Alignment == right.Alignment &&
            string.Equals(left.FontFamily, right.FontFamily, StringComparison.OrdinalIgnoreCase) &&
            left.Color.Equals(right.Color) &&
            NearlyEqual(left.Alpha, right.Alpha) &&
            left.HighlightColor.Equals(right.HighlightColor) &&
            !left.PreventCoalesce &&
            !right.PreventCoalesce &&
            NearlyEqual(left.Y, right.Y) &&
            NearlyEqual(left.Height, right.Height) &&
            NearlyEqual(left.ClipX, right.ClipX) &&
            NearlyEqual(left.ClipY, right.ClipY) &&
            NearlyEqual(left.ClipWidth, right.ClipWidth) &&
            NearlyEqual(left.ClipHeight, right.ClipHeight) &&
            NearlyEqual(left.FontSize, right.FontSize) &&
            NearlyEqual(left.CharacterSpacing, right.CharacterSpacing) &&
            NearlyEqual(left.BaselineOffset, right.BaselineOffset) &&
            NearlyEqual(left.RotationDegrees, right.RotationDegrees) &&
            NearlyEqual(left.RotationCenterX, right.RotationCenterX) &&
            NearlyEqual(left.RotationCenterY, right.RotationCenterY) &&
            Math.Abs((left.X + left.Width) - right.X) <= Math.Max(1d, left.FontSize * 0.08d);
    }

    private static bool NearlyEqual(double left, double right)
    {
        return Math.Abs(left - right) <= 0.001d;
    }

    private static string FontKey(TextRun run)
    {
        string familyName = string.IsNullOrWhiteSpace(run.FontFamily) ? "Arial" : run.FontFamily!;
        return familyName + "\u001f" + run.Bold.ToString(CultureInfo.InvariantCulture) + "\u001f" + run.Italic.ToString(CultureInfo.InvariantCulture);
    }

    private static void DrawWrappedRun(PdfGraphicsBuilder graphics, string resourceName, PdfEmbeddedFont embedded, TextRun run, bool syntheticBold, bool syntheticItalic)
    {
        graphics.SaveState();
        if (Math.Abs(run.RotationDegrees) > 0.001d)
        {
            ApplyTextRotation(graphics, run.RotationDegrees, run.RotationCenterX, run.RotationCenterY);
        }

        graphics.ClipRectangle(run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
        TextGlyphRun? glyphRun = BuildTextGlyphRun(resourceName, embedded, run, syntheticBold, syntheticItalic);
        if (glyphRun is not null)
        {
            bool transparentText = run.Alpha < 0.999d;
            if (transparentText)
            {
                graphics.SaveState();
                graphics.SetAlpha(run.Alpha, 1d);
            }

            DrawGlyphText(graphics, glyphRun);
            if (syntheticBold)
            {
                DrawGlyphText(graphics, glyphRun with { X = glyphRun.X + 0.35d });
            }

            if (run.Underline)
            {
                graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                double underlineScale = run.FontSize / embedded.Font.UnitsPerEm;
                double underlineThickness = Math.Max(0.5d, Math.Abs(embedded.Font.Post.UnderlineThickness) * underlineScale);
                double underlineY = glyphRun.BaselineY + (embedded.Font.Post.UnderlinePosition - Math.Abs(embedded.Font.Post.UnderlineThickness)) * underlineScale;
                graphics.FillRectangle(glyphRun.X, underlineY, glyphRun.Width, underlineThickness);
            }

            if (run.Strike)
            {
                graphics.SetFillRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                graphics.FillRectangle(glyphRun.X, glyphRun.BaselineY + run.FontSize * 0.211d, glyphRun.Width, Math.Max(0.5d, run.FontSize * 0.05d));
            }

            if (transparentText)
            {
                graphics.RestoreState();
            }
        }

        graphics.RestoreState();
    }

    private static TextGlyphRun? BuildTextGlyphRun(string resourceName, PdfEmbeddedFont embedded, TextRun run, bool syntheticBold, bool syntheticItalic)
    {
        string glyphHex = embedded.EncodeGlyphHex(run.Text);
        double baselineY = run.Y + run.BaselineOffset;
        if (glyphHex.Length == 0 || !BaselineIntersectsClip(run, baselineY))
        {
            return null;
        }

        double lineWidth = MeasureRenderedText(embedded, run.Text, run.FontSize, run.CharacterSpacing, run.KerningEnabled);
        double x = run.Alignment switch
        {
            TextAlignment.Center => run.X + Math.Max(0, run.Width - lineWidth) / 2d,
            TextAlignment.Right => run.X + Math.Max(0, run.Width - lineWidth),
            _ => run.X
        };
        string? positioningArray = embedded.EncodeGlyphPositioningArray(run.Text, run.CharacterSpacing, run.FontSize, forcePositioningArray: true, run.KerningEnabled);
        return new TextGlyphRun(run, resourceName, embedded, glyphHex, positioningArray, BuildTextGlyphAtoms(embedded, run), x, baselineY, lineWidth, syntheticBold, syntheticItalic);
    }

    private static IReadOnlyList<TextGlyphAtom> BuildTextGlyphAtoms(PdfEmbeddedFont embedded, TextRun run)
    {
        var atoms = new List<TextGlyphAtom>();
        ushort previousGlyph = 0;
        foreach (Rune rune in run.Text.EnumerateRunes())
        {
            ushort glyph = embedded.Font.MapCodePoint(rune.Value);
            if (glyph == 0)
            {
                continue;
            }

            double adjustmentBefore = 0d;
            if (atoms.Count > 0)
            {
                adjustmentBefore += run.CharacterSpacing;
                if (run.KerningEnabled && previousGlyph != 0)
                {
                    adjustmentBefore += embedded.Font.GetKerning(previousGlyph, glyph) * run.FontSize / embedded.Font.UnitsPerEm;
                }
            }

            double advance = embedded.Font.GetAdvanceWidth(glyph) * run.FontSize / embedded.Font.UnitsPerEm;
            atoms.Add(new TextGlyphAtom(rune.Value, glyph, advance, adjustmentBefore));
            previousGlyph = glyph;
        }

        return atoms;
    }

    private static void DrawHighlightRun(PdfGraphicsBuilder graphics, PdfEmbeddedFont embedded, TextRun run)
    {
        if (run.HighlightColor is not { } highlight)
        {
            return;
        }

        double baselineY = run.Y + run.BaselineOffset;
        if (!BaselineIntersectsClip(run, baselineY))
        {
            return;
        }

        double lineWidth = MeasureRenderedText(embedded, run.Text, run.FontSize, run.CharacterSpacing, run.KerningEnabled);
        graphics.SaveState();
        if (Math.Abs(run.RotationDegrees) > 0.001d)
        {
            ApplyTextRotation(graphics, run.RotationDegrees, run.RotationCenterX, run.RotationCenterY);
        }

        graphics.ClipRectangle(run.ClipX, run.ClipY, run.ClipWidth, run.ClipHeight);
        graphics.SetFillRgb(highlight.Red, highlight.Green, highlight.Blue);
        double fontScale = run.FontSize / embedded.Font.UnitsPerEm;
        double highlightY = baselineY - (embedded.Font.Os2.WindowsDescender + 32d) * fontScale;
        double highlightHeight = (embedded.Font.Os2.WindowsAscender + embedded.Font.Os2.WindowsDescender) * fontScale;
        graphics.FillRectangle(run.X, highlightY, lineWidth, highlightHeight);
        graphics.RestoreState();
    }

    private static bool BaselineIntersectsClip(TextRun run, double baselineY)
    {
        return baselineY >= run.ClipY && baselineY <= run.ClipY + run.ClipHeight + run.FontSize;
    }

    private static void ApplyTextRotation(PdfGraphicsBuilder graphics, double rotationDegrees, double centerX, double centerY)
    {
        double radians = -rotationDegrees * Math.PI / 180d;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);
        double e = centerX - cos * centerX + sin * centerY;
        double f = centerY - sin * centerX - cos * centerY;
        graphics.Transform(cos, sin, -sin, cos, e, f);
    }

    private static void DrawGlyphText(PdfGraphicsBuilder graphics, TextGlyphRun glyphRun)
    {
        TextRun run = glyphRun.Source;
        if (glyphRun.PositioningArray is null)
        {
            graphics.DrawGlyphText(glyphRun.ResourceName, run.FontSize, glyphRun.X, glyphRun.BaselineY, run.Color.Red, run.Color.Green, run.Color.Blue, glyphRun.GlyphHex, glyphRun.SyntheticItalic);
        }
        else
        {
            graphics.DrawGlyphPositionedText(glyphRun.ResourceName, run.FontSize, glyphRun.X, glyphRun.BaselineY, run.Color.Red, run.Color.Green, run.Color.Blue, glyphRun.PositioningArray, glyphRun.SyntheticItalic);
        }
    }

    private static IEnumerable<string> WrapWords(string text, double maxWidth, double fontSize, double characterSpacing, PdfEmbeddedFont embedded)
    {
        if (MeasureRenderedText(embedded, text, fontSize, characterSpacing, kerningEnabled: true) <= maxWidth)
        {
            yield return text;
            yield break;
        }

        string[] words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0)
        {
            yield break;
        }

        var line = new StringBuilder();
        foreach (string word in words)
        {
            string candidate = line.Length == 0 ? word : line + " " + word;
            if (line.Length > 0 && MeasureRenderedText(embedded, candidate, fontSize, characterSpacing, kerningEnabled: true) > maxWidth)
            {
                yield return line.ToString();
                line.Clear();
                line.Append(word);
            }
            else
            {
                line.Clear();
                line.Append(candidate);
            }
        }

        if (line.Length > 0)
        {
            yield return line.ToString();
        }
    }

    private static double MeasureRenderedText(PdfEmbeddedFont embedded, string text, double fontSize, double characterSpacing, bool kerningEnabled)
    {
        double width = embedded.MeasureTextPoints(text, fontSize, kerningEnabled);
        int runeCount = text.EnumerateRunes().Count();
        return Math.Max(0d, width + Math.Max(0, runeCount - 1) * characterSpacing);
    }
}
