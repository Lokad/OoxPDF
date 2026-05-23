using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool TryReadSolidColor(XElement? element, PptxTheme theme, out RgbColor color)
    {
        return TryReadSolidColorWithAlpha(element, theme, out color, out _);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, out RgbColor color, out double alpha)
    {
        return TryReadSolidColorWithAlpha(element, theme, placeholderColorContainer: null, out color, out alpha);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, XElement? placeholderColorContainer, out RgbColor color, out double alpha)
    {
        XElement? solidFill = element?.Element(DrawingNamespace + "solidFill");
        XElement? colorContainer = element?.Name == DrawingNamespace + "solidFill"
            ? element
            : solidFill ?? element;
        alpha = ReadAlpha(colorContainer);
        XElement? srgbColor = colorContainer?.Element(DrawingNamespace + "srgbClr");
        string? hex = (string?)srgbColor?.Attribute("val");
        if (RgbColor.TryParse(hex, out color))
        {
            color = ApplyColorTransforms(srgbColor, color);
            return true;
        }

        XElement? schemeColorElement = colorContainer?.Element(DrawingNamespace + "schemeClr");
        string? schemeColor = (string?)schemeColorElement?.Attribute("val");
        if (schemeColor == "phClr" &&
            placeholderColorContainer is not null &&
            TryReadSolidColorWithAlpha(placeholderColorContainer, theme, placeholderColorContainer: null, out color, out double placeholderAlpha))
        {
            color = ApplyColorTransforms(schemeColorElement, color);
            alpha *= placeholderAlpha;
            return true;
        }

        if (schemeColor is not null && theme.TryResolveColor(schemeColor, out color))
        {
            color = ApplyColorTransforms(schemeColorElement, color);
            return true;
        }

        XElement? systemColor = colorContainer?.Element(DrawingNamespace + "sysClr");
        string? systemHex = (string?)systemColor?.Attribute("lastClr") ?? (string?)systemColor?.Attribute("val");
        if (RgbColor.TryParse(systemHex, out color))
        {
            color = ApplyColorTransforms(systemColor, color);
            return true;
        }

        XElement? presetColor = colorContainer?.Element(DrawingNamespace + "prstClr");
        if (TryResolvePresetColor((string?)presetColor?.Attribute("val"), out color))
        {
            color = ApplyColorTransforms(presetColor, color);
            return true;
        }

        XElement? scRgbColor = colorContainer?.Element(DrawingNamespace + "scrgbClr");
        if (scRgbColor is not null)
        {
            color = new RgbColor(
                ReadPercentageByte(scRgbColor, "r"),
                ReadPercentageByte(scRgbColor, "g"),
                ReadPercentageByte(scRgbColor, "b"));
            color = ApplyColorTransforms(scRgbColor, color);
            return true;
        }

        XElement? hslColor = colorContainer?.Element(DrawingNamespace + "hslClr");
        if (hslColor is not null)
        {
            color = ReadHslColor(hslColor);
            color = ApplyColorTransforms(hslColor, color);
            return true;
        }

        return false;
    }

    private static double ReadAlpha(XElement? colorContainer)
    {
        XElement? alpha = colorContainer?
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName is "srgbClr" or "schemeClr" or "sysClr" or "prstClr" or "scrgbClr" or "hslClr")
            ?.Element(DrawingNamespace + "alpha");
        if (alpha?.Attribute("val") is { } value &&
            int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
        {
            return Math.Clamp(parsed / 100000d, 0d, 1d);
        }

        return 1d;
    }

    private static RgbColor ApplyColorTransforms(XElement? colorElement, RgbColor color)
    {
        if (colorElement is null)
        {
            return color;
        }

        double red = color.Red;
        double green = color.Green;
        double blue = color.Blue;
        foreach (XElement transform in colorElement.Elements())
        {
            double value = ParseOptionalLongAttribute(transform, "val", 100000) / 100000d;
            switch (transform.Name.LocalName)
            {
                case "lumMod":
                case "shade":
                    red *= value;
                    green *= value;
                    blue *= value;
                    break;
                case "lumOff":
                    red += 255d * value;
                    green += 255d * value;
                    blue += 255d * value;
                    break;
                case "tint":
                    red += (255d - red) * value;
                    green += (255d - green) * value;
                    blue += (255d - blue) * value;
                    break;
            }
        }

        return new RgbColor(ToByte(red), ToByte(green), ToByte(blue));
    }

    private static bool TryResolvePresetColor(string? name, out RgbColor color)
    {
        return PptxPresetColors.TryResolve(name, out color);
    }

    private static byte ReadPercentageByte(XElement element, string attributeName)
    {
        double ratio = ParseOptionalLongAttribute(element, attributeName, 0) / 100000d;
        return ToByte(255d * ratio);
    }

    private static RgbColor ReadHslColor(XElement element)
    {
        double hue = (ParseOptionalLongAttribute(element, "hue", 0) / 60000d) % 360d;
        double saturation = Math.Clamp(ParseOptionalLongAttribute(element, "sat", 0) / 100000d, 0d, 1d);
        double luminosity = Math.Clamp(ParseOptionalLongAttribute(element, "lum", 0) / 100000d, 0d, 1d);
        double chroma = (1d - Math.Abs(2d * luminosity - 1d)) * saturation;
        double segment = hue / 60d;
        double second = chroma * (1d - Math.Abs(segment % 2d - 1d));
        (double r1, double g1, double b1) = segment switch
        {
            >= 0d and < 1d => (chroma, second, 0d),
            >= 1d and < 2d => (second, chroma, 0d),
            >= 2d and < 3d => (0d, chroma, second),
            >= 3d and < 4d => (0d, second, chroma),
            >= 4d and < 5d => (second, 0d, chroma),
            _ => (chroma, 0d, second)
        };
        double match = luminosity - chroma / 2d;
        return new RgbColor(ToByte((r1 + match) * 255d), ToByte((g1 + match) * 255d), ToByte((b1 + match) * 255d));
    }

    private static byte ToByte(double value)
    {
        return (byte)Math.Clamp((int)Math.Round(value, MidpointRounding.AwayFromZero), 0, 255);
    }

    private static bool TryReadLine(XElement shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth)
    {
        return TryReadLineWithAlpha(shapeProperties, theme, out color, out lineWidth, out _);
    }

    private static bool TryReadLineWithAlpha(XElement shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth, out double alpha)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        lineWidth = line?.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        return TryReadSolidColorWithAlpha(line, theme, out color, out alpha);
    }
}
