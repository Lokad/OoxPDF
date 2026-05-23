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

        return false;
    }

    private static double ReadAlpha(XElement? colorContainer)
    {
        XElement? alpha = colorContainer?
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName is "srgbClr" or "schemeClr" or "sysClr" or "prstClr")
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
        if (name is not null && PresetColors.TryGetValue(name, out color))
        {
            return true;
        }

        color = default;
        return false;
    }

    private static readonly IReadOnlyDictionary<string, RgbColor> PresetColors = new Dictionary<string, RgbColor>(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = new(0x00, 0x00, 0x00),
        ["white"] = new(0xFF, 0xFF, 0xFF),
        ["red"] = new(0xFF, 0x00, 0x00),
        ["green"] = new(0x00, 0x80, 0x00),
        ["blue"] = new(0x00, 0x00, 0xFF),
        ["yellow"] = new(0xFF, 0xFF, 0x00),
        ["cyan"] = new(0x00, 0xFF, 0xFF),
        ["magenta"] = new(0xFF, 0x00, 0xFF),
        ["orange"] = new(0xFF, 0xA5, 0x00),
        ["purple"] = new(0x80, 0x00, 0x80),
        ["gray"] = new(0x80, 0x80, 0x80),
        ["grey"] = new(0x80, 0x80, 0x80),
        ["lime"] = new(0x00, 0xFF, 0x00),
        ["navy"] = new(0x00, 0x00, 0x80),
        ["teal"] = new(0x00, 0x80, 0x80),
        ["maroon"] = new(0x80, 0x00, 0x00),
        ["olive"] = new(0x80, 0x80, 0x00),
        ["silver"] = new(0xC0, 0xC0, 0xC0),
        ["aqua"] = new(0x00, 0xFF, 0xFF),
        ["fuchsia"] = new(0xFF, 0x00, 0xFF),
        ["darkBlue"] = new(0x00, 0x00, 0x8B),
        ["darkGreen"] = new(0x00, 0x64, 0x00),
        ["darkRed"] = new(0x8B, 0x00, 0x00),
        ["gold"] = new(0xFF, 0xD7, 0x00),
        ["cornflowerBlue"] = new(0x64, 0x95, 0xED)
    };

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
