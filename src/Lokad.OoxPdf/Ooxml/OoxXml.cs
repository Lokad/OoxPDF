using System.Globalization;
using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Ooxml;

// Numeric Parse readers reject malformed values. Numeric Read readers supply
// the caller's default. Boolean readers delegate to the shared OOXML vocabulary.
internal static class OoxXml
{
    public static string? ReadDrawingSvgRelationshipId(XElement? blip, CancellationToken cancellationToken)
    {
        // Only a direct extension of the selected drawing blip owns the SVG relationship.
        foreach (XElement extension in blip?.Element(DrawingNamespace + "extLst")?.Elements(DrawingNamespace + "ext") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (extension.Element(DrawingSvgNamespace + "svgBlip") is { } svgBlip)
                return (string?)svgBlip.Attribute(RelationshipsNamespace + "embed");
        }

        return null;
    }

    // DrawingML accepts thousandths of a percent or a decimal percent literal.
    // Keep the existing Int32 policy for strict/try readers; color channels and
    // transforms below retain their existing Int64 fallback policy separately.
    public static bool TryParsePercentage(string? value, out double ratio)
    {
        ReadOnlySpan<char> text = value.AsSpan().Trim();
        if (!text.IsEmpty && text[^1] == '%')
        {
            if (decimal.TryParse(text[..^1], NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture, out decimal percent) &&
                percent >= int.MinValue / 1000m && percent <= int.MaxValue / 1000m)
            {
                // Decimal source units give both spellings the same rounded ratio.
                ratio = (double)(percent / 100m);
                return true;
            }
        }
        else if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int thousandths))
        {
            ratio = thousandths / 100000d;
            return true;
        }
        ratio = 0d;
        return false;
    }

    public static double ParsePercentage(string value) =>
        TryParsePercentage(value, out double ratio) ? ratio
            // Retain the strict readers' established format/overflow failures.
            : int.Parse(value, CultureInfo.InvariantCulture) / 100000d;

    public static double ReadOptionalPercentage(XElement? element, string name, double defaultValue)
    {
        string? value = (string?)element?.Attribute(name);
        return value.AsSpan().TrimEnd().EndsWith("%")
            ? TryParsePercentage(value, out double ratio) ? ratio : defaultValue
            : long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long thousandths)
                ? thousandths / 100000d : defaultValue;
    }

    public static long ParseRequiredLong(XElement element, XName name, string feature)
    {
        string? value = (string?)element.Attribute(name);
        if (value is null)
        {
            throw new InvalidDataException("Missing required " + feature + " attribute [" + name.LocalName + "].");
        }

        try
        {
            return long.Parse(value, CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (ex is FormatException or OverflowException)
        {
            throw new InvalidDataException("Malformed " + feature + " attribute [" + name.LocalName + "].", ex);
        }
    }

    public static long ParseOptionalLong(XElement element, string name, long defaultValue)
    {
        if (element.Attribute(name) is not { } attribute)
        {
            return defaultValue;
        }

        try
        {
            return long.Parse(attribute.Value, CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (ex is FormatException or OverflowException)
        {
            throw new InvalidDataException("Malformed attribute [" + name + "].", ex);
        }
    }

    public static long ReadOptionalLong(XElement? element, string name, long defaultValue)
    {
        return element?.Attribute(name) is { } attribute &&
            long.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long value)
            ? value
            : defaultValue;
    }

    public static int ReadOptionalInt(XElement? element, string name, int defaultValue)
    {
        return element?.Attribute(name) is { } attribute &&
            int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)
            ? value
            : defaultValue;
    }

    public static double ReadOptionalDouble(XElement? element, string name, double defaultValue)
    {
        return element?.Attribute(name) is { } attribute &&
            double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
            ? value
            : defaultValue;
    }

    public static bool ParseOptionalBool(XElement? element, string name)
    {
        return OoxBoolean.ParseAttribute(element, name, defaultValue: false);
    }

    public static bool ReadBool(XElement element, string name)
    {
        return OoxBoolean.IsXmlTrue((string?)element.Attribute(name));
    }
}
