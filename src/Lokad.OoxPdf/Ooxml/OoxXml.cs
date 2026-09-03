using System.Globalization;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

// Centralized OOXML attribute readers. Parser and renderer classes each grew
// their own ParseLong/ReadLong/ParseOptional copies with subtly different
// missing-vs-malformed behavior; the names below make that choice explicit:
// Parse-prefixed members throw on malformed values, Read-prefixed members
// fall back to the default.
internal static class OoxXml
{
    public static long ParseRequiredLong(XElement element, XName name, string feature)
    {
        string? value = (string?)element.Attribute(name);
        if (value is null)
        {
            throw new InvalidDataException("Missing required " + feature + " attribute [" + name.LocalName + "].");
        }

        return long.Parse(value, CultureInfo.InvariantCulture);
    }

    public static long ParseOptionalLong(XElement element, string name, long defaultValue = 0L)
    {
        return element.Attribute(name) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : defaultValue;
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

    public static bool ParseBoolOrDefault(XElement? element, string name, bool defaultValue)
    {
        return OoxBoolean.ParseAttribute(element, name, defaultValue);
    }

    public static bool ReadBool(XElement element, string name)
    {
        string? value = (string?)element.Attribute(name);
        return value is "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }
}
