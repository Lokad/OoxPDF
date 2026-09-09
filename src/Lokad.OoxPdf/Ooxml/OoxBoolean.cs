using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

internal static class OoxBoolean
{
    public static bool IsTrue(string? value)
    {
        return value is not null &&
            (value == "1" ||
            value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("on", StringComparison.OrdinalIgnoreCase));
    }

    public static bool ParseAttribute(XElement? element, XName attributeName, bool defaultValue)
    {
        string? value = (string?)element?.Attribute(attributeName);
        return value is null ? defaultValue : IsTrue(value);
    }

    public static bool? ParseOptionalAttribute(XElement? element, XName attributeName)
    {
        XAttribute? attribute = element?.Attribute(attributeName);
        return attribute is null ? null : IsTrue(attribute.Value);
    }

    // Explicit off-values for display flags; unknown or absent values fail
    // open (shown), so garbage never drops content (S01).
    public static bool IsOff(string? value)
    {
        return value is not null &&
            (value == "0" ||
            value.Equals("false", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("off", StringComparison.OrdinalIgnoreCase));
    }

    public static bool ParseElement(XElement? element, bool defaultValue, XName? valueAttributeName)
    {
        if (element is null)
        {
            return defaultValue;
        }

        string? value = (string?)element.Attribute(valueAttributeName ?? "val");
        return value is null || IsTrue(value);
    }

}
